Attribute VB_Name = "banktovp"
Sub vp_BankToVP()
Attribute vp_BankToVP.VB_ProcData.VB_Invoke_Func = " \n14"
'
' vp_BankToVP Macro
'

'
    Dim fldr As FileDialog
    Dim exdr As FileDialog
    Dim dir As String
    Dim dirVP As String
    Dim wbVP As Workbook
    
    
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    ' Set .xlsx Folder
    With fldr
        .Title = "Escolha a Pasta com os relatorios do banco .xlsx"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path
        If .Show <> -1 Then GoTo NextStepOne
        dir = .SelectedItems(1)
    End With
NextStepOne:
    'GetFolder = dir
    Set fldr = Nothing

    'Set VP dir
    Set exdr = Application.FileDialog(msoFileDialogFilePicker)
    With exdr
        .Title = "Escolha a Planilha VP"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path
        If .Show <> -1 Then GoTo NextStepTwo
        dirVP = .SelectedItems(1)
    End With
NextStepTwo:
    'GetFolder = dirVP
    Set exdr = Nothing
    
    ' Main Interface
    If dir <> "" And dirVP <> "" Then
        Dim oFSO As Object
        Dim oFolder As Object
        Dim oFile As Object
        
        'Reset Workplace
        Rows("11:1048576").EntireRow.Hidden = False
        Rows("11:1048576").Select
        Selection.ClearContents
        
        Application.AskToUpdateLinks = False
        Application.DisplayAlerts = False
        
        'Open VP
        Set wbVP = Application.Workbooks.Open(dirVP)
        
        'Load one xlsx at a time
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        Set oFolder = oFSO.GetFolder(dir)
        For Each oFile In oFolder.files
            'If found File is xlsx
            If UBound(Split(oFile.name, ".xlsx")) > 0 Then
                'Open Workbook
                Application.Workbooks.Open (dir & "\" & oFile.name)
                'run all
                If Not oFile.name = "Bretas.xlsx" Then
                    AllRolls oFile.name, wbVP.name
                Else
                    AllBretas oFile.name, wbVP.name
                End If
                'close workbook
                Workbooks(oFile.name).Close SaveChanges:=False
            End If
            
        Next oFile
        'End
        ThisWorkbook.Activate
        Application.DisplayAlerts = True
        Application.AskToUpdateLinks = True
        'spaghetti code for global variable, used in the fixer
        Range("L4").value = wbVP.name
    End If
    
End Sub

Public Sub InsertBank(filesArray As Variant)
    Dim dirVP As String
    Dim dir As String
    Dim str As String
    
    dirVP = Cells(1, 2).value
    dir = Cells(3, 2).value

    'Reset Workplace
    Sheets("Relatorio de Faltas").Activate
    Rows("2:1048576").EntireRow.Hidden = False
    Rows("2:1048576").Select
    Selection.ClearContents
    
    'Msg Disable
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    
    'Open VP
    Set wbVP = Application.Workbooks.Open(dirVP)

    For Each file In filesArray
        str = file
        Application.Workbooks.Open (dir & "\" & file)
        If InStr(1, file, "bretas", vbTextCompare) Then
            AllBretas str, wbVP.name
        Else
            AllRolls str, wbVP.name
        End If
    Next file
    
    ThisWorkbook.Activate
    'Msg Enable
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
End Sub

Private Sub AllBretas(Workbook As String, target As String)
    Dim nextRow As Integer
    nextRow = 1
    Windows(Workbook).Activate
    Range("A1").Select
    'Find first item
    Do While OrElse(Selection.value)
        Selection.Offset(1, 0).Select
    Loop

    Do While nextRow > 0
        OneBretas Workbook, target
        nextRow = GetNextRow
        Selection.Offset(nextRow, 0).Select
    Loop
End Sub

Private Sub AllRolls(Workbook As String, target As String)
    Dim nextRow As Integer
    nextRow = 1
    Windows(Workbook).Activate
    Range("A1").Select
    'Find first item
    Do While OrElse(Selection.value)
        Selection.Offset(1, 0).Select
    Loop

    Do While nextRow > 0
        OneRoll Workbook, target
        nextRow = GetNextRow
        Selection.Offset(nextRow, 0).Select
    Loop
End Sub

Private Function OrElse(selcV As Variant) As Boolean
    If IsEmpty(selcV) Then
        OrElse = True
    Else
        OrElse = Not IsNumeric(Split(SpaceRemover(selcV), "-")(0))
    End If
    
End Function

Private Function GetNextRow() As Variant
    If IsEmpty(Selection.Offset(1, 0)) And Not IsEmpty(Selection.Offset(2, 0)) Then
        GetNextRow = 2
    ElseIf IsEmpty(Selection.Offset(1, 0)) And IsEmpty(Selection.Offset(2, 0)) Then
        GetNextRow = 0
    Else
        GetNextRow = 1
    End If
End Function

Private Sub OneRoll(Workbook As String, target As String)
    Dim BR As clsBankRoll
    Set BR = SetBankRoll
    
    Windows(target).Activate
    
    
    If CustomFinder(BR) Then
        'If found, write at
        Selection.Offset(0, -2).FormulaLocal = WriteFormula(Selection.Offset(0, -2).formula, BR)
    Else
        Windows(Workbook).Activate
        WriteMissingRoll
    End If
    
    Windows(Workbook).Activate
End Sub

Private Sub OneBretas(Workbook As String, target As String)
    Dim BR As clsBankRoll
    Set BR = SetBankRollBretas
    
    Windows(target).Activate
    On Error Resume Next
    Finder (BR.GetNumber)
    'MsgBox ("Selection: " & Selection.Offset(0, -2).value & vbNewLine & "BR: " & BR.GetValue)
    If Selection.Offset(0, -2).value = BR.GetValue Then
        'If found, write at
        Selection.Offset(0, -2).FormulaLocal = WriteFormula(Selection.Offset(0, -2).formula, BR)
    Else
        Windows(Workbook).Activate
        WriteMissingRoll
    End If
    
    Windows(Workbook).Activate
End Sub

Private Function WriteFormula(formula As Variant, BR As clsBankRoll) As String

    formula = "=" & formula & "-" & BR.GetValue
    If BR.hasExtra Then
        formula = formula & "-" & BR.GetExtra
    End If
    formula = Replace(formula, "==", "=")
    formula = Replace(formula, ".", ",")
    formula = SpaceRemover(formula)
    WriteFormula = formula

End Function

Private Function SetBankRoll() As clsBankRoll
    Dim rBR As New clsBankRoll
    rBR.SetNumber (Selection.Offset(0, 1).value)
    rBR.name = Selection.Offset(0, 2).value
    rBR.SetDate (Selection.Offset(0, 3).value)
    rBR.SetValue (Selection.Offset(0, 4).value)
    rBR.SetExtra (Selection.Offset(0, 6).value)
    'If next row is empty but the one after isn't, this is a bankroll with interest
    'Two empty rows means end of document
    rBR.hasExtra = IsEmpty(Selection.Offset(1, 0)) And Not IsEmpty(Selection.Offset(2, 0))
    Set SetBankRoll = rBR
End Function

Private Function SetBankRollBretas() As clsBankRoll
    Dim rBR As New clsBankRoll
    'If starts with 00, nf is at H
    rBR.SetNumber (Split(Selection.value, "-")(0))
    rBR.name = "Bretas linha: " & Selection.row
    rBR.SetDate (Selection.Offset(0, 2).value)
    rBR.SetValue (Selection.Offset(0, 4).value)
    rBR.SetExtra (0)
    rBR.hasExtra = False
    Set SetBankRollBretas = rBR
End Function


Private Function CustomFinder(BR As clsBankRoll) As Boolean
    'Find number
    'Big Check
    Dim isValid As Integer
    isValid = 0
    Dim tries As Integer
    
    isValid = False
    tries = 0
    Do While (isValid = 0) And tries < 5
        On Error GoTo err_handle
            Finder (BR.GetNumber)
            isValid = FullCheck(BR)
        
        tries = tries + 1
    Loop
    If isValid = 0 Then
        CustomFinder = False
    Else
        CustomFinder = True
    End If
    Exit Function
    
err_handle:
            'Every false return is cataloged in 'blank' row 11
            CustomFinder = False
            Exit Function
    
    
    
End Function

Private Sub WriteMissingRoll() 'invoice As clsNF)
    Dim cWB As String
    Dim cRow As Integer
    Dim nRow As Integer
    Dim vendedor As String
    Dim lRow As Long
    
    cWB = ActiveWorkbook.name
    cRow = Selection.row
    Range(cRow & ":" & cRow).Copy
    ThisWorkbook.Activate
    Range("A2").Select
    lRow = Cells(Rows.Count, 1).End(xlUp).row
    'Do While Not IsEmpty(Selection.value)
        'Selection.Offset(1, 0).Select
    'Loop
    nRow = lRow + 1
    Range(nRow & ":" & nRow).PasteSpecial Paste:=xlPasteValues
    'ActiveSheet
    Range("A" & nRow).value = cWB
    Windows(cWB).Activate
    
End Sub

' Checks if bank roll is equivalent to found VP, by checking if the value matches the payment
' 0 = Invalid Find
' 1 = Valid
' 2 = Valid but already writting
Private Function FullCheck(BR As clsBankRoll) As Integer
    Dim quotas As Integer
    Dim quotaIndex As Integer
    Dim quotasAlreadyPaid As Integer
    Dim quotaValue As Variant
    Dim total As Variant
    Dim arrDate() As String
    Dim arrFormula() As String
    
    'Set Array
    total = Selection.Offset(0, -2).formula
    total = Replace(total, "=", "")
    total = Replace(total, ".", ",")
    arrFormula = Split(total, "-")
    'Get total
    total = Selection.Offset(0, -2).formula
    On Error Resume Next
    total = Split(total, "-")(0)
    total = Replace(total, "=", "")
    total = Replace(total, ".", ",")
    total = CDec(total)
    'Get All dates
    arrDate = Split(Selection.Offset(0, -1))
    ' Check how many quotas this roll has
    quotas = UBound(arrDate) + 1
    'If looking at trash
    If Not IsNumeric(total) Or Not IsNumeric(BR.GetValue) Then
        FullCheck = 0
    'If grand totals are equal, this might be the correct row
    ElseIf IsSameValue(BR.GetValue * quotas, total, quotas) Then
        'Check if it has already been subtracted
        'quota index tells which quota in being paid. Starts at 1. Zero means not found
        If quotas > 1 Then
            quotaIndex = InArrayIndex(BR.GetDateAsString, arrDate) + 1
        Else
            ' no quotas, no need to check date
            quotaIndex = 1
        End If
        'already paid tells how many quotas have been paid already. If there are no quotas it returns 1 because it finds the total
        quotasAlreadyPaid = CountOccurances(BR.GetValue, arrFormula)
        ' If the quota is ahead from what has been paid or there are no quotas
        If (quotaIndex > quotasAlreadyPaid) Or (quotaIndex = quotasAlreadyPaid And quotas = 1) Then
            FullCheck = 1
        ' If already there
        ElseIf (quotaIndex = quotasAlreadyPaid) Or (quotaIndex = quotas And quotasAlreadyPaid = 2) And quotaIndex <> 0 Then
            FullCheck = 0
        End If
    Else
        FullCheck = 0
    End If
    
    
End Function

Private Function IsSameValue(first As Variant, second As Variant, quotas As Integer) As Boolean
    ' Due to rounding, one payment might be up to one cent off for each quota (if floored) or one cent pass (if ceiling)
    Dim res As Variant
    res = Abs(CDec(first) - CDec(second))
    IsSameValue = (0 <= res And res <= 0.1 * quotas)
End Function

Private Function InArrayIndex(stringToBeFound As String, arr As Variant) As Integer
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            InArrayIndex = i
            Exit Function
        End If
    Next i
    InArrayIndex = -1

End Function

Private Function CountOccurances(value As Variant, arr As Variant) As Integer
    Dim i As Integer
    Dim occur As Integer
    occur = 0
    For i = LBound(arr) To UBound(arr)
        If IsSameValue(value, arr(i), 1) Then
            occur = occur + 1
        End If
    Next i
    CountOccurances = occur
End Function


Private Function Finder(search As Variant)
        
    Cells.Find(What:=search, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
    :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    False, SearchFormat:=False).Activate
    
End Function

Function SpaceRemover(toRemove As Variant) As Variant
    toRemove = Replace(toRemove, " ", "")
    toRemove = Replace(toRemove, " ", "")
    SpaceRemover = toRemove
End Function
