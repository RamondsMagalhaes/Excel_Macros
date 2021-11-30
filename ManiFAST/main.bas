Attribute VB_Name = "main"
Sub interface()
Attribute interface.VB_ProcData.VB_Invoke_Func = " \n14"
'
' interface Macro
'

'
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "asd"
    Range("B2").Select
End Sub

Public Sub EscolherDistribuicao()
    Dim exdr As FileDialog
    Dim dirVP As String
    
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
    If dirVP <> "" Then Range("B1").Value = dirVP
    If dirVP = "" Then Exit Sub
    'SETUP
    'Find Manifesto
    Dim wbDis
    Dim wbThis
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Set wbThis = Application.ActiveWorkbook
    Range("B5:Z5").ClearContents
    Range("B7:Z7").ClearContents
    Range("B8:Z8").ClearContents
    Set wbDis = Application.Workbooks.Open(Range("B1").Value)
    wbDis.Activate
    Cells.Find(What:="manifesto", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Do While StrComp(Selection.Value, "total vendedores", vbTextCompare)
        Selection.Offset(0, 1).Select
        Dim grabvalue As String
        grabvalue = Selection.Value
        grabvalue = grabvalue & " (" & Selection.row & ", " & Selection.Column & ")"
        If Not (StrComp(Selection.Value, "Preço") = 0 Or StrComp(Selection.Value, "TOTAL", vbTextCompare) = 0 Or StrComp(Selection.Value, "total vendedores", vbTextCompare) = 0) Then
            wbThis.Activate
            Range("B5").Activate
            Do While Not IsEmpty(Selection)
                Selection.Offset(0, 1).Select
            Loop
            Selection.Value = grabvalue
            wbDis.Activate
        End If
    Loop
    Dim dir As String
    dir = wbDis.Path
    wbDis.Close SaveChanges:=False
    wbThis.Activate
    
    'WRITE FILE NAMES
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim index As Integer
              
    index = 0
    'Load one xlsx at a time
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(dir)
    For Each oFile In oFolder.Files
        'If found File is xlsx
        If UBound(Split(oFile.Name, ".xls")) > 0 And InStr(oFile.Name, "manif") Then
            'Write name
            Cells(7, 2 + index).Value = oFile.Name
            Cells(6, 2 + index).Value = "" & (index + 1)
            index = index + 1
        End If
    
    Next oFile
    
End Sub

Public Sub EscolherBaseManif()
    Dim exdr As FileDialog
    Dim dirVP As String
    
    'Set VP dir
    Set exdr = Application.FileDialog(msoFileDialogFilePicker)
    With exdr
        .Title = "Escolha a Planilha VP"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path
        If .Show <> -1 Then GoTo NextStepTowo
        dirVP = .SelectedItems(1)
    End With
NextStepTowo:
    'GetFolder = dirVP
    Set exdr = Nothing
    If dirVP <> "" Then Range("B3").Value = dirVP
End Sub
Sub setup()
Attribute setup.VB_ProcData.VB_Invoke_Func = " \n14"
'
' setup Macro
'

'
    Cells.Find(What:="manifesto", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Cells.FindNext(After:=ActiveCell).Activate
    Range("CR19").Select
    Cells.Find(What:="manifesto", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Range("CU1").Select
    ActiveWindow.SmallScroll ToRight:=13
    Range("DJ1").Select
    ActiveWindow.SmallScroll ToRight:=12
    Range("DT1").Select
    ActiveWindow.SmallScroll ToRight:=7
    Range("EA1").Select
    Windows("maniFAST v1.0.xlsm").Activate
End Sub

Sub comecar()
    'extract dic
    Dim folderDir As String
    Dim thisWB As Workbook
    Dim BaseWB As Workbook
    Dim DisWB As Workbook
    folderDir = extractDir(Range("B1").Value)
    Set thisWB = Application.ActiveWorkbook
    Set BaseWB = Workbooks.Open(Range("B3").Value)
    thisWB.Activate
    Set DisWB = Workbooks.Open(Range("B1").Value)
    thisWB.Activate
    Range("B7").Select
    Application.DisplayAlerts = False 'switching off the alert button
    While Not IsEmpty(Selection.Value)
        'CORE
        'OPEN SHEETS
        Dim manif As Workbook
        Set manif = Workbooks.Open(folderDir & "\" & Selection.Value)
        While manif.Sheets.Count > 3
            manif.Sheets(1).Delete
        Wend
        manif.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        manif.ActiveSheet.Name = "PRODUTOS"
        manif.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        manif.ActiveSheet.Name = Day(Date) & "M"
        manif.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        manif.ActiveSheet.Name = Day(Date) & "R"
        manif.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        manif.ActiveSheet.Name = Day(Date) & "A"
        BaseWB.Activate
        Cells.Select
        Selection.Copy
        manif.Activate
        Cells.Select
        ActiveSheet.Paste

        
        'PASTE DIST
        Dim row, col As Integer
        Dim cord As String
        Dim cords() As String
        thisWB.Activate
        cord = Selection.Offset(1, 0).Value
        cord = Split(cord, "(")(1)
        cord = Split(cord, ")")(0)
        cords = Split(cord, ",")
        row = CInt(cords(0))
        col = CInt(cords(1))
        DisWB.Activate
        Dim msgstr As String
        Cells(row + 450, col + 2).Select
        msgstr = "Vendedor: " & Cells(row, col).Value
        msgstr = msgstr & vbNewLine & "Total: " & Cells(row + 450, col + 2).Text
        MsgBox msgstr
        Range(Cells(row + 1, col), Cells(row + 449, col + 1)).Select
        Selection.Copy
        manif.Activate
        Range("D2").PasteSpecial Paste:=xlPasteValues
        Cells.Select
        Selection.Columns.AutoFit
        Cells.Copy
        Sheets(Day(Date) & "R").Activate
        Cells.Select
        ActiveSheet.Paste
        Cells.Copy
        Sheets(Day(Date) & "M").Activate
        Cells.Select
        ActiveSheet.Paste
        
        'FILTER AND DELETE
        Range("G:Z").Select
        Selection.Delete
        Columns("D:D").Select
        Selection.AutoFilter Field:=1, Criteria1:="-"
        Range("A2:AA9999").SpecialCells(xlCellTypeVisible).Delete
        ActiveSheet.ShowAllData
        Cells.Copy
        Sheets("PRODUTOS").Activate
        ActiveSheet.Paste
        '------------------------------------->
        'next
        manif.Save
        manif.Close
        thisWB.Activate
        Selection.Offset(0, 1).Select
    Wend
    BaseWB.Close
    DisWB.Close
    Application.DisplayAlerts = True 'switching on the alert button
End Sub

Private Function extractDir(fullDir As String) As String
    Dim dir As String
    Dim dirsplit() As String
    dirsplit = Split(fullDir, "\")
    Dim index As Integer
    index = 0
    For Each word In dirsplit
        If Not index = UBound(dirsplit) Then
            dir = dir & word & "\"
        End If
        index = index + 1
    Next
    extractDir = dir
End Function
