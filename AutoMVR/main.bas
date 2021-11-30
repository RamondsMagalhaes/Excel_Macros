Attribute VB_Name = "main"
Option Compare Text

Sub mainMVR()

    'variables
    Dim mvr_list As Object
    Dim file_name_list As Object
    Dim unique_names As Object
    Set mvr_list = CreateObject("System.Collections.ArrayList")
    Set unique_names = CreateObject("System.Collections.ArrayList")
    Set file_name_list = CreateObject("System.Collections.ArrayList")
    Set d = CreateObject("Scripting.Dictionary")
    Dim one_mvr As Object
    Dim folder_dir As String
    'Build all MVR
    
    folder_dir = GetFolder & "\"
    If folder_dir <> "" Then
        'GET FULL NAME OF TXT
        Dim file_name As String
        file_name = dir(folder_dir)
        Do While file_name <> ""
            file_name_list.Add file_name
            file_name = dir()
        Loop
        'GET UNIQUE NAMES OF TXT FILES
        For Each e In file_name_list
            e = Split(e)(0)
            d(e) = e
        Next
        For Each name In d.items
            unique_names.Add name
        Next
        ' SET ALL MVR TO BE USED
        For Each name In unique_names
            Set one_mvr = New MVR
            one_mvr.name = name
            one_mvr.manifestotxt = folder_dir & GetCorrectTxt(file_name_list, one_mvr.name, "manifesto")
            one_mvr.vendatxt = folder_dir & GetCorrectTxt(file_name_list, one_mvr.name, "venda")
            one_mvr.retornotxt = folder_dir & GetCorrectTxt(file_name_list, one_mvr.name, "retorno")
            mvr_list.Add one_mvr
        Next
        
        ' CREATE ALL WORKBOOKS
        Dim c_mvr As Object
        For Each c_mvr In mvr_list
            Dim wb As Workbook
            Dim wbAuto As Workbook
            Set wbAuto = ActiveWorkbook
            Set wb = Workbooks.Add
            wbAuto.Activate
            GenericMVR wb, wbAuto, c_mvr, "retorno"
            GenericMVR wb, wbAuto, c_mvr, "venda"
            GenericMVR wb, wbAuto, c_mvr, "manifesto"
            Dim strName As String
            strName = c_mvr.vendatxt
            strName = Split(strName, "\")(UBound(Split(strName, "\")))
            strName = Split(strName, ".txt")(0)
            strName = Replace(strName, "vendas", "MVR")
            wb.SaveAs strName
            wb.Close
        Next
        
    End If
    
    
    
    
End Sub


Function GetFolder() As String
    Application.DefaultFilePath = "D:\AutoMVR"
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

Function GetCorrectTxt(allNames As Variant, vendedor As String, operation As String) As String
    For Each name In allNames
        If Not InStr(name, vendedor) = 0 And Not InStr(name, operation) = 0 Then
            GetCorrectTxt = name
        End If
    Next
    
End Function

Sub GenericMVR(wb As Workbook, wbAuto As Workbook, mvr_p As MVR, operation As String)

    Dim operationtxt As String
    operation = operation ' & mvr_p.name
    operationtxt = operation & "txt"

    'grab formatation & base and paste
    wbAuto.Activate
    ActiveSheet.Range("A1:E450").Select
    Selection.Copy
    'set up wb
    wb.Activate
    If Sheets.Count > 1 Or Not StrComp(operation, "retorno", vbTextCompare) = 0 Then
        Sheets.Add
    End If
    ActiveSheet.Paste
    Columns("A:E").AutoFit
    ActiveSheet.name = operation
    Sheets.Add
    'import txt
    'MsgBox (mvr_p.manifestotxt)
    importToA1 mvr_p.AskDir(operation)
    ActiveSheet.name = operationtxt
    'WRITE FORMULAS
    'C
    Sheets(operation).Activate
    Range("C2").Select
    Selection.FormulaLocal = "=SEERRO(PROCV(A2;" & operationtxt & "!A$1:J$400;5;FALSO);0)"
    Selection.AutoFill Range("C2:C450")
    'E
    Sheets(operation).Activate
    Range("E2").Select
    Selection.FormulaLocal = "=SEERRO(PROCV(A2;" & operationtxt & "!A$1:J$400;7;FALSO);0)"
    Selection.AutoFill Range("E2:E450")
    'D
    Sheets(operation).Activate
    Range("D2").Select
    Selection.FormulaLocal = "=SEERRO(E2/C2;0)"
    Selection.AutoFill Range("D2:D450")
    
    Cells.Copy
    Range("A1").PasteSpecial Paste:=xlPasteValues
    
    Sheets(operationtxt).Delete
    Sheets(operation).Select
    Cells.Replace What:="#N/D", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="#N/D", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Range("C451").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-449]C:R[-1]C)"
    Range("D451").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-449]C:R[-1]C)"
    Range("E451").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-449]C:R[-1]C)"
    Columns("A:E").AutoFit
    Range("A1").Select
End Sub
