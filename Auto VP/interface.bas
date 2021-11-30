Attribute VB_Name = "Interface"
Public isCancel As Boolean

Sub vp_ChooseDir()

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
    If dirVP <> "" Then Cells(1, 2).value = dirVP
    
    
End Sub

Sub reports_ChooseDir()
    Dim fldr As FileDialog
    Dim dir As String
    
    
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
    If dir <> "" Then
        Cells(3, 2).value = dir
        
        Range("C3:L4").ClearContents
        
        'WRITE FILE NAMES
        Dim oFSO As Object
        Dim oFolder As Object
        Dim oFile As Object
        Dim Index As Integer
                  
        Index = 0
        'Load one xlsx at a time
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        Set oFolder = oFSO.GetFolder(dir)
        For Each oFile In oFolder.files
            'If found File is xlsx
            If UBound(Split(oFile.name, ".xlsx")) > 0 Then
                'Write name
                Cells(3, 3 + Index).value = oFile.name
                Cells(4, 3 + Index).value = "Relatorio " & (Index + 1)
                Index = Index + 1
            End If
            
        Next oFile
    End If
End Sub

Sub onlyOne()
    Dim c_file As String
    isCancel = False
    pick.Show
    If Not isCancel Then
        c_file = pick.ComboBox1.value
        InsertBank Array(c_file)
        Unload pick
        ChangeColor False
    End If
    
End Sub

Sub sendAll()
    Dim files As Object
    Set files = CreateObject("System.Collections.ArrayList")
    For Each cell In Range("C3:L3").SpecialCells(xlConstants)
        files.Add (cell.value)
    Next cell
    InsertBank files
    ChangeColor False
End Sub

Sub FixVPButton()
    mainvpfix (True)
    ThisWorkbook.Activate
    ChangeColor (True)
End Sub

Sub FixVPButtonFalse()
    mainvpfix (False)
    ThisWorkbook.Activate
    ChangeColor (True)
End Sub

Sub OpenVPButton()
    Workbooks.Open (Range("B1").value)
End Sub

Sub ChangeColor(isOK As Boolean)
    Sheets("main").Activate
    If isOK Then
        'GREEN
        Range("B7").Select
        ActiveCell.FormulaR1C1 = "Não Falta Organizar"
        Range("B7:B8").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
    Else
        'RED
        Range("B7").Select
        ActiveCell.FormulaR1C1 = "Falta Organizar"
        Range("B7:B8").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If
End Sub


