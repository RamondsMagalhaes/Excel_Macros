Attribute VB_Name = "interface"
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
    If dirVP <> "" Then Cells(1, 2).Value = dirVP
    
    
End Sub

Sub reports_ChooseDir()
    Dim fldr As FileDialog
    Dim dir As String
    
    
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    ' Set .txt Folder
    With fldr
        .Title = "Escolha a Pasta com as remessas do banco .txt"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path
        If .Show <> -1 Then GoTo NextStepOne
        dir = .SelectedItems(1)
    End With
NextStepOne:
    'GetFolder = dir
    Set fldr = Nothing
    If dir <> "" Then
        Cells(3, 2).Value = dir
        
        Range("C3:Z4").ClearContents
        
        'WRITE FILE NAMES
        Dim oFSO As Object
        Dim oFolder As Object
        Dim oFile As Object
        Dim Index As Integer
                  
        Index = 0
        'Load one xlsx at a time
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        Set oFolder = oFSO.GetFolder(dir)
        For Each oFile In oFolder.Files
            'If found File is xlsx
            If UBound(Split(oFile.Name, ".txt")) > 0 Then
                'Write name
                Cells(3, 3 + Index).Value = oFile.Name
                Cells(4, 3 + Index).Value = "Remessa " & (Index + 1)
                Index = Index + 1
            End If
            
        Next oFile
    End If
End Sub

Sub interface_vpfix()
    Dim wbVP As Workbook
    Set wbVP = Application.Workbooks.Open(Range("B1").Value)
    Windows(wbVP.Name).Activate
    vp_green (wbVP.Name)
End Sub


Public Sub open_vpdir()
    Dim wbVP As Workbook
    Set wbVP = Application.Workbooks.Open(Range("B1").Value)
    Windows(wbVP.Name).Activate
End Sub
