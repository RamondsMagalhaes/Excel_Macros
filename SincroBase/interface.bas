Attribute VB_Name = "interface"
Public Sub chooseMain()
    Dim file As String
    file = GetFile
    If file <> "" Then
        Range("D1").Value = file
    End If
End Sub

Public Sub showUserform()
    based.Show
End Sub

Function GetFile() As String
    Dim exdr As FileDialog
    Dim dirVP As String
    
    'Set VP dir
    Set exdr = Application.FileDialog(msoFileDialogFilePicker)
    With exdr
        .Title = "Escolha a Planilha BASE CORRETA"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path
        If .Show <> -1 Then GoTo NextStepTwo
        dirVP = .SelectedItems(1)
    End With
NextStepTwo:
    'GetFolder = dirVP
    Set exdr = Nothing
    GetFile = dirVP
    
    
End Function
