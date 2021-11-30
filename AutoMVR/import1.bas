Attribute VB_Name = "import1"
Public Sub importToA1(dir As String)
Attribute importToA1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' improtbloco Macro
'

'
    Range("A1").Select
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & dir, Destination:= _
        Range("$A$1"))
        .name = "AutoMVR"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlFixedWidth
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileFixedColumnWidths = Array(12, 38, 9, 3, 8, 6, 9, 11, 3)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Range("A1").Select
End Sub
Sub ImportMan()
Attribute ImportMan.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ImportMan Macro
'

'

End Sub
