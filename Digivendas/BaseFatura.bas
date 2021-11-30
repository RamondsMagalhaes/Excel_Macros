Attribute VB_Name = "BaseFatura"
Sub Copia_Base_Faturada()
'
' Esta Macro faz a cópia da Base faturada do vendedor para a base do Excel
'
    Sheets("BF").Select
    Rows("2:1048576").Select
    Selection.ClearContents
    Range("B2").Select
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;D:\Video.txt", Destination:=Range("$B$2"))
        .Name = "Video"
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
        .TextFileStartRow = 8
        .TextFileParseType = xlFixedWidth
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileFixedColumnWidths = Array(10, 40, 15, 13, 11, 11, 10, 12, 11, 8, 11, 8)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("B2").Select

    Sheets("A").Select
    Range("C3").Select
    ActiveWorkbook.Save
    
End Sub
