Attribute VB_Name = "import2"
Sub wholetramit()
Attribute wholetramit.VB_ProcData.VB_Invoke_Func = " \n14"
'
' wholetramit Macro
'

'
    Windows("AutoMVR.xlsm").Activate
    Cells.Select
    Selection.Copy
    ''CREATE
    Windows("Pasta1").Activate
    ActiveSheet.Paste
    Sheets("Planilha1").Select
    Sheets("Planilha1").name = "manifesto"
    Sheets.Add After:=ActiveSheet
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "manifest"
    Sheets("Planilha2").Select
    Sheets("Planilha2").name = "manifesto txt"
    Selection.ClearContents
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;D:\AutoMVR\01 0301 0801\Adao MANIFESTO 0201 0801.txt", Destination:= _
        Range("$A$1"))
        .CommandType = 0
        .name = "Adao MANIFESTO 0201 0801"
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
        .TextFileFixedColumnWidths = Array(12, 38, 9, 3, 8, 7, 8, 11, 5)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    ActiveWindow.SmallScroll Down:=-39
    Sheets("manifesto").Select
    Range("C2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-2],'manifesto txt'!R[5]C[-2]:R[216]C[7],4,FALSE)"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-2],'manifesto txt'!R[5]C[-2]:R[216]C[7],5,FALSE)"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C9")
    Range("C2:C9").Select
    Selection.AutoFill Destination:=Range("C2:C450"), Type:=xlFillDefault
    Range("C2:C450").Select
    ActiveWindow.SmallScroll Down:=-465
End Sub
