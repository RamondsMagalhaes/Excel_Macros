Attribute VB_Name = "autovp"
Sub vp_readAllTxt()
'
' vp_readAllTxt Macro
'

'

'OPEN DIR
    Dim fldr As FileDialog
    Dim exdr As FileDialog
    Dim dir As String
    Dim dirVP As String
    Dim wbVP As Workbook
    
    dirVP = Range("B1")
    dir = Range("B3")
    
    
    If dir <> "" And dirVP <> "" Then
        Dim oFSO As Object
        Dim oFolder As Object
        Dim oFile As Object
        
        'Reset Workplace
        Sheets("Relatorio de Faltas").Select
        DeleteAllSheets
        Rows("2:1048576").EntireRow.Hidden = False
        Rows("2:1048576").Select
        Selection.ClearContents
        
        'Load all Txt
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        Set oFolder = oFSO.GetFolder(dir)
        For Each oFile In oFolder.Files
            If UBound(Split(oFile.Name, ".txt")) > 0 Then
                Sheets.Add(After:=Sheets("Relatorio de Faltas")).Name = oFile.Name
                OneTxt (dir & "\" & oFile.Name)
            End If
            
        Next oFile
        'Open vp
        Set wbVP = Application.Workbooks.Open(dirVP)
        'Go back to auto
        Windows("Auto VP v2.0.xlsm").Activate
        ActiveWorkbook.Worksheets(2).Select
        'Read All sheets
        Do While ReadOneSheet(wbVP.Name)
            'Read all imported txt
        Loop
        'Green all missing
        Windows(wbVP.Name).Activate
        vp_green wbVP.Name
        Range("A1").Select
        'Go back
        Windows("Auto VP v2.0.xlsm").Activate
        ActiveWorkbook.Worksheets(2).Select
        DeleteAllSheets
    End If
    
End Sub

Private Function ReadOneSheet(vpFilename As String) As Boolean
    'Next Sheet
    If ActiveSheet.Index = Worksheets.Count Then
        'End
        'MsgBox ("Fim")
        ReadOneSheet = False
    Else
        ActiveSheet.Next.Activate
        Range("H2").Select
        vp_WriteAllInvoice (vpFilename)
        ReadOneSheet = True
    End If
End Function

Private Sub WriteMissingInvoice(Optional ByVal isGreen As Boolean = False, Optional ByVal vpFilename As String = "none", Optional ByVal vendedor As String = "none")
    Dim cSheet As Integer
    Dim cRow As Integer
    Dim nRow As Integer
    
    cSheet = ActiveSheet.Index
    cRow = Selection.row
    If isGreen Then
        Selection.Copy
        Windows("Auto VP v2.0.xlsm").Activate
        Sheets("Relatorio de Faltas").Select
    Else
        Range(cRow & ":" & cRow).Copy
        Sheets("Relatorio de Faltas").Select
    End If
    Range("A1").Select
    'Find next empty line
    'Do While Not IsEmpty(Selection.Value)
        Selection.Offset(1, 0).Select
    'Loop
    nRow = Cells(Rows.Count, 1).End(xlUp).row + 1
    If isGreen Then
        Range("B" & nRow).PasteSpecial Paste:=xlPasteValues
        Range("A" & nRow).Value = vendedor
        Windows(vpFilename).Activate
    Else
        Range(nRow & ":" & nRow).PasteSpecial Paste:=xlPasteValues
        'ActiveSheet
        Range("A" & nRow).Value = Worksheets(cSheet).Name
        Application.Worksheets(cSheet).Select
    End If

    
End Sub

Private Sub OneTxt(filename As String)
'
' vpToExcel Macro
'

'
    Dim vendedor As String
    
    ' SET VENDEDOR
    'remove .txt
    vendedor = Split(filename, ".txt")(0)
    'split at \, get last split
    vendedor = Split(vendedor, "\")(UBound(Split(vendedor, "\"), 1))
    'split again, remove last
    vendedor = Replace(vendedor, Split(vendedor)(UBound(Split(vendedor), 1)), "")
    
    WriteTxtToExcel (filename)
    

End Sub


Private Function WriteTxtToExcel(filename As String)
    Rows("2:1048576").Select
    Selection.ClearContents
    Range("B2").Select
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & filename, Destination:=Range("$B$2"))
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
        .TextFileStartRow = 7
        .TextFileParseType = xlFixedWidth
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, xlTextFormat, 1, 1, 1, xlDMYFormat, 1)
        'Nº do banco / Nº Cliente / Nome / Nº Documento / Data Emissao / Data Vencimento / Valor
        .TextFileFixedColumnWidths = Array(21, 11, 52, 10, 17, 10, 14)
        .TextFileTrailingMinusNumbers = False
        .Refresh BackgroundQuery:=False
    End With
    Cells.Select
    Cells.EntireColumn.AutoFit
'   Range("B2").Select
'
'   Sheets("A").Select
'   Range("C3").Select
'   ActiveWorkbook.Save
End Function

Private Sub DeleteAllSheets()

    Application.DisplayAlerts = False 'switching off the alert button
    

    Do While ActiveWorkbook.Worksheets.Count > 2
        ActiveWorkbook.Worksheets(3).Delete
    Loop
    
    'ActiveSheet.Delete
    
    Application.DisplayAlerts = True 'switching on the alert button

End Sub


Private Function vp_WriteAllInvoice(vpFilename As String) As Boolean
Attribute vp_WriteAllInvoice.VB_ProcData.VB_Invoke_Func = " \n14"
'
' vp_WriteAllSheets Macro
'

'
    Do While IsNumeric(Selection.Value)
        OneInvoice (vpFilename)
        Selection.Offset(GetNextRoll, 0).Select
    Loop
    
End Function

Private Function OneInvoice(vpFilename As String)
'
' blank Macro
'

'
    Dim invoice As clsNF
    Dim search As Variant
    Dim txtDumpFilename As String

    txtDumpFilename = "Auto VP v2.0.xlsm"
    
    Set invoice = SetInvoice
    
    search = invoice.valor
    
'    MsgBox invoice.valor
'    MsgBox invoice.GetDate
'    MsgBox invoice.number
    
    Windows(vpFilename).Activate
    If (CustomFinder(invoice)) Then
        'Write Values
        Selection.Offset(0, 1).Value = invoice.GetDate
        Selection.Offset(0, 1).NumberFormat = "dd/mmm"
        Selection.Offset(0, 2).NumberFormat = "@"
        Selection.Offset(0, 2).Value = invoice.number
        'Text Format
        Selection.Offset(0, 2).NumberFormat = "@"
    Else
        'Write not found invoice
        Windows(txtDumpFilename).Activate
        WriteMissingInvoice '(invoice)
    End If

    Windows(txtDumpFilename).Activate
    
End Function


Private Function SetInvoice() As clsNF
    'If this Nº Nota is equal to bellow
    Dim rInvoice As New clsNF
    Dim total As Variant
    Do While isSameInvoice()
        total = total + Selection.Value
        rInvoice.addDate (Selection.Offset(0, -1).Value)
        Selection.Offset(GetNextRoll(), 0).Select
    Loop
    total = total + Selection.Value
    rInvoice.addDate (Selection.Offset(0, -1).Value)
    rInvoice.number = Selection.Offset(0, -3).Value
    rInvoice.valor = total
    Set SetInvoice = rInvoice
End Function

Private Function GetNextRoll() As Integer
    'Check for the token in column I
    If Not IsEmpty(Selection.Offset(1, 1)) And IsNumeric(Selection.Offset(1, 1)) Then
        'Skip 7
        GetNextRoll = 7
    Else
        'Skip 1
        GetNextRoll = 1
    End If
End Function

Private Function isSameInvoice() As Boolean
    Dim cNumber As Variant
    Dim nNumber As Variant
    cNumber = Selection.Offset(0, -3).Value
    nNumber = Selection.Offset(GetNextRoll(), -3).Value
    isSameInvoice = (cNumber = nNumber)
End Function

Private Function CustomFinder(invoice As clsNF) As Boolean
    Dim isValid As Boolean
    Dim tries As Integer
    
    isValid = False
    tries = 0
    Range("A1").Select
    Do While Not isValid And tries < 5
        On Error GoTo err_handle
            Finder (invoice.valor)
        'If "found value AND its new" OR If "found value and current meta data is equal to invoice"
        If (IsEmpty(Selection.Offset(0, 1)) And IsEmpty(Selection.Offset(0, 2))) Or _
        (Selection.Offset(0, 1).Value = invoice.GetDate And Selection.Offset(0, 2).Value = invoice.number) Then
            isValid = True
        End If
        
        tries = tries + 1
    Loop
    
    CustomFinder = isValid
    Exit Function
    
err_handle:
            'Every false return is cataloged in blank row 11
            CustomFinder = False
            Exit Function
    
    
    
End Function

Private Function Finder(search As Variant)
        
    Cells.Find(What:=search, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
    :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    False, SearchFormat:=False).Activate
    
End Function
''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''' GREEN '''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''

Sub vp_green(vpFilename As String)
'
' FixVP Macro
'
' Atalho do teclado: Ctrl+a
'
    Dim allColumns As Variant
    allColumns = Array("B", "F", "J", "N", "R", "V", "Z", "AD", "AH", "AL", "AP", "AT", "AX", "BB", "BE")
    For Each i In allColumns
        Range(i & "1").Select
        GreenWholeColumn vpFilename
    Next i
    
    
    

End Sub

Private Sub GreenWholeColumn(vpFilename As String)
    Dim col As Integer
    Dim row As Integer
    col = Selection.Column
    row = Selection.row
    Do While row < 583
        col = Selection.Column
        row = Selection.row
        GreenOne row, col, vpFilename
        Cells(row + 1, col).Select
    Loop
End Sub

Private Sub GreenOne(row As Integer, col As Integer, vpFilename As String)
    ' If Value and Name but no date and no number
    If (IsNumeric(Selection.Value) And Not IsEmpty(Selection.Value)) And Not IsEmpty(Selection.Offset(0, -1)) And _
    IsEmpty(Selection.Offset(0, 1)) And IsEmpty(Selection.Offset(0, 2)) Then
        Range(Cells(row, col - 1), Cells(row, col + 2)).Select
        With Selection.Interior
            .PatternColorIndex = xlAutomatic
            .Color = 5296274
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        'Data
        Cells(row, col + 1).NumberFormat = "[$-pt-BR]dd-mmm;@"
        'Number
        Cells(row, col + 2).NumberFormat = "@"
        'WriteMissingInvoice True, vpFilename, Cells(1, col - 1).Value
    ' If Value, Name and date and number
    ElseIf (IsNumeric(Selection.Value) And Not IsEmpty(Selection.Value)) And Not IsEmpty(Selection.Offset(0, -1)) And _
        Not IsEmpty(Selection.Offset(0, 1)) And Not IsEmpty(Selection.Offset(0, 2)) Then
        'paint white
        Cells(row, col - 1).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Range(Cells(row, col + 1), Cells(row, col + 2)).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        'Data
        Cells(row, col + 1).NumberFormat = "[$-pt-BR]dd-mmm;@"
        'Number
        Cells(row, col + 2).NumberFormat = "@"
        'WriteMissingInvoice True, vpFilename, Cells(1, col - 1).Value
        'paint yellow
        Cells(row, col).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

    End If
    
    Cells(row, col).Select
    
End Sub
