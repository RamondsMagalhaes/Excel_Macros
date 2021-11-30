Attribute VB_Name = "vpfix"
Public Sub mainvpfix(bool As Boolean)
'
' FixVP Macro
'
' Atalho do teclado: Ctrl+a
'
    Dim allColumns As Variant
    Dim wbName As String
    wbName = Range("B1").value
    wbName = Split(wbName, "\")(UBound(Split(wbName, "\"), 1))
    Dim isOpen As Boolean
    For Each wb In Workbooks
    If wb.name = wbName Then
        isOpen = True
    End If
    Next
    If isOpen Then
        'obsolet
    ElseIf Not isOpen And (wbName = "a" Or wbName = "") Then
        'obsolet
    Else
        Workbooks.Open (Range("B1").value)
    End If
    ThisWorkbook.Activate
    Dim columnsString As String
    columnsString = ""
    Range("C8").Select
    While (Not IsEmpty(Selection.value))
        columnsString = columnsString & Selection.value & ","
        Selection.Offset(0, 1).Select
    Wend
    Windows(wbName).Activate
    allColumns = Split(columnsString, ",")
    For Each i In allColumns
        If (i <> "") Then
            Range(i & "1").Select
            ClearWholeColumn (bool)
            Range(i & "1").Select
            ClearAgain
        End If
    Next i
    
    Range("BG24").Select

End Sub

Private Sub RemoveOne(row As Integer, col As Integer)
    Range(Cells(row, col - 1), Cells(row, col + 2)).Select
    Selection.ClearContents
    Range(Cells(row + 1, col - 1), Cells(583, col + 2)).Select
    Selection.Cut
    Cells(row, col - 1).Select
    ActiveSheet.Paste
    
    Range(Cells(582, col - 1), Cells(582, col + 2)).Select
    Selection.Copy
    Range(Cells(583, col - 1), Cells(583, col + 2)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Cells(row, col).Select
End Sub

Private Sub PaintYellow(row As Integer, col As Integer)
    If (IsNumeric(Selection.value) And Not IsEmpty(Selection.value)) And Not IsEmpty(Selection.Offset(0, -1)) And _
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

End Sub

Private Sub FixFormat(row As Integer, col As Integer)
    'If date
    If (Not IsEmpty(Selection.Offset(0, -1)) And IsEmpty(Selection.value) And IsEmpty(Selection.Offset(0, 1)) And IsEmpty(Selection.Offset(0, 2))) Then
        Selection.Offset(0, -1).NumberFormat = "[$-pt-BR]d-mmm-yy;@"
    'If roll
    ElseIf (IsNumeric(Selection.value) And Not IsEmpty(Selection.value)) And Not IsEmpty(Selection.Offset(0, -1)) And _
        Not IsEmpty(Selection.Offset(0, 1)) And Not IsEmpty(Selection.Offset(0, 2)) Then
        'nome
        Selection.Offset(0, -1).NumberFormat = "@"
        
        'valor
        'Selection.NumberFormat = "0.00"
        
        'data
        Selection.Offset(0, 1).NumberFormat = "[$-pt-BR]d-mmm;@"
        
        'numero
        With Selection.Offset(0, 2)
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Offset(0, 2).NumberFormat = "@"
        
    End If
End Sub

Private Sub SendToBottom(row As Integer, col As Integer, bottomRow As Integer)
    'Send to Bottom
    
    'Get and cut negative line
    Range(Cells(row, col - 1), Cells(row, col + 2)).Select
    Selection.Copy
    'Paste at bottom row
    Cells(bottomRow, col - 1).Select
    ActiveSheet.Paste
    'Write zero at removed line, so the macro will clear it
    Cells(row, col).value = 0
    'Dont skip, the current cell is the next target
    Cells(row, col).Select
End Sub

Private Sub ClearWholeColumn(bool As Boolean)
    Dim col As Integer
    Dim row As Integer
    Dim bottomRow As Integer
    Dim answer As Integer
    col = Selection.Column
    row = Selection.row
    bottomRow = FindBottom(row, col)
    Do While row < 583
        col = Selection.Column
        row = Selection.row
        If bool Then PaintYellow row, col
        'If value is zero OR if double dates
        If (Cells(row, col).value = 0 And Not IsEmpty(Cells(row, col))) Or _
        ((IsEmpty(Cells(row, col).value) And Not IsEmpty(Cells(row, col - 1).value)) And (IsEmpty(Cells(row + 1, col).value) And Not IsEmpty(Cells(row + 1, col - 1).value))) Then
            RemoveOne row, col
        ElseIf (0 < Cells(row, col).value) And (Cells(row, col).value <= 3) Then
            SendToBottom row, col, bottomRow
            bottomRow = bottomRow + 1
        ElseIf Cells(row, col).value < 0 Then
            If InStr(1, Cells(row, col - 1).value, "credito", vbTextCompare) Or _
            InStr(1, Cells(row, col - 1).value, "crédito", vbTextCompare) Then
                answer = False
            Else
                answer = True
            End If
            'answer = 'MsgBox("Enviar esse boleto para baixo?: " & vbNewLine & Cells(row, col - 1).value & vbNewLine & vbNewLine & "Nomes com CREDITO não devem ser enviados para baixo", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
            If answer Then
                SendToBottom row, col, bottomRow
                bottomRow = bottomRow + 1
            Else
                Cells(row + 1, col).Select
            End If
        Else
            Cells(row + 1, col).Select
        End If
    Loop
End Sub

Private Sub ClearAgain()
    Dim col As Integer
    Dim row As Integer
    Dim bottomRow As Integer
    Dim answer As Integer
    col = Selection.Column
    row = Selection.row
    'bottomRow = FindBottom(row, col)
    Do While row < 583
        col = Selection.Column
        row = Selection.row
        FixFormat row, col
        'if double dates
        If ((IsEmpty(Cells(row, col).value) And Not IsEmpty(Cells(row, col - 1).value)) And (IsEmpty(Cells(row + 1, col).value) And Not IsEmpty(Cells(row + 1, col - 1).value))) Then
            RemoveOne row, col
        Else
            Cells(row + 1, col).Select
        End If
    Loop
End Sub

Private Function FindBottom(row As Integer, col As Integer) As Integer

    Dim bottomRow As Integer
    bottomRow = 585
    Do While Not (IsEmpty(Cells(bottomRow, col)) And IsEmpty(Cells(bottomRow + 1, col)))
        bottomRow = bottomRow + 1
    Loop
    'Get correct bottom row
    bottomRow = bottomRow + 1
    
    FindBottom = bottomRow

End Function


