Attribute VB_Name = "extra"

Sub quickEXE()
Attribute quickEXE.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' quickEXE Macro
'
' Atalho do teclado: Ctrl+e
'
    Copia_Base_Faturada
    Dim formula As String
    Dim lastRow As Long
    Dim lastBFRow As Long
    Dim lastARow As Long
    Dim totalProdutosA As Double
    Dim totalProdutosBF As Double
    
    Sheets("BF").Select
    procvFormula = "=PROCV(B2;A!A$3:H$451;4;FALSO)"
    lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    Range("O2").FormulaLocal = procvFormula
    Range("O2").Select
    Selection.AutoFill Destination:=Range("O2:O" & lastRow)
    For i = 2 To lastRow
        If Not IsNumeric(Range("B" & i)) Or IsEmpty(Range("B" & i)) Then
        Range("O" & i).ClearContents
        End If
        
    Next i
    lastBFRow = Cells(Rows.Count, 6).End(xlUp).Row
    totalProdutosBF = Range("F" & lastBFRow).Value
    Sheets("A").Select
    lastARow = Cells(Rows.Count, 4).End(xlUp).Row
    totalProdutosA = Range("D" & lastARow).Value
    If totalProdutosBF <> totalProdutosA Then
        MsgBox "ERRO! Total de Produtos NÃO bate"
        Sheets("BF").Select
    Else
        MsgBox "CORRETO! Total de Produtos batem"
        Sheets("A").Select
    End If
    
    'Sheets("A").Select
    'Range("H3:H451").Select
    'Selection.Copy
End Sub
