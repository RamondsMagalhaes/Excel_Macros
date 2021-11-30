VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AutoStockForm 
   Caption         =   "Lançador de Produtos"
   ClientHeight    =   2310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7020
   OleObjectBlob   =   "AutoStockForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AutoStockForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
    Dim side As Integer
    If OptionButton1.Value = True Then
        side = 1
    Else
        side = 455
    End If
    On Error Resume Next
    Cells(side, CInt(ComboBox1.column(1))).Select
    colunaTarget.Caption = ComboBox1.Value
End Sub

Private Sub CommandButton1_Click()
    writeAt
    TextBox1.Value = ""
    TextBox2.Value = ""
    TextBox3.Value = ""
    TextBox1.SetFocus
    
End Sub


Private Sub CommandButton3_Click()
    Unload Me
End Sub


Private Sub CommandButton5_Click()
    Cells(2, CInt(ComboBox1.column(1))).Select
    For Each rng In Range(Cells(2, CInt(ComboBox1.column(1))), Cells(450, CInt(ComboBox1.column(1))))
        If Not IsEmpty(rng) Then
            rng.Select
            Dim answer As Integer
            Dim text As String
            If OptionButton1.Value = True Then
                'Estoque
                text = Cells(rng.row, 2).Value
            Else
                'C. Fria
                text = Cells(rng.row, 1).Value
            End If
            text = text & vbNewLine & vbTab & rng.Value
            answer = MsgBox(text, vbYesNo + vbDefaultButton1)
            If answer = vbYes Then
                rng.Interior.ColorIndex = 6
            Else
                rng.Interior.ColorIndex = 3
            End If
        End If
    Next
    
End Sub

Private Sub OptionButton1_Click()
    SetEstoque
End Sub

Private Sub OptionButton2_Click()
    SetCFria
End Sub


Private Sub up_Click()
    MoveUpDown (True)
End Sub

Private Sub down_Click()
    MoveUpDown (False)
End Sub

Private Function MoveUpDown(isUp As Boolean)
    If isUp Then
        If Selection.row > 2 Then
            Selection.Offset(-1, 0).Select
        End If
    Else
        Selection.Offset(1, 0).Select
    End If
    Dim code As String
    Dim row As Integer
    Dim curWB As String
    curWB = ActiveWorkbook.name
    row = Selection.row
    Workbooks("Stock3000.xlsm").Sheets("base").Activate
    code = Cells(row, 1).Value
    Workbooks(curWB).Activate
    TextBox1.Value = code
    produtoTarget.Caption = Cells(row, 1).Offset(0, 1).Value
    UpdateQtt
End Function



Private Sub UserForm_Initialize()

    'Set buttom
    OptionButton1.Value = True
    'SetEstoque
    

End Sub

Private Sub TextBox1_AfterUpdate()
    Dim target As Range
    Set target = FindRange
    Cells(target.row, CInt(ComboBox1.column(1))).Select
    UpdateQtt
End Sub


Private Function FindRange() As Range
    Dim rng As Range
    Dim err As Boolean
    err = False
    If TextBox1.Value = "" Then
        Set rng = Selection
        'produtoTarget.Caption = "Digite o codigo"
    Else
        Dim curWB As String
        curWB = ActiveWorkbook.name
        Workbooks("Stock3000.xlsm").Sheets("base").Activate
        Columns("A:A").Select
        With Worksheets(1).Range(Cells.Address)
            If IsNumeric(TextBox1.Value) Then
                Set rng = .find(What:=TextBox1.Value, After:=ActiveCell, LookIn:=xlFormulas, _
                    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                    MatchCase:=False, SearchFormat:=False)
            Else
                'FIND NEXT
                Set rng = Worksheets(1).Range("B:B").find _
                    (What:=TextBox1.Value, After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
                    , SearchFormat:=False).Activate
            End If
            If rng Is Nothing Then
                MsgBox ("Codigo invalido")
                Set rng = Selection
                TextBox1.Value = ""
                TextBox1.SetFocus
                err = True
            End If
        End With
        If Not err Then
            produtoTarget.Caption = Cells(rng.row, 2).Value
        End If
        Workbooks(curWB).Activate
        If err Then TextBox1.SetFocus
    End If
    Set FindRange = rng
End Function

Private Function SetEstoque()
    Do While ComboBox1.ListCount > 0
        ComboBox1.RemoveItem (0)
    Loop

    'Find Range
    Dim rngV As Variant
    Dim rngS As Variant
    For Each rng In Range("C1:AAA1")
        rngV = rng.Value
        If (StrComp(rngV, "Saída") = 0) Or (rngV = "") Then
            rngV = rng.column - 1
            Exit For
        End If
    Next
    'Find Saída
    rngS = "fail"
    For Each rng In Range("C1:AAA1")
        rngS = rng.Value
        If (StrComp(rngS, "Saída") = 0) Then
            rngS = rng.column + 1
            Exit For
        End If
    Next
    If StrComp(rngS, "fail") = 0 Then
        rngS = 30
    End If
    displabelTarget.Caption = rngS
    'Cells(1, rngV).Select
    'Set
    For Each rng In Range(Cells(1, 3), Cells(1, rngV))
        ComboBox1.AddItem rng.Value
        ComboBox1.list(ComboBox1.ListCount - 1, 1) = rng.column
    Next
    ComboBox1.ListIndex = ComboBox1.ListCount - 1
    TextBox3.TabStop = False
End Function

Private Function SetCFria()
    Do While ComboBox1.ListCount > 0
        ComboBox1.RemoveItem (0)
    Loop
    For Each rng In Range("C455:AAA455")
        If Not IsEmpty(rng.Value) And Not IsNumeric(rng.Value) And Not IsError(rng.Value) Then
            rng.Select
            ComboBox1.AddItem rng.Value
            ComboBox1.list(ComboBox1.ListCount - 1, 1) = rng.column
        End If
    Next
    ComboBox1.ListIndex = ComboBox1.ListCount - 1
    TextBox3.TabStop = True
End Function

Private Function UpdateQtt()
    atualTarget = Selection.formula
    If OptionButton1.Value = True Then
        If Not displabelTarget.Caption = "" Then
            disponivelTarget = Cells(Selection.row, CInt(displabelTarget.Caption)).Value
        Else
            'MsgBox ("Algo esta errado, você selecionou a opção C.Fria / Estoque corretamente?")
        End If
    Else
        disponivelTarget = "-"
    End If
    
End Function

Private Function writeAt()
    Selection.Value = TextBox2.Value
    If OptionButton2.Value = True Then
        Selection.Offset(0, 1).Value = TextBox3.Value
    End If
End Function

