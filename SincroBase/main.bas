Attribute VB_Name = "main"
Sub mainAtt()
    Dim wb, wbSinc As Workbook
    Dim basedRange As Range
    Dim basedepartment As Object
    Set basedepartment = CreateObject("System.Collections.ArrayList")
    
    For Each cell In Range("D7:D999")
        If Not IsEmpty(cell) Then
            basedepartment.Add cell.Value
        End If
    Next cell
    
    Set wbmaster = Workbooks.Open(Range("D1").Value)
    
    For Each basedir In basedepartment
        'open target base
        Application.DisplayAlerts = False
        Set wb = Workbooks.Open(basedir)
        wb.SaveAs Filename:="backup\BACKUP " & wb.Name, FileFormat:=xlWorkbookNormal
        wb.Close
        Set wb = Workbooks.Open(basedir)
        wbmaster.Activate
        Dim control As Boolean
        control = False
        Do
            'GETTING CORE BASE, IF NULL ABORT
            wbmaster.Activate
            Set basedRange = findBase(control)
            If Not basedRange Is Nothing Then
                basedRange.Select
                Selection.Copy
                wb.Activate
                Dim targetbase As Range
                Set targetbase = findBase(control)
                'IF NULL, DO NOTHING
                If Not targetbase Is Nothing Then
                    targetbase.PasteSpecial (xlPasteValues)
                End If
            Else
                Exit Sub
            End If
            control = Not control
        Loop While control
        wb.Save
        wb.Close
        Application.DisplayAlerts = True
    Next basedir
    wbmaster.Close
End Sub

Function findBase(isCode As Boolean) As Range
    Dim findstr As String
    Dim uprange As Range
    If isCode Then
        findstr = "206167"
    Else
        findstr = "Polpa Iogurte Bi Sabor 540g"
    End If
    
    Cells.Select
    Set uprange = Cells.Find(What:=findstr, After:=ActiveCell, LookIn _
        :=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False)
    If uprange Is Nothing Then
        Set findBase = Nothing
    Else
        Set findBase = Range(uprange, uprange.Offset(448, 0))
    End If

End Function

Function findCode() As Range

End Function
