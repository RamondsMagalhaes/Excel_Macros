Attribute VB_Name = "filterdel"
Sub filterdel()
Attribute filterdel.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FILTERDEL Macro
'

'
    Windows("manif Thiago.xls").Activate
    Columns("D:D").Select
    Selection.AutoFilter
    ActiveSheet.Range("$D$1:$D$463").AutoFilter Field:=1, Criteria1:="-"
    ActiveWindow.SmallScroll Down:=-9
    Rows("10:464").Select
    Selection.Delete Shift:=xlUp
    Columns("D:D").Select
    ActiveSheet.ShowAllData
    Windows("maniFAST v1.0.xlsm").Activate
End Sub
