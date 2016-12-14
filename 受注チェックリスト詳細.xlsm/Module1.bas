Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$EB$936").AutoFilter Field:=2, Criteria1:="<>"
End Sub
