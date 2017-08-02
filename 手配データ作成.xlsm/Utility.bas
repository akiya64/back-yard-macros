Attribute VB_Name = "Utility"
Option Explicit

Sub PrepareSheet(ByRef TargetSheet As Variant)
'シートのクリア

If Not IsEmpty(TargetSheet.Range("A2").Value) Then
    TargetSheet.Range("A1").CurrentRegion.Offset(1, 0).Delete Shift:=xlShiftUp
    
    If TargetSheet.Buttons.Count > 0 Then
        TargetSheet.Buttons(1).Delete
    End If

End If

End Sub

Function FetchWorkBook(path As String) As Workbook

'引数で渡されたパスのブックを開きます。引数のブックを開いていれば、そのブックを戻り値にします。

Dim WorkBookName As String
WorkBookName = Dir(path)

Dim wb As Workbook

For Each wb In Workbooks
    
    If wb.Name = WorkBookName Then
        
        GoTo ret
    
    End If

Next

Set wb = Workbooks.Open(path)

ret:
Set FetchWorkBook = wb

End Function

