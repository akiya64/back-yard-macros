Attribute VB_Name = "UpdateDisallowFlag"
Option Explicit
Sub 不可事由転記()

Dim Arr As Variant
Arr = Array("廃番", "不可", "不明")

'JANで、不可事由が空欄をまずフィルター
With Worksheets("商品情報").Range("A1").CurrentRegion
    '.AutoFilter Field:=2, Criteria1:="???????*"
    .AutoFilter Field:=16, Criteria1:="="
End With

'設定した文言それぞれで、フィルターして手配不可事由へ記入を実行
Dim s As Variant
For Each s In Arr
    Call InputReason(s)
Next

Worksheets("商品情報").Range("A1").AutoFilter


End Sub

Sub InputReason(ByVal Str As String)
Attribute InputReason.VB_ProcData.VB_Invoke_Func = " \n14"
'廃番フラグの未転記をフィルター

Worksheets("商品情報").Range("A1").CurrentRegion.AutoFilter Field:=4, Criteria1:="*" & Str & "*"

Dim r As Range, TargetRange As Range
Set TargetRange = Intersect(Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible), Range("P2:P300000"))

If TargetRange Is Nothing Then Exit Sub

For Each r In TargetRange
    r.Offset(0, -1).Value = 1
    r.Offset(0, 1).Value = Date
    r.Value = Str
Next

End Sub
