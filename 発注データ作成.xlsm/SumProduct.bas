Attribute VB_Name = "SumProduct"
Option Explicit

Sub SumPuchaseRequest()

'セラー分の商品コードを重複なしでコピー
'URL http://www.eurus.dti.ne.jp/yoneyama/Excel/vba/vba_jyufuku.html

Worksheets("セラー分").Range("C2:C48").AdvancedFilter xlFilterCopy, Unique:=True, CopyToRange:=Range("B2")

'手配件数を略号で入れる
Dim r As Range, CodeRange As Range, EndRow As Long
EndRow = Cells(1, 2).SpecialCells(xlLastCell).Row
Set CodeRange = Range(Cells(2, 2), Cells(EndRow, 2))

'商品毎の手配依頼数量の合計を算出

For Each r In CodeRange
    r.Offset(0, -1).Value = BuildMark(r.Value)
Next

'卸分の商品コードを入れる
Worksheets("卸分").Range("C2:C5").AdvancedFilter xlFilterCopy, Unique:=True, CopyToRange:=Range("B46")

Range("A2").End(xlDown).Offset(1, 1).Select
Range(Selection, Selection.End(xlDown)).Offset(0, -1).Select
Selection.Value = "V"
Range("A1").Select


'卸分の手配依頼数量の合計を算出


End Sub

Private Function BuildMark(ByVal Code As String) As String

Dim Counter(3) As Long, Mall As String, FoundCode As String, EndRow As Long, k As Long, Mark As String

EndRow = Worksheets("セラー分").UsedRange.Rows.Count

For k = 2 To EndRow
    
    FoundCode = Worksheets("セラー分").Cells(k, 3).Value
    
    If FoundCode = Code Then
    
        Mall = Worksheets("セラー分").Cells(k, 1).Value
        
        Select Case Mall
            Case "A"
                Counter(0) = Counter(0) + 1
            Case "R"
                Counter(1) = Counter(1) + 1
            Case "Y"
                Counter(2) = Counter(2) + 1
            Case Else
                Counter(3) = Counter(3) + 1
        End Select
    
    End If

Next

Mark = IIf(Counter(0) > 0, "A" & Counter(0), "")
Mark = Mark & IIf(Counter(1) > 0, "R" & Counter(1), "")
Mark = Mark & IIf(Counter(2) > 0, "Y" & Counter(2), "")
Mark = Mark & IIf(Counter(3) > 0, "SP" & Counter(3), "")

'このままだとA1というように、1件の時も数字が出るので、1は削除
Mark = Replace(Mark, "1", "")

BuildMark = Mark

End Function
