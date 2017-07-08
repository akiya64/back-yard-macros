Attribute VB_Name = "SumPurcheseReq"
Option Explicit

Sub SumPuchaseRequest()
'商品別に手配依頼数を集計

Worksheets("手配数量決定シート").Activate

'セラー分の商品コードを重複なしでコピー
'URL http://www.eurus.dti.ne.jp/yoneyama/Excel/vba/vba_jyufuku.html
Worksheets("セラー分").Range("C2:D" & Worksheets("セラー分").UsedRange.Rows.Count).AdvancedFilter xlFilterCopy, CriteriaRange:=Range("C2:C" & Worksheets("セラー分").UsedRange.Rows.Count), Unique:=True, CopyToRange:=Range("G2")

'手配件数の合計略号と、手配依頼数の合計を入れる
Dim r As Range, CodeRange As Range, Endrow As Long
Endrow = Cells(1, 7).SpecialCells(xlLastCell).Row
Set CodeRange = Range(Cells(2, 7), Cells(Endrow, 7))

For Each r In CodeRange
    r.Offset(0, -1).Value = CountMallOrder(r.Value)
    r.Offset(0, 2).Value = SumRequestQuantity(r.Value, True)
Next

'卸分の手配依頼数量の合計を算出
'卸分の商品コードを入れる
Worksheets("卸分").Range("C2:D" & Worksheets("卸分").UsedRange.Rows.Count).AdvancedFilter xlFilterCopy, CriteriaRange:=Range("C2:C" & Worksheets("セラー分").UsedRange.Rows.Count), Unique:=True, CopyToRange:=Range("G" & Endrow + 1)

Dim WholeSaleJanRange As Range
Set WholeSaleJanRange = Range(Cells(Endrow + 1, 7), Cells(2, 7).End(xlDown))

WholeSaleJanRange.Offset(0, -1).Value = "V"

For Each r In WholeSaleJanRange
    r.Offset(0, 2).Value = SumRequestQuantity(r.Value, False)
Next

End Sub

Private Function CountMallOrder(ByVal Code As String) As String
'モール別の件数略号を作る。A2Rはアマゾン2件、楽天1件の手配依頼あり。

Dim Counter(3) As Long, Mall As String, FoundCode As String, Endrow As Long, k As Long, Mark As String

Endrow = Worksheets("セラー分").UsedRange.Rows.Count

For k = 2 To Endrow
    
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

CountMallOrder = Mark

End Function

Private Function SumRequestQuantity(ByVal Code As String, ByVal IsSeller As Boolean) As Long
    '商品コード別の手配依頼数量を合計する。
    
    Dim TargetSheet As Worksheet, Endrow As Long, TmpSum As Long

    '数量列はどちらもE列なので、対象シートを切り換えるだけでよい。
    Set TargetSheet = Worksheets(IIf(IsSeller, "セラー分", "卸分"))
    
    Endrow = TargetSheet.UsedRange.Rows.Count
     
    Dim TargetRange As Range, QuantityRange As Range
    Set TargetRange = TargetSheet.Range("C2").Resize(Endrow - 1, 1)
    Set QuantityRange = TargetSheet.Range("E2").Resize(Endrow - 1, 1)
    
    TmpSum = WorksheetFunction.SumIf(TargetRange, Code, QuantityRange)
    
    SumRequestQuantity = TmpSum

End Function



