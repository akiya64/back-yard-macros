Attribute VB_Name = "SumPurcheseReq"
Option Explicit

Sub SumPuchaseRequest()
'商品別に手配依頼数を集計

Worksheets("手配数量入力シート").Activate

'セラー分の商品コードをコピーして、重複削除
If Worksheets("セラー分").Range("A2").Value = "" Then GoTo Vendor
Worksheets("セラー分").Range("C2:D" & Worksheets("セラー分").UsedRange.Rows.Count).Copy Destination:=Range("G2")
Range("A1").CurrentRegion.RemoveDuplicates Columns:=7, Header:=xlYes


'手配件数の合計略号と、手配依頼数の合計を入れる
Dim r As Range, CodeRange As Range

Set CodeRange = Range(Cells(2, 7), Cells(2, 7).End(xlDown))

For Each r In CodeRange
    r.Offset(0, -1).Value = CountMallOrder(r.Value)
    r.Offset(0, 2).Value = SumRequestQuantity(r.Value, True)
Next

Vendor:

'卸分の手配依頼数量の合計を算出
'卸分の商品コードを入れる
If Worksheets("卸分").Range("A2").Value = "" Then GoTo Quit
Worksheets("卸分").Range("C2:D" & Worksheets("卸分").UsedRange.Rows.Count).Copy Destination:=Range("G2").End(xlDown).Offset(1, 0)

'コピーして重複削除
Range(Cells(2, 6).End(xlDown).Offset(1, 1), Cells(2, 7).End(xlDown).Offset(0, 1)).RemoveDuplicates 1, xlNo

Dim WholeSaleJanRange As Range
Set WholeSaleJanRange = Range(Cells(2, 6).End(xlDown).Offset(1, 1), Cells(2, 7).End(xlDown))
WholeSaleJanRange.Offset(0, -1).Value = "V"

For Each r In WholeSaleJanRange
    r.Offset(0, 2).Value = SumRequestQuantity(r.Value, False)
Next

Quit:

End Sub

Private Function CountMallOrder(ByVal Code As String) As String
'モール別の件数略号を作る。A2Rはアマゾン2件、楽天1件の手配依頼あり。

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

Mark = IIf(Counter(0) > 0, "A" & IIf(Counter(0) > 1, Counter(0), ""), "")
Mark = Mark & IIf(Counter(1) > 0, "R" & IIf(Counter(1) > 1, Counter(1), ""), "")
Mark = Mark & IIf(Counter(2) > 0, "Y" & IIf(Counter(2) > 1, Counter(2), ""), "")
Mark = Mark & IIf(Counter(3) > 0, "SP" & IIf(Counter(3) > 1, Counter(3), ""), "")

CountMallOrder = Mark

End Function

Private Function SumRequestQuantity(ByVal Code As String, ByVal IsSeller As Boolean) As Long
    '商品コード別の手配依頼数量を合計する。
    
    Dim TargetSheet As Worksheet, EndRow As Long, TmpSum As Long

    '数量列はどちらもE列なので、対象シートを切り換えるだけでよい。
    Set TargetSheet = Worksheets(IIf(IsSeller, "セラー分", "卸分"))
    
    EndRow = TargetSheet.UsedRange.Rows.Count
     
    Dim TargetRange As Range, QuantityRange As Range
    Set TargetRange = TargetSheet.Range("C2").Resize(EndRow - 1, 1)
    Set QuantityRange = TargetSheet.Range("E2").Resize(EndRow - 1, 1)
    
    TmpSum = WorksheetFunction.SumIf(TargetRange, Code, QuantityRange)
    
    SumRequestQuantity = TmpSum

End Function
