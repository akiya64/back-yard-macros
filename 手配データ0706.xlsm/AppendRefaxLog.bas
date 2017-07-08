Attribute VB_Name = "AppendRefaxLog"
Option Explicit

Sub AppendRefaxList()

'FAX納期回答リストを開く
Dim RefaxBook As Workbook, RefaxSheet As Worksheet, WriteCell As Range
Set RefaxBook = FetchWorkBook("\\Server02\商品部\ネット販売関連\発注関連\半自動発注バックアップ\FAX納期回答リスト.xlsm")
Set RefaxSheet = RefaxBook.Worksheets("納期リスト")

RefaxSheet.Activate

If Cells(Range("A1").CurrentRegion.Rows.Count, 6).Value = Format(Date, "Mdd") Then
    RefaxBook.Close
    Exit Sub
End If
'返信FAXの最終の空白行へ書き込む
Set WriteCell = Cells(Range("A1").CurrentRegion.Rows.Count, 1).Offset(1, 0)

'発注商品リストからデータをコピー
Dim DataCol As Range
ThisWorkbook.Worksheets("発注商品リスト").Activate
Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
Set DataCol = Range(Cells(2, 1), Cells(2, 1).End(xlDown))

'DataColレンジとWriteCellを右へオフセットしながらデータをコピーしていく。

DataCol.Offset(0, 8).Copy Destination:=WriteCell '発注数量
DataCol.Offset(0, 0).Copy Destination:=WriteCell.Offset(0, 1) '注番
DataCol.Offset(0, 1).Copy Destination:=WriteCell.Offset(0, 3) '仕入先
DataCol.Offset(0, 2).Copy Destination:=WriteCell.Offset(0, 4) 'モール識別記号
DataCol.Offset(0, 3).Copy Destination:=WriteCell.Offset(0, 5) '日付
DataCol.Offset(0, 7).Copy Destination:=WriteCell.Offset(0, 8) '手配時商品コード
DataCol.Offset(0, 6).Copy Destination:=WriteCell.Offset(0, 10) '商品名


'同様に保留シートからデータをコピー

ThisWorkbook.Worksheets("保留").Activate
Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
Set DataCol = Range(Cells(2, 1), Cells(2, 1).End(xlDown))

RefaxSheet.Activate
Set WriteCell = Cells(Range("A1").CurrentRegion.Rows.Count, 1).Offset(1, 0)

'数量の頭に「保留」文言を入れて貼り付け
Dim HoldQty As Variant, i As Long
HoldQty = DataCol.Offset(0, 6).Value

For i = 1 To UBound(HoldQty)
    HoldQty(i, 1) = "保留：" & HoldQty(i, 1)
Next

WriteCell.Resize(UBound(HoldQty), 1).Value = HoldQty

DataCol.Offset(0, 1).Copy Destination:=WriteCell.Offset(0, 1) '注番
DataCol.Offset(0, 2).Copy Destination:=WriteCell.Offset(0, 3) '仕入先
DataCol.Offset(0, 3).Copy Destination:=WriteCell.Offset(0, 4) 'モール識別記号
DataCol.Offset(0, 4).Copy Destination:=WriteCell.Offset(0, 5) '日付
DataCol.Offset(0, 5).Copy Destination:=WriteCell.Offset(0, 8) '手配時商品コード
DataCol.Offset(0, 7).Copy Destination:=WriteCell.Offset(0, 10) '商品名

DataCol.Offset(0, 0).Copy
WriteCell.Offset(0, 22).PasteSpecial Paste:=xlPasteValues   '保留理由

RefaxBook.Close SaveChanges:=True

End Sub


