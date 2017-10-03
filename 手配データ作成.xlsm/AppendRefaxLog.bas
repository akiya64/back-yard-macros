Attribute VB_Name = "AppendRefaxLog"
Option Explicit

Sub AppendRefaxList()
'FAX納期回答リストに本日の発注商品・発注保留を追記する

'FAX納期回答リストを開く
Dim RefaxBook As Workbook, RefaxSheet As Worksheet, WriteCell As Range
Set RefaxBook = FetchWorkBook("\\Server02\商品部\ネット販売関連\発注関連\半自動発注バックアップ\FAX納期回答リスト.xlsm")
Set RefaxSheet = RefaxBook.Worksheets("納期リスト")

RefaxSheet.Activate

If DateDiff("d", Date, Cells(Range("A1").CurrentRegion.Rows.Count, 5).Value) >= 0 Then
    RefaxBook.Close SaveChanges:=False
    Exit Sub
End If
'返信FAXの最終の空白行へ書き込む
Set WriteCell = RefaxSheet.Cells(Range("A1").CurrentRegion.Rows.Count, 1).Offset(1, 0)

'発注商品リストからデータをコピー
Dim DataCol As Range, PurchaseSheet As Worksheet
ThisWorkbook.Worksheets("発注商品リスト").Activate
ThisWorkbook.Worksheets("発注商品リスト").Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous

If Range("A2").Value = "" Then
    RefaxBook.Close SaveChanges:=True
    Exit Sub
Else
    Set DataCol = ThisWorkbook.Worksheets("発注商品リスト").Range(Cells(2, 1), Cells(2, 1).End(xlDown))
End If

'DataColレンジとWriteCellを右へオフセットしながらデータをコピーしていく。

DataCol.Offset(0, 6).Copy Destination:=WriteCell '発注数量
DataCol.Offset(0, 0).Copy Destination:=WriteCell.Offset(0, 1) '注番
DataCol.Offset(0, 1).Copy Destination:=WriteCell.Offset(0, 2) '仕入先
DataCol.Offset(0, 2).Copy Destination:=WriteCell.Offset(0, 3) 'モール識別記号
DataCol.Offset(0, 3).Copy Destination:=WriteCell.Offset(0, 4) '日付
DataCol.Offset(0, 4).Copy Destination:=WriteCell.Offset(0, 5) '商品コード
DataCol.Offset(0, 5).Copy Destination:=WriteCell.Offset(0, 6) '商品名


'同様に保留シートからデータをコピー

Dim HoldSheet As Worksheet
ThisWorkbook.Worksheets("保留").Activate

Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous

If Range("A2").Value <> "" Then
    Set DataCol = Worksheets("保留").Range(Cells(2, 1), Cells(1, 1).End(xlDown))
    
    RefaxSheet.Activate
    Set WriteCell = RefaxSheet.Cells(Range("A1").CurrentRegion.Rows.Count, 1).Offset(1, 0)
    
    '数量の頭に「保留」文言を入れて貼り付け
    Dim HoldQty As Variant, i As Long
    'HoldQtyは二次元配列で格納される
    HoldQty = DataCol.Offset(0, 6).Value
    
    '保留が1行の時は、配列にならない
    If IsArray(HoldQty) Then
        
        For i = 1 To UBound(HoldQty)
            HoldQty(i, 1) = "保留：" & HoldQty(i, 1)
        Next
        WriteCell.Resize(UBound(HoldQty), 1).Value = HoldQty
    
    Else
    
        WriteCell.Value = "保留：" & HoldQty
    
    End If
    
    DataCol.Offset(0, 1).Copy Destination:=WriteCell.Offset(0, 1) '注番
    DataCol.Offset(0, 2).Copy Destination:=WriteCell.Offset(0, 2) '仕入先
    DataCol.Offset(0, 3).Copy Destination:=WriteCell.Offset(0, 3) 'モール識別記号
    DataCol.Offset(0, 4).Copy Destination:=WriteCell.Offset(0, 4) '日付
    DataCol.Offset(0, 5).Copy Destination:=WriteCell.Offset(0, 5) '手配時商品コード
    DataCol.Offset(0, 7).Copy Destination:=WriteCell.Offset(0, 6) '商品名
    
    DataCol.Offset(0, 0).Copy
    WriteCell.Offset(0, 9).PasteSpecial Paste:=xlPasteValues   '保留理由

End If

On Error Resume Next
    Application.Run "FAX納期回答リスト.xlsm!一ヶ月以前転記"
    Application.Run "FAX納期回答リスト.xlsm!入荷日の算出式を入力"
    Application.Run "FAX納期回答リスト.xlsm!条件付き書式範囲修正"
On Error GoTo 0

RefaxBook.Save

End Sub


