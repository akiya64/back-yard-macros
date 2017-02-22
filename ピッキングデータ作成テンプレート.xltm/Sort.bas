Attribute VB_Name = "Sort"
Option Explicit

Sub 振分用シート_ソート(Sheet As Worksheet)
Attribute 振分用シート_ソート.VB_ProcData.VB_Invoke_Func = " \n14"

Dim SortRange As Range
Set SortRange = Sheet.Range("A1").CurrentRegion

Dim CodeRange As Range
Set CodeRange = Sheet.Range("A2:A" & SortRange.Rows.Count)

'ソート条件をセット
With Sheet.Sort
    
    '一旦ソートをクリア
    .SortFields.Clear
    
    'ソートキーをセット 第一キー 商品コード：色、第二キー 商品コード：昇順
    .SortFields.Add Key:=CodeRange, SortOn:=xlSortOnCellColor, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add Key:=CodeRange, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    'ソート対象のデータが入ってる範囲を指定して
    .SetRange SortRange
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    
    'セットした条件を適用
    .Apply

End With

'カレントリージョンがセレクトされているので、選択セルをセルA1にセットし直す
Sheet.Activate
Range("A1").Select

End Sub
