Attribute VB_Name = "Sort"
Option Explicit

Sub 振分用シート_ソート()
Attribute 振分用シート_ソート.VB_ProcData.VB_Invoke_Func = " \n14"
    
Dim SortRange As Range
Set SortRange = Worksheets("振分け用一覧シート").Range("A1").CurrentRegion

Dim CodeRange As Range
Set CodeRange = Worksheets("振分け用一覧シート").Range("C2:C" & SortRange.Rows.Count)

'ソート条件をセット
With Worksheets("振分け用一覧シート").Sort
    
    '一旦ソートをクリア
    .SortFields.Clear
    
    'ソートキーをセット 第一キー 商品コード：色、第二キー 商品コード：昇順
    'セル背景に色がついている棚無しを下に固める。背景なし＝棚あり、背景色つき＝棚無し それぞれの中で6ケタ昇順
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

End Sub
