Attribute VB_Name = "Transfer"
Option Explicit

Sub 作業シートへデータ抽出()

'産直シートをフィルターして必要列のレンジのみをコピーする

Sheet1.Range("A1").CurrentRegion.AutoFilter Field:=2, Criteria1:="<>"

Dim FilteredRange As Range
Set FilteredRange = Range("A1").CurrentRegion

'必要な列のレンジを指定
Dim RequireColumns As Range
Set RequireColumns = Columns("A:P")

Dim TargetRange As Range

'フィルター後の必要項目列のみをコピーする
Intersect(FilteredRange, RequireColumns).Copy

'作業シートへ貼り付け、セルの調整
With Worksheets.Add
    .Paste
    .Name = "作業シート"
    
    '列幅調整 商品名、届け先、住所は固定幅 単位：ポイント
    .Columns("A:B").AutoFit
    .Columns("D:I").AutoFit
    .Columns("C").ColumnWidth = 40
    .Columns("K").ColumnWidth = 20
    .Columns("L").AutoFit
    .Columns("M:P").ColumnWidth = 20
    
End With

'オートフィルター解除
Sheet1.Range("A1").CurrentRegion.AutoFilter

'後の処理のために、先に列を挿入
Columns("L").Insert
Range("L1").Value = "届け先住所"

Columns("C").Insert
Range("C1").Value = "JANコード"

Range("A1").Select

End Sub

Sub 住所結合()
'届け先都道府県、届け先市区町村、届け先住所1、届け先住所2、届け先住所3 列が分かれている。
'「届け先住所」列へ結合して格納。

Worksheets("作業シート").Activate

Dim i As Long
i = 2

Do
    'L列に住所を結合
    Cells(i, 13).Value = Cells(i, 14).Value & Cells(i, 15).Value & Cells(i, 16).Value & Cells(i, 17).Value & Cells(i, 18).Value
    
    i = i + 1

Loop Until IsEmpty(Cells(i, 1))


End Sub

Sub JAN転記()
'商品コード列は、商魂コード か 空白として、6ケタ以外はJAN列へ転記する

Worksheets("作業シート").Activate

Dim i As Long
i = 2

Do
    Dim Code As String, Jan As String
    Code = Cells(i, 2).Value
    
    '数字5ケタ化
    If Code Like String(6, "#") And InStr(1, Code, "0") = 1 Then
        
        Code = Right(Code, 5)
        Jan = ""
        
        Cells(i, 2).Resize(1, 2).Value = Array(Code, Jan)
    
    '5ケタでも6ケタでもない場合、JAN列へ入れる
    ElseIf Not Code Like String(5, "#") And Not Code Like "5" & String(5, "#") Then
        
        Jan = Code
        Code = ""
    
        Cells(i, 2).Resize(1, 2).Value = Array(Code, Jan)
    
    End If

    i = i + 1

Loop Until IsEmpty(Cells(i, 1))

End Sub

Sub 楽天商品名修正()

'商品名から、楽天のキャンペーン情報を削除する
'≪≫か【】で先頭に記載されているので、今回はInstrで ≫ 】の位置を検出して前を削除
'正規表現での実装は、whatnot受注取込マクロにある。

Worksheets("作業シート").Activate

'キャンペーン情報を囲っている閉じカッコ配列を定義
Dim CloseBrackets() As Variant
CloseBrackets = Array(Array("【", "】"), Array("≪", "≫"))

'行カウンタ
Dim i As Long
i = 2

Do
    Dim ProductName As String
    ProductName = Cells(i, 3).Value
    
    '閉じ括弧が何文字目に出てくるか調べる
    Dim k As Integer
    For k = 0 To UBound(CloseBrackets)
    
        'キャンペーン情報の括弧が冒頭にあるかチェック
        If InStr(1, ProductName, CloseBrackets(k)(0)) = 1 Then
            
            Dim CampaignInfoCharEnd As Integer
            CampaignInfoCharEnd = InStr(1, ProductName, CloseBrackets(k)(1)) + 1
            
            Debug.Assert CampaignInfoCharEnd = 0
            
            Cells(i, 3) = Mid(ProductName, CampaignInfoCharEnd)
            
        End If
        
    Next
        
    i = i + 1

Loop Until IsEmpty(Cells(i, 1))

End Sub

Sub 提出用シートへ転記()

Worksheets("作業シート").Activate

'A2～最終行まで、セット商品以外を転記
Dim i As Long, k As Long
i = 2
k = 2

Do

    '7777始まりは転記しない
    If Cells(i, 3).Value Like "77777" Then GoTo Continue
    
    '1行分のレンジを定義
    Dim Record As Range
    Set Record = Range(Cells(i, 1), Cells(i, 5))
    
    Set Record = Union(Record, Range(Cells(i, 7), Cells(i, 13)))
    
    '行をコピー、コピー先行カウンタをインクリメント
    Record.Copy Worksheets("提出シート").Cells(k, 1)
    k = k + 1

Continue:
    i = i + 1
    
Loop Until IsEmpty(Cells(i, 1))

Worksheets("提出シート").Activate

End Sub
