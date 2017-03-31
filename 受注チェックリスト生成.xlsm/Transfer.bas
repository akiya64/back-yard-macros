Attribute VB_Name = "Transfer"
Option Explicit

Sub 作業シートへデータ抽出()

'受注データシートをフィルターして必要項目の楽天レンジのみをコピーする

Sheet1.Range("A1").CurrentRegion.AutoFilter Field:=10, Criteria1:="楽天店"

Dim FilteredRange As Range
Set FilteredRange = Range("A1").CurrentRegion

'必要な列のレンジを指定
Dim RequireColumns As Range
Set RequireColumns = Columns("A:O")

Dim TargetRange As Range

'フィルター後の必要項目列のみをコピーする
Intersect(FilteredRange, RequireColumns).Copy

'作業シートへ貼り付け、セルの調整
'Pasteメソッドの失敗ダイアログが出る場合があるので、レジュームネクストとする。
'ダイアログの表示は再現性が特定できず、クリップボードの内容などが関係している模様。
On Error Resume Next
With Worksheets.Add
    .Paste
    .Name = "作業シート"
    
    '列幅調整 商品名、届け先、住所は固定幅 単位：ポイント
    .Columns("A:B").AutoFit
    .Columns("D:I").AutoFit
    .Columns("C").ColumnWidth = 40
    .Columns("K").ColumnWidth = 20
    .Columns("L").AutoFit
    .Columns("M:Q").ColumnWidth = 20
    
End With
On Error GoTo 0

'受注データシートのオートフィルター解除
Sheet1.Range("A1").CurrentRegion.AutoFilter

'後の処理のために、先に列を挿入
Columns("L").Insert
Range("L1").Value = "届け先住所"

Columns("C").Insert
Range("C1").Value = "JANコード"

Range("A1").Select

End Sub
Sub 店舗識別番号振替()

    
Worksheets("作業シート").Activate

Dim i As Long
i = 2

Do
    'モール名から、社内処理用のモール番号へ振り替えて、店舗コード列を上書き
    Dim Mall As String, MallId As Integer
    Mall = Cells(i, 11).Value
    
    Select Case Mall
        Case "Amazon店"
            MallId = 1
        Case "楽天店"
            MallId = 2
        Case "Yahoo店"
            MallId = 4
    End Select
    
    '納品書区分はDBでは数値型
    Cells(i, 10).NumberFormatLocal = "#"
    Cells(i, 10) = MallId
    
    i = i + 1

Loop Until IsEmpty(Cells(i, 1))

End Sub

Sub 住所結合()
'届け先都道府県、届け先市区町村、届け先住所1、届け先住所2、届け先住所3 列が分かれている。
'「届け先住所」列へ結合して格納。

Worksheets("作業シート").Activate

Dim i As Long
i = 2

Do
    'L列に住所を結合
    Cells(i, 13).Value = Cells(i, 14).Value & Cells(i, 15).Value & Cells(i, 16).Value
    
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

Sub 商品名修正()

'商品名から、楽天のキャンペーン情報を削除する
'≪≫か【】で先頭に記載されているので、正規表現で検出して括弧ごと削除、複数括弧対応
'また、DBのフィールドサイズが50文字なので、45文字でカットする。

Worksheets("作業シート").Activate

'ループ内で使う行カウンタ
Dim i As Long
i = 2

'正規表現オブジェクトと、パターンをセット
Dim Reg As New RegExp
Reg.Global = True
Reg.Pattern = "^((≪|【).*?(】|≫))*"

Do
    Dim ProductName As String
    
    ProductName = Cells(i, 4).Value
    ProductName = Reg.Replace(ProductName, "")
            
    Cells(i, 4) = Left(ProductName, 45)
        
    i = i + 1

Loop Until IsEmpty(Cells(i, 1))

End Sub

Sub 書式と型の変更()

Dim i As Long
i = 2

Do
    '受注番号の修正
    Cells(i, 1).NumberFormatLocal = "#"
    Cells(i, 1).Value = CDbl(Cells(i, 1).Value)
    
    '日付の表示を修正
    Cells(i, 7).NumberFormatLocal = "yyyy/M/dd"
    Cells(i, 7).Value = Format(Cells(i, 7).Value, "yyyy/M/dd")
    
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
    If Cells(i, 3).Value Like "77777*" Then GoTo Continue
    
    '1行分、商品コードと住所をコピー
    Dim Record As Range
    Set Record = Range(Cells(i, 1), Cells(i, 5))
    Set Record = Union(Record, Range(Cells(i, 7), Cells(i, 13)), Cells(i, 17))
    
    Record.Copy Worksheets("アップロードシート").Cells(k, 1)
    
    '受注明細枝番は全て1でよい
    Worksheets("アップロードシート").Cells(k, 14).Value = "1"

    'コピー先行カウンタをインクリメント
    k = k + 1

Continue:
    i = i + 1
    
Loop Until IsEmpty(Cells(i, 1))

Worksheets("アップロードシート").Activate

End Sub

Function ValidateName(Name As String) As String

Dim Reg As New RegExp

Reg.Global = True
Reg.Pattern = "^((≪|【).*?(】|≫))*"
Name = Reg.Replace(Name, "")

ValidateName = Name

End Function

