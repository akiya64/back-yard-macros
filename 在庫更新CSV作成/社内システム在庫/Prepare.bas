Attribute VB_Name = "Prepare"
Sub FetchYahooCSV()
'ヤフーのDataCSVをヤフーデータシートにコピーします。

'オートフィルターを解除

yahoo6digit.Activate

If Not yahoo6digit.AutoFilter Is Nothing Then yahoo6digit.Range("A1").AutoFilter

Dim DataCsvPath As String
' ｢ファイルを開く｣のフォームでファイル名の指定を受ける
DataCsvPath = Application.GetOpenFilename(Title:="ヤフーの商品情報CSVを指定")

' キャンセルされた場合はFalseが返るので以降の処理は行なわない
If VarType(DataCsvPath) = 8 Then Exit Sub

Workbooks.Open DataCsvPath

Dim CsvName As String
CsvName = Dir(DataCsvPath)

'「ヤフーデータ」をクリア
yahoo6digit.Cells.Clear

Dim header As Variant
header = Array("sub-code", "original-price", "options", "caption")  '"headline"

With Workbooks(CsvName).Sheets(1)

    'ヤフーCSVをXLSMへコピー
    'ヘッダーを調べてAbstractまでの間に、sub-code/original-price/options/headline/captionがあれば列を削除
    i = 1
    Do Until IsEmpty(.Cells(1, i))
        For Each v In header
            If Cells(1, i) = v Then
                .Columns(i).Delete
            End If
        Next
            
        i = i + 1
    
    Loop
    
    .Range("A1").CurrentRegion.WrapText = False
    .Range("A1").CurrentRegion.Copy Destination:=yahoo6digit.Range("A1")

    ActiveWindow.Close saveChanges:=False

End With

End Sub

Sub SetRangeName()
'各シートのコードレンジを「名前」で呼べるよう、定義し直す
'連想配列とかつかってイテレート回す様にすべきだが
'代わる部分は各々…シート名、最初のレンジ、レンジ名 三つだとコピペ書き換えの方が楽か

'ヤフーシート「YahooCodeRange」の範囲を再定義
With yahoo6digit
    Set rng = .Range("C1").Resize(.Range("C1").SpecialCells(xlCellTypeLastCell).Row, 1)
    Names.Add Name:="YahooCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'処分・在廃の「StockOnlyCodeRange」の範囲を再定義
With StockOnly
    Set rng = .Range("C1").Resize(.Range("C1").SpecialCells(xlCellTypeLastCell).Row, 1)
    Names.Add Name:="StockOnlyCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'商魂マスターシート「SyokonCodeRange」の範囲を再定義
With SyokonMaster
    Set rng = .Range("A1").Resize(.Range("A1").SpecialCells(xlCellTypeLastCell).Row, 1)
    Names.Add Name:="SyokonCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'在庫セット除外シート
With ExceptQty
    Set rng = .Range("C1").Resize(.Range("C1").SpecialCells(xlCellTypeLastCell).Row, 1)
    Names.Add Name:="ExceptCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'廃番シート「EolCodeRange」
With Eol
    Set rng = .Range("C1").Resize(.Range("C1").SpecialCells(xlCellTypeLastCell).Row, 1)
    Names.Add Name:="EolCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'Slims在庫CSV「SlimsCodeRange」
With Slims
    Set rng = .Range("B1").Resize(.Range("B1").SpecialCells(xlCellTypeLastCell).Row, 1)
    Names.Add Name:="SlimsCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

End Sub

