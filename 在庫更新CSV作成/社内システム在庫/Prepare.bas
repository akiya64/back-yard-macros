Attribute VB_Name = "Prepare"
Sub FetchYahooCSV()
'ヤフーのDataCSVをヤフーデータシートにコピーします。

'オートフィルターを解除

yahoo6digit.Activate

If Not yahoo6digit.AutoFilter Is Nothing Then yahoo6digit.Range("A1").AutoFilter

'「ヤフーデータ」をクリア
yahoo6digit.Cells.Clear

Dim DataCsvPath As String
' ｢ファイルを開く｣のフォームでファイル名の指定を受ける
DataCsvPath = Application.GetOpenFilename(Title:="ヤフーの商品情報CSVを指定")

' キャンセルされた場合はFalseが返るので以降の処理は行なわない
If VarType(DataCsvPath) = vbBoolean Then Exit Sub

Workbooks.Open DataCsvPath

Dim CsvName As String
CsvName = Dir(DataCsvPath)

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


Sub FetchSecondInventry()

'棚無在庫確認表を開いて棚無データ=SecondInventryにコピー

Application.ScreenUpdating = False

Const SECOND_INVENTRY_FILE As String = "\\server02\商品部\ネット販売関連\棚無在庫確認表.xlsm"
Const SECOND_INVENTRY_SHEET_NAME As String = "棚無データ"

SecondInventry.Cells.Clear

'在庫表を開いてシートをコピー
'1.在庫表の存在チェック

Dim WbName As String
WbName = Dir(SECOND_INVENTRY_FILE)

If WbName = "" Then
    MsgBox "棚無しの在庫表が存在しません", vbExclamation
    Exit Sub
End If

'2.同名ブックを開いていないかチェック
Dim wb As Workbook

For Each wb In Workbooks
    If wb.Name = WbName Then
        MsgBox WbName & vbCrLf & "はすでに開いています", vbExclamation
        Exit Sub
    End If
Next wb

'ここでブックを開く
Workbooks.Open SECOND_INVENTRY_FILE

'在庫表よりシートをコピー
For i = 1 To Workbooks.Count
    
    If Workbooks(i).Name = WbName Then
        
        With Workbooks(i).Sheets(SECOND_INVENTRY_SHEET_NAME)
            
            Dim LastRow As Long
            LastRow = .Range("A1").SpecialCells(xlCellTypeLastCell).row
            
            .Range("A1").Resize(LastRow - 1, 4).Copy
        
        End With
        
        SecondInventry.Range("A1").PasteSpecial (xlPasteValues)
        
        Application.DisplayAlerts = False
        
        'コピーが終われば速やかに在庫表を閉じる
        Workbooks(WbName).Close saveChanges:=False
        
        Application.DisplayAlerts = True
        
        Exit For
        
    End If

Next

Worksheets(SECOND_INVENTRY_SHEET_NAME).Range("A1").AutoFilter

With Worksheets(SECOND_INVENTRY_SHEET_NAME).AutoFilter.Sort

    .SortFields.Clear 'ソートフィールドを一旦クリアー
    
    'ソートフィールドを指定
    .SortFields.Add Key:=Range("C1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    
    'ソート順序指定
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    
    'ソート適用
    .Apply

End With

End Sub

Sub SetRangeName()
'各シートのコードレンジを「名前」で呼べるよう、定義し直す
'連想配列とかつかってイテレート回す様にすべきだが
'代わる部分は各々…シート名、最初のレンジ、レンジ名 三つだとコピペ書き換えの方が楽か

'ヤフーシート「YahooCodeRange」の範囲を再定義
With yahoo6digit
    Set rng = .Range("C1").Resize(.Range("C1").SpecialCells(xlCellTypeLastCell).row, 1)
    Names.Add Name:="YahooCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'処分・在廃の「StockOnlyCodeRange」の範囲を再定義
With StockOnly
    Set rng = .Range("C1").Resize(.Range("C1").SpecialCells(xlCellTypeLastCell).row, 1)
    Names.Add Name:="StockOnlyCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'商魂マスターシート「SyokonCodeRange」の範囲を再定義
With SyokonMaster
    Set rng = .Range("A1").Resize(.Range("A1").SpecialCells(xlCellTypeLastCell).row, 1)
    Names.Add Name:="SyokonCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'在庫セット除外シート
With ExceptQty
    Set rng = .Range("C1").Resize(.Range("C1").SpecialCells(xlCellTypeLastCell).row, 1)
    Names.Add Name:="ExceptCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'廃番シート「EolCodeRange」
With Eol
    Set rng = .Range("C1").Resize(.Range("C1").SpecialCells(xlCellTypeLastCell).row, 1)
    Names.Add Name:="EolCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'棚無在庫表シート「SecondInventryCodeRange」
With SecondInventry
    Set rng = .Range("B1").Resize(.Range("B1").SpecialCells(xlCellTypeLastCell).row, 1)
    Names.Add Name:="SecondInventryCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

End Sub

