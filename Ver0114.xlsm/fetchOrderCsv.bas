Attribute VB_Name = "fetchOrderCsv"
Option Explicit

'明細と注文ヘッダーのあるフォルダを指定、最後必ず\マーク
Const CSV_PATH As String = "\\MOS10\Users\mos10\Desktop\ヤフー\"

Sub 梱包室受注ファイル読込()

Dim LineBuf As Variant
Dim OrderDetail As Variant

'ファイル操作オブジェクト生成
Dim FSO As New FileSystemObject

' Meisai.csvとtyumon_H.csvの存在チェック
Dim MeisaiPath As String
MeisaiPath = CSV_PATH & "Meisai.csv"

If FSO.FileExists(MeisaiPath) = False Then
    
    MsgBox "Meisai.csvが見つかりません" & vbLf & _
            CreateObject("WScript.Network").UserName & "ではMOS10のデータを参照できないので、別PCで実行してください。" & vbLf & _
            "もしくは、ヤフーの管理画面からダウンロードして、「個別読込」で指定してください。" & vbLf & _
            vbLf & "処理を終了します。"
    
    End

End If

Dim TyumonhPath As String
TyumonhPath = CSV_PATH & "tyumon_H.csv"

If FSO.FileExists(TyumonhPath) = False Then
    
    MsgBox "tyumon_H.csvが見つかりません" & vbLf & _
            CreateObject("WScript.Network").UserName & "ではMOS10のデータを参照できないので、別PCで実行してください。" & vbLf & _
            "もしくは、ヤフーの管理画面からダウンロードして、「個別読込」で指定してください。" & vbLf & _
            vbLf & "処理を終了します。"
    
    End

End If

' 本日分、読込済か確認
If LogSheet.Range("LastFetchNewOrder").Value = Date Then
    
    Dim mb As Variant
    mb = MsgBox("本日分は読込済です。" & vbLf & "処理を続けますか？", vbYesNo + vbExclamation)
        
    If mb = vbNo Then
        MsgBox "処理をキャンセルしました。"
        Exit Sub
    
    End If
End If

Call readMeisai(MeisaiPath)

Call sortOrderId

Call readTyumonH(TyumonhPath)

LogSheet.Range("LastFetchNewOrder").Value = Date

ThisWorkbook.Save

'要望列を展開します。
OrderSheet.Outline.ShowLevels ColumnLevels:=2

MsgBox prompt:=Format(Date, "m月d日") & "受注分  読込完了しました", Buttons:=vbInformation

End Sub

Function Meisai個別読込(Optional str As String = "") As Variant

Dim FilePath As String

'meisai.csvをファイルダイアログで指定"
FilePath = setCsvPath("meisai.csv")

If FilePath = "" Then
    MsgBox "ファイル指定がキャンセルされました。"
    Exit Function
End If

Call readMeisai(FilePath)

Call sortOrderId

MsgBox "読込完了"

End Function

Function tyumon_H個別読込(Optional str As String = "") As Variant

Dim FilePath As String

'tyumon_H.csvをファイルダイアログから指定する
FilePath = setCsvPath("tyumon_H.csv")

If FilePath = "" Then
    MsgBox "ファイル指定がキャンセルされました。"
    Exit Function
End If

Call readTyumonH(FilePath)

MsgBox "読込完了"

End Function

Function setCsvPath(CsvName As String)

'Applicationオブジェクト取得
Dim xlApp As Application
Set xlApp = Application

'｢ファイルを開く｣のフォームでファイル名の指定を受ける
Dim FileName As Variant
FileName = xlApp.GetOpenFilename(filefilter:="CSVファイル,*.csv" _
                                    , Title:=CsvName & "を指定してください")

'キャンセルされた場合はFalseが返るので以降の処理は行なわない
If VarType(FileName) = vbBoolean Then End

setCsvPath = FileName
    
End Function

Private Sub readMeisai(Path As String)

'Meisai.CSVをOrderSheet=注文一覧に追記する

'ダブりチェックのために読込前の注残シートのレンジを指定
Dim LastRow As Integer
LastRow = OrderSheet.Cells.SpecialCells(xlCellTypeLastCell).Row

Dim RngLoadedOrders As Range
Set RngLoadedOrders = OrderSheet.Range(Cells(2, 2), Cells(LastRow, 2))

Dim FSO As Object
Set FSO = New FileSystemObject

Dim TS As Textstream
Set TS = FSO.OpenTextFile(Path, ForReading)


Dim i As Long
i = LastRow '項目行は出力しないので、iは終端行から開始
    
Do Until TS.AtEndOfStream
    
' 行を取り出して必要な項目のみを配列に入れ直す
    Dim LineBuf As Variant
    LineBuf = Split(TS.ReadLine, """,""")
    
    Dim OrderDetail As Variant
    
    '0=1列目=注文番号、1=2列目=1注文内で何アイテム目か、2=3列目=個数、4=5列目=コード 5=6列目=商品名
    'ハードコーディングされているので、注文管理画面から出力項目を変更したら、読み取れなくなります。

    OrderDetail = Array(LineBuf(0), LineBuf(1), LineBuf(2), LineBuf(3), LineBuf(4))
    
    Dim j As Integer
    For j = 0 To UBound(OrderDetail)
        OrderDetail(j) = Trim(Replace(OrderDetail(j), Chr(34), "")) 'chr(34)で " [半角二重引用符]らしい
    
    Next
    
    'ループ一回目ではヘッダーなので、インクリメントへ飛ぶ
    If OrderDetail(0) = "Order ID" Then GoTo increment
    
    '注文番号が既に読込済のセル範囲にある場合もインクリメントへ
    If WorksheetFunction.CountIf(RngLoadedOrders, OrderDetail(0)) <= 0 Then
    
        Cells(i, 1).Value = Date
        Cells(i, 2).Value = OrderDetail(0)
        Cells(i, 4).Value = OrderDetail(1)
        Cells(i, 5).Value = OrderDetail(3)
        Cells(i, 6).Value = OrderDetail(4)
        Cells(i, 7).Value = OrderDetail(2)
    
    Else
        GoTo increment
    
    End If
    
increment:
    i = i + 1

Loop

Call sortOrderId

'ユーザーフォーム呼び出しボタンの位置調整
OrderSheet.Shapes("ShowFormButton").Top = OrderSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Offset(2, 1).Top
OrderSheet.Shapes("hideWishCol").Top = OrderSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Offset(2, 1).Top

' TextStreamを切断
TS.Close

End Sub

Private Sub readTyumonH(Path As String)

Dim FSO As Object
Set FSO = New FileSystemObject

Dim TS As Textstream
'Set TS = FSO.OpenTextFile("C:\Users\mos9\Downloads\tyumon_H_3.csv", ForReading)
Set TS = FSO.OpenTextFile(Path, ForReading)

'読込済注文番号のレンジをセット
Dim LoadedOrderRange As Range
Set LoadedOrderRange = OrderSheet.Cells(2, 2).Resize(OrderSheet.Cells(2, 2).SpecialCells(xlCellTypeLastCell).Row, 1)

Do Until TS.AtEndOfStream
    
' 行を取り出して必要な項目のみを配列に入れ直す
    Dim LineBuf As Variant
    LineBuf = Split(TS.ReadLine, ",")
    
    '0=1列目=注文番号、注文者名、要望、決済方法、クーポン値引き
    Dim order As Variant
    order = Array(LineBuf(0), LineBuf(5), LineBuf(36), LineBuf(35), LineBuf(43))
        
    Dim j As Integer
    For j = 0 To UBound(order)
        order(j) = Trim(Replace(order(j), Chr(34), "")) 'chr(34)で " [半角二重引用符]
    
    Next

    '注文番号の行を調べる
    Dim FindRow As Double 'Match関数の返値はDouble型
    
    On Error Resume Next
        
        FindRow = WorksheetFunction.Match(CDbl(order(0)), LoadedOrderRange, 0) + 1  'コードレンジはB2から始まるので行数は1加えた数
        
        If Err Then GoTo continue
        
    On Error GoTo 0
    
    Range("C" & FindRow).Value = order(1) '注文者名を入れる
    
    '一旦、tmpに備考欄内容を保持
    Dim tmp As String
    tmp = Range("S" & FindRow).Value
    
    If order(4) < 0 Then tmp = tmp & "クーポン利用 " & order(3)
    If order(3) = "銀行振込（前払い）" Then tmp = tmp & "振込 口座案内 未"
    
    Range("S" & FindRow).Value = tmp 'tmpをセルに書き戻す
    
    If order(2) <> "" Then Range("Q" & FindRow).Value = order(2) '要望を転記
    
continue:
    
Loop

' オブジェクトを破棄
TS.Close
Set TS = Nothing
Set FSO = Nothing

End Sub

Sub sortOrderId()

'OrderIDの列を探す B2決め打ちでもよくないかな？
Dim col_orderID As Range
Set col_orderID = OrderSheet.Rows(1).Find("Order ID")

With OrderSheet.Sort

    .SortFields.Clear '一旦、条件をクリアーして
    .SortFields.Add key:=col_orderID, order:=xlAscending 'ソート条件をセット

    'ソートを実行する
    .SetRange Range("A1").CurrentRegion
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply 'ソート適用

End With

End Sub
