Attribute VB_Name = "FetchOrder"
Option Explicit

'明細と注文ヘッダーのあるフォルダを指定、最後必ず\マーク
Const CSV_PATH As String = "C:\Users\yahoo_operator\Desktop\ヤフー\"
Const ALTER_CSV_PATH As String = "\\shiperPC\Users\shipper\Desktop\ヤフー\"

Sub 受注ファイル読込()

OrderSheet.Activate

Dim LineBuf As Variant

'ファイル操作オブジェクト生成
Dim FSO As New FileSystemObject

' Meisai.csvとtyumon_H.csvのCSVファイルのパスをセット
Dim MeisaiPath As String, TyumonhPath As String

If FSO.FileExists(CSV_PATH & "Meisai.csv") Then

    MeisaiPath = CSV_PATH & "Meisai.csv"
    TyumonhPath = CSV_PATH & "tyumon_H.csv"

ElseIf FSO.FileExists(ALTER_CSV_PATH & "Meisai.csv") Then
   
    MeisaiPath = ALTER_CSV_PATH & "Meisai.csv"
    TyumonhPath = ALTER_CSV_PATH & "tyumon_H.csv"

Else
    
    MsgBox "meisai.csv ファイルなし"
    End

End If

Call ReadMeisai(MeisaiPath)

Call ReadTyumonH(TyumonhPath)

Dim LastRow As Long
LastRow = Range("D1").SpecialCells(xlCellTypeLastCell).Row

'アドイン用の行・列 指定
Range("L1").Value = "アドイン指定 台帳：9998"
Range("L2:O2") = Array(2, 4, LastRow, 12)

MsgBox "アドインを実行して下さい。"

End Sub

Private Sub ReadMeisai(Path As String)

'Meisai.CSVをOrderSheet=注文一覧に追記する

Dim FSO As Object
Set FSO = New FileSystemObject

Dim TS As Textstream
Set TS = FSO.OpenTextFile(Path, ForReading)

Dim i As Long
i = 1 '項目行は出力しないので、iは1行目から開始
    
Do Until TS.AtEndOfStream
    
' 行を取り出して必要な項目のみを配列に入れ直す
    Dim LineBuf As Variant
    LineBuf = Split(TS.ReadLine, ",")
       
    Dim j As Integer
    For j = 0 To UBound(LineBuf)
        LineBuf(j) = Trim(Replace(LineBuf(j), Chr(34), "")) 'chr(34)で " [半角二重引用符]
    
    Next
    
    'ループ一回目ではヘッダーなので、インクリメントへ飛ぶ
    If LineBuf(0) = "Order ID" Then GoTo increment
    
    'CSV側ヘッダー 0:Order ID/1:Line ID/2:Quantity/3:Product Code/4:Description/5:Option Name/6:Option Value/7:Unit Price/
    
    With Worksheets("受注データシート")
        .Range("A" & i).Value = LineBuf(0)
        .Range("C" & i).Value = LineBuf(1)
        
        .Range("C" & i).NumberFormatLocal = "@"
        .Range("C" & i).Value = LineBuf(3)
        
        .Range("D" & i).NumberFormatLocal = "@"
        .Range("D" & i).Value = LineBuf(3)
        
        .Range("E" & i).Value = LineBuf(4)
        .Range("F" & i).Value = LineBuf(2)
        .Range("G" & i).Value = LineBuf(7)
        
        'E列をアドイン用に社内コードへ変換
        Dim ycode As String
        ycode = .Range("D" & i).Value
        
        If ycode Like "#####" Then
        
            .Range("D" & i).Value = "0" & ycode
        
        End If
        
        'セット分解
        If ycode Like "7777*" Then
            
            Call SetParser.ParseItems(.Range("D" & i))
        
        End If
    
        '単体セット分解
    
    End With
    
increment:
    i = i + 1

Loop

TS.Close

End Sub

Private Sub ReadTyumonH(Path As String)

Dim FSO As Object
Set FSO = New FileSystemObject

Dim TS As Textstream
Set TS = FSO.OpenTextFile(Path, ForReading)

'読込済注文番号のレンジをセット、A1からA列の番号入り最終セルまで
Dim LoadedOrderRange As Range
Set LoadedOrderRange = OrderSheet.Cells(1, 1).Resize(OrderSheet.Cells(1, 1).SpecialCells(xlCellTypeLastCell).Row, 1)

Do Until TS.AtEndOfStream
    
' 行を取り出して必要な項目のみを配列に入れ直す
    Dim LineBuf As Variant
    LineBuf = Split(TS.ReadLine, """,""")
    
    '0=1列目=注文番号、注文者名、要望、決済方法、クーポン値引き
    Dim Order As Variant
    Order = Array(LineBuf(0), LineBuf(5), LineBuf(36), LineBuf(34), LineBuf(43))
        
    Dim j As Integer
    For j = 0 To UBound(Order)
        Order(j) = Trim(Replace(Order(j), Chr(34), "")) 'chr(34)で " [半角二重引用符]
    
    Next

    '注文番号の行を調べる、注文番号はDobule型で入っている。CSVはString型。
    'Match関数の返値はDouble型
    
    Dim FindRow As Double
    
    On Error Resume Next
        
        FindRow = WorksheetFunction.Match(CDbl(Order(0)), LoadedOrderRange, 0)
        
        If Err Then
            Err.Clear
            GoTo Continue
        End If
    
    On Error GoTo 0
        
    Dim i As Long
    i = 0
    
    '注文者名を記入 オフセットしつつ、全ての行へ記入
    Do While Range("A" & FindRow).Offset(i, 0).Value = CDbl(Order(0))
        
        Range("A" & FindRow).Offset(i, 1).Value = LineBuf(5)
        i = i + 1
    
    Loop
    
    '備考欄へ追記 クーポン利用かつ代引き・銀行振込・ヤフーマネー決済 確認
    Dim tmp As String
    tmp = ""
    
    If Order(3) = "payment_d1" And Order(4) < 0 Then tmp = "代引き クーポン利用 "
    If Order(3) = "payment_b1" Then tmp = tmp & "振込 口座案内 未"
    If Order(3) = "payment_a16" Then tmp = tmp & "Yahoo!マネー払い"
    
    Range("K" & FindRow).Value = tmp 'tmpをセルに書き戻す
        
Continue:
    
Loop

' オブジェクトを破棄
TS.Close
Set TS = Nothing
Set FSO = Nothing

End Sub
