Attribute VB_Name = "FetchOrder"
Option Explicit

Sub ReadMeisai(Path As String)

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
    LineBuf = Split(TS.ReadLine, """,""")
       
    Dim j As Integer
    For j = 0 To UBound(LineBuf)
        LineBuf(j) = Trim(Replace(LineBuf(j), Chr(34), "")) 'chr(34)で " [半角二重引用符]
    
    Next
    
    'CSV側ヘッダー 0:Order ID/1:Line ID/2:Quantity/3:Product Code/4:Description/5:Option Name/6:Option Value/7:Unit Price/
    'ループ一回目ではヘッダーなので、インクリメントへ飛ぶ
    
    If LineBuf(0) = "Order ID" Then GoTo increment
        
    ':ToDo ここからシート、セルの値のが入るので分割した方がいいかもしれない。
    
    With Worksheets("受注データシート")
        'A列、注文番号
        .Range("A" & i).Value = LineBuf(0)
        
        'C列、受注時の商品コード
        .Range("C" & i).NumberFormatLocal = "@"
        .Range("C" & i).Value = LineBuf(3)
        
        'D列、アドイン実行用に6ケタ化したコード、もしくはJAN
        '空欄がありえるので、ピッキングデータ・振分リストに転記時に空欄判定する
        
        '6ケタならそのまま入れる
        If LineBuf(3) Like "######" Then
            .Range("D" & i).NumberFormatLocal = "@"
            .Range("D" & i).Value = LineBuf(3)
        
        '数字5ケタは頭にゼロを追記
        ElseIf LineBuf(3) Like "#####" Then
            
            .Range("D" & i).NumberFormatLocal = "@"
            .Range("D" & i).Value = "0" & LineBuf(3)
        
        'JANもそのまま入れる
        ElseIf LineBuf(3) Like String(13, "#") Then
            
            .Range("D" & i).NumberFormatLocal = "@"
            .Range("D" & i).Value = LineBuf(3)
        
        End If
        
        'E列：商品名  F列：受注数量  G列：売価
        .Range("E" & i).Value = LineBuf(4)
        .Range("F" & i).Value = LineBuf(2)
        .Range("G" & i).Value = LineBuf(7)
        
        'CSV1行をリード完了
        

        'セット分解 7777始まり
        If .Range("D" & i).Value Like "7777*" Then
            
            Call SetParser.ParseItems(.Range("D" & i))
            
            'ParseItemsで行が挿入されるので、行カウンタをセットし直す
            i = OrderSheet.Range("A1").CurrentRegion.Rows.Count
            
        
        End If
    
        '単体○個セット分解 ハイフン含むコードなら分解処理へ投げる
        
        If .Range("D" & i).Value Like "*-*" Then
        
            Call SetParser.ParseScalingSet(.Range("D" & i))
        
        End If
    
    End With
    
increment:
    i = i + 1

Loop

TS.Close

SetParser.CloseSetMasterBook

End Sub

Sub ReadTyumonH(Path As String)

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

    '注文番号の行を調べる
    '注文番号はDobule型で入っている。CSVはString型、Match関数の返値はDouble型
    
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
    
    '注文者名を記入 オフセットしつつ、該当注文番号の全ての行へ記入
    Do While Range("A" & FindRow).Offset(i, 0).Value = CDbl(Order(0))
        
        Range("A" & FindRow).Offset(i, 1).Value = LineBuf(5)
        i = i + 1
    
    Loop

        
Continue:
    
Loop

TS.Close

End Sub
