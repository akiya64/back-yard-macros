Attribute VB_Name = "FetchOrder"
Option Explicit

Type RawOrder

    Serial As String        'クロスモールで採番する連番
    
    MallId As String        '受注モールコード  01楽天 02ヤフー 03Amazon
    MallName As String      '受注モール名称
    OrderId As String       '各モールの受注番号
    
    Addressee As String     '送り先名
    
    Code As String          '受注時の商品コード
    ProductName As String   'モール掲載の商品名
    Quantity As String      '受注数量
    
    Price As String         '受注金額

End Type

Sub ReadClossMallCsv(Path As String)

'クロスモールの受注CSVを受注データシートに読み込む

Dim FSO As Object
Set FSO = New FileSystemObject

Dim TS As Textstream
Set TS = FSO.OpenTextFile(Path, ForReading)

Dim WriteRow As Long
WriteRow = 1 '項目行は出力しないので、iは1行目から開始
    
Do Until TS.AtEndOfStream
    
' 行を取り出して必要な項目のみを配列に入れ直す
    Dim Col As Variant
    Col = Split(TS.ReadLine, ",")
    
    'ループ一回目ではヘッダーなので、Continue
    If Col(0) = "管理番号" Then GoTo Continue
    
    Dim Order As RawOrder
    
    With Order
        .Serial = Col(0)
        
        .Code = Col(1)
        .ProductName = Col(2)
        .Quantity = Col(3)
        .Price = Col(4)
                
        .MallName = Col(8)
        .Addressee = Col(10)
        .OrderId = Col(13)

    End With
    
    Call WriteSheet(Order)

    '最終行を特定してセット分解
    Dim CurrentCodeCell As Range
    Set CurrentCodeCell = Cells(Range("A1").SpecialCells(xlCellTypeLastCell).Row, 2)
    
    
    '7777始まりセット分解
    If CurrentCodeCell.Value Like "7777*" Then

        Call SetParser.ParseItems(CurrentCodeCell)
    
    End If

    '単体○個セット分解
    If CurrentCodeCell.Value Like "*-*" Then
    
        Call SetParser.ParseScalingSet(CurrentCodeCell)
    
    End If

Continue:

Loop

TS.Close

SetParser.CloseSetMasterBook

End Sub

Sub WriteSheet(ByRef Order As RawOrder)
'注文データの配列を受け取って、最終行の直下へ追記
    With Worksheets("受注データシート")
    
        Dim WriteRow As Long
        WriteRow = .Range("A1").SpecialCells(xlCellTypeLastCell).Row + 1
        
        'A列 クロスモール連番
        .Range("A" & WriteRow).NumberFormatLocal = "@"
        .Range("A" & WriteRow).Value = Order.Serial
        
        'B列、受注時の商品コード
        .Range("B" & WriteRow).NumberFormatLocal = "@"
        .Range("B" & WriteRow).Value = Order.Code
        
        
        'C列：商品名  D列：売価   E列：受注数量  F列：受注番号 G列：モール名  H列：お届け先名
        .Range("C" & WriteRow).Value = Order.ProductName
        .Range("D" & WriteRow).Value = Order.Price
        .Range("E" & WriteRow).Value = Order.Quantity
        .Range("F" & WriteRow).Value = Order.OrderId
        .Range("G" & WriteRow).Value = Order.MallName
        .Range("H" & WriteRow).Value = Order.Addressee
        
        'I列、アドイン実行用に6ケタ化したコード、もしくはJAN
        '空欄がありえるので、ピッキングデータ・振分リストに転記時に空欄判定する
        .Range("I" & WriteRow).NumberFormatLocal = "@"
        
        '6ケタならそのまま入れる
        If Order.Code Like String(6, "#") Then
            .Range("I" & WriteRow).Value = Order.Code
        
        '数字5ケタは頭にゼロを追記
        ElseIf Order.Code Like String(5, "#") Then
            
            .Range("I" & WriteRow).Value = "0" & Order.Code
        
        'JANもそのまま入れる
        ElseIf Order.Code Like String(13, "#") Then
            
            .Range("I" & WriteRow).Value = Order.Code
        
        End If
    
        'J列 必要数量 6ケタ/JANに対して必要な数量
        'セット分解後に書き換えられるので、一旦受注数量を入れる。
        .Range("J" & WriteRow).Value = Order.Quantity
    
    End With
End Sub

