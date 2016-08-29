Attribute VB_Name = "ComputeQuantity"
Type Syokon
    Quantity As Long
    Status As String
    VenderCode As String
    
End Type
Sub PutQtyCsv()

'商魂の区分、ヤフーデータのAbstract、在庫限りシート、廃番シートをチェックして、
'ヤフーにアップローする在庫数、Allow-overdraftをセットします。

If Not SecondInventry.Cells(1, 1).Value = "JAN" Then
    
    MsgBox "棚無データなし、処理を継続しますか？"

End If

'時間計測をします

Dim startTime As Long
startTime = Timer

'準備

Call FetchSecondInventry

'Call FetchYahooCSV

'各シートのコード範囲を名前で呼び出せるよう再定義
Call SetRangeName


'---準備完了---

'商魂データから全アイテムに在庫をセット
Call SetQuantity

'ヤフーデータシートからCSVを保存
'Call PutQtyCsv

'終了時刻を入れる
Dim endTime As Long
endTime = Timer

Call ApendProcessingTime(endTime - startTime)

MsgBox "実行時間：" & endTime - startTime & " 秒"

End Sub

Sub buildSecondInventry()

Dim startTime As Long
startTime = Timer

'準備

Call FetchSecondInventry
'各シートのコード範囲を名前で呼び出せるよう再定義
Call SetRangeName

'商魂データから全アイテムに在庫をセット
Call SetQuantity

'終了時刻を入れる
Dim endTime As Long
endTime = Timer

Call ApendProcessingTime(endTime - startTime)

MsgBox "実行時間：" & endTime - startTime & " 秒"

End Sub

Sub SetQuantity()

'PreDecreaseQty：棚卸し前の商魂-実在庫の見込み誤差数
'2月8月は商魂から5を引いてから0.6掛けした在庫数をセットしたい

yahoo6digit.Activate

Application.ScreenUpdating = False


'ヘッダーを追記
yahoo6digit.Range("A1").End(xlToRight).Offset(0, 1).Resize(1, 3) = Array("quantity", "allow-overdraft", "status")

Dim item As New item

Dim colAbstract As Integer, colQuantity As Integer, colAllow As Integer, colStatus As Integer 'なんぞもっとスマートな定義方法があるだろ
colAbstract = yahoo6digit.Rows(1).Find("abstract").Column

colQuantity = yahoo6digit.Rows(1).Find("quantity").Column
colAllow = yahoo6digit.Rows(1).Find("allow-overdraft").Column
colStatus = yahoo6digit.Rows(1).Find("status").Column

Dim r As Range

With yahoo6digit 'With構文内ではオブジェクト参照が繰り返されないため、少しだけ高速になるらしい

    For Each r In .Range("YahooCodeRange")
        
        Set item = New item
        item.Code = r.Value
        
        Dim i As Long  'TODO:行番号を格納するiは要らないのでは…
        i = r.Row
        
        'Debug.Assert i < 1000
        
        '在庫設定除外シートにあれば、以下の処理は行わない、Continueへ飛ぶ
        If WorksheetFunction.CountIf(ExceptQty.Range("ExceptCodeRange"), item.Code) > 0 Then GoTo Continue
        
        'Abstractを拾う。16年1月現在 全商品で使っているので記載有無判定はなし
        item.Abstract = yahoo6digit.Cells(i, colAbstract).Value

               
        '商魂シートから商魂の値を取得、登録無しはSy.Quantity=0
        Dim sy As Syokon
        sy = SyokonMaster.GetSyokonQtyKubun(item.Code)
        
        'Itemオブジェクトに商魂の値をセット
        item.Status = sy.Status
        item.VenderCode = sy.VenderCode
        
        '棚なし、廃番、在庫限りの各シートをチェック
        item.CheckSecondInventry
        item.CheckEol
        item.CheckStockOnly
        
        '設定在庫数のセット、在庫数はスリムスに一本化
        
        If Slims.HasLocation(item.Code) Then
            
            item.Quantity = Slims.getQuantity(item.Code)
        
        Else
            
            item.Quantity = 0
        
        End If
        
        '手配可否をセット
        item.SetAvailablePurchase
        
        '算出した在庫と、判定したAllow-overDraftを書き出す
        
        'Debug.Assert item.Quantity > 0
        
        .Cells(i, colQuantity).Value = item.Quantity
        
        If item.AvailablePurchase Then  'Allow-overdraftはBool値なので1/0に置き換えて出力
            .Cells(i, colAllow).Value = 1
        Else
            .Cells(i, colAllow).Value = 0
        End If
        
        .Cells(i, colStatus).Value = item.Status
       
Continue:
    
       Set item = Nothing
       
    Next r

End With


'一時停止を上書き
Call halt.setHalt

'在廃、処分で0個は廃番・終了へ移動
Call CheckEolInStockOnly

'今回の棚無しアイテムの在庫有りを別シートにコピーしておく
Call StackLastQty

End Sub

Sub UpdateSecondInventryQty()

Call FetchSecondInventry

End Sub

Sub PutCsv()
'FileSystemObjectのテキストストリームでCSVファイルを生成して、TextStreamで内容を流し込みます。
'数秒で終わります。

With yahoo6digit 'ヤフーデータの下準備

    .Activate
    
    '「"登録なし"」と「"空白"」 この2つ以外をフィルターで表示…TODO：1列目からフィルターの状況をチェックさせた方がいい
    '16-2-29 廃番の区分が「メ廃番」になりました。
    
    Dim StatusCol As Integer
    StatusCol = .Rows(1).Find("status").Column
    
    .Range("A1").CurrentRegion.AutoFilter Field:=StatusCol, Criteria1:=Array( _
            "ＳＰ扱い", "限定品", "在庫廃番", "在庫処分", "棚なしに有", "棚なし完売", "直送扱い", "登録のみ", "メ廃番品", "販路限定", "販売中止", "標準" _
            ), Operator:=xlFilterValues
    
    'フィルターしたレンジをセット、CSVのヘッダーは別途書き込んでおくので、2行目以降のレンジ。
    Dim A As Range
    Set A = .Range("C1").CurrentRegion.SpecialCells(xlCellTypeVisible)
    
    Dim B As Range
    Set B = .Range("C2").Resize(Range("C1").SpecialCells(xlCellTypeLastCell).Row - 1, 1)
    
    Dim CodeRange As Range
    Set CodeRange = Application.Intersect(A, B)

End With

'書き出し用CSVを用意
Dim day As String
day = Format(Date, "mm") & Format(Date, "dd")

Dim OutputCsvName As String
OutputCsvName = "商魂在庫アップ用" & day & ".csv"

Dim FSO As Object 'TODO:事前バインディングに変更
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim TS As Object
    
Set TS = FSO.CreateTextFile(Filename:=ThisWorkbook.Path & "\" & OutputCsvName, _
                            OverWrite:=True)
                            
'ヘッダーを書き込み
header = "code,quantity,allow-overdraft"

TS.WriteLine header

colQuantity = yahoo6digit.Rows(1).Find("quantity").Column
colAllow = yahoo6digit.Rows(1).Find("allow-overdraft").Column
colStatus = yahoo6digit.Rows(1).Find("status").Column

'コードレンジに対して、r.rowで行番号を取り出して同じ行のQuantity/Allowの値を取得する
For Each r In CodeRange
    
    Code = r.Value
    
    qty = Cells(r.Row, colQuantity).Value
    pur = Cells(r.Row, colAllow).Value
    
    TS.WriteLine Code & "," & qty & "," & pur

Next

TS.Close

End Sub


Sub StackLastQty()
'前回棚無しをクリアー
'ヤフーデータから棚なしに有をフィルターして前回棚無しにコピー
LastSecondInventry.Cells.Clear

yahoo6digit.Activate

Dim StatusCol As Integer
StatusCol = Rows(1).Find("status").Column

Range("A1").CurrentRegion.AutoFilter Field:=9, Criteria1:="棚なしに有"

'フィルターした表示領域
Dim A As Range
Set A = Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible)

'ヤフーデータ全体のレンジ
Dim B As Range
Set B = yahoo6digit.AutoFilter.Range

'ABの交差レンジをデータレンジとしてセット、棚なしに有商品のレンジ
Dim InSecondInventryRange As Range
Set InSecondInventryRange = Application.Intersect(A, B)

InSecondInventryRange.Copy LastSecondInventry.Range("A1")

End Sub
