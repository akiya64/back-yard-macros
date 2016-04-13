Attribute VB_Name = "ComputeQuantity"
'棚卸し前の社内システム-実在庫の見込み誤差数、2月8月はこの数を減じてから在庫数を丸める
Public PreDecreaseQty As Long

Type Syokon
    Quantity As Long
    Status As String
    VenderCode As String
    
End Type
Sub BuildQtyCsv()

'社内システムの区分、ヤフーデータのAbstract、在庫限りシート、廃番シートをチェックして、
'ヤフーにアップローする在庫数、Allow-overdraftをセットします。

'社内システムデータのDLはOKか、確認Popup
go = MsgBox(prompt:="社内システムデータが用意できていますか？" & vbLf & "また、ヤフーデータの前回実行分はクリアしてよろしいですか？", Buttons:=vbYesNo)

If go <> vbYes Then
    MsgBox "処理を終了します。"
    End
End If

'処理前のデータチェック
If Not SyokonMaster.Cells(1, 1).Value = "Code" Then
    
    MsgBox "社内システムデータなし、処理を終了します。"
    Exit Sub

End If

If Not SecondInventry.Cells(1, 1).Value = "JAN" Then
    
    MsgBox "ネット用在庫データなし、処理を継続しますか？"

End If

'時間計測をします
'7万6000行の処理で420秒ぐらい

Dim startTime As Long
startTime = Timer

'準備

Call FetchSecondInventry

Call FetchYahooCSV

'各シートのコード範囲を名前で呼び出せるよう再定義
Call SetRangeName


'---準備完了---

'社内システムデータから全アイテムに在庫をセット
Call SetQuantity

'ヤフーデータシートからCSVを保存
Call PutQtyCsv

'終了時刻を入れる
Dim endTime As Long
endTime = Timer

Call ApendProcessingTime(endTime - startTime)

MsgBox "実行時間：" & endTime - startTime & " 秒"

End Sub

Sub buildSecondInventryQty()

Dim startTime As Long
startTime = Timer

'準備

Call FetchSecondInventry
'各シートのコード範囲を名前で呼び出せるよう再定義
Call SetRangeName

'社内システムデータから全アイテムに在庫をセット
Call SetQuantity

'終了時刻を入れる
Dim endTime As Long
endTime = Timer

Call ApendProcessingTime(endTime - startTime)
With yahoo6digit

    .Activate
       
    Dim StatusCol As Integer
    StatusCol = .Rows(1).Find("status").Column
    
    .Range("A1").CurrentRegion.AutoFilter Field:=StatusCol, Criteria1:=Array( _
            "ネット用在庫に有", "ネット用在庫完売" _
            ), Operator:=xlFilterValues
    
    'フィルターしたレンジをセット
    Dim A As Range
    Set A = .Range("C1").CurrentRegion.SpecialCells(xlCellTypeVisible)
    
    Dim B As Range
    Set B = .Range("C2").Resize(Range("C1").SpecialCells(xlCellTypeLastCell).row - 1, 1)
    
    Dim CodeRange As Range
    Set CodeRange = Application.Intersect(A, B)

End With

'書き出し用CSVシートを用意
Worksheets("CSV").Cells.Clear

'ヘッダーを書き込み
header = Array("code", "quantity", "allow-overdraft")

Worksheets("CSV").Range("A1:C1") = header

Worksheets("ヤフーデータ").Activate

colQuantity = yahoo6digit.Rows(1).Find("quantity").Column
colAllow = yahoo6digit.Rows(1).Find("allow-overdraft").Column

Dim i As Long
i = 2
'コードレンジに対して、r.rowで行番号を取り出して同じ行のQuantity/Allowの値を取得する
For Each r In CodeRange

    Code = r.Value

    Qty = Cells(r.row, colQuantity).Value
    pur = Cells(r.row, colAllow).Value

    Worksheets("CSV").Range("A" & i & ":C" & i) = Array(Code, Qty, pur)

    i = i + 1

Next

Worksheets("CSV").Activate

'CSV追記準備
Dim FSO As New FileSystemObject
Dim Csv As Object

'追記モード ForAppending でファイルを開く
Set Csv = FSO.OpenTextFile(Filename:=ThisWorkbook.Path & "\" & "ヤフー在庫更新0413.csv", IOMode:=8)


For i = 94 To 709
    
    With Worksheets("Csv")
        Csv.WriteLine (CStr(.Cells(i, 1).Value) & "," & CStr(.Cells(i, 2).Value) & "," & CStr(.Cells(i, 3).Value))
    End With

Next

MsgBox "実行時間：" & endTime - startTime & " 秒"

End Sub

Sub SetQuantity()

'PreDecreaseQty：棚卸し前の社内システム-実在庫の見込み誤差数
'2月8月は社内システムから5を引いてから0.6掛けした在庫数をセットする

PreDecreaseQty = 0

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
        i = r.row
        
        
        'Debug.Assert i < 2902

        
        '在庫設定除外シートにあれば、以下の処理は行わない、Continueへ飛ぶ
        If WorksheetFunction.CountIf(ExceptQty.Range("ExceptCodeRange"), item.Code) > 0 Then GoTo continue
        
        'Abstractを拾う。16年1月現在 全商品で使っているので記載有無判定はなし
        item.Abstract = yahoo6digit.Cells(i, colAbstract).Value

               
        '社内システムシートから社内システムの値を取得、登録無しはSy.Quantity=0
        Dim sy As Syokon
        sy = SyokonMaster.GetSyokonQtyKubun(item.Code)
        
        'Itemオブジェクトに社内システムの値をセット
        item.Status = sy.Status
        item.VenderCode = sy.VenderCode
        
        'ネット用在庫、廃番、在庫限りの各シートをチェック
        item.CheckSecondInventry
        item.CheckEol
        item.CheckStockOnly
        
        
        '設定在庫数を算出セット、ネット用在庫/社内システムをスイッチして渡す。
        
        If item.IsSecondInventry Then
            item.SetQuantity (item.SecondInventryQuantity)
        Else
            item.SetQuantity (sy.Quantity)
        End If
        
        '発注可否をセット
        item.SetAvailablePurchase
        
        '算出した在庫と、判定したAllow-overDraftを書き出す
        
        .Cells(i, colQuantity).Value = item.Quantity
        
        If item.AvailablePurchase Then  'Allow-overdraftはBool値なので1/0に置き換えて出力
            .Cells(i, colAllow).Value = 1
        Else
            .Cells(i, colAllow).Value = 0
        End If
        
        .Cells(i, colStatus).Value = item.Status
       
continue:
    
       Set item = Nothing
       
    Next r

End With


'一時停止を上書き
Call halt.setHalt

'在廃、処分で0個は廃番・終了へ移動
Call CheckEolInStockOnly

'今回のネット用在庫アイテムの在庫有りを別シートにコピーしておく
Call StackLastQty

'次回実行時に社内システムデータDLしてないとエラー出るよう、社内システムの今回データを移動

SyokonMaster.Range("A1").CurrentRegion.Cut Destination:=SyokonMaster.Range("K1")


End Sub

Sub UpdateSecondInventryQty()

Call FetchSecondInventry

End Sub

Sub PutQtyCsv()
'FileSystemObjectのテキストストリームでCSVファイルを生成して、TextStreamで内容を流し込みます。
'数秒で終わります。

With yahoo6digit 'ヤフーデータの下準備

    .Activate
    
    '「"登録なし"」と「"空白"」 この2つ以外をフィルターで表示…TODO：1列目からフィルターの状況をチェックさせた方がいい

    
    Dim StatusCol As Integer
    StatusCol = .Rows(1).Find("status").Column
    
    .Range("A1").CurrentRegion.AutoFilter Field:=StatusCol, Criteria1:=Array( _
             "在庫廃番", "在庫処分", "ネット用在庫有", "ネット用在庫完売", "直送扱い", "登録のみ", "廃番品", "販売中止", "標準" _
            ), Operator:=xlFilterValues
    
    'フィルターしたレンジをセット、CSVのヘッダーは別途書き込んでおくので、2行目以降のレンジ。
    Dim A As Range
    Set A = .Range("C1").CurrentRegion.SpecialCells(xlCellTypeVisible)
    
    Dim B As Range
    Set B = .Range("C2").Resize(Range("C1").SpecialCells(xlCellTypeLastCell).row - 1, 1)
    
    Dim CodeRange As Range
    Set CodeRange = Application.Intersect(A, B)

End With

'書き出し用CSVを用意
Dim day As String
day = Format(Date, "mm") & Format(Date, "dd")

Dim OutputCsvName As String
OutputCsvName = "社内システム在庫アップ用" & day & ".csv"

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
    
    Qty = Cells(r.row, colQuantity).Value
    pur = Cells(r.row, colAllow).Value
    
    TS.WriteLine Code & "," & Qty & "," & pur

Next

TS.Close

End Sub


Sub StackLastQty()
'前回ネット用在庫をクリアー
'ヤフーデータからネット用在庫に有をフィルターして前回ネット用在庫にコピー
LastSecondInventry.Cells.Clear

yahoo6digit.Activate

Dim StatusCol As Integer
StatusCol = Rows(1).Find("status").Column

Range("A1").CurrentRegion.AutoFilter Field:=StatusCol, Criteria1:="ネット用在庫に有"

'フィルターした表示領域
Dim A As Range
Set A = Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible)

'ヤフーデータ全体のレンジ
Dim B As Range
Set B = yahoo6digit.AutoFilter.Range

'ABの交差レンジをデータレンジとしてセット、ネット用在庫に有商品のレンジ
Dim InSecondInventryRange As Range
Set InSecondInventryRange = Application.Intersect(A, B)

InSecondInventryRange.Copy LastSecondInventry.Range("A1")

End Sub
