Attribute VB_Name = "PutYahooInvoiceFile"
Option Explicit

Sub 送り状番号一括ファイル作成()
'注残管理シートの未発送を元に、出荷詳細CSVから処理実績一括アップロード用CSVを生成します。
'2015年5月19日作成

'2015/6/2、クラスモジュールを使ってリファクタリングしました
'2015/6/3、送り状番号を出荷詳細一覧.xlsxから取得するように変更しました。
'2015/10/22 進捗表示を付けました。型チェックもします。
'2015/12/02 出荷通知除外機能はなくてもいいかもしれない。商魂在庫を判定して、佐川の送り状システムへの送り込みデータ作成してるらしいので。
            '現状､間違って送り状作成したとかが､どこでもチェックされていないので､出荷通知除外は引き続き機能させます｡
'2016/1/6 ↑この商魂在庫がなければ送り状起票しないマクロ機能してない時がある。
            
'-------------------------------------切り取り線-----------------------------------------------

OpPanel.Hide

Application.ScreenUpdating = False

'プログレスバーの準備 表示方法はこのサイトがよくまとまってる http://hideprogram.web.fc2.com/vba/vba_ProgressBarForm.html
ShippingFileProgress.ProgressBar.Min = 1
ShippingFileProgress.ProgressBar.Max = 5

'進捗ウィンドウの状況表示をセット
ShippingFileProgress.ProgressBar.Value = 1
ShippingFileProgress.ShowCurrentProcess.Caption = "受注取込/ピッキング取込 チェック中"

'進捗ウィンドウを表示 モードレス指定だとバックグラウンドで処理が進む
ShippingFileProgress.Show vbModeless

'本日の受注、ピッキングシートが転記済かチェックします。

If LogSheet.Range("LastUpdatePickingSheet").Value <> Date Then

    On Error Resume Next 'ピッキングシートファイルが開けなくても続行、送り状ファイル生成に必須ではないので
        
        Call CheckPickingProducts(IsMsgBox:=False)
   
    On Error GoTo 0
    
    LogSheet.Range("B9").Value = Date
    
    If FetchPickingSheet.IsFileNewOpen Then Workbooks(FetchPickingSheet.PickingFileName).Close

End If

'本日分の送り状番号配列を作成 一カ所だけShippingFileProgressの更新処理をInvoces内でやってます
Dim TodaysInvoices As Invoices
Set TodaysInvoices = New Invoices

TodaysInvoices.fetchReportXlsx

ShippingFileProgress.ProgressBar.Value = 3
ShippingFileProgress.ShowCurrentProcess.Caption = "未発送注文 リスト作成中"

'未発送注文の配列を作成
Dim TodaysUndispatch As Dictionary
Set TodaysUndispatch = OrderSheet.getUndispatchOrders

'未発送注文dictionaryが出来ているかチェック
If TodaysUndispatch.Count = 0 Then
    
    MsgBox prompt:="未出荷注文は0件です。" & vbLf & "注残管理シートを確認してください。" & vbLf _
                    & "Dictionary.count = 0 in ""OrderSheet""" _
                    , Buttons:=vbExclamation
    End

End If

ShippingFileProgress.ProgressBar.Value = 4
ShippingFileProgress.ShowCurrentProcess.Caption = "未発送注文の送り状番号を転記中"

Dim TodaysShipping As ShippingOrders
Set TodaysShipping = New ShippingOrders

Call TodaysShipping.createShippingList(TodaysUndispatch, TodaysInvoices)

ShippingFileProgress.ProgressBar.Value = 5
ShippingFileProgress.ShowCurrentProcess.Caption = "一括アップロード用ファイルを保存中"

TodaysShipping.putCsv

ShippingFileProgress.Hide

ThisWorkbook.Save

Call 発送列の空欄のみ表示

Application.ScreenUpdating = True

MsgBox prompt:="ヤフー送り状番号一括" & Format(Date, "mmdd") & "   保存しました。" & vbLf _
                & "ゆうパケット発送分は手動で入力をお願いします。" _
                , Buttons:=vbInformation

End

End Sub
