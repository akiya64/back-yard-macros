# 出荷一覧詳細.csvからヤフーへの一括アップロード用CSVを生成します。

#佐川の送り状システムのファイル出力先フォルダへ移動
Set-Location "\\Server02\出荷通知\" 

#送り状システムのダンプするデータにはヘッダーがないためWPSで処理できない。ヘッダー用の文字列配列を定義
$Header = "受注日","受注受付日","発送日","何かの番号","送り状番号","受注番号","空列1","注文者郵便番号","注文者住所","空列2","注文者電話番号","注文者名","受注番号-2","空列3","届け先郵便番号","届け先住所１","空列4","届け先電話番号","届け先名","送り状種別番号","送り状","明細番号","処理状況","商品コード","商品名","単価","数量","小計","空列5","空列6","注文番号"

#ヤフーの注文番号フォーマットのレコードのみを抽出する
$TodayInvoices = Get-Content .\出荷一覧詳細.csv
$YahooInvoices = $TodayInvoices | ConvertFrom-Csv -Header $Header | where {$_.注文番号 -like "100*" -and $_.送り状番号 -ne ""} | sort 送り状番号 -Unique | Select-Object 注文番号,送り状番号

#日付文字列を作成する

#ファイルとして書き出すための配列を作成
$UpdateOrders = New-Object System.Collections.ArrayList

#YahooInvoicesをイテレートして、UpdateOrdersに追記

$YahooInvoices | ForEach {

    $Order = New-Object PSObject | Select-Object OrderId,ShipMethod,ShipInvoiceNumber1,ShipDate,ShipStatus

    $Order.OrderId = $_.注文番号
    $Order.ShipMethod = "postage1"
    $Order.ShipInvoiceNumber1 = $_.送り状番号
    $Order.ShipDate = Get-date -Format "yyyy/MM/dd"
    $Order.ShipStatus = "2"
    
    [void]$UpdateOrders.add($Order)
}

#CSVファイルとしてUpdateFileを出力

$OutputFolder = $HOME + "\Desktop\ヤフー\"
$TodayDate = Get-Date -Format "MMdd"

If (Test-Path $OutputFolder) {

    $OutputFullPath = $OutputFolder + "ヤフー送り状番号一括" + $TodayDate + ".csv"

}else{
    $OutputFolder = $OutputFolder -replace "ヤフー\\"
    $OutputFullPath = $OutputFolder + "ヤフー送り状番号一括" + $TodayDate + ".csv"
}

$UpdateOrders | Export-Csv $OutputFullPath -Encoding Default -NoTypeInformation
