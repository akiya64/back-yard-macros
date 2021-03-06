function FetchSlimsCsv {
    $Slims = Return Get-content ~\desktop\SISJM.csv | ConvertFrom-Csv
    return = $Slims
}

function FetchclossMallCsv {
    $Crossmall = Get-content ~\desktop\order.csv | ConvertFrom-Csv
    return $Crossmall

}

function GetOrderedProducts($OrderId) {
# クロスモールCSVから管理番号が一致する行を取得
    $RecordSet = $Crossmall | Where-Object { $_.管理番号 -like $OrderId }

    return $RecordSet 

}

function GetSlimsInventry($Code) {
# ロケーションと在庫数を判定して、在庫数を返す
# Slimsにデータなしは在庫ゼロとみなす
    
    # 受注時商品コードでハイフンがあれば、それより後ろは削除
    $Code = $Code -replace "\-.",""
    
    if ($Code.length -le 7) {
        $Result = $Slims | Where-Object { $_.WSHOCD -eq [int]$Code }
    
        [Long]$Qty = 0
    
        if ($null -eq $Result) {
            Return $Qty = 0
        } else {
            # ロケーション集約データでレコードが2個ある時にキャストエラーがでるので考える
            Return $Qty = $Result.WKSBQT                
        }
    } else {
        # 5ケタ・6ケタでないコードはSlims在庫 0で返す
        Return $Qty = 0
    }
    
}

function AllowShipping($RecordSet) {
    # 商品配列に対して、全ての商品がSLIMS在庫有りなら、Trueを返す
    
    # 商品コードをキーとして、在庫数を格納するハッシュを作成
    $Products = New-Object 'System.Collections.Generic.Dictionary[string, long]'
    
    foreach ($Record in $RecordSet){
        $Code = $Record.商品コード        
        $Qty = GetSlimsInventry ($Code)            
        
        # 同一商品で行が分かれる場合があるので、キー重複時はそのままContinueする
        try {
            $Products.add($Code, $Qty)            
        } 
        catch {
            continue
        }
    }

    # 1点でも在庫数0の商品があれば、False
    if ( $Products.ContainsValue(0) ) {
        return $false
    } else {
        return $true
    }

}

function UpdateYamato($Csv) {
# ヤマトの出荷可能なお客様側管理番号に対して、出荷予定日を本日で更新する

    # 処理時間確認用のEcho
    Get-date -Format "HH:mm:ss"
    
    $TodayDate= get-date -Format "yyyy/M/d"    

    $Script:Yamato = Get-Content $Csv.Fullname | ConvertFrom-Csv

    $Script:Yamato | ForEach-Object {

        $OrderedProducts = GetOrderedProducts($_.お客様管理番号)
            
        if (AllowShipping($OrderedProducts)) {
            $Yamato | Where-Object {$_.お客様管理番号 -like $OrderId} | foreach { $_.出荷予定日 = $TodayDate}
        }

    }

    $OutPutPath = "~\desktop\test_data_yamato_" + $Csv.Name.Replace(".csv","_today") + ".csv" 
    $yamato | Export-Csv $OutPutPath -Encoding default -noType

}

function UpdatePostal($Csv) {
# ゆうパックの出荷予定日を本日で更新する

    # 処理時間確認用のEcho
    Get-date -Format "HH:mm:ss"
    
    $TodayDate= get-date -Format "yyyy/M/d"    

    $Script:Postal = Get-Content $Csv.Fullname | ConvertFrom-Csv
    
    $Script:Postal | ForEach-Object {

        $OrderedProducts = GetOrderedProducts($_.お客様側管理番号)
            
        if (AllowShipping($OrderedProducts)) {
            $Postal | Where-Object {$_.お客様側管理番号 -like $OrderId} | foreach { $_.発送予定日 = $TodayDate}
        }

    }

    $OutPutPath = "~\desktop\test_data_postal_" + $Csv.Name.Replace(".csv","_today") + ".csv" 
    $Postal | Export-Csv $OutPutPath -Encoding default -noType

}

$BasePath = "\\server02\商品部\ネット販売関連\梱包室データ\送り状データ\"
$today = Get-Date -Format "MMdd"
$path = $BasePath + $today + "\出荷データ"

$Script:Slims = FetchSlimsCsv
$Script:Crossmall = FetchClossmallCsv

GetSlimsInventry("10269-2")
Read-Host "終了するにはENTERキーを押して下さい" 

echo "ヤマト処理中"
$Script:i = 1

$AllYamatoCsv = ls $path -filter *ヤマト*.csv
ForEach ($Csv in $AllYamatoCsv) {

    UpdateYamato($Csv)

}

echo "ゆうパック処理中"
$i = 1

$AllPostalCsv = ls $path -filter *ゆうパック*.csv
ForEach ($Csv in $AllPostalCsv) {
    UpdatePostal($Csv)
}

Echo "ヤマト・ゆうパック処理完了"
Get-date -Format "HH:mm:ss"

Read-Host "終了するにはENTERキーを押して下さい" 