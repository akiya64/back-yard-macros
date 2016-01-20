Attribute VB_Name = "Module1"
'ヤフーショッピングWebAPIを使って、JANで検索して価格とショップ名を抜き出す
'GETメソッドで特定JANの価格順など、検索結果をXMLを取得できるので、それをパースしてセルに転記する

'参照設定で、MicroSoft XML,V6.0　ライブラリにチェックを入れること
'エクセルでMXLを扱うためのライブラリ、MSXML2オブジェクトの生成に必要

'ヤフーショッピングWebAPIリファレンス　WebAPIを呼び出すには要アプリケーションコード
'http://developer.yahoo.co.jp/webapi/shopping/shopping/v1/itemsearch.html

'MSDN MSXML2.XMLHTTPオブジェクトの公式リファレンスは下記
'http://msdn.microsoft.com/en-us/library/ms759148%28v=vs.85%29.aspx

'MSDN 初心者のための XML DOM ガイド　更新日時が古いけど基本は変わらないはず
'http://msdn.microsoft.com/ja-jp/library/aa468547.aspx


'宣言セクション

'sleepを使うための宣言
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'全プロシージャで共有する変数
'ショップ/価格リストを記入する開始列
Dim startcolumn_price As Integer

'Yahoo API呼び出しに使う アプリケーションID
Const APP_ID = ""

'janの入っているセル
Dim c As Range

Sub 価格調査()

'店舗名・価格を書き出す最初の列を指定
startcolumn_price = ActiveSheet.UsedRange.Columns.Count + 1

'Janの列を特定します。
'社内システムでは「JANコード」YahooCsvでは「jan」なのでどちらかを調べる
Dim jan_col As Range

If Not Range("1:1").Find("jan") Is Nothing Then
    
    Set jan_col = Range("1:1").Find("jan")

ElseIf Not Range("1:1").Find("JANコード") Is Nothing Then

    Set jan_col = Range("1:1").Find("JANコード")

Else
    
    MsgBox "Janの見出しが見つかりません。" & vbLf & _
           "1行目のヘッダーに「jan」か「JANコード」を指定"
    
    Exit Sub

End If

'JANのリストをレンジで取得します

Dim rng_jan As Range

'rng_janにJANコードのレンジをセット
Set rng_jan = jan_col.Offset(1, 0).Resize(ActiveSheet.UsedRange.Rows.Count - 1, 1)

For Each c In rng_jan
       
    'サーバーアクセスとか、モジュール化orオブジェクト化して並列処理できると速いんでは？
    'このコードでも十分速いけど…ヤフーサーバーのレスポンスが速い、腐ってもWEBポータル
    
    Dim jan As String
    jan = c.Value
    
    'セルから取得したjanが数字13ケタかチェック
    If Not jan Like "#############" Then
        Call writeError("正しくないJAN")
        GoTo continue:
    End If

    Call loadXml(jan)
    
continue:

Next c

End Sub

Private Sub loadXml(jan As String)

 'xmlオブジェクトのインスタンス生成
 Dim xml As Object
 Set xml = CreateObject("MSXML2.DOMDocument")
 
 'サーバーからXMLを読むためのMSXML2オブジェクトの設定
 xml.async = False
 xml.setProperty "ServerHTTPRequest", True
 
 'GETするためのurlを生成します
 Dim url As String
 url = makeUrl(jan)
 
 'サーバーにGETメソッドでアドレスを投げて、XMLを取得します
 xml.Load (url)

 'スリープタイムのカウンタjを初期化
 j = 0
 
 'サーバーからのレスポンス待ち　Sleep100しながら待機　ビジーウェイト
 Do
     DoEvents
     Sleep 10
     j = j + 1
     
     If j > 100 Then  '10msec*100=1秒応答がなければループアウト
         Cells(c.Row, startcolumn_price).Value = "サーバー応答なし"
         Exit Do
     End If
         
 Loop While xml.readyState <> 4
         
 'xmlオブジェクトがXMLを取得できてなければ、Continue
 
 If Not xml.HasChildNodes Then
     Call writeError("結果が正しく取得できませんでした")
     Exit Sub
 End If

    
'xmlのResultSetからHitを取り出す、簡単なツリー構造
'ResultSet>Result>Hit>Store>Name
'                    >Price

'jan引数が指定なしだと<Error><Message>BadRequestが返ってくるので、Resultsetがあるかチェック
If xml.getElementsByTagName("ResultSet").Length > 0 Then
    
    'ResultSetのTotalResultAvailable/TotalResult属性
    '現状では特にセルに書き戻さないが、有効なHit要素数のチェックと、
    '注文可能な店舗数・掲載店舗数が把握できるので変数に格納する
    'TotalResultsReturned=0だと空のHit要素が1個返ってくる
    
    Dim total_results_counts As Integer
    total_results_counts = xml.SelectSingleNode("ResultSet").Attributes.getNamedItem("totalResultsReturned").Text
    
    Dim available_results_counts As Integer
    available_results_counts = xml.SelectSingleNode("ResultSet").Attributes.getNamedItem("totalResultsAvailable").Text
    
    If total_results_counts > 0 Then
        
        Call parseWriteRanking(xml.SelectNodes("ResultSet/Result/Hit"))
    
    Else
        'totalRsultsAvailableが0
        Call writeError("掲載ショップなし")
        Exit Sub
    
    End If
    
Else

    'Resultsetがない
    Call writeError("該当JANがヤフーに登録なし")
    Exit Sub
    
End If

End Sub

Private Function makeUrl(jan As String)
'janを渡してもらって、GETメソッドで投げるURLを生成

Dim base_url As String 'WEB APIをGETで呼び出すベースURL
base_url = "http://shopping.yahooapis.jp/ShoppingWebService/V1/itemSearch"

Dim sort As String
sort = "%2Bprice" '価格順、＋－で降順・昇順指定できる、URLエンティティに変換が必要

get_url = base_url
get_url = get_url & "?appid=" & APP_ID 'アプリケーションID
get_url = get_url & "&sort=" & sort    'ソート種別
get_url = get_url & "&hits=5"          '最大数5、最安5件で
get_url = get_url & "&jan=" & jan      'JANをセット
    
makeUrl = get_url

End Function

Private Sub parseWriteRanking(hits As Object)
'<Hit>ノードリストhitsをイテレートしつつセルに書き込み

'hを書き換えながらイテレートするので、hitノードを格納できる変数hをセット、
'hはXML DOM ElementかNodeクラス、NodeListは複数ノードを格納するクラス
Dim h As Object

'書き戻しセルの列カウンターkを設定
Dim k As Integer
k = 0

For Each h In hits

    store_name = h.SelectSingleNode("Store/Name").Text              '各HitノードのStore>Name＝ショップ名
    
    Cells(c.Row, startcolumn_price + k).Value = store_name          '列を+1しながらショップ・価格・ショップ・価格の順でCellに記入
    
    k = k + 1
            
    sale_price = h.SelectSingleNode("Price").Text                   '各HitノードのPrice＝販売価格
    Cells(c.Row, startcolumn_price + k).Value = sale_price
    k = k + 1
    
Next h

End Sub

Private Sub writeError(s As String)
'エラーメッセージをセルに返す

Cells(c.Row, startcolumn_price).Value = s

End Sub
