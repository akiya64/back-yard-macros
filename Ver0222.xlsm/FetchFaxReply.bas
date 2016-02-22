Attribute VB_Name = "FetchFaxReply"
Option Explicit

'FAX納期回答リストのあるディレクトリ、最後必ず\マーク
Const PURCHASE_LOG_FOLDER As String = "\\Server02\商品部\ネット販売関連\発注関連\半自動発注バックアップ\"
Const REPLY_XLSM_FILE As String = "FAX納期回答リスト.xlsm"
Const REPLY_SHEET_NAME As String = "納期リスト"

Dim RangeFaxReplyCode As Range

'Sub FetchFaxReply() '単体テストのためのSub切り替えコメントアウト、Functionにしてるとマクロに表示しないので
Function FetchFaxReply(Optional arg As Variant = "")

'注残「アイテム」リストを作る、注文番号・コード・発注日
'返信FAXのファイルを開く
'コードでFind＝検索をして探す
'Yか+があればヤフーの注文分
'日付でない＝メーカー状況に転記、入荷日があれば転記

Application.ScreenUpdating = False

'手配中商品Dictionaryを作成
Dim CurrentPurchase As Dictionary
Set CurrentPurchase = OrderSheet.getCurrentPurchase

'FAX納期回答リストを開く


Dim wb As Workbook
Dim WsFaxReply As Worksheet

'ワークブックを開いていればそれを使う
For Each wb In Workbooks
    
    If wb.Name = REPLY_XLSM_FILE Then
    
        Set WsFaxReply = wb.Sheets(REPLY_SHEET_NAME)

    End If
    
Next wb

'ワークブックを開いてセット
If WsFaxReply Is Nothing Then

    Set wb = Workbooks.Open(PURCHASE_LOG_FOLDER & REPLY_XLSM_FILE)
    Set WsFaxReply = wb.Sheets(REPLY_SHEET_NAME)
            
End If

Workbooks(REPLY_XLSM_FILE).Activate

'If WsFaxReply.AutoFilterMode = True Then WsFaxReply.Range("A1").AutoFilter

'FAX返信リストの商品コードレンジ

Set RangeFaxReplyCode = WsFaxReply.Range("I2").Resize(WsFaxReply.Range("I2").CurrentRegion.Rows.Count, 1)

'FAX納期回答リストを開いて、手配済み商品リストを取得完了
Dim v As Variant

For Each v In CurrentPurchase
    
    Dim FirstFoundCell As Range
    Set FirstFoundCell = RangeFaxReplyCode.Find(CurrentPurchase(v).Code)
       
    '注残リストに該当商品コードがなければ次のProductへ
    If FirstFoundCell Is Nothing Then GoTo continue
                  
    Call FindArrivalDate(CurrentPurchase(v), FirstFoundCell)
        
continue:

Next v

'For Each v In CurrentPurchase
'    Debug.Print CurrentPurchase(v).OrderId & ":" & CurrentPurchase(v).Code
'Next v

For Each v In CurrentPurchase
    Call OrderSheet.WriteEstimatedArrivalDate(CurrentPurchase(v))
Next v

'FAX納期回答リストを閉じる、開きっぱなしだとエクセルが重すぎる。
'2015-09-15時点でファイルが4MBぐらいある
Workbooks(REPLY_XLSM_FILE).Close SaveChanges:=False

Call 未発送のみ表示

ThisWorkbook.Save

Application.ScreenUpdating = True

MsgBox prompt:="返信リスト読込完了"

End

End Function

Private Sub FindArrivalDate(Product As Product, FoundCell As Range)
'単体Finder
'返信FAXの返信記載列をハードコーディングしているので、注意
'引数 FoundCellのレンジでいいのだろうか？なんか変

Dim FirstFoundCellAddress As String
FirstFoundCellAddress = FoundCell.Address

Do
    
    'Debug.Assert FoundCell.Address <> "$I$252" '特定アドレスの時に停止
    'Debug.Print FoundCell.Address
    
    Dim PurDate As String, Identifier As String, VenderReply As Variant, EstimatedArrivalDate As Variant
    
    PurDate = CStr(Range("F" & FoundCell.Row).Value)
    Identifier = Range("E" & FoundCell.Row).Value
    VenderReply = Range("W" & FoundCell.Row).Value
    EstimatedArrivalDate = Range("Y" & FoundCell.Row).Value
            
    'ヤフー注残とFAX返信リストの商品が一致するとみなす条件
    If PurDate = Format(Product.PurchaseDate, "mdd") Then '発注日が一致する
        
        If InStr(Identifier, "Y") > 0 Or InStr(Identifier, "+") > 0 Then 'モール識別子が「Yを含む」か「+を含む」
            
            'メーカー返信内容が日付でない場合、メーカー状況プロパティに転記
            If Not IsDate(VenderReply) Then 'IsDate関数は日付型に変換可能かを判定するらしく、厳密な型検査ではない
                Product.VenderStatus = VenderReply
            End If
            
            '入荷予定日が日付型なら、入荷予定プロパティに転記
            If IsDate(EstimatedArrivalDate) Then
                Product.EstimatedArrivalDate = EstimatedArrivalDate
            End If
            
            Exit Do
            
        End If
    
    End If
    
    '次を検索する、最初の検索行と一致するまでLoop継続
    Set FoundCell = RangeFaxReplyCode.FindNext(FoundCell)
    
    If FoundCell Is Nothing Then Exit Do
    
Loop Until FirstFoundCellAddress = FoundCell.Address

'参照渡しでオブジェクトもらってるので、値の返却は不要

End Sub
