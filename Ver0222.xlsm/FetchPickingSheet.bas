Attribute VB_Name = "FetchPickingSheet"
Option Explicit

'注文番号、注文者名など、見出し名と列番号を格納するHead型を定義
'作成者や作成日によって、列の位置が変わることがあるので、毎回特定する
'基本的には、マニュアルとテンプレートでデータをやりとりすれば不要ではあるが、
'列番号のハードコーディングはしないほうがいい

Private Type Head
    
    Caption As String
    Columns As Long

End Type

'他のモジュールから開いたピッキングシートファイルをクローズするのに必要なパブリック変数
Public IsFileNewOpen As Boolean
Public PickingFileName As String

Function CheckPickingProducts(Optional IsMsgBox As Boolean = True)

'テスト用
'Sub CheckPickingProducts()

'ヤフースカイプ分日付ファイルで、緑に塗られていない=ピッキングできなかった商品を、
'注残シート、センター在庫列に「なし」転記、本日日付で手配かかったとみなして日付入れます。

'本日注文を取り込んでいるかチェック
Application.ScreenUpdating = False
Application.DisplayAlerts = False

If LogSheet.Range("LastFetchNewOrder") <> Date Then
    
    Call fetchOrderCsv.梱包室受注ファイル読込
    LogSheet.Range("B7").Value = Date

End If

Application.DisplayAlerts = True

'「ヤフースカイプ分xx.xlsx」を開く＝梱包室からのチェック済みの商品リストのエクセルファイル開く
'PickingFileName = Range("PickingSheetBaseName") & Format(Date, "mmdd") & ".xlsx"

'フォームのTextBox4が
PickingFileName = OpPanel.TextBox4


Dim PickingFilePath As String

PickingFilePath = Range("PickingSheetFolder").Value

If Right(PickingFilePath, 1) <> "\" Then PickingFilePath = PickingFilePath & "\" '末尾\マークでないとき補完

PickingFilePath = PickingFilePath & PickingFileName


''「ファイルを開く」のフォームでファイルを指定 一応残します
'PickingFilePath = Application.GetOpenFilename("エクセルファイル,*.xls?", , "ヤフーピッキングリストを指定")


Dim WsPicking As Worksheet
Dim wb As Workbook

'ピッキングシートが開いていればそのまま利用する、
'Todo:定型処理で切り出してしまった方がいい、Invoicesクラスにも似たような処理がある
For Each wb In Workbooks
    If wb.Name = PickingFileName Then
        Set WsPicking = wb.Sheets(1)
    End If
Next wb

'ワークブックを開いてセット
If WsPicking Is Nothing Then

    'ネット販売関連の所定のフォルダにピッキングシートがない場合、Exit
    If Not Dir(PickingFilePath) Like "*.xlsx" Then
        
        MsgBox "ピッキングシートの転記ができませんでした。" & vbLf & _
                "ピッキングシートファイルなし。ヤフーピッキングシートのファイル有無、ファイル名を確認" & vbLf & _
                "他の処理は継続して可能です。"

        Exit Function
    
    End If

    Set wb = Workbooks.Open(PickingFilePath)
    Set WsPicking = wb.Sheets(1)
    
    IsFileNewOpen = True
        
End If

'ピッキング対象の商品コードレンジ
Dim LastRow As Integer
LastRow = WsPicking.Range("B1").SpecialCells(xlCellTypeLastCell).Row

'----ピッキングシートのオープン処理完了-----

Dim TodaysOrders As Dictionary
Set TodaysOrders = OrderSheet.getTodaysOrders '本日受注のOrderListを生成


'一旦、全てのアイテムのIsPickingDoneをFalseにセット、OrderObject生成の度にセルを読んで空ならTrueなので"
Dim v As Variant
Dim w As Variant
For Each v In TodaysOrders
    For Each w In TodaysOrders(v).Products
        TodaysOrders(v).Products(w).IsPickingDone = False
    Next
Next


'ピッキングシートの注文番号列、注文者名列、商品コード列、備考列を特定する
'ヘッダー配列を用意
Dim Header(3) As Head
Header(0).Caption = "注文番号"
Header(1).Caption = "届け先名"
Header(2).Caption = "商品コード"
Header(3).Caption = "ロケーション" '備考の取得に使う

Dim h As Integer
For h = 0 To 3
    Header(h).Columns = WsPicking.Rows(1).Find(Header(h).Caption).Column
Next

'検索フォームを戻すために空検索
WsPicking.Rows(1).Find ("")

'ピッキングシートの注文者名は「発送先名」、注残管理は注文者名
'注残管理の名前で探して、なければ「ピッキングシート該当無し」
Dim i As Integer
For i = 2 To LastRow
    
    Dim CurrentId As String
    CurrentId = WsPicking.Cells(i, Header(0).Columns).Value
    
    Dim CurrentBuyerName As String
    CurrentBuyerName = WsPicking.Cells(i, Header(1).Columns).Value
    
    Dim CurrentCode As String
    CurrentCode = WsPicking.Cells(i, Header(2).Columns).Value
        
    Dim CurrentNote As String
    CurrentNote = WsPicking.Cells(i, Header(3).Columns + 1).Value
        
    'コードをヤフーの形式に変換 012345->12345
    If CurrentCode Like "0#####" Then CurrentCode = Right(CurrentCode, 5)
    
    Dim o As order
    
    '背景色が白ではない商品=センター在庫有り、ピッキング可能
    If Not WsPicking.Cells(i, 1).Interior.Color = 16777215 Then
        
        '注文者名から、どの注文の商品か特定
        Set o = FindByBuyerName(CurrentBuyerName, CurrentCode, TodaysOrders)
        
        '注残管理シートは注文者名、ピッキングシートは宛先氏名のため、
        'FindByBuyerNameメソッドで注文の一致がとれず､戻り値が空の場合がある｡
        
        If Not o Is Nothing Then

            o.Products(CurrentCode).IsPickingDone = True
                   
        End If
    
    End If
    
    'H列に何か書いてる＝梱包室で把握している在庫状況
    If CurrentNote <> "" Then
 
       '注文者名から、どの注文の商品か特定
        Set o = FindByBuyerName(CurrentBuyerName, CurrentCode, TodaysOrders)
        
        'その注文の商品オブジェクトにピッキング可のフラグを登録
        o.Products(CurrentCode).CenterStockState = CurrentNote

    End If
    
Next i

'TodaysOrderの各注文の各商品のIsPickingDoneを転記

'Dim w As Variant
'Dim v As Variant

For Each v In TodaysOrders
    For Each w In TodaysOrders(v).Products

        If TodaysOrders(v).Products(w).IsPickingDone = False Then
            
            'ピッキングステータスをシートに転記、Falseだと「なし」＋本日手配扱い
            'Todo:OrderかProductオブジェクトを渡して、OrderSheetに判定して転記させる
            Call OrderSheet.writePickingStatus(CStr(v), CStr(w), TodaysOrders(v).Products(w).CenterStockState)
        
        End If
    
    Next
Next

'チェック日をLogSheetに書きこむ
LogSheet.Range("LastUpdatePickingSheet") = Date

ThisWorkbook.Save

Application.ScreenUpdating = False

If IsMsgBox Then

    MsgBox prompt:="ピッキングファイルの転記完了", Buttons:=vbInformation

End If

End Function

Private Function FindByBuyerName(Name As String, Code As String, OrderList As Dictionary) As order
'注文リスト配列と注文者名を受け取って、Orderオブジェクトを返す。
'注文者名で注文を探して、その受注アイテムProductsに該当コードがあるか判定、
'Orderオブジェクトを返す

'Orderの配列をまず名前で調べて、該当すればProducts内のコードを調べる
Dim v As Variant
For Each v In OrderList

    If OrderList(v).BuyerName = Name Then
        
        Dim w As Variant
        
        For Each w In OrderList(v).Products
            
            If OrderList(v).Products.Exists(Code) Then
            
                Set FindByBuyerName = OrderList(v)
            
                Exit Function
            
            End If
        
        Next w
    
    End If

Next v

End Function

Private Function FindCol(Caption As String, PickingWorkSheet As Worksheet) As Integer



End Function
