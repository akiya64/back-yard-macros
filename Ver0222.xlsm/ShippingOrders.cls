VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShippingOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const SERVER02_YAHOO_SHIPPMET_PATH As String = "\\server02\商品部\ネット販売関連\出荷通知\出荷通知_ヤフー\"
Const CSV_BASE_NAME As String = "ヤフー送り状番号一括"

Public Orders As Dictionary

Public Sub createShippingList(UndispatchOrders As Dictionary, Invoices As Invoices)
'Orderの入っているDictionaryとInvoicesオブジェクトを受け取って、
'InvoiceListの注文ID1件ずつについて、対応する注文番号が未発送リストにあるか調べる
'あれば、発送リストに加えて、送り状番号、発送日を代入

Set Orders = New Dictionary

Dim v As Variant
For Each v In Invoices.InvoiceNumberList
    
    Dim Id As String
    Id = Invoices.InvoiceNumberList(v).Id

    'undispatchOrderリストに該当OrderIdがあれば、発送リストに加える
    If UndispatchOrders.Exists(Id) Then
        
        On Error Resume Next
            
            Orders.Add Id, UndispatchOrders(Id)
        
        On Error GoTo 0
    
    Else
    
    'Todo:出荷詳細の注文番号が、注残リストにない場合
    
        
    End If
    
    '発送リストの該当注文番号に、送り状番号と発送日を代入
    If Orders.Exists(Id) Then
        
        With Orders(Id)
            
            If .InvoiceNumber = "" Then
               .InvoiceNumber = Invoices.InvoiceNumberList(Id).InvoiceNumber
               .ShippingDate = Invoices.InvoiceNumberList(Id).ShippingDate
            End If
        
        End With
    
    End If
     
Next

'発送リストが出来ているかチェック
If Orders.Count = 0 Then
    
    MsgBox Prompt:="本日の出荷はありません。" & vbLf & "注残管理シート、出荷詳細ファイルを確認してください。" _
    , Buttons:=vbExclamation
    
    End

End If

End Sub

Public Sub putCsv()
'OrdersをヤフーにアップロードできるCSV形式で書きだします。

'1．FSOオブジェクトでcsvファイルを用意
'2．テキストストリームオブジェクトで行追記
'参考：http://www.asahi-net.or.jp/~ef2o-inue/vba_o/sub05_110_050.html

Dim day As String
day = Format(Date, "mmdd")

Dim DefaultOutputDir As String
DefaultOutputDir = Environ("HOMEDRIVE") & Environ("HOMEPATH") & "\Desktop\ヤフー"

'保存先フォルダのチェック
Dim OutputPath As String

If Dir(DefaultOutputDir, vbDirectory) = "ヤフー" Then 'ヤフーフォルダがデスクトップにある時

    OutputPath = DefaultOutputDir & "\" & CSV_BASE_NAME & day & ".csv"


Else
    
    MsgBox "デスクトップ内に「ヤフー」フォルダがありません。" & vbLf & "デスクトップへ保存します。"
    OutputPath = Replace(DefaultOutputDir, "ヤフー", "") & CSV_BASE_NAME & day & ".csv"

End If

'OutputPath = SERVER02_YAHOO_SHIPPMET_PATH & CSV_BASE_NAME & day & ".csv"


Dim FSO As FileSystemObject
Set FSO = New FileSystemObject

Dim res As Integer

'本日のファイルがあるかチェック
If FSO.FileExists(OutputPath) Then
    res = MsgBox("ヤフー送り状番号一括" & day & ".csv" & vbLf & "既に存在します、上書きしますか？", _
    vbYesNo + vbExclamation)
    
    If res <> 6 Then

        MsgBox Prompt:="処理を中止しました。", Buttons:=vbInformation
        End

    End If

End If

Dim TS As Object
Set TS = FSO.CreateTextFile(FileName:=OutputPath, OverWrite:=True)

'ヘッダーを用意して記入
Dim Header As String
Header = "OrderId,ShipMethod,ShipInvoiceNumber1,ShipDate,ShipStatus"

TS.WriteLine Header

'disableFlgをチェックしつつCSVへ出力
Dim v As Variant
For Each v In Orders
    
    '発送不可フラグが真なのに、送り状番号があった場合は注残シートに「発送除外」を書き戻し
    If Orders(v).CanNotShipping Then
        
        Call OrderSheet.writeAssertIvoiceNumber(Orders(v))
    
    End If
        
    '1行分の文字列を作る
    Dim RowText As String
    
    RowText = Orders(v).Id
    RowText = RowText & "," & "postage1"
    RowText = RowText & "," & Orders(v).InvoiceNumber
    RowText = RowText & "," & Format(Orders(v).ShippingDate, "yyyy/mm/dd")
    RowText = RowText & "," & "2"
    
    'CSVへ追記
    TS.WriteLine RowText
    
Next

TS.Close

End Sub
