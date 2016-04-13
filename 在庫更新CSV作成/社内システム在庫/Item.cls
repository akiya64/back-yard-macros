VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Code As String           'ヤフーの商品情報
Public Abstract As String
Public LastUpdateQuantity As Integer

Public Quantity As Integer 'ヤフーショッピングに設定する在庫数はMAX100
Public AllowOrverDraft As Integer

Public AvailablePurchase As Boolean

Public Status As String  '商魂の情報を格納するメンバー

Public VenderCode As String
Public SyokonQuantity As Integer

Public SecondInventryQuantity As Integer '第二倉庫の棚なし商品のメンバー
Public IsSecondInventry As Boolean

Private IsEol As Boolean   'このエクセルで控えている商品リストのフラグ格納メンバー
Private IsStockOnly As Boolean
Private IsException As Boolean

Sub CheckEol()

    '廃番リストに転記済みかチェック
    '廃番リストに転記されていなくて
    '1.商魂で区分が廃番、販売中止になっている 2.Abstractに廃番、販売終了を記入している
    '16-2-29 商魂の区分にメーカー廃番の「メ廃番」が加わりました。
    
    '廃番リストに転記済みの廃番
    If WorksheetFunction.CountIf(Eol.Range("EolCodeRange"), Code) > 0 Then
        
        IsEol = True
        
        Exit Sub
        
    End If
    
    
    If InStr(Status, "廃番") > 0 Or InStr(Status, "販売中止") > 0 Then
    
        Call addCode(Code, "EolCodeRange")
        IsEol = True
    
    End If
        
    '2.廃番リストに転記未、Abstractに廃番
    If InStr(Abstract, "終了") > 0 Or _
        InStr(Abstract, "廃番") > 0 Or _
        InStr(Abstract, "完売") > 0 Then
      
        Call addCode(Code, "EolCodeRange")
        IsEol = True
        Status = "廃番品"
        
    End If

End Sub

Sub CheckStockOnly()

    'EOLがTrueなら、転記済でQuantity=0 Allow-overdraft=0でセットされるので、CheckStokcOnlyは特に何もする必要はない
    If IsEol Then Exit Sub

    '在庫限りリストに転記済みかチェック
    If WorksheetFunction.CountIf(StockOnly.Range("StockOnlyCodeRange"), Code) > 0 Then
        
        IsStockOnly = True
        
        Exit Sub
        
    End If
    
    '在庫限りリストに転記されていない
    '1.商魂で区分が在廃、処分品になっている
    If InStr(Status, "在廃") > 0 Or InStr(Status, "処分品") > 0 Then
    
        Call addCode(Code, "StockOnlyCodeRange")
        IsStockOnly = True
    
    End If
    
    '2.Abstractに「在庫限り」「在庫僅少」を記入している
    If InStr(Abstract, "僅少") > 0 Or _
        InStr(Abstract, "在庫限り") > 0 Then
        
        Call addCode(Code, "StockOnlyCodeRange")
        IsStockOnly = True
       
    End If

End Sub

Sub CheckSecondInventry()
'コードがJANの場合は、SecondInventry=棚なしシートを調べる
'あればIsSecondInventry=棚なしフラグがTrueで、Statusは"棚なしに有"

'商魂に登録があるコードは調べる必要なし、
If Not Status = "登録なし" Then Exit Sub

LastUpdateQuantity = LastSecondInventry.getQuantity(Code)

'Debug.Assert LastUpdateQuantity = 0  '前回在庫が1以上で停止

If WorksheetFunction.CountIf(SecondInventry.Range("SecondInventryCodeRange"), Code) > 0 Then  '棚なしシートにJANの有無
    
    Status = "棚なしに有"
    IsSecondInventry = True
    SecondInventryQuantity = SecondInventry.getQuantity(Code)
    
Else 'ない場合、前回設定数を調べる
    If LastUpdateQuantity > 0 Then  '前回設定が1以上で棚なしシートにない=他で売れて在庫がなくなっている状態
        
        Status = "棚なし完売"
        SecondInventryQuantity = 0  '前回設定で0->登録なしに戻す
    
    Else
    
        Status = "登録なし"
        SecondInventryQuantity = 0
    
    End If

End If

End Sub
Sub SetAvailablePurchase()

If IsEol Then
    AvailablePurchase = False
    Exit Sub
End If

If IsStockOnly Then
    AvailablePurchase = False
    Exit Sub
End If

'仕入れ先別、入荷可否フラグが1か判定
On Error GoTo e:
    
    If WorksheetFunction.VLookup(VenderCode, purFlagSheet.Range("PurchaseFlgByVender"), 3, False) = 1 Then
        AvailablePurchase = True
        Exit Sub
    End If
    
    
On Error GoTo 0

Exit Sub
'コードがJANは基本1とする…返信FAXなどで廃番・販売終了・ロット過多はEol/StockOnlyに控えているのが前提
e:
    If Status = "棚なし完売" Or SecondInventryQuantity = 0 Then
    
        AvailablePurchase = True
        Exit Sub
    End If

End Sub

Sub SetQuantity(Qty As Long)
'商魂の在庫数量を、ヤフーへアップする形へ丸めてセット

If IsEol Then
    Quantity = 0
    Exit Sub
End If

'0以下＝0、上限100、商魂在庫数に0.6かけで小数点以下は切り捨て

'棚卸の前1ヵ月は商魂と実在庫の誤差が大きくなるので、いくらか減じてから丸める
Qty = Qty - PreDecreaseQty

Select Case Qty
    
    Case Is <= 0  '0個以下＝バックオーダー有で手配中or廃番
        
        Quantity = 0

    Case Is >= 1
                                    
        Quantity = WorksheetFunction.RoundUp(Qty * 0.6, 0)
                    
        If Quantity > 100 Then Quantity = 100
        
    Case Else
                    
        Quantity = 0
         
End Select

End Sub


