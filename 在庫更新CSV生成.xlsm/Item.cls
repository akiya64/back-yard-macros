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
Public Status As String

Public IsEol As Boolean
Public IsStockOnly As Boolean

Private Quantity As Integer 'ヤフーショッピングに設定する在庫数はMAX100

Private AvailablePurchase As Boolean

Private SyokonData As Syokon  '商魂の情報を格納するメンバー

Private IsException As Boolean
Private HasSlimsLocation As Boolean

Sub Constractor(ArgsCode As String)

'命名は初期化でいいのだろうか。別にコンストラクタは一応ある。

    Code = ArgsCode
    SyokonData = SyokonMaster.GetSyokonQtyKubun(Code)

    Status = SyokonData.Status

    '廃番、在庫限りの各シートをチェック
    Me.CheckEol
    Me.CheckStockOnly
    
    '設定在庫数のセット
    
    SetQuantity (SyokonData.Quantity)
        
    '手配可否をセット
    Me.SetAvailablePurchase

End Sub

Function GetQuantity() As Long

    GetQuantity = Quantity

End Function

Function GetAvailablePurchase() As Boolean
    
    GetAvailablePurchase = AvailablePurchase

End Function
Function GetSyokonData() As Syokon

    GetSyokonData = SyokonData
    
End Function

Sub CheckEol()

    '廃番リストに転記済かチェック、前回までの商魂廃番は廃番リストに控えられている。
    '商魂で区分が廃番、販売中止になっていれば、廃番リストへ転記
    '16-2-29 商魂の区分がメーカー廃番の「メ廃番」に変わりました。

    '廃番リストに転記済みの廃番
    If WorksheetFunction.CountIf(Eol.Range("EolCodeRange"), Code) > 0 Then
        
        IsEol = True
        
        Exit Sub
        
    End If
    
    If InStr(Status, "メ廃番") > 0 Or InStr(Status, "販売中止") > 0 Then
    
        IsEol = True
    
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
    If InStr(Status, "在庫処分") > 0 Or InStr(Status, "在庫廃番") > 0 Then
    
        IsStockOnly = True
    
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
    
    If WorksheetFunction.VLookup(SyokonData.VenderCode, purFlagSheet.Range("PurchaseFlgByVender"), 3, False) = 1 Then
        AvailablePurchase = True
        Exit Sub
    End If
    
    
On Error GoTo 0

Exit Sub

'コードがJANは基本1とする…返信FAXなどで廃番・販売終了・ロット過多はEol/StockOnlyに控えているのが前提
e:
        
        AvailablePurchase = False
        Exit Sub

End Sub

Sub SetQuantity(Qty As Long, Optional SlimsFlag As Boolean)
'商魂の在庫数量を、ヤフーへアップする形へ丸めてセット

If IsEol Then
    Quantity = 0
    Exit Sub
End If

'Slimsの在庫データを使用する際はロケーション有無を確認
If SlimsFlag = True Then
        
        If Slims.HasLocation(Code) Then
        
            Qty = Slims.GetQuantity(Code)
    
        Else
        
            Quantity = 0
            Exit Sub
    
        End If

End If

'0以下＝0、上限100、商魂在庫数に0.6かけで小数点以下は切り捨て
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

