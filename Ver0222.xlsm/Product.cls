VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'受注1アイテムに対応するクラス

'シートから読み取ってセットするフィールド

Public OrderId As String

Public Code As String

Public Name As String
Public Jan As String

Public CenterStockState As String

Public ReqQuantity As Long '必要個数

Public PurchaseDate As Date

Public VenderStatus As String

Public EstimatedArrivalDate As Date

'Product自身でメソッド走らせてセットするフィールド

Public IsUnavailable As Boolean
Public IsPickingDone As Boolean

Sub setIsUnavailable()
    
    If InStr(CenterStockState, "廃番") > 0 Or _
        InStr(CenterStockState, "未定") > 0 Or _
        InStr(CenterStockState, "終了") > 0 Or _
        InStr(CenterStockState, "中止") > 0 Then
        
        IsUnavailable = True
    
    ElseIf InStr(VenderStatus, "廃番") > 0 Or _
        InStr(VenderStatus, "未定") > 0 Or _
        InStr(VenderStatus, "終了") > 0 Or _
        InStr(VenderStatus, "中止") > 0 Or _
        InStr(VenderStatus, "欠品") > 0 Then
        
        IsUnavailable = True
    
    End If
    
End Sub
Sub setIsPickingDone()
    
    If CenterStockState = "" Then
        
        IsPickingDone = True
    
    End If

End Sub
