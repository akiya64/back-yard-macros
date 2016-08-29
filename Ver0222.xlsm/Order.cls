VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'注文1件のクラス 日付型の初期値は #00:00:00#

Public Id As String                    '注文番号
Public OrderDate As Date               '受注日
Public InvoiceNumber As String         '送り状番号
Public ShippingDate As Date            '発送日
Public ShippingMethod As String        '発送方法
Public BuyerName As String             '注文者名

Public IsWaitingPayment As Boolean     '入金待ち

Public Products As Dictionary          '注文商品リスト

Public CanNotShipping As Boolean       '発送ができる＝False　できない＝True

Public SendMailDate As Date            '納期連絡日

Public IsCancel As Boolean             'キャンセルフラグ　True=キャンセル
Public IsExceptShipping As Boolean     '

Public AlertPiriod As Date             'アラートの基準日


Public Sub Class_Initialize()  'コンストラクタ、オブジェクト生成時にデフォルトのAlertPiriod アラート起点日をセットした方がいい



End Sub

Sub setCanNotShipping()
'発送不可能な注文であるか？
'アイテムの状況をチェックして発送=Shipping不可フラグを立てます。
    
    Dim v As Variant
    

    For Each v In Products
    
        With Products(v)
        

        'ピッキングできていない商品があればTrue
        If Not .IsPickingDone Then
            CanNotShipping = True
            
        End If
        
        '手配日が本日ならTrue
        If .PurchaseDate = Date Then
             
            'ただし、ArrivalDateが本日ならFalse
            If DateDiff("d", Products(v).EstimatedArrivalDate, Date) = 0 Then
                
                CanNotShipping = False
            
            Else
                
                CanNotShipping = True
            
            End If
        
        End If

        End With
        
    Next
    
    Call setAlertPiriod        'alert起算日をセット
    Call cancelCanNotShipping  'CanNotShippingを解除できるかチェック

End Sub

Private Sub setAlertPiriod()  'アラートの起点日

Dim p As New Product

'デフォルトで一旦、受注日をセットする
'何もなければ、基準日は受注日となり、ここから三日経過でアラート対象となる

AlertPiriod = OrderDate

Dim v As Variant

For Each v In Products
    Set p = Products(v)
    
    '毎回DateDiffします、常に一番遅い日付を指すようにセットしたいので。

    'ピッキングされていない商品なら、手配日を調べてセット
    If Not p.IsPickingDone Then  '
        If DateDiff("d", AlertPiriod, p.PurchaseDate) >= 0 Then
            AlertPiriod = p.PurchaseDate
        End If
    End If
    
    '手配した商品があれば、入荷予定日をセット
    If p.PurchaseDate <> #12:00:00 AM# Then
        If DateDiff("d", AlertPiriod, p.PurchaseDate) > 0 Then
            AlertPiriod = p.PurchaseDate
        End If
    End If

    '入荷予定日が定まっていれば、入荷予定日をセット
    If p.EstimatedArrivalDate <> #12:00:00 AM# Then
        If DateDiff("d", AlertPiriod, p.EstimatedArrivalDate) > 0 Then
            AlertPiriod = p.EstimatedArrivalDate
        End If
    End If

Next

End Sub

Private Sub cancelCanNotShipping()
'Order.IsShippingDisableが解除できる条件
'Product.IsPickingがFalseの商品全てについてEstimatedArrivalDateを超えている

'一旦IsAllProductsAvailable=全商品が揃っているかフラグはTrue
'一点でも商品状況がFalse=発送不可ならFalseにする。VBAのBoolean初期値はFalse

Dim IsAllProductsAvailable As Boolean
IsAllProductsAvailable = True

Dim p As Product
Dim v As Variant

For Each v In Products
    
    
    Set p = Products(v)
    
    'センター、メーカーステータスで廃番とかならTrue
    '解除できる条件は発生しないので、チェック終了
    If p.IsUnavailable Then
        CanNotShipping = True
        Exit Sub
    End If

    'Productがピッキング済かどうか
    If p.IsPickingDone Then
        
        IsAllProductsAvailable = True
        GoTo continue
        
    End If
    
    'Productがピッキングされていなくて、入荷予定日に日付が入っている時、どちらかがFalseならAvailableはFalse
    If p.IsPickingDone = False And p.EstimatedArrivalDate <> #12:00:00 AM# Then
                            
        'IsAllItemAvailableとの論理和を取る、この時IsAllProductsAvailableは、イテレータの指す商品の一個前の商品まで全てでAvalableでなければTrueを示さない
        'TODO:もっと綺麗に、一商品でもAvailableでなければ、CanNotShipping=Falseができるのではないかと、再帰とかセルのOR関数とか使えれば。
        
        If DateDiff("d", Date, p.EstimatedArrivalDate) <= 0 And IsAllProductsAvailable Then
            IsAllProductsAvailable = True
            
            GoTo continue
        
        Else
            IsAllProductsAvailable = False
            GoTo continue
        
        End If
    
    Else
    
        IsAllProductsAvailable = False
    
    End If

continue:

Next

'全ての商品がAvailableならばCanNotShippingはFalse、出荷可能

If IsAllProductsAvailable Then CanNotShipping = False


End Sub
