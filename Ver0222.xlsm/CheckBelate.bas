Attribute VB_Name = "CheckBelate"
Public Sub 遅延チェック()
'belated arrivalで延着のこと､遅延はBelateで統一します。
'OrderListを作成して、Belate=遅延チェックをして、該当注文をBelateListに追加します。
'最後にリストをMsgBoxでポップアップ
'アラートの表示方法はまた何かいい方法があれば変えます。

'注残一覧はOrderSheet

OrderSheet.Activate

Application.ScreenUpdating = False

'未発送商品のOrderListを作ります
Dim UndispatchList As Dictionary
Set UndispatchList = OrderSheet.getUndispatchOrders

Dim o As order
Dim v As Variant

Dim BelateList As Dictionary
Set BelateList = New Dictionary

For Each v In UndispatchList 'OrderListの個々のOrderについて、チェック
    Set o = UndispatchList(v)
      
    If checkBelateDispatch(o) Then 'checkBelateDispatch Functionでチェック
              
        'AlertPiriodよりSendMailDateが後ろ=Purchase後に一度連絡している注文は、遅延リストに加えない
        'AlertPiriodがEstimatedArrivalDate=入荷日を指していれば、通常ルーティンでは入荷日は連絡日より後ろになるため、DateDiffで正の値になる。
        'Dayでしか判定をしないので、1ヶ月を超えるとアラート上がらないが、入荷予定から三日経過とかのスパンでの発送漏れ、連絡漏れを把握したいので構わない。
                  
        If DateDiff("d", o.AlertPiriod, o.SendMailDate) < 0 Then
        
            BelateList.Add o.Id, o
        
        End If
    
    End If
    
Next

Dim IdList As String

For Each v In BelateList  '遅延リストをMsgBoxで表示するためStringで出力して連結。
    
    IdList = IdList & vbLf & Format(BelateList(v).OrderDate, "MM/dd") & "  No." & BelateList(v).Id

Next

Application.ScreenUpdating = True

If BelateList.Count > 0 Then

    MsgBox prompt:="未発送/未連絡で3日以上経過している注文" & vbLf _
            & IdList & vbLf _
            & vbLf _
            & BelateList.Count & "件あります。", _
            Buttons:=vbExclamation, _
            Title:="チェック結果"

Else

    MsgBox "未発送/未連絡で3日経過している注文はありません。", _
            Buttons:=vbInformation, _
            Title:="チェック結果"
   

End If

End Sub

Private Function checkBelateDispatch(order As order) As Boolean

'単体の未発送チェッカー

    '振込の場合は入金連絡済で、7日を超えていれば遅延。ヤフーの自動処理で注文日より14日後にポイント自動確定するので
    If order.IsWaitingPayment And DateDiff("d", order.AlertPiriod, Date) > 7 Then
        TmpBl = True
        GoTo re
    Else
        TmpBl = False
    End If

    'Orderオブジェクトのアラート起算日と本日の差が2を超える＝三日以上でTrue

    If DateDiff("d", order.AlertPiriod, Date) > 2 Then
        TmpBl = True
    Else
        TmpBl = False
    End If

re:
checkBelateDispatch = TmpBl

End Function
