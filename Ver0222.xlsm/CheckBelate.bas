Attribute VB_Name = "CheckBelate"
Public Sub 遅延チェック()

Dim BelateList As Dictionary
Set BelateList = MakeBelateList()

For Each v In BelateList  '遅延リストをMsgBoxで表示するためStringで出力して連結。
    
Dim IdList As String
    
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

Public Sub 遅延リスト出力()

Dim BelateList As Dictionary
Set BelateList = MakeBelateList()

Workbooks.Add

'新規追加ファイルにヘッダー作成
With ActiveSheet
    
    .Range("A1").Value = "出荷状況確認 " & Format(Date, "m月d日")
    .Range("A2:I2") = Array("受注日", "注文番号", "注文者名", "Line", "商品コード", "商品名", "数量", "出荷状況", "出荷日", "送り状番号")

End With

Dim i As Integer
i = 3


For Each v In BelateList
  
    Id = BelateList(v).Id
    
    '注文番号から、注残一覧シートの行番号を特定、注文情報を配列に格納
    With ThisWorkbook.Worksheets("注残一覧")
        Dim r As Long, rng As Range
        r = .Range("B:B").Find(Id).Row
        Dim arr
        arr = .Range("A" & r & ":" & "G" & r)
    
    End With
    
    '作成した新規ブックに貼り付けて行く
    ActiveSheet.Range(Cells(i, 1), Cells(i, 7)) = arr
    
    i = i + 1

Next

Debug.Print i
'Line番号列を削除
ActiveSheet.Columns("d:d").Delete

ActiveSheet.Range(Cells(3, 1), Cells(1, 1).End(xlDown)).NumberFormatLocal = "m""月""d""日"";@"

End Sub

Private Function MakeBelateList() As Dictionary
'belated arrivalで延着のこと､遅延はBelateで統一します。
'OrderListを作成して、Belate=遅延チェックをして、該当注文をBelateListに追加します。

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

Set MakeBelateList = BelateList

End Function

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
