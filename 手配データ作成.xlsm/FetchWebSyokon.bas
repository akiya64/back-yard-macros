Attribute VB_Name = "FetchWebSyokon"
Option Explicit

'VBAでIEを使ってWEBのデータを取得する方法
'http://excel-ubara.com/excelvba5/EXCELVBA222.html
'http://www.vba-ie.net/ieobject/refresh.html

Type PurchaseLog
    
    Code As String
    PurchaseDate As Date
    WarehouseNum As Integer
    PurchaseQuantity As Long
    NonArrivalQty As Long
    Po As Long
    LastArrival As Date

End Type

Sub CheckNonArrival()
'WEB商魂から、注残がないか調べて未入荷が有れば備考列へ追記

'取得できなくても処理はとにかく続行
'On Error Resume Next

Dim CodeRange As Range, r As Range
Set CodeRange = Worksheets("手配数量入力シート").Range(Cells(2, 7), Cells(2, 7).End(xlDown))

For Each r In CodeRange
    
    Dim Code As String
    
    Code = r.Value
    If Len(Code) > 6 Then GoTo Continue
    
    Dim LastPur As PurchaseLog, InitPur As PurchaseLog
    LastPur = InitPur
    LastPur = FetchRecentPurchase(Code)
    
    If LastPur.NonArrivalQty > 1 Then
        
        Dim CautionCell As Range
        Set CautionCell = r.Offset(0, -5)
        
        CautionCell.Value = CautionCell.Value & IIf(CautionCell.Value = "", "", " ") & "未入荷" & LastPur.NonArrivalQty & "個 " & Format(LastPur.PurchaseDate, "M月d日") & "手配分"
    
    End If

Continue:

Next

End Sub

Private Function FetchRecentPurchase(ByVal Code As String) As PurchaseLog
'WEB商魂からHTML経由で直近の手配状況を1件取得する

    Dim CurrentCode As PurchaseLog
    CurrentCode.Code = Code

    On Error Resume Next
            
        Dim SyokonPage As InternetExplorerMedium
        Set SyokonPage = New InternetExplorerMedium
    
        SyokonPage.Navigate "http://server02/gyomu/SK_IZoom.asp?ICode=" & Code & "&C5="
        Call untilReady(SyokonPage)
        
        'オブジェクト変数はDOM
        Dim DivPurchaseLog As Object
        Set DivPurchaseLog = SyokonPage.Document.getElementsByName("t1") '最近の発注状況と入荷案内 DivタグのIDがt1
        Dim RecentRow As Object
    
        Set RecentRow = DivPurchaseLog(0).all.Item(13) '最近の発注状況と入荷案内テーブル2行目のDOM
        
        With CurrentCode
            .PurchaseDate = CDate(RecentRow.all.Item(0).innertext)
            .WarehouseNum = RecentRow.all.Item(1).innertext
            .PurchaseQuantity = RecentRow.all.Item(2).innertext
            .NonArrivalQty = IIf(RecentRow.all.Item(3).innertext = "無し", 0, RecentRow.all.Item(3).innertext)
            .Po = RecentRow.all.Item(4).innertext
            If Not RecentRow.all.Item(5).innertext Like "*-*" Then
                .LastArrival = CDate(RecentRow.all.Item(5).innertext)
            End If
        End With
        
        SyokonPage.Quit

    On Error GoTo 0
    
    FetchRecentPurchase = CurrentCode

End Function

Private Sub untilReady(objIE As Object, Optional ByVal WaitTime As Integer = 20)
'WEB商魂のレスポンス待ちのためのプロシージャ

    'サーバーレスポンス待機
    Dim starttime As Date
    starttime = Now()
    Do While objIE.Busy = True Or objIE.ReadyState <> READYSTATE_COMPLETE
        DoEvents
        If Now() > DateAdd("S", WaitTime, starttime) Then
            Exit Do
        End If
    Loop
    
    'ローディング画面の表示後に、詳細データが動的に再描画されるので2秒待機
    Dim WaitEnd As Date
    WaitEnd = DateAdd("S", 2, Now())
    Do While Now() < WaitEnd
        DoEvents
    Loop
    
    DoEvents

End Sub
