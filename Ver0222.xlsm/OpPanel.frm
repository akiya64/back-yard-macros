VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OpPanel 
   Caption         =   "ヤフー注文処理"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6405
   OleObjectBlob   =   "OpPanel.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "OpPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'全プロシージャで共通するファイルの所在とファイル名

'明細と注文ヘッダーの在処
Const MEISAI_PATH As String = "\\MOS10\Users\mos10\Desktop\ヤフー\meisai.csv"
Const TYUMON_PATH As String = "\\MOS10\Users\mos10\Desktop\ヤフー\tyumon_H.csv"

Private Sub CommandButton1_Click()
    
    梱包室受注ファイル読込
    
End Sub

Private Sub CommandButton12_Click()
    
    CheckPickingProducts
    
End Sub

Private Sub CommandButton13_Click()
    
    FetchFaxReply.FetchFaxReply

End Sub

Private Sub CommandButton14_Click()
    
    送り状番号一括ファイル作成
    
End Sub

Private Sub CommandButton15_Click()
    
    遅延チェック
    
End Sub

Private Sub CommandButton16_Click()
    'アウトラインの展開、折りたたみについて  http://www.relief.jp/itnote/archives/017927.php
    
    OrderSheet.Outline.ShowLevels ColumnLevels:=2
    
End Sub

Private Sub CommandButton17_Click()

    OrderSheet.setProtect

End Sub

Private Sub CommandButton18_Click()
    
    OrderSheet.setUnprotect
    MsgBox "30秒間のみ、注残一覧の記入/削除がフリーになります。"
    End

End Sub

Private Sub CommandButton19_Click()

    Call checkMesaiFileExistance

End Sub

Private Sub CommandButton2_Click()
    
    Meisai個別読込

End Sub
Private Sub CommandButton3_Click()

    tyumon_H個別読込

End Sub
Private Sub CommandButton4_Click()

    fetchShippingDone

End Sub
Private Sub CommandButton5_Click()

    Unload Me

End Sub

Private Sub CommandButton8_Click()
    
    archiveCompletedOrder

End Sub

Private Sub UserForm_Initialize()
'ユーザーフォームが開いた時の処理

'ページ1を表示
Me.MultiPage1.Value = 0

Dim PickingFilePath As String

TextBox1.Value = "存在チェック 未"
TextBox2.Value = "存在チェック 未"


If Format(Now(), "hh") < 10 Then '10時までのフォームオープンなら、Meisaiファイル存在チェック

    Call checkMesaiFileExistance

Else

    TextBox1.Value = "要 手動チェック"
    TextBox2.Value = "要 手動チェック"

End If

PickingFilePath = Range("PickingSheetFolder").Value

If Right(PickingFilePath, 1) <> "\" Then PickingFilePath = PickingFilePath & "\" '末尾\マークでないとき補完

PickingFileName = Range("PickingSheetBaseName").Value


'テキストBox4、ピッキングファイルのファイル情報
TextBox4.Value = Dir(PickingFilePath & PickingFileName & "*")

'テキストBox5、出荷一覧詳細のファイル情報

If Format(Now(), "hh") > 15 Then '15時過ぎてのフォームオープンなら、出荷一覧詳細のファイル存在チェック
    
'出荷一覧詳細のありか ファイルチェックに使ってます
    Const REPORT_FILE_PATH As String = "\\server02\商品部\ネット販売関連\出荷通知\出荷通知_楽天\旧出荷一覧詳細\"
    Const REPORT_FILE_BASE As String = "出荷一覧詳細_"
    
    Dim ReportFileName As String
    ReportFileName = REPORT_FILE_BASE & Format(Date, "yymmdd") & ".xlsx"
    TextBox5.Value = Dir(REPORT_FILE_PATH & ReportFileName)
    
End If

End Sub

Private Sub checkMesaiFileExistance()

'ファイル操作オブジェクト生成
Dim FSO As Object
Set FSO = New FileSystemObject

If FSO.FileExists(MEISAI_PATH) Then
    
    TextBox1.Value = FSO.GetFile(MEISAI_PATH).DateCreated
    TextBox2.Locked = True
    
Else
    
    TextBox1.Value = "ファイルなし"

End If

If FSO.FileExists(TYUMON_PATH) Then
    
    TextBox2.Value = FSO.GetFile(TYUMON_PATH).DateCreated
    TextBox2.Locked = True
    
Else
    
    TextBox2.Value = "ファイルなし"

End If

End Sub

Private Sub fetchShippingDone()

Dim LineBuf As Variant
Dim order As Variant

'連絡状況シート＝OrderSheetの注文番号のレンジ
Dim search_range As Range
Set search_range = OrderSheet.Cells(2, 2).Resize(OrderSheet.Cells(2, 2).SpecialCells(xlCellTypeLastCell).Row, 1)

'ループ内で使うFind関係のレンジ
Dim firstCell As Range
Dim FoundCell As Range

' ファイルダイアログからパスを指定して、FSOで開く
Dim file_path As String
file_path = fetchOrderCsv.setCsvPath("shipping.csv")

If file_path = "" Then
    MsgBox "ファイル指定がキャンセルされました。"
    Exit Sub
End If

Dim FSO As Object
Set FSO = New FileSystemObject

' CSVをテキストストリームとして処理する
Dim TS As Textstream
Set TS = FSO.OpenTextFile(file_path, ForReading)
       
Do Until TS.AtEndOfStream
    
'注文番号、送り先氏名、（処理）状況、問い合わせ番号を配列tmpに入れる

    LineBuf = Split(TS.ReadLine, ",")
    
    'tmp[0]=OrderID=Column"A"
    'tmp[1]=Ship name=送り先名=Column"B",
    'tmp[2]=status=状況=Column"C"
    'tmp[3]=shipDate=発送日=Column"D"
    'tmp[4]=Shipping Number=問い合わせ番号=Column"E"
    
    tmp = Array(LineBuf(0), LineBuf(1), LineBuf(2), LineBuf(3), LineBuf(4))
    
    For j = 0 To UBound(tmp)
        tmp(j) = Trim(Replace(tmp(j), Chr(34), "")) 'chr(34)で " [半角二重引用符]
    
    Next

'注残一覧シートの該当する注文番号に読み取った情報を入れる
                
    Set FoundCell = search_range.Find(what:=tmp(0))
    
    If Not FoundCell Is Nothing Then
        
        FirstCellAddress = FoundCell.Address
        
        If Not OrderSheet.Cells(FoundCell.Row, 3).Value = tmp(1) Then '注文者名がなければ最初のFoundCellに入れる
        
            OrderSheet.Cells(FoundCell.Row, 3) = tmp(1)
        
        End If
                
        Do
            '発送状況「済」はFindして見つかった注文番号すべてに入れる
            'IF 状況が完了 AND 発送セルが空白 AND 問い合わせ番号がある Then 処理状況列に済、発送日列に発送日
            'TODO:FINDでループ回さないようにリファクタリング
            
            If Cells(FoundCell.Row, "O").Value = "" And Not tmp(4) = "" Then
                OrderSheet.Cells(FoundCell.Row, 15) = "済" '処理状況=発送はO列
                OrderSheet.Cells(FoundCell.Row, 16) = tmp(3)
    
            End If
            
            Set FoundCell = Cells.FindNext(FoundCell)
            
            If FoundCell Is Nothing Or FoundCell.Address = FirstCellAddress Then Exit Do
        
         Loop

    End If

Loop


' 指定ファイルをCLOSE
TS.Close
Set TS = Nothing
Set FSO = Nothing

OpPanel.Hide

Call CheckBelate.遅延チェック

ThisWorkbook.Save

End Sub

Private Sub archiveCompletedOrder()

'前々月最終日までの発送済み、キャンセルを別シートに移動します

Application.ScreenUpdating = False

'オートフィルターがかかっていれば一旦解除します、コピー後の行が非表示になるため
'sheetオブジェクトに、オートフィルターの有無を示すFilterModeプロパティ:Boolean型があります
If OrderSheet.FilterMode Then OrderSheet.Range("A1").CurrentRegion.AutoFilter

'本日日付から今日の「日」を引く＝前月最終日を算出
Dim LastDay As Date
LastDay = DateAdd("d", -(Format(Date, "d")), Date)

Dim i As Long
i = 2

With ThisWorkbook.Sheets("注残一覧")

'Cell(1,1)の日付を比較して前月最終日を超えない限り処理をする

Do Until DateDiff("d", LastDay, .Cells(i, 1)) > 1
    
    '空欄の場合も自動キャストで比較してしまい、処理が止まらない場合がありえる
    If IsEmpty(.Cells(i, 1)) Then Exit Sub
    
    If .Cells(i, "O").Value = "済" Then

        .Rows(i).Cut Destination:=Sheets("完了").Rows(Sheets("完了").UsedRange.Rows.Count + 1)
        .Rows(i).Delete
        
    ElseIf OrderSheet.Cells(i, "O").Value = "キャンセル" Then
    
        .Rows(i).Cut Destination:=Sheets("キャンセル").Rows(Sheets("キャンセル").UsedRange.Rows.Count + 1)
        .Rows(i).Delete
    
    Else
        
        '列Oが済orキャンセルでない場合＝空欄の時、行ポインタを進める
        i = i + 1
    
    End If

Loop

End With

Application.ScreenUpdating = True

OrderSheet.Activate

'ボタンを再配置
OrderSheet.Shapes("ShowFormButton").Top = OrderSheet.Range("A1").End(xlDown).Offset(2, 1).Top
OrderSheet.Shapes("ButtonHideWish").Top = OrderSheet.Range("A1").End(xlDown).Offset(2, 1).Top

End Sub

