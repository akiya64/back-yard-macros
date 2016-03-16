VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Invoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'送り状の集合クラス、InvoiceNumberListの中身はOrderオブジェクトを前提として処理している。
'トリッキーになるがVBAでもタイプセーフな配列が作れるらしい。

'fetchReportXlsx() 定数宣言しているパスの出荷一覧詳細.xlsxを読み込んで送り状Dictionalyを作成
'fetchReportCsv()　定数宣言しているパスの出荷一覧詳細.csvを読み込んで送り状Dictionalyを作成

'出荷一覧詳細.csvのパス
Const REPORT_CSV_PATH As String = "\\server02\商品部\ネット販売関連\出荷通知\出荷通知_楽天\出荷一覧詳細.csv"

'出荷一覧詳細.xlsxのパスとファイル名
Const REPORT_XLSX_FOLDER As String = "\\server02\商品部\ネット販売関連\出荷通知\出荷通知_楽天\旧出荷一覧詳細\"
Const REPORT_XLSX_BASENAME As String = "出荷一覧詳細"

'注文番号をキーとした、Invoiceオブジェクトの辞書配列
Public InvoiceNumberList As Dictionary

Public Sub fetchReportXlsx()
'出荷詳細一覧エクセルファイルの方からInvoiceListを生成

Application.ScreenUpdating = False

'ファイル名は出荷一覧詳細_150603.xlsxの形
Dim FileName As String
FileName = REPORT_XLSX_BASENAME & "_" & Format(Date, "yymmdd") & ".xlsx"

'テスト用のパス
'FileName = REPORT_XLSX_BASENAME & "_" & "160311" & ".xlsx"

Dim Path As String
Path = REPORT_XLSX_FOLDER & FileName

Call findReportFile(Path)

'ブックを開く処理、送り状レポートワークブックを取得
Dim InvoiceReportWb As Workbook
Dim wb As Workbook

'開いているブックの中に出荷一覧詳細がないかチェック
'開いてればそのまま使えばいい、開いてなければ開いてセット
For Each wb In Workbooks
    If wb.Name = FileName Then
        Set InvoiceReportWb = wb
    End If
Next wb

'ワークブックを開いてセット
If InvoiceReportWb Is Nothing Then

    Set InvoiceReportWb = Workbooks.Open(Path)
    
    Dim IsOpenReport As Boolean
    IsOpenReport = True
        
End If


ShippingFileProgress.ProgressBar.Value = 2
ShippingFileProgress.ShowCurrentProcess.Caption = "出荷一覧詳細 読み込み中"


'出荷一覧詳細で注文番号はAE1列
'ちょいちょい見出しがない列がある。
'昇順ソートでヤフーの注文番号フォーマット100######が上に来るはずだが、
'自信がないので、実行時間が遅くなっても全行調べます。－＞CSV読込よりずっと速かった。

'Invoceオブジェクトを入れるDictionary
Set InvoiceNumberList = New Dictionary

Dim LastRow As Long
LastRow = InvoiceReportWb.Sheets(1).Range("AE1").SpecialCells(xlCellTypeLastCell).Row

Dim i As Long

'全ての行を調べる
For i = 2 To LastRow
    
    '発送日C列、送り状番号E列、注文番号AE列
    
    With InvoiceReportWb.Sheets(1)
        
        Dim CurrentId As String
        CurrentId = .Range("AE" & i).Value
        
        Dim CurrentInvoiceNum As String
        CurrentInvoiceNum = .Range("E" & i).Value
        
        Dim CurrentShipDate As Date
        CurrentShipDate = .Range("C" & i).Value
        
    End With
    
    'ヤフーの注文番号の形式なら、Invoiceインスタンスを生成してInvoiceNumberList に追加
    If CurrentId Like "10######" And Len(CurrentInvoiceNum) > 0 Then
            
        Dim invoice As invoice
        Set invoice = New invoice
        
        invoice.Id = CurrentId
        invoice.InvoiceNumber = CurrentInvoiceNum
        invoice.ShippingDate = CurrentShipDate
        
        On Error Resume Next
        
            InvoiceNumberList.Add CurrentId, invoice
        
        On Error GoTo 0
                  
    End If

Next

checkInvoiceList

'このモジュールでブックを開いていれば閉じる。
If IsOpenReport Then InvoiceReportWb.Close SaveChanges:=False

Application.ScreenUpdating = True

End Sub

Private Function findReportFile(Path As String) As String
'出荷詳細のパスを返します。

Dim FSO As FileSystemObject
Set FSO = New FileSystemObject

Dim res As Long 'MsgBoxからの返り値を受ける

If FSO.FileExists(Path) Then '出荷詳細一覧ファイルの存在チェック
    
    Dim dtc As Date
    
    dtc = FSO.GetFile(Path).DateLastModified

    res = MsgBox(FSO.GetFileName(Path) & vbLf _
                & "最終更新：" & Format(dtc, "mm月dd日 hh:mm") & vbLf _
                & "このファイルから生成しますか？" _
                , vbYesNo + vbQuestion)
                
    If res <> 6 Then

        MsgBox prompt:="出荷一覧ファイルの読込がキャンセルされました。", Buttons:=vbInformation
        End

    End If

    findReportFile = Path

Else

    MsgBox prompt:="本日の出荷一覧ファイルがありません。", Buttons:=vbCritical
    End

End If

End Function

Private Sub checkInvoiceList()
'送り状リストの件数をチェックして0件ならマクロ終了

If InvoiceNumberList.Count = 0 Then

    MsgBox prompt:="本日の出荷はありません。", Buttons:=vbCritical
    End

End If

End Sub

Public Sub fetchReportCsv()
 '出荷詳細CSVを読み込んでInvoiceListを作ります。
 
Dim Path As String
Path = findReportFile(REPORT_CSV_PATH)

Set InvoiceNumberList = New Dictionary

Dim FSO As FileSystemObject
Set FSO = New FileSystemObject

Dim TS As Textstream
Set TS = FSO.OpenTextFile(Path, ForReading)

Dim LineBuf() As String

Do Until TS.AtEndOfStream

    LineBuf = Split(TS.ReadLine, ",")
    
    Dim tmp() As Variant
    '30列目の注文番号=tmp(0)、4列目の送り状番号tmp(1)　2列目の発送日tmp(2)
    tmp() = Array(LineBuf(30), LineBuf(4), LineBuf(2))

    Dim j As Long
    For j = 0 To UBound(tmp())
        tmp(j) = Trim(Replace(tmp(j), Chr(34), "")) 'chr(34)で " [半角二重引用符]らしい
    
    Next
    
    '注文番号が10始まりの数字8ケタ＝ヤフーの注文番号フォーマット
    'かつ、送り状番号が入っている注文番号のみInvoiceListに格納する

    If tmp(0) Like "10######" And Len(tmp(1)) > 0 Then
        
        Dim invoice As invoice
        Set invoice = New invoice
        
        invoice.Id = tmp(0)
        invoice.InvoiceNumber = tmp(1)
        invoice.ShippingDate = CDate(tmp(2))
        
        On Error Resume Next
        
            InvoiceNumberList.Add tmp(0), invoice
        
        On Error GoTo 0
        
    End If
        
Loop

checkInvoiceList

End Sub
