Attribute VB_Name = "Main"
Option Explicit
Sub 受注ファイル読込()

OrderSheet.Activate

If Not Range("A2").Value = "" Then
    MsgBox "データ取得済です。"
    End
End If

'プログレスバーの準備
With ShowProgress
    .ProgressBar.Min = 1
    .ProgressBar.Max = 4

    .Show vbModeless
End With

Dim CsvPath As String
CsvPath = Application.GetOpenFilename(Title:="CSVを指定", FileFilter:="クロスモールCSV,*.csv", FilterIndex:="2")

If CsvPath = "" Then
    MsgBox "ファイル指定がキャンセルされました。" & vbLf & "マクロを終了します。"
    End
End If

ShowProgress.ProgressBar.Value = 3
ShowProgress.StepMessageLabel = "CSV読込中"

Call ReadClossMallCsv(CsvPath)

ShowProgress.ProgressBar.Value = 4
ShowProgress.StepMessageLabel = "CSV読込完了"

'マクロ起動ボタンを消去
'OrderSheet.Shapes(1).Delete

'アドイン用の行・列 表示
Dim LastRow As Long
LastRow = Range("D1").SpecialCells(xlCellTypeLastCell).Row

Range("L1").Value = "アドイン指定 台帳：9998"
Range("L2:O2") = Array(2, 9, LastRow, 12)

ShowProgress.Hide

MsgBox "アドインを実行して下さい。"

'アドインでロケーション取得前の処理終了

End Sub

'この位置に、アドインでのロケーション取得が必要。
'DB接続してデータとってこれればMain処理は1クリックになる。

Sub 電算提出_振分けシート作成()

'アドイン後の処理
OrderSheet.Activate

'アドイン未実行の際は、ダイアログで警告を出して終了
If InStr(Range("L1").Value, "アドイン指定") > 0 Then
    MsgBox "アドインを実行して下さい。"
    End
End If

'プログレスバーの準備
With ShowProgress
    .ProgressBar.Min = 1
    .ProgressBar.Max = 9
    
    Dim ProgressStep As Long
    ProgressStep = 1
    
    .ProgressBar.Value = ProgressStep
    .Show vbModeless
End With


'無効なロケーションをカット
DataVaridate.ModifyOrderSheet

'受注一覧シートの修正終わり、シートを保護、データロックをかける。
OrderSheet.Protect

'モール毎の電算室提出データ保存、振分けシート作成
Dim Mall As Variant, Malls As Variant

Malls = Array("ヤフー", "楽天", "Amazon")

For Each Mall In Malls

    ProgressStep = ProgressStep + 1
    ShowProgress.ProgressBar.Value = ProgressStep
    ShowProgress.StepMessageLabel = Mall & "データ処理中"
    
    'ピッキングシート作成・保存
    'Call BuildSheets.OutputPickingData(CStr(Mall))
    
    '振分け用シート作成
    Call BuildSheets.CreateSorterSheet(CStr(Mall))

Next

'アラートダイアログを抑止
Application.DisplayAlerts = False

'テンプレートシートを削除
Worksheets("ピッキングシート提出用テンプレート").Delete
Worksheets("振分用テンプレート").Delete

'このファイルを保存
Dim PutFileName As String
PutFileName = "ピッキング・振分" & Format(Date, "MMdd") & ".xlsx"

ShowProgress.ProgressBar.Value = ProgressStep + 1
ShowProgress.StepMessageLabel = Mall & "保存処理中"

'擬似的なTry-Catchで保存を実行
On Error Resume Next
    
    'Try
     ThisWorkbook.SaveAs FileName:="C:" & Environ("HOMEPATH") & "\Desktop\" & PutFileName
    
    'catch
    If Err Then
        MsgBox "ファイルを保存できませんでした。手動で名前を付けて保存してください。"
    End If

'On Error Goto 0 宣言でErrは解除される
On Error GoTo 0


ShowProgress.ProgressBar.Value = ProgressStep + 2
ShowProgress.StepMessageLabel = Mall & "振分シート プリント"

'実行PCデフォルトのプリンタでプリントアウト
'プリンタの指定してなければ、Windowsのデフォルトプリンタを使う。
Dim i As Long
For i = 2 To Worksheets.Count

    'Worksheets(i).PrintOut

Next

OrderSheet.Activate

'プログレスバーを消して終了メッセージ
ShowProgress.Hide
MsgBox Prompt:="処理完了", Buttons:=vbInformation

'この後、ThisWorkBookのコードへ処理を戻さない
End

End Sub
