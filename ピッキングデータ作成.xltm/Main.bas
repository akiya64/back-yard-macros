Attribute VB_Name = "Main"
Option Explicit

'あす楽・Amazonプライム分の作成かを記録するフラグ BuildSheets.PreparePickingBookで使用
Public IsSecondPicking As Boolean
Public IsTimeStampMode As Boolean

Sub ピッキング_振分作成()
'ボタンから起動するプロシージャ

OrderSheet.Activate

If Not Range("A2").Value = "" Then
    MsgBox "データ取得済です。"
    End
End If

'プログレスバーの準備
With ShowProgress
    .ProgressBar.Min = 1
    .ProgressBar.Max = 15

    .Show vbModeless
End With

ShowProgress.ProgressBar.Value = 2
ShowProgress.StepMessageLabel = "CSV読込中"

Call LoadCsv

'アドイン用にコード修正、セット分解
Call DataValidate.FixForAddin
Call SetParser.SetParse7777

ShowProgress.ProgressBar.Value = 3
ShowProgress.StepMessageLabel = "ロケーションデータ取得中"
Application.Wait Now + TimeValue("00:00:02")
'2秒待機してプログレスバーを表示更新を待つ。梱包室で行うとプログレスバーが更新されない対策 効果未検証

'DB接続、ロケーション取得、受注データの修正
Call ConnectDB.Make_List
Call DataValidate.FilterLocation

'モール毎の電算室提出データ保存、振分けシート作成

'テンプレートの1行目に本日日付を入れる
Worksheets("ピッキングシート提出用テンプレート").Range("C1").Value = Format(Date, "M月dd日")

'クロスモール側で使用しているモール名の配列を作成し、イテレートする。
Dim Mall As Variant, Malls As Variant, ProgressStep As Long
Malls = Array("Amazon", "楽天", "Yahoo")
ProgressStep = 3

For Each Mall In Malls
        
    ProgressStep = ProgressStep + 1
    ShowProgress.ProgressBar.Value = ProgressStep
    ShowProgress.StepMessageLabel = Mall & "ピッキング作成中"
    
    'モール毎の受注件数がゼロ件ならファイル生成しない。
    If WorksheetFunction.CountIf(OrderSheet.Range("F:F"), Mall & "*") = 0 Then GoTo Continue1
    
    'ピッキングシート作成・保存
    Call BuildSheets.OutputPickingData(CStr(Mall))


Continue1:

Next

'振分用セット商品リストをコード昇順にするために、ソート
Call OrderSheet.SortAscend("受注時商品コード")

For Each Mall In Malls
        
    ProgressStep = ProgressStep + 1
    ShowProgress.ProgressBar.Value = ProgressStep
    ShowProgress.StepMessageLabel = Mall & "振分シート作成中"
    
    'モール毎の受注件数がゼロ件ならファイル生成しない。
    If WorksheetFunction.CountIf(OrderSheet.Range("F:F"), Mall & "*") = 0 Then GoTo Continue2
    
    '振分け用シート作成
    Call BuildSheets.CreateSorterSheet(CStr(Mall))

Continue2:

Next

Call OrderSheet.SortAscend("管理番号")

'受注データシートでの処理終了、シート保護をかける
OrderSheet.Activate
OrderSheet.Range("A1").Select
OrderSheet.Protect

'シート削除、保存時のアラートダイアログを抑止
Application.DisplayAlerts = False

'テンプレートシートを削除
Worksheets("ピッキングシート提出用テンプレート").Delete
Worksheets("振分用テンプレート").Delete

ShowProgress.ProgressBar.Value = 14
ShowProgress.StepMessageLabel = Mall & "保存処理中"

Dim DeskTop As String, SaveFileName As String, SavePath As String
Const SAVE_FOLDER = "\\server02\商品部\ネット販売関連\ピッキング\クロスモール\過去データ\"

SaveFileName = "ピッキング・振分" & Format(Date, "MMdd")

OrderSheet.Activate
If Dir(SAVE_FOLDER, vbDirectory) <> "" Then
    '既に本日ファイルがあれば、時刻付けて保存
    If Dir(SAVE_FOLDER & SaveFileName & ".xlsx") = "" Then
        SavePath = SAVE_FOLDER & SaveFileName
    Else
        SavePath = SAVE_FOLDER & SaveFileName & "-" & Format(Time, "hhmm")
    End If
    
        ActiveWorkbook.SaveAs Filename:=SavePath, FileFormat:=xlWorkbookDefault

Else
    
    Dim DeskTopPath As String
    If Dir(DeskTopPath & SaveFileName & ".xlsx") = "" Then
        DeskTopPath = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\" & SaveFileName
    Else
        DeskTopPath = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\" & SaveFileName & "-" & Format(Time, "hhmm")
    End If
    
    MsgBox "ネット販売関連に繋がらないため、" & SaveFileName & "をデスクトップに保存します。"
        
    ActiveWorkbook.SaveAs Filename:=DeskTopPath, FileFormat:=xlWorkbookDefault

End If

ShowProgress.ProgressBar.Value = 15
ShowProgress.StepMessageLabel = Mall & "振分シート プリント"

'実行PCデフォルトのプリンタでプリントアウト
'プリンタの指定してなければ、Windowsのデフォルトプリンタを使う。
Dim i As Long
For i = 2 To Worksheets.Count

    Worksheets(i).Protect
    Worksheets(i).PrintOut

Next

'プログレスバーを消して終了メッセージ
ShowProgress.Hide
MsgBox prompt:="処理完了", Buttons:=vbInformation, Title:="処理終了"

End Sub
