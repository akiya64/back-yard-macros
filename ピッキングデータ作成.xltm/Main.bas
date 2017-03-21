Attribute VB_Name = "Main"
Option Explicit

Sub ピッキング_振分作成()

OrderSheet.Activate

If Not Range("A2").Value = "" Then
    MsgBox "データ取得済です。"
    End
End If

'プログレスバーの準備
With ShowProgress
    .ProgressBar.Min = 1
    .ProgressBar.Max = 8

    .Show vbModeless
End With

ShowProgress.ProgressBar.Value = 2
ShowProgress.StepMessageLabel = "CSV読込中"

Call LoadCsv

ShowProgress.ProgressBar.Value = 3
ShowProgress.StepMessageLabel = "ロケーションデータ取得中"
Application.Wait Now + TimeValue("00:00:01")
'1秒待機してプログレスバーを更新

'DB接続、ロケーション取得、受注データの修正
Call ConnectDB.Make_List
Call DataVaridate.ModifyOrderSheet

'受注データシートでの処理終了、シート保護をかける
OrderSheet.Protect

'モール毎の電算室提出データ保存、振分けシート作成
Dim Mall As Variant, Malls As Variant, ProgressStep As Long

Malls = Array("Amazon", "楽天", "Yahoo")
ProgressStep = 3

For Each Mall In Malls
        
    ProgressStep = ProgressStep + 1
    ShowProgress.ProgressBar.Value = ProgressStep
    ShowProgress.StepMessageLabel = Mall & "データ処理中"
    
    'モール毎の受注件数がゼロ件ならファイル生成しない。
    If WorksheetFunction.CountIf(OrderSheet.Range("F:F"), Mall & "*") = 0 Then GoTo Continue
    
    'ピッキングシート作成・保存
    Call BuildSheets.OutputPickingData(CStr(Mall))
    
    '振分け用シート作成
    Call BuildSheets.CreateSorterSheet(CStr(Mall))

Continue:

Next

'シート削除、保存時のアラートダイアログを抑止
Application.DisplayAlerts = False

'テンプレートシートを削除
Worksheets("ピッキングシート提出用テンプレート").Delete
Worksheets("振分用テンプレート").Delete

ShowProgress.ProgressBar.Value = 7
ShowProgress.StepMessageLabel = Mall & "保存処理中"

Dim DeskTop As String, SaveFileName As String, SavePath As String
Const SAVE_FOLDER = "\\server02\商品部\ネット販売関連\ピッキング\クロスモール\過去データ\"

SaveFileName = "ピッキング・振分" & Format(Date, "MMdd") & ".xlsx"


If Dir(SAVE_FOLDER, vbDirectory) <> "" Then
    '既に本日ファイルがあれば、時刻付けて保存
    If Dir(SAVE_FOLDER & SaveFileName) = "" Then
        SavePath = SAVE_FOLDER & SaveFileName
    Else
        SavePath = SAVE_FOLDER & Format(Time, "hhmm") & SaveFileName
    End If
    
        ActiveWorkbook.SaveAs Filename:=SavePath, FileFormat:=xlWorkbookDefault

Else
    
    Dim DeskTopPath As String
    If Dir(DeskTopPath & SaveFileName) = "" Then
        DeskTopPath = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\" & SaveFileName
    Else
        DeskTopPath = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\" & Format(Time, "hhmm") & SaveFileName
    End If
    
    MsgBox "ネット販売関連に繋がらないため、" & SaveFileName & "をデスクトップに保存します。"
        
    ActiveWorkbook.SaveAs Filename:=DeskTopPath, FileFormat:=xlWorkbookDefault

End If

ShowProgress.ProgressBar.Value = 8
ShowProgress.StepMessageLabel = Mall & "振分シート プリント"

'実行PCデフォルトのプリンタでプリントアウト
'プリンタの指定してなければ、Windowsのデフォルトプリンタを使う。
Dim i As Long
For i = 2 To Worksheets.Count

    Worksheets(i).Protect
    Worksheets(i).PrintOut

Next

OrderSheet.Activate

'プログレスバーを消して終了メッセージ
ShowProgress.Hide
MsgBox Prompt:="処理完了", Buttons:=vbInformation, Title:="処理終了"

End Sub
