Attribute VB_Name = "Main"
Option Explicit

Sub ヤフー在庫更新ファイル生成()

'商魂の区分、在庫限りシート、廃番シートをチェックして、
'ヤフーにアップローする在庫数、Allow-overdraftをセットします。

'時間計測をします
Dim startTime As Long
startTime = Timer

'ヤフーCSVをインポート
Prepare.ImportYahooCSV

Prepare.ImportSyokonAddinData

'各シートのコード範囲を名前で呼び出せるよう再定義
Prepare.SetRangeName

'---準備完了---

'設定在庫数算出、取り寄せ可否算出

Compute.UploadQuantity

'一時停止を上書き
HaltSheet.setHalt

'在廃、処分で0個は廃番・終了へ移動
CheckEolInStockOnly

'区分の表示をフィルター
SetStatusFilter

'ヤフーデータシートからCSVを保存
Call Output.QtyCsv

ThisWorkbook.Save

'終了時刻を入れる
Dim endTime As Long
endTime = Timer

MsgBox "実行時間：" & endTime - startTime & " 秒"

End Sub
