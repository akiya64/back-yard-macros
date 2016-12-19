Attribute VB_Name = "Main"
Option Explicit

Sub 受注チェックリスト生成()

'CSV読込、作業シートへコピー
Importer.CSV読込
Transfer.作業シートへデータ抽出

'作業シートでのデータ修正処理
Worksheets("作業シート").Activate

SetParser.セット分解
Transfer.住所結合
Transfer.JAN転記
Transfer.楽天商品名修正


Transfer.提出用シートへ転記

'提出ファイル保存
Sheets("提出シート").Copy

Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs FileName:="提出" & Format(Date, "MMdd") & ".xlsx"
    ActiveWorkbook.Close
Application.DisplayAlerts = True

Dim w As Workbook

For Each w In Workbooks
    If w.Name = "ｾｯﾄ商品ﾘｽﾄ.xls" Then w.Close False
Next

MsgBox "ファイル作成 完了"

ThisWorkbook.Close False

End Sub
