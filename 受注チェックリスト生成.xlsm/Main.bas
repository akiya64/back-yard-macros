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


Transfer.アップロード用シートへ転記

'セット商品リストブックを閉じる
Dim w As Workbook
For Each w In Workbooks
    If w.Name = "ｾｯﾄ商品ﾘｽﾄ.xls" Then w.Close False
Next

Worksheets("アップロードシート").Range("A1").Select

'取込、作業シート込みのエクセルファイルを保存
Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs FileName:=ActiveWorkbook.Path & "\受注チェックリスト_" & Format(Date, "yyyymmdd") & ".xlsm", FileFormat:=52
    'ActiveWorkbook.SaveAs FileName:=ActiveWorkbook.Path & "\受注チェックリスト.xlsx", FileFormat:=xlOpenXMLWorkbook
    
Application.DisplayAlerts = True

'データベースへの登録処理実行
Call InsertDB.CodeI2JAN_E

End Sub


