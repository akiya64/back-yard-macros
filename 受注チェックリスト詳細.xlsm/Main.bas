Attribute VB_Name = "Main"
Option Explicit

Sub 生成()

'Importer.産直CSVインポート

Transfer.作業シートへデータ抽出

Worksheets("作業シート").Activate

SetParser.セット分解
Transfer.住所結合
Transfer.FixJAN

End Sub
