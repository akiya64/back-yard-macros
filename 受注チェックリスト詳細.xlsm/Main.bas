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


'提出用ファイル作成処理
Transfer.提出用シートへ転記

MsgBox "シート作成 完了"

End Sub
