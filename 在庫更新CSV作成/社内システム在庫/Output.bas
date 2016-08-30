Attribute VB_Name = "Output"
Option Explicit

Sub QtyCsv()
'FileSystemObjectのテキストストリームでCSVファイルを生成して、TextStreamで内容を流し込みます。
'数秒で終わります。

With yahoo6digit 'ヤフーデータの下準備

    .Activate
    
    '「"登録なし"」と「"空白"」 この2つ以外をフィルターで表示…TODO：1列目からフィルターの状況をチェックさせた方がいい
    '16-2-29 廃番の区分が「メ廃番」になりました。
    
    Dim StatusCol As Integer
    StatusCol = .Rows(1).Find("status").Column
    
    .Range("A1").CurrentRegion.AutoFilter Field:=StatusCol, Criteria1:=Array( _
            "ＳＰ扱い", "限定品", "在庫廃番", "在庫処分", "棚なしに有", "棚なし完売", "直送扱い", "登録のみ", "メ廃番品", "販路限定", "販売中止", "標準" _
            ), Operator:=xlFilterValues
    
    'フィルターしたレンジをセット、CSVのヘッダーは別途書き込んでおくので、2行目以降のレンジ。
    Dim A As Range
    Set A = .Range("C1").CurrentRegion.SpecialCells(xlCellTypeVisible)
    
    Dim B As Range
    Set B = .Range("C2").Resize(Range("C1").SpecialCells(xlCellTypeLastCell).Row - 1, 1)
    
    Dim CodeRange As Range
    Set CodeRange = Application.Intersect(A, B)

End With

'書き出し用CSVを用意
Dim day As String
day = Format(Date, "mm") & Format(Date, "dd")

Dim OutputCsvName As String
OutputCsvName = "商魂在庫アップ用" & day & ".csv"

Dim FSO As Object 'TODO:事前バインディングに変更
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim TS As Object
    
Set TS = FSO.CreateTextFile(Filename:=ThisWorkbook.Path & "\" & OutputCsvName, _
                            OverWrite:=True)
                            
'ヘッダーを書き込み
Dim header As Variant
header = "code,quantity,allow-overdraft"

TS.WriteLine header

Dim colQuantity As Long, colAllow As Long, colStatus As Long

colQuantity = yahoo6digit.Rows(1).Find("quantity").Column
colAllow = yahoo6digit.Rows(1).Find("allow-overdraft").Column
colStatus = yahoo6digit.Rows(1).Find("status").Column

'コードレンジに対して、r.rowで行番号を取り出して同じ行のQuantity/Allowの値を取得する
Dim r As Range, Code As String, Qty As String, Pur As String

For Each r In CodeRange
    
    Code = r.Value
    
    Qty = Cells(r.Row, colQuantity).Value
    Pur = Cells(r.Row, colAllow).Value
    
    TS.WriteLine Code & "," & Qty & "," & Pur

Next

TS.Close

End Sub

