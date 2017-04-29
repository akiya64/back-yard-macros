Attribute VB_Name = "Output"
Option Explicit

Sub QtyCsv()
'フィルター済のヤフーデータシート
'FileSystemObjectのテキストストリームでCSVファイルを生成して、TextStreamで内容を流し込みます。
'数秒で終わります。

With yahoo6digit 'ヤフーデータの下準備

    .Activate
          
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

