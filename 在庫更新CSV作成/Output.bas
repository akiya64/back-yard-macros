Attribute VB_Name = "Output"
Option Explicit

Sub AppendQtyCsv()

Dim startTime As Long
startTime = Timer

'準備
Call FetchSecondInventry
'各シートのコード範囲を名前で呼び出せるよう再定義
Call SetRangeName

'商魂データから全アイテムに在庫をセット
Call SetQuantity

With yahoo6digit

    .Activate
       
    Dim StatusCol As Integer
    StatusCol = .Rows(1).Find("status").Column
    
    .Range("A1").CurrentRegion.AutoFilter Field:=StatusCol, Criteria1:=Array( _
            "棚なしに有", "棚なし完売" _
            ), Operator:=xlFilterValues
    
    'フィルターしたレンジをセット
    Dim A As Range
    Set A = .Range("C1").CurrentRegion.SpecialCells(xlCellTypeVisible)
    
    Dim B As Range
    Set B = .Range("C2").Resize(Range("C1").SpecialCells(xlCellTypeLastCell).Row - 1, 1)
    
    Dim CodeRange As Range
    Set CodeRange = Application.Intersect(A, B)

End With

'書き出し用CSVシートを用意
Worksheets("CSV").Cells.Clear

'ヘッダーを書き込み
Dim header As Variant
header = Array("code", "quantity", "allow-overdraft")

Worksheets("CSV").Range("A1:C1") = header

Worksheets("ヤフーデータ").Activate

Dim colQuantity As Long, colAllow As Long
colQuantity = yahoo6digit.Rows(1).Find("quantity").Column
colAllow = yahoo6digit.Rows(1).Find("allow-overdraft").Column

Dim i As Long
i = 2
'コードレンジに対して、r.rowで行番号を取り出して同じ行のQuantity/Allowの値を取得する
Dim r As Range
For Each r In CodeRange

    Dim Code As String
    Code = r.Value

    Dim qty As Long, pur As String
    qty = Cells(r.Row, colQuantity).Value
    pur = Cells(r.Row, colAllow).Value

    Worksheets("CSV").Range("A" & i & ":C" & i) = Array(Code, qty, pur)

    i = i + 1

Next

Worksheets("CSV").Activate

'CSV追記準備
Dim FSO As New FileSystemObject
Dim Csv As Object

'追記モード ForAppending でファイルを開く
Set Csv = FSO.OpenTextFile(Filename:=ThisWorkbook.Path & "\" & "ヤフー在庫更新" & Format(Date, "mmdd") & ".csv", IOMode:=8)

For i = 2 To Worksheets("CSV").UsedRange.Rows.Count
    
    With Worksheets("Csv")
        Csv.WriteLine (CStr(.Cells(i, 1).Value) & "," & CStr(.Cells(i, 2).Value) & "," & CStr(.Cells(i, 3).Value))
    End With

Next

'終了時刻を格納
Dim endTime As Long
endTime = Timer

'ログシートへ処理時間を記録
Call ApendProcessingTime(endTime - startTime)

End Sub

