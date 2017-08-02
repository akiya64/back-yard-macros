Attribute VB_Name = "AppendHoldBook"
Option Explicit

Sub AppendHoldPurWokbook(ByVal HoldBook As Workbook)
'発注保留リストに、本日の手配保留商品を追記

With Worksheets(1)
    .Activate

    '列数を合わせる
    Columns("B").Insert
    Columns("B").ColumnWidth = 10
    Range("B1").Value = "備考A"
    
    Columns("D").Insert
    Columns("D").ColumnWidth = 10
    Range("D1").Value = "連番"
    
    Columns("H").Insert
    Columns("H").ColumnWidth = 10
    Range("H1").Value = "備考"
    
    Columns("J").Insert
    Range("I:I").Copy Destination:=Range("J:J")
    

    Dim EndRow As Integer
    EndRow = .Range("A1").SpecialCells(xlCellTypeLastCell).Row
    
    '発注保留シートのコピー範囲選択、A～M列で2行目から最終行の範囲取得
    Dim HoldProductRange As Range
    Set HoldProductRange = .Range("A2:M" & EndRow)
    
End With

'発注保留を開く
Dim HoldXlsxPath As String
HoldXlsxPath = "\\Server02\商品部\ネット販売関連\発注関連\c注文保留分.xlsx"

Dim HoldLogWorkbook As Workbook
Set HoldLogWorkbook = FetchWorkBook(HoldXlsxPath)

'保留一覧へコピー
With HoldLogWorkbook.Worksheets("保留一覧")

    'フィルターを解除
    If Not .AutoFilter Is Nothing Then
        Range("A1").AutoFilter
    End If
    
    '空行削除、保留後に手配した際に保留リストからデータ移動させるため、空行があるかも
    Call DeleteEmptyRow(HoldLogWorkbook.Worksheets("保留一覧"))
    
    '最終行の日付チェック、保留一覧シートでは文字列で保持しているため、文字列同士で比較する
    If CStr(.Range("G1").End(xlDown).Value) = Format(Date, "Mdd") Then
        HoldLogWorkbook.Close
        Exit Sub
    End If
    
    'コピー実行
    Dim DestinationRange As Range
    Set DestinationRange = .Range("A1").End(xlDown).Offset(1, 0)
    
    HoldProductRange.Copy
    DestinationRange.PasteSpecial (xlPasteValues)
    
    Range("A1").Select
    
    HoldLogWorkbook.Save

End With

HoldLogWorkbook.Close

End Sub

Private Sub DeleteEmptyRow(HoldWorkSheet As Worksheet)
'空白行を削除、行を走査、空白行のレンジを取得してRangeオブジェクトのメソッドでまとめて削除
'参考URL  https://www.moug.net/tech/exvba/0050065.html

With HoldWorkSheet

    'UsedRangeプロパティなら、空行も含めて最終セルを取得できる
    Dim UsedRowsCount As Long
    UsedRowsCount = .UsedRange.Rows.Count
    
    Dim i As Long, Target As Range

    '1列目のセルが空なら、Targetレンジにレンジを追加していく
    For i = 2 To UsedRowsCount
        
        Dim c As Range
        Set c = .Cells(i, 3)
        
        If c.Value = "" Then
            
            If Target Is Nothing Then
                Set Target = c.EntireRow
            Else
                Set Target = Union(Target, c.EntireRow)
            End If
            
        End If
    Next

    'Targetレンジ＝空行を一括削除
    If Not Target Is Nothing Then
        Target.Delete
    End If

End With

End Sub
