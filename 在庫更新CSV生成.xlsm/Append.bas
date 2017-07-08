Attribute VB_Name = "Append"
Sub AppendCode(ByVal Code As String, ByVal RangeName As String, Optional RowNumber As Variant)
'リストにコードを加える

'既にリストアップ済みのコードでないかチェック
If WorksheetFunction.CountIf(ThisWorkbook.Names(RangeName).RefersToRange, Code) > 0 Then Exit Sub

Dim N As Name
Set N = ThisWorkbook.Names(RangeName)

Dim CutLength As Integer
CutLength = InStr(2, N.Value, "!") - 2

Dim SheetName As String
SheetName = Mid(N, 2, CutLength)

If IsMissing(RowNumber) Then 'IsMissing関数を使う、判定したい引数がVariantでないと判定ビットが含まれない

    'Yahoo6digitsシートの該当行を特定する
    On Error Resume Next
        
        RowNumber = WorksheetFunction.Match(CDbl(Code), yahoo6digit.Range("YahooCodeRange"), 0)
        If Err Then Exit Sub
    
    On Error GoTo 0
End If

'商品レコードをコピー
With ThisWorkbook.Worksheets(SheetName)
    
    yahoo6digit.Rows(RowNumber).Copy Destination:=.Rows(.UsedRange.Rows.Count + 1)
    
    'ヤフーデータの方はグレーで塗る
    yahoo6digit.Range("A" & RowNumber & ":I" & RowNumber).Interior.ColorIndex = 15

End With

End Sub
