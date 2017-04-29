Attribute VB_Name = "FixJan"
Option Explicit
Sub FixAllJAN()

Dim i As Long
For i = 2 To ActiveSheet.UsedRange.Rows.Count
        
    FixJanDigit (Cells(i, 1))
        
Next

End Sub

Sub FixJanDigit(ByVal Cell As Range)

If Not Cell.Value Like String(13, "#") And Cell.Value <> "" Then
    
    Dim AddZeroCount As Long
    AddZeroCount = 13 - Len(Cell.Value)
    
    If AddZeroCount > 0 Then
        Cell.NumberFormatLocal = "@"
        Cell.Value = String(AddZeroCount, "0") & Cell.Value
    End If

End If

End Sub
