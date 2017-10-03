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
Sub FillJan()

For i = 2 To 182338

If IsEmpty(Cells(i, 1).Value) Then

    Sku = Cells(i, 2).Value

    If Len(Sku) = 13 And Not Sku Like "77777*" And Not Sku Like "88888*" Then
    
        Cells(i, 1).NumberFormatLocal = "@"
        Cells(i, 1).Value = Sku
    
    End If

End If

Next

End Sub

