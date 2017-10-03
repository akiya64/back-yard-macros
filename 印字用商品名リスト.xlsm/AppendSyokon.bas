Attribute VB_Name = "AppendSyokon"
Option Explicit

Sub AppendSyokonCode()

Dim i As Long, Code As String, Descript As String
i = 2

'開いているワークブックの二つ目なので、毎回指定し直す必要がある。
With Workbooks(2).Worksheets(1)

    Do
        Code = .Cells(i, 1).Value
        
        If Code Like "#####" Then Code = "0" & Code
        
        Descript = .Cells(i, 2).Value
        
        Call AppendProduct(Code, Descript)
        
        i = i + 1
    
    Loop Until IsEmpty(Cells(i, 1).Value)

End With

End Sub

Sub AppendProduct(ByVal Code As String, ByVal ProductName As String)

Dim MapperSheet As Worksheet
Set MapperSheet = ThisWorkbook.Sheets("最終")

If WorksheetFunction.CountIf(MapperSheet.Range("A1:A400000"), Code) = 0 Then

    With Worksheets("最終")
        Dim FinalRow As Long
        FinalRow = .Range("A1").SpecialCells(xlCellTypeLastCell).Row
        
        .Cells(FinalRow + 1, 1).NumberFormatLocal = "@"
        
        .Cells(FinalRow + 1, 1).Value = Code
        .Cells(FinalRow + 1, 2).Value = ProductName

    End With
    
End If

End Sub
