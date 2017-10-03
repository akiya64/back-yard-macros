Attribute VB_Name = "UpdateProductName"
Sub Update()

'Do構文内の商品コードの列番号、商品名列番号、ワークブック・シート番号、3カ所指定し直す。

'更新元データのブック・シートを指定
Dim SourceSheet As Worksheet
Set SourceSheet = Workbooks(1).Worksheets(1)

Worksheets("最終").Activate

'「最終」シートに既に入っている商品コードのレンジ
Dim CodeRange As Range
Set CodeRange = Range(Cells(1, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 1))

Dim i As Long, Code As String, Name As String
i = 2

Do
    '更新したい商品名のシートの商品コード列
    Code = SourceSheet.Cells(i, 1).Value

    Dim HitRow As Long
    HitRow = WorksheetFunction.Match(Code, CodeRange, 0)
    
    Cells(HitRow, 2).Value = SourceSheet.Cells(i, 13).Value
    
    i = i + 1
    
Loop Until IsEmpty(SourceSheet.Cells(i, 1).Value)

End Sub

