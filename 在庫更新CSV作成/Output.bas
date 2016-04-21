Attribute VB_Name = "Module1"
Sub AppendQtyCsv()

Const CSV_SH_NAME As String = "アップロード用在庫"

Application.ScreenUpdating = False

'CSV追記準備
Dim FSO As New FileSystemObject
Dim Csv As Object

Dim CsvFileName As String
CsvFileName = "在庫更新" & Format(Date, "mmdd") & ".csv"

'追記モード ForAppending でファイルを開く
Set Csv = FSO.OpenTextFile(FileName:=ThisWorkbook.Path & "\" & CsvFileName, IOMode:=8)

Dim LastRow As Long
LastRow = Worksheets(CSV_SH_NAME).UsedRange.Rows.Count

For i = 2 To LastRow
    
    With Worksheets(CSV_SH_NAME)
        Csv.WriteLine (CStr(.Cells(i, 1).Value) & "," & CStr(.Cells(i, 2).Value) & "," & CStr(.Cells(i, 3).Value))
    End With

Next
End Sub

