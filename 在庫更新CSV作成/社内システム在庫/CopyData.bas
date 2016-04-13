Attribute VB_Name = "CopyData"
Sub CopyDataBySyokonVendor()
Attribute CopyDataBySyokonVendor.VB_ProcData.VB_Invoke_Func = "r\n14"
'商魂マスターを仕入れ先別でフィルター表示して、
'その表示されているコードについて、ヤフーデータの行を別のブックへコピー
'1仕入れ先につき、1シートにコピー

'コードリストの準備
SyokonMaster.Activate

'フィルターした表示領域
Dim A As Range
Set A = Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible)

'コード列のレンジ
Dim B As Range
Set B = Range("A2").Resize(Range("A1").SpecialCells(xlCellTypeLastCell).row - 1, 1)

'ABの交差レンジをCodeレンジとしてセット=特定仕入れ先のコードレンジを取得できる。
Dim CodeRange As Range
Set CodeRange = Application.Intersect(A, B)


'コピー先ブックの指定
Dim DestinationBook As Workbook
Set DestinationBook = Workbooks.Add

Dim VendorName As String
VendorName = CodeRange.Cells(1, 4).Value

'コピー先ブックに新しいシートを用意
Set NewSheet = DestinationBook.Worksheets.Add()
NewSheet.Name = VendorName

Dim DestinationSheet As Worksheet
Set DestinationSheet = DestinationBook.Worksheets(VendorName)

yahoo6digit.Rows(1).Copy Destination:=DestinationSheet.Rows(1)

SyokonMaster.Activate

Dim r As Range

For Each r In CodeRange
    
    Code = Right(r.Value, 5)
    
    On Error Resume Next
        FindRow = WorksheetFunction.Match(CDbl(Code), yahoo6digit.Range("YahooCodeRange"), 0)
    
    If Err Then
        GoTo continue
    Else
        yahoo6digit.Rows(FindRow).Copy Destination:=DestinationSheet.Rows(DestinationSheet.UsedRange.Rows.Count + 1)
    End If
    
    On Error GoTo 0

continue:

Next

MsgBox VendorName & " コピー完了"

End Sub

Sub ExtractYahooData()

'コピー先ブックの指定
Dim DestinationBook As Workbook
Set DestinationBook = Workbooks.Add

ThisWorkbook.Worksheets("ヤフーデータ").Rows(1).Copy Destination:=DestinationBook.Sheets(1).Rows(1)

'抽出したいコードリストの用意
Dim CodeRange As Range
Set CodeRange = Workbooks(2).Sheets(1).Range("B2:B1410")

Dim r As Range

For Each r In CodeRange
    
    Code = r.Value
    
    On Error Resume Next
        FindRow = WorksheetFunction.Match(CDbl(Code), yahoo6digit.Range("YahooCodeRange"), 0)
    
    If Err Then
        r.Interior.ColorIndex = 6
        GoTo continue
    Else
        yahoo6digit.Rows(FindRow).Copy Destination:=DestinationBook.Sheets(1).Rows(DestinationBook.Sheets(1).UsedRange.Rows.Count + 1)
    End If
    
    On Error GoTo 0

continue:

Next

MsgBox VendorName & " コピー完了"

End Sub

