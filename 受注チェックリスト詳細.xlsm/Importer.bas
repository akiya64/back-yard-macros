Attribute VB_Name = "Importer"
Option Explicit
Sub CSV読込()

Worksheets("Santyoku受注データ").Activate

Dim CsvPath As String
CsvPath = GetOrderCheckListPath()

With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & CsvPath, Destination:=Range("$A$2"))
    .Name = "受注チェックリスト詳細読込"
    .FieldNames = False
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .TextFilePromptOnRefresh = False
    .TextFilePlatform = 932
    .TextFileStartRow = 2
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierDoubleQuote
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = False
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = True
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(2, 2, 2, 1, 9, 9, 9, 2, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, _
    9, 9, 9, 9, 5, 5, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 2, 1, 2, 9, 9, 9, 9 _
    , 9, 9, 9, 2, 9, 9, 9, 2, 2, 2, 2, 2, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, _
    9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 5, 9, 9, 9, 9, 9, 9, 9)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With

ActiveWorkbook.Connections(1).Delete


'取込日の日付チェック 最初の注文行と、最後の注文行の日付に対して

Dim LastRow As Long
LastRow = Range("Q1").SpecialCells(xlCellTypeLastCell).Row

If DateDiff("D", Cells(2, 17).Value, DateValue(Date)) <> 0 _
    Or DateDiff("D", Cells(LastRow, 17).Value, DateValue(Date)) <> 0 Then

    Dim ContinueWrongDate As VbMsgBoxResult
        ContinueWrongDate = MsgBox(Buttons:=vbExclamation + vbOKCancel, Prompt:="産直への取込日が本日ではありません。" & vbLf & "処理を続行します。" & vbLf & vbLf & "取込データ記載の取込日:" & Range("Q2").Value)
    
    If ContinueWrongDate <> vbOK Then
        '続行しない場合、データを消してマクロ終了
        Worksheets("Santyoku受注データ").UsedRange.Offset(1, 0).Clear
        End
        
    End If
    
End If

End Sub
Private Function GetOrderCheckListPath() As String
'ピッキングシートの-a＝棚無のセット分解前ファイルを探してフルパスをセット

Const SANTYOKU_DUMP_FOLDER As String = "\\Server02\商品部\ネット販売関連\梱包室データ\ARY受注チェックリスト\" '末尾\マーク必須

'実行時バインディング
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim f As Object, TodayCSV As Object
      
'指定フォルダー内のFileNameを含むファイル名を調べて、本日 日付ファイルを一つ取得する

For Each f In FSO.GetFolder(SANTYOKU_DUMP_FOLDER).Files

    If DateDiff("D", f.DateLastModified, DateValue(Date)) = 0 Then
    
        Set TodayCSV = f
    
        Exit For
    End If

Next

'本日日付のファイルがなければ、一旦マクロ終了
'TODO:本日日付ファイルがなければファイル指定ダイアログを出して手動セット
If TodayCSV Is Nothing Then End

GetOrderCheckListPath = SANTYOKU_DUMP_FOLDER & TodayCSV.Name

End Function
