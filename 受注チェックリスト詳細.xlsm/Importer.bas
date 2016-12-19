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


'読み込み後、取込日の日付チェック 最初の注文行と最後の注文行の日付に対して

Dim LastRow As Long
LastRow = Range("Q1").SpecialCells(xlCellTypeLastCell).Row


'本日日付ならば、この処理は完了
If DateDiff("D", Cells(2, 17).Value, DateValue(Date)) = 0 _
    And DateDiff("D", Cells(LastRow, 17).Value, DateValue(Date)) = 0 Then
    
    Exit Sub

End If


'読込データが本日取込でなかった場合、続行可否をダイアログで決めてもらう。
Dim ContinueWrongDate As VbMsgBoxResult
    ContinueWrongDate = MsgBox(Buttons:=vbExclamation + vbYesNo, Prompt:="産直への取込日が本日ではありません。" & vbLf & "処理を続行しますか？" & vbLf & vbLf & "取込データ記載の取込日:" & Range("Q2").Value)

If ContinueWrongDate = vbNo Then
    
    '続行しない場合、データを消すかそのままにするか選択
    Dim ChooseDataClear As VbMsgBoxResult
    ChooseDataClear = MsgBox(Buttons:=vbExclamation + vbYesNo, Prompt:="読込済データを消去しますか？")
    
    If ChooseDataClear = vbYes Then
        'データ消去してマクロ全体を終了
        Worksheets("Santyoku受注データ").UsedRange.Offset(1, 0).Clear
        End
    
    Else
        'データ確認の上で続行する場合、続行用ボタンを追加。
        With ActiveSheet.Buttons.Add(709.5, 54, 201, 42)
            .OnAction = "作業シートへデータ抽出"
            .Characters.Text = "読込済データで処理を続行"
        End With
    
    End If
    
End If

End Sub
Private Function GetOrderCheckListPath() As String
'本日CSVを指定フォルダより探す。

Const SANTYOKU_DUMP_FOLDER As String = "\\Server02\商品部\ネット販売関連\梱包室データ\ARY受注チェックリスト\" '末尾\マーク必須

'実行時バインディングでファイルオブジェクト
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

'本日日付ファイルがなければファイル指定ダイアログを出して手動セット
If TodayCSV Is Nothing Then
    
    MsgBox Prompt:="本日の受注チェックリスト ファイルが見つかりませんでした。" & vbLf & "ファイルを指定して下さい。", _
            Buttons:=vbCritical
    
    'カレントフォルダを移動して梱包室データフォルダでファイル指定ダイアログを開く
    '@url http://officetanaka.net/other/extra/tips15.htm
    CreateObject("WScript.Shell").CurrentDirectory = "\\server02\商品部\ネット販売関連\梱包室データ\"
    
    Dim FilePath As String
    FilePath = Application.GetOpenFilename()
    
    Set TodayCSV = FSO.getfile(FilePath)

End If

GetOrderCheckListPath = SANTYOKU_DUMP_FOLDER & TodayCSV.Name

End Function
