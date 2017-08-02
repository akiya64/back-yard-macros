Attribute VB_Name = "Importer"
Option Explicit
Sub CSV読込()

Worksheets("受注データ").Activate

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
    .TextFileColumnDataTypes = Array(2, 2, 2, 1, 1, 5, 5, 2, 2, 2, 2, 2, 2, 2, 2)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=True
End With

ActiveWorkbook.Connections(1).Delete

'マクロ起動ボタン削除
Worksheets("受注データ").Shapes(1).Delete

End Sub
Private Function GetOrderCheckListPath() As String
'フォルダを指定してファイル指定ダイアログからファイル指定

Const CSV_DL_FOLDER As String = "\\server02\商品部\ネット販売関連\ピッキング\クロスモール" '末尾\マーク必須

Dim FilePath As String

'カレントフォルダを移動して梱包室データフォルダでファイル指定ダイアログを開く
'@url http://officetanaka.net/other/extra/tips15.htm
CreateObject("WScript.Shell").CurrentDirectory = CSV_DL_FOLDER

FilePath = Application.GetOpenFilename("クロスモールCSV,*.csv", 2, "クロスモールのピッキングCSVを指定")

If FilePath = "False" Then
    MsgBox "ファイル指定がキャンセルされました。" & vbLf & "マクロを終了します。"
    End
End If

GetOrderCheckListPath = FilePath

End Function
