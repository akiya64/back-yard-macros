Attribute VB_Name = "LoadOrderCsv"
Option Explicit

Sub LoadCsv(Optional ByVal bool As Boolean)
'クロスモールからダウンロードしたCSVの読込
'ファイル指定ダイアログを表示し、ネット販売関連\ピッキング\クロスモール\のCSVを指定する。

'クロスモールのCSVをダウンロードしているフォルダへ移動、ファイル指定ダイアログを開く
'@url http://officetanaka.net/other/extra/tips15.htm
CreateObject("WScript.Shell").CurrentDirectory = "\\server02\商品部\ネット販売関連\ピッキング\クロスモール\"

Dim FilePath As String
FilePath = Application.GetOpenFilename("クロスモールCSV,*.csv", 2, "クロスモールのピッキングCSVを指定")

If FilePath = "False" Then
    MsgBox "ファイル指定がキャンセルされました。" & vbLf & "マクロを終了します。"
    End
End If

If DateDiff("D", FileDateTime(FilePath), Date) <> 0 Then
    Dim IsContinue As Integer
    IsContinue = MsgBox(prompt:="本日のダウンロードファイルではありません。" & vbLf & "ピッキングシートを生成しますか？", Buttons:=vbYesNo + vbQuestion)

    If IsContinue = vbNo Then
        MsgBox "処理を終了します。"
        End
    End If
End If

'マクロ起動ボタン削除
OrderSheet.Shapes(1).Delete

'データ接続を利用してCSVデータを読み込み
With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & FilePath, Destination:=Range("$A$2"))
    .Name = "受注チェックリスト詳細読込"
    .FieldNames = False
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .RefreshStyle = xlOverwriteCells '既定のxlInsertDeleteCellsでは空行が挿入されることがある
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
    
    .TextFileColumnDataTypes = Array(2, 2, 2, 2, 1, 9, 9, 9, 9, 1, 1, 9, 9, 9, 1)
    
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=True
End With

'CSVへのデータ接続削除、クエリーテーブルに名前が付くので削除
ActiveWorkbook.Connections(1).Delete
ActiveWorkbook.Names(1).Delete

'クロスモールのCSVが読み込まれたかチェック クロスモール側で採番する連番は数字8ケタ
If Not Range("A2").Value Like String(8, "#") Then
    MsgBox prompt:="読込んだファイルにクロスモールの連番がありません。" & vbLf & "処理を終了します。", Buttons:=vbCritical, Title:="正しくないファイル"
    End
End If

End Sub
