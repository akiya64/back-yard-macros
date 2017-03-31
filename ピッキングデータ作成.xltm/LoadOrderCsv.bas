Attribute VB_Name = "LoadOrderCsv"
Option Explicit

Sub LoadCsv(Optional ByVal bool As Boolean)
'クロスモールからダウンロードしたCSV読込

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
    
    .TextFileColumnDataTypes = Array(2, 2, 2, 2, 1, 9, 9, 9, 9, 1, 1, 9, 9, 9, 1)
    
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=True
End With

'クロスモールのCSVが読み込まれたかチェック クロスモール側で採番する連番は数字8ケタ
If Not Range("A2").Value Like String(8, "#") Then
    MsgBox prompt:="読込んだファイルにクロスモールの連番がありません。" & vbLf & "処理を終了します。", Buttons:=vbCritical, Title:="正しくないファイル"
    End
End If

'アドイン用のコード修正、セット分解
Call FixForAddin
Call SetParse

ActiveWorkbook.Connections(1).Delete

End Sub

Private Sub FixForAddin()
Dim CodeRange As Range, c As Range
Set CodeRange = Range(Cells(2, 2), Cells(Range("B1").SpecialCells(xlCellTypeLastCell).Row, 2))

'アドイン用のコードを記入する
For Each c In CodeRange
    
    Dim CurrentCodeCell As Range
    Set CurrentCodeCell = c
    
    'I列、アドイン実行用に6ケタ化したコード、もしくはJANを入れる
    Dim ForAddinCell As Range
    Set ForAddinCell = Cells(c.Row, 9)
    
    ForAddinCell.NumberFormatLocal = "@"
    
    '6ケタならそのまま入れる
    If CurrentCodeCell.Value Like String(6, "#") Then
        ForAddinCell.Value = CurrentCodeCell.Value
    
    '数字5ケタは頭にゼロを追記
    ElseIf CurrentCodeCell.Value Like String(5, "#") Then
        
        ForAddinCell.Value = "0" & CurrentCodeCell.Value
    
    'JANもそのまま入れる
    ElseIf CurrentCodeCell.Value Like String(13, "#") Then
        
        ForAddinCell.Value = CurrentCodeCell.Value
    
    End If

    '必要数量、一旦受注の数量で埋める。セット分解後に書き換えられる。
    Cells(c.Row, 10).Value = Cells(c.Row, 4).Value

    '○個組分解
    If c.Value Like "*-*" Then
    
        Call SetParser.ParseScalingSet(c)
    
    End If

Next

End Sub

Private Sub SetParse()
'77777 セット分解  行の挿入を伴う処理なので単体で全レコードへ行う

Dim ForAddinRange As Range, c As Range
Set ForAddinRange = Range(Cells(2, 9), Cells(Range("B1").SpecialCells(xlCellTypeLastCell).Row, 9))

For Each c In ForAddinRange
    '7777始まりセット分解
    If c.Value Like "7777*" Then

        Call SetParser.ParseItems(c)
    
    End If
    
Next c

'セット商品ブックを閉じる
Call SetParser.CloseSetMasterBook

End Sub
