Attribute VB_Name = "BuildShippingNumSheet"
Option Explicit
Sub 佐川ヤマト_シート作成()

'ファイル保存
Application.DisplayAlerts = False
    ThisWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\出荷確認_" & Format(Date, "yyyyMMdd") & ".xlsx", FileFormat:=xlWorkbookDefault
Application.DisplayAlerts = True


'TSV/CSVファイル読込
Call LoadAllCsv

'ボタン削除
Worksheets("トップ").Shapes(1).Delete

'データ取得後処理  データリンク削除＆セルの「名前」削除
Dim qt As QueryTable
For Each qt In Worksheets("トップ").QueryTables
    qt.Delete
Next qt

Dim nm As Name
For Each nm In ActiveWorkbook.Names
    nm.Delete
Next nm

'運送会社別にシートへコピー
'運送会社毎の送り状番号冒頭5ケタは、SortByCarrierプロシージャにてハードコーディング
Call SortByCarrier("佐川急便")
Call SortByCarrier("ヤマト運輸")

'列幅調整
Dim i As Long
For i = 1 To Worksheets.Count
    Worksheets(i).Range("A1").CurrentRegion.Columns.AutoFit
Next i

'振分後の保存と完了メッセージ
ThisWorkbook.Save

MsgBox "完了"
End Sub

Private Sub LoadAllCsv()
'送り状番号CSVパスを取得、3ファイルまで同時指定可能

'ファイルダイアログにてAmazon・楽天・ヤフーのTSV/CSVファイルを複数同時指定してもらう
Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd
    'ファイル選択ダイアログの設定
    .Filters.Clear
    .Filters.Add "Amazon,楽天,Yahoo!", "*.tsv; *.csv"
    .InitialFileName = "\\Server02\商品部\ネット販売関連\梱包室データ\出荷通知"
    
    'ダイアログ表示
    .Show
    
    'ファイル選択後の処理
    If .SelectedItems.Count = 0 Then
    
        MsgBox "ファイル指定がキャンセルされました。"
        End
    
    ElseIf .SelectedItems.Count >= 4 Then
        MsgBox "ファイル指定が3つを超えています。"
        End
    
    End If
    
    Dim Paths(2) As String, CurrentPath As String, i As Long
    
    '選択されたファイルパスから中身調べてモール毎にセット
    For i = 1 To .SelectedItems.Count
        Call LoadCsv(.SelectedItems.Item(i))
    Next
    
End With

End Sub

Private Sub LoadCsv(ByVal Path As String)
'引数のパスにテキストストリームで接続してヘッダーを調べ、モール名を返す。
Dim FSO As Object, TS As Object, i As Long, CurrentMall As String, CurrentRow As Variant

Set FSO = CreateObject("Scripting.FileSystemObject")
Set TS = FSO.OpenTextFile(Path)
        
Do Until TS.AtEndOfStream Or i > 3
    CurrentRow = TS.ReadLine
    
    'タブがあればAmazon
    If InStr(CurrentRow, Chr(9)) > 0 Then
        Call LoadAmazon(Path)
        Exit Do
    
    '受注番号 の文言があれば楽天
    ElseIf InStr(CurrentRow, "受注番号") > 0 Then
        Call LoadRakuten(Path)
        Exit Do
        
    'OrderId の文言があればヤフー
    ElseIf InStr(CurrentRow, "OrderId") > 0 Then
        Call LoadYahoo(Path)
        Exit Do
    
    End If
    
    i = i + 1

Loop

End Sub

Sub SortByCarrier(ByVal CarrierName As String)
'運送会社名を受け取って、運送会社毎のシートへ送り状番号をコピー
'引数でコピー先シートを指定するので、シート名はCase構文の運送会社と合わせること。

Dim Criteria As Variant, Operator As Integer

Worksheets("トップ").Activate

'佐川の送り状番号冒頭4ケタ、フィルターのプロパティになるので配列へ代入
Criteria = Array("4031*", "4012*")

Select Case CarrierName
    
    Case "佐川急便"
        'Criteria配列に入っている送り状番号でフィルター
        Operator = xlFilterValues
        
        With Range("A1").CurrentRegion
            .AutoFilter Field:=3, Criteria1:=Criteria, Operator:=Operator
            .Copy Worksheets(CarrierName).Range("A1")
            .AutoFilter 'オートフィルター解除
        End With
        
    Case "ヤマト運輸"
        '佐川以外の送り状番号をフィルター、Criteriaは2までしかセット出来ない。
        
        Operator = xlAnd
        
        With Range("A1").CurrentRegion
            .AutoFilter Field:=3, Criteria1:="<>" & Criteria(0), Operator:=Operator, Criteria2:="<>" & Criteria(1)
            .Copy Worksheets(CarrierName).Range("A1")
            .AutoFilter 'オートフィルター解除
        End With
End Select

End Sub

Sub LoadAmazon(ByVal Path As String)

With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & Path, Destination:=GetDestRange())
    .Name = "Amazon"
    .FieldNames = True
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
    .TextFileStartRow = 4
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierDoubleQuote
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = True
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = False
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(2, 9, 9, 9, 9, 9, 2, 9, 9)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With

Call FillMallName("Amazon")

End Sub

Sub LoadRakuten(ByVal Path As String)

With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & Path, Destination:=GetDestRange())
    .Name = "楽天"
    .FieldNames = True
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
    .TextFileColumnDataTypes = Array(2, 9, 9, 2, 9)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With

Call FillMallName("楽天")

End Sub

Sub LoadYahoo(ByVal Path As String)

With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & Path, Destination:=GetDestRange()) 'パスと書き出し先は動的に決める
    .Name = "yahoo"
    .FieldNames = True
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
    .TextFileColumnDataTypes = Array(2, 9, 2, 9, 9, 9)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With

Call FillMallName("Yahoo")

End Sub

Private Sub FillMallName(ByVal MallName As String)
'CSV読み込み後にA列をモール名で埋めます。

Dim StartRow As Double, EndRow As Double, i As Double
StartRow = IIf(Range("A2").Value = "", 2, Range("A1").End(xlDown).Row + 1)
EndRow = Range("B1").End(xlDown).Row

For i = StartRow To EndRow
    Cells(i, 1).Value = MallName
Next i

End Sub

Private Function GetDestRange() As Range

'書き出し先セルを決める、B2が空の時にEndコマンドでは1,048,576行まで飛んでしまうので。
Dim r As Range
If IsEmpty(Range("B2")) Then
    Set r = Range("B2")
Else
    Set r = Range("B1").End(xlDown).Offset(1, 0)
End If

Set GetDestRange = r

End Function
