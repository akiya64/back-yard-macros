Attribute VB_Name = "Module1"
Sub 読込集計_フルオート()
    'MeisaiSheet/paymentMethodを読み込んで式も入れて、新規シートへ保存して終了
    'CSV:MeisaiSheet,PaymentMethod
    
    Dim MonthName As String
    MonthName = Format(DateAdd("M", -1, Date), "yy年M月")
    
    Sheets("商品別集計").Range("A1") = MonthName & " ヤフー月次"
    
    Call meisaiCSVインポート
    
    Call 転記と重複削除
    Call 集計式の挿入
    Call 罫線を引く
    
    Dim FileName As String
    FileName = "ヤフー月次" & MonthName & "_作業中.xlsm"
    
    Dim Folder As String
    Folder = Environ("USERPROFILE") & "\Documents\"
    
    Dim Path As String
    Path = Folder & FileName
    
    ThisWorkbook.SaveAs FileName:=Path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    
    Call 商品別集計を新規シートへコピー
    
End Sub

Private Function findColum(str As String) As Integer

findColum = WorksheetFunction.Match(str, MeisaiSheet.Range("A1").Resize(1, 20), 0)

End Function

Sub meisaiCSVインポート()

Dim FilePath
FilePath = setCsvPath("Meisai")

If FilePath = "" Then
    MsgBox "ファイル指定がキャンセルされました。"
    Exit Sub
End If

MeisaiSheet.Activate

With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & FilePath, Destination:=Range("$A$1"))
    .Name = "Meisai"
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
    .TextFileStartRow = 1
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierDoubleQuote
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = False
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = True
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(2, 1, 1, 2, 2, 1, 1, 1, 1, 1, 1, 1)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With

End Sub

Private Function setCsvPath(CsvName As String)
'ファイル選択ダイアログを開いてファイル指定、パスを返す

' ｢ファイルを開く｣のフォームでファイル名の指定を受ける
Path = Application.GetOpenFilename(Title:=CsvName & "を指定")

' キャンセルされた場合はFalseが返るので以降の処理は行なわない
If VarType(Path) = vbBoolean Then Exit Function

setCsvPath = Path
    
End Function

Sub 転記と重複削除()
'売上商品の一意な表を用意します。
'MeisaiSheetのコードと商品名を集計シートに転記して重複削除します。
'RangeメソッドのRemoveDuplicatesを使う。

ItemTotalSheet.Activate

'商品別集計シートへの転記

On Error GoTo ErrorMes
   'Code列の特定
    Dim codeCol As Integer
    codeCol = WorksheetFunction.Match("Product Code", MeisaiSheet.Range("A1").Resize(1, 20), 0)
    
    'Description列の特定
    Dim DescriptionCol As Integer
    DescriptionCol = WorksheetFunction.Match("Description", MeisaiSheet.Range("A1").Resize(1, 20), 0)

On Error GoTo 0


'商品別集計へMeisaiシートから転記
'Code=商品コードとDescription=商品名を1行ずつ

With ItemTotalSheet
    
    Dim i As Long
    i = 2

    Do While MeisaiSheet.Cells(i, codeCol).Value <> ""
        
        '数量0はキャンセルなので飛ばす
        If MeisaiSheet.Cells(i, 3).Value = 0 Then GoTo Continue
        
        Dim WriteRow As Long
        WriteRow = .Range("A1").SpecialCells(xlCellTypeLastCell).Row + 1
        
        .Cells(WriteRow, 1).Value = MeisaiSheet.Cells(i, DescriptionCol)
        .Cells(WriteRow, 2).Value = MeisaiSheet.Cells(i, codeCol)
        
Continue:
    i = i + 1
    Loop
    
    'Rangeを指定してRangeオブジェクトのRemoveDuplicateメソッドで一発重複削除｡エクセル2010以降。
    
    .Range("A2:B2").Resize(.UsedRange.Rows.Count, 2).Name = "商品リスト"
    .Range("商品リスト").RemoveDuplicates Columns:=2, Header:=xlYes

End With


Exit Sub

ErrorMes:
MsgBox "処理を中止しました。" & vbLf & "MeisaiシートにProduct CodeとDescriptionがありません。"

End Sub

Sub 集計式の挿入()

'商品コードの最初のセルから最終行までのRangeを格納
'SUM関数を使うために、数値型に変換が必要なレンジもあわせて格納しておく。

Dim sh1EndRow As Long
sh1EndRow = MeisaiSheet.UsedRange.Rows.Count

'合計、個数、件数など集計対象の列をダブル型にキャスト
  
Dim Rng As Range
Set Rng = Union(MeisaiSheet.Cells(2, findColum("Quantity")).Resize(sh1EndRow, 1), _
                MeisaiSheet.Cells(2, findColum("Unit Price")).Resize(sh1EndRow, 1), _
                MeisaiSheet.Cells(2, findColum("Line Sub Total")).Resize(sh1EndRow, 1))

Dim c As Range
For Each c In Rng
    c.NumberFormat = "General" '表示形式を「標準」にセット
    c.Value = CDbl(c.Value)    'ダブル型にキャストして格納
Next

'Code列の特定
On Error GoTo ErrorMes
    Dim codeCol As Integer
    codeCol = WorksheetFunction.Match("Product Code", MeisaiSheet.Range("A1").Resize(1, 20), 0)
    
On Error GoTo 0

Dim MeisaiCodeRange As Range
Set MeisaiCodeRange = MeisaiSheet.Cells(2, codeCol).Resize(sh1EndRow - 1, 1)

i = 3

With ItemTotalSheet
    Do Until IsEmpty(.Cells(i, 2))
        
        '商品コードに対する売上金額を合計するSUMIF式、
        .Cells(i, "C").Formula = "=SUMIF(Meisai!" & MeisaiCodeRange.Address & ",B" & i & ",Meisai!" & MeisaiCodeRange.Offset(0, 6).Address & ")"
            
        '商品コードに対する注文件数を合計するCOUNTIF式、
        .Cells(i, "D").Formula = "=COUNTIF(Meisai!" & MeisaiCodeRange.Address & ",B" & i & ")"
    
        '商品コードに対する売上個数を合計するSUMIF式
        .Cells(i, "F").Formula = "=SUMIF(Meisai!" & MeisaiCodeRange.Address & ",B" & i & ",Meisai!" & MeisaiCodeRange.Offset(0, -1).Address & ")"
        
        .Cells(i, "E").Formula = "=C" & i & "/F" & i '除算があるので最後に行う
        
    i = i + 1
    
    Loop
    
    Dim EndRow As Integer
    EndRow = .Range("C3").End(xlDown).Row
    
    .Range("C1").Formula = "=SUM(C3:C" & EndRow & ")"
    .Range("H1").Formula = "=SUM(H3:H" & EndRow & ")"

End With

Exit Sub

ErrorMes:

MsgBox "処理を中止しました。" & vbLf & "MeisaiシートにProduct CodeとDescriptionがありません。"

End Sub

Private Sub 罫線を引く()
'商品別集計のシートに罫線を引きます

ResizeRow = ItemTotalSheet.Range("A2").End(xlDown).Row - 1

With ItemTotalSheet.Range("A2").Resize(ResizeRow, 9).Borders
        
        .LineStyle = xlContinuous
        .Weight = xlThin
    
End With

End Sub

Private Sub 商品別集計を新規ファイルへコピー()

ItemTotalSheet.Activate

'式を値に直す
'CD FG列の3行目から最終行までを値のみにします

'範囲を選択して
Dim RngCd As Range, RngFg As Range, Rng As Range

Set RngCd = Range("C3:D3").Resize(Range("A1").End(xlDown).Row - 2, 2)
Set RngFg = Range("F3:G3").Resize(Range("A1").End(xlDown).Row - 2, 2)

Set Rng = Union(RngCd, RngFg)

'式を削除して値のみとする場合は、Valueを格納し直すだけでいい
'http://www.relief.jp/itnote/archives/003686.php

For Each c In Rng
    
    c.Value = c.Value

Next

'コードを6ケタに修正
Dim RngCode As Range
Set RngCode = Range("B3").Resize(Range("A1").End(xlDown).Row - 2, 1)

For Each c In RngCode
    
    c.NumberFormatLocal = "@"
    If Len(c.Value) = 5 Then c.Value = "0" & c.Value

Next

'新規ワークブックへ商品別集計をコピー
Dim FileName As String, FolderPath As String, Path As String

FolderPath = Environ("USERPROFILE") & "\Documents\"
FileName = "ヤフー月次" & Format(DateAdd("M", -1, Date), "yy年M月") & ".xlsx"

Path = FolderPath & FileName

Sheets("商品別集計").Copy

ActiveWorkbook.SaveAs Path

ThisWorkbook.Close SaveChanges:=False

End Sub

Private Sub 商品別集計を新規シートへコピー()

ItemTotalSheet.Activate
ItemTotalSheet.Copy After:=Worksheets("商品別集計")

ActiveSheet.Name = "原価入力"

'式を値に直す
'CD FG列の3行目から最終行までを値のみにします

'範囲を選択して
Dim RngCd As Range, RngFg As Range, Rng As Range

Set RngCd = Range("C3:D3").Resize(Range("A1").End(xlDown).Row - 2, 2)
Set RngFg = Range("F3:G3").Resize(Range("A1").End(xlDown).Row - 2, 2)

Set Rng = Union(RngCd, RngFg)

'式を削除して値のみとする場合は、Valueを格納し直すだけでいい
'http://www.relief.jp/itnote/archives/003686.php

For Each c In Rng
    
    c.Value = c.Value

Next

'コードを6ケタに修正
Dim RngCode As Range
Set RngCode = Range("B3").Resize(Range("A1").End(xlDown).Row - 2, 1)

For Each c In RngCode
    
    c.NumberFormatLocal = "@"
    If Len(c.Value) = 5 Then c.Value = "0" & c.Value

Next

ActiveSheet.Shapes(1).Delete

End Sub
