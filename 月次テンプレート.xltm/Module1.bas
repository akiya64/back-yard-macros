Attribute VB_Name = "Module1"
Sub 読込集計_フルオート()
    'MeisaiSheet/paymentMethodを読み込んで式も入れて、新規シートへ保存して終了
    '計3個のファイルが必要ですので、マクロ起動後のファイル選択ウィンドウで指定して下さい。
    'CSV:MeisaiSheet,PaymentMethod Xlsx:価格チェック○月.xlsx
    
    Dim MonthName As String
    MonthName = Format(DateAdd("M", -1, Date), "yy年M月")
    
    Sheets("商品別集計").Range("A1") = MonthName & " ヤフー月次"
    
    Call PaymentCsv読込
    Call MeisaiSheetCsv読込
    
    Call 転記と重複削除
    Call 集計式の挿入
    Call 罫線を引く
    
    Call 原価情報の転記
    
    Dim FileName As String
    FileName = "ヤフー月次" & MonthName & "_作業中.xlsm"
    
    Dim Folder As String
    Folder = Environ("USERPROFILE") & "\Documents\"
    
    Dim Path As String
    Path = Folder & FileName
    
    ThisWorkbook.SaveAs FileName:=Path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    
    Call 商品別集計を新規ファイルへコピー
    
End Sub

Private Function findColum(str As String) As Integer

findColum = WorksheetFunction.Match(str, MeisaiSheet.Range("A1").Resize(1, 20), 0)

End Function

Sub MeisaiSheetCsv読込()

'Meisai.csvファイルを指定
'1行ずつ配列に入れて、シートへ転記
'Quantity=0はキャンセルなので弾きます

'-----ここから-----------------'
Dim FilePath
FilePath = setCsvPath("Meisai")

If FilePath = "" Then
    MsgBox "ファイル指定がキャンセルされました。"
    Exit Sub
End If

'CSV読込のTextStreamを準備
Dim LineBuf As Variant
Dim FSO As Object

Set FSO = CreateObject("Scripting.FileSystemObject")

Dim TS As Textstream
Set TS = FSO.OpenTextFile(FilePath, ForReading)
    
'-------ここまでPaymentMethodでも同じ処理をしていますCSV名が違うだけ----------"
    
Dim Header As Variant
Header = Split(TS.ReadLine, """,""")

'1項目目の"と、最後の項目の"が残るので削除します、chr(34)で"です
Header(0) = (Replace(Header(0), Chr(34), ""))
Header(UBound(Header)) = (Replace(Header(UBound(Header)), Chr(34), ""))

'キャンセルの注文を集計しないために、個数の列は何番目か特定する
For j = 0 To UBound(Header)
    
    Dim QtyCol As Integer
    
    If Header(j) = "Quantity" Then
        QtyCol = j
        Exit For
    End If

Next

'ヘッダーをシートに転記
Sheets("Meisai").Range("A1").Resize(1, UBound(Header) + 1).Value = Header

Dim i As Long
i = 1

'続いてMeisaiSheetデータをシートへ転記
Do Until TS.AtEndOfStream
    
    'LineBuf配列に1項目ずつ入れる
    LineBuf = Split(TS.ReadLine, """,""")
        
    For j = 0 To UBound(LineBuf)
        LineBuf(j) = Trim(Replace(LineBuf(j), Chr(34), "")) '念のため再度chr(34)で " [半角二重引用符]を除去してトリム
        
        If j = QtyCol Then  'qty=0ならキャンセルの注文なので、さっさとContinueへ飛ぶ
            If LineBuf(j) = 0 Then GoTo Continue
        
        End If
    
    Next
    
    'A1セルから、オフセット＋リサイズしつつ転記
    Sheets("Meisai").Range("A1").Offset(i, 0).Resize(1, UBound(LineBuf) + 1).Value = LineBuf
    
    i = i + 1

Continue:

Loop

' 指定ファイルをCLOSE
TS.Close

End Sub

Sub PaymentCsv読込()

'Csvを指定する
Dim FilePath
FilePath = setCsvPath("PaymentMethod")

If FilePath = "" Then
    MsgBox "ファイル指定がキャンセルされました。"
    Exit Sub
End If

'CSV読込用のTextStreamオブジェクトを用意
Dim LineBuf As Variant
Dim FSO As Object

Set FSO = CreateObject("Scripting.FileSystemObject")

Dim TS As Textstream
Set TS = FSO.OpenTextFile(FilePath, ForReading)
    
Dim Header As Variant
Header = Split(TS.ReadLine, """,""")

'1項目目の"と、最後の項目の"が残るので削除します、chr(34)で"です
Header(0) = (Replace(Header(0), Chr(34), ""))
Header(UBound(Header)) = (Replace(Header(UBound(Header)), Chr(34), ""))

'ヘッダーをシートに転記
Sheets("PaymentMethod").Range("A1").Resize(1, UBound(Header) + 1).Value = Header

Dim i As Long
i = 1

'続いてPaymentMethodのレコードをシートへ転記
Do Until TS.AtEndOfStream
    
    'LineBuf配列に1項目ずつ入れる
    LineBuf = Split(TS.ReadLine, """,""")
        
    For j = 0 To UBound(LineBuf)
        LineBuf(j) = Trim(Replace(LineBuf(j), Chr(34), "")) '念のため再度chr(34)で " [半角二重引用符]を除去してトリム
        
        If j = SaleTotalCol Then  'SaleTotalCol=0ならキャンセルの注文なので、さっさとContinueへ飛ぶ
            If LineBuf(j) = 0 Then GoTo Continue
        
        End If
    
    Next
    
    'A1セルから、オフセット＋リサイズしつつ転記
    Sheets("PaymentMethod").Range("A1").Offset(i, 0).Resize(1, UBound(LineBuf) + 1).Value = LineBuf
    
    i = i + 1

Continue:

Loop

' 指定ファイルをCLOSE
TS.Close

' 読み込み後の集計 ここから別プロシージャの方がいいかも
With PaymentSheet
    
    .Activate 'アクティブでないとダメかも、キャストが
    
    Dim EndRow As Long
    EndRow = .Range("A1").End(xlDown).Row
    
    For i = 2 To EndRow
        .Cells(i, 8).NumberFormat = "0"
        .Cells(i, 8).Value = CDbl(Cells(i, 8).Value) 'SUMIFで計算するのでセルの値はダブル型

    Next

    'Totalのデータレンジ、文字列で
    Dim TotalRangeStr As String
    TotalRangeStr = .Rows(1).Find(what:="Total", LookAt:=xlWhole).Offset(1, 0).Resize(EndRow, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    'PaymentMethodのデータレンジ、文字列で
    Dim PaymentMethodRangeStr As String
    PaymentMethodRangeStr = .Rows(1).Find(what:="Payment Method", LookAt:=xlWhole).Offset(1, 0).Resize(EndRow, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    '集計用の式をセルへ格納
        
    .Range("K3").Formula = "=COUNTIF(" & PaymentMethodRangeStr & ",J16)"
    .Range("L3").Formula = "=SUMIF(" & PaymentMethodRangeStr & ",J16," & TotalRangeStr & ")"
    
    .Range("K4").Formula = "=COUNTIF(" & PaymentMethodRangeStr & ",J28)"
    .Range("L4").Formula = "=SUMIF(" & PaymentMethodRangeStr & ",J28," & TotalRangeStr & ")"
    
    .Range("K5").Formula = "=COUNTIF(" & PaymentMethodRangeStr & ",""payment_b1"")"
    .Range("L5").Formula = "=SUMIF(" & PaymentMethodRangeStr & ",""payment_b1""," & TotalRangeStr & ")"
        
    Set AllSaleTotalRange = .Rows(1).Find(what:="", LookAt:=xlPart) '検索ウィンドウの設定を戻すために空検索
        
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


Private Sub 転記と重複削除()
'売上商品の一意な表を用意します。
'MeisaiSheetのコードと商品名を集計シートに転記して重複削除します。
'RangeメソッドのRemoveDuplicatesを使う。

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
    
    For i = 2 To MeisaiSheet.UsedRange.Rows.Count
        .Cells(i + 1, 1).Value = MeisaiSheet.Cells(i, DescriptionCol)
        .Cells(i + 1, 2).Value = MeisaiSheet.Cells(i, codeCol)
    Next

    
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

Sub 罫線を引く()
'商品別集計のシートに罫線を引きます

ResizeRow = ItemTotalSheet.Range("A2").End(xlDown).Row - 1

With ItemTotalSheet.Range("A2").Resize(ResizeRow, 9)

    With .Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With .Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With .Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With .Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With .Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With .Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
End With

End Sub

Sub 原価情報の転記()

'受注チェックxlsmファイルを指定
Dim FilePath
FilePath = setCsvPath("ヤフー価格チェックを指定")

If FilePath = "" Then
    MsgBox "ファイル指定がキャンセルされました。"
    Exit Sub
End If


'価格チェックワークシートのコピー

Workbooks.Open FilePath
  
With ActiveWorkbook
    
    .Worksheets("価格チェック").Copy After:=ThisWorkbook.Worksheets("商品別集計")
    .Close

End With


'原価チェックシートの商品コードを修正＞VlookupでヒットさせるためにStr型で格納し直す。

Worksheets("価格チェック").Activate 'Withで括るとレンジ指定が面倒になるので、Activateして作業
    
Dim EndRow As Integer
EndRow = Worksheets("価格チェック").UsedRange.Rows.Count

For i = 2 To EndRow
    Cells(i, 1).NumberFormatLocal = "@"
    Cells(i, 1).Value = CStr(Cells(i, 1).Value)
Next

'Vlookupで検索する範囲を指定
Dim SearchRange As Range
Set SearchRange = Range("A1").Resize(EndRow, 5)


Dim SearchRangeAddress As String
SearchRangeAddress = "価格チェック!" & SearchRange.Address(RowAbsolute:=False, ColumnAbsolute:=False)

'Vlookup式を送り込む
ItemTotalSheet.Activate

j = 3 '行カウンタ初期化、集計シートは3行目から商品コードが始まる

Do Until IsEmpty(Cells(j, 2))
    
    Dim CodeAddress As String
    CodeAddress = Cells(j, 2).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    Cells(j, 2).Offset(0, 5).Formula = "=VLOOKUP(" & CodeAddress & "," & SearchRangeAddress & ",5,FALSE)"
    Cells(j, 2).Offset(0, 6).Formula = "=F" & j & "*" & "G" & j
    
    With Cells(j, 2).Offset(0, 7)
        .Formula = "=H" & j & "/" & "C" & j
        .NumberFormatLocal = "0.00%"
    End With
    
    j = j + 1
    
Loop

Range("H1").Formula = "=SUM(H3:H" & j - 1 & ")"

End Sub

Sub 商品別集計を新規ファイルへコピー()

ItemTotalSheet.Activate

'式を値に直す
'CD FG列の3行目から最終行までを値のみにします

'範囲を選択して
Dim RngCd As Range, RngFg As Range, Rng As Range

Set RngCd = Range("C3:D3").Resize(Range("A1").End(xlDown).Row - 2, 2)
Set RngFg = Range("F3:G3").Resize(Range("A1").End(xlDown).Row - 2, 2)

Set Rng = Union(RngCd, RngFg)

'式を削除して値のみとする場合は、Valueを格納し直すだけでいい！
'http://www.relief.jp/itnote/archives/003686.php全

For Each c In Rng
    
    c.Value = c.Value

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
