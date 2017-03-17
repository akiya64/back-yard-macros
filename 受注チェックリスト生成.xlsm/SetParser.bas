Attribute VB_Name = "SetParser"
Option Explicit

'モール登録コードのうちセット商品について、内容商品の6ケタ・JAN・数量へ分解する。

'Public Sub ParseTiedItems(CodeCell As Range)
'行の挿入を行うため、引数はRangeとする
'数量セルはコードのセルからのオフセットで取得する


'Public Sub ParseMultipleSet(CodeCell As Range)
'セルに入っている数量の書き換えを伴うため、引数はRange
'数量セルはコードのセルからのオフセットで取得する

'定数 セット商品リストのパス
Const TIED_ITEM_LIST_BOOK As String = "ｾｯﾄ商品ﾘｽﾄ.xls"
Const LIST_BOOK_FOLDER As String = "\\server02\商品部\ネット販売関連\"

Sub セット分解()

Worksheets("作業シート").Activate

Dim CodeRange As Range
Set CodeRange = Range(Cells(2, 2), Cells(Cells.SpecialCells(xlCellTypeLastCell).Row, 2))

Dim c As Range
For Each c In CodeRange

    Dim Code As String
    Code = c.Value
    
    '77777始まりセットコードなら
    If Code Like "77777*" Then
        
        Call ParseTiedItem(c)
    
    '-02 -04 -120 ハイフン-数量 セットなら ハイフンを含みアルファベット始まりでない
    ElseIf InStr(Code, "-") > 1 And Not Code Like "[a-zA-Z]*" Then
        
        Call ParseMultipleSet(c)
    
    End If

Next

End Sub


Private Sub ParseTiedItem(CodeCell As Range)

Dim ComponentItems As Collection
Set ComponentItems = GetComponentItems(CodeCell.Value)

Dim i As Long, v As Variant

For i = 1 To ComponentItems.Count
    
    Set v = ComponentItems.Item(i)
    
    Rows(CodeCell.Offset(1, 0).Row).Insert (xlShiftDown)
    
    '6ケタあれば6ケタ、なければJAN
    CodeCell.Offset(1, 0).NumberFormatLocal = "@"
    
    If v.SyokonCode <> "" Then

        CodeCell.Offset(1, 0).Value = v.SyokonCode
    
    Else
    
        CodeCell.Offset(1, 0).Value = v.Jan
    
    End If
    
    '商品名出力と必要数量をかける
    CodeCell.Offset(1, 2).Value = v.Name
    CodeCell.Offset(1, 3).Value = v.Quantity * CodeCell.Offset(0, 3).Value
    
    '挿入後の行に注文番号を入れる
    CodeCell.Offset(1, -1).Value = CodeCell.Offset(0, -1).Value
    
    '挿入後の行にE列以降の注文情報を入れる
    CodeCell.Offset(1, 4).Resize(1, 12).Value = CodeCell.Offset(0, 4).Resize(1, 12).Value
    
Next

End Sub

Private Sub ParseMultipleSet(CodeCell As Range)
'012345-02など、ハイフン 数字のセットを分解します。

'コード文字列をハイフンの位置で分解
Dim Code As String
Code = CodeCell.Value

Dim SeparatedCode As Variant
SeparatedCode = Split(Code, "-")

'単体コードを一旦格納
Dim ComponentCode As String
ComponentCode = SeparatedCode(0)

'IsNumericメソッドで、ハイフンの後ろで数値に変換可能な値があるかチェック
'変換可能なら一個目の数字に基づいて○個セットと見なす
Dim i As Long

For i = 1 To UBound(SeparatedCode)

    If IsNumeric(SeparatedCode(i)) Then
        
        'セット数量を格納
        Dim MultipleRatio As Long
        MultipleRatio = SeparatedCode(i)
        
        Exit For
    
    End If

Next

'分解後コード・数量を出力

With CodeCell
    .NumberFormatLocal = "@"
    .Value = ComponentCode
End With

'○個組の数字が乗算できない値だと0なので、受注数量をそのまま入れる。
'出荷数量はセット数量×受注数量

If MultipleRatio > 0 Then

    CodeCell.Offset(0, 3).Value = CodeCell.Offset(0, 3) * MultipleRatio

Else
    
    CodeCell.Offset(0, 3).Value = CodeCell.Offset(0, 3).Value
    
End If

End Sub

Private Function SearchTiedItemSheet(Code As String) As Worksheet
    '「セット品リスト」エクセルファイルから、該当コードのあるワークシートを探します。
    'ワークシート毎にCountIfで該当コードがあるかチェック、あればそのワークシートを返す。
    
    Call OpenListBook
    
    Dim Hits As Long
    Dim i As Long
    
    For i = 1 To Workbooks(TIED_ITEM_LIST_BOOK).Worksheets.Count
        
        Dim LastRow As Long
        LastRow = Workbooks(TIED_ITEM_LIST_BOOK).Worksheets(i).Cells(2, 1).SpecialCells(xlCellTypeLastCell).Row
        
        Dim TiedItemCodeList As Range
        Set TiedItemCodeList = Workbooks(TIED_ITEM_LIST_BOOK).Worksheets(i).Range("A1:A" & LastRow)
        
        Hits = WorksheetFunction.CountIf(TiedItemCodeList, Code)

        If Hits > 0 Then
            
            Set SearchTiedItemSheet = Workbooks(TIED_ITEM_LIST_BOOK).Worksheets(i)
            Exit Function
        
        End If
    
    Next

End Function

Private Function GetComponentItems(TiedCode As String) As Collection

'コードから、セット内容の

'セット商品コードのあるシートを探す
Dim HitSheet As Worksheet
Set HitSheet = SearchTiedItemSheet(TiedCode)

'登録コードのレンジ、ここをMatch関数で調べて、Codeの行番号を出す
Dim CodeRange As Range
Set CodeRange = HitSheet.Range("A1:A" & HitSheet.Cells(2, 1).SpecialCells(xlCellTypeLastCell).Row)

Dim HitRow As Double
HitRow = WorksheetFunction.Match(TiedCode, CodeRange, 0)

'ループ内で使う変数など初期化
'セット内容商品を格納するコレクションを初期化。
'列カウンタ、配列数カウンタ、セット商品の商品情報を格納する配列の初期化

Dim ComponetItems As Collection
Set ComponetItems = New Collection

Dim i As Integer
i = HitSheet.Rows(1).Find("商品情報1").Column

'E列=5から、セット内容はスタート
'ヘッダー  SKU(連番77777始まり)/売価(税込)/JAN単位の総数量(●点ｾｯﾄ)     /JAN /商魂SKU /数量 / 商品名

'IsEmptyだと空白セル拾う場合がある
Do Until HitSheet.Cells(HitRow, i) = ""

    Dim UnitCell As Range, Unit As ComponentItem
    Set Unit = New ComponentItem
    
    Set UnitCell = HitSheet.Cells(HitRow, i)
            
    With Unit
        
        .Jan = UnitCell.Value
        .SyokonCode = UnitCell.Offset(0, 1).Value
        .Name = UnitCell.Offset(0, 3).Value
        .Quantity = CLng(UnitCell.Offset(0, 2).Value)
    
    End With
    
    ComponetItems.Add Unit
    
    i = i + 4

Loop

Set GetComponentItems = ComponetItems

End Function

Private Sub OpenListBook()

'セットリストのエクセルファイルを開きます。1分後に自動で閉じます。

Dim wb As Workbook

For Each wb In Workbooks
    
    If wb.Name = TIED_ITEM_LIST_BOOK Then
        
        GoTo ret
    
    End If

Next

Set wb = Workbooks.Open(LIST_BOOK_FOLDER & TIED_ITEM_LIST_BOOK, ReadOnly:=True)

'開いたワークブックがアクティブになるので、このブックをアクティブ化し直す。
ThisWorkbook.Activate

ret:
'retラベル以下は毎回実行されます。

Application.OnTime Now + TimeSerial(0, 1, 0), "CloseDataBook"

End Sub

Private Sub CloseDataBook()

Dim wb As Workbook

For Each wb In Workbooks
    
    If wb.Name = TIED_ITEM_LIST_BOOK Then
        
        wb.Close SaveChanges:=False
    
    End If

Next

End Sub
