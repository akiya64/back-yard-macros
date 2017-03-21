Attribute VB_Name = "SetParser"
Option Explicit

Const TIED_ITEM_LIST_BOOK As String = "ｾｯﾄ商品ﾘｽﾄ.xls"
Const LIST_BOOK_FOLDER As String = "\\server02\商品部\ネット販売関連\"
Sub ParseItems(r As Range)

'セット商品リストのブックを開く
Call OpenListBook

ThisWorkbook.Activate

Dim ComponentItems As Collection
Set ComponentItems = GetComponentItems(r.Value)

If ComponentItems Is Nothing Then Exit Sub

'セット内容書き出し処理
Call InsertComponetRow(r, ComponentItems)

End Sub

Private Sub InsertComponetRow(c As Range, ComponentItems As Collection)

Dim i As Long
For i = 1 To ComponentItems.Count
    
    Dim Record As Range
    Set Record = Range(Cells(c.Row, 1), Cells(c.Row, 15))
    
    '一旦セット品の行をコピー
    Record.Copy
    Record.Offset(1, 0).Insert (xlShiftDown)
    
    '挿入した行を示す行番号
    Dim wr As Long
    wr = c.Row + 1
    
    '挿入後の行をセット内容の商品情報で書き換える
    Dim Component As Variant
    Set Component = ComponentItems.Item(i)
    
    'アドイン用のコードは6ケタあれば6ケタ、なければJANで上書き
    If Component.Code <> "" Then

        Cells(wr, 9).Value = Component.Code
    
    Else
    
        Cells(wr, 9).Value = Component.Jan
    
    End If
    
    '商品名上書き
    Cells(wr, 3).Value = Component.Name
    
    '数量と必要数量上書き
    Cells(wr, 4).Value = Component.Quantity * Cells(c.Row, 4).Value
    Cells(wr, 10).Value = Component.Quantity * Cells(c.Row, 4).Value
    
    '1個目のアイテムにのみ販売価格を付け替える
    '売価転記済フラグ
    Dim Flg As Boolean
    
    If Component.Quantity = 1 And Flg = False Then
    
        Cells(wr, 5) = Cells(c.Row, 5).Value
        Cells(c.Row, 5).Value = 0
                
        Flg = True
        
    Else
        Cells(c.Row, 5).Value = 0
        
    End If
    
Next

End Sub

Private Function GetComponentItems(TiedCode As String) As Collection
'渡されたコードから、セット内容Collectionを返します。
'セット商品リストは呼び出し側のプロシージャで開いているものとします。

'セット商品リストから該当コードのあるシートと行を探す

Dim i As Long
For i = 1 To Workbooks(TIED_ITEM_LIST_BOOK).Worksheets.Count
        
    Dim TiedCodeList As Worksheet
    Set TiedCodeList = Workbooks(TIED_ITEM_LIST_BOOK).Worksheets(i)

    Dim CodeRange As Range
    Set CodeRange = TiedCodeList.Range("A1:A" & TiedCodeList.Cells(2, 1).SpecialCells(xlCellTypeLastCell).Row)
        
    On Error Resume Next
        
        Dim HitRow As Double
        HitRow = WorksheetFunction.Match(TiedCode, CodeRange, 0)
        
        If HitRow > 0 Then Exit For
        
    On Error GoTo 0

Next

If HitRow = 0 Then
    Exit Function
End If

Dim ComponetItems As Collection
Set ComponetItems = New Collection

'E列=5から、セット内容はスタート
'ヘッダー  SKU(連番77777始まり)/売価(税込)/JAN単位の総数量(●点ｾｯﾄ)     /JAN /商魂SKU /数量 / 商品名

'列カウンタ
Dim k As Integer
k = TiedCodeList.Rows(1).Find("商品情報1").Column

'IsEmptyだと空白セル拾う場合がある
Do Until TiedCodeList.Cells(HitRow, k) = ""

    Dim UnitCell As Range
    
    Dim Unit As ComponentItem
    Set Unit = New ComponentItem
    
    Set UnitCell = TiedCodeList.Cells(HitRow, k)
    
    With Unit
        
        .Jan = UnitCell.Value
        .Code = UnitCell.Offset(0, 1).Value
        .Name = UnitCell.Offset(0, 3).Value
        .Quantity = CLng(UnitCell.Offset(0, 2).Value)
    
    End With
        
    ComponetItems.Add Unit
    
    k = k + 4

Loop

Set GetComponentItems = ComponetItems

End Function

Private Sub OpenListBook()

'セットリストのエクセルファイルを開くか、開いていればそのまま終了します。
'1つのピッキングシートの処理で何回か開く場合があるので、閉じるのは呼び出し側でセット分解終了のタイミングで行います。

Dim wb As Workbook

For Each wb In Workbooks
    
    If wb.Name = TIED_ITEM_LIST_BOOK Then
        
        GoTo ret
    
    End If

Next

Set wb = Workbooks.Open(LIST_BOOK_FOLDER & TIED_ITEM_LIST_BOOK, ReadOnly:=True)

ret:

End Sub

Function CloseSetMasterBook(Optional ByVal arg As Variant) As Boolean

Dim wb As Workbook

For Each wb In Workbooks
    
    If wb.Name = TIED_ITEM_LIST_BOOK Then
        
        wb.Close SaveChanges:=False
    
    End If

Next

End Function

Sub ParseScalingSet(r As Variant)

Dim Code As String, FixedCode As String
Code = r.Value

Dim SeparatedCode As Variant
SeparatedCode = Split(Code, "-", 2)

If SeparatedCode(0) Like String(5, "#") Then
    FixedCode = "0" & SeparatedCode(0)
Else
    FixedCode = SeparatedCode(0)
End If

'単体コードをI列に入れる
Range("I" & r.Row).NumberFormatLocal = "@"
Range("I" & r.Row).Value = FixedCode

'IsNumericメソッドで、ハイフンの後ろが数値に変換可能かチェック
'変換可能なら、○個セットと見なす

If Not IsNumeric(SeparatedCode(1)) Then
    Exit Sub
End If

'セットなら、必要数量はセット数量×受注数量に書き換え
Range("J" & r.Row).Value = Range("J" & r.Row).Value * CLng(Val(SeparatedCode(1)))

End Sub
