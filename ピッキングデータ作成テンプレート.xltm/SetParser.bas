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

Private Sub InsertComponetRow(c As Range, d As Collection)

Dim i As Long, v As Variant

For i = 1 To d.Count
    
    Rows(c.Offset(1, 0).Row).Insert (xlShiftDown)
    
    Set v = d.Item(i)
    
    '6ケタあれば6ケタ、なければJAN
    c.Offset(1, 0).NumberFormatLocal = "@"
    c.Offset(1, 0).Value = c.Value
    
    If v.Code <> "" Then

        c.Offset(1, 7).Value = v.Code
    
    Else
    
        c.Offset(1, 7).Value = v.Jan
    
    End If
    
    '商品名書き込み
    c.Offset(1, 1).Value = v.Name
    
    '単体商品コードと必要数量書き込み
    c.Offset(1, 3).Value = v.Quantity * c.Offset(0, 3).Value
    
    c.Offset(1, 7).Value = v.Code
    c.Offset(1, 8).Value = v.Quantity * c.Offset(0, 3).Value
    
    '1個目のアイテムにのみ販売価格を付け替える
    '売価転記済フラグ
    Dim Flg As Boolean
    
    If v.Quantity = 1 And Flg = False Then
    
        c.Offset(1, 2).Value = c.Offset(0, 2).Value
        c.Offset(0, 2).Value = 0
                
        Flg = True
        
    Else
        c.Offset(1, 2).Value = 0
        
    End If
    
    '挿入後の行に、受注時コードはセットの7777コードを入れる
    c.Offset(1, 0).Value = c.Value
    
    '同じく、挿入後の行に注文番号を入れる
    c.Offset(1, -1).Value = c.Offset(0, -1).Value
    
    c.Offset(1, 4).Value = c.Offset(0, 4).Value
    c.Offset(1, 5).Value = c.Offset(0, 5).Value
    c.Offset(1, 6).Value = c.Offset(0, 6).Value
    
Next

End Sub

Private Function GetComponentItems(TiedCode As String) As Collection
'渡されたコードから、セット内容Collectionを返します。
'セット商品リストは呼び出し側のプロシージャで開いているものとします。

'セット商品リストのブックを取得する

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

Function CloseSetMasterBook(Optional ByVal Arg As Variant) As Boolean

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

