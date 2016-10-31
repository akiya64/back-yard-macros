Attribute VB_Name = "SetParser"
Option Explicit

Const TIED_ITEM_LIST_BOOK As String = "ｾｯﾄ商品ﾘｽﾄ.xls"
Const LIST_BOOK_FOLDER As String = "\\server02\商品部\ネット販売関連\"
Sub ParseItems(r As Range)

Call OpenListBook

ThisWorkbook.Activate

Dim HitSheet As Worksheet
Set HitSheet = SetParser.SearchTiedItemSheet(r.Value)

Dim ComponentItems As Collection
Set ComponentItems = GetComponentItems(r.Value, HitSheet)

'セット内容書き出し処理
Call InsertComponetRow(r, ComponentItems)

End Sub

Private Sub InsertComponetRow(c As Range, d As Collection)

Dim i As Long, v As Variant

For i = 1 To d.Count
    
    Set v = d.Item(i)
    
    Rows(c.Offset(1, 0).Row).Insert (xlShiftDown)
    
    '6ケタあれば6ケタ、なければJAN
    c.Offset(1, 0).NumberFormatLocal = "@"
    
    If v.Code <> "" Then

        c.Offset(1, 0).Value = v.Code
    
    Else
    
        c.Offset(1, 0).Value = v.Jan
    
    End If
    
    '商品名出力と必要数量をかける
    c.Offset(1, 1).Value = v.Name
    c.Offset(1, 2).Formula = "=" & v.Quantity & "*" & c.Offset(0, 2).Value
    c.Offset(1, 2).Value = c.Offset(1, 2).Value
    
    '1個目のアイテムにのみ販売価格を付け替える
    
    '売価転記済フラグ
    Dim Flg As Boolean
    
    If v.Quantity = 1 And Flg = False Then
    
        c.Offset(1, 3).Value = c.Offset(0, 3).Value
        c.Offset(0, 3).Value = 0
                
        Flg = True
        
    Else
        c.Offset(1, 3).Value = 0
        
    End If
    
    '挿入後の行に、ヤフー登録コードはセットの7777コードを入れる
    c.Offset(1, -1).Value = c.Value
    
    '同じく、挿入後の行に注文番号を入れる
    c.Offset(1, -3).Value = c.Offset(0, -3).Value
    
Next

End Sub

Private Function SearchTiedItemSheet(Code As String) As Worksheet
    '該当コードのあるワークシートを探します。
      
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

Private Function GetComponentItems(TiedCode As String, TiedCodeList As Worksheet) As Collection

'渡されたシートとコードから、セット内容Collectionを返します。
'呼び出し側でエラーハンドリングを行うので、On Errorステートメントは不要

'登録コードのレンジ、ここをMatch関数で調べて、Codeの行番号を出す
Dim CodeRange As Range
Set CodeRange = TiedCodeList.Range("A1:A" & TiedCodeList.Cells(2, 1).SpecialCells(xlCellTypeLastCell).Row)

Dim HitRow As Double
HitRow = WorksheetFunction.Match(TiedCode, CodeRange, 0)


Dim ComponetItems As Collection
Set ComponetItems = New Collection

'E列=5から、セット内容はスタート
'ヘッダー  SKU(連番77777始まり)/売価(税込)/JAN単位の総数量(●点ｾｯﾄ)     /JAN /商魂SKU /数量 / 商品名

'列カウンタ
Dim i As Integer
i = TiedCodeList.Rows(1).Find("商品情報1").Column

'IsEmptyだと空白セル拾う場合がある
Do Until TiedCodeList.Cells(HitRow, i) = ""

    Dim UnitCell As Range
    
    Dim Unit As ComponentItem
    Set Unit = New ComponentItem
    
    Set UnitCell = TiedCodeList.Cells(HitRow, i)
    
    With Unit
        
        .Jan = UnitCell.Value
        .Code = UnitCell.Offset(0, 1).Value
        .Name = UnitCell.Offset(0, 3).Value
        .Quantity = CLng(UnitCell.Offset(0, 2).Value)
    
    End With
        
    ComponetItems.Add Unit
    
    i = i + 4

Loop

Set GetComponentItems = ComponetItems

End Function

Private Sub OpenListBook()

'セットリストのエクセルファイルを開きます。

Dim wb As Workbook

For Each wb In Workbooks
    
    If wb.Name = TIED_ITEM_LIST_BOOK Then
        
        GoTo ret
    
    End If

Next

Set wb = Workbooks.Open(LIST_BOOK_FOLDER & TIED_ITEM_LIST_BOOK, ReadOnly:=True)

ret:

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

Sub ParseScalingSet(r As Variant)

Dim Code As String
Code = r.Value

Dim SeparatedCode As Variant
SeparatedCode = Split(Code, "-", 2)

'IsNumericメソッドで、ハイフンの後ろが数値に変換可能かチェック
'変換可能なら、○個セットと見なす

If Not IsNumeric(SeparatedCode(1)) Then
    Exit Sub
End If

'セットなら、D列は単体コード、数量はセット数量×受注数量
r.NumberFormatLocal = "@"
r.Value = CStr(SeparatedCode(0))

Range("F" & r.Row).Value = Range("F" & r.Row).Value * CLng(Val(SeparatedCode(1)))

'備考に、セット分解済記入
Range("K" & r.Row).Value = Range("K" & r.Row).Value & "セット分解 済"

End Sub

