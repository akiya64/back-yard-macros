Attribute VB_Name = "Parser"
Option Explicit

Const TIED_ITEM_LIST_BOOK As String = "商品ﾘｽﾄ.xls"
Const LIST_BOOK_FOLDER As String = "\\server02\商品部\ネット販売関連\"
Sub ParseItems()

    If Selection.Columns.Count > 1 Then
        
        MsgBox "2列以上選択しないで下さい。"
        End
    
    End If

    Dim CurrentWorkBook As Workbook
    Set CurrentWorkBook = ActiveWorkbook
    
    Call OpenListBook

    CurrentWorkBook.Activate

    Dim rng As Range, r As Range, sh As Worksheet
    Set rng = Selection

    For Each r In rng
        
        If Not r.Value Like "#####*" Then GoTo Continue
        
        'セット内容を取得する 擬似的なTry-Catch
        'Try
        On Error Resume Next
            
            Dim HitSheet As Worksheet
            Set HitSheet = Parser.SearchTiedItemSheet(r.Value)
        
            Dim Items As Collection
            Set Items = GetComponentItems(r.Value, HitSheet)
            
            'セット内容書き出し処理
            
            Call InsertComponetRow(r, Items)
            
            'Catch
            If Err Then
                
                'エラー時書き戻しメソッド
                
                Err.Clear
                GoTo Continue
            
            End If
            
        On Error GoTo 0
            
        'セット品コードはリスト内にあるが、商品登録がない場合
        If Items.Count = 0 Then
                            
            'エラー時書き戻しメソッド
            
            GoTo Continue
        
        End If
     
Continue:
        
    Next

End Sub

Sub InsertComponetRow(c As Range, d As Collection)

Dim v As Variant

For Each v In d
    
    Rows(c.Offset(1, 0).Row).Insert (xlShiftDown)
    
    '社内コードあれば社内コード、なければJAN
    If v.Code <> "" Then
        
        c.Offset(1, 0).Value = v.Code
    
    Else
    
        c.Offset(1, 0).Value = v.Jan
    
    End If
    
    '商品名出力と必要数量をかける
    If TypeName(c.Offset(0, 1).Value) = "String" Then
        
        c.Offset(1, 1).Value = v.Name
        c.Offset(1, 2).Formula = "=" & v.Quantity & "*" & c.Offset(0, 2).Value
        
    Else
    
        c.Offset(1, 2).Value = v.Name
        c.Offset(1, 1).Formula = "=" & v.Quantity & "*" & c.Offset(0, 1).Value
    
    End If
    
    
Next

End Sub


Function SearchTiedItemSheet(Code As String) As Worksheet
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


Function GetComponentItems(TiedCode As String, TiedCodeList As Worksheet) As Collection

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
'ヘッダー  連番/売価(税込)/JAN単位の総数量(●点ｾｯﾄ)     /JAN /社内コード /数量 / 商品名

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

Sub OpenListBook()

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

Sub CloseDataBook()

Dim wb As Workbook

For Each wb In Workbooks
    
    If wb.Name = TIED_ITEM_LIST_BOOK Then
        
        wb.Close SaveChanges:=False
    
    End If

Next

End Sub
