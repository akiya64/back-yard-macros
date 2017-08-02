Attribute VB_Name = "Module1"
Option Explicit
'http://officetanaka.net/excel/vba/tips/tips36.htm

Sub ピッキングファイル集計()
'日付を1日足しながらファイルを探してテンポラリーシートへコピーして集計シートへ集計結果を入れていく

    Application.ScreenUpdating = False

    Dim DateCount As Date
    
    For DateCount = #1/1/2012# To #10/30/2016#
        
        TmpSheet.Cells.Clear
        
        Call CopyPickingData(DateCount)
        
        If TmpSheet.Range("A1") <> "" Then
            
            '日付を入れて、作業シートを集計
            ResultSheet.Range("A1").End(xlDown).Offset(1, 0) = DateCount
            Call AggregatePicking

        End If
        
    Next
    
End Sub

Private Sub CopyPickingData(ByVal TargetDay As Date)
    
    Dim FSO As New FileSystemObject, Folder As Variant, File As File
    
    Dim Path As String
    Path = "C:\Users\mos9\Documents\ピッキング過去データ\" & Year(TargetDay) & "\" & Format(TargetDay, "M月")
    
    If FSO.FolderExists(Path) = False Then Exit Sub
    
    For Each File In FSO.GetFolder(Path).Files
        
        If File.Name Like "*" & Format(TargetDay, "MMdd") & ".xls*" _
            And Not File.Name Like "*棚*" Then
        
            Call CopySheet(File.Path)
            
        End If
        
    Next File
    
End Sub

Private Sub CopySheet(ByVal Path As String)
    Application.ScreenUpdating = False
            On Error GoTo ee:
    Workbooks.Open Filename:=Path
    On Error GoTo 0
    
    Dim DestBaseCell As Range
    
    If TmpSheet.Range("A1").Value = "" Then
        Set DestBaseCell = TmpSheet.Range("A1")
    Else
        Set DestBaseCell = TmpSheet.Range("A1").End(xlDown).Offset(1, 0)
    End If
        
    With ActiveSheet
        '開いたピッキングデータブックから、SKU列・数量列・ロケーション列のみコピー
        
        Dim Header As Range, TargetRange As Range, headerArray As Variant
        If Dir(Path) Like "*ヤフー*" Then
            headerArray = Array("商品コード", "数量")
        Else
            headerArray = Array("SKU", "個数")
        End If
        
        Set Header = .Range("A1:AA2").Find(headerArray(0))
        Set TargetRange = .Range(Header.Offset(1, 0), Header.End(xlDown))
        TargetRange.Copy Destination:=DestBaseCell.Offset(0, 0)
    
        Set Header = .Range("A1:AA2").Find(headerArray(1))
        Set TargetRange = .Range(Header.Offset(1, 0), Header.End(xlDown))
        TargetRange.Copy Destination:=DestBaseCell.Offset(0, 1)

        Set Header = .Range("A1:AA2").Find("ロケーション")
        On Error GoTo e:
        Set TargetRange = .Range(Header.Offset(1, 0), Header.End(xlDown))
        TargetRange.Copy Destination:=DestBaseCell.Offset(0, 2)
    End With
e:
    Workbooks(Dir(Path)).Close SaveChanges:=False
ee:
End Sub

Private Sub AggregatePicking()

With TmpSheet

    .Activate

    '商品コード文字列化
    Dim r As Range
    For Each r In .Range(.Cells(1, 1), .Cells(1, 1).End(xlDown))
        r.NumberFormatLocal = "@"
        r.Value = CStr(r.Value)
    Next

    '算出結果を格納する変数
    Dim OrderCount As Long, OrderQuantity As Long, OrderedItemCount As Long, RegisterdItemCount As Long, RegularItemCount As Long
    
    OrderCount = .Range("A1").CurrentRegion.Rows.Count
    OrderQuantity = WorksheetFunction.Sum(.Range(.Cells(1, 2), .Cells(1, 2).End(xlDown)))
    

    '重複コードを削除
    .Range(Cells(1, 1), Cells(1, 1).End(xlDown).Offset(0, 2)).RemoveDuplicates Columns:=1, Header:=xlNo
    
    OrderedItemCount = .Range("A1").CurrentRegion.Rows.Count

    Dim CodeRange As Range
    Set CodeRange = .Range(Cells(1, 1), Cells(1, 1).End(xlDown))

    RegisterdItemCount = WorksheetFunction.CountIf(CodeRange, "0?????") + WorksheetFunction.CountIf(CodeRange, "5?????")

    RegularItemCount = WorksheetFunction.CountIf(CodeRange.Offset(0, 2), "")

End With

ResultSheet.Activate
ResultSheet.Range("A1").End(xlDown).Offset(0, 1).Resize(1, 5).Value = Array(OrderCount, OrderQuantity, OrderedItemCount, RegisterdItemCount, RegularItemCount)

End Sub
