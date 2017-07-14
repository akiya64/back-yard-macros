Attribute VB_Name = "Module1"
Option Explicit
'http://officetanaka.net/excel/vba/tips/tips36.htm

Sub ピッキングファイル集計()
'日付を1日足しながらファイルを探してテンポラリーシートへコピーして集計シートへ集計結果を入れていく

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
    Path = "D:\Doc\ピッキング過去データ\" & Year(TargetDay) & "\" & Format(TargetDay, "M月")
    
    If FSO.FolderExists(Path) = False Then Exit Sub
    
    For Each File In FSO.GetFolder(Path).Files
        
        If File.Name Like "*" & Format(TargetDay, "MMdd") & ".xls*" Then
        
            Call CopySheet(File.Path)
            
        End If
        
    Next File
    
End Sub

Private Sub CopySheet(ByVal Path As String)
    
    Workbooks.Open Filename:=Path
    
    Dim DestBaseCell As Range
    
    If TmpSheet.Range("A1").Value = "" Then
        Set DestBaseCell = TmpSheet.Range("A1")
    Else
        Set DestBaseCell = TmpSheet.Range("A1").End(xlDown).Offset(1, 0)
    End If
        
    With ActiveSheet
        '開いたピッキングデータブックから、SKU列・数量列・ロケーション列のみコピー
        Dim Header As Range, TargetRange As Range
            
        Set Header = .Range("A1:AA2").Find("SKU")
        Set TargetRange = .Range(Header.Offset(1, 0), Header.End(xlDown))
        TargetRange.Copy Destination:=DestBaseCell.Offset(0, 0)
    
        Set Header = .Range("A1:AA2").Find("個数")
        Set TargetRange = .Range(Header.Offset(1, 0), Header.End(xlDown))
        TargetRange.Copy Destination:=DestBaseCell.Offset(0, 1)

        Set Header = .Range("A1:AA2").Find("ロケーション")
        Set TargetRange = .Range(Header.Offset(1, 0), Header.End(xlDown))
        TargetRange.Copy Destination:=DestBaseCell.Offset(0, 2)
    End With
    
    ActiveWorkbook.Close SaveChanges:=False

End Sub

Private Sub AggregatePicking()

With TmpSheet

    Dim OrderCount As Long, OrderQuantity As Long, OrderedItemCount As Long, RegisterdItemCount As Long, RegularItemCount As Long
    
    OrderCount = .Range("A1").CurrentRegion.Rows.Count
    OrderQuantity = WorksheetFunction.Sum(.Range(.Cells(1, 2), .Cells(1, 2).End(xlDown)))

    '重複コードを削除
    'OrderedItemCount=

    'RegisterdItemCount=

    'RegularItemCount=

End With

ResultSheet.Range("A1").End(xlDown).Offset(0, 1).Resize(1, 2).Value = Array(OrderCount, OrderQuantity)

End Sub
