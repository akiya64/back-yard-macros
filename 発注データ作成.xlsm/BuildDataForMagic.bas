Attribute VB_Name = "BuildDataForMagic"
Option Explicit
Const OPERATOR_CODE As Integer = 329

Type Purchase

    Code As String
    ProductName As String
    
    VendorCode As Long
    VendorName As String
    
    UnitCost As Long
    
    PurchaseQuantity As Long
    RequireQuantity As Long
    
    WarehouseNumber As Integer
    
    IsPickup As Integer
    
    IsHold As Boolean
    HoldReason As String

End Type

Sub BuildPurcahseData()
    
Worksheets("手配数量決定シート").Activate

Dim i As Long
For i = 2 To Range("A1").End(xlDown).Row

    Dim CurrentPurchase As Purchase
    CurrentPurchase = ReadPurchase(i)
    
    If Not CurrentPurchase.IsHold Then
        'Call WriteHoldList(CurrentPurchase)
    
        If CurrentPurchase.Code Like "######" Then
            Call WriteMagicTxt(CurrentPurchase)
        Else
            Call WriteMagicManualInput(CurrentPurchase)
        End If
        
    End If

Next

Worksheets("Magic一括登録").Columns("A:E").AutoFit
Worksheets("Magic手入力用").Columns("A:I").AutoFit

'Magic一括登録シートを新規ブックにコピー、CSVで保存
Worksheets("Magic一括登録").Copy
ActiveSheet.Rows(1).Delete

Dim FileName As String
FileName = "\Magic登録用" & Format(Date, "MMdd") & ".txt"

If Dir(ThisWorkbook.path & FileName) <> "" Then
    FileName = Replace(FileName, Format(Date, "MMdd"), Format(Date, "MMdd") & "-" & Format(Time, "hhmm"))
End If

ActiveWorkbook.SaveAs FileName:=ThisWorkbook.path & FileName, FileFormat:=xlText
ActiveWorkbook.Close

End Sub

Private Function ReadPurchase(ByVal Row As Long) As Purchase

'手配数量決定シートから1行を1変数に読み込む
Dim TmpPur As Purchase

With TmpPur
    .Code = Cells(Row, 7).Value  '発注時の商品コード、JANか6ケタ
    .ProductName = Cells(Row, 8).Value '商品名、JAN手配分のみ必須
    
    .VendorCode = Cells(Row, 4).Value '手配先コード
    .VendorName = Cells(Row, 5).Value '手配先名称
     
    .WarehouseNumber = IIf(Cells(Row, 6).Value = "V", "187", "180")  '倉庫番号

    .RequireQuantity = Cells(Row, 9).Value '手配依頼数量

    '発注保留に該当するかチェックして、フラグを立てる

    '保留として数量を注意書きで上書きしているか？
    If IsNumeric(Cells(Row, 1).Value) Then
        TmpPur.PurchaseQuantity = Cells(Row, 1).Value
    Else
        TmpPur.HoldReason = Cells(Row, 1).Value
        TmpPur.IsHold = True
    End If
    
    .UnitCost = Cells(Row, 10).Value '原価データの有無
    If .UnitCost = 0 Then
        TmpPur.IsHold = True
        TmpPur.HoldReason = "原価不明"
    End If
    
    If .VendorCode = 0 And .VendorName = "" Then '仕入先が分かっているか
        TmpPur.IsHold = True
        TmpPur.HoldReason = "仕入先不明"
    End If
    
    .IsPickup = GetPickupFlag(.VendorCode) '引取で手配するか
        
End With

ReadPurchase = TmpPur

End Function

Private Sub WriteMagicTxt(ByRef Purchase As Purchase)
    
    Dim WriteRow As Long, TargetSheet As Worksheet, Record As Variant
    
    With Purchase
        Record = Array( _
                    .WarehouseNumber, _
                    .Code, _
                    .PurchaseQuantity, _
                    .IsPickup, _
                    OPERATOR_CODE _
                    )
    End With
    
    Set TargetSheet = Worksheets("Magic一括登録")
    WriteRow = TargetSheet.UsedRange.Rows.Count + 1
    
    With TargetSheet
        .Cells(WriteRow, 2).NumberFormatLocal = String(9, "0")
        .Cells(WriteRow, 3).NumberFormatLocal = String(8, "0")
    
        .Cells(WriteRow, 1).Resize(1, 5).Value = Record
    End With
    
End Sub

Private Sub WriteMagicManualInput(ByRef Purchase As Purchase)

    Dim WriteRow As Long, TargetSheet As Worksheet, Record As Variant
    
    With Purchase
        Record = Array( _
                    .WarehouseNumber, _
                    .VendorCode, _
                    .VendorName, _
                    .Code, _
                    .ProductName, _
                    .PurchaseQuantity, _
                    .UnitCost, _
                    .IsPickup, _
                    OPERATOR_CODE _
                    )
    End With
    
    Set TargetSheet = Worksheets("Magic手入力用")
    WriteRow = TargetSheet.UsedRange.Rows.Count + 1
    
    TargetSheet.Cells(WriteRow, 4).NumberFormatLocal = "@"
    TargetSheet.Cells(WriteRow, 1).Resize(1, 9).Value = Record
    
End Sub
