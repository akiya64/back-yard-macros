Attribute VB_Name = "UpdateVendorCode"
Option Explicit

Sub 仕入先コード追記()

With Worksheets("商品情報")
    .Activate
    
    Dim EndRow As Long
    EndRow = .UsedRange.Rows.Count

End With

Dim i As Long

For i = 2 To EndRow

    Call UpdateVendorCode(i)

Continue:
Next

End Sub

Sub UpdateVendorCode(ByVal CurrentRow As Double)

Dim VenderTerm As String, Terms As Variant, v As Variant
VenderTerm = Cells(CurrentRow, 4).Value

'仕入先名をSplit
Terms = Split(VenderTerm, " ")

'分離した文字列配列、各々について仕入先シートと一致する仕入先名があるか
For Each v In Terms

    Dim TmpVendorCode As String
    TmpVendorCode = GetVendorCode(CStr(v))
    
    '仕入先コードが取得できなければ次の要素へ
    If TmpVendorCode <> "" Then
    
        Cells(CurrentRow, 32).NumberFormatLocal = "@"
        Cells(CurrentRow, 32).Value = TmpVendorCode
        
        '注記フィールドを更新して、保留フラグを立てておく
        Dim VendorName As String, ExclamationNote As String
        VendorName = v
        
        '発注時の注意書き文言は、仕入先から仕入先名称を削除した残りの文言
        ExclamationNote = Trim(Replace(VenderTerm, VendorName, ""))
        
        If ExclamationNote <> "" Then
            If IsEmpty(Cells(CurrentRow, 35).Value) Then Cells(CurrentRow, 35).Value = ExclamationNote
            Cells(CurrentRow, 34).Value = 1
        End If

    End If

Next

End Sub

Function GetVendorCode(ByVal VendorTerm As String)

On Error Resume Next
    
Dim VendorCode As String
VendorCode = WorksheetFunction.VLookup(VendorTerm, Worksheets("仕入先").Range("B2:AA500"), 26, 0)

GetVendorCode = VendorCode

End Function
