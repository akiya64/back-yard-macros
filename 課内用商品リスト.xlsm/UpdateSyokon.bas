Attribute VB_Name = "UpdateSyokon"
Option Explicit

Type Syokon

    Code As String
    Jan As String
    VendorCode As String
    
End Type

Sub SKUがJANを商魂6ケタで置き換え()

'元のエクセルブック名、元シートの範囲は毎回指定のこと
'イミディエイトで、Workbooks(1).nameでワークブック名が確認できる。
Dim Rng As Range
Set Rng = Workbooks("商品マスタA.xlsx").Sheets(1).Range("A2:A112378")

Dim r As Range
For Each r In Rng

    'Debug.Assert r.Row < 1000

    Dim sy As Syokon
    
    'ToDo 5始まりコードは、先頭0を落とす
    
    sy.Code = r.Value
    sy.Jan = r.Offset(0, 3)
    
    '9始まり、1始まり6ケタは資材・什器のため飛ばす
    If sy.Code Like "9#####" Or sy.Code Like "1#####" Then
        
        GoTo Continue
            
    End If

    Call UpdateJan(sy)
    
Continue:

Next

ThisWorkbook.Close savechanges:=True

End Sub


Private Sub UpdateJan(Syokon As Syokon)

Dim c As Range

'A列の該当JANを探す
With Workbooks("発注用商品情報.xlsm").Worksheets("商品情報").Columns(1)

'完全一致で
Set c = .Find(what:=Syokon.Jan, LookIn:=xlValues, LookAt:=xlWhole)

If c Is Nothing Then Exit Sub
If Cells(c.Row, 2).Value Like "######" Then Exit Sub
'最初のセルのアドレスを控える
Dim FirstAddress As String
FirstAddress = c.Address

'繰返し検索し、条件を満たすすべてのセルを検索する
Do

    Dim SkuCell As Range, Sku As String
    Set SkuCell = c.Offset(0, 1)
    Sku = SkuCell.Value
    
    'あえてローマ字変数にする、F列見出しの元の意図が不明なので
    Dim KijunSku As Range
    Set KijunSku = Cells(c.Row, 6)
    
    'B列にハイフンがなく、6ケタでもなければ、6ケタで上書き
    If Not Sku = Syokon.Code And InStr(Sku, "-") < 1 Then
         c.Offset(0, 1).Value = IIf(Len(Syokon.Code) = 5, "0" & Syokon.Code, Syokon.Code)
    End If
    
    'F列は全て上書き
    If KijunSku.Value <> Syokon.Code Then
        KijunSku.Value = IIf(Len(Syokon.Code) = 5, "0" & Syokon.Code, Syokon.Code)
    End If
    
    '次の検索をセット
    Set c = .FindNext(c)
    If c Is Nothing Then Exit Do

Loop Until c.Address = FirstAddress

End With

End Sub

Sub 商魂の仕入先に合わせる()

Dim FinalRow As Long, i As Long
FinalRow = Worksheets("商品情報").UsedRange.Rows.Count

For i = 110000 To FinalRow
    
    Call UpdateVendor(i)

Next


ThisWorkbook.Close savechanges:=True

End Sub

Private Sub UpdateVendor(ByVal Row As Long)

Dim CurrentVendor As String
CurrentVendor = Cells(Row, 4).Value

Dim CurrentCode As String
CurrentCode = Cells(Row, 2).Value

'仕入先名が空なら、商品マスタ-仕入先コード-手配書作成「仕入先」の名称に基づいた仕入先名を入れる
If CurrentVendor = "" Then
    Dim NewVendorName
    NewVendorName = GetVendorName(GetSyokonVendor(CurrentCode))
    
    If NewVendorName <> "" Then
    
        Cells(Row, 4).Value = NewVendorName
    
    End If
    
    Exit Sub
End If

'手配書作成-仕入先シートの仕入先コードを取得
Dim CurrentVendorCode As String
CurrentVendorCode = GetVendorCodeFromPurBook(CurrentVendor)

'商品マスタの仕入先コードを取得
Dim SyokonVendorCode As String
SyokonVendorCode = GetSyokonVendor(CurrentCode)
If SyokonVendorCode = "" Then Exit Sub

'商品マスタの仕入先コードと一致するかチェック
'不一致ならば、仕入先シートの名称で上書きする
If CurrentVendorCode <> "" And CurrentVendorCode <> SyokonVendorCode Then
    Cells(Row, 4).Value = GetVendorName(SyokonVendorCode)
End If

End Sub

Private Function GetVendorCodeFromPurBook(ByVal VendorName As String) As String
On Error Resume Next

GetVendorCodeFromPurBook = WorksheetFunction.VLookup(VendorName, Worksheets("仕入先").Range("B2:AA490"), 26, False)

End Function

Private Function GetSyokonVendor(ByVal Code As String) As String
On Error Resume Next

GetSyokonVendor = Workbooks("商品マスタA.xlsx").Worksheets(1).Range("A:A").Find(CDbl(Code)).Offset(0, 4).Value

End Function

Private Function GetVendorName(ByVal VendorCode As String) As String
On Error Resume Next

GetVendorName = WorksheetFunction.VLookup(VendorCode, Worksheets("仕入先").Range("AA2:AB550"), 2, False)

End Function
