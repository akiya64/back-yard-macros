 	

Option Explicit

Type InhouseRecord

    Code As String
    Jan As String

End Type

Sub SKUがJANを社内コードで置き換え()

'DBダンプのエクセルブック名、元シートの範囲は毎回指定のこと
'イミディエイトで、Workbooks(1).nameでワークブック名が確認できる。
Dim Rng As Range
Set Rng = Workbooks("社内データ.xlsx").Sheets(1).Range("A2:A99579")

Dim r As Range
For Each r In Rng

    'Debug.Assert r.Row < 1000

    Dim ir As InhouseRecord
    
    ir.Code = r.Value
    ir.Jan = r.Offset(0, 2)

    '09始まりコードは資材のためとばす、05始まりは先頭0を落とす
    
    If ir.Code Like "09#####" Then
        
        GoTo continue
            
    ElseIf ir.Code Like "05#####" Then
        
        ir.Code = Mid(ir.Code, 2, 6)
    
    End If

    Call UpdateJan(ir)
    
continue:

Next

End Sub


Private Sub UpdateJan(item As InhouseRecord)

Dim c As Range

'A列の該当JANを探す
With Workbooks("商品情報.xlsm").Worksheets("商品情報").Columns(1)

'完全一致で
Set c = .Find(what:=item.Jan, LookIn:=xlValues, LookAt:=xlWhole)

If c Is Nothing Then Exit Sub

'最初のセルのアドレスを控える Rangeで見た方がいいかも？
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
    
    'B列にハイフンがなく、社内コードでもなければ、社内コードで上書き
    If Not Sku = item.Code And InStr(Sku, "-") < 1 Then
         c.Offset(0, 1).Value = item.Code
    End If
    
    'F列は全て上書き
    If KijunSku.Value <> item.Code Then
        KijunSku.Value = item.Code
    End If
    
    '次の検索をセット
    Set c = .FindNext(c)
    If c Is Nothing Then Exit Do

Loop Until c.Address = FirstAddress

End With

End Sub
