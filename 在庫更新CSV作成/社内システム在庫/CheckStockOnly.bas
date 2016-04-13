Attribute VB_Name = "CheckStockOnly"
Sub CheckEolInStockOnly()
'発注不可、在庫のみシートの商品コードについて、
'区分をチェック、区分がメ廃番、販売中止->廃番・終了にコピー、すでにリストアップ済なら行削除
'区分が処分・在廃で在庫数が0->廃番・終了にコピー、すでにリストアップ済なら行削除

StockOnly.Activate

Application.ScreenUpdating = False

Call SetRangeName

Dim i As Long
i = 2

Do Until IsEmpty(Cells(i, 3))
      
   Call CheckEol(i)  '行番号を参照渡しで該当行の6ケタについてチェック、参照渡しなので処理先でiを1行進めることができる。
    
Loop

Application.ScreenUpdating = True

End Sub

Private Sub CheckEol(i As Long)

Dim Code As String
Code = Cells(i, 3)

'廃番・終了リストに転記済みかチェック

If WorksheetFunction.CountIf(Range("EolCodeRange"), Code) > 0 Then
        
        Rows(i).Delete
        Exit Sub

End If

Dim sy As Syokon
sy = SyokonMaster.GetSyokonQtyKubun(Code)


'廃番・販売終了リストへの転記条件に一致するか？

If InStr(sy.Status, "廃番") > 0 Or InStr(sy.Status, "販売中止") > 0 Then

        Call PostEol(i)
        Exit Sub

ElseIf InStr(sy.Status, "在廃") > 0 Or InStr(sy.Status, "処分品") > 0 Then
    
    If sy.Quantity <= 0 Then
        
        Call PostEol(i)
        Exit Sub
    
    End If

End If

i = i + 1 '行を削除しなかった場合のみ、iを1進める

End Sub

Private Sub PostEol(i As Long)

Dim Code As String
Code = Cells(i, 3).Value

Call addCode(Code, "EolCodeRange")
Rows(i).Delete

End Sub
