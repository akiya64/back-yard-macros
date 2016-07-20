Attribute VB_Name = "Module5"
Sub ASIN追加()



    Dim amazon As String
    amazon = "amazon"
    
    Dim rakuten As String
    rakuten = "楽天"
    
    Dim yahoo As String
    yahoo = "ヤフー"
    
    Dim data As String
    data = "商品情報"
    
    Dim pasted As String
    pasted = "貼付"
    
    Dim config As String
    config = "設定"
    
    Dim lt As String
    lt = "LT"
    
    '------------------------------
    Dim nextRow As String
    For i = ThisWorkbook.Worksheets(lt).Range("A200000").End(xlUp).Row To 1 Step -1
    
        If ThisWorkbook.Worksheets(lt).Cells(i, 1).Value <> "" And ThisWorkbook.Worksheets(lt).Cells(i, 2).Value <> "" Then
            
            nextRow = Application.Match(ThisWorkbook.Worksheets(lt).Cells(i, 2).Value, ThisWorkbook.Worksheets(data).Range("B:B"), 0)
            ThisWorkbook.Worksheets(data).Cells(nextRow, 20).Value = ThisWorkbook.Worksheets(lt).Cells(i, 1).Value
        
        End If
        
    Next i
    


End Sub
Sub ASIN追加項目削除()



    Dim amazon As String
    amazon = "amazon"
    
    Dim rakuten As String
    rakuten = "楽天"
    
    Dim yahoo As String
    yahoo = "ヤフー"
    
    Dim data As String
    data = "商品情報"
    
    Dim pasted As String
    pasted = "貼付"
    
    Dim config As String
    config = "設定"
    
    Dim lt As String
    lt = "LT"
    
    '------------------------------
    ThisWorkbook.Worksheets(lt).Range("A:B").ClearContents


End Sub
