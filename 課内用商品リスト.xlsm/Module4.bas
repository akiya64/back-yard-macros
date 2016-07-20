Attribute VB_Name = "Module4"
Sub SKU欄文字列化()

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
    
    '------------------------------

    For i = 2 To ThisWorkbook.Worksheets(data).Range("B200000").End(xlUp).Row
    
        With ThisWorkbook.Worksheets(data)
            '仕入先名
            .Cells(i, 2).NumberFormat = "@"
            .Cells(i, 2).Value = ThisWorkbook.Worksheets(data).Cells(i, 2).Value
            
            'JAN
            .Cells(i, 5).NumberFormat = "@"
            .Cells(i, 5).Value = ThisWorkbook.Worksheets(data).Cells(i, 2).Value
            
        End With
        
        
    Next i


End Sub
