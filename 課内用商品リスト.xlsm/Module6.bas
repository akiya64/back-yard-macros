Attribute VB_Name = "Module6"
Sub バックアップコピー()

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
    '------------------------

    Dim fileName As String
    fileName = ThisWorkbook.Name
    
    Dim before As String
    before = ThisWorkbook.Path & "\" & fileName
    
    Dim after As String
    after = ThisWorkbook.Worksheets(config).Range("C22").Value & "\BU" & Format(Date, "yyyymmdd") & "_" & Format(Now, "hhmmss") & ".xls"

    Workbooks.Add
    
    Dim finalrow As String
    finalrow = ThisWorkbook.Worksheets(data).Range("B500000").End(xlUp).Row
    
    Dim finalRange As String
    finalRange = "A1:Z" & finalrow
    
    ActiveWorkbook.Worksheets(1).Range(finalRange).Value = ThisWorkbook.Worksheets(data).Range(finalRange).Value
    
    ActiveWorkbook.SaveAs fileName:=after
    ActiveWorkbook.Close True

    


End Sub

Sub アマゾン未登録分抽出()

    Call バックアップコピー
    Call アマゾン未登録分抽出A
    Call アマゾン未登録分抽出B
    Call アマゾン未登録分抽出C
    
    MsgBox ("取込完了!!" & vbNewLine & vbNewLine & "これ以降の発注作業に影響が出るため、次回の発注までに必ずSKUと商品名以外の情報を追加してください。")

End Sub

Sub アマゾン未登録分抽出A()

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
    ThisWorkbook.Worksheets(pasted).Cells.Clear


    Dim filepath As String
    filepath = ThisWorkbook.Worksheets(config).Range("C3").Value & "/" & ThisWorkbook.Worksheets(config).Range("B3").Value & ".txt"

    Dim fileOpen As Workbook
    Set fileOpen = Workbooks.Open(filepath)
    
    Dim finalrow As String
    finalrow = fileOpen.Worksheets(1).Range("A300000").End(xlUp).Row
    
    Dim copyRange As String
    copyRange = "A1:AZ" & finalrow
    
    '貼付
    With ThisWorkbook.Worksheets(pasted)
    
        
        .Range("C:C").NumberFormat = "@"
        .Range(copyRange).Value = fileOpen.Worksheets(1).Range(copyRange).Value
        .Range("D:AZ").Clear
    
    End With

    fileOpen.Close False



End Sub

Sub アマゾン未登録分抽出B()

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



    'SKUを編集
    Dim pastedFinalRow As String
    For i = 2 To ThisWorkbook.Worksheets(pasted).Range("A300000").End(xlUp).Row
    
        If Len(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) = 5 And IsNumeric(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) Then
        
            With ThisWorkbook.Worksheets(pasted)
            
                .Cells(i, 3).NumberFormat = "@"
                .Cells(i, 3).Value = "0" & ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
            
            End With
        
        ElseIf Len(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) = 6 And IsNumeric(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) Then
        
            With ThisWorkbook.Worksheets(pasted)
            
                .Cells(i, 3).NumberFormat = "@"
                .Cells(i, 3).Value = "" & ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
            
            End With
        
        ElseIf Len(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) = 10 And IsNumeric(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) Then
        
            With ThisWorkbook.Worksheets(pasted)
            
                .Cells(i, 3).NumberFormat = "#########################"
                .Cells(i, 3).Value = ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
                .Cells(i, 3).NumberFormat = "@"
                .Cells(i, 3).Value = "000" & ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
            
            End With
        
        ElseIf Len(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) = 11 And IsNumeric(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) Then
        
            With ThisWorkbook.Worksheets(pasted)
            
                .Cells(i, 3).NumberFormat = "#########################"
                .Cells(i, 3).Value = ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
                .Cells(i, 3).NumberFormat = "@"
                .Cells(i, 3).Value = "00" & ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
            
            End With
        
        ElseIf Len(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) = 12 And IsNumeric(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) Then
        
            With ThisWorkbook.Worksheets(pasted)
            
                .Cells(i, 3).NumberFormat = "#########################"
                .Cells(i, 3).Value = ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
                .Cells(i, 3).NumberFormat = "@"
                .Cells(i, 3).Value = "0" & ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
            
            End With
        
        ElseIf Len(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) = 13 And IsNumeric(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) Then
        
            With ThisWorkbook.Worksheets(pasted)
            
                .Cells(i, 3).NumberFormat = "#########################"
                .Cells(i, 3).Value = ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
                .Cells(i, 3).NumberFormat = "@"
                .Cells(i, 3).Value = "" & ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
            
            End With
        
        End If
    
    Next i


End Sub
Sub アマゾン未登録分抽出C()

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

    Dim finalrow As String
    For i = 2 To ThisWorkbook.Worksheets(pasted).Range("B300000").End(xlUp).Row
    
        If Application.CountIf(ThisWorkbook.Worksheets(data).Range("B:B"), ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) < 1 Then
        
            finalrow = ThisWorkbook.Worksheets(data).Range("B300000").End(xlUp).Row
            With ThisWorkbook.Worksheets(data)
            
                .Cells(finalrow + 1, 2).NumberFormat = "@"
                .Cells(finalrow + 1, 2).Value = ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
                .Cells(finalrow + 1, 3).NumberFormat = "@"
                .Cells(finalrow + 1, 3).Value = ThisWorkbook.Worksheets(pasted).Cells(i, 1).Value

            
            End With
        
        End If
    
    Next i

ThisWorkbook.Worksheets(pasted).Cells.Clear


End Sub

