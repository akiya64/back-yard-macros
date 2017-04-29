Attribute VB_Name = "Module2"

Sub 反映処理()

    Dim entry As String
    entry = "手入力"
    
    Dim upload As String
    upload = "最終"
    
    Dim pasted As String
    pasted = "出品詳細レコード"
    
    Dim config As String
    config = "設定"
    

    With ThisWorkbook.Worksheets(upload).Range("2:1000000")
    
        .Clear
        .Interior.ColorIndex = xlNone
    
    End With
    
    For i = ThisWorkbook.Worksheets(entry).Range("A500000").End(xlUp).Row To 2 Step -1
    
        If ThisWorkbook.Worksheets(entry).Cells(i, 2).Value <> "" And ThisWorkbook.Worksheets(entry).Cells(i, 4).Value <> "" And ThisWorkbook.Worksheets(entry).Cells(i, 6).Value <> "" Then
        
            Dim newLine As String
            newLine = ThisWorkbook.Worksheets(upload).Range("A500000").End(xlUp).Row + 1
        
            With ThisWorkbook.Worksheets(upload)
                
                
                .Cells(newLine, 1).NumberFormat = "@"
                .Cells(newLine, 1).Value = ThisWorkbook.Worksheets(entry).Cells(i, 1).Value
                
                .Cells(newLine, 2).Value = ThisWorkbook.Worksheets(entry).Cells(i, 2).Value & ThisWorkbook.Worksheets(entry).Cells(i, 3).Value & ThisWorkbook.Worksheets(entry).Cells(i, 4).Value & ThisWorkbook.Worksheets(entry).Cells(i, 5).Value & ThisWorkbook.Worksheets(entry).Cells(i, 6).Value & ThisWorkbook.Worksheets(entry).Cells(i, 7).Value & ThisWorkbook.Worksheets(entry).Cells(i, 8).Value & ThisWorkbook.Worksheets(entry).Cells(i, 9).Value & ThisWorkbook.Worksheets(entry).Cells(i, 10).Value
            
            End With
            
            If ThisWorkbook.Worksheets(entry).Cells(i, 11).Value > ThisWorkbook.Worksheets(config).Range("B2").Value Then
            
                ThisWorkbook.Worksheets(upload).Range(newLine & ":" & newLine).Interior.Color = RGB(255, 0, 0)
            
            End If
        
        End If
    
    Next i
    
    Dim finalLine As String
    finalLine = ThisWorkbook.Worksheets(upload).Range("A500000").End(xlUp).Row
    
    ThisWorkbook.Worksheets(upload).Range((finalLine + 1) & ":" & (finalLine + 500000)).Clear
    

End Sub
