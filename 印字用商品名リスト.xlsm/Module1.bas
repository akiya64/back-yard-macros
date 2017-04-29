Attribute VB_Name = "Module1"
Sub 情報更新()

    Dim entry As String
    entry = "手入力"
    
    Dim upload As String
    upload = "最終"
    
    Dim pasted As String
    pasted = "出品詳細レコード"
    
    Dim config As String
    config = "設定"
    
    'SKUの調整
    ThisWorkbook.Worksheets(pasted).Select
    
    For i = ThisWorkbook.Worksheets(pasted).Range("A500000").End(xlUp).Row To 2 Step -1
    
        If (g + 10000) Mod 10000 = 2 Then
        
            ThisWorkbook.Worksheets(pasted).Cells(i, 1).Activate
        
        End If
    
        If Len(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) = 5 And IsNumeric(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) Then
        
            With ThisWorkbook.Worksheets(pasted).Cells(i, 3)
            
                .NumberFormat = "@"
                .Value = "0" & ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
            
            End With
        
        ElseIf IsNumeric(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) Then
        
            With ThisWorkbook.Worksheets(pasted).Cells(i, 3)
            
                .NumberFormat = "###############################"
                .Value = "0" & ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
            
            End With
        
        End If
    

    
    
    '登録済ならば削除
    
        If Application.CountIf(ThisWorkbook.Worksheets(entry).Range("A:A"), ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) >= 1 Then
        
            ThisWorkbook.Worksheets(pasted).Range(i & ":" & i).Delete
        
        End If
    
    Next i

    '残った情報（新商品）のみ貼り付け
    ThisWorkbook.Worksheets(entry).Select
    
    For k = 2 To ThisWorkbook.Worksheets(pasted).Range("A500000").End(xlUp).Row

        If (k + 10000) Mod 10000 = 2 Then
        
            ThisWorkbook.Worksheets(entry).Cells(k, 1).Activate
        
        End If

        Dim cfr As String
        cfr = ThisWorkbook.Worksheets(entry).Range("A500000").End(xlUp).Row + 1
    
        With ThisWorkbook.Worksheets(entry)
            
            .Cells(cfr, 1).Value = ThisWorkbook.Worksheets(pasted).Cells(k, 3).Value
            
            If Len(Cells(cfr, 1).Value) = 5 And IsNumeric(Cells(cfr, 1).Value) Then
            
                .Cells(cfr, 1).NumberFormat = "@"
                .Cells(cfr, 1).Value = "0" & ThisWorkbook.Worksheets(entry).Cells(cfr, 1).Value
            
            ElseIf IsNumeric(Cells(cfr, 1).Value) Then
            
                .Cells(cfr, 1).NumberFormat = "##############################"
                .Cells(cfr, 1).Value = ThisWorkbook.Worksheets(entry).Cells(cfr, 1).Value
            
            End If
            
            .Cells(cfr, 11).Value = "=LEN(B" & cfr & ")+LEN(C" & cfr & ")+LEN(D" & cfr & ")+LEN(E" & cfr & ")+LEN(F" & cfr & ")+LEN(G" & cfr & ")+LEN(H" & cfr & ")+LEN(I" & cfr & ")+LEN(J" & cfr & ")"
            .Cells(cfr, 14).Value = ThisWorkbook.Worksheets(pasted).Cells(k, 1).Value
            
        End With
    
    Next k

End Sub
