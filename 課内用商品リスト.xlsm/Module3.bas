Attribute VB_Name = "Module3"
Sub ¤°ƒf[ƒ^‚Ìæ()

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With
    '''''''''''''''''''''''
    Dim top As Worksheet
    Set top = ThisWorkbook.Worksheets("ƒgƒbƒv")
    Dim data As Worksheet
    Set data = ThisWorkbook.Worksheets("¤•iî•ñ")
    Dim config As Worksheet
    Set config = ThisWorkbook.Worksheets("İ’è")
    Dim master As Worksheet
    Set master = ThisWorkbook.Worksheets("m")
    ''''''''''''''''''''''''''''''''''''''''''
    '‚UŒ…ã‚Ì‚à‚Ì‚ğˆê“x®—
    For a = master.Range("A500000").End(xlUp).Row To 2 Step -1
        With master
            .Cells(a, 1).NumberFormat = "@"
            .Cells(a, 1).Value = Right(master.Cells(a, 1).Value, 6)
        End With
    Next a
    
    
    
    '—áŠOˆ—ˆÈŠO‚ğİ’è
    Dim bairitsu As String
    Dim skurow As String
    For i = 2 To data.Range("B500000").End(xlUp).Row
        With data
            If data.Cells(i, 27).Value = "-" Then
                .Cells(i, 28).Value = "•ÏX‚µ‚È‚¢"
            ElseIf data.Cells(i, 27).Value <> "" And data.Cells(i, 27).Value > Date And IsDate(data.Cells(i, 27).Value) Then
                .Cells(i, 28).Value = "¤°æ‚è‚Ü‚È‚¢"
            Else
            '‚»‚Ì‘¼‚Ì‚İæ
                '”{—¦Œˆ’è
                bairitsu = 1
                If data.Cells(i, 7).Value <> "" And IsNumeric(data.Cells(i, 7).Value) Then
                    bairitsu = data.Cells(i, 7).Value
                End If
                
                'JAN‚Åˆø‚Á’£‚è
                If Not IsError(Application.Match(data.Cells(i, 1).Value, master.Range("C:C"), 0)) Then
                    skurow = Application.Match(data.Cells(i, 1).Value, master.Range("C:C"), 0)
                    If Application.RoundUp(bairitsu * master.Cells(skurow, 6).Value, 0) <> data.Cells(i, 13).Value Then
                        .Cells(i, 13).Value = Application.RoundUp(bairitsu * master.Cells(skurow, 6).Value, 0)
                        .Cells(i, 14).Value = Format(Date, "yyyy/mm/dd")
                    End If
                    
                    If InStr(data.Cells(i, 4).Value, "”p") = 0 Or InStr(data.Cells(i, 4).Value, "ˆ•ª") = 0 Or InStr(data.Cells(i, 4).Value, "’†~") = 0 Then
                        If InStr(master.Cells(skurow, 15).Value, "”p") > 0 Then
                            .Cells(i, 4).Value = data.Cells(i, 4).Value & " ”p”Ô"
                        ElseIf InStr(master.Cells(skurow, 15).Value, "ˆ•ª") > 0 Then
                            .Cells(i, 4).Value = data.Cells(i, 4).Value & " ˆ•ª•iŠ®”„"
                        ElseIf InStr(master.Cells(skurow, 15).Value, "’†~") > 0 Then
                            .Cells(i, 4).Value = data.Cells(i, 4).Value & " ”Ì”„’†~"
                        End If
                    End If
                End If
                
                'SKU‚Åˆø‚Á’£‚è
                skurow = ""
                If Not IsError(Application.Match(data.Cells(i, 2).Value, master.Range("A:A"), 0)) Then
                    skurow = Application.Match(data.Cells(i, 2).Value, master.Range("A:A"), 0)
                    If Application.RoundUp(bairitsu * master.Cells(skurow, 6).Value, 0) <> data.Cells(i, 13).Value Then
                        .Cells(i, 13).Value = Application.RoundUp(bairitsu * master.Cells(skurow, 6).Value, 0)
                        .Cells(i, 14).Value = Format(Date, "yyyy/mm/dd")
                    End If
                    
                    If InStr(data.Cells(i, 4).Value, "”p") = 0 Or InStr(data.Cells(i, 4).Value, "ˆ•ª") = 0 Or InStr(data.Cells(i, 4).Value, "’†~") = 0 Then
                        If InStr(master.Cells(skurow, 15).Value, "”p") > 0 Then
                            .Cells(i, 4).Value = data.Cells(i, 4).Value & " ”p”Ô"
                        ElseIf InStr(master.Cells(skurow, 15).Value, "ˆ•ª") > 0 Then
                            .Cells(i, 4).Value = data.Cells(i, 4).Value & " ˆ•ª•iŠ®”„"
                        ElseIf InStr(master.Cells(skurow, 15).Value, "’†~") > 0 Then
                            .Cells(i, 4).Value = data.Cells(i, 4).Value & " ”Ì”„’†~"
                        End If
                    End If
                End If
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
            End If
    
        End With
    Next i
    

    '''''''''''''''''''''''
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With

End Sub

