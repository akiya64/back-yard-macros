Attribute VB_Name = "Module1"
Sub 出品商品情報取込()
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With
    '''''''''''''''''''''''
    Dim top As Worksheet
    Set top = ThisWorkbook.Worksheets("トップ")
    Dim data As Worksheet
    Set data = ThisWorkbook.Worksheets("商品情報")
    Dim config As Worksheet
    Set config = ThisWorkbook.Worksheets("設定")
    Dim master As Worksheet
    Set master = ThisWorkbook.Worksheets("m")
    ''''''''''''''''''''''''''''''''''''''''''
    '出品詳細レポートの呼び出し
    Dim filepath As String
    filepath = config.Range("C3").Value & "\s" & Format(Date, "yyyymmdd") & ".txt"
    Dim fileOpen As Workbook
    Set fileOpen = Workbooks.Open(filepath)
    Dim finalrow As String
    finalrow = fileOpen.Worksheets(1).Range("C500000").End(xlUp).Row
    Dim skurow As String
    fileOpen.Worksheets(1).Range("C:C").NumberFormat = "@"
    For i = 2 To finalrow
        With fileOpen.Worksheets(1)
            If Len(fileOpen.Worksheets(1).Cells(i, 3).Value) = 10 And IsNumeric(fileOpen.Worksheets(1).Cells(i, 3).Value) Then
                .Cells(i, 3).Value = "000" & fileOpen.Worksheets(1).Cells(i, 3).Value
            ElseIf Len(fileOpen.Worksheets(1).Cells(i, 3).Value) = 11 And IsNumeric(fileOpen.Worksheets(1).Cells(i, 3).Value) Then
                .Cells(i, 3).Value = "00" & fileOpen.Worksheets(1).Cells(i, 3).Value
            ElseIf Len(fileOpen.Worksheets(1).Cells(i, 3).Value) = 12 And IsNumeric(fileOpen.Worksheets(1).Cells(i, 3).Value) Then
                .Cells(i, 3).Value = "0" & fileOpen.Worksheets(1).Cells(i, 3).Value
            ElseIf Len(fileOpen.Worksheets(1).Cells(i, 3).Value) = 5 And IsNumeric(fileOpen.Worksheets(1).Cells(i, 3).Value) Then
                .Cells(i, 3).Value = "0" & fileOpen.Worksheets(1).Cells(i, 3).Value
            Else
                .Cells(i, 3).Value = "" & fileOpen.Worksheets(1).Cells(i, 3).Value
            End If
        End With
        
        With data
            If Application.CountIf(data.Range("B:B"), fileOpen.Worksheets(1).Cells(i, 3).Value) = 0 Then
                finalrow = data.Range("B500000").End(xlUp).Row + 1
                If fileOpen.Worksheets(1).Cells(i, 7).Value = 3 Then
                    .Cells(finalrow, 1).Value = fileOpen.Worksheets(1).Cells(i, 12).Value
                    If Len(data.Cells(finalrow, 1).Value) = 10 And IsNumeric(data.Cells(finalrow, 1).Value) Then
                        Cells(finalrow, 1).Value = "000" & data.Cells(finalrow, 1).Value
                    ElseIf Len(data.Cells(finalrow, 1).Value) = 11 And IsNumeric(data.Cells(finalrow, 1).Value) Then
                        Cells(finalrow, 1).Value = "00" & data.Cells(finalrow, 1).Value
                    ElseIf Len(data.Cells(finalrow, 1).Value) = 12 And IsNumeric(data.Cells(finalrow, 1).Value) Then
                        Cells(finalrow, 1).Value = "0" & data.Cells(finalrow, 1).Value
                    ElseIf Len(data.Cells(finalrow, 1).Value) = 5 And IsNumeric(data.Cells(finalrow, 1).Value) Then
                        Cells(finalrow, 1).Value = "0" & data.Cells(finalrow, 1).Value
                    Else
                        Cells(finalrow, 1).Value = "" & data.Cells(finalrow, 1).Value
                    End If
                End If
                .Cells(finalrow, 2).Value = fileOpen.Worksheets(1).Cells(i, 3).Value
                .Cells(finalrow, 3).Value = fileOpen.Worksheets(1).Cells(i, 1).Value
                If fileOpen.Worksheets(1).Cells(i, 7).Value = 1 Then
                    .Cells(finalrow, 20).Value = fileOpen.Worksheets(1).Cells(i, 12).Value
                End If
            End If
        End With
    Next i
    fileOpen.Close False
    
    
    '''''''''''''''''''''''
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
End Sub
