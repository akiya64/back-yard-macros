Attribute VB_Name = "Module2"
Sub データの整理()

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
    'ここはデータのトリム化
    For i = 2 To data.Range("B500000").End(xlUp).Row
        With data
            .Cells(i, 1).Value = Trim(data.Cells(i, 1).Value)
            .Cells(i, 2).Value = Trim(data.Cells(i, 2).Value)
            .Cells(i, 6).Value = Trim(data.Cells(i, 6).Value)
            .Cells(i, 20).Value = Trim(data.Cells(i, 20).Value)
        End With
    Next i
    
    
    '原価の埋め込み
    Dim finalrow As String
    finalrow = data.Range("B500000").End(xlUp).Row
    data.Activate
    data.Range("A2:AZ" & finalrow).Sort key1:=Range("A2"), order1:=xlAscending, key2:=Range("AA2"), order2:=xlDescending, key3:=Range("N2"), order3:=xlDescending
    
    For i = 2 To finalrow
        With data
            
            '原価情報更新
            If data.Cells(i, 1).Value <> "" And data.Cells(i, 1).Value = data.Cells(i + 1, 1).Value And data.Cells(i, 13).Value <> "" Then
            
                If data.Cells(i, 13).Value = 0 Then
                    .Cells(i, 13).Value = ""
                End If
                If data.Cells(i, 7).Value = 0 Then
                    .Cells(i, 7).Value = ""
                End If
                If data.Cells(i, 7).Value = "" Then
                    motonum = 1
                    tanka = Application.RoundUp(data.Cells(i, 13).Value, 0)
                Else
                    motonum = data.Cells(i, 7).Value
                    tanka = Application.RoundUp(data.Cells(i, 13).Value / motonum, 0)
                End If
                
                If data.Cells(i, 14).Value = "" Then
                    motodate = 0
                Else
                    motodate = data.Cells(i, 14).Value
                End If

                If data.Cells(i + 1, 7).Value = "" Then
                    henkonum = 1
                Else
                    henkonum = data.Cells(i + 1, 7).Value
                End If
                
                .Cells(i + 1, 6).Value = data.Cells(i, 6).Value
                .Cells(i + 1, 13).Value = tanka * henkonum
                .Cells(i + 1, 14).Value = Format(motodate, "yyyy/mm/dd")
                .Cells(i + 1, 27).Value = data.Cells(i, 27).Value
            End If
        End With
    Next i
    
    
    
    '''''''''''''''''''''''
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With

End Sub
