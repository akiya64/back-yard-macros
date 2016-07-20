Attribute VB_Name = "Module8"
Sub 情報入れ込み()
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
    Dim chosa As Worksheet
    Set chosa = ThisWorkbook.Worksheets("単価調査リスト")
    Dim entry As Worksheet
    Set entry = ThisWorkbook.Worksheets("情報追加")
    ''''''''''''''''''''''''''''''''''''''''''
    Dim datarow As String
    Dim finalrow As String
    finalrow = entry.Range("B500000").End(xlUp).Row
    With data
        For i = 2 To finalrow
            If Not IsError(Application.Match(entry.Cells(i, 2).Value, data.Range("B:B"), 0)) Then
            '登録がある場合
                datarow = Application.Match(entry.Cells(i, 2).Value, data.Range("B:B"), 0)
                For k = 1 To 30
                    If k <> 2 Then
                        If data.Cells(datarow, k).Value <> entry.Cells(i, k).Value And entry.Cells(i, k).Value <> "" Then
                            .Cells(datarow, k).Value = entry.Cells(i, k).Value
                        End If
                    End If
                Next k
            Else
            '登録がない場合
                datarow = data.Range("B500000").End(xlUp).Row + 1
                For k = 1 To 30
                    If data.Cells(datarow, k).Value <> entry.Cells(i, k).Value And entry.Cells(i, k).Value <> "" Then
                        .Cells(datarow, k).Value = entry.Cells(i, k).Value
                    End If
                Next k
            End If

        Next i
    End With

    MsgBox ("原価入れ込み完了")
    ''''''''''''''''''''''''''''''''''''''''''
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
End Sub
Sub 情報入れ込みクリア()
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
    Dim chosa As Worksheet
    Set chosa = ThisWorkbook.Worksheets("単価調査リスト")
    Dim entry As Worksheet
    Set entry = ThisWorkbook.Worksheets("情報追加")
    ''''''''''''''''''''''''''''''''''''''''''
    Dim datarow As String
    Dim finalrow As String
    finalrow = entry.Range("B500000").End(xlUp).Row
    entry.Range("A2:AZ" & finalrow).Cells.Clear
    ''''''''''''''''''''''''''''''''''''''''''
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
End Sub



