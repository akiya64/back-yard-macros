Attribute VB_Name = "Module7"
Sub 単価調査抽出()

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
    ''''''''''''''''''''''''''''''''''''''''''
    With chosa
        .Activate
        .Cells.Clear
        .Range("A1:AZ1").Value = data.Range("A1:AZ1").Value
        .Range("A:B").NumberFormat = "@"
        .Range("F:F").NumberFormat = "@"
    End With
    finalrow = data.Range("B500000").End(xlUp).Row
    With data
        .Activate
        .Range("A2:AZ" & finalrow).Sort key1:=Range("D2"), order1:=xlAscending, key2:=Range("A2"), order2:=xlDescending, key3:=Range("N2"), order3:=xlDescending
    End With
    With chosa
        .Activate
        If Not IsNumeric(config.Range("A32").Value) Or config.Range("A32").Value = "" Then
            config.Range("A32").Value = 60
        End If
    End With
    Dim janrow As String
    For i = 2 To finalrow
        janrow = chosa.Range("B500000").End(xlUp).Row
        With chosa
            If data.Cells(i, 1).Value <> "" And data.Cells(i, 1).Value <> data.Cells(i - 1, 1).Value And (data.Cells(i, 13).Value <> "" Or Date - config.Range("A32").Value > data.Cells(i, 14).Value) Then
                .Range(janrow + 1 & ":" & janrow + 1).Value = data.Range(i & ":" & i).Value
            End If
        End With
    Next i
    finalrow = chosa.Range("B500000").End(xlUp).Row
    With chosa
        .Activate
        .Range("A2:AZ" & finalrow).Sort key1:=Range("D2"), order1:=xlAscending, key2:=Range("A2"), order2:=xlDescending
    End With
    '''''''''''''''''''''''
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With

End Sub

