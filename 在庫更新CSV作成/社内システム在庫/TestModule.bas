Attribute VB_Name = "TestModule"
Option Explicit

Sub test_LastSecondInbentryGetQuantity()

    If IsNumeric(LastSecondInventry.getQuantity("4953571171524")) Then
        MsgBox "OK"
    Else
        MsgBox "NG!"
    End If

End Sub

Sub test_appendtime()

    Call Log.ApendProcessingTime(21)

End Sub

Sub test_GetSecondInventryQuantity()

    Debug.Print SecondInventry.getQuantity("4953571172101")
    
End Sub

Sub test_SortSecondInventry()
    
    Worksheets("ネット用在庫").Range("A1").AutoFilter
    
    With Worksheets("ネット用在庫").AutoFilter.Sort
    
        .SortFields.Clear 'ソートフィールドを一旦クリアー
        
        'ソートフィールドを指定
        .SortFields.Add Key:=Range("C1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        
        'ソート順序指定
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        
        'ソート適用
        .Apply

    End With


End Sub
