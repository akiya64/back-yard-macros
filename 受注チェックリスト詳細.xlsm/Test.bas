Attribute VB_Name = "Test"
Option Explicit

Sub test_ParseCode()

Dim c As Range, v As Variant

For Each v In Range("B2:B32")

    Dim Code As String
    Code = v.Value
    
    '77777始まりセットコードなら
    If Code Like "77777*" Then
        
        Debug.Print v.Address & Chr(9) & v.Value & Chr(9) & "7777セット"
        Call ParseTiedItem(Cells(v.Row, 2))
        
    '-02 -04 -120 ハイフン-数量 セットなら ハイフンを含みアルファベット始まりでない
    ElseIf InStr(Code, "-") > 1 And Not Code Like "[a-zA-Z]*" Then
        
        Debug.Print v.Address & Chr(9) & v.Value & Chr(9) & "ハイフンセット"
        Call ParseMultipleSet(Cells(v.Row, 2))
    
    End If

Next

End Sub

Sub test_ParseTiedItem()

    Call ParseTiedItem(Cells(25, 2))

End Sub
