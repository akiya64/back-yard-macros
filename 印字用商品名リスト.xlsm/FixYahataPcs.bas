Attribute VB_Name = "FixYahataPcs"
Option Explicit

Sub ExtractProcuctName()

'正規表現オブジェクトはRegExp
Dim re As Object
Set re = CreateObject("VBScript.RegExp")

re.Global = True

'○個入り検出パターン
re.Pattern = "[0-9]{1,4}個入り"

Dim i As Long
For i = 131070 To 135086
    
    If Not Cells(i, 1).Value Like "*-*" Then GoTo Continue
    
    Dim str As String
    str = Cells(i, 2).Value
    Cells(i, 2).Value = re.Replace(str, "") & Cells(i, 3).Value
    
Continue:

Next

End Sub

Sub ExtractYahataJanPcs()

'正規表現オブジェクトはRegExp
Dim re As Object
Set re = CreateObject("VBScript.RegExp")

re.Global = True

'○個入り検出パターン
re.Pattern = "[0-9]{1,4}個入り"

Dim i As Long
For i = 131070 To 135086
    
    If Cells(i, 1).Value Like "*-*" Then GoTo Continue
    Dim str As String
    
    str = Cells(i, 2).Value
    
    Dim Matches
    
    Set Matches = re.Execute(str)
    
    If Matches.Count = 0 Then GoTo Continue
    Cells(i, 3).Value = Matches(0)

Continue:

Next

End Sub

Sub ReCalcurationSetQty(RowCount As Long)

Do
    Dim Code As String
    Code = Cells(RowCount, 1).Value
    
    Dim SinglexQty As String, SinglexJAN As String, CurrentSinglexJan As String
    
    'ハイフン有りの時は分解して再算出
    If Code Like "*-*" Then
             
        Dim SeparatedCode As Variant
        SeparatedCode = Split(Code, "-", 2)
        
        Dim SetQty As String
        SetQty = CStr(Val(SeparatedCode(1)))
        
        Dim SetQtyStr As String
        SetQtyStr = SinglexQty & "×" & SetQty
       
        '算出し直した個数表記をC列へ格納
        Cells(RowCount, 3).Value = SetQtyStr
        
        CurrentSinglexJan = SeparatedCode(0)
        
    Else
    'ハイフンなし＝単体JANの場合、単体での内容量を格納
    
        SinglexQty = Cells(RowCount, 3).Value
        SinglexJAN = Cells(RowCount, 1).Value
        
        CurrentSinglexJan = Cells(RowCount, 1).Value
    
    End If
    
    '参照渡しなので元カウンタを進めておく
    RowCount = RowCount + 1

Loop While CurrentSinglexJan = SinglexJAN

End Sub

Sub Recalc()

Dim i As Long
For i = 131070 To 140000

    Call ReCalcurationSetQty(i)

Next

End Sub
