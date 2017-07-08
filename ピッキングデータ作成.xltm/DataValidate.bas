Attribute VB_Name = "DataValidate"
Option Explicit
Sub FixForAddin(Optional ByVal arg As Boolean)
'社内DBと照合できるように受注データシートに対して、受注商品コード列のコードをアドイン用商品コードへ転記する。
'○個組セット分解もここで行う。

Worksheets("受注データシート").Activate

Dim CodeRange As Range, c As Range
Set CodeRange = Range(Cells(2, 2), Cells(Range("B1").SpecialCells(xlCellTypeLastCell).Row, 2))

'アドイン用のコードを記入する
For Each c In CodeRange
    
    Dim CurrentCodeCell As Range
    Set CurrentCodeCell = c
    
    'I列、アドイン実行用に6ケタ化したコード、もしくはJANを入れる
    Cells(c.Row, 9).NumberFormatLocal = "@"
    Cells(c.Row, 9).Value = DataValidate.ValidateCode(c.Value)
    
    '必要数量、一旦受注の数量で埋める。セット分解後に書き換えられる。
    Cells(c.Row, 10).Value = Cells(c.Row, 4).Value

    '○個組分解
    If c.Value Like "*-*" Then
    
        Call SetParser.ParseScalingSet(c)
    
    End If

Next

End Sub
Sub FilterLocation(Optional ByVal arg As Boolean)
'受注データシートの全ての行に対して、ロケーション列から無効なロケーション文字を削除して有効ロケーション列へ転記。

OrderSheet.Activate

Dim LastRow As Long, i As Long
LastRow = Range("A1").SpecialCells(xlCellTypeLastCell).Row

For i = 2 To LastRow
    
    'ロケーション修正、商品名バリデーション
    Cells(i, 11).Value = CutOffUnlocation(Cells(i, 15).Value)
    
Next

End Sub

Function CutOffUnlocation(Location As String) As String
'正規表現でロケーション[0-0-0-0][0- -0- - ][1-0-0-0-0]などを削除して返します。

Dim Reg As New RegExp

Reg.Global = True

'ロケーションの並び 階-通路-棚番-段-順  棚番はA～Qアルファベット
Reg.Pattern = "\[[0-9|\s]\-[0-2|\s]\-[0-9|\s]\-[0|\s]\-(([0-9]{1,})|\s)\]"

CutOffUnlocation = Reg.Replace(Location, "")

End Function

Function ValidateName(Name As String) As String
'正規表現で商品名の修正。
'カンマ・ピリオドなどを削除、冒頭の【】≪≫で括られた楽天のセール文言削除


Dim Reg As New RegExp

Reg.Global = True
Reg.Pattern = ",|\!|\.|&"

Name = Reg.Replace(Name, "")


Reg.Pattern = "^((≪|【).*?(】|≫))*"
Name = Reg.Replace(Name, "")

ValidateName = Name

End Function

Function ValidateCode(Code As String) As String
'コードを受け取って、数字以外を削除・13ケタ/6ケタに足りない場合は冒頭0を補完したコードを返す

Dim FixedCode As String

'アルファベットを削除
Dim Reg As New RegExp
Reg.Global = True
Reg.Pattern = "[a-zA-Z]"
Code = Reg.Replace(Code, "")

'6ケタならそのまま入れる
If Code Like String(6, "#") Then
    FixedCode = Code

'数字5ケタは頭にゼロを追記
ElseIf Code Like String(5, "#") Then
    
    FixedCode = "0" & Code

'JANもそのまま入れる
ElseIf Code Like String(13, "#") Then
    
    FixedCode = Code
    
'数字7ケタ以上12ケタなら、13ケタになるよう先頭に0を追記
ElseIf Code Like (String(7, "#") & "*") And Len(Code) <= 12 Then

    FixedCode = String(13 - Len(Code), "0") & Code
    
Else
'どの条件にも一致しない場合でも、値は返す
    
    FixedCode = Code
    
End If

ValidateCode = FixedCode

End Function

Sub テスト_棚なしロケーション削除()

Dim UnlocaitonStrings As Variant, LocString As Variant, TestResult As Boolean

TestResult = True

UnlocaitonStrings = Array( _
                            "[4-0- -0-0]", _
                            "[5-0- -0-0]", _
                            "[2-0- -0-0]", _
                            "[3-0- -0-5]", _
                            "[0-0-0-0-0]", _
                            "[ -1- -0-0]", _
                            "[ -0- -0-0]", _
                            "[ -0-2-0-0]", _
                            "[ - - - - ]", _
                            "[1-0- -0-15]", _
                            "[3-15-1-2-6]" _
                        )

For Each LocString In UnlocaitonStrings

    If CutOffUnlocation(CStr(LocString)) <> "" Then
        TestResult = False
        Debug.Print "Miss! ;" & LocString
    End If

Next

Dim ValidLocationStrings As Variant
ValidLocationStrings = Array( _
                            "[3-14-I-4-6]", _
                            "[1-6-R-4-3]", _
                            "[9-55-A-2-3-9]" _
                            )

For Each LocString In ValidLocationStrings

    If CutOffUnlocation(CStr(LocString)) = "" Then
        TestResult = False
        Debug.Print "Don't Cut! ;" & LocString
    End If

Next

If TestResult = True Then
    Debug.Print "Test Passed!"
Else
    Debug.Print "Test Missed"
End If

End Sub

