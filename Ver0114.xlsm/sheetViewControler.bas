Attribute VB_Name = "sheetViewControler"
Option Explicit


Sub 読込フォームを表示()
    
''月曜日の一回目の起動は、フィルター解除、受注データ整合性チェックを促すアラート
''月曜日でFirstOpenがTrueならアラート
'
'If Weekday(Date) = 2 Then
'
'    If LogSheet.Range("B5").Value = True Then
'
'        MsgBox prompt:="月曜日の初回起動です。" & vbLf & "注残一覧の前週分を確認してください。", _
'                Buttons:=vbExclamation
'
'        OrderSheet.AutoFilterMode = False
'
'        LogSheet.Range("B5").Value = False
'
'        End
'
'    End If
'
'Else
'
''月曜以外なら､Trueをセットしておく
'
'    LogSheet.Range("B5").Value = True
'
'End If

OpPanel.Show

End Sub


Sub hideWishCol()
    
    OrderSheet.Outline.ShowLevels ColumnLevels:=1
    
End Sub


Sub 未発送のみ表示()
Attribute 未発送のみ表示.VB_Description = "発送列 空欄 と 「出荷通知除外」を表示"
Attribute 未発送のみ表示.VB_ProcData.VB_Invoke_Func = "u\n14"

OrderSheet.Activate

Application.ScreenUpdating = False

'オートフィルターがセットされていなければ、15列目の「発送」空欄と「出荷通知除外」のみ表示で設定
If Not OrderSheet.AutoFilterMode Then
    
    Range("A1").AutoFilter Field:=15, Criteria1:="=出荷通知除外", Operator:=xlOr, Criteria2:="="

Else
    'フィルターがセットされていれば、セットし直す設定
    Dim i As Integer
    For i = 1 To 17

        If i = 15 Then
           Range("A1").AutoFilter Field:=i, Criteria1:="=出荷通知除外", Operator:=xlOr, Criteria2:="="

        Else
           Range("A1").AutoFilter i  '他はフィルター解除、Criteria指定を省略で「全て」表示
        
        End If

    Next

End If

End Sub

Sub 発送列の空欄のみ表示()
Attribute 発送列の空欄のみ表示.VB_ProcData.VB_Invoke_Func = " \n14"

'fillterShippingNull

OrderSheet.Activate

Application.ScreenUpdating = False

'オートフィルターがセットされていなければ、15列目の「発送」空欄のみ表示
If Not OrderSheet.AutoFilterMode Then
    
    Range("A1").AutoFilter Field:=15, Criteria1:="="

Else
    'フィルターがセットされていれば、セットし直す設定
    Dim i As Integer
    For i = 1 To 17
        
        If i = 15 Then
           Range("A1").AutoFilter Field:=i, Criteria1:="="
        
        Else
           Range("A1").AutoFilter i  '他はフィルター解除、Criteria指定を省略で「全て」表示
        
        End If
    
    Next

End If

End Sub
