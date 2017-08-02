Attribute VB_Name = "Module1"
Option Explicit

Sub PrintAllPo()

Dim Flg As Integer

Flg = MsgBox(Prompt:="開いているファイル全て印刷します。" & vbLf & "よろしいですか？", Buttons:=vbOKCancel)

If Flg <> 1 Then Exit Sub

Dim Book As Variant

For Each Book In Workbooks

    If Book.Name Like "発注書*" Then
        Book.PrintOut
    End If

Next

End Sub

Sub CloseAllPo()

Dim Flg As Integer

Flg = MsgBox(Prompt:="開いているファイル全て閉じます。" & vbLf & "よろしいですか？", Buttons:=vbOKCancel)

If Flg <> 1 Then Exit Sub

Dim Book As Variant

For Each Book In Workbooks

    If Book.Name Like "発注書*" Then
        Book.Close SaveChanges:=False
    End If

Next

End Sub

