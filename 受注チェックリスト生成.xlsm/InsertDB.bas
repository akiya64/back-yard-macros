Attribute VB_Name = "InsertDB"
Option Explicit

Sub CodeI2JAN_E()
'EC変換専用（受注チェックリスト生成.xltm組込み専用ルーチン
'参照設定でADO_Libraryが必須（2.X 系推奨)

Dim Cnn As New ADODB.Connection
Dim Cmd As New ADODB.Command
Dim Rs As New ADODB.Recordset

Dim SQL_W As String
Dim SQL_W1 As String
Dim R As Long
Dim i As Long
Dim j As Long
Dim IC As Long
Dim JC As String
Dim T_D As Variant
Dim MB_C As Variant
Dim ReceiveNo As Long
Dim Old_Status As String

Old_Status = Application.StatusBar

MB_C = MsgBox("JAN変換とDBへの書き込みを開始します", vbOKCancel)

If MB_C = vbOK Then

    Cnn.Open "PROVIDER=SQLOLEDB;Server=server;Database=WINSQL;UID=xxx;PWD=xxx;"
    Cmd.CommandTimeout = 180
    Set Cmd.ActiveConnection = Cnn

    Range("A2").Select
    Selection.End(xlDown).Select
    R = ActiveCell.Row

    ReceiveNo = Range("A2").Value
    Application.StatusBar = "伝票番号重複チェック中・・・"
    Application.ScreenUpdating = False

        For i = 2 To R
            Cells(i, 1).Select
            If Len(Cells(i, 1).Value) <> 0 Then
                IC = Val(Cells(i, 1).Value)
                SQL_W = "SELECT 受注番号 From EC_order_base Where 受注番号 = " & IC & ";"
                Set Rs = Cnn.Execute(SQL_W)
                If Not Rs.EOF Then
                    Cells(i, 15).Value = 1
                End If
            End If
       Next i
    

    SQL_W = "SELECT 商品コード, JANコード FROM 商品マスタ WHERE (商品コード = "
    
    Application.ScreenUpdating = True
    ActiveWorkbook.RefreshAll

    Application.StatusBar = "受注データーをDBに書き込み中・・・"
    
    Application.ScreenUpdating = False

        For i = 2 To R
            Cells(i, 2).Select
            If Len(Cells(i, 2).Value) <> 0 Then
                IC = Val(Cells(i, 2).Value)
                SQL_W1 = SQL_W & IC & ")"
                Set Rs = Cnn.Execute(SQL_W1)
                If Not Rs.EOF() Then
                    JC = Rs(1)
                    Cells(i, 3).Value = JC
                End If
            End If
        Next i

    Application.ScreenUpdating = True
    ActiveWorkbook.RefreshAll

    Range("A2").Select
    
    Application.ScreenUpdating = False

    For i = 2 To R
        If Cells(i, 15).Value <> 1 Then
            T_D = Cells(i, 1).Value & ","
        
            If Cells(i, 2).Value = "" Then
                T_D = T_D & "Null,"
            Else
                T_D = T_D & Cells(i, 2).Value & ","
            End If
        
            T_D = T_D & "'" & Cells(i, 3).Value & "',"
            T_D = T_D & "'" & Cells(i, 4).Value & "',"
            T_D = T_D & Cells(i, 5).Value & ","
            T_D = T_D & "'" & Cells(i, 6).Value & "',"
            T_D = T_D & "'" & Cells(i, 7).Value & "',"
            T_D = T_D & "'" & Cells(i, 8).Value & "',"
            T_D = T_D & Cells(i, 9).Value & ","
            T_D = T_D & "'" & Cells(i, 10).Value & "',"
            T_D = T_D & "'" & Cells(i, 11).Value & "',"
            T_D = T_D & "'" & Cells(i, 12).Value & "',"
            T_D = T_D & "'" & Cells(i, 13).Value & "',"
            T_D = T_D & Cells(i, 14).Value & ","
            T_D = T_D & "0"
            

            SQL_W = "INSERT INTO EC_Order_Base VALUES (" & T_D & ")"

            Cnn.Execute (SQL_W)

            Cells(i, 15).Value = 1

        End If

    Next i
    
    Application.ScreenUpdating = True
    ActiveWorkbook.RefreshAll

    Application.StatusBar = "ハンディ用データーをDBに書き込み中・・・"
    
    Application.ScreenUpdating = False

    SQL_W = "INSERT INTO EC_order ( 納品書番号, 商品コード, バーコード, 商品名, 受注数量, モール, 顧客名, フェーズ変更日時, 処理フェーズ, 住所, キャンセル, 受注メモ欄, 受注明細枝番) "
    SQL_W = SQL_W & "SELECT EC_Order_Base.受注番号, EC_Order_Base.商品コード, EC_Order_Base.JANコード, EC_Order_Base.商品名称, EC_Order_Base.数量, EC_Order_Base.納品書区分, EC_Order_Base.届け先名称, GETDATE() AS 式1, 1 AS 式2, 届け先住所, 0, EC_Order_Base.受注メモ欄, EC_Order_Base.受注明細枝番 "
    SQL_W = SQL_W & "FROM EC_Order_Base WHERE EC_Order_Base.転送 = 0;"

    Cnn.Execute (SQL_W)
    
    SQL_W = "UPDATE EC_Order_Base SET 転送 = 1 WHERE 転送 = 0;"
    Cnn.Execute (SQL_W)
    
    Application.ScreenUpdating = True
    ActiveWorkbook.RefreshAll


    MsgBox "終了", vbInformation

    Cnn.Close

Else

MsgBox "キャンセルしました", vbCritical

End If

Set Cnn = Nothing
Set Cmd = Nothing
Set Rs = Nothing

Application.StatusBar = Old_Status


End Sub

