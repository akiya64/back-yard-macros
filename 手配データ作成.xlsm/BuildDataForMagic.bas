Attribute VB_Name = "BuildDataForMagic"
Option Explicit
Const DEFAULT_OPERATOR_CODE As Integer = 329

Dim CurrentOperator As Integer

Type Purchase
'手配数量入力シート1行分に相当するユーザー定義型

    Code As String
    ProductName As String
    
    VendorCode As Long
    VendorName As String
    
    UnitCost As Long
    
    PurchaseQuantity As Long
    RequireQuantity As Long
    
    RequireMallCount As String
    
    WarehouseNumber As Integer
    
    IsPickup As Integer
    
    IsHold As Boolean
    HoldReason As String

End Type

Sub BuildPurcahseData()
'発注数量決定シートを元に、ファイルを書き出す
'「発注システム用データ出力」ボタンで呼び出される

'各シートが空かチェック
Dim Sh As Variant
For Each Sh In Array(Worksheets("Magic一括登録"), Worksheets("Magic手入力用"), Worksheets("発注商品リスト"), Worksheets("保留"))
    Call PrepareSheet(Sh)
Next

'担当者コード欄が空か数字3ケタでなければ、担当者コード329でデータ作成
Dim InputOpCode As Variant
InputOpCode = InputSp.Range("G3").Value

If Not InputOpCode Like "###" Then
    MsgBox Prompt:="手入力シートで設定された担当者コードの形式が合いません。" & vbCrLf & "担当者コード329でデータ作成します。", _
            Buttons:=vbExclamation
Else
    CurrentOperator = InputOpCode
End If

'データ出力用のシートに、1行ずつコピー
Worksheets("手配数量入力シート").Activate

Dim i As Long
For i = 2 To Range("A1").End(xlDown).Row

    Dim CurrentPurchase As Purchase
    CurrentPurchase = ReadPurchase(i)
    
    If CurrentPurchase.IsHold Then
        Call WriteHoldList(CurrentPurchase)
    Else
        
        Call WriteBackupSheet(CurrentPurchase)
        
        If CurrentPurchase.Code Like "######" Then
            Call WriteMagicTxt(CurrentPurchase)
        Else
            Call WriteMagicManualInput(CurrentPurchase)
        End If
        
    End If

Next

Worksheets("Magic一括登録").Columns("A:E").AutoFit
Worksheets("Magic手入力用").Columns("A:I").AutoFit

'出力用シートをファイルとして保存していく
'発注システム用テキストファイルを出力
Call PutTxtFileForMagic

'バックアップを保存
ThisWorkbook.Worksheets("発注商品リスト").Copy

With ActiveSheet
    .Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
    .Rows(1).Insert
    .Range("B1").Value = "ﾊﾞｯｸｱｯﾌﾟ日時 : " & Format(Date, "YYYY/MM/dd") & " " & Format(Time, "hh:mm:ss")
End With

ActiveWorkbook.SaveAs FileName:="\\Server02\商品部\ネット販売関連\発注関連\半自動発注バックアップ\BU" & Format(Date, "YYYYMMDD") & "-" & Format(Now, "hhmmss") & ".xlsx"
ActiveWorkbook.Close

'保留を保存
Worksheets("保留").Copy

Dim FileName As String
FileName = "\保留" & Format(Date, "MMdd") & ".xlsx"

If Dir(ThisWorkbook.Path & FileName) <> "" Then
    FileName = Replace(FileName, Format(Date, "MMdd"), Format(Date, "MMdd") & "-" & Format(Time, "hhmm"))
End If

ActiveWorkbook.SaveAs FileName:=ThisWorkbook.Path & FileName

'c保留へ追記してから閉じる
Call AppendHoldPurWokbook(ActiveWorkbook)
ActiveWorkbook.Close

'Magic入力用Excelファイルを保存
Sheets(Array("Magic一括登録", "Magic手入力用")).Copy

FileName = "\Magic入力データ" & Format(Date, "MMdd") & ".xlsx"

If Dir(ThisWorkbook.Path & FileName) <> "" Then
    FileName = Replace(FileName, Format(Date, "MMdd"), Format(Date, "MMdd") & "-" & Format(Time, "hhmm"))
End If

ActiveWorkbook.SaveAs FileName:=ThisWorkbook.Path & FileName
ActiveWorkbook.Close

'ファイル出力完了、このブックを保存
ThisWorkbook.Save

Application.DisplayAlerts = True

MsgBox Prompt:="ファイル保存が完了しました。" & vbLf & "続けて返信FAXリスト転記を行います。", Buttons:=vbInformation

'返信FAXリストへ追記
Call AppendRefaxList

MsgBox Prompt:="処理が完了しました。", Buttons:=vbInformation

End Sub

Private Function ReadPurchase(ByVal Row As Long) As Purchase
'手配数量入力シートから1行を1変数に読み込む

Dim TmpPur As Purchase

With TmpPur
    .Code = Cells(Row, 7).Value  '発注時の商品コード、JANか6ケタ
    .ProductName = Cells(Row, 8).Value '商品名、JAN手配分のみ必須
    
    .VendorCode = Cells(Row, 4).Value '手配先コード
    .VendorName = Cells(Row, 5).Value '手配先名称
     
    .WarehouseNumber = IIf(Cells(Row, 6).Value = "V", "187", "180")  '倉庫番号

    .RequireQuantity = Cells(Row, 9).Value '手配依頼数量
    
    .RequireMallCount = Cells(Row, 6).Value 'モール別の依頼件数

    '発注保留に該当するかチェックして、フラグを立てる

    '保留として数量を注意書きで上書きしているか？
    If IsNumeric(Cells(Row, 1).Value) Then
        TmpPur.PurchaseQuantity = Cells(Row, 1).Value
    Else
        TmpPur.HoldReason = Cells(Row, 1).Value
        TmpPur.IsHold = True
    End If
    
    .UnitCost = Cells(Row, 10).Value
    
    '引取で手配するか
    If Trim(Cells(Row, 11).Value) = "" Then
        .IsPickup = 2
    Else
        .IsPickup = Cells(Row, 11).Value
    End If
        
End With

ReadPurchase = TmpPur

End Function

Private Sub WriteMagicTxt(ByRef Purchase As Purchase)
    
    Dim WriteRow As Long, TargetSheet As Worksheet, Record As Variant
    
    With Purchase
        Record = Array( _
                    .WarehouseNumber, _
                    .Code, _
                    .PurchaseQuantity, _
                    .IsPickup, _
                    CurrentOperator _
                    )
    End With
    
    With Worksheets("Magic一括登録")
        
        WriteRow = IIf(.Range("A2").Value = "", 2, .Range("A1").End(xlDown).Row + 1)
        
        .Cells(WriteRow, 2).NumberFormatLocal = String(9, "0")
        .Cells(WriteRow, 3).NumberFormatLocal = String(8, "0")
    
        .Cells(WriteRow, 1).Resize(1, UBound(Record) + 1).Value = Record
    End With
    
End Sub

Private Sub WriteMagicManualInput(ByRef Purchase As Purchase)

    Dim WriteRow As Long, TargetSheet As Worksheet, Record As Variant
    
    With Purchase
        Record = Array( _
                    .WarehouseNumber, _
                    .VendorCode, _
                    .VendorName, _
                    .Code, _
                    .ProductName, _
                    .PurchaseQuantity, _
                    .UnitCost, _
                    .IsPickup, _
                    CurrentOperator _
                    )
    End With
    
    With Worksheets("Magic手入力用")
        WriteRow = IIf(.Range("A2").Value = "", 2, .Range("A1").End(xlDown).Row + 1)
    
        .Cells(WriteRow, 4).NumberFormatLocal = "@"
        
        .Cells(WriteRow, 1).Resize(1, UBound(Record) + 1).Value = Record
    End With
    
End Sub

Private Sub WriteHoldList(ByRef Purchase As Purchase)

    Dim WriteRow As Long, TargetSheet As Worksheet, Record As Variant
    
    With Purchase
        Record = Array( _
                    .HoldReason, _
                    .WarehouseNumber, _
                    .VendorName, _
                    .RequireMallCount, _
                    Date, _
                    .Code, _
                    .RequireQuantity, _
                    .ProductName)
    End With
    
    With Worksheets("保留")
        WriteRow = IIf(.Range("A2").Value = "", 2, .Range("A1").End(xlDown).Row + 1)
    
        .Cells(WriteRow, 5).NumberFormatLocal = "Mdd"
        .Cells(WriteRow, 1).Resize(1, UBound(Record) + 1).Value = Record
    End With
    
End Sub

Private Sub WriteBackupSheet(ByRef Purchase As Purchase)

    Dim WriteRow As Long, TargetSheet As Worksheet, Record As Variant
    
    With Purchase
        Record = Array( _
                    .WarehouseNumber, _
                    .VendorName, _
                    .RequireMallCount, _
                    Date, _
                    .Code, _
                    .ProductName, _
                    .PurchaseQuantity _
                    )
    End With
    
    With Worksheets("発注商品リスト")
        WriteRow = IIf(.Range("A2").Value = "", 2, .Range("A1").End(xlDown).Row + 1)
    
        .Cells(WriteRow, 4).NumberFormatLocal = "Mdd"
        
        .Cells(WriteRow, 1).Resize(1, UBound(Record) + 1).Value = Record
    End With
    
End Sub

Sub PutTxtFileForMagic()

'Magic一括登録シートを新規ブックにコピー、拡張子.txt、カンマ区切り、ヘッダー無しで保存

Worksheets("Magic一括登録").Copy
ActiveSheet.Rows(1).Delete

Dim FileName As String
FileName = "\Magic登録用" & Format(Date, "MMdd") & ".txt"

If Dir(ThisWorkbook.Path & FileName) <> "" Then
    FileName = Replace(FileName, Format(Date, "MMdd"), Format(Date, "MMdd") & "-" & Format(Time, "hhmm"))
End If

Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs FileName:=ThisWorkbook.Path & FileName, FileFormat:=xlCSV
    ActiveWorkbook.Close

End Sub

