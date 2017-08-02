Attribute VB_Name = "InsertFormula"
Option Explicit

Sub 入荷日の算出式を入力()
'算出式が未入力の行を調べる
'A列からK列の商品名列までは、手配データ作成から入力される

ThisWorkbook.Worksheets("納期リスト").Activate

Dim InsertEndCell As Range, InsertStartCell As Range, TargetRange As Range
'最終行のU列のセル
Set InsertEndCell = Cells(Range("A1").End(xlDown).Row, 21)
'X列で式が入っている最終行から一行下のU列のセル
Set InsertStartCell = Cells(InsertEndCell.Row, 24).End(xlUp).Offset(1, -3)

Set TargetRange = Range(InsertStartCell, InsertEndCell)

'セットしたレンジに対して式を入れる
Dim r As Variant

For Each r In TargetRange

    'U列、V列はメーカーシートの入荷に関する文言
    On Error Resume Next
        r.Offset(0, 0).Value = WorksheetFunction.VLookup(Cells(r.Row, 4).Value, Worksheets("メーカー").Range("B3:D1000"), 2, False)
        r.Offset(0, 1).Value = WorksheetFunction.VLookup(Cells(r.Row, 4).Value, Worksheets("メーカー").Range("B3:D1000"), 3, False)
    On Error GoTo 0

    'X , Y列は､W列の日付から入荷日を算出する式を入れる、行番号を＠と置いて置換して式の文字列を作る
    r.Offset(0, 3).Formula = Replace("=IFERROR(VALUE(IF($W@="""",IF($V@=""当日"",$F@,IF($V@=""翌日"",$F@+1,IF($V@=""翌々日"",$T@+2,""""))),$W@+1)),"""")", "@", r.Row)
    r.Offset(0, 4).Formula = Replace("=IF($X@="""","""",IF(MOD($X@,7)=0,$X@+2,IF(MOD($X@,7)=1,$X@+1,$X@)))", "@", r.Row)

Next

End Sub
