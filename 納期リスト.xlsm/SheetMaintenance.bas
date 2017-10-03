Attribute VB_Name = "SheetMaintenance"
Option Explicit

Sub 条件付き書式範囲修正()
'表示の更新処理が設定された範囲全てに走るので、最小限の最終2週間分にしておく。

'日々の処理では手配データ作成から追記するときに実行

'@URL:https://msdn.microsoft.com/ja-jp/library/office/ff837422.aspx (Excell2013以降とあるが、2010でも動く）
'@URL:http://www.relief.jp/docs/excel-vba-check-if-has-format-conditions.html

With Worksheets("納期リスト")
    .Activate

    Dim WeekBeforeLastRow As Long, Endrow As Long, LatestArrivalRange
    Endrow = .Range("E2").End(xlDown).Row
    
    'T列で14日前の日付の行を取得、土日は手配入力がないため検索条件は「以下」でセット、日付の検索時はダブル型でシリアル化する
    WeekBeforeLastRow = WorksheetFunction.Match(CDbl(DateAdd("d", -14, Date)), .Range(.Cells(1, 5), .Cells(Endrow, 5)), 1) + 1

    Set LatestArrivalRange = Range(Cells(WeekBeforeLastRow, 12), Cells(Endrow, 12))

    .Cells.FormatConditions(1).ModifyAppliesToRange LatestArrivalRange

End With

End Sub

Sub 一ヶ月以前転記()
'手配データ作成から追記するときに実行

'日々の処理では手配データ作成から追記するときに実行

'ETA; Estimated time of arrival 飛行機の到着予定時刻のこと
Dim Endrow As Long, EtaSheet As Worksheet, LogSheet As Worksheet, i As Long

Set EtaSheet = ThisWorkbook.Worksheets("納期リスト")
Set LogSheet = ThisWorkbook.Worksheets("旧")

EtaSheet.Activate

Endrow = EtaSheet.Range("F2").End(xlDown).Row

'最初のデータが1ヶ月以上前なら、旧シートへコピーするために2行目をコピー対象範囲としてをセット
If DateDiff("d", Date, Cells(2, 5).Value) < -30 Then
    Dim TargetRange As Range
    Set TargetRange = Range("A2:M2")
Else
    Exit Sub
End If

'1ヶ月以上前は、納期リストの3行目からRangeを連結していく
i = 2

Do

    Set TargetRange = Union(TargetRange, EtaSheet.Range(Cells(i, 1), Cells(i, 13)))
    i = i + 1

Loop While DateDiff("d", Date, Cells(i, 5).Value) < -30

'1ヶ月以上前の範囲をコピーして削除
TargetRange.Copy
LogSheet.Cells(LogSheet.UsedRange.Rows.Count + 1, 1).PasteSpecial Paste:=xlPasteValues

TargetRange.EntireRow.Delete shift:=xlShiftUp

End Sub
