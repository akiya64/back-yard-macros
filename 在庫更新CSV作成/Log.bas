Attribute VB_Name = "Log"
Option Explicit

Sub ApendProcessingTime(Sec As Long)

Dim TimeRange As Range
Set TimeRange = Sheet1.Range("A141")

Dim DateCell As Range
Set DateCell = TimeRange.End(xlDown).Offset(1, 0)

DateCell.Value = Format(Date, "MŒŽD“ú")

Dim TimeCell As Range
Set TimeCell = DateCell.Offset(0, 1)

TimeCell.Value = Sec

End Sub
