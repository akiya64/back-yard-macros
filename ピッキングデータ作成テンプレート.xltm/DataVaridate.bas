Attribute VB_Name = "DataVaridate"
Option Explicit

Function ModifyOrderSheet(Optional var As Variant) As Boolean
'アドインで取得したデータを転記

OrderSheet.Activate

Dim LastRow As Long, i As Long
LastRow = Range("A1").SpecialCells(xlCellTypeLastCell).Row

For i = 2 To LastRow
    
    'ロケーション修正
    Cells(i, 8).Value = CutOffUnlocation(Cells(i, 18).Value)

Next

'Columns("M:AB").Delete

End Function

Private Function CutOffUnlocation(Location As String) As String
' 正規表現でロケーション[0-0-0-0][0- -0- - ][1-0-0-0-0]などを削除して返します。

Dim Reg As New RegExp

Reg.Global = True
Reg.Pattern = "\[[0-3|\s]\-[0-3|\s]\-[0|\s]\-[0|\s]\-[0|\s]\]"

CutOffUnlocation = Reg.Replace(Location, "")

End Function
