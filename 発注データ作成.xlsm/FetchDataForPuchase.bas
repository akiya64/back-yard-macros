Attribute VB_Name = "FetchDataForPuchase"
Option Explicit
Sub FetchSyokonData()
'�������i�}�X�^�[�̃f�[�^�擾
'���ۂɓǂ݂ɍs���e�[�u���́A�󒍏ڍ׊m�F�p�ɖ������v���P�[�V��������Server3��EC�ۗp�}�X�^

'�ڑ��̂��߂̃I�u�W�F�N�g���`�ADB�ڑ��ݒ��Z�b�g
Dim DbCnn As New ADODB.Connection
Dim DbCmd  As New ADODB.Command
Dim DbRs As New ADODB.Recordset

DbCnn.ConnectionTimeout = 0
DbCnn.Open "PROVIDER=SQLOLEDB;Server=;Database=;UID=;PWD=;"
DbCmd.CommandTimeout = 180
Set DbCmd.ActiveConnection = DbCnn

'���i�R�[�h�̃����W��Z�b�g�A1�Z������SQL���s
Dim CodeRange As Range, r As Range
Set CodeRange = Range(Cells(2, 7), Cells(2, 7).End(xlDown))

For Each r In CodeRange
    Dim sql As String, Code As String
    Code = r.Value
    
    If Code Like String(13, "#") Then
        
    Else
    
    End If
    
    sql = "SELECT ���i�R�[�h, �戵�敪, ���b�g��, �d������, �d����, �d����}�X�^.�d���旪�� " & _
          "FROM ���i�}�X�^ JOIN �d����}�X�^ ON ���i�}�X�^.�d���� = �d����}�X�^.�d����R�[�h " & _
          "WHERE ���i�R�[�h = " & Code & "OR JAN�R�[�h = '" & Code & "'"
    
    Set DbRs = DbCnn.Execute(sql)

    If Not DbRs.EOF Then
        Cells(r.Row, 3).Value = DbRs("���b�g��")
        Cells(r.Row, 4).Value = DbRs("�d����")
        Cells(r.Row, 5).Value = DbRs("�d���旪��")
        Cells(r.Row, 10).Value = DbRs("�d������")
        Cells(r.Row, 2).Value = GetKubun(DbRs("�戵�敪"))
        
        'JAN�󒍕��̏��i�R�[�h�u���A���Amazon���p
        If Len(r.Value) > 6 Then
            r.NumberFormatLocal = "@"
            r.Value = IIf(Len(DbRs("���i�R�[�h")) = 5, "0" & DbRs("���i�R�[�h"), DbRs("���i�R�[�h"))
        End If
    
    End If

Next

End Sub

Sub FetchExcellForPurchase()
'�����p���i���̃f�[�^�擾

Dim DataBook As Workbook, DataSheet As Worksheet, PurDataCodeRange As Range, PurDataJanRange As Range
Set DataSheet = OpenPurDataBook().Worksheets("���i���")
Set PurDataJanRange = DataSheet.Range(Cells(1, 1), Cells(DataSheet.UsedRange.Rows.Count, 1))
Set PurDataCodeRange = DataSheet.Range(Cells(1, 2), Cells(DataSheet.UsedRange.Rows.Count, 2))

ThisWorkbook.Activate
Dim CodeRange As Range, r As Range
Set CodeRange = Range(Cells(2, 7), Cells(2, 7).End(xlDown))

For Each r In CodeRange

    Dim Code As String, HitRow As Double
        
    Code = r.Value

    On Error Resume Next
        HitRow = WorksheetFunction.Match(Code, PurDataCodeRange, 0)

        If Err Then
            Err.Clear
            HitRow = WorksheetFunction.Match(Code, PurDataJanRange, 0)
            
            If Err And IsEmpty(Cells(r.Row, 4).Value) Then '�d����R�[�h�������ɂ�G�N�Z���ɂ�Ȃ���΁A�����ł��Ȃ��̂Œ��ӏ���������
                Cells(r.Row, 2).Value = "�����p���i��� �f�[�^�Ȃ�"
                GoTo Continue
            End If
        
        End If
    
    On Error GoTo 0
        
    '��z�����ӁA���[�J�[���b�g��
    Cells(r.Row, 2).Value = Cells(r.Row, 2).Value & DataSheet.Cells(HitRow, 35).Value '��z������
    Cells(r.Row, 11).Value = DataSheet.Cells(HitRow, 5).Value '�����p���i���̃��b�g��
    
    '�d����R�[�h�A�����A�d���於��6�P�^�ɂȂ����̂ݓ����
    If IsEmpty(Cells(r.Row, 4).Value) Then
    
        Cells(r.Row, 4).Value = DataSheet.Cells(HitRow, 32).Value '�d����R�[�h
        Cells(r.Row, 5).Value = DataSheet.Cells(HitRow, 4).Value '�d���於
        Cells(r.Row, 10).Value = DataSheet.Cells(HitRow, 13).Value '����

    End If

Continue:

Next

End Sub

Sub FetchExcellJanInventory()
'�I�Ȃ��݌ɕ\�f�[�^�̊m�F


End Sub

Sub CalcPurchaseQuantity()
Dim i As Long
For i = 2 To 60
    Dim Rot As Double, Qty As Long, RequestQty As Double
    If IsEmpty(Cells(i, 11).Value) Then
        Rot = 1
    Else
        Rot = Cells(i, 11).Value
    End If
    
    RequestQty = Cells(i, 9).Value
    
    Qty = WorksheetFunction.Ceiling(Rot, RequestQty)

    Cells(i, 1).Value = Qty
Next

End Sub

Private Function GetKubun(ByVal KubunCode As Integer) As String
Dim tmp As String

Select Case KubunCode
    Case 3
        tmp = "����:�̔����~"
    Case 7
        tmp = "����:�݌ɔp��"
    Case 8
        tmp = "����:�݌ɏ���"
    Case 9
        tmp = "����:���[�J�[�p��"
    Case Else
        tmp = ""
End Select

GetKubun = tmp

End Function

Private Function OpenPurDataBook() As Workbook

'�����p���i���̃t�@�C����J���܂��B
'���s���̃G�N�Z���Ŕ����p���i���̃t�@�C����J���Ă���΁A���̃��[�N�u�b�N��Ԃ��܂��B

Const PUR_DATA_EXCELL_PATH As String = "\\Server02\���i��\�l�b�g�̔��֘A\�����֘A\�����p���i���.xlsm"

Dim WorkBookName As String
WorkBookName = Dir(PUR_DATA_EXCELL_PATH)

Dim wb As Workbook

For Each wb In Workbooks
    
    If wb.Name = WorkBookName Then
        wb.Activate
        GoTo ret
    
    End If

Next

Set wb = Workbooks.Open(PUR_DATA_EXCELL_PATH)

ret:

Set OpenPurDataBook = wb

End Function
