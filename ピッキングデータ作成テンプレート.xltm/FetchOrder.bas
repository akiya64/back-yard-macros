Attribute VB_Name = "FetchOrder"
Option Explicit

'���ׂƒ����w�b�_�[�̂����t�H���_���w���A�Ō��K��\�}�[�N
Const CSV_PATH As String = "C:\Users\mos10\Desktop\���t�[\"
Const ALTER_CSV_PATH As String = "\\MOS10\Users\mos10\Desktop\���t�[\"

Sub �󒍃t�@�C���Ǎ�()

OrderSheet.Activate

If Not Range("B2").Value = "" Then
    MsgBox "�f�[�^�擾�ςł��B"
    End
End If

Dim LineBuf As Variant

'�t�@�C�������I�u�W�F�N�g����
Dim FSO As New FileSystemObject

' Meisai.csv��tyumon_H.csv��CSV�t�@�C���̃p�X���Z�b�g
Dim MeisaiPath As String, TyumonhPath As String

If FSO.FileExists(CSV_PATH & "Meisai.csv") Then

    MeisaiPath = CSV_PATH & "Meisai.csv"
    TyumonhPath = CSV_PATH & "tyumon_H.csv"

ElseIf FSO.FileExists(ALTER_CSV_PATH & "Meisai.csv") Then
   
    MeisaiPath = ALTER_CSV_PATH & "Meisai.csv"
    TyumonhPath = ALTER_CSV_PATH & "tyumon_H.csv"

Else
    
    MsgBox "meisai.csv �t�@�C���Ȃ�"
    End

End If

Call ReadMeisai(MeisaiPath)

Call ReadTyumonH(TyumonhPath)

OrderSheet.Shapes(1).Delete

'�A�h�C���p�̍s�E�� �\��
Dim LastRow As Long
LastRow = Range("D1").SpecialCells(xlCellTypeLastCell).Row

Range("L1").Value = "�A�h�C���w�� �䒠�F9998"
Range("L2:O2") = Array(2, 4, LastRow, 12)

MsgBox "�A�h�C�������s���ĉ������B"

End Sub

Private Sub ReadMeisai(Path As String)

'Meisai.CSV��OrderSheet=�����ꗗ�ɒǋL����

Dim FSO As Object
Set FSO = New FileSystemObject

Dim TS As Textstream
Set TS = FSO.OpenTextFile(Path, ForReading)

Dim i As Long
i = 1 '���ڍs�͏o�͂��Ȃ��̂ŁAi��1�s�ڂ����J�n
    
Do Until TS.AtEndOfStream
    
' �s�������o���ĕK�v�ȍ��ڂ݂̂��z���ɓ��꒼��
    Dim LineBuf As Variant
    LineBuf = Split(TS.ReadLine, """,""")
       
    Dim j As Integer
    For j = 0 To UBound(LineBuf)
        LineBuf(j) = Trim(Replace(LineBuf(j), Chr(34), "")) 'chr(34)�� " [���p���d���p��]
    
    Next
    
    '���[�v�����ڂł̓w�b�_�[�Ȃ̂ŁA�C���N�������g�֔���
    If LineBuf(0) = "Order ID" Then GoTo increment
    
    'CSV���w�b�_�[ 0:Order ID/1:Line ID/2:Quantity/3:Product Code/4:Description/5:Option Name/6:Option Value/7:Unit Price/
        
    ':ToDo ���������V�[�g�A�Z���̒l�̂������̂ŕ������������������������Ȃ��B
    With Worksheets("�󒍃f�[�^�V�[�g")
        .Range("A" & i).Value = LineBuf(0)
        .Range("C" & i).Value = LineBuf(1)
        
        .Range("C" & i).NumberFormatLocal = "@"
        .Range("C" & i).Value = LineBuf(3)
        
        .Range("D" & i).NumberFormatLocal = "@"
        .Range("D" & i).Value = LineBuf(3)
        
        .Range("E" & i).Value = LineBuf(4)
        .Range("F" & i).Value = LineBuf(2)
        .Range("G" & i).Value = LineBuf(7)
        
        'Yahoo!�o�^�R�[�h���`�F�b�N
        '�Z�b�g���� 7777�n�܂�
        Dim YahooCode As String
        YahooCode = .Range("D" & i).Value
        
        If YahooCode Like "7777*" Then
            
            Call SetParser.ParseItems(.Range("D" & i))
            
            'ParseItems�ōs���}���������̂ŁA�s�J�E���^���Z�b�g������
            i = OrderSheet.Range("A1").CurrentRegion.Rows.Count
            
        
        End If
    
        '�P�́��Z�b�g���� �n�C�t���܂ރR�[�h�Ȃ番���\���`�F�b�N
        
        If YahooCode Like "*-*" Then
        
            Call SetParser.ParseScalingSet(.Range("D" & i))
        
        End If
    
        'D�����A�h�C���p��6�P�^�ɏC��
        
        If YahooCode Like "#####" Then
                    
            .Range("D" & i).NumberFormatLocal = "@"
            .Range("D" & i).Value = "0" & YahooCode
        
        End If
    
    End With
    
increment:
    i = i + 1

Loop

TS.Close

End Sub

Private Sub ReadTyumonH(Path As String)

Dim FSO As Object
Set FSO = New FileSystemObject

Dim TS As Textstream
Set TS = FSO.OpenTextFile(Path, ForReading)

'�Ǎ��ϒ����ԍ��̃����W���Z�b�g�AA1����A���̔ԍ������ŏI�Z���܂�
Dim LoadedOrderRange As Range
Set LoadedOrderRange = OrderSheet.Cells(1, 1).Resize(OrderSheet.Cells(1, 1).SpecialCells(xlCellTypeLastCell).Row, 1)

Do Until TS.AtEndOfStream
    
' �s�������o���ĕK�v�ȍ��ڂ݂̂��z���ɓ��꒼��
    Dim LineBuf As Variant
    LineBuf = Split(TS.ReadLine, """,""")
    
    '0=1����=�����ԍ��A�����Җ��A�v�]�A���ϕ��@�A�N�[�|���l����
    Dim Order As Variant
    Order = Array(LineBuf(0), LineBuf(5), LineBuf(36), LineBuf(34), LineBuf(43))
        
    Dim j As Integer
    For j = 0 To UBound(Order)
        Order(j) = Trim(Replace(Order(j), Chr(34), "")) 'chr(34)�� " [���p���d���p��]
    
    Next

    '�����ԍ��̍s�𒲂ׂ�
    '�����ԍ���Dobule�^�œ����Ă����BCSV��String�^�AMatch�֐��̕Ԓl��Double�^
    
    Dim FindRow As Double
    
    On Error Resume Next
        
        FindRow = WorksheetFunction.Match(CDbl(Order(0)), LoadedOrderRange, 0)
        
        If Err Then
            Err.Clear
            GoTo Continue
        End If
    
    On Error GoTo 0
        
    Dim i As Long
    i = 0
    
    '�����Җ����L�� �I�t�Z�b�g���A�Y�������ԍ��̑S�Ă̍s�֋L��
    Do While Range("A" & FindRow).Offset(i, 0).Value = CDbl(Order(0))
        
        Range("A" & FindRow).Offset(i, 1).Value = LineBuf(5)
        i = i + 1
    
    Loop
    
    '���l���֒ǋL �N�[�|�����p���������E���s�U���E���t�[�}�l�[���� �m�F����
    Dim tmp As String
    tmp = ""
    
    If Order(3) = "payment_d1" And Order(4) < 0 Then tmp = "������ �N�[�|�����p "
    If Order(3) = "payment_b1" Then tmp = tmp & "���s�U��"
    If Order(3) = "payment_a16" Then tmp = tmp & "Yahoo!�}�l�[����"
    
    Range("K" & FindRow).Value = tmp 'tmp���Z���ɏ����߂�
        
Continue:
    
Loop

' �I�u�W�F�N�g���j��
TS.Close
Set TS = Nothing
Set FSO = Nothing

End Sub

