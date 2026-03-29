Attribute VB_Name = "ModStep10"
Option Explicit

' ============================================================
' �X�e�b�v�I: ���ёւ�
'
' �S���f�����ʂ̗D�揇�ʁi��ʃL�[����j:
'   1. MODEL(U��): ����
'   2. �@�B�i��(H��): ���� �������e�n�̑�2�L�[�AV8/V9�͋󗓂Ȃ̂Ŏ����X�L�b�v
'   3. ���^ss�o�ד�(N��): ����
'   4. �����w�����s��(M��): ����
'   5. KP-No(R��): ����
'   6. ����(I��): �~��
'   7. �q�於(C��): ����
'   8. ���Y�v��No(B��): ����
'
' �� V8/V9�̋@�B�i�Ԃ͋󗓂̂��߁A���� V8/V9�ƃ����e��MODEL�ŕ�������A
'   V8/V9�͏o�ד������s����KP-No���������q�於���v��No�̏��ɂȂ�
' ============================================================
Public Sub Step10_���ёւ�(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        Call ���O����("Step10_���ёւ�", "���", "�f�[�^�Ȃ��A�X�L�b�v")
        Exit Sub
    End If

    Dim lastCol As Long
    lastCol = ws.UsedRange.Columns.Count

    Dim sortRange As Range
    Set sortRange = ws.Range(ws.Cells(g_DataStartRow, 1), ws.Cells(lastRow, lastCol))

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Columns(g_ColModel),         Order:=xlAscending   ' 1. MODEL
        .SortFields.Add Key:=ws.Columns(g_ColKikiHinban),    Order:=xlAscending   ' 2. �@�B�i��
        .SortFields.Add Key:=ws.Columns(g_ColShukkaDate),    Order:=xlAscending   ' 3. �o�ד�
        .SortFields.Add Key:=ws.Columns(g_ColJunjoHakkoDate),Order:=xlAscending   ' 4. ���s��
        .SortFields.Add Key:=ws.Columns(g_ColKPNo),          Order:=xlAscending   ' 5. KP-No
        .SortFields.Add Key:=ws.Columns(g_ColZokusei),       Order:=xlDescending  ' 6. �����i�~���j
        .SortFields.Add Key:=ws.Columns(g_ColKyakusakiName), Order:=xlAscending   ' 7. �q�於
        .SortFields.Add Key:=ws.Columns(g_ColSeisanNo),      Order:=xlAscending   ' 8. ���Y�v��No
        .SetRange sortRange
        .Header = xlNo
        .Apply
    End With

    Call ���O����("Step10_���ёւ�", "����", "���ёւ������i" & lastRow - g_DataStartRow + 1 & "�s�j")
End Sub
