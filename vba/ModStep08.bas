Attribute VB_Name = "ModStep08"
Option Explicit

' ============================================================
' �X�e�b�v�G: �v�搶�Y�s�W�J�i1��1�s���j
'
' �Ώ�: F��i�@�햼�j�Ɂu�v�搶�Y�v���܂ލs ����
'       N��i�o�ד��j�������`3�����ȓ�
'
' ����: L��i���ʁj�̐������s���R�s�[���ēW�J���A
'       B��i���Y�v��No�j������ -01,-02... �ƘA�Ԃ�t�^����
' ============================================================
Public Sub Step08_�v�搶�Y�s�W�J(ws As Worksheet)
    Dim months3Later As Date
    Dim expandedCount As Long
    expandedCount = 0
    months3Later = DateSerial(Year(g_BaseDate), Month(g_BaseDate) + 3, Day(g_BaseDate))

    ' �����珈�����邱�Ƃōs�}����̃C���f�b�N�X�Y����h��
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = lastRow To g_DataStartRow Step -1
        Dim kishuName As String
        kishuName = Trim(CStr(ws.Cells(i, g_ColKishuName).Value))
        If InStr(kishuName, "�v�搶�Y") = 0 Then GoTo NextRow

        Dim shukkaDate As Variant
        shukkaDate = ws.Cells(i, g_ColShukkaDate).Value
        If IsEmpty(shukkaDate) Or CStr(shukkaDate) = "" Then GoTo NextRow
        If Not IsDate(shukkaDate) Then GoTo NextRow
        If CDate(shukkaDate) > months3Later Then GoTo NextRow

        Dim rawSuryo As Variant
        rawSuryo = ws.Cells(i, g_ColSuryo).Value
        If IsEmpty(rawSuryo) Or Not IsNumeric(rawSuryo) Then GoTo NextRow
        Dim suryo As Long
        suryo = CLng(rawSuryo)
        If suryo <= 1 Then GoTo NextRow

        ' ���̐��Y�v��No���擾
        Dim baseNo As String
        baseNo = Trim(CStr(ws.Cells(i, g_ColSeisanNo).Value))

        ' Phase 1: suryo-1 �s��}�����ăR�s�[
        Dim j As Long
        For j = 1 To suryo - 1
            ws.Rows(i + 1).Insert Shift:=xlDown
            ws.Rows(i).Copy ws.Rows(i + 1)
        Next j
        ' Phase 2: �A�ԕt�^�i�s i ���� i+suryo-1 ���ׂăR�s�[�ς݁j
        For j = 1 To suryo
            ws.Cells(i + j - 1, g_ColSeisanNo).Value = baseNo & "-" & Format(j, "00")
            ws.Cells(i + j - 1, g_ColSuryo).Value = 1
        Next j

        expandedCount = expandedCount + 1
NextRow:
    Next i

    Call ���O����("Step08_�v�搶�Y�s�W�J", "����", expandedCount & "���̍s�W�J�����{���܂���")
End Sub
