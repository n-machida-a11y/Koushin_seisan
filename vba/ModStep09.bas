Attribute VB_Name = "ModStep09"
Option Explicit

' ============================================================
' �X�e�b�v�H: ���ʃ`�F�b�N�iV8/V9 3�����ȓ��j
'
' MODEL�iU��j���uV8�v�܂��́uV9�v�i�����e�ȊO�j��
' N��i�o�ד��j��3�����ȓ��̍s�ɐ���1�ȊO������Ώ�����~�G���[
' ============================================================
Public Sub Step09_���ʃ`�F�b�N(ws As Worksheet)
    Dim months3Later As Date
    months3Later = DateSerial(Year(g_BaseDate), Month(g_BaseDate) + 3, Day(g_BaseDate))

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = g_DataStartRow To lastRow
        Dim model As String
        model = Trim(CStr(ws.Cells(i, g_ColModel).Value))

        ' V8�܂���V9�i�����e�����j�̂݃`�F�b�N
        If model <> "V8" And model <> "V9" Then GoTo NextRow

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
        If suryo <> 1 Then
            Call ������~�G���[(ws, i, _
                "MODEL�u" & model & "�v�Ő��ʂ�1�ł͂���܂���i����=" & suryo & "�j�B" & vbCrLf & _
                "�I�������S���҂ւ̖₢���킹���K�v�ł��B" & vbCrLf & _
                "���Y�v��No: " & ws.Cells(i, g_ColSeisanNo).Value)
        End If
NextRow:
    Next i

    Call ���O����("Step09_���ʃ`�F�b�N", "����", "V8/V9��3�����ȓ����ʃ`�F�b�N����")
End Sub
