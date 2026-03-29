Attribute VB_Name = "ModError"
Option Explicit

' ============================================================
' ������~�G���[
' �Y���s�����F�n�C���C�g �� ���O�L�^ �� �|�b�v�A�b�v�\�� �� End �ŏ�����~
' ws: ���H�ΏۃV�[�g, rowNum: ���̂���s�ԍ�, message: �\�����b�Z�[�W
' ============================================================
Public Sub ������~�G���[(ws As Worksheet, rowNum As Long, message As String)
    ' �Y���s�����F�Ńn�C���C�g
    ws.Rows(rowNum).Interior.Color = RGB(255, 255, 0)

    ' ���O�ɋL�^
    Call ���O����("�G���[���o", "�G���[", "�s" & rowNum & ": " & message)

    ' �|�b�v�A�b�v�\��
    MsgBox "�y������~�z" & vbCrLf & vbCrLf & _
           message & vbCrLf & vbCrLf & _
           "�s�ԍ�: " & rowNum & vbCrLf & vbCrLf & _
           "�I�������S���҂ɖ₢���킹��A�f�[�^���C�����čŏ�����Ď��s���Ă��������B", _
           vbCritical, "���Y�v�掩���� - ������~"

    ' Application��Ԃ̕���
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    ' �������I���i�ĊJ�Ȃ��j
    End
End Sub

' ============================================================
' �x�����O�i���s�j
' ���O�ɋL�^����̂݁B�|�b�v�A�b�v�Ȃ��E�����p��
' ============================================================
Public Sub �x�����O(stepName As String, rowNum As Long, message As String)
    Call ���O����(stepName, "�x��", "�s" & rowNum & ": " & message)
End Sub
