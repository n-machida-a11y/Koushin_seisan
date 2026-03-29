Attribute VB_Name = "ModMain"
Option Explicit

' ============================================================
' ���C�����s: Phase 1-A�i�X�e�b�v�D�`�I�j
' �u�ݒ�v�V�[�g�̃{�^������Ăяo��
' ============================================================
Public Sub ���C�����s()
    ' �J�n�m�F�_�C�A���O
    Dim ans As VbMsgBoxResult
    ans = MsgBox("���Y�v�掩�����iPhase 1-A�j���J�n���܂��B" & vbCrLf & vbCrLf & _
                 "�y���O�m�F�z" & vbCrLf & _
                 "�EBH�v�����̏o�̓t�@�C���ixlsx�j�� input�t�H���_�ɒu���Ă�������" & vbCrLf & _
                 "�E�ݒ�V�[�g�̃t�H���_�p�X�����������Ƃ��m�F���Ă�������" & vbCrLf & vbCrLf & _
                 "���s���܂����H", vbYesNo + vbQuestion, "���Y�v�掩����")
    If ans = vbNo Then Exit Sub

    ' �ݒ�ǂݍ���
    Call �ݒ�ǂݍ���()

    ' ���H�Ώۃt�@�C�����J��
    Dim targetWb As Workbook
    Set targetWb = �Ώۃt�@�C�����J��()
    If targetWb Is Nothing Then Exit Sub

    ' �ΏۃV�[�g���擾
    Dim targetWs As Worksheet
    On Error Resume Next
    Set targetWs = targetWb.Sheets(g_TargetSheetName)
    On Error GoTo 0
    If targetWs Is Nothing Then
        MsgBox "�V�[�g�u" & g_TargetSheetName & "�v��������܂���B" & vbCrLf & _
               "�ݒ�V�[�g�́u���H�ΏۃV�[�g���v���m�F���Ă��������B", _
               vbCritical, "�V�[�g��������܂���"
        targetWb.Close SaveChanges:=False
        Exit Sub
    End If

    ' ���O�ɊJ�n���L�^
    Call ���O����("���C�����s", "���", "�����J�n: " & targetWb.Name)

    ' �p�t�H�[�}���X�œK��
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrHandler

    ' ===== �X�e�b�v�D�`�I�����ԂɎ��s =====
    Call Step05_�v�搶�Y�Ώۍ폜(targetWs)
    Call Step06_�o�׍ς݃f�[�^�폜(targetWs)
    Call Step07_�^���⊮(targetWs)
    Call Step08_�v�搶�Y�s�W�J(targetWs)
    Call Step09_���ʃ`�F�b�N(targetWs)
    Call Step10_���ёւ�(targetWs)
    ' =========================================

    On Error GoTo 0

    ' Application��Ԃ̕���
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    ' �㏑���ۑ�
    targetWb.Save

    Call ���O����("���C�����s", "����", "Phase 1-A ��������: " & targetWb.Name)
    MsgBox "�������������܂����B" & vbCrLf & _
           "�u���O�v�V�[�g�ŏ������ʂ��m�F���Ă��������B", vbInformation, "����"
    Exit Sub

ErrHandler:
    ' Application��Ԃ̕���
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    Call ���O����("���C�����s", "�G���[", "�\�����Ȃ��G���[: " & Err.Description)
    MsgBox "�\�����Ȃ��G���[�����������܂����B" & vbCrLf & _
           "�t�@�C���͕ۑ�����Ă��܂���B" & vbCrLf & vbCrLf & _
           "�G���[: " & Err.Description, vbCritical, "�G���["
    targetWb.Close SaveChanges:=False
End Sub

' ============================================================
' input�t�H���_���� xlsx �t�@�C�����J���ĕԂ�
' ��������ꍇ�͍Ō�ɍX�V���ꂽ���̂�I��
' ============================================================
Private Function �Ώۃt�@�C�����J��() As Workbook
    Dim folderPath As String
    folderPath = g_BHPlanFolder
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ' �t�H���_����xlsx���������čŐV�t�@�C�����擾
    Dim fileName As String
    Dim latestFile As String
    Dim latestDate As Date
    Dim dirErrNum As Long
    dirErrNum = 0
    On Error Resume Next
    fileName = Dir(folderPath & "*.xlsx")
    dirErrNum = Err.Number
    On Error GoTo 0
    If dirErrNum <> 0 Then
        MsgBox "�t�H���_�ւ̃A�N�Z�X�Ɏ��s���܂���(Error " & dirErrNum & ")�B" & vbCrLf & _
               "�t�H���_: " & folderPath & vbCrLf & _
               "�ݒ�V�[�g�́uBH�v�����ۑ��t�H���_�v���m�F���Ă��������B", _
               vbCritical, "�t�H���_�A�N�Z�X�G���["
        Set �Ώۃt�@�C�����J�� = Nothing
        Exit Function
    End If

    Do While fileName <> ""
        Dim fileDate As Date
        fileDate = FileDateTime(folderPath & fileName)
        If fileDate > latestDate Then
            latestDate = fileDate
            latestFile = fileName
        End If
        fileName = Dir()
    Loop

    If latestFile = "" Then
        MsgBox "input�t�H���_��xlsx�t�@�C����������܂���B" & vbCrLf & _
               "�t�H���_: " & folderPath & vbCrLf & vbCrLf & _
               "BH�v�����̏o�̓t�@�C�����t�H���_�ɔz�u���Ă���Ď��s���Ă��������B", _
               vbCritical, "�t�@�C���Ȃ�"
        Set �Ώۃt�@�C�����J�� = Nothing
        Exit Function
    End If

    Set �Ώۃt�@�C�����J�� = Workbooks.Open(folderPath & latestFile)
End Function
