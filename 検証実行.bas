Attribute VB_Name = "���؎��s"
Option Explicit
Sub �V�[�g�}�b�`��������()
    Dim �n�� As Date, �I�� As Date, �����l As Long, �i�� As Long
    Dim ��������(1 To 10000)
    �n�� = Now
    ���s��.Show vbModeless
    For �����l = 1 To 10000
        ��������(�����l) = �V�[�g�}�b�`(�����l, "A:A")
        ���s��.�v���O���X�o�[�X�V (Int(�����l / 100))
    Next
    Unload ���s��
    �I�� = Now
    MsgBox "�������ԁF" & �I�� - �n��
    Sheets("����").Cells(3, 3) = �I�� - �n��
End Sub
Sub �z��}�b�`��������()
    Dim �n�� As Date, �I�� As Date, �����l As Long
    Dim ��������(1 To 10000)
    Dim �z��()
    �n�� = Now
    ���s��.Show vbModeless
    With Sheets("�f�[�^")
        �z�� = Range("A1:A10000")
    End With
    For �����l = 1 To 10000
        ��������(�����l) = �z��}�b�`(�����l, �z��)
        ���s��.�v���O���X�o�[�X�V (Int(�����l / 100))
    Next
    Unload ���s��
    �I�� = Now
    MsgBox "�������ԁF" & �I�� - �n��
    Sheets("����").Cells(4, 3) = �I�� - �n��
End Sub
Sub �V�[�g�l�N�X�g��������()
    Dim �n�� As Date, �I�� As Date, �����l As Long, �I�s As Long
    Dim ��������(1 To 10000)
    �n�� = Now
    ���s��.Show vbModeless
    �I�s = Sheets("�f�[�^").Cells(Rows.Count, 1).End(xlUp).Row
    For �����l = 1 To 10000
        ��������(�����l) = �V�[�g�l�N�X�g(�����l, �I�s)
        ���s��.�v���O���X�o�[�X�V (Int(�����l / 100))
    Next
    Unload ���s��
    �I�� = Now
    MsgBox "�������ԁF" & �I�� - �n��
    Sheets("����").Cells(5, 3) = �I�� - �n��
End Sub
Sub �z��l�N�X�g��������()
    Dim �n�� As Date, �I�� As Date, �����l As Long
    Dim ��������(1 To 10000)
    Dim �z��()
    �n�� = Now
    ���s��.Show vbModeless
    With Sheets("�f�[�^")
        �z�� = Range("A1:A10000")
    End With
    For �����l = 1 To 10000
        ��������(�����l) = �z��l�N�X�g(�����l, �z��)
        ���s��.�v���O���X�o�[�X�V (Int(�����l / 100))
    Next
    Unload ���s��
    �I�� = Now
    MsgBox "�������ԁF" & �I�� - �n��
    Sheets("����").Cells(6, 3) = �I�� - �n��
End Sub
Sub �V�[�gVLU��������()
    Dim �n�� As Date, �I�� As Date, �����l As Long
    Dim ��������(1 To 10000)
    �n�� = Now
    ���s��.Show vbModeless
    For �����l = 1 To 10000
        ��������(�����l) = �V�[�gVLU(�����l, "A:A")
        ���s��.�v���O���X�o�[�X�V (Int(�����l / 100))
    Next
    Unload ���s��
    �I�� = Now
    MsgBox "�������ԁF" & �I�� - �n��
    Sheets("����").Cells(7, 3) = �I�� - �n��
End Sub
Sub �z��VLU��������()
    Dim �n�� As Date, �I�� As Date, �����l As Long
    Dim ��������(1 To 10000)
    Dim �z��()
    �n�� = Now
    ���s��.Show vbModeless
    With Sheets("�f�[�^")
        �z�� = Range("A1:A10000")
    End With
    For �����l = 1 To 10000
        ��������(�����l) = �z��VLU(�����l, �z��)
        ���s��.�v���O���X�o�[�X�V (Int(�����l / 100))
    Next
    Unload ���s��
    �I�� = Now
    MsgBox "�������ԁF" & �I�� - �n��
    Sheets("����").Cells(8, 3) = �I�� - �n��
End Sub
