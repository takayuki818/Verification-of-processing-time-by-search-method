Attribute VB_Name = "���؎��s"
Option Explicit
Sub �V�[�g�}�b�`��������()
    Dim �n�� As Date, �����l As Long, �i�� As Long
    Dim ��������(1 To 10000)
    �n�� = Now
    ���s��.Show vbModeless
    Call ���s��.�v���O���X�o�[�X�V(0)
    For �����l = 1 To 10000
        ��������(�����l) = �V�[�g�}�b�`(�����l, "A:A")
        If Int(�����l / 100) - �i�� >= 10 Then
            �i�� = Int(�����l / 100)
            Call ���s��.�v���O���X�o�[�X�V(�i��)
        End If
    Next
    Unload ���s��
    Sheets("����").Cells(3, 3) = Int((Now - �n��) * 24 * 60 * 60 * 100) / 100
End Sub
Sub �z��}�b�`��������()
    Dim �n�� As Date, �����l As Long, �i�� As Long
    Dim ��������(1 To 10000)
    Dim �z��()
    �n�� = Now
    ���s��.Show vbModeless
    Call ���s��.�v���O���X�o�[�X�V(0)
    With Sheets("�f�[�^")
        �z�� = Range("A1:A10000")
    End With
    For �����l = 1 To 10000
        ��������(�����l) = �z��}�b�`(�����l, �z��)
        If Int(�����l / 100) - �i�� >= 10 Then
            �i�� = Int(�����l / 100)
            Call ���s��.�v���O���X�o�[�X�V(�i��)
        End If
    Next
    Unload ���s��
    Sheets("����").Cells(4, 3) = Int((Now - �n��) * 24 * 60 * 60 * 100) / 100
End Sub
Sub �V�[�g�l�N�X�g��������()
    Dim �n�� As Date, �����l As Long, �I�s As Long, �i�� As Long
    Dim ��������(1 To 10000)
    �n�� = Now
    ���s��.Show vbModeless
    Call ���s��.�v���O���X�o�[�X�V(0)
    �I�s = Sheets("�f�[�^").Cells(Rows.Count, 1).End(xlUp).Row
    For �����l = 1 To 10000
        ��������(�����l) = �V�[�g�l�N�X�g(�����l, �I�s)
        If Int(�����l / 100) - �i�� >= 10 Then
            �i�� = Int(�����l / 100)
            Call ���s��.�v���O���X�o�[�X�V(�i��)
        End If
    Next
    Unload ���s��
    Sheets("����").Cells(5, 3) = Int((Now - �n��) * 24 * 60 * 60 * 100) / 100
End Sub
Sub �z��l�N�X�g��������()
    Dim �n�� As Date, �����l As Long, �i�� As Long
    Dim ��������(1 To 10000)
    Dim �z��()
    �n�� = Now
    ���s��.Show vbModeless
    Call ���s��.�v���O���X�o�[�X�V(0)
    With Sheets("�f�[�^")
        �z�� = Range("A1:A10000")
    End With
    For �����l = 1 To 10000
        ��������(�����l) = �z��l�N�X�g(�����l, �z��)
        If Int(�����l / 100) - �i�� >= 10 Then
            �i�� = Int(�����l / 100)
            Call ���s��.�v���O���X�o�[�X�V(�i��)
        End If
    Next
    Unload ���s��
    Sheets("����").Cells(6, 3) = Int((Now - �n��) * 24 * 60 * 60 * 100) / 100
End Sub
Sub �V�[�gVLU��������()
    Dim �n�� As Date, �����l As Long, �i�� As Long
    Dim ��������(1 To 10000)
    �n�� = Now
    ���s��.Show vbModeless
    Call ���s��.�v���O���X�o�[�X�V(0)
    For �����l = 1 To 10000
        ��������(�����l) = �V�[�gVLU(�����l, "A:A")
        If Int(�����l / 100) - �i�� >= 10 Then
            �i�� = Int(�����l / 100)
            Call ���s��.�v���O���X�o�[�X�V(�i��)
        End If
    Next
    Unload ���s��
    Sheets("����").Cells(7, 3) = Int((Now - �n��) * 24 * 60 * 60 * 100) / 100
End Sub
Sub �z��VLU��������()
    Dim �n�� As Date, �����l As Long, �i�� As Long
    Dim ��������(1 To 10000)
    Dim �z��()
    �n�� = Now
    ���s��.Show vbModeless
    Call ���s��.�v���O���X�o�[�X�V(0)
    With Sheets("�f�[�^")
        �z�� = Range("A1:A10000")
    End With
    For �����l = 1 To 10000
        ��������(�����l) = �z��VLU(�����l, �z��)
        If Int(�����l / 100) - �i�� >= 10 Then
            �i�� = Int(�����l / 100)
            Call ���s��.�v���O���X�o�[�X�V(�i��)
        End If
    Next
    Unload ���s��
    Sheets("����").Cells(8, 3) = Int((Now - �n��) * 24 * 60 * 60 * 100) / 100
End Sub
