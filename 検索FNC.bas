Attribute VB_Name = "����FNC"
Option Explicit
Function �V�[�g�}�b�`(�����l, �͈�)
    With Sheets("�f�[�^")
        On Error Resume Next
        �V�[�g�}�b�` = WorksheetFunction.Match(�����l, .Range(�͈�), 0)
    End With
End Function
Function �z��}�b�`(�����l, �z��)
    With Sheets("�f�[�^")
        On Error Resume Next
        �z��}�b�` = WorksheetFunction.Match(�����l, �z��, 0)
    End With
End Function
Function �V�[�g�l�N�X�g(�����l, �I�s)
    Dim �s As Long
    With Sheets("�f�[�^")
        For �s = 1 To �I�s
            If .Cells(�s, 1) = �����l Then
                �V�[�g�l�N�X�g = �s
                Exit For
            End If
        Next
    End With
End Function
Function �z��l�N�X�g(�����l, �z��)
    Dim �s As Long
    With Sheets("�f�[�^")
        For �s = 1 To UBound(�z��, 1)
            If �z��(�s, 1) = �����l Then
                �z��l�N�X�g = �s
                Exit For
            End If
        Next
    End With
End Function
Function �V�[�gVLU(�����l, �͈�)
    With Sheets("�f�[�^")
        On Error Resume Next
        �V�[�gVLU = WorksheetFunction.VLookup(�����l, .Range(�͈�), 1, 0)
    End With
End Function
Function �z��VLU(�����l, �z��)
    With Sheets("�f�[�^")
        On Error Resume Next
        �z��VLU = WorksheetFunction.VLookup(�����l, �z��, 1, 0)
    End With
End Function
