Option Compare Database
Option Explicit
'2.1.0 ADD

Public Function IsGuidePieceJ(in_Hinban As String, in_��Spec As Variant) As Boolean
'   *************************************************************
'   �K�C�h�s�[�X��t��K�o�͑Ώۃ`�F�b�N
'
'   �߂�l:Boolean
'       ��True              ��K����
'       ��False             ��K����
'
'    Input����
'       in_Hinban           ����i��
'       in_��Spec         ��Spec

'   *************************************************************

    Dim strSpec_BRAND As String
    Dim strSPEC As String
        
    IsGuidePieceJ = False
    
    If IsYukazukeRail(in_Hinban) Then
    
        If Not IsNull(in_��Spec) Then
            strSpec_BRAND = left(in_��Spec, 3)
            strSPEC = right(in_��Spec, 4)
            
            If strSpec_BRAND = "BRD" And strSPEC >= "1507" Then
                IsGuidePieceJ = True
            End If
            
        End If
    End If

End Function

Public Function fncdblGuidePiece_SH(in_Hinban As String, dblDH As Double) As Double
'   *************************************************************
'   �K�C�h�s�[�XSH�l
'       �K�C�h�s�[�X�i�Ԃ��ǂ����͖{�֐��ł̓`�F�b�N���Ȃ��̂�
'       ���O��IsGuidePieceJ�֐��Ń`�F�b�N���Ă�������

'   �߂�l:Double
'       ��SH�l
'
'    Input����
'       in_Hinban           ����i��
'       dblDH               DH

'   *************************************************************

    If in_Hinban Like "*VL-####*" Or in_Hinban Like "*VM-####*" Or in_Hinban Like "*VN-####*" Then
        fncdblGuidePiece_SH = dblDH - 47
    Else
        fncdblGuidePiece_SH = dblDH - 41
    End If
    
End Function

Public Function fncstrGuidePiece_MM(dblDH As Double) As String
'   *************************************************************
'   �K�C�h�s�[�Xmm�l
'       �K�C�h�s�[�X�i�Ԃ��ǂ����͖{�֐��ł̓`�F�b�N���Ȃ��̂�
'       ���O��IsGuidePieceJ�֐��Ń`�F�b�N���Ă�������

'   �߂�l:String
'       ��mm�l
'
'    Input����
'       dblDH               DH

'   *************************************************************
    If dblDH < 2411 Then
        fncstrGuidePiece_MM = "4mm"
    Else
        fncstrGuidePiece_MM = "7mm"
    End If
    
End Function

Public Function fncstrGuidePiece_Name(in_Hinban As String, in_��Spec As Variant) As String
'   *************************************************************
'   �K�C�h�s�[�X���́i�R�[�h�j�擾
'       �K�C�h�s�[�X�i�Ԃ��ǂ����͖{�֐��ł̓`�F�b�N���Ȃ��̂�
'       ���O��IsGuidePieceJ�֐��Ń`�F�b�N���Ă�������

'       ���O�Ɍ�Spec��Null�łȂ����Ƃ��m�F���邱�Ɓi������Null�̏ꍇ�͋󗓂�Ԃ��j

'   �߂�l:String
'       ���K�C�h�s�[�X����

'           1608���O�� A

'           1608�ȍ~
'               ��DM�� A
'               ������ȊO�� B

'           1701�ȍ~��V���[��
'               ��VM��C

'           ������ȊO��D

'    Input����
'       in_Hinban           ����i��
'       in_��Spec         ��Spec

'   *************************************************************

    If IsNull(in_��Spec) Then
        fncstrGuidePiece_Name = ""
        
    ElseIf right(in_��Spec, 4) < "1608" Then
        fncstrGuidePiece_Name = "A"
    Else
        
        If in_Hinban Like "*DM-####*" Then
            fncstrGuidePiece_Name = "A"
        ElseIf in_Hinban Like "*DL-####*" Or in_Hinban Like "*DN-####*" Then
            fncstrGuidePiece_Name = "B"
        ElseIf in_Hinban Like "*VM-####*" Then
            fncstrGuidePiece_Name = "C"
        ElseIf in_Hinban Like "*VL-####*" Or in_Hinban Like "*VN-####*" Then
            fncstrGuidePiece_Name = "D"
        End If
        
    End If

End Function