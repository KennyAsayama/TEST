Attribute VB_Name = "md_���i�֐�"
Option Compare Database
Option Explicit

Public Function IsCasing(in_strWakuHinban As Variant) As Boolean
'   *************************************************************
'   �P�[�V���O�m�F
'
'   �߂�l:Boolean
'       ��True              �P�[�V���O
'       ��False             �P�[�V���O�ȊO
'
'    Input����
'       in_strHinban        �g�i��

'   *************************************************************
    IsCasing = False
    
    If in_strWakuHinban Like "*X*KH*-####*" Or in_strWakuHinban Like "*Y*KH*-####*" Then
        IsCasing = True
    End If
    
End Function

Public Function IsCloset(in_strSetHinban As Variant) As Boolean
'   *************************************************************
'   �N���[�[�b�g�m�F
'
'   �߂�l:Boolean
'       ��True              �N���[�[�b�g
'       ��False             �N���[�[�b�g�ȊO
'
'    Input����
'       in_strHinban        ����i��

'   *************************************************************
    IsCloset = False
    
    If in_strSetHinban Like "M??-?-?####*-*" Or in_strSetHinban Like "�� M??-?-?####*-*" Then
        IsCloset = True
    End If
    
End Function

Public Function IsCloset_Isehara(in_strHinban As Variant) As Boolean
'   *************************************************************
'   �ɐ������Y�N���[�[�b�g�m�F
'
'   �߂�l:Boolean
'       ��True              �ɐ������Y�N���[�[�b�g
'       ��False             �ɐ������Y�N���[�[�b�g�ȊO
'
'    Input����
'       in_strHinban        ����i��

'   *************************************************************
    IsCloset_Isehara = False
    
    If in_strHinban Like "*CME-####*-*" Or in_strHinban Like "*CSA-####*-*" Then
        IsCloset_Isehara = True
    End If
    
End Function

Public Function IsStealth(in_strHinban As Variant) As Boolean
'   *************************************************************
'   �X�e���X�m�F
'
'   �߂�l:Boolean
'       ��True              �X�e���X
'       ��False             �X�e���X�ȊO
'
'    Input����
'       in_strHinban        ���n�i��

'   *************************************************************
    IsStealth = False
    
    If Not in_strHinban Like "*KG*-####*" Then
        IsStealth = True
    End If
    
End Function
Public Function IsStealth_Seizo(in_strHinban As Variant) As Boolean
'   *************************************************************
'   �X�e���X�i�����j�m�F
'
'   �߂�l:Boolean
'       ��True              �X�e���X�i�����j
'       ��False             �X�e���X�i�����j�ȊO
'
'    Input����
'       in_strHinban        ���n�i��

'   *************************************************************
    '20150820���ݖ��g�p
    IsStealth_Seizo = False
    
'    If (in_strHinban Like "*SG*-####*" Or in_strHinban Like "*NG*-####*" Or in_strHinban Like "*AG*-####*" Or in_strHinban Like "*BG*-####*") _
'        And Not in_strHinban Like "*ML-####*" And Not in_strHinban Like "*MK-####*"  And Not in_strHinban Like "*MT-####*" And Not in_strHinban Like "*DU-####*" And Not in_strHinban Like "*DN-####*" And Not in_strHinban Like "*CTSG*MK-####*" And Not in_strHinban Like "*CTSG*ML-####*"  And Not in_strHinban Like "*CTSG*MT-####*"  And Not in_strHinban Like "*KU-####*"  And Not in_strHinban Like "*KN-####*" Then
'        IsStealth_Seizo = True
'    End If
    
End Function
Public Function intFncSeizokubun(in_strShurui As String, in_varHinban As Variant) As Integer
'   *************************************************************
'   �����敪�擾
'
'   �߂�l:Integer
'       ��                  �����敪
'
'    Input����
'       in_strShurui        ���
'       in_varHinban        �i��

'   *************************************************************
    
    intFncSeizokubun = 0
    
    Select Case in_strShurui
    
        Case "����", "�q��"
            
            If IsKamachi(in_varHinban) Then
            
                intFncSeizokubun = 3
                
            ElseIf IsFkamachi(in_varHinban) Then
            
                intFncSeizokubun = 2
                
            Else
            
                intFncSeizokubun = 1
                
            End If
            
        Case "�۾ޯ�"
        
            If IsCloset_Isehara(in_varHinban) Then  '�۾ޯ�(�ɐ������Y)
                intFncSeizokubun = 1
            End If
        Case "�g"
        
            If IsCasing(in_varHinban) Then
                intFncSeizokubun = 5
            Else
                intFncSeizokubun = 4
            End If
        Case "���n"
        
            If IsStealth_Seizo(in_varHinban) Then
                intFncSeizokubun = 7
            Else
                intFncSeizokubun = 6
            End If
            
    End Select
    
End Function

Public Function intFncSeihinkubun(in_strShurui As String, in_varHinban As Variant) As Integer
'   *************************************************************
'   �����敪�擾
'   ���i�敪������R�[�h���擾
'
'   �߂�l:Integer
'       ��                  �����敪
'
'    Input����
'       in_strShurui        ���
'       in_varHinban        �i��

'   *************************************************************

'*****************************************
'1.���ɃR�[�h����敪�����擾����t��������
'�@(�X�V�̍ۂ͓�������邱��)
'2.�R�[�h�̒ǉ��ύX�폜�̏ꍇ�`�F�b�N���X�g�̊֐����C��
'�@(�֐���:intFncSeizoKubunToChecklistCode)
'*****************************************

    Dim intChecklistikubun As Integer
    
    intFncSeihinkubun = 0
    
    Select Case in_strShurui
    
        Case "����", "�q��"

            intFncSeihinkubun = 1
            
        Case "�۾ޯ�"

            intFncSeihinkubun = 5
            
        Case "�g"
        
            If IsCasing(in_varHinban) Then
                intFncSeihinkubun = 4
            Else
                intFncSeihinkubun = 2
            End If

            
        Case "���n"
        
            intFncSeihinkubun = 3
            
        Case "�����"
        
            intFncSeihinkubun = 6
            
        Case "���֎��["
        
            intFncSeihinkubun = 7
            
        Case "����"
        
            intFncSeihinkubun = 8
            
        Case "�z����"
        
            intFncSeihinkubun = 9
            
        Case "����"

            intFncSeihinkubun = 10

        Case "�K�i"

            intFncSeihinkubun = 11

        Case "̧����"

            intFncSeihinkubun = 12
           
    End Select
    
        
End Function

Public Function strFncSeihinkubunMei(in_intSeihinkubun As Integer) As String
'   *************************************************************
'   �����敪���擾
'   ���i�R�[�h����敪�����擾
'
'   �߂�l:Integer
'       ��                  �����敪
'
'    Input����
'       in_strShurui        ���
'       in_varHinban        �i��

'   *************************************************************

'*****************************************
'��ɋ敪������R�[�h���擾����t��������
'(�X�V�̍ۂ͓�������邱��)
'*****************************************
    strFncSeihinkubunMei = ""
    
    Select Case in_intSeihinkubun

        Case 5
            strFncSeihinkubunMei = "�۾ޯ�"

        Case 1
            strFncSeihinkubunMei = "����"

        Case 2
            strFncSeihinkubunMei = "�g"

        Case 4
            strFncSeihinkubunMei = "�O���g"

        Case 3
            strFncSeihinkubunMei = "���n"

        Case 6
            strFncSeihinkubunMei = "�����"

        Case 7
            strFncSeihinkubunMei = "���֎��["

        Case 8
            strFncSeihinkubunMei = "����"
        
        Case 9
            strFncSeihinkubunMei = "�z����"
            
        Case 10
            strFncSeihinkubunMei = "����"

        Case 11
            strFncSeihinkubunMei = "�K�i"

        Case 12
            strFncSeihinkubunMei = "̧����"

          
    End Select
    
End Function

Public Function IsFkamachi(in_strHinban As Variant) As Boolean
'   *************************************************************
'   Flush�y�m�F
'
'   �߂�l:Boolean
'       ��True              F�y
'       ��False             F�y�ȊO
'
'    Input����
'       in_strHinban        ����i��

'   *************************************************************
    
    IsFkamachi = False
    
    If IsNull(in_strHinban) Then Exit Function
       
    If in_strHinban Like "*-####G*-*" Or in_strHinban Like "*-####MF*-*" Or in_strHinban Like "*O*-####P*-*" Then
        IsFkamachi = True
    End If
    
End Function

Public Function IsKamachi(in_strHinban As Variant) As Boolean
'   *************************************************************
'   �y�m�F
'
'   �߂�l:Boolean
'       ��True              �y
'       ��False             �y�ȊO
'
'    Input����
'       in_strHinban        ����i��

'   *************************************************************

    IsKamachi = False
    
End Function

Public Function IsThruGlass(in_strHinban As Variant) As Boolean
'   *************************************************************
'   �X���[�K���X�m�F
'   �T�u�t�H�[���̏����t��������̌Ăяo���ŏ��������ەs�v�ȌĂяo������������̂ŃG���[���W�b�N��ǉ�
'
'   �߂�l:Boolean
'       ��True              �X���[
'       ��False             �X���[�ȊO
'
'    Input����
'       in_strHinban        ����i��

'   *************************************************************
    On Error GoTo Err_IsThruGlass
    
    IsThruGlass = False
    
    If IsNull(in_strHinban) Then Exit Function
     
    If in_strHinban Like "*-####S*-*" Or in_strHinban Like "*-####C*-*" Or in_strHinban Like "*-####D*-*" _
        Or in_strHinban Like "*-####A*-*" Or in_strHinban Like "*-####B*-*" Or in_strHinban Like "*-####O*-*" _
        Or in_strHinban Like "*ME-####M*-*" Or in_strHinban Like "*SA-####M*-*" Or IsVertica(in_strHinban) Then
        
        IsThruGlass = True
    Else
        IsThruGlass = False
    End If
    
    Exit Function
    
Err_IsThruGlass:
    IsThruGlass = False
End Function

Public Function IsOyatobira(in_strHinban As Variant) As Boolean
'   *************************************************************
'   �e���m�F
'
'   �߂�l:Boolean
'       ��True              �e��
'       ��False             �e���ȊO
'
'    Input����
'       in_strHinban        ����i��
'   *************************************************************

    IsOyatobira = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "*DO-####*" Or in_strHinban Like "*DOS-####*" _
       Or in_strHinban Like "*CO-####*" Or in_strHinban Like "*COS-####*" _
                                                                            Then
        IsOyatobira = True
    Else
        IsOyatobira = False
    End If
    
End Function

Public Function IsKotobira(in_strHinban As Variant) As Boolean
'   *************************************************************
'   �q���m�F
'
'   �߂�l:Boolean
'       ��True              �q��
'       ��False             �q���ȊO
'
'    Input����
'       in_strHinban        ����i��
'   *************************************************************

    IsKotobira = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "*DK-####*" Or in_strHinban Like "*DKS-####*" Then
        IsKotobira = True
    Else
        IsKotobira = False
    End If
    
End Function

Public Function IsSxL(in_strHinban As Variant, out_strKamiyahinban As Variant) As Boolean
'   *************************************************************
'   �G�X�o�C�G���m�F
'
'   �߂�l:Boolean
'       ��True              �G�X�o�C�G��
'       ��False             �G�X�o�C�G���ȊO
'
'    Input����
'       in_strHinban        ����i��
'       out_strSxLhinban    �_�J�i��(False�̏ꍇ��Null)
'   *************************************************************
    
    Dim objLOCALDB As New cls_LOCALDB
    Dim strHinban As String
    Dim bolMentori As Boolean
    
    IsSxL = False
    
    On Error GoTo Err_IsSxL
    
    If IsNull(in_strHinban) Then GoTo Exit_IsSxL
    
    '���n�Ŗʎ��L��������ꍇ�͊O��
    If in_strHinban Like "*�@?�A?�B?�C*" Then
        strHinban = left(in_strHinban, Len(in_strHinban) - 10)
        bolMentori = True
    Else
        strHinban = in_strHinban
        bolMentori = False
    End If
    '1.10.3 K.Asayama 20151119 SxL�i�ԓǑ֕\���[�J���e�[�u�����ύX
    If objLOCALDB.ExecSelect("select �u�����h�i�� from WK_SxL�i�ԓǑ֕\ where S�~L�i�� = '" & Trim(strHinban) & "'") Then
        If Not objLOCALDB.GetRS.EOF Then
            out_strKamiyahinban = objLOCALDB.GetRS![�u�����h�i��]
            If bolMentori Then
                out_strKamiyahinban = out_strKamiyahinban & right(in_strHinban, 10)
            End If
            IsSxL = True
        End If
    End If
    
    GoTo Exit_IsSxL
    
Err_IsSxL:
    IsSxL = False
    
Exit_IsSxL:
'�N���X�̃C���X�^���X��j��
    Set objLOCALDB = Nothing
End Function

Public Function valfncHinmei(in_objRemoteDB As cls_BRAND_MASTER, in_Rs As ADODB.Recordset, in_strHinban As Variant, in_intSeihinkubun As Integer, in_strSpec As Variant) As Variant
'   *************************************************************
'   �i�����o
'   20151116 1.10.2 ��Spec��Variant�ɕύX�iNull�̉\�������邽�߁j
'   �߂�l:Variant �� �i���i������Ȃ��ꍇ��NULL�j
'
'    Input����
'       in_objREMOTEDB      �f�[�^�x�[�X�T�[�o
'       in_strHinban        ����i�ԁi�����͊O���Ă����j
'       in_intSeihinkubun   �i�ԋ敪
'       in_strSpec          ��Spec
'   *************************************************************
    Dim strSQL As String
    
    strSQL = ""
    valfncHinmei = Null
    
    On Error GoTo Err_valfncHinmei
    
    If IsNull(in_strHinban) Then GoTo Exit_valfncHinmei
    
    '1.10.2 �p�~***********************************
    'If in_strSpec = "" Then GoTo Exit_valfncHinmei
    '**********************************************
    
    Select Case in_intSeihinkubun
        Case 1, 5 '����,�۾ޯ�
            strSQL = "select top 1 �i�� from T_����i��Ͻ� where "
                If IsKotobira(in_strHinban) Then
                    strSQL = strSQL & " �q���i�� = '" & in_strHinban & "'"
                Else
                    strSQL = strSQL & " ����i�� = '" & in_strHinban & "'"
                End If
        Case 2, 4 '�g,�O���g
            strSQL = "select top 1 �i�� from T_�g�i��Ͻ� where �g�i�� = '" & in_strHinban & "'"
            
        Case 3 '���n�g
            strSQL = "select top 1 �i�� from T_���n�ޕi��Ͻ� where ���n�ޕi�� = '" & in_strHinban & "'"
          
        Case 6 '�����
            strSQL = "select top 1 �i�� from T_����ޕi��Ͻ� where ����ޕi�� = '" & in_strHinban & "'"
            
        Case 7 '���֎��[
            strSQL = "select top 1 �i�� from T_���֎��[Ͻ� where �i�� = '" & in_strHinban & "'"
            
        Case 8 '����
            strSQL = "select top 1 �i�� from T_�����i��Ͻ� where �����i�� = '" & in_strHinban & "'"
        
    End Select
    
    If strSQL = "" Then
        GoTo Exit_valfncHinmei
    Else
        '1.10.2 ****************************************************************************************************************
        'strSQL = strSQL & " and �d�l = '" & left(in_strSpec, 3) & "' and '" & right(in_strSpec, 4) & "' between �J�n and �I�� "
        If Not IsNull(in_strSpec) And in_strSpec <> "" Then
            strSQL = strSQL & " and �d�l = '" & left(in_strSpec, 3) & "' and '" & right(in_strSpec, 4) & "' between �J�n and �I�� "
        End If
        '***********************************************************************************************************************
    End If
    
    
    If in_objRemoteDB.ExecSelect_ExternalRS(in_Rs, strSQL) Then
        If Not in_Rs.EOF Then
            valfncHinmei = in_Rs![�i��]
        End If
    End If
    
    GoTo Exit_valfncHinmei
    
Err_valfncHinmei:
    'MsgBox Err.Description
Exit_valfncHinmei:

End Function

Public Function bolFncTokuHinban(in_varHinban As Variant, in_varTokuHinban As Variant, ByRef out_varTokuhinban As Variant) As Boolean
'   *************************************************************
'   �����i�Ԋm�F
'   �i�Ԃ������i�Ԃ��m�F�������i�Ԃ̏ꍇ�͒ʏ�i�Ԃ�Ԃ�
'   SxL�i�ԂɊY������ꍇ�_�J�i�Ԃ�Ԃ�
'
'   �߂�l:Boolean
'       ��True              ����
'       ��False             �����ȊO
'
'    Input����
'       in_varHinban        �󒍕i��
'       in_varTokuHinban    �����󒍕i��
'       out_varTokuhinban   �󒍕i�ԁi�����̏ꍇ--���́u�� �v���O�������́j
'   *************************************************************
    
    Dim varHinban As Variant
    
    out_varTokuhinban = Null
    
    If Not IsNull(in_varHinban) And in_varHinban <> "" Then
        Select Case in_varHinban
            Case "TATEGU", "TATEGUM", "TATEGUW", "TATEGUS", "BUZAI", "BUZAIK", "BUZAIG"
                varHinban = Mid(in_varTokuHinban, 3)
                bolFncTokuHinban = True
            Case Else
                varHinban = in_varHinban
        End Select
        
        'SxL�i�ԃ`�F�b�N
        If Not IsSxL(varHinban, out_varTokuhinban) Then
            out_varTokuhinban = varHinban
        End If
        
    End If
End Function

Public Function intFncChecklistCode(in_Kubun As String) As Integer
'   *************************************************************
'   �`�F�b�N���X�g�̋敪�擾
'   �R�[�h�̓��[�J�����[��
'
'   �߂�l:Integer
'       ��                  �`�F�b�N���X�g�p�R�[�h
'                           ��;�q��;�g;��;���n;��;��;��;�K;�t
'    Input����
'       in_Kubun            T_�`�F�b�N���X�g�̋敪
'                          �Y������敪�������ꍇ��0��Ԃ�
'   *************************************************************

    Select Case in_Kubun
        Case "��"
            intFncChecklistCode = 1
        Case "�q��"
            intFncChecklistCode = 1
        Case "�g"
            intFncChecklistCode = 2
        Case "���n"
            intFncChecklistCode = 3
        Case "��"
            intFncChecklistCode = 6
        Case "��"
            intFncChecklistCode = 8
        Case "��"
            intFncChecklistCode = 7
        Case "��"
            intFncChecklistCode = 10
        Case "�K"
            intFncChecklistCode = 11
        Case "�t"
            intFncChecklistCode = 12
        Case Else
            intFncChecklistCode = 0
    End Select

End Function

Public Function intFncSeizoKubunToChecklistCode(in_Kubun As Integer) As Integer
'   *************************************************************
'   �����敪�ɑΉ�����`�F�b�N���X�g�̃R�[�h�擾
'
'   �߂�l:Integer
'       ��                  �`�F�b�N���X�g�p�R�[�h
'                           ��;�q��;�g;��;���n;��;��;��;�K;�t
'    Input����
'       in_Kubun           �����敪
'                          �Y������敪�������ꍇ��0��Ԃ�
'   *************************************************************

    '�����敪�ɑΉ�����`�F�b�N���X�g�̃R�[�h��Ԃ�

    Select Case in_Kubun
        Case 1, 2, 3, 5 'Flush,F�y,�y,�۾ޯ�
            intFncSeizoKubunToChecklistCode = 1
        Case 4, 5 '�g�ƎO���g
            intFncSeizoKubunToChecklistCode = 2
        Case 6, 7 '���n�ƃX�e���X
            intFncSeizoKubunToChecklistCode = 3
        Case Else '���̑��͖���
            intFncSeizoKubunToChecklistCode = 0
    End Select

End Function

Public Function intFncSeihinKubunToChecklistCode(in_intSeihinkubun As Integer) As Integer
'   *************************************************************
'   ���i�敪�ɑΉ�����`�F�b�N���X�g�̃R�[�h�擾
'
'   �߂�l:Integer
'       ��                  �`�F�b�N���X�g�p�R�[�h
'                           ��;�q��;�g;��;���n;��;��;��;�K;�t
'    Input����
'       in_Kubun           ���i�敪
'                          �Y������敪�������ꍇ��0��Ԃ�
'   *************************************************************
    
    intFncSeihinKubunToChecklistCode = 0
    
    Select Case in_intSeihinkubun
    
        Case 5  '����A�۾ޯ�

            intFncSeihinKubunToChecklistCode = 1
            
        Case Else
        
            intFncSeihinKubunToChecklistCode = in_intSeihinkubun
           
    End Select
    
End Function

Public Function isCaro(in_varHinban As Variant) As Boolean
'   *************************************************************
'   Caro�m�F
'
'   �߂�l:Boolean
'       ��True              Caro
'       ��False             Caro�ȊO
'
'    Input����
'       in_strHinban        ����i��

'   *************************************************************

    isCaro = False
    
    If in_varHinban Like "F_C*-####A*-*" Or in_varHinban Like "F_C*-####B*-*" Or in_varHinban Like "F_C*-####O*-*" _
        Or in_varHinban Like "�� F_C*-####A*-*" Or in_varHinban Like "�� F_C*-####B*-*" Or in_varHinban Like "�� F_C*-####O*-*" Then
    
        isCaro = True
        
    End If
    
End Function

Public Function strfncDaibunrui_Kamui(in_strShurui As String, in_varHinban As Variant) As String
'   *************************************************************
'   ��ނ���J���C�̑啪�ނ��擾
'
'   �߂�l:String
'       ��                  �J���C�̑啪��
'                           �Y������敪�������ꍇ��"00"��Ԃ�

'    Input����
'       in_strShurui        ���
'       in_varHinban        �i��
'   *************************************************************
'
'�O���g�̂ݕi�Ԃ��K�v
    
    strfncDaibunrui_Kamui = "00"
    
    Select Case in_strShurui
    
        Case "����", "�q��"

            strfncDaibunrui_Kamui = "11"
            
        Case "�۾ޯ�"

            strfncDaibunrui_Kamui = "21"
            
        Case "�g"
        
            If IsCasing(in_varHinban) Then
                strfncDaibunrui_Kamui = "03"
            Else
                strfncDaibunrui_Kamui = "02"
            End If

            
        Case "���n"
        
            strfncDaibunrui_Kamui = "01"
            
        Case "�����"
        
            strfncDaibunrui_Kamui = "41"
            
        Case "���֎��["
        
            strfncDaibunrui_Kamui = "61"
            
        Case "����"
        
            strfncDaibunrui_Kamui = "51"
            
        Case "�z����"
        
            
        Case "����"


        Case "�K�i"


        Case "̧����"
    
    End Select
    
End Function

Public Function strfncSyobunrui_Kamui(in_strDaibunrui_Kamui As String, in_varHinban As Variant) As String
'   *************************************************************
'   �J���C�̑啪�ނƕi�Ԃ�����J���C�����ނ��擾
'
'   �߂�l:String
'       ��                              �J���C�̏�����
'
'    Input����
'       in_strDaibunrui_Kamui           �J���C�̑啪��
'       in_varHinban                    �i��
'   *************************************************************

    Dim strHinbanKigou As String
    
    Select Case in_strDaibunrui_Kamui
    
        Case "01" '���n
            strHinbanKigou = left(in_varHinban, 1)
            
            Select Case strHinbanKigou
                Case "S", "N", "A", "B"
                    strfncSyobunrui_Kamui = strHinbanKigou
                    
                Case Else
                    strfncSyobunrui_Kamui = "W"
                    
            End Select
        Case "02" '�g
            strfncSyobunrui_Kamui = "W"
            
        Case "03" '�O���g
            strfncSyobunrui_Kamui = "C"
            
        Case "11" '�o����
            strHinbanKigou = left(in_varHinban, 1)
            
            If in_varHinban Like "F_V*-####*" Then 'Vertica
                strfncSyobunrui_Kamui = "V"
                
            ElseIf in_varHinban Like "F_C*-####*" Then 'Caro
                strfncSyobunrui_Kamui = "A"
            
            Else
                Select Case strHinbanKigou
                    Case "F" '�W���i��CUBE�̃R�[�h�𑗂�
                        strfncSyobunrui_Kamui = "C"
                    Case "S" 'F/S
                        strfncSyobunrui_Kamui = "K"
                    Case "A" 'Air
                        strfncSyobunrui_Kamui = "F"
                    Case Else
                        strfncSyobunrui_Kamui = strHinbanKigou
                End Select
            End If
            
        Case "21" '�N���[�b�g
            strfncSyobunrui_Kamui = "M"
        
        Case "31" '�E�H�[�N�X���[
            If in_varHinban Like "*-####G*" Then        '�K���X
                strfncSyobunrui_Kamui = "G"
            ElseIf in_varHinban Like "*-####L*" Then    '���[�o�[
                strfncSyobunrui_Kamui = "L"
            Else
                strfncSyobunrui_Kamui = "C"             '�R���r
            End If
            
        Case "41" '�����
            strfncSyobunrui_Kamui = "99999" '�\�����Ȃ�
            
        Case "51" '����
            strfncSyobunrui_Kamui = 1
            
        Case "61" '���֎��[
            strfncSyobunrui_Kamui = 1
            
    End Select


End Function

Public Function IsGikan(in_strHinban As Variant) As Boolean
'   *************************************************************
'   �Z�������m�F
'   �T�u�t�H�[���̏����t��������̌Ăяo���ŏ��������ەs�v�ȌĂяo������������̂ŃG���[���W�b�N��ǉ�
'   'ADD by Asayama 20150903
'   �߂�l:Boolean
'       ��True              �Z������
'       ��False             �Z�������ȊO
'
'    Input����
'       in_strHinban        ����i��

'   *************************************************************
    On Error GoTo Err_IsGikan
    
    IsGikan = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    '�X���[�K���X
    If IsThruGlass(in_strHinban) Then
        IsGikan = True
    
    '�����背�X�iVertica�j
    ElseIf IsVertica(in_strHinban) Then
        IsGikan = True
    
    'Air
    ElseIf IsAir(in_strHinban) Then
        IsGikan = True
     
    Else
        IsGikan = False
        
    End If
    
    Exit Function
    
Err_IsGikan:
    IsGikan = False
End Function

Public Function IsVertica(in_strHinban As Variant) As Boolean
'   *************************************************************
'   �����背�X���ˊm�F
'   �T�u�t�H�[���̏����t��������̌Ăяo���ŏ��������ەs�v�ȌĂяo������������̂ŃG���[���W�b�N��ǉ�
'   'ADD by Asayama 20150903
'   �߂�l:Boolean
'       ��True              �����背�X
'       ��False             �����背�X�ȊO
'
'    Input����
'       in_strHinban        ����i��

'   *************************************************************
    On Error GoTo Err_IsVertica
    
    IsVertica = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "??V*-####*-*" Or in_strHinban Like "�� ??V*-####*-*" Then
        IsVertica = True
    Else
        IsVertica = False
    End If
    
    Exit Function
    
Err_IsVertica:
    IsVertica = False
    
End Function

Public Function IsAir(in_strHinban As Variant) As Boolean
'   *************************************************************
'   FullHeight Air�m�F
'   �T�u�t�H�[���̏����t��������̌Ăяo���ŏ��������ەs�v�ȌĂяo������������̂ŃG���[���W�b�N��ǉ�
'   'ADD by Asayama 20150903
'   �߂�l:Boolean
'       ��True              Air
'       ��False             Air�ȊO
'
'    Input����
'       in_strHinban        ����i��

'   *************************************************************
    On Error GoTo Err_IsAir
    
    IsAir = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "A*-####SC*-*" Or in_strHinban Like "A*-####SL*-*" Or in_strHinban Like "�� A*-####SC*-*" Or in_strHinban Like "�� A*-####SL*-*" Then
        IsAir = True
    Else
        IsAir = False
    End If
    
    Exit Function
    
Err_IsAir:
    IsAir = False
    
End Function

Public Function IsPainted(in_strHinban As Variant) As Boolean
'   *************************************************************
'   �h�����m�F
'   �T�u�t�H�[���̏����t��������̌Ăяo���ŏ��������ەs�v�ȌĂяo������������̂ŃG���[���W�b�N��ǉ�
'   'ADD by Asayama 201510**
'   �߂�l:Boolean
'       ��True              �h����
'       ��False             �h�����ȊO
'
'    Input����
'       in_strHinban        ����i��

'   *************************************************************
    On Error GoTo Err_IsPainted
    
    IsPainted = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "R*-####*-*" Or in_strHinban Like "�� R*-####*-*" Or in_strHinban Like "B*-####*-*" Or in_strHinban Like "�� B*-####*-*" Then
        IsPainted = True
    Else
        IsPainted = False
    End If
    
    Exit Function
    
Err_IsPainted:
    IsPainted = False
    
End Function

Public Function IsMonster(in_strHinban As Variant) As Boolean
'   *************************************************************
'   �����X�^�[���m�F
'   �T�u�t�H�[���̏����t��������̌Ăяo���ŏ��������ەs�v�ȌĂяo������������̂ŃG���[���W�b�N��ǉ�
'   'ADD by Asayama 201510**
'   �߂�l:Boolean
'       ��True              �����X�^�[���m�F
'       ��False             �����X�^�[���m�F�ȊO
'
'    Input����
'       in_strHinban        ����i��

'   *************************************************************
    On Error GoTo Err_IsMonster
    
    IsMonster = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "O*-####*-*" Or in_strHinban Like "�� O*-####*-*" Then
        IsMonster = True
    Else
        IsMonster = False
    End If
    
    Exit Function
    
Err_IsMonster:
    IsMonster = False
    
End Function

Public Function IsStealth_Seizo_TEMP(in_strHinban As Variant) As Boolean
'   *************************************************************
'   �X�e���X�i�����j�m�F�i��L IsStealth_Seizo�g�p�J�n���ɂ͍����ւ��j
'
'   �߂�l:Boolean
'       ��True              �X�e���X�i�����j
'       ��False             �X�e���X�i�����j�ȊO
'
'    Input����
'       in_strHinban        ���n�i��

'   *************************************************************
    '
    IsStealth_Seizo_TEMP = False
    
    If (in_strHinban Like "*SG*-####*" Or in_strHinban Like "*NG*-####*" Or in_strHinban Like "*AG*-####*" Or in_strHinban Like "*BG*-####*") _
        And Not in_strHinban Like "*ML-####*" And Not in_strHinban Like "*MK-####*" And Not in_strHinban Like "*MT-####*" And Not in_strHinban Like "*DU-####*" And Not in_strHinban Like "*DN-####*" And Not in_strHinban Like "*CTSG*MK-####*" And Not in_strHinban Like "*CTSG*ML-####*" And Not in_strHinban Like "*CTSG*MT-####*" And Not in_strHinban Like "*KU-####*" And Not in_strHinban Like "*KN-####*" Then
        IsStealth_Seizo_TEMP = True
    End If
    
End Function

Public Function fncbolSxL_Replace() As Boolean
'   *************************************************************
'   SxL�i�ԓǑ֕\�u��������
'   1.10.3 K.Asayama ADD 20151119 SxL�i�ԕ\�����[�g����R�s�[
'
'   �����[�g�f�[�^�x�[�X���烍�[�J����SxL�i�ԓǑ֕\���R�s�[����
'
'   �߂�l:Boolean
'       ��True              �u������
'       ��False             �u�����s
'
'   *************************************************************

    fncbolSxL_Replace = False
    
    Dim objREMOTEDB As New cls_BRAND_MASTER
    Dim objLOCALDB As New cls_LOCALDB
    
    On Error GoTo Err_fncbolSxL_Replace
    
    Dim strSQL_Insert As String
    Dim strSQL As String
    strSQL_Insert = "Insert into WK_SxL�i�ԓǑ֕\(S�~L�i��,�u�����h�i��,DH,DW,CH) values ("
    
    '�H��p�R�s�[�iT_Calendar_�H��)
    If objLOCALDB.ExecSQL("delete from WK_SxL�i�ԓǑ֕\") Then
        strSQL = "select distinct [S�~L�i��],�u�����h�i��,DW,DH,CH from SxL�i�ԓǑ֕\ "
        If objREMOTEDB.ExecSelect(strSQL) Then
            Do While Not objREMOTEDB.GetRS.EOF
                If Not objLOCALDB.ExecSQL(strSQL_Insert & "'" & objREMOTEDB.GetRS![S�~L�i��] & "','" & objREMOTEDB.GetRS![�u�����h�i��] & "'," & objREMOTEDB.GetRS![DW] & "," & objREMOTEDB.GetRS![DH] & "," & objREMOTEDB.GetRS![CH] & ")") Then
                    Err.Raise 9999, , "SxL�i�ԓǑ֕\ ���[�J���R�s�[�G���["
                End If
                objREMOTEDB.GetRS.MoveNext
            Loop
        End If
    End If
    
    fncbolSxL_Replace = True
    
    GoTo Exit_fncbolSxL_Replace
    
Err_fncbolSxL_Replace:
    MsgBox Err.Description
    
Exit_fncbolSxL_Replace:

    Set objREMOTEDB = Nothing
    Set objLOCALDB = Nothing
    
End Function
