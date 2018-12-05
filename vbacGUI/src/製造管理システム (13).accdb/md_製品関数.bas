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
'       in_strSetHinban     �Z�b�g�i��

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
    
'    If in_strHinban Like "*PW*-####*" Then '�G�X�p�X���C�h�E�H�[���̓C���Z�b�g
'        IsStealth_Seizo = False
'        Exit Function
'    End If
'
'    If (in_strHinban Like "*SG*-####*" Or in_strHinban Like "*NG*-####*" Or in_strHinban Like "*AG*-####*" Or in_strHinban Like "*BG*-####*") _
'        And Not in_strHinban Like "*ML-####*" And Not in_strHinban Like "*MK-####*" And Not in_strHinban Like "*MT-####*" And Not in_strHinban Like "*DU-####*" And Not in_strHinban Like "*DN-####*" And Not in_strHinban Like "*CTSG*MK-####*" And Not in_strHinban Like "*CTSG*ML-####*" And Not in_strHinban Like "*CTSG*MT-####*" And Not in_strHinban Like "*KU-####*" And Not in_strHinban Like "*KN-####*" And Not in_strHinban Like "*DV-####*" Then
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

'2.7.0
'   ���t���K���X�͋敪0�i�������Ȃ��j
'   *************************************************************
    
    intFncSeizokubun = 0
    
    Select Case in_strShurui
    
        Case "����", "�q��"
            
            If IsFullGlass(in_varHinban) Then
                
                intFncSeizokubun = 0
            
            ElseIf IsKamachi(in_varHinban) Then
            
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

'   1.10.11 20160302 K.Asayama ADD
'           ���G�X�p�X���C�h�E�H�[���ǉ�
'   1.10.12 20160322 K.Asayama Change
'           ��AF1�`AF3�i�J���j�ǉ�
'   1.10.19 K.Asayama Change
'           ��1608�ȍ~�̃~���[��Flush�i�X���[�K���X�j
'   1.11.0
'           ���e���X�h�A�ǉ�
'   1.11.3
'           �������X�^�[�i�ԕύX�Ή��i�֐��j
'   2.3.0
'           ��1801�d�l�ǉ��@G9�^
'   2.5.2
'           ��1801�d�l�ǉ��@�i�q��
'   *************************************************************
    
    IsFkamachi = False
    
    If IsNull(in_strHinban) Then Exit Function
       
    '1.10.19
    'If in_strHinban Like "*-####G*-*" Or in_strHinban Like "*-####MF*-*" Or in_strHinban Like "*O*-####P*-*" Then
    If in_strHinban Like "*-####G*-*" Or in_strHinban Like "F?B*-####MF*-*" Or in_strHinban Like "�� F?B*-####MF*-*" Or IsMonster(in_strHinban) Then
        IsFkamachi = True
       
    'Caro
    ElseIf in_strHinban Like "F?B??*-####A*-*" Or in_strHinban Like "F?B??*-####B*-*" Or in_strHinban Like "F?B??*-####O*-*" Then
         IsFkamachi = True
    
    'Terrace(YG6�^,YG5�^)
    ElseIf in_strHinban Like "Y?B??*-####W*-*" Then
         IsFkamachi = True
         
    'G9�^
    ElseIf IsG9(in_strHinban) Then
         IsFkamachi = True
         
    '�i�q�^
    ElseIf IsKousi(in_strHinban) Then
         IsFkamachi = True
    End If
    
    '1.10.11 ADD �G�X�p�X���C�h�E�H�[��
    If in_strHinban Like "*PSW*-####FV*-*" Then
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
'
'   1.10.9 201602** K.Asayama ADD
'           ���y�����쐬
'   1.10.11 20160302 K.Asayama ADD
'           ���G�X�p���A���[�g���O
'   *************************************************************
    
    IsKamachi = False
    
    '1.10.9 ADD
    On Error GoTo Err_IsKamachi
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "??R*-####*-*" Or in_strHinban Like "�� ??R*-####*-*" Then
        '1.10.11 Change
            'IsKamachi = True
            If Not in_strHinban Like "HER*-####*-*" And Not in_strHinban Like "�� HER*-####*-*" Then
            
                IsKamachi = True
            End If
        '1.10.11 Change END
    End If
    
    Exit Function
    
Err_IsKamachi:
    IsKamachi = False
    '1.10.9 ADD END
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
'
'   1.10.12 20160322 K.Asayama Change
'           ��AF1�`AF3�����O�iF�y��)
'   1.10.19 K.Asayama Change
'           ��1608���7�^��Flush�i�K���X�j����
'   1.11.0
'           ���e���X�h�A(YG6�^)
'   2.5.2
'           ��YG6�^�̓X���[�K���X����O��
'   *************************************************************
    On Error GoTo Err_IsThruGlass
    
    IsThruGlass = False
    
    If IsNull(in_strHinban) Then Exit Function
     
    If in_strHinban Like "*-####S*-*" Or in_strHinban Like "*-####C*-*" Or in_strHinban Like "*-####D*-*" _
        Or in_strHinban Like "F?C??*-####A*-*" Or in_strHinban Like "F?C??*-####B*-*" Or in_strHinban Like "F?C??*-####O*-*" _
        Or in_strHinban Like "*ME-####M*-*" Or in_strHinban Like "*SA-####M*-*" Or IsVertica(in_strHinban) Or in_strHinban Like "F?C??*-####MF*-*" Then
        
        IsThruGlass = True
    'YG6�^
'    ElseIf in_strHinban Like "Y*-####T*-*" Then
'        IsThruGlass = True
        
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

'   1.10.6 K.Asayama 1610�d�l�i�B�����ԁj�ǉ�
'   *************************************************************

    IsOyatobira = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "*DO-####*" Or in_strHinban Like "*DOS-####*" _
       Or in_strHinban Like "*CO-####*" Or in_strHinban Like "*COS-####*" _
        Or in_strHinban Like "*KO-####*" Or in_strHinban Like "*KOS-####*" _
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

'   1.10.6 K.Asayama 1610�d�l�i�B�����ԁj�ǉ�
'   *************************************************************

    IsKotobira = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "*DK-####*" Or in_strHinban Like "*DKS-####*" _
            Or in_strHinban Like "*KK-####*" Or in_strHinban Like "*KKS-####*" _
                                                                                Then
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
'    Output����
'       out_strSxLhinban    �_�J�i��(False�̏ꍇ��Null)

'   1.10.6 K.Asayama SxL�R�s�[����̂ݎ��s�ɕύX�������ߖ{�����ɒǉ�
'   *************************************************************
    
    Dim objLOCALdb As New cls_LOCALDB
    Dim strHinban As String
    Dim bolMentori As Boolean
    
    IsSxL = False
    
    On Error GoTo Err_IsSxL
    
    If IsNull(in_strHinban) Then GoTo Exit_IsSxL

    '1.10.6 K.Asayama ADD 20161211********
    If Not fncbolSxL_Replace() Then
        MsgBox "SxL�i�ԃ}�X�^�̃R�s�[�Ɏ��s���܂���" & vbCrLf & "�l�b�g���[�N�ɖ�肪����ꍇ�͉񕜌�ēx���s���Ă�������"
        Err.Raise 9999, , "Quit"
    End If
    '*************************************
    
    
    '���n�Ŗʎ��L��������ꍇ�͊O��
    If in_strHinban Like "*�@?�A?�B?�C*" Then
        strHinban = left(in_strHinban, Len(in_strHinban) - 10)
        bolMentori = True
    Else
        strHinban = in_strHinban
        bolMentori = False
    End If
    '1.10.3 K.Asayama 20151119 SxL�i�ԓǑ֕\���[�J���e�[�u�����ύX
    If objLOCALdb.ExecSelect("select �u�����h�i�� from WK_SxL�i�ԓǑ֕\ where S�~L�i�� = '" & Trim(strHinban) & "'") Then
        If Not objLOCALdb.GetRS.EOF Then
            out_strKamiyahinban = objLOCALdb.GetRS![�u�����h�i��]
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
    Set objLOCALdb = Nothing
End Function

Public Function valfncHinmei(in_objRemoteDB As cls_BRAND_MASTER, in_RS As ADODB.Recordset, in_strHinban As Variant, in_intSeihinkubun As Integer, in_strSpec As Variant) As Variant
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
    
    
    If in_objRemoteDB.ExecSelect_ExternalRS(in_RS, strSQL) Then
        If Not in_RS.EOF Then
            valfncHinmei = in_RS![�i��]
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

'   1.10.6 K.Asayama 1610�d�l�iAF1�`AF3�j�ǉ�
'   1.10.19 K.Asayama Change
'           �����C���h�J�[�h������(_��?)
'   *************************************************************

    isCaro = False
    
    If in_varHinban Like "F?C*-####A*-*" Or in_varHinban Like "F?C*-####B*-*" Or in_varHinban Like "F?C*-####O*-*" _
        Or in_varHinban Like "�� F?C*-####A*-*" Or in_varHinban Like "�� F?C*-####B*-*" Or in_varHinban Like "�� F?C*-####O*-*" _
            Or in_varHinban Like "F?B*-####A*-*" Or in_varHinban Like "F?B*-####B*-*" Or in_varHinban Like "F?B*-####O*-*" _
                Or in_varHinban Like "�� F?B*-####A*-*" Or in_varHinban Like "�� F?B*-####B*-*" Or in_varHinban Like "�� F?B*-####O*-*" _
                                                                                                                                        Then
        
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

'1.11.0
'       �����ޕύX�ɑΉ�(�ꕔ�֐����j
'1.11.3
'       �����ޕύX�ɑΉ�
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
            
            '�֐���
'            If in_varHinban Like "F_V*-####*" Then 'Vertica
'                strfncSyobunrui_Kamui = "V"
'
'            ElseIf in_varHinban Like "F_C*-####*" Then 'Caro
'                strfncSyobunrui_Kamui = "A"

            If IsVertica(in_varHinban) Then  'Vertica
                strfncSyobunrui_Kamui = "V"

            ElseIf isCaro(in_varHinban) Then 'Caro
                strfncSyobunrui_Kamui = "A"
                
            Else
                Select Case strHinbanKigou
                    Case "F" '�W���i��CUBE�̃R�[�h�𑗂�i�������ꂽ�番����K�v����j
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
'   '1.10.4 Change by Asayama 20151207
'       ���S�ʉ����i���A���[�g�ɖ��h�����ł���̂ŐF�R�[�h�x�[�X�ɕύX�j
'
'   �߂�l:Boolean
'       ��True              �h����
'       ��False             �h�����ȊO
'
'    Input����
'       in_strHinban        ����i��

'   1.10.11 K.Asayama ADD
'           ���G�X�p�̃��A���[�g�͓h��
'   1.12.3
'           �����A���[�g�V�F�ǉ�
'   2.1.1
'           ��1801�V�F��s�ǉ�
'   *************************************************************
    On Error GoTo Err_IsPainted
    
    IsPainted = False
    
    If IsNull(in_strHinban) Then Exit Function
    
'    If in_strHinban Like "R*-####*-*" Or in_strHinban Like "�� R*-####*-*" Or in_strHinban Like "B*-####*-*" Or in_strHinban Like "�� B*-####*-*" Then
'        IsPainted = True
'    Else
'        IsPainted = False
'    End If
    
    If in_strHinban Like "*-####*-*(NW)*" Or in_strHinban Like "*-####*-*(NO)*" Or in_strHinban Like "*-####*-*(NC)*" Or in_strHinban Like "*-####*-*(NK)*" Or in_strHinban Like "*-####*-*(NA)*" Or in_strHinban Like "*-####*-*(NB)*" Then
        IsPainted = True
    Else
        IsPainted = False
    End If
    
    '1.12.3 ADD
    If in_strHinban Like "*-####*-*(NH)*" Then
        IsPainted = True
    End If
    
    If in_strHinban Like "*-####*-*(NT)*" Or in_strHinban Like "*-####*-*(NY)*" Then
        IsPainted = True
    End If
    
    '1.10.11 ADD
    If in_strHinban Like "*HER*-####*-*" Then
        IsPainted = True
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

'1.10.9 K.Asayama
'       �������J�l��DV�̓C���Z�b�g���n
'1.10.13 K.Asayama
'       ���G�X�p���n�i�Ԃ̓C���Z�b�g���n
'1.11.4 K.Asayama
'       ��1701�V�i�Ԓǉ�(VM)
'2.9.0
'       ��1808�V�i�Ԓǉ�(GU)
'   *************************************************************
    '
    IsStealth_Seizo_TEMP = False
    
    '1.10.13
    If in_strHinban Like "*PW*-####*" Then
        IsStealth_Seizo_TEMP = False
        Exit Function
    End If
    
    If (in_strHinban Like "*SG*-####*" Or in_strHinban Like "*NG*-####*" Or in_strHinban Like "*AG*-####*" Or in_strHinban Like "*BG*-####*") _
        And Not in_strHinban Like "*ML-####*" And Not in_strHinban Like "*MK-####*" And Not in_strHinban Like "*MT-####*" And Not in_strHinban Like "*DU-####*" And Not in_strHinban Like "*DN-####*" And Not in_strHinban Like "*VN-####*" And Not in_strHinban Like "*CTSG*MK-####*" And Not in_strHinban Like "*CTSG*ML-####*" And Not in_strHinban Like "*CTSG*MT-####*" And Not in_strHinban Like "*KU-####*" And Not in_strHinban Like "*KN-####*" And Not in_strHinban Like "*DV-####*" And Not in_strHinban Like "*GU-####*" Then
        IsStealth_Seizo_TEMP = True
    End If
    
End Function

Public Function fncbolSxL_Replace() As Boolean
'   *************************************************************
'   SxL�i�ԓǑ֕\�u��������
'   1.10.3 K.Asayama ADD 20151119 SxL�i�ԕ\�����[�g����R�s�[
'   1.10.6 K.Asayama ADD 20151211 �R�s�[�ς݂̏ꍇ(bolSxLCopy=True�j�͏������Ȃ�
'
'   �����[�g�f�[�^�x�[�X���烍�[�J����SxL�i�ԓǑ֕\���R�s�[����
'
'   �߂�l:Boolean
'       ��True              �u������
'       ��False             �u�����s
'
'   *************************************************************

    fncbolSxL_Replace = False
    
    If bolSxLCopy Then
        fncbolSxL_Replace = True
        Exit Function
    End If
    
    Dim objREMOTEdb As New cls_BRAND_MASTER
    Dim objLOCALdb As New cls_LOCALDB
    
    On Error GoTo Err_fncbolSxL_Replace
    
    Dim strSQL_Insert As String
    Dim strSQL As String
    strSQL_Insert = "Insert into WK_SxL�i�ԓǑ֕\(S�~L�i��,�u�����h�i��,DH,DW,CH) values ("
    
    '�H��p�R�s�[�iT_Calendar_�H��)
    If objLOCALdb.ExecSQL("delete from WK_SxL�i�ԓǑ֕\") Then
        strSQL = "select distinct [S�~L�i��],�u�����h�i��,DW,DH,CH from SxL�i�ԓǑ֕\ "
        If objREMOTEdb.ExecSelect(strSQL) Then
            Do While Not objREMOTEdb.GetRS.EOF
                If Not objLOCALdb.ExecSQL(strSQL_Insert & "'" & objREMOTEdb.GetRS![S�~L�i��] & "','" & objREMOTEdb.GetRS![�u�����h�i��] & "'," & objREMOTEdb.GetRS![DW] & "," & objREMOTEdb.GetRS![DH] & "," & objREMOTEdb.GetRS![CH] & ")") Then
                    Err.Raise 9999, , "SxL�i�ԓǑ֕\ ���[�J���R�s�[�G���["
                End If
                objREMOTEdb.GetRS.MoveNext
            Loop
        End If
    End If
    
    '1.10.6 K.Asayama ADD 20151211 �R�s�[�����̏ꍇ���ʃt���O��True�ɂ���
    bolSxLCopy = True
    
    fncbolSxL_Replace = True
    
    GoTo Exit_fncbolSxL_Replace
    
Err_fncbolSxL_Replace:
    MsgBox Err.Description
    
Exit_fncbolSxL_Replace:

    Set objREMOTEdb = Nothing
    Set objLOCALdb = Nothing
    
End Function

Public Function IsREALART(in_strHinban As Variant) As Boolean
'   *************************************************************
'   REALART�m�F
'   �T�u�t�H�[���̏����t��������̌Ăяo���ŏ��������ەs�v�ȌĂяo������������̂ŃG���[���W�b�N��ǉ�
'   '1.10.4 ADD by Asayama 20151207
'   �߂�l:Boolean
'       ��True              REALART
'       ��False             REALART�ȊO
'
'    Input����
'       in_strHinban        ����i��

'2.3.0�i�R�����g�̂ݑ}���j
'   ��1801�d�l�Ń��A���[�g�ƃ^���ɃV���[�Y�������ꂽ���V�X�e���̓��e��͈ꏏ�̕���
'     �s�����悢�̂ł��̊֐��̓^�����܂ނ悤�ɂ���
'   *************************************************************
    On Error GoTo Err_IsREALART
    
    IsREALART = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "R*-####*-*" Or in_strHinban Like "�� R*-####*-*" Then
        IsREALART = True
    Else
        IsREALART = False
    End If
    
    Exit Function
    
Err_IsREALART:
    IsREALART = False
    
End Function

Public Function IsPALIO(in_strHinban As Variant) As Boolean
'   *************************************************************
'   PALIO�m�F
'   �T�u�t�H�[���̏����t��������̌Ăяo���ŏ��������ەs�v�ȌĂяo������������̂ŃG���[���W�b�N��ǉ�
'   '1.10.4 ADD by Asayama 20151207
'   �߂�l:Boolean
'       ��True              PALIO
'       ��False             PALIO�ȊO
'
'    Input����
'       in_strHinban        ����i��

'   *************************************************************
    On Error GoTo Err_IsPALIO
    
    IsPALIO = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "B*-####*-*" Or in_strHinban Like "�� B*-####*-*" Then
        IsPALIO = True
    Else
        IsPALIO = False
    End If
    
    Exit Function
    
Err_IsPALIO:
    IsPALIO = False
    
End Function

Public Function fncvalDoorColor(inHinban As String) As Variant
'   *************************************************************
'   �F�m�F
'   �i�Ԃ���F��Ԃ��B�Ԃ��Ȃ��ꍇ�͋󗓂�Ԃ��iNull�ł͂Ȃ��j
'   '1.10.7 ADD by Asayama 20160108
'   �߂�l:Variant
'       ���F�R�[�h�i�F�R�[�h�������ꍇ�͋󗓁A�G���[�̏ꍇ��Null�j
'
'    Input����
'       inHinban            ����i��

'   *************************************************************
    Dim i As Integer
    Dim byteHirakiIchi As Byte
    Dim byteTojiIchi As Byte
    
    On Error GoTo Err_fncvalDoorColor
    
    fncvalDoorColor = Null
    
    byteTojiIchi = 0
    byteHirakiIchi = 0
    
    For i = Len(inHinban) To 1 Step -1
        If Mid(inHinban, i, 1) = ")" Then
            byteTojiIchi = i
        ElseIf Mid(inHinban, i, 1) = "(" Then
            byteHirakiIchi = i + 1
        End If
        
        If byteHirakiIchi <> 0 And byteTojiIchi <> 0 Then Exit For
        
    Next
        
    If byteHirakiIchi <> 0 And byteTojiIchi <> 0 And byteTojiIchi > byteHirakiIchi Then
        fncvalDoorColor = Mid(inHinban, byteHirakiIchi, byteTojiIchi - byteHirakiIchi)
    End If
    
    GoTo Exit_fncvalDoorColor
    
Err_fncvalDoorColor:
    fncvalDoorColor = Null
    MsgBox Err.Description, , "�i�Ԃ���F�R�[�h���擾�ł��܂���"
    
Exit_fncvalDoorColor:
    
End Function

Public Function fncIntHalfGlassMirror_Maisu(in_strHinban As Variant, in_Maisu As Integer) As Integer
'   *************************************************************
'   �������ŕБ��̂݃K���X�E�~���[�̕i�Ԋm�F���A�K���X������Ԃ�
'   �T�u�t�H�[���̏����t��������̌Ăяo���ŏ��������ەs�v�ȌĂяo������������̂ŃG���[���W�b�N��ǉ�
'1.10.10 ADD by Asayama
'   �߂�l:Integer
'       ���K���X������
'
'    Input����
'       in_strHinban        ����i��
'        in_Maisu �����
'   *************************************************************
    On Error GoTo Err_fncIntHalfGlassMirror_Maisu
    
    fncIntHalfGlassMirror_Maisu = in_Maisu
    
    If IsNull(in_strHinban) Then Exit Function
    
    '2�Ŋ���؂�Ȃ��ꍇ���̂܂ܕԂ�
    If in_Maisu Mod 2 <> 0 Then Exit Function
    
    If in_strHinban Like "*ME-####MR*-*" Or in_strHinban Like "*ME-####ML*-*" Then
        
        fncIntHalfGlassMirror_Maisu = in_Maisu / 2
    End If
    
    Exit Function
    
Err_fncIntHalfGlassMirror_Maisu:
    fncIntHalfGlassMirror_Maisu = in_Maisu
End Function

Public Function IsGranArt(in_strHinban As Variant) As Boolean
'   *************************************************************
'   �O�����A�[�g�m�F
'   �T�u�t�H�[���̏����t��������̌Ăяo���ŏ��������ەs�v�ȌĂяo������������̂ŃG���[���W�b�N��ǉ�
'   '1.10.16 ADD
'
'   �߂�l:Boolean
'       ��True              �O�����A�[�g
'       ��False             �O�����A�[�g�ȊO
'
'    Input����
'       in_strHinban        ����i��

'   *************************************************************
    On Error GoTo Err_IsGranArt
    
    IsGranArt = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "G*-####*-*" Or in_strHinban Like "�� G*-####*-*" Then
        IsGranArt = True
    Else
        IsGranArt = False
    End If
    
    Exit Function
    
Err_IsGranArt:
    IsGranArt = False
    
End Function
Public Function IsInset(in_strWakuHinban As Variant) As Boolean
'   *************************************************************
'   �C���Z�b�g�g�m�F
'   '1.10.16 ADD
'
'   �߂�l:Boolean
'       ��True              �C���Z�b�g�g
'       ��False             �C���Z�b�g�g�ȊO
'
'    Input����
'       in_strHinban        �g�i��

'1.11.1 Change K70�i�Ԃ�False�ɂȂ��Ă��܂����Ή�
'   *************************************************************
    On Error GoTo Err_IsInset
    
    IsInset = False

    If in_strWakuHinban Like "K##*-####*" Or in_strWakuHinban Like "�� K##*-####*" Then
        IsInset = True
    End If
    
    Exit Function

Err_IsInset:
    IsInset = False
End Function
Public Function IsHirakido(in_strHinban As Variant) As Boolean
'   *************************************************************
'   �J���ˊm�F�i�e�q�܂ށj
'   '1.10.16 ADD
'
'   �߂�l:Boolean
'       ��True              �J����
'       ��False             �J���ˈȊO
'
'    Input����
'       in_strHinban        ����i�g�A���n�j�i��
'   1.10.19 K.Asayama Change
'           ���B�����Ԑe�q�ǉ�
'   *************************************************************
    
    On Error GoTo Err_IsHirakido
    
    If in_strHinban Like "*CA-####*" Or in_strHinban Like "*CAS-####*" _
        Or in_strHinban Like "*DA-####*" Or in_strHinban Like "*DAS-####*" _
        Or in_strHinban Like "*PA-####*" Or in_strHinban Like "*PAS-####*" _
        Or in_strHinban Like "*KA-####*" Or in_strHinban Like "*KAS-####*" _
        Or in_strHinban Like "*DO-####*" Or in_strHinban Like "*DOS-####*" _
        Or in_strHinban Like "*DK-####*" Or in_strHinban Like "*DKS-####*" _
        Or in_strHinban Like "*KO-####*" Or in_strHinban Like "*KOS-####*" _
        Or in_strHinban Like "*KK-####*" Or in_strHinban Like "*KKS-####*" _
    Then
        
        IsHirakido = True
        
    Else
    
        IsHirakido = False
        
    End If
    
    Exit Function
    
Err_IsHirakido:
    IsHirakido = False
End Function

Public Function IsWallThru(in_strHinban As Variant) As Boolean
'   *************************************************************
'   �E�H�[���X���[�m�F
'   1.11.0 ADD
'
'   �߂�l:Boolean
'       ��True              WallThrough
'       ��False             WallThrough�ȊO
'
'    Input����
'       in_strHinban        ���n�i��

'   *************************************************************
    '
    IsWallThru = False

    If in_strHinban Like "*WS*-####*" Then
        IsWallThru = True
        Exit Function
    End If

    
End Function

Public Function IsTerrace(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �e���X�h�A�m�F
'
'   �߂�l:Boolean
'       ��True              Terrace
'       ��False             Terrace�ȊO
'
'    Input����
'       in_strHinban        ����i��

'   1.11.0 ADD
'   *************************************************************

    IsTerrace = False
    
    On Error GoTo Err_IsTerrace
        
    If IsNull(in_varHinban) Then
        Exit Function
    End If
    
    If in_varHinban Like "Y*-####*-*" Then
        
        IsTerrace = True
        
    End If
    
    Exit Function
    
Err_IsTerrace:
    IsTerrace = False
    
End Function

Public Function IsMirror(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �~���[���m�F
'
'   �߂�l:Boolean
'       ��True              �~���[
'       ��False             �~���[�ȊO
'
'    Input����
'       in_strHinban        ����i��

'   1.11.0 ADD
'   *************************************************************

    IsMirror = False
    
    On Error GoTo Err_IsMirror
        
    If IsNull(in_varHinban) Then
        Exit Function
    End If
    
    If in_varHinban Like "*-####M*-*" Then
        
        IsMirror = True
        
    End If
    
    Exit Function
    
Err_IsMirror:
    IsMirror = False
End Function

Public Function IsCloset_Hiraki(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �N���[�b�g�i�Ԋm�F�i�J���ˁj

'   �����^�ЊJ���i���n�g���p�j �X���C�h���[�͑ΏۂƂ��Ȃ�

'   �߂�l:Boolean
'       ��True              �N���[�b�g�J��
'       ��False             �N���[�b�g�J���ȊO
'
'    Input����
'       in_varHinban        ����i�ԁ^���n�i��

'   1.12.0 ADD
'   *************************************************************
    On Error GoTo Err_IsCloset_Hiraki
    
    IsCloset_Hiraki = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*MA-####*" Or in_varHinban Like "*MB-####*" Or in_varHinban Like "*MAS-####*" Or in_varHinban Like "*MBS-####*" Then
        IsCloset_Hiraki = True
    End If
    
    Exit Function

Err_IsCloset_Hiraki:
    IsCloset_Hiraki = False
End Function

Public Function IsCloset_Oredo(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �N���[�b�g�i�Ԋm�F�i�܂�ˁj


'   �߂�l:Boolean
'       ��True              �N���[�b�g�܂��
'       ��False             �N���[�b�g�܂�ˈȊO
'
'    Input����
'       in_varHinban        ����i�ԁ^���n�i��

'   1.12.0 ADD
'   *************************************************************
    On Error GoTo Err_IsCloset_Oredo
    
    IsCloset_Oredo = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*ML-####*" Or in_varHinban Like "*MK-####*" Or in_varHinban Like "*MT-####*" Then
        IsCloset_Oredo = True
    End If
    
    Exit Function

Err_IsCloset_Oredo:
    IsCloset_Oredo = False
End Function

Public Function IsSoftMotion(ByVal in_varHinban As Variant) As Boolean
'   *************************************************************
'   �\�t�g���[�V�����m�F

'   �߂�l:Boolean
'       ��True              �\�t�g���[�V��������
'       ��False             �\�t�g���[�V�����ȊO
'
'    Input����
'       in_varHinban        ����i�ԁ^���n�i��

'   1.12.0 ADD
'   *************************************************************
    IsSoftMotion = False
    
    If in_varHinban Like "*CA-####*" Or in_varHinban Like "*CO-####*" Or in_varHinban Like "*CAS-####*" Or in_varHinban Like "*COS-####*" Then
    
        IsSoftMotion = True
    
    End If
    

End Function

Public Function IsCloset_Slide(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �X���C�h���[�m�F

'   �߂�l:Boolean
'       ��True              �X���C�h���[
'       ��False             �X���C�h���[�ȊO
'
'    Input����
'       in_strHinban        ����i��,���͉��n�i��


'   1.12.0 ADD
'   *************************************************************
    
    Dim strHinban As String
    
    On Error GoTo Err_IsCloset_Slide
    
    IsCloset_Slide = False
    
    
    If IsNull(in_varHinban) Then Exit Function

    If in_varHinban Like "*SA-####*" Then
        IsCloset_Slide = True
    End If
    
    Exit Function

Err_IsCloset_Slide:
    IsCloset_Slide = False
    
End Function

Public Function IsYukazukeRail(in_varHinban As Variant) As Boolean
'   *************************************************************
'   ���t�����[���i�Ԋm�F

'   ����݂�A���͊܂܂Ȃ�

'   �߂�l:Boolean
'       ��True              ���t�����[��
'       ��False             ���t�����[���ȊO
'
'    Input����
'       in_varHinban        �i��

'   1.12.0 ADD
'   *************************************************************
    On Error GoTo Err_IsYukazukeRail
    
    IsYukazukeRail = False
    
    If in_varHinban Like "*DM-####*" Or in_varHinban Like "*DL-####*" Or in_varHinban Like "*DN-####*" Then
        IsYukazukeRail = True
    'V���[��
    ElseIf in_varHinban Like "*VM-####*" Or in_varHinban Like "*VL-####*" Or in_varHinban Like "*VN-####*" Then
        IsYukazukeRail = True
    End If
    
    Exit Function

Err_IsYukazukeRail:
    IsYukazukeRail = False
End Function

Public Function IsLUCENTE(in_varHinban As Variant) As Boolean
'   *************************************************************
'   ���`�F���e�m�F

'   �߂�l:Boolean
'       ��True              ���`�F���e
'       ��False             ���`�F���e�ȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.1.0 ADD
'   *************************************************************

    Dim strHinban As String
    
    IsLUCENTE = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "�� ", "")
    
    If strHinban Like "P*-####*-*" Then
        If strHinban Like "*(XW)" Or strHinban Like "*(XB)" Then
            IsLUCENTE = True
        End If
    End If
    
End Function

Public Function IsSINA(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �V�i�m�F
'   'ADD by Asayama 20150903
'   �߂�l:Boolean
'       ��True              �V�i�i��
'       ��False             �V�i�i�ԈȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.1.0 ADD
'   *************************************************************

    Dim strHinban As String
    
    IsSINA = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "�� ", "")
    
    If strHinban Like "T*-####*-*" Then
        If IsSINAColor(strHinban) Then
            IsSINA = True
        End If
    End If
    
End Function

Public Function IsSINAColor(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �V�i�F�m�F
'
'   �߂�l:Boolean
'       ��True              �F���V�i�F
'       ��False             �V�i�F�ȊO
'
'    Input����
'       in_strHinban        ����i��

'   2.1.0 ADD
'   *************************************************************

    IsSINAColor = False
    
    If IsNull(in_varHinban) Then Exit Function

    If in_varHinban Like "*-*-*(ZZ)" Or in_varHinban Like "*-*-*(AA)" Or in_varHinban Like "*-*-*(BB)" Or in_varHinban Like "*-*-*(CC)" Or in_varHinban Like "*-*-*(DD)" Then
        IsSINAColor = True
    Else
        IsSINAColor = False
    End If
    
End Function

Public Function IsFs(in_varHinban As Variant) As Boolean
'   *************************************************************
'   F/S�m�F

'   �߂�l:Boolean
'       ��True              F/S
'       ��False             F/S�ȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.1.0 ADD
'   *************************************************************

    Dim strHinban As String
    
    IsFs = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "�� ", "")
    
    If strHinban Like "S*-####*-*" Then
        IsFs = True
    End If
    
End Function

Public Function IsCloset_Hikichigai(in_varHinban As Variant) As Boolean
'   *************************************************************
'   ����������Ⴂ�m�F

'   �߂�l:Boolean
'       ��True              ����������Ⴂ
'       ��False             ����������Ⴂ�ȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.1.0 ADD
'   *************************************************************
    
    IsCloset_Hikichigai = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*ME-####*-*" Then
        IsCloset_Hikichigai = True
    End If
    
End Function

Public Function IsSideThru(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �T�C�h�X���[�m�F

'   �߂�l:Boolean
'       ��True              �T�C�h�X���[
'       ��False             �T�C�h�X���[�ȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.1.0 ADD

'2.3.0
'   ��1801�d�l�ǉ�
'   *************************************************************
    
    IsSideThru = False
    
    If IsNull(in_varHinban) Then Exit Function

    If in_varHinban Like "*-####ST*-*" Or in_varHinban Like "*-####SS*-*" Or in_varHinban Like "*-####SG*-*" Or in_varHinban Like "*-####SH*-*" Then
        IsSideThru = True
    End If
    
End Function

Public Function IsCenterThru(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �Z���^�[�X���[�m�F

'   �߂�l:Boolean
'       ��True              �Z���^�[�X���[
'       ��False             �Z���^�[�X���[�ȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.1.0 ADD
'   *************************************************************
    
    IsCenterThru = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*-####C*-*" Then
        IsCenterThru = True
    End If
    
End Function

Public Function IsWideThru(in_varHinban As Variant) As Boolean
'   *************************************************************
'   ���L�X���[�K���X�m�F

'   �߂�l:Boolean
'       ��True              ���L�Z���^�[�X���[
'       ��False             ���L�Z���^�[�X���[�ȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.1.0 ADD
'   *************************************************************
    
    IsWideThru = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*-####D*-*" Then
        IsWideThru = True
    End If
    
End Function

Public Function IsG7_Flush(in_varHinban As Variant) As Boolean
'   *************************************************************
'   G7�^(1608�d�l�ȍ~)�m�F

'   �߂�l:Boolean
'       ��True              G7�^
'       ��False             G7�^�ȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.1.0 ADD
'   *************************************************************

    Dim strHinban As String
    
    IsG7_Flush = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "�� ", "")
    
    If strHinban Like "??C*-####MF*" Then
        IsG7_Flush = True
    End If

End Function

Public Function IsHikido(ByVal in_varHinban As Variant) As Boolean
'   *************************************************************
'   ���ˊm�F

'   �߂�l:Boolean
'       ��True              ����
'       ��False             ���ˈȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.1.0 ADD

'2.3.0
'   ��1801�d�l�ǉ�
'2.7.0
'   ��1808�d�l�ǉ�
'   *************************************************************

    IsHikido = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*DC-####*" Or in_varHinban Like "*DT-####*" Or _
        in_varHinban Like "*KC-####*" Or in_varHinban Like "*KT-####*" Or _
        in_varHinban Like "*DM-####*" Or in_varHinban Like "*DL-####*" Or _
        in_varHinban Like "*DP-####*" Or in_varHinban Like "*DH-####*" Or _
        in_varHinban Like "*DE-####*" Or in_varHinban Like "*DJ-####*" Or _
        in_varHinban Like "*DF-####*" Or in_varHinban Like "*DQ-####*" Or _
        in_varHinban Like "*DU-####*" Or in_varHinban Like "*DN-####*" Or in_varHinban Like "*KU-####*" Or _
        in_varHinban Like "*DI-####*" Or in_varHinban Like "*DG-####*" Or _
        in_varHinban Like "*DD-####*" Or in_varHinban Like "*DV-####*" Or _
        in_varHinban Like "*VM-####*" Or in_varHinban Like "*VL-####*" Or in_varHinban Like "*VN-####*" Or _
        in_varHinban Like "*VF-####*" Or in_varHinban Like "*VQ-####*" Or _
        in_varHinban Like "*VI-####*" Or in_varHinban Like "*VG-####*" Or _
        in_varHinban Like "*JC-####*" Or in_varHinban Like "*GU-####*" Or _
        in_varHinban Like "*DY-####*" _
    Then

        IsHikido = True
        
    End If
    

End Function

Public Function IsKabetsukeGuide(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �Ǖt�K�C�h���ˊm�F

'   �߂�l:Boolean
'       ��True              �Ǖt�K�C�h����
'       ��False             �Ǖt�K�C�h���ˈȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.1.0 ADD
'   2.5.0
'       ���o�O�C�� KT��KU�̓���[*]�������Ă���
'   *************************************************************
    
    IsKabetsukeGuide = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*KC-####*-*" Or in_varHinban Like "*KT-####*-*" Or in_varHinban Like "*KU-####*-*" Then
        IsKabetsukeGuide = True
    End If
    
End Function

Public Function IsEndWakunashi(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �G���h�g�����m�F

'   �߂�l:Boolean
'       ��True              �G���h�g����
'       ��False             �G���h�g�����ȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.1.0 ADD
'   *************************************************************
    
    IsEndWakunashi = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*DU-####*" Or in_varHinban Like "*DN-####*" Or in_varHinban Like "*KU-####*" Or in_varHinban Like "*VN-####*" Then
        IsEndWakunashi = True
    End If
    
End Function

Public Function IsCaro_Panel(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �J���iAF-1�`3�^�@�p�l���ˁj�m�F

'   �߂�l:Boolean
'       ��True              Caro�i�p�l���j
'       ��False             Caro�i�p�l���j�ȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.1.0 ADD
'   *************************************************************
    
    Dim strHinban As String

    IsCaro_Panel = False
    
    If IsNull(in_varHinban) Then Exit Function

    strHinban = Replace(in_varHinban, "�� ", "")
    
    If strHinban Like "??B*-####A*-*" Or strHinban Like "??B*-####B*-*" Or strHinban Like "??B*-####O*-*" Then
        IsCaro_Panel = True
    End If
    
End Function

Public Function IsTerraceGlass(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �e���X�K���X�h�A�m�F

'   �߂�l:Boolean
'       ��True              �e���X�K���X�h�A�^
'       ��False             �e���X�K���X�h�A�^�ȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.1.0 ADD
'   *************************************************************

    Dim strHinban As String

    IsTerraceGlass = False
    
    If IsNull(in_varHinban) Then Exit Function

    strHinban = Replace(in_varHinban, "�� ", "")
    
    If strHinban Like "Y*-####?A*" Or strHinban Like "Y*-####?C*" Or strHinban Like "Y*-####?D*" Or strHinban Like "Y*-####?P*" Or strHinban Like "Y*-####?V*" Then
        IsTerraceGlass = True
    End If

End Function

Public Function IsHidden_Hinge(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �u�B�����ԁv�m�F

'   �߂�l:Boolean
'       ��True              �B������
'       ��False             �B�����ԂłȂ�
'
'    Input����
'       in_varHinban        ����i��

'   2.1.0 ADD
'   *************************************************************

    IsHidden_Hinge = False
      
     If IsNull(in_varHinban) Then Exit Function
     
    If in_varHinban Like "*KA-####*" Or in_varHinban Like "*KAS-####*" Or in_varHinban Like "*KO-####*" Or in_varHinban Like "*KOS-####*" Or in_varHinban Like "*KK-####*" Or in_varHinban Like "*KKS-####*" Then
        IsHidden_Hinge = True
    End If
    
End Function

Public Function IsYG6(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �e���X�ʉ��K���X�h�A�m�F

'   �߂�l:Boolean
'       ��True              �e���X�ʉ��K���X�h�A
'       ��False             �e���X�ʉ��K���X�h�A�ȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.5.2 ADD
'   *************************************************************

    Dim strHinban As String

    IsYG6 = False
    
    If IsNull(in_varHinban) Then Exit Function

    strHinban = Replace(in_varHinban, "�� ", "")
    
    If strHinban Like "Y*-####T*" Then
        IsYG6 = True
    End If

End Function

Public Function IsPALIOBlack(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �p���I�u���b�N�i�r�A���R�j�m�F
'   �Ƃ肠�������~�ɂȂ����̂�False�݂̂�Ԃ�

'   �߂�l:Boolean
'       ��True              �p���I�u���b�N
'       ��False             �p���I�u���b�N�ȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.1.0 ADD
'   *************************************************************
'    If in_varHinban Like "*-*-*(NN)" Then
'        IsPALIOBlack = True
'    Else
        IsPALIOBlack = False
'    End If
    
End Function

Public Function IsTateguInset(in_varHinban As Variant) As Boolean
'   *************************************************************
'   ����i�Ԃ̘g�d�l���C���Z�b�g���m�F


'   �߂�l:boolen
'       ��True �g�d�l���C���Z�b�g
'
'    Input����
'       in_varHinban        �i��

'   2.1.0 ADD
'   *************************************************************
    Dim strHinban As String
    
    IsTateguInset = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "�� ", "")
    
    If strHinban Like "?Z*-####*" Or strHinban Like "?Y*-####*" Or strHinban Like "?T*-####*" Then
        IsTateguInset = True
    End If
    
End Function

Public Function IsGuidePiece_ShitaanaKakou(in_varHinban As Variant, in_varTobiraIchi As Variant, in_varSpec As Variant, Optional in_SekkeiBikou As Variant) As Boolean
'   *************************************************************
'   �K�C�h�s�[�X�������H�m�F

'   �߂�l:Boolean
'       ��True              �������H����
'       ��False             �������H�Ȃ�
'
'    Input����
'       in_varHinban        ����i��
'       in_varTobiraIchi    ���ʒu�i�E�A���A���j�A����ȊO�͖�������False��Ԃ�
'       in_varSpec          ��Spec 20160923���_�ł͎g�p���Ȃ�
'       in_SekkeiBikou      ����݌v���l
       
'   2.1.0 ADD
'   *************************************************************
    
    Dim strTsurimoto As String
    
    On Error GoTo Err_IsGuidePiece_ShitaanaKakou
    
    If in_varTobiraIchi <> "�E" And in_varTobiraIchi <> "��" And in_varTobiraIchi <> "��" Then
        Err.Raise 9999, , "ErrEnd"
    End If
    
    If IsNull(in_varHinban) Then
        Err.Raise 9999, , "ErrEnd"
    End If
    
    '�G�X�p�X���C�h�E�H�[���͏��O
    If IsSlideWall_Espacio(in_varHinban) Then
        IsGuidePiece_ShitaanaKakou = False
        Exit Function
    End If
    
    '�݌v���l�Ɉȉ��R�����g������ꍇ�͏��O
    If Not IsMissing(in_SekkeiBikou) Then
        If in_SekkeiBikou Like "*�}�ʂ���*�ˎ�ˎ�*" Then
            IsGuidePiece_ShitaanaKakou = False
            Exit Function
        End If
    End If
    
    If in_varHinban Like "*DF-####*-*" Or in_varHinban Like "*VF-####*-*" Then

    
        If in_varTobiraIchi = "��" Then
            IsGuidePiece_ShitaanaKakou = True
        End If
    ElseIf in_varHinban Like "*DH-####*-*" Then
        strTsurimoto = Mid(in_varHinban, InStr(1, in_varHinban, "(") - 1, 1)
        If strTsurimoto = "L" And in_varTobiraIchi = "�E" Then
            IsGuidePiece_ShitaanaKakou = True
        ElseIf strTsurimoto = "R" And in_varTobiraIchi = "��" Then
            IsGuidePiece_ShitaanaKakou = True
        End If
    ElseIf in_varHinban Like "*DJ-####*-*" Then
        If in_varTobiraIchi = "��" Then
            IsGuidePiece_ShitaanaKakou = True
        Else
            strTsurimoto = Mid(in_varHinban, InStr(1, in_varHinban, "(") - 1, 1)
            If strTsurimoto = "L" And in_varTobiraIchi = "�E" Then
                IsGuidePiece_ShitaanaKakou = True
            ElseIf strTsurimoto = "R" And in_varTobiraIchi = "��" Then
                IsGuidePiece_ShitaanaKakou = True
            End If
        End If
    End If
    
    Exit Function
    
Err_IsGuidePiece_ShitaanaKakou:
    IsGuidePiece_ShitaanaKakou = False
    
End Function

Public Function IsSlideWall_Espacio(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �X���C�h�E�H�[���i�G�X�p�j�m�F

'   �߂�l:Boolean
'       ��True              �X���C�h�E�H�[��
'       ��False             �X���C�h�E�H�[���ȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.1.0 ADD
'   *************************************************************

    On Error GoTo Err_IsSlideWall_Espacio
    
    Dim strHinban As String
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "�� ", "")
    
    If strHinban Like "PSW*-####FV-*" Or strHinban Like "ESW*-####FV-*" Then
        IsSlideWall_Espacio = True
    Else
        IsSlideWall_Espacio = False
    End If
    
    Exit Function

Err_IsSlideWall_Espacio:
    IsSlideWall_Espacio = False
    
End Function

Public Function dblfncGuidePiece_ShitaanaSunpo(in_varHinban As Variant, in_varSpec As Variant) As Double
'   *************************************************************
'   �K�C�h�s�[�X�������H���@

'   �߂�l:Double
'       ���������H���@�i�Y�����Ȃ��ꍇ��0��Ԃ��j
'
'    Input����
'       in_varHinban        ����i��
'       in_varSpec          ��Spec 20160923���_�ł͎g�p���Ȃ�

'   2.1.0 ADD
'   *************************************************************

    On Error GoTo Err_dblfncGuidePiece_ShitaanaSunpo
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*DF-####*-*" Or in_varHinban Like "*VF-####*-*" Then
    
        dblfncGuidePiece_ShitaanaSunpo = 60
    ElseIf in_varHinban Like "*DH-####*-*" Then
        dblfncGuidePiece_ShitaanaSunpo = 52.5
    ElseIf in_varHinban Like "*DJ-####*-*" Then
        dblfncGuidePiece_ShitaanaSunpo = 52.5
    Else
        dblfncGuidePiece_ShitaanaSunpo = 0
    End If
    
    Exit Function
    
Err_dblfncGuidePiece_ShitaanaSunpo:
    dblfncGuidePiece_ShitaanaSunpo = 0
End Function

Public Function intfncGuidePiece_ShitaanaSu(in_varHinban As Variant, in_varSpec As Variant) As Integer
'   *************************************************************
'   �K�C�h�s�[�X�������H��

'   �߂�l:Integer
'       ���������H���i�Y�����Ȃ��ꍇ��0��Ԃ��j
'
'    Input����
'       in_varHinban        ����i��
'       in_varSpec          ��Spec 20160923���_�ł͎g�p���Ȃ�
        
'   2.1.0 ADD
'   *************************************************************
    
    On Error GoTo Err_intfncGuidePiece_ShitaanaSu
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*DF-####*-*" Or in_varHinban Like "*VF-####*-*" Then
    
        intfncGuidePiece_ShitaanaSu = 2
    ElseIf in_varHinban Like "*DH-####*-*" Then
        intfncGuidePiece_ShitaanaSu = 1
    ElseIf in_varHinban Like "*DJ-####*-*" Then
        intfncGuidePiece_ShitaanaSu = 1
    Else
        intfncGuidePiece_ShitaanaSu = 0
    End If
    
    Exit Function
    
Err_intfncGuidePiece_ShitaanaSu:
    intfncGuidePiece_ShitaanaSu = 0

End Function

Public Function IsMirrorUsed(in_varHinban As Variant, Optional in_Tobiraichi As Variant = Null) As Boolean
'   *************************************************************
'   ���g�p���m�F

'   �������g�p����i�Ԃ��m�F

'   �߂�l:Boolean
'       ��True              ���g�p
'       ��False             �����g�p
'
'    Input����
'       in_varHinban        ����i��
'       in_Tobiraichi       ���ʒu(L or R or C or LC or RC) --�I�v�V�����i�������Ȃ��ꍇ�͕i�Ԃ݂̂Ŕ��f�j--201708���_�ł͎g�p���Ȃ�

'   *************************************************************
    Dim strHinban As String
    
    On Error GoTo Err_IsMirrorUsed
    
    IsMirrorUsed = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "�� ", "")
            
    '���ʒu�Ɋ֌W�Ȃ�True�̕i��
    If strHinban Like "*-####MF*" Or strHinban Like "*-####MM*" Then
        
        IsMirrorUsed = True
        
    '���ʒu�ɂ���Ĕ�������ꍇ
    ElseIf strHinban Like "*-####ML*" Or strHinban Like "*-####MR*" Then
        
        '���ʒu�̎w�����Ȃ��ꍇ��True
        If Nz(in_Tobiraichi, "") = "" Then
            IsMirrorUsed = True
        Else
            Select Case in_Tobiraichi
                Case "L"
                    If strHinban Like "*-####ML*" Then IsMirrorUsed = True
                        
                Case "R"
                    If strHinban Like "*-####MR*" Then IsMirrorUsed = True
            End Select
        End If
    End If
    
    'SxL����
    '1701�d�l���_�ł͂Ȃ�

    Exit Function

Err_IsMirrorUsed:
    IsMirrorUsed = False
End Function

Public Function valfncHinmei_Local(in_strHinban As Variant, in_intSeihinkubun As Integer, in_strSpec As Variant) As Variant
'   *************************************************************
'   �i�����o�i�i����Ԃ��̂ݔŁj

'   �߂�l:Variant �� �i���i������Ȃ��ꍇ��NULL�j
'
'    Input����
'       in_strHinban        ����i��
'       in_intSeihinkubun   �i�ԋ敪
'       in_strSpec          ��Spec
'   *************************************************************
    Dim objREMOTEdb As cls_BRAND_MASTER
    
    Dim strSQL As String
    Dim strHinban As String
    
    strSQL = ""
    valfncHinmei_Local = Null
    
    On Error GoTo Err_valfncHinmei_Local
    
    If IsNull(in_strHinban) Then GoTo Exit_valfncHinmei_Local
    
    strHinban = Replace(in_strHinban, "�� ", "")
    
    Select Case in_intSeihinkubun
        Case 1, 5 '����,�۾ޯ�
            strSQL = "select top 1 �i�� from T_����i��Ͻ� where "
                If IsKotobira(strHinban) Then
                    strSQL = strSQL & " �q���i�� = '" & strHinban & "'"
                Else
                    strSQL = strSQL & " ����i�� = '" & strHinban & "'"
                End If
        Case 2, 4 '�g,�O���g
            strSQL = "select top 1 �i�� from T_�g�i��Ͻ� where �g�i�� = '" & strHinban & "'"
            
        Case 3 '���n�g
            strSQL = "select top 1 �i�� from T_���n�ޕi��Ͻ� where ���n�ޕi�� = '" & strHinban & "'"
          
        Case 6 '�����
            strSQL = "select top 1 �i�� from T_����ޕi��Ͻ� where ����ޕi�� = '" & strHinban & "'"
            
        Case 7 '���֎��[
            strSQL = "select top 1 �i�� from T_���֎��[Ͻ� where �i�� = '" & strHinban & "'"
            
        Case 8 '����
            strSQL = "select top 1 �i�� from T_�����i��Ͻ� where �����i�� = '" & strHinban & "'"
        
    End Select
    
    If strSQL = "" Then
        GoTo Exit_valfncHinmei_Local
    Else

        If Not IsNull(in_strSpec) And in_strSpec <> "" Then
            strSQL = strSQL & " and �d�l = '" & left(in_strSpec, 3) & "' and '" & right(in_strSpec, 4) & "' between �J�n and �I�� "
        End If

    End If
    
    With objREMOTEdb
        If .ExecSelect(strSQL) Then
            If Not .GetRS.EOF Then
                valfncHinmei_Local = .GetRS![�i��]
            End If
        End If
    End With
    
    GoTo Exit_valfncHinmei_Local
    
Err_valfncHinmei_Local:
    'MsgBox Err.Description
Exit_valfncHinmei_Local:
    Set objREMOTEdb = Nothing
End Function

Public Function IsTerraceKamachi(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �e���X�y(YF5/YG5)�h�A�m�F

'   �߂�l:Boolean
'       ��True              �e���X�h�A�^
'       ��False             �e���X�h�A�^�ȊO
'
'    Input����
'       in_strHinban        ����i��

'   2.1.0 ADD
'   *************************************************************

    On Error GoTo Err_IsTerraceKamachi
    
    Dim strHinban As String
    
    IsTerraceKamachi = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "�� ", "")
    
    If strHinban Like "Y?B*-####*" Then
        IsTerraceKamachi = True
    Else
        IsTerraceKamachi = False
    End If
    
    Exit Function

Err_IsTerraceKamachi:
    IsTerraceKamachi = False
End Function

Public Function IsG9(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �f�X�^�i�מy�ˁj�m�F
'
'   �߂�l:Boolean
'       ��True              G9�^
'       ��False             G9�^�ȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.3.0 ADD
'   *************************************************************

    Dim strHinban As String
    
    On Error GoTo Err_IsG9
    
    IsG9 = False
    
    If IsNull(in_varHinban) Then Exit Function

    strHinban = Replace(in_varHinban, "�� ", "")
    
    If strHinban Like "??B*-####E*-*" Then
        IsG9 = True
    End If
    
    Exit Function

Err_IsG9:
    IsG9 = False
    
End Function

Public Function IsTamo(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �^���V���[�Y�iJF1,JG1,JG2)�m�F
'
'   �߂�l:Boolean
'       ��True              �^���V���[�Y
'       ��False             �^���V���[�Y�ȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.3.0 ADD
'   *************************************************************
    Dim strHinban As String
    
    On Error GoTo Err_IsTamo
    
    IsTamo = False
    
    If IsNull(in_varHinban) Then Exit Function

    strHinban = Replace(in_varHinban, "�� ", "")
    
    If strHinban Like "R*-####*-*(NT)*" Or strHinban Like "R*-####*-*(ZT)*" Then
        IsTamo = True
    End If
    
    Exit Function

Err_IsTamo:
    IsTamo = False
    
End Function

Public Function IsRendouTategu(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �A������m�F
'
'   �߂�l:Boolean
'       ��True              �A������i�K�C�h�s�[�X�p����������j
'       ��False             �A������ȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.5.0 ADD
'   *************************************************************
    Dim strHinban As String
    
    On Error GoTo Err_IsRendouTategu
    
    IsRendouTategu = False
    
    If IsNull(in_varHinban) Then Exit Function

    strHinban = Replace(in_varHinban, "�� ", "")
    
    If strHinban Like "*VF-####*-*" Or strHinban Like "*DF-####*-*" Or strHinban Like "*DH-####*-*" Or strHinban Like "*DJ-####*-*" Then
        IsRendouTategu = True
    End If
    
    Exit Function

Err_IsRendouTategu:
    IsRendouTategu = False

End Function

Public Function IsHiRendouTategu(in_varHinban As Variant) As Boolean
'   *************************************************************
'   ��A������m�F
'
'   �߂�l:Boolean
'       ��True              ��A������
'       ��False             ��A������ȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.5.0 ADD
'   *************************************************************
    Dim strHinban As String
    
    On Error GoTo Err_IsHiRendouTategu
    
    IsHiRendouTategu = False
    
    If IsNull(in_varHinban) Then Exit Function

    strHinban = Replace(in_varHinban, "�� ", "")
    
    If strHinban Like "*DQ-####*-*" Or strHinban Like "*VQ-####*-*" Then
        IsHiRendouTategu = True
    End If
    
    Exit Function

Err_IsHiRendouTategu:
    IsHiRendouTategu = False

End Function

Public Function IsKousi(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �i�q���m�F
'
'   �߂�l:Boolean
'       ��True              �i�q��
'       ��False             �i�q���ȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.5.2 ADD
'   *************************************************************
    Dim strHinban As String
    
    On Error GoTo Err_IsKousi
    
    IsKousi = False
    
    If IsNull(in_varHinban) Then Exit Function

    strHinban = Replace(in_varHinban, "�� ", "")
    
    If strHinban Like "Z?B*-####*-*" Then
        IsKousi = True
    End If
    
    Exit Function

Err_IsKousi:
    IsKousi = False
    
End Function

Public Function IsReversible(in_varHinban As Variant, varTateguSekkeiBikou As Variant, varSpec As Variant) As Boolean
'   *************************************************************
'   ���o�[�V�u���m�F
'
'   �߂�l:Boolean
'       ��True                  ���o�[�V�u��
'       ��False                 ���o�[�V�u���ȊO
'
'    Input����
'       in_varHinban            ����i��
'       varTateguSekkeiBikou    ����݌v���l
'       varSpec                 ��Spec

'   2.5.3 ADD
'   2.7.0
'       ��ME��1808�ɂă��o�[�V�u���p�~
'   *************************************************************

    Dim strTateguSekkeiBikou As String
    
    On Error GoTo Err_IsReversible
    
    IsReversible = False

    If IsNull(in_varHinban) Or IsNull(varSpec) Then
        Exit Function
    End If
    
    strTateguSekkeiBikou = Nz(varTateguSekkeiBikou, "")


    '   *************************************************************
    '   ����݌v���l�Ɂu���o�[�V�u���v���܂܂�Ă���ꍇ
    '   ���o�[�V�u������
    '   *************************************************************
    
    If strTateguSekkeiBikou Like "*���o�[�V�u��*" Or strTateguSekkeiBikou Like "*��ް����*" Then
        IsReversible = True
        
    '   *************************************************************
    '   F/S�V���[�Y
    '   ZZ�F(KF1)�̓��o�[�V�u���łȂ�
    '   *************************************************************
    
    ElseIf IsFs(CStr(in_varHinban)) Then
        If in_varHinban Like "*(ZZ)" Then 'KF1�^
            Exit Function
        Else
            IsReversible = True 'KF7�^�̓��o�[�V�u��
        End If
        
    '   *************************************************************
    '   ����������Ⴂ
    '   PH�F�̓��o�[�V�u���łȂ�
    '   �~���[�I�v�V������PH�ł����o�[�V�u������
    '   *************************************************************
    
    ElseIf IsCloset_Hikichigai(CStr(in_varHinban)) Then '���������Ⴂ

        
        '�~���[�̓��o�[�V�u���i�������邽�߁j
        If in_varHinban Like "*-####M*" Then
            
            IsReversible = True
                
        '1808�d�l�ȍ~�̓��o�[�V�u���łȂ�
        ElseIf right(varSpec, 4) >= "1808" Then
            
            IsReversible = False
            
        '1801�ȑO�ł����̓��o�[�V�u���łȂ�
        ElseIf in_varHinban Like "*-####*-*(PH)" Or (in_varHinban Like "*-####*-*(SH)" And right(varSpec, 4) >= "1701") Then

            IsReversible = False
        
        '1801�ȑO�̎c��S�Ă̓��o�[�V�u��
        Else
            IsReversible = True

        End If
        
    End If
    
    Exit Function
    
Err_IsReversible:
    IsReversible = False
    
End Function

Public Function IsFullGlass(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �u�f�P�^�i�t���K���X�j�m�F
'
'   �߂�l:Boolean
'       ��True              VG1�^
'       ��False             VG1�^�ȊO
'
'    Input����
'       in_varHinban        ����i��

'   2.7.0 ADD
'   *************************************************************

    Dim strHinban As String
    
    On Error GoTo Err_IsFullglass
    
    IsFullGlass = False
    
    If IsNull(in_varHinban) Then Exit Function

    strHinban = Replace(in_varHinban, "�� ", "")
    
    If strHinban Like "X*-####X*-*" Then
        IsFullGlass = True
    End If
    
    Exit Function

Err_IsFullglass:
    IsFullGlass = False
    
End Function