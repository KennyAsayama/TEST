Option Compare Database
Option Explicit
                
Public Function bolfncUwawakuShitaji_GroupADD() As Boolean
'   *************************************************************
'   ��g���H�L���m�F
'   20170301 K.Asayama ADD
'
'   ��g���n�J�b�g�\�ɏW�v�O���[�v�ƌ�����ǉ�����

'   1.12.1
'       ���o�O�C��(SQL���� like������ *[�A�X�^���X�N]���g�p���Ă������� %[�p�[�Z���g]�ɕύX
'       ���}�ʂ���̎��ɏ�g���n������i�Ԃ��m�F����
'   1.12.2
'        ��Err.Raise���Ɉ����̐����Ԉ���Ă���Ƃ�����C��
'   *************************************************************
    
    Dim cnADO As ADODB.Connection
    Dim rsADO As ADODB.Recordset
    Dim strSQL As String
    Dim intGroupName As Integer
    Dim dblcutLength() As Double
    Dim dblGroupSu As Double
    Dim i As Integer
    Dim bolStealth As Boolean
    Dim varHinban As Variant
    Dim bolToku As Boolean
    
    Set cnADO = CurrentProject.Connection
    Set rsADO = New ADODB.Recordset
    
    bolfncUwawakuShitaji_GroupADD = False
    
    On Error GoTo Err_fncbolUwawakuShitaji_GroupADD
    
'    DoCmd.RunSQL "delete from WK_��ٽ��g���n�W�v�\"
'    DoCmd.RunSQL "delete from WK_�ݾ�ď�g���n�W�v�\"
    DoCmd.RunSQL "delete from WK_��g���n�O���[�v�W�v"
    
    strSQL = ""
    strSQL = strSQL & "select S.*,(���� - �݌Ɉ�����) AS ���H�� from TMP_�����w�������n�� S "
    strSQL = strSQL & "inner join TMP_�������� T "
    strSQL = strSQL & "on S.�_��ԍ� = T.�_��ԍ� and S.���ԍ� = T.���ԍ� and S.�����ԍ� = T.�����ԍ� and S.�� = T.�� "
    strSQL = strSQL & "where (���� - �݌Ɉ�����) > 0 and ���n�ތ��iFLG = False "
    strSQL = strSQL & "and (���n�g�݌v���l like '%�}�ʂ���%' or ��g���nW <> 0) "
    
    rsADO.Open strSQL, cnADO, adOpenStatic, adLockPessimistic
    
    Do Until rsADO.EOF
        strSQL = ""
        dblGroupSu = 0
        bolStealth = False
        
        '�i�Ԏ擾
        bolToku = bolFncTokuHinban(rsADO![���n�ޕi��], rsADO![�������n�ޕi��], varHinban)
        
        '�X�e���X�m�F
        If IsStealth_Seizo_TEMP(varHinban) Then bolStealth = True
        
        If Not Nz(rsADO![���n�g�݌v���l], "") Like "*�}�ʂ���*" Then
            intGroupName = intFncUwawakuShitajiLengthGroup(rsADO![��g���nW])
            
            If bolfncUwawakuShitajiLength(rsADO![��g���nW], dblcutLength()) Then
                For i = 1 To 3
                    If dblcutLength(i) > 0 Then
                        strSQL = ""
                        strSQL = strSQL & "insert into WK_��g���n�O���[�v�W�v(  "
                        strSQL = strSQL & "�_��ԍ�,���ԍ�,�����ԍ�,�@��,��,���n�ޕi��,�J����,�N���[�[�b�g,�X�e���X"
                        strSQL = strSQL & ",���i�� "
                        strSQL = strSQL & ",��g���nW,�������g���nW,��g���nW�����O���[�v,�㑍��,��Œ�l��,��ϓ��l��,����,����,�J�b�g�㐔��,���l,���͏� "
                        strSQL = strSQL & ") values ( "
                        strSQL = strSQL & varNullChk(rsADO![�_��ԍ�], 1)
                        strSQL = strSQL & "," & varNullChk(rsADO![���ԍ�], 1) & " "
                        strSQL = strSQL & "," & varNullChk(rsADO![�����ԍ�], 1) & " "
                        strSQL = strSQL & "," & varNullChk(rsADO![������], 1) & " "
                        strSQL = strSQL & "," & varNullChk(rsADO![��], 1) & " "
                        If bolToku Then
                            strSQL = strSQL & "," & varNullChk(rsADO![�������n�ޕi��], 1) & " "
                        Else
                        
                            strSQL = strSQL & "," & varNullChk(rsADO![���n�ޕi��], 1) & " "
                        End If
                        If IsHirakido(varHinban) _
                            Or IsOyatobira(varHinban) _
                                Or IsCloset_Hiraki(varHinban) Then
                            strSQL = strSQL & "," & True & " "
                        Else
                            strSQL = strSQL & "," & False & " "
                        End If
                        
                        If IsCloset_Hiraki(varHinban) _
                            Or IsCloset_Oredo(varHinban) Then
                                strSQL = strSQL & "," & True & " "
                        Else
                                strSQL = strSQL & "," & False & " "
                        End If
                        strSQL = strSQL & "," & bolStealth & " "
                        strSQL = strSQL & "," & varNullChk(rsADO![���i��], 1) & " "
                        strSQL = strSQL & "," & varNullChk(rsADO![��g���nW], 1) & " "
                        strSQL = strSQL & "," & varNullChk(dblcutLength(i), 1) & " "
                        strSQL = strSQL & "," & IIf(intFncUwawakuShitajiLengthGroup(dblcutLength(i)) > 0, CStr(intFncUwawakuShitajiLengthGroup(dblcutLength(i))), Null)
                        strSQL = strSQL & "," & varNullChk(rsADO![��g���n����], 1) & " "
                        strSQL = strSQL & "," & varNullChk(rsADO![��g���n�Œ�l��], 1) & " "
                        strSQL = strSQL & "," & varNullChk(rsADO![��g���n�ϓ��l��], 1) & " "
                        strSQL = strSQL & "," & varNullChk(fncstrUwawakuShitajiT(varHinban, rsADO![�������], rsADO![�{�[�h��]), 1) & " "

                        If i = 1 Then
                            strSQL = strSQL & "," & varNullChk(rsADO![���H��], 1) & " "
                        Else
                            strSQL = strSQL & ",0 "
                        End If
                        
                        If dblcutLength(i) > 300 And dblcutLength(i) <= 900 Then
                            strSQL = strSQL & "," & varNullChk(rsADO![���H��] * 0.5, 1) & " "
                        Else
                            strSQL = strSQL & "," & varNullChk(rsADO![���H��], 1) & " "
                        End If
                        strSQL = strSQL & ",Null "
                        strSQL = strSQL & "," & varNullChk(rsADO![���͏�], 1) & " "
                        strSQL = strSQL & ") "
                        
                        cnADO.Execute strSQL
                        
                    Else
                        Exit For
                    End If
                Next
            Else
                Err.Raise 9999, , "�ȍ~�̒��[�̏o�͂𒆎~���܂�" '1.12.2
            End If
        Else '�}�ʂ���
            If fncstrUwawakuShitajiT(varHinban, rsADO![�������], rsADO![�{�[�h��]) <> "" Then
                strSQL = ""
                strSQL = strSQL & "insert into WK_��g���n�O���[�v�W�v(  "
                strSQL = strSQL & "�_��ԍ�,���ԍ�,�����ԍ�,�@��,��,���n�ޕi��,�J����,�N���[�[�b�g,�X�e���X"
                strSQL = strSQL & ",���i�� "
                strSQL = strSQL & ",��g���nW,�������g���nW,��g���nW�����O���[�v,�㑍��,��Œ�l��,��ϓ��l��,����,����,�J�b�g�㐔��,���l,���͏� "
                strSQL = strSQL & ") values ( "
                strSQL = strSQL & varNullChk(rsADO![�_��ԍ�], 1)
                strSQL = strSQL & "," & varNullChk(rsADO![���ԍ�], 1) & " "
                strSQL = strSQL & "," & varNullChk(rsADO![�����ԍ�], 1) & " "
                strSQL = strSQL & "," & varNullChk(rsADO![������], 1) & " "
                strSQL = strSQL & "," & varNullChk(rsADO![��], 1) & " "
                If bolToku Then
                    strSQL = strSQL & "," & varNullChk(rsADO![�������n�ޕi��], 1) & " "
                Else
                
                    strSQL = strSQL & "," & varNullChk(rsADO![���n�ޕi��], 1) & " "
                End If
                If IsHirakido(varHinban) _
                    Or IsOyatobira(varHinban) _
                        Or IsCloset_Hiraki(varHinban) Then
                    strSQL = strSQL & "," & True & " "
                Else
                    strSQL = strSQL & "," & False & " "
                End If
                
                If IsCloset_Hiraki(varHinban) _
                    Or IsCloset_Oredo(varHinban) Then
                        strSQL = strSQL & "," & True & " "
                Else
                        strSQL = strSQL & "," & False & " "
                End If
                strSQL = strSQL & "," & bolStealth & " "
                strSQL = strSQL & "," & varNullChk(rsADO![���i��], 1) & " "
                strSQL = strSQL & "," & "0,0,Null,0,0,0,Null "
                strSQL = strSQL & "," & varNullChk(rsADO![���H��], 1) & " "
                strSQL = strSQL & "," & varNullChk(rsADO![���H��], 1) & " "
                strSQL = strSQL & ",'" & rsADO![������] & "�@" & rsADO![��] & "�@" & rsADO![���n�g�݌v���l] & "' "
                strSQL = strSQL & "," & varNullChk(rsADO![���͏�], 1) & " "
                strSQL = strSQL & ") "
                
                cnADO.Execute strSQL
            End If
            
        End If
        
        rsADO.MoveNext
    Loop
    
    bolfncUwawakuShitaji_GroupADD = True
    
    GoTo Exit_fncbolUwawakuShitaji_GroupADD
    
Err_fncbolUwawakuShitaji_GroupADD:
    MsgBox Err.Description
    Debug.Print strSQL
    
Exit_fncbolUwawakuShitaji_GroupADD:
    If rsADO.State = adStateOpen Then rsADO.Close
    Set rsADO = Nothing
    
    If cnADO.State = adStateOpen Then cnADO.Close
    Set cnADO = Nothing
    
End Function

Public Function bolfncUwawakuShitajiLength(ByVal in_varLength As Variant, ByRef out_dblLength() As Double) As Boolean
'   *************************************************************
'   ��g���n�������o
'   'ADD by K.Asayama 20170301
'   �߂�l:Boolean
'       ��True              ��g���n�������o
'       ��False             ��g���n�������o�s��
'
'    Input����
'       in_varLength        ��g���nW
'       out_dblLength()     �������ꂽW(False�̍ۂ͑S��0) --Output

'   *************************************************************

    Dim dblLength As Double
    Dim dblWork As Double
    
    On Error GoTo Err_bolfncUwawakuShitajiLength
    
    '�ϐ�������
    bolfncUwawakuShitajiLength = False
    
    ReDim out_dblLength(1 To 3)
    dblWork = 0
    
    '�����������ȊO�̏ꍇ��False
    If IsNumeric(in_varLength) Then
        '�����_��Q�ʈȉ��؂�̂�
        dblLength = RoundDown(CDbl(in_varLength), 1)
        dblLength = dblFIVEorZERO(dblLength)
    Else
        Exit Function
    End If
    
    '��g���n�擾
    '2420mm�ȉ��͂��̂܂�1�{
    '����ȏ�͏����ɂ�蕪������
    
    Select Case dblLength
        
        Case Is < 2420.5
            out_dblLength(1) = dblLength
            
        Case 2420.5 To 2720
            out_dblLength(1) = dblLength - 300
            out_dblLength(2) = 300
        
        Case 2720.5 To 4840
            out_dblLength(1) = dblLength - 2420
            out_dblLength(2) = 2420
        
        Case Is > 4840
            dblWork = dblFIVEorZERO(Roundx(dblLength / 3, 1))
            out_dblLength(1) = dblWork
            out_dblLength(2) = dblWork
            out_dblLength(3) = dblLength - (dblWork * 2)
            
    End Select
    
    bolfncUwawakuShitajiLength = True
    Exit Function

Err_bolfncUwawakuShitajiLength:
    MsgBox Err.Description, vbCritical, "��g���nW�����G���["
    
End Function

Public Function intFncUwawakuShitajiLengthGroup(in_dblLength As Double) As Integer
'   *************************************************************
'   ��g���n�����W�v�O���[�v���o
'   'ADD by K.Asayama 20170301
'   �߂�l:Integer
'       ������
'
'    Input����
'       in_dblLength        ��g���nW

'   *************************************************************
    intFncUwawakuShitajiLengthGroup = 0
    
    Select Case in_dblLength
        
        Case Is = 300
            intFncUwawakuShitajiLengthGroup = 300
            
        Case Is > 1820
            intFncUwawakuShitajiLengthGroup = 2430
                       
        Case Else
            intFncUwawakuShitajiLengthGroup = 1820
            
    End Select
    
End Function

Public Function fncstrUwawakuShitajiT(ByVal in_varHinban As Variant, ByVal in_varSagari As Variant, ByVal in_BoardT As Variant) As String
'   *************************************************************
'   ��g���n���ݒ��o
'   'ADD by K.Asayama 20170301
'   �߂�l:String
'       �����݁i���l����A+B�̕\�L������̂ŕ�����^���ŏo�́j
'
'    Input����
'       in_varHinban        ���n�ޕi��
'       in_varSagari        �������

'1.12.2
'   ���E�H�[���X���[���O�ǉ�
'3.0.0
'   ���{�[�h���l�� �����ǉ��iBRD1908�j
'   *************************************************************
    Dim varHinban As String
    Dim strSagari As String
    Dim strBoardT As String
    
    fncstrUwawakuShitajiT = ""
    
    On Error GoTo Err_fncstrUwawakuShitajiT
    
    If IsWallThru(in_varHinban) Then Exit Function
    
    If Not IsNull(in_varHinban) Then
        varHinban = Replace(in_varHinban, "�� ", "")
    Else
        Exit Function
    End If
    
    If in_varSagari = "�L" Then
        strSagari = "�L"
    Else
        strSagari = "��"
    End If
    
    Select Case Nz(in_BoardT, 0)
        Case 9.5
            strBoardT = "21"
        
        Case 12.5
            strBoardT = "18"
            
        Case 15
            strBoardT = "15"
            
        Case Else
            strBoardT = "30"
    End Select
    
    Select Case IsStealth_Seizo_TEMP(varHinban)
        '�X�e���X
        Case True

           Select Case strSagari
                '�������
                Case "�L"

                        If IsHirakido(varHinban) Or IsOyatobira(varHinban) Then
                            If varHinban Like "*G114*-####*" Then
                                fncstrUwawakuShitajiT = "30+9"
                            ElseIf IsSoftMotion(varHinban) Then
                                fncstrUwawakuShitajiT = "18+9"
                            Else
                                fncstrUwawakuShitajiT = "12+9"
                            End If
                        Else
                            If IsCloset_Hiraki(varHinban) Then
                                If varHinban Like "*G114*-####*" Then
                                    fncstrUwawakuShitajiT = "30"
                                Else
                                    fncstrUwawakuShitajiT = "12"
                                End If
                            Else
                                fncstrUwawakuShitajiT = strBoardT
                            End If
                        End If
                    
                '�V����܂�
                Case "��"
                    
                    If IsSoftMotion(varHinban) Then
                        fncstrUwawakuShitajiT = "18"
                        
                    ElseIf IsOyatobira(varHinban) Then
                        fncstrUwawakuShitajiT = "12"
                        
                    ElseIf IsCloset_Hiraki(varHinban) Then
                        
                        fncstrUwawakuShitajiT = "12"
                        
                    ElseIf Not IsHirakido(varHinban) And Not IsCloset_Slide(varHinban) And Not IsYukazukeRail(varHinban) Then
                        
                        fncstrUwawakuShitajiT = strBoardT
                        
                    End If

                    
            End Select
            
        '�C���Z�b�g���n
        Case False
            
            Select Case strSagari
            
                '�������
                Case "�L"
                    If IsHirakido(varHinban) Or IsOyatobira(varHinban) Then
                        If varHinban Like "*G114*-####*" Then
                            fncstrUwawakuShitajiT = "30"
                        ElseIf IsSoftMotion(varHinban) Then
                            fncstrUwawakuShitajiT = "18"
                        Else
                            fncstrUwawakuShitajiT = "12"
                        End If
                    Else
                        fncstrUwawakuShitajiT = strBoardT
                    End If
                    
                '�V����܂�
                Case "��"
                    If IsSoftMotion(varHinban) Then
                        fncstrUwawakuShitajiT = "18"
                    ElseIf IsOyatobira(varHinban) Then
                        fncstrUwawakuShitajiT = "12"
                    ElseIf Not IsHirakido(varHinban) Then
                        fncstrUwawakuShitajiT = strBoardT
                    End If
                    
                        
            End Select
    End Select
    
    Exit Function

Err_fncstrUwawakuShitajiT:
    
End Function

Public Function strfncFullHeightHinge(in_varHinban As Variant, in_varSpec As Variant) As String
'   *************************************************************
'   �t���n�C�g�q���W���[�J�[�m�F

'   ����:
'       ������i��
'         ��Spec

'
'   �߂�l:String
'       �����[�J�[��
'
'  2.7.0 ADD

'   *************************************************************

    Dim datNOW As Date
    
    strfncFullHeightHinge = ""
    
    '2018/9/6�ȍ~����K�p�Ƃ���**************
    datNOW = Date
    If datNOW < #9/6/2018# Then Exit Function
    '****************************************
    
    If IsNull(in_varHinban) Then Exit Function
    If IsNull(in_varSpec) Then Exit Function
    
    If (IsHirakido(CStr(in_varHinban)) Or IsOyatobira(CStr(in_varHinban)) Or IsKotobira(CStr(in_varHinban))) And Not IsHidden_Hinge(in_varHinban) Then
    
        If left(in_varSpec, 3) = "BRD" Then
        
            If right(in_varSpec, 4) >= "1808" Then
                strfncFullHeightHinge = "NISHIMURA"
            Else
                strfncFullHeightHinge = "YOGO"
            End If

        Else
        
            strfncFullHeightHinge = "YOGO"
            
        End If
    End If
    
End Function

Public Function bolFncCloset_IseharaToso(in_varHinban As Variant) As Boolean
'   *************************************************************
'   �ɐ����H��h���N���[�[�b�g�m�F
'
'   �߂�l:Boolean
'       ��True              �ɐ����H��h��
'       ��False             �ɐ����H��h���ȊO
'
'    Input����
'       in_varHinban        ����i�܌ˁj�i��

'   2.14.0 ADD
'   *************************************************************
    
    Dim strHinban As String
    
    bolFncCloset_IseharaToso = False
    
    If IsNull(in_varHinban) Then Exit Function
        
    If IsCloset_Oredo(in_varHinban) Or IsCloset_Hiraki(in_varHinban) Then
        If in_varHinban Like "*-####*(NI)*" Then
            bolFncCloset_IseharaToso = True
        End If
    End If
    
End Function

Public Function strFncFuchibariColor(in_varHinban As Variant, in_strColor As String, in_varSpec As Variant) As String

'   *************************************************************
'   ���\��F�m�F

'   �߂�l:Boolean
'       �����\��F
'
'    Input����
'       in_varHinban        ����i��
'       in_strColor         �F
'       in_varSpec          ��Spec

'   2.14.0 ADD

'   3.0.0
'   ���W�����A�̉��\��F�ύX(BA��MO)
'   *************************************************************
    
    Dim strHinban As String
    
    strFncFuchibariColor = ""
    
    On Error GoTo Err_strFncFuchibariColor
    
    If in_strColor = "" Then Exit Function
    
    If IsNull(in_varHinban) Then
        strFncFuchibariColor = in_strColor
        Exit Function
    End If
    
    strHinban = Replace(in_varHinban, "�� ", "")
    
    If IsCarloGiulia(strHinban) Then
        If in_strColor = "SB" Then
            strFncFuchibariColor = "EW"
        ElseIf in_strColor = "SH" Then
            If Is40mm(in_varSpec) Then
                strFncFuchibariColor = "MO"
            Else
                strFncFuchibariColor = "BA"
            End If
        End If
    Else
        strFncFuchibariColor = in_strColor
    End If
    
    Exit Function

Err_strFncFuchibariColor:
    MsgBox Err.Description
    strFncFuchibariColor = ""
    
End Function

Public Function varFncKanamonoSeizoBi(ByVal strKeiyakuNo As String, ByVal strTouNo As String, ByVal strHeyaNo As String, ByVal varDate As Variant, ByVal bytHinbanKubun As Byte) As Variant
'   *************************************************************
'   varFncKanamonoSeizoBi
'       �o�׋���������ꍇ�A�����w���f�[�^���琻�����i�ő�l�j��Ԃ�
'
'   �߂�l:Variant
'       ��Date�^            ������
'       ��Null              �������Ȃ��A���̓G���[
'
'    Input����
'       strKeiyakuNo        �_��ԍ�
'       strTouNo            ���ԍ�
'       strHeyaNo           �����ԍ�
'       varDate             ������
'       bytHinbanKubun      �i�ԋ敪 (1������j

'3.0.0 ADD
'   *************************************************************
    Dim strSQL As String
    
    varFncKanamonoSeizoBi = Null
    
    On Error GoTo Err_varFncKanamonoSeizoBi
    
    strSQL = ""
    strSQL = strSQL & "�_��ԍ� = '" & strKeiyakuNo & "' and  ���ԍ� = '" & strTouNo & "' and �����ԍ� = '" & strHeyaNo & "' "
    strSQL = strSQL & "and  �o�׋������� = True AND ������ < #" & Format(varDate, "yyyy/MM/dd") & "# "
    
    varFncKanamonoSeizoBi = DMax("������", "TEMP_���ޓW�J_������", strSQL)
    
    Exit Function
    
Err_varFncKanamonoSeizoBi:
    'Debug.Print Err.Description
    varFncKanamonoSeizoBi = Null
End Function

Public Function varFncKanamonoDaisha(ByVal strKeiyakuNo As String, ByVal strTouNo As String, ByVal strHeyaNo As String, ByVal bytHinbanKubun As Byte) As Variant
'   *************************************************************
'   varFncKanamonoDaisha
'       �o�׋�������ԃf�[�^�ɓo�^����Ă���ꍇ�A�����w���f�[�^�����ԃR�[�h��Ԃ�

'
'   �߂�l:Variant
'       ��String�^          ��ԃR�[�h
'       ��Null              �������Ȃ��A���̓G���[
'
'    Input����
'       strKeiyakuNo        �_��ԍ�
'       strTouNo            ���ԍ�
'       strHeyaNo           �����ԍ�
'       bytHinbanKubun      �i�ԋ敪 (1������j

'3.0.0 ADD
'   *************************************************************
    Dim strSQL As String
    Dim strKeiyakuBango As String
    Dim bytKubun As Byte
    
    varFncKanamonoDaisha = Null
    
    On Error GoTo Err_varFncKanamonoDaisha
    
    strKeiyakuBango = strKeiyakuNo & "-" & strTouNo & "-" & strHeyaNo
    
    Select Case bytHinbanKubun
        Case 1
            bytKubun = 10
    End Select
    
    If Not IsNull(bytKubun) Then
            strSQL = ""
            strSQL = strSQL & "�_��ԍ� = '" & strKeiyakuNo & "' and  ���ԍ� = '" & strTouNo & "' and �����ԍ� = '" & strHeyaNo & "' "
    
            varFncKanamonoDaisha = DMax("��ԃR�[�h", "TEMP_���ޓW�J_������", strSQL)
    End If
    
    Exit Function
    
Err_varFncKanamonoDaisha:
    'Debug.Print Err.Description
    varFncKanamonoDaisha = Null
End Function