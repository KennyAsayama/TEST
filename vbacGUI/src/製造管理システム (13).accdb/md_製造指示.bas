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
                        strSQL = strSQL & "," & varNullChk(fncstrUwawakuShitajiT(varHinban, rsADO![�������]), 1) & " "
                        
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
                Err.Raise "�ȍ~�̒��[�̏o�͂𒆎~���܂�"
            End If
        Else '�}�ʂ���
            If fncstrUwawakuShitajiT(varHinban, rsADO![�������]) <> "" Then
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

Public Function fncstrUwawakuShitajiT(ByVal in_varHinban As Variant, ByVal in_varSagari As Variant) As String
'   *************************************************************
'   ��g���n���ݒ��o
'   'ADD by K.Asayama 20170301
'   �߂�l:String
'       �����݁i���l����A+B�̕\�L������̂ŕ�����^���ŏo�́j
'
'    Input����
'       in_varHinban        ���n�ޕi��
'       in_varSagari        �������

'   *************************************************************
    Dim varHinban As String
    Dim strSagari As String
    
    fncstrUwawakuShitajiT = ""
    
    On Error GoTo Err_fncstrUwawakuShitajiT
    
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
                                fncstrUwawakuShitajiT = "30"
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
                        
                        fncstrUwawakuShitajiT = "30"
                        
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
                        fncstrUwawakuShitajiT = "30"
                    End If
                    
                '�V����܂�
                Case "��"
                    If IsSoftMotion(varHinban) Then
                        fncstrUwawakuShitajiT = "18"
                    ElseIf IsOyatobira(varHinban) Then
                        fncstrUwawakuShitajiT = "12"
                    ElseIf Not IsHirakido(varHinban) Then
                        fncstrUwawakuShitajiT = "30"
                    End If
                    
                        
            End Select
    End Select
    
    Exit Function

Err_fncstrUwawakuShitajiT:
    
End Function