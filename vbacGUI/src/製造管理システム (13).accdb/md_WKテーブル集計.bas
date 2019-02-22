Option Compare Database
Option Explicit

Public Function SetOrderData(ByVal inDate As Date, ByVal inDateKbn As Byte, inSeizoKbn As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'�����f�[�^��WK�t�@�C���ɓ]������
'
'   :����
'       inDate          :�[�i��
'       inDateKbn       :1:�[�i���x�[�X�W�v�A2:�������x�[�X�W�v
'       inSeizoKbn      :����A�g�A���n
'
'   :�߂�l
'       True            :����
'       False           :���s
'1.10.7 K.Asayama ADD 20160108
'       ���uF_�@��_���ʁv�H���\�{�^�����g�p�\�ɂ��������ǉ�
'       ���uWK_�D�f�[�^�v�ɏo�ו��@�A�F�i�h���̂݁j��ǉ�
'       �� �������x�[�X�̎��͖��m����W�v
'       �� inDate��[9999/12/31]�̎��͓��tNull�̃f�[�^���o�́i�������x�[�X�j
'       �� inDate��[9999/12/30]�̎��͓��t�͊֌W�Ȃ����m��̃f�[�^���o�́i�������x�[�X�j
'1.10.8 K.Asayama ADD 20160114
'       �����F���`�J����
'1.10.10 K.Asayama Change 20160212
'       ������������Ⴂ�Б��~���[�I�v�V�����ǉ�
'1.10.14 K.Asayama Change 20160418
'       ���o�O�C���v�Z�o�ד���Null�Ŗ߂����ꍇ�̑Ή�
'1.10.16 K.Asayama Change
'       �����n�A�X�e���X����
'2.5.0
'       ���o�ד��v�Z�����[�h�^�C���ɕύX
'2.5.2
'       ��F�y�̓h���W�v�Ή�
'2.8.0
'       ���A���W���I�t�B�X�[�����f�[�^��荞��
'2.9.0
'       �������[�g�Ƃ̐ڑ����ԒZ�k���̂��߃��[�N�e�[�u���쐬
'       �����[�h�^�C������o�ד��v�Z���T�[�o�T�C�h���i�p�t�H�[�}���X���P�j
'       ��DoEvents�ǉ�
'2.13.0
'       ��Vertica�V���N�������Ή�
'--------------------------------------------------------------------------------------------------------------------
    Dim objREMOTEdb As New cls_BRAND_MASTER
    Dim objLOCALdb As New cls_LOCALDB
    Dim objLOCALDB_2 As New cls_LOCALDB
    
    Dim strSQL As String
    Dim strSQLWK As String
    
    Dim bolTRAN As Boolean
    Dim strKeiyakuNo As String
    Dim varCalcShukkaBi As Variant
    Dim intMinusDays As Integer
    Dim dblWindowTop As Double, dblWindowLeft As Double, dblWindowHight As Double, dblWindowWidth As Double
    Dim bolFormOpen As Boolean
    Dim strKubun As String
    Dim strLT As String
    Dim i As Integer
    
    bolFormOpen = False
    
    On Error GoTo Err_SetOrderData
    
'    Me.Painting = False
    Application.Echo False
    
    SetOrderData = False
    bolTRAN = False
    strKeiyakuNo = ""
    
    Select Case inSeizoKbn
        Case "����"
            strKubun = "1,2,3"
            strLT = "����LT"
        Case "�g"
            strKubun = "4,5"
            strLT = "�gLT"
        '1.10.16 Change
'        Case "���n"
'            strKubun = "6,7"
        Case "���n"
            strKubun = "6"
            strLT = "���n��LT"
        Case "�X�e���X"
            strKubun = "7"
            strLT = "���n��LT"
        Case Else
            Err.Raise 9999, , "�����敪�]���G���["
    End Select
    
    strSQL = ""
    strSQL = strSQL & "select * "
    strSQL = strSQL & ",case when �o�ד� is null then dbo.fnc�o�ד��擾_LT�̂�(�m���," & strLT & ") "
    strSQL = strSQL & " else null end �v�Z�o�ד� "
    strSQL = strSQL & "from ( "
    strSQL = strSQL & "select s.�_��ԍ�,s.���ԍ�,s.�����ԍ� "
    strSQL = strSQL & ",s.�_��ԍ� + '-' + s.���ԍ� + '-' + s.�����ԍ� AS �_��No "
    strSQL = strSQL & ",s.�m��� "
    strSQL = strSQL & ",dbo.fncSeizosyukkaDate(s.�_��ԍ�,s.���ԍ�,s.�����ԍ�,s.��,"
    strSQL = strSQL & "case s.�����敪 "
    strSQL = strSQL & "when 1 then 1 "
    strSQL = strSQL & "when 2 then 1 "
    strSQL = strSQL & "when 3 then 1 "
    strSQL = strSQL & "when 4 then 2 "
    strSQL = strSQL & "when 5 then 5 "
    strSQL = strSQL & "when 6 then 3 "
    strSQL = strSQL & "when 7 then 3 "
    strSQL = strSQL & "else 999 "
    strSQL = strSQL & "end) as �o�ד� "
    strSQL = strSQL & ",dbo.fncNohinAddress(s.�_��ԍ�,s.���ԍ�,s.�����ԍ�,s.��,"
    strSQL = strSQL & "case s.�����敪 "
    strSQL = strSQL & "when 1 then 1 "
    strSQL = strSQL & "when 2 then 1 "
    strSQL = strSQL & "when 3 then 1 "
    strSQL = strSQL & "when 4 then 2 "
    strSQL = strSQL & "when 5 then 5 "
    strSQL = strSQL & "when 6 then 3 "
    strSQL = strSQL & "when 7 then 3 "
    strSQL = strSQL & "else 999 "
    strSQL = strSQL & "end) as �[�i�Z�� "
    strSQL = strSQL & ",s.������ "
    strSQL = strSQL & ",s.�� "
    strSQL = strSQL & ",s.�����敪 "
    strSQL = strSQL & ",s.���� "
    strSQL = strSQL & ",m.������ "
    strSQL = strSQL & ",m.�{�H�X "
    strSQL = strSQL & ",case s.�����敪 when 1  then s.���� else 0 end AS [Flush��] "
    strSQL = strSQL & ",case s.�����敪 when 2  then s.���� else 0 end AS [F�y��] "
    strSQL = strSQL & ",case s.�����敪 when 3  then s.���� else 0 end AS [�y��] "
    strSQL = strSQL & ",case s.�����敪 when 4  then s.���� else 0 end AS [�g��] "
    strSQL = strSQL & ",case s.�����敪 when 5  then s.���� else 0 end AS [�O���g��] "
    strSQL = strSQL & ",case s.�����敪 when 6  then s.���� else 0 end AS [���n�g��] "
    strSQL = strSQL & ",case s.�����敪 when 7  then s.���� else 0 end AS [�X�e���X�g��] "
    strSQL = strSQL & ",s.�o�^���i�� "
    strSQL = strSQL & ",s.���� "
    strSQL = strSQL & ",s.�m�� "
    strSQL = strSQL & ",y.�R�����g as ���l "
    '1.10.7 ADD
    strSQL = strSQL & ",dbo.fncNohinHaiso(s.�_��ԍ�,s.���ԍ�,s.�����ԍ�,s.��,"
    strSQL = strSQL & "case s.�����敪 "
    strSQL = strSQL & "when 1 then 1 "
    strSQL = strSQL & "when 2 then 1 "
    strSQL = strSQL & "when 3 then 1 "
    strSQL = strSQL & "when 4 then 2 "
    strSQL = strSQL & "when 5 then 5 "
    strSQL = strSQL & "when 6 then 3 "
    strSQL = strSQL & "when 7 then 3 "
    strSQL = strSQL & "else 999 "
    strSQL = strSQL & "end) as �o�ו��@ "
    '1.10.7 ADD End
    strSQL = strSQL & ",����LT,�gLT,���n��LT,WTLT,����LT,���֎��[LT,�����LT "
    
    If inSeizoKbn = "����" Then
        strSQL = strSQL & ",�K���X���ד�,���[�o�[���ד�,���̑����ד�,�o�׋������ד� "
    Else
        strSQL = strSQL & ",Null as �K���X���ד�,Null as ���[�o�[���ד�,Null as ���̑����ד�,Null as �o�׋������ד� "
    End If
    
    strSQL = strSQL & "from T_�����w�� s "
    strSQL = strSQL & "inner join T_��Ͻ� m "
    strSQL = strSQL & "on s.�_��ԍ� = m.�_��ԍ� and s.���ԍ� = m.���ԍ� and s.�����ԍ� = m.�����ԍ� "
    strSQL = strSQL & "inner join T_��Ͻ�_2 m2 "
    strSQL = strSQL & "on s.�_��ԍ� = m2.�_��ԍ� and s.���ԍ� = m2.���ԍ� and s.�����ԍ� = m2.�����ԍ� "
    strSQL = strSQL & "left join T_�����\�� y "
    strSQL = strSQL & "on s.�_��ԍ� = y.�_��ԍ� and s.���ԍ� = y.���ԍ� and s.�����ԍ� = y.�����ԍ� and s.�����敪 = y.�����敪 "
    
    If inSeizoKbn = "����" Then
        strSQL = strSQL & "left join (select �_��ԍ�,���ԍ�,�����ԍ�,�� "
        strSQL = strSQL & ",max(case when ���ގ��CD like '%��׽%' or ���ގ��CD like '%�װ%'  then ���ד� end) �K���X���ד� "
        strSQL = strSQL & ",max(case when ���ގ��CD like '%ٰ�ް�Ư�%' then ���ד� end) ���[�o�[���ד� "
        strSQL = strSQL & ",max(case when ���ގ��CD not like '%��׽%' and ���ގ��CD not like '%�װ%' and ���ގ��CD not like '%ٰ�ް�Ư�%' and (�����i is null or �����i <> '��') then ���ד� end) ���̑����ד� "
        strSQL = strSQL & ",max(case when �����i = '��' then ���ד� end) �o�׋������ד� "
        strSQL = strSQL & "from T_AO���ޔ[����� AO "
        strSQL = strSQL & "where �i�ԋ敪 = 1 and �����敪 = 1 "
        strSQL = strSQL & "group by �_��ԍ�,���ԍ�,�����ԍ�,�� "
        strSQL = strSQL & ") AO "
        strSQL = strSQL & "on s.�_��ԍ� = AO.�_��ԍ� and s.���ԍ� = AO.���ԍ� and s.�����ԍ� = AO.�����ԍ� and s.�� = AO.�� "
    End If
    
    If inDateKbn = 1 Then
        strSQL = strSQL & "where s.�m��� = '" & Format(inDate, "yyyy/mm/dd") & "' "
    Else
        '1.10.7 ADD
        If inDate = #12/31/9999# Then
            strSQL = strSQL & "where s.������ is Null "
            
        ElseIf inDate = #12/30/9999# Then
            strSQL = strSQL & " where �m�� < 2 "
        Else
        '1.10.7 ADD End
            strSQL = strSQL & "where s.������ = '" & Format(inDate, "yyyy/mm/dd") & "' "
            '1.10.7 DEL
            'strSQL = strSQL & " and �m�� > 0 "
            '1.10.7 DEL END
        
        '1.10.7 ADD
        End If
        '1.10.7 ADD End
    End If
    strSQL = strSQL & " and s.�����敪 in ( " & strKubun & ") "
    
    '�E�H�[���X���[�͐����������Ă��Ȃ��̂őΏۊO
    If inSeizoKbn = "���n" Then
        strSQL = strSQL & " and s.�o�^���i�� not like 'WS%' "
    End If
    
    strSQL = strSQL & " ) WKTABLE "
    
    If Not objLOCALdb.ExecSQL("delete from WK_�D�f�[�^") Then
        Err.Raise 9999, , "�����w���f�[�^���[�N�i���[�J���j�������G���["
    End If
    
    '�ŏ��͂Ȃ��̂ŃG���[�͖���
    objLOCALdb.ExecSQL ("drop table TMP_�����w���f�[�^")
    
    strSQLWK = ""
    strSQLWK = strSQLWK & "CREATE TABLE TMP_�����w���f�[�^( "
    strSQLWK = strSQLWK & " �_��ԍ�            TEXT(10) "
    strSQLWK = strSQLWK & ",���ԍ�              TEXT(10) "
    strSQLWK = strSQLWK & ",�����ԍ�            TEXT(10) "
    strSQLWK = strSQLWK & ",�_��No              TEXT(30) "
    strSQLWK = strSQLWK & ",�m���              DATE "
    strSQLWK = strSQLWK & ",�o�ד�              DATE "
    strSQLWK = strSQLWK & ",�[�i�Z��            TEXT(255) "
    strSQLWK = strSQLWK & ",������              DATE "
    strSQLWK = strSQLWK & ",��                  INT "
    strSQLWK = strSQLWK & ",�����敪            INT "
    strSQLWK = strSQLWK & ",����                INT "
    strSQLWK = strSQLWK & ",������              TEXT(255) "
    strSQLWK = strSQLWK & ",�{�H�X              TEXT(255) "
    strSQLWK = strSQLWK & ",Flush��             INT "
    strSQLWK = strSQLWK & ",F�y��               INT "
    strSQLWK = strSQLWK & ",�y��                INT "
    strSQLWK = strSQLWK & ",�g��                INT "
    strSQLWK = strSQLWK & ",�O���g��            INT "
    strSQLWK = strSQLWK & ",���n�g��            INT "
    strSQLWK = strSQLWK & ",�X�e���X�g��        INT "
    strSQLWK = strSQLWK & ",�o�^���i��          TEXT(50) "
    strSQLWK = strSQLWK & ",����                INT "
    strSQLWK = strSQLWK & ",�m��                INT "
    strSQLWK = strSQLWK & ",���l                TEXT(255) "
    strSQLWK = strSQLWK & ",�o�ו��@            TEXT(50) "
    strSQLWK = strSQLWK & ",����LT              INT "
    strSQLWK = strSQLWK & ",�gLT                INT "
    strSQLWK = strSQLWK & ",���n��LT            INT "
    strSQLWK = strSQLWK & ",WTLT                INT "
    strSQLWK = strSQLWK & ",����LT              INT "
    strSQLWK = strSQLWK & ",���֎��[LT          INT "
    strSQLWK = strSQLWK & ",�����LT            INT "
    strSQLWK = strSQLWK & ",�K���X���ד�        DATE "
    strSQLWK = strSQLWK & ",���[�o�[���ד�      DATE "
    strSQLWK = strSQLWK & ",���̑����ד�        DATE "
    strSQLWK = strSQLWK & ",�o�׋������ד�      DATE "
    strSQLWK = strSQLWK & ",�v�Z�o�ד�          DATE "
    strSQLWK = strSQLWK & ") "
        
    If Not objLOCALdb.ExecSQL(strSQLWK) Then
        Err.Raise 9999, , "�����w���f�[�^���[�N�i���[�J���j�쐬�G���["
    End If
    
    
    With objREMOTEdb
        If .ExecSelect(strSQL) Then
            If Not bolfncTableCopyToLocal(.GetRS, "TMP_�����w���f�[�^", False) Then
                Err.Raise 9999, , "TMP_�����w���f�[�^���[�J���R�s�[�G���[�B�Ǘ��҂ɘA�����Ă�������"
            End If
        Else
            Err.Raise 9999, , ""
        End If
    End With
    
    strSQL = ""
    strSQL = strSQL & "select * from TMP_�����w���f�[�^ "
    
    i = 0
    
    With objLOCALDB_2
        If .ExecSelect(strSQL) Then
            If objLOCALdb.ExecSelect_Writable("select * from WK_�D�f�[�^") Then
            
                objLOCALdb.BeginTrans
                bolTRAN = True
                
                Do While Not .GetRS.EOF
                        objLOCALdb.GetRS.AddNew

                        objLOCALdb.GetRS![�_��ԍ�] = .GetRS![�_��ԍ�]
                        objLOCALdb.GetRS![���ԍ�] = .GetRS![���ԍ�]
                        objLOCALdb.GetRS![�����ԍ�] = .GetRS![�����ԍ�]
                        objLOCALdb.GetRS![������] = .GetRS![������]
                        objLOCALdb.GetRS![�{�H�X] = .GetRS![�{�H�X]
                        objLOCALdb.GetRS![�_��No] = .GetRS![�_��No]
                        objLOCALdb.GetRS![��] = .GetRS![��]
                        objLOCALdb.GetRS![�����敪] = .GetRS![�����敪]
                        objLOCALdb.GetRS![�m���] = .GetRS![�m���]
'                        If IsNull(.GetRS![�o�ד�]) Then
'                            objLOCALDB.GetRS![�o�ד��o�^] = False
'                            'If fncbolSyukkaBiFromAddress(.GetRS![�[�i�Z��], .GetRS![�m���], varCalcShukkaBi, intMinusDays) Then
'                            If fncbolSyukkaBiFromLeadTime(.GetRS(strLT), .GetRS![�m���], varCalcShukkaBi, intMinusDays) Then
'                                '1.10.14
'                                If Not IsNull(varCalcShukkaBi) Then
'                                    objLOCALDB.GetRS![�o�ד�] = CDate(varCalcShukkaBi)
'                                End If
'                            Else
'                                objLOCALDB.GetRS![�o�ד�] = .GetRS![�o�ד�]
'                            End If
'                        Else
'                            objLOCALDB.GetRS![�o�ד��o�^] = True
'                            objLOCALDB.GetRS![�o�ד�] = .GetRS![�o�ד�]
'                        End If
                        
                        If IsNull(.GetRS![�o�ד�]) Then
                            objLOCALdb.GetRS![�o�ד��o�^] = False
                            If Not IsNull(.GetRS![�v�Z�o�ד�]) Then
                                objLOCALdb.GetRS![�o�ד�] = CDate(.GetRS![�v�Z�o�ד�])
                            Else
                                objLOCALdb.GetRS![�o�ד�] = Null
                            End If
                        Else
                            objLOCALdb.GetRS![�o�ד��o�^] = True
                           objLOCALdb.GetRS![�o�ד�] = .GetRS![�o�ד�]
                        End If

                        objLOCALdb.GetRS![������] = .GetRS![������]
                        objLOCALdb.GetRS![�[�i�Z��] = .GetRS![�[�i�Z��]
                        
                        'If IsNull(.GetRS![�m��]) Or .GetRS![�m��] = 0 Then
                        '    objLocalDB.GetRS![�m��] = 0
                        'Else
                        '    objLocalDB.GetRS![�m��] = -1
                        'End If
                        
                        objLOCALdb.GetRS![�m��] = .GetRS![�m��]
                        
                        '1.10.7 ADD
                        objLOCALdb.GetRS![�o�ו��@] = .GetRS![�o�ו��@]
                        '1.10.7 ADD End
                        
                        objLOCALdb.GetRS![Flush��] = .GetRS![Flush��] + .GetRS![F�y��]
                        objLOCALdb.GetRS![F�y��] = .GetRS![F�y��]
                        objLOCALdb.GetRS![�y��] = .GetRS![�y��]
                        objLOCALdb.GetRS![�g��] = .GetRS![�g��]
                        objLOCALdb.GetRS![�O���g��] = .GetRS![�O���g��]
                                                
                        If IsStealth_Seizo_TEMP(Nz(.GetRS![�o�^���i��], "nz")) Then
                            objLOCALdb.GetRS![���n�g��] = 0
                            '1.10.16 change
                            'objLOCALDB.GetRS![�X�e���X�g��] = .GetRS![���n�g��]
                            If .GetRS![�����敪] = 7 Then
                                objLOCALdb.GetRS![�X�e���X�g��] = .GetRS![�X�e���X�g��]
                            Else
                                objLOCALdb.GetRS![�X�e���X�g��] = .GetRS![���n�g��]
                            End If
                        Else
                            objLOCALdb.GetRS![�X�e���X�g��] = 0
                            objLOCALdb.GetRS![���n�g��] = .GetRS![���n�g��]
                        End If
                        
                        If .GetRS![�����敪] >= 1 And .GetRS![�����敪] <= 3 Then
                            If IsThruGlass(.GetRS![�o�^���i��]) Then
                                If IsVertica(.GetRS![�o�^���i��]) Then
                                    objLOCALdb.GetRS![�X���[�K���X��] = IsVertica_Maisu(.GetRS![�o�^���i��], .GetRS![Flush��])
                                Else
                                    '1.10.10 K.Asayama Change
                                    'objLOCALDB.GetRS![�X���[�K���X��] = .GetRS![Flush��]
                                    objLOCALdb.GetRS![�X���[�K���X��] = fncIntHalfGlassMirror_Maisu(.GetRS![�o�^���i��], .GetRS![Flush��])
                                    '1.10.10 K.Asayama Change End
                                End If
                            Else
                                objLOCALdb.GetRS![�X���[�K���X��] = 0
                            End If
                            
                            If IsAir(.GetRS![�o�^���i��]) Then
                                objLOCALdb.GetRS![���[�o�[����] = .GetRS![Flush��]
                            Else
                                objLOCALdb.GetRS![���[�o�[����] = 0
                            End If
                            
                            If IsPainted(.GetRS![�o�^���i��]) Then
                                If .GetRS![F�y��] > 0 Then
                                    objLOCALdb.GetRS![�h������] = .GetRS![F�y��]
                                Else
                                    objLOCALdb.GetRS![�h������] = .GetRS![Flush��]
                                End If
                                '1.10.7 ADD
                                objLOCALdb.GetRS![�F] = fncvalDoorColor(.GetRS![�o�^���i��])
                                '1.10.7 ADD End
                            Else
                                objLOCALdb.GetRS![�h������] = 0
                            End If
                            
                            If IsMonster(.GetRS![�o�^���i��]) Then
                                objLOCALdb.GetRS![�����X�^�[��] = .GetRS![F�y��]
                            Else
                                objLOCALdb.GetRS![�����X�^�[��] = 0
                            End If
                            '1.10.8 ADD
                            If IsVertica(.GetRS![�o�^���i��]) Then
                                'objLOCALdb.GetRS![���F���`�J��] = .GetRS![Flush��]
                                objLOCALdb.GetRS![���F���`�J��] = IsVertica_Maisu(.GetRS![�o�^���i��], .GetRS![Flush��])
                            Else
                                objLOCALdb.GetRS![���F���`�J��] = 0
                            End If
                            '1.10.8 ADD End
                        Else
                            objLOCALdb.GetRS![�X���[�K���X��] = 0
                            objLOCALdb.GetRS![���[�o�[����] = 0
                            objLOCALdb.GetRS![�h������] = 0
                            objLOCALdb.GetRS![�����X�^�[��] = 0
                            '1.10.8 ADD
                            objLOCALdb.GetRS![���F���`�J��] = 0
                            '1.10.8 ADD End
                        End If
                        
                        objLOCALdb.GetRS![���l] = .GetRS![���l]
                    
                    objLOCALdb.GetRS![�K���X���ד�] = .GetRS![�K���X���ד�]
                    objLOCALdb.GetRS![���[�o�[���ד�] = .GetRS![���[�o�[���ד�]
                    objLOCALdb.GetRS![���̑����ד�] = .GetRS![���̑����ד�]
                    objLOCALdb.GetRS![�o�׋������ד�] = .GetRS![�o�׋������ד�]
                    
                    objLOCALdb.GetRS.Update
                                        
                    i = i + 1
                    
                    If i Mod 100 = 0 Then
                        DoEvents
                    End If
                    
                    .GetRS.MoveNext
                Loop
                
                If bolTRAN Then objLOCALdb.Commit
                bolTRAN = False
            Else
                Err.Raise 9999, , "�`�F�b�N���X�g���[�N�i���[�J���j�I�[�v���G���["
            
            End If
        Else
            Err.Raise 9999, , "�`�F�b�N���X�g���o�G���["
        End If
    End With
    
    '1.10.7 ADD ���l�f�[�^�Ăяo��
    If Not SetBikouData() Then
        Err.Raise 9999, , "���l���Ăяo���ُ�"
    End If
    
    DoCmd.SetWarnings False
    
    
    If Form_IsLoaded("F_�@��_����") Then
        bolFormOpen = True
    End If
    
    
    If Not bolFormOpen Then
        DoCmd.OpenForm "F_�@��_����", acNormal, , , , , inDateKbn
    Else
        '1.10.7 Change
        'If Not Form_F_�@��_����.bolfncData_Update(inSeizoKbn) Then
        If Not Form_F_�@��_����.bolfncData_Update(inSeizoKbn, inDateKbn) Then
        '1.10.7 Change End
            DoCmd.Close acForm, "F_�@��_����", acSaveNo
        End If
    End If
    
    If bolFormOpen = True Then
        DoCmd.SelectObject acForm, "F_�@��_����"
    End If
    
    DoCmd.SetWarnings True
    
    SetOrderData = True
    GoTo Exit_SetOrderData
    
Err_SetOrderData:
    If bolTRAN Then objLOCALdb.Rollback
    bolTRAN = False
    MsgBox Err.Description

Exit_SetOrderData:
    Set objREMOTEdb = Nothing
    Set objLOCALdb = Nothing
    Set objLOCALDB_2 = Nothing
    
    Application.Echo True
    'Me.Painting = True
    
End Function

Public Function SetOrderCount(ByVal inDateKbn As Byte, ByRef Captionctl() As cls_Labelset, ByRef Graphctl() As cls_Labelset, ByRef Graphctl_Kakutei() As cls_Labelset, ByRef Graphctl_Temp() As cls_Labelset, ByRef Itemctl() As cls_Labelset, ByVal in_HinbanKubun As Integer, ByVal in_KojoCD As Integer)
'--------------------------------------------------------------------------------------------------------------------
'���ʏW�v����
'
'   :����
'       inDateKbn       :1:�[�i���x�[�X�W�v�A2:�������x�[�X�W�v
'       Captionctl      :���t�\�����x���i�R���g���[���z��j
'       Graphctl        :���ʕ\�����x���i�R���g���[���z��j
'       Graphctl_Kakutei:�m�萔�ʕ\�����x���i�R���g���[���z��j
'       Graphctl_Temp   :���m�萔�ʕ\�����x���i�R���g���[���z��j
'       Itemctl         :���i�\�����x���i2�����R���g���[���z��[���t,���i]�j
'       in_HinbanKubun  :1,����A2,�g�A3,���n
'       in_KojoCD       :�H��CD

'
'   :�߂�l
'       True            :����
'       False           :���s
'---------------------------
'   �ύX
'       1.10.1 K.Asayama ���n���A�X�e���X�������x���\���i�e�X�K���X���A�����X�^�[���𗬗p�j
'       1.10.7 K.Asayama �����ɉ��m��iGraphctl_Temp�j�ǉ��A�m�萔�W�v�ǉ� �b��ŕ\���͂��Ȃ��B�m�萔�̃��x���ɃJ�b�R�Ƃ��Ő��ʕ\��
'       1.10.8 K.Asayama
'                       ���y�p���x�������F���`�J�p���x���ɕύX
'                       ���O���t���x���̐��ʁiCaption�j��ControlTipText�Ή�
'       1.10.10 K.Asayama Change 20160212
'                       ������������Ⴂ�Б��~���[�I�v�V�����ǉ�
'       2.0.0
'                       ���H��CD���g�p���Ȃ�

'       2.13.0
'                       ��Vertica�V���N�������Ή�
'--------------------------------------------------------------------------------------------------------------------
    Dim objREMOTEdb As New cls_BRAND_MASTER
    Dim strSQL_C As String
    Dim strSQL As String
    Dim strKubun As String
    Dim i As Integer
    Dim bolToku As Boolean
    
    Dim intFlushM As Integer
    Dim intFkamachiM As Integer
    Dim intKamachiM As Integer
    Dim intThruM As Integer
    Dim intPaintM As Integer
    Dim intAirM As Integer
    Dim intMonsterM As Integer
    Dim intKakuteiM As Integer
    Dim intShitajiM As Integer
    Dim intStealthM As Integer
    '1.10.7 ADD
    Dim intKakuteiTempM As Integer
    '1.10.7 ADD End
    '1.10.8 ADD
    Dim intVerticaM As Integer
    '1.10.8 ADD End
    
    On Error GoTo Err_SetOrderCount
    
    '���n�Ƌ��p���郉�x������U����ŏ�����
    For i = 0 To UBound(Itemctl)
        Itemctl(i, 1).CaptionSet ("�K���X")
        Itemctl(i, 4).CaptionSet ("Monster")
        
        If inDateKbn = 1 Then
            Itemctl(i, 1).SetWidth (197)
            Itemctl(i, 4).SetWidth (197)
        Else
            Itemctl(i, 1).SetWidth (107)
            Itemctl(i, 4).SetWidth (107)
        End If
    Next
    
    Select Case in_HinbanKubun
        Case 1 'Flush
            strKubun = "1,2,3"
        Case 2 'Waku
            strKubun = "4,5"
        '1.10.16 Change
'        Case 3 'Shitaji
'            strKubun = "6,7"
        Case 3 'Shitaji
            strKubun = "6"
        Case 4 'Stealth
            strKubun = "7"
        Case Else
            strKubun = CStr(in_HinbanKubun)
    End Select
    
    strSQL_C = "select s.�o�^���i��, s.�����敪, s.�m��, s.���� as ���� from T_�����w�� s "
    strSQL_C = strSQL_C & "inner join T_�󒍖��� m "
    strSQL_C = strSQL_C & "on m.�_��ԍ� = s.�_��ԍ� and m.���ԍ� = s.���ԍ� and m.�����ԍ� = s.�����ԍ� and m.�� = s.�� "
    
    
    For i = 0 To UBound(Captionctl)
        strSQL = strSQL_C
        If inDateKbn = 1 Then
            strSQL = strSQL & " where s.�m��� = '" & Captionctl(i).GetTag & "'"
        Else
            strSQL = strSQL & " where s.������ = '" & Captionctl(i).GetTag & "'"
            '1.10.7 K.Asayama Change
            'strSQL = strSQL & " and �m�� > 0 "
            '1.10.7 K.Asayama Change End
        End If
        strSQL = strSQL & " and s.�����敪 in ( " & strKubun & ")"
'        Select Case in_HinbanKubun
'            Case 1
'                strSQL = strSQL & " and �����敪 = 1"
'            Case 6
'                strSQL = strSQL & " and �����敪 between 6 and 7"
'            Case Else
'                strSQL = strSQL & " and �����敪= " & in_HinbanKubun
'        End Select
'
        '�E�H�[���X���[�͐����������Ă��Ȃ��̂őΏۊO
        If in_HinbanKubun = 3 Then
            strSQL = strSQL & " and s.�o�^���i�� not like 'WS%' "
        End If
    
'        strSQL = strSQL & " and s.�H��CD = " & in_KojoCD
        strSQL = strSQL & " "
        
        With objREMOTEdb
           If .ExecSelect(strSQL) Then
                intFlushM = 0
                intFkamachiM = 0
                intKamachiM = 0
                intThruM = 0
                intPaintM = 0
                intAirM = 0
                intMonsterM = 0
                intKakuteiM = 0
                intShitajiM = 0
                intStealthM = 0
                '1.10.7 ADD
                intKakuteiTempM = 0
                '1.10.7 ADD End
                '1.10.8 ADD
                intVerticaM = 0
                '1.10.8 ADD End
    
                Do Until .GetRS.EOF
                    
                    intFlushM = intFlushM + .GetRS("����")
                    
                    '1.10.7 ADD �������x�[�X�̎��͖��m��͏W�v���Ȃ�
                    If (inDateKbn = 1) Or (inDateKbn = 2 And .GetRS("�m��") <> 0) Then
                    '1.10.7 ADD End
                        Select Case .GetRS("�����敪")
                            Case 1, 2, 3
                                If .GetRS("�����敪") = 2 Then intFkamachiM = intFkamachiM + .GetRS("����")
                                If .GetRS("�����敪") = 3 Then intKamachiM = intKamachiM + .GetRS("����")
                                '1.10.10 K.Asayama Change
                                'If IsThruGlass(.GetRS("�o�^���i��")) Then intThruM = intThruM + .GetRS("����")
                                If IsThruGlass(.GetRS("�o�^���i��")) Then
                                    If IsVertica(.GetRS("�o�^���i��")) Then
                                        intThruM = intThruM + IsVertica_Maisu(.GetRS("�o�^���i��"), .GetRS("����"))
                                    Else
                                        intThruM = intThruM + fncIntHalfGlassMirror_Maisu(.GetRS("�o�^���i��"), .GetRS("����"))
                                    End If
                                End If
                                '1.10.10 K.Asayama Change End
                                If IsPainted(.GetRS("�o�^���i��")) Then intPaintM = intPaintM + .GetRS("����")
                                If IsAir(.GetRS("�o�^���i��")) Then intAirM = intAirM + .GetRS("����")
                                If IsMonster(.GetRS("�o�^���i��")) Then intMonsterM = intMonsterM + .GetRS("����")
                                '1.10.8 K.Asayama ADD
                                'If IsVertica(.GetRS("�o�^���i��")) Then intVerticaM = intVerticaM + .GetRS("����")
                                If IsVertica(.GetRS("�o�^���i��")) Then intVerticaM = intVerticaM + IsVertica_Maisu(.GetRS("�o�^���i��"), .GetRS("����"))
                                '1.10.8 K.Asayama ADD End
                            Case 6
                                If IsStealth_Seizo_TEMP(.GetRS("�o�^���i��")) Then
                                    intStealthM = intStealthM + .GetRS("����")
                                Else
                                    intShitajiM = intShitajiM + .GetRS("����")
                                End If
                                'intShitajiM = intShitajiM + .GetRS("����")
                            Case 7
                                intStealthM = intStealthM + .GetRS("����")
                                
                        End Select
                    '1.10.7 ADD
                    End If
                    '1.10.7 ADD End
                    
                    '1.10.7 Change
                    'If .GetRS("�m��") <> 0 Then intKakuteiM = intKakuteiM + .GetRS("����")
                    If .GetRS("�m��") = 2 Then
                        intKakuteiM = intKakuteiM + .GetRS("����")
                    ElseIf .GetRS("�m��") = 1 Then
                        intKakuteiTempM = intKakuteiTempM + .GetRS("����")
                    End If
                    '1.10.7 Change End
                    
                    .GetRS.MoveNext
                Loop
                
                Graphctl(i).SetTag (CStr(intFlushM))
                Graphctl(i).CaptionSet Graphctl(i).GetTag
                '1.10.8 ADD
                Graphctl(i).SetControlTipText Graphctl(i).GetTag
                '1.10.8 ADD End
                
                If intKakuteiM > 0 Then
                    Graphctl_Kakutei(i).SetTag (CStr(intKakuteiM))
                    Graphctl_Kakutei(i).myVisible (True)
                Else
                    Graphctl_Kakutei(i).SetTag "0"
                    Graphctl_Kakutei(i).myVisible (False)
                End If
                               
                
                Graphctl_Kakutei(i).CaptionSet Graphctl_Kakutei(i).GetTag
                '1.10.8 ADD
                Graphctl_Kakutei(i).SetControlTipText Graphctl_Kakutei(i).GetTag
                '1.10.8 ADD End
                
                '1.10.7 ADD
                If intKakuteiTempM + intKakuteiM > 0 Then
                    Graphctl_Temp(i).SetTag (CStr(intKakuteiM + intKakuteiTempM))
                    Graphctl_Temp(i).myVisible (True)
                Else
                    Graphctl_Temp(i).SetTag "0"
                    Graphctl_Temp(i).myVisible (False)
                End If
                
                Graphctl_Temp(i).CaptionSet Graphctl_Temp(i).GetTag
                '1.10.7 ADD End
                '1.10.8 ADD
                Graphctl_Temp(i).SetControlTipText Graphctl_Temp(i).GetTag
                '1.10.8 ADD End
                
                If intFkamachiM > 0 Then Itemctl(i, 0).myVisible (True): Itemctl(i, 0).SetControlTipText (intFkamachiM) Else Itemctl(i, 0).myVisible (False)
                If intThruM > 0 Then Itemctl(i, 1).myVisible (True): Itemctl(i, 1).SetControlTipText (intThruM) Else Itemctl(i, 1).myVisible (False)
                If intPaintM > 0 Then Itemctl(i, 2).myVisible (True): Itemctl(i, 2).SetControlTipText (intPaintM) Else Itemctl(i, 2).myVisible (False)
                If intAirM > 0 Then Itemctl(i, 3).myVisible (True): Itemctl(i, 3).SetControlTipText (intAirM) Else Itemctl(i, 3).myVisible (False)
                If intMonsterM > 0 Then Itemctl(i, 4).myVisible (True): Itemctl(i, 4).SetControlTipText (intMonsterM) Else Itemctl(i, 4).myVisible (False)
                '1.10.8 Change
                'If intKamachiM > 0 Then Itemctl(i, 5).myVisible (True): Itemctl(i, 5).SetControlTipText (intKamachiM) Else Itemctl(i, 5).myVisible (False)
                If intVerticaM > 0 Then Itemctl(i, 5).myVisible (True): Itemctl(i, 5).SetControlTipText (intVerticaM) Else Itemctl(i, 5).myVisible (False)
                '1.10.8 Change End
                
                If intShitajiM > 0 Then Itemctl(i, 1).myVisible (True): Itemctl(i, 1).SetControlTipText ("���n��"): Itemctl(i, 1).CaptionSet (CStr(intShitajiM)): Itemctl(i, 1).SetWidth (240)
                If intStealthM > 0 Then Itemctl(i, 4).myVisible (True): Itemctl(i, 4).SetControlTipText ("�X�e���X��"): Itemctl(i, 4).CaptionSet (CStr(intStealthM)): Itemctl(i, 4).SetWidth (240)
                
           End If
        End With
    Next
    
    GoTo Exit_SetOrderCount
    
Err_SetOrderCount:
    MsgBox Err.Description

Exit_SetOrderCount:
    Set objREMOTEdb = Nothing


End Function

Public Function fncbolSetComboKubun(inKubun As String, inCombobox As Access.ComboBox) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'�R���{�{�b�N�X�Z�b�g�i���ʁj
'
'   :����
'       inKubun         :�R���{�{�b�N�X�敪��
'       inCombobox      :�R���{�{�b�N�X�I�u�W�F�N�g
'
'   :�߂�l
'       True            :����
'       False           :���s

'1.10.16 K.Asayama ADD
'   ���X�e���X�敪�ǉ�
'--------------------------------------------------------------------------------------------------------------------
    On Error GoTo Err_fncbolSetComboKubun
    
    inCombobox.RowSourceType = "Value List"
    
    If inKubun = "�����敪" Then
        inCombobox.AddItem "����,1", 0
        inCombobox.AddItem "�g,2", 1
        inCombobox.AddItem "���n,3", 2
        '1.10.16 ADD
        inCombobox.AddItem "�X�e���X,4", 3
        inCombobox.value = inCombobox.ItemData(0)
    End If
    
    
    fncbolSetComboKubun = True
    
    GoTo Exit_fncbolSetComboKubun
    
Err_fncbolSetComboKubun:
    fncbolSetComboKubun = False
    MsgBox Err.Description
    
Exit_fncbolSetComboKubun:
    
End Function

Public Function SetBikouData() As Boolean
'--------------------------------------------------------------------------------------------------------------------
'WK_�D�f�[�^_���l�t�@�C�����쐬����
'
'
'   :�߂�l
'       True            :����
'       False           :���s
'1.10.7 K.Asayama ADD 20160108
'       ���쐬����WK_�D�f�[�^������l�t�@�C�����쐬����
'1.10.8 K.Asayama Change 20160114
'       ���o�O�C�� First���Ƃ��܂��f�[�^���o�Ȃ��̂�Max�ɕύX
'--------------------------------------------------------------------------------------------------------------------
    Dim objLOCALdb As New cls_LOCALDB
    
    Dim strSQL As String
    Dim strErrMsg As String
    
    SetBikouData = False
     
    On Error GoTo Err_SetBikouData
    
    strSQL = ""
    
    strSQL = strSQL & "select �_��ԍ�,���ԍ�,�����ԍ�"
'1.10.8 Change
'    strSQL = strSQL & ",First(IIf([�����敪] = 1,[���l],Null)) as Flush���l "
'    strSQL = strSQL & ",First(IIf([�����敪] = 2,[���l],Null)) as F�y���l "
'    strSQL = strSQL & ",First(IIf([�����敪] = 3,[���l],Null)) as �y���l "
'    strSQL = strSQL & ",First(IIf([�����敪] = 4,[���l],Null)) as �g���l "
'    strSQL = strSQL & ",First(IIf([�����敪] = 5,[���l],Null)) as �O���g���l "
'    strSQL = strSQL & ",First(IIf([�����敪] = 6,[���l],Null)) as ���n���l "
'    strSQL = strSQL & ",First(IIf([�����敪] = 7,[���l],Null)) as �X�e���X�g���l "
    strSQL = strSQL & ",Max(IIf([�����敪] = 1,[���l],Null)) as Flush���l "
    strSQL = strSQL & ",Max(IIf([�����敪] = 2,[���l],Null)) as F�y���l "
    strSQL = strSQL & ",Max(IIf([�����敪] = 3,[���l],Null)) as �y���l "
    strSQL = strSQL & ",Max(IIf([�����敪] = 4,[���l],Null)) as �g���l "
    strSQL = strSQL & ",Max(IIf([�����敪] = 5,[���l],Null)) as �O���g���l "
    strSQL = strSQL & ",Max(IIf([�����敪] = 6,[���l],Null)) as ���n���l "
    strSQL = strSQL & ",Max(IIf([�����敪] = 7,[���l],Null)) as �X�e���X�g���l "
'1.10.8 Change End
    strSQL = strSQL & "from WK_�D�f�[�^ "
    strSQL = strSQL & "where ���l is not null "
    strSQL = strSQL & "group by �_��ԍ�,���ԍ�,�����ԍ� "
    
    If Not objLOCALdb.ExecSQL("delete from WK_�D�f�[�^_���l") Then
        Err.Raise 9999, , "���l�f�[�^���[�N�i���[�J���j�������G���["
    End If
    
    With objLOCALdb
        If .ExecSelect(strSQL) Then
            
            Do While Not .GetRS.EOF
                strSQL = "insert into WK_�D�f�[�^_���l ("
                strSQL = strSQL & "�_��ԍ�,���ԍ�,�����ԍ� "
                strSQL = strSQL & ",Flush���l,F�y���l,�g���l,�O���g���l,���n���l,�X�e���X�g���l"
                strSQL = strSQL & ") values ( "
                strSQL = strSQL & "'" & .GetRS![�_��ԍ�] & "','" & .GetRS![���ԍ�] & "','" & .GetRS![�����ԍ�] & "'"
                strSQL = strSQL & "," & varNullChk(.GetRS![Flush���l], 1) & " "
                strSQL = strSQL & "," & varNullChk(.GetRS![F�y���l], 1) & " "
                strSQL = strSQL & "," & varNullChk(.GetRS![�g���l], 1) & " "
                strSQL = strSQL & "," & varNullChk(.GetRS![�O���g���l], 1) & " "
                strSQL = strSQL & "," & varNullChk(.GetRS![���n���l], 1) & " "
                strSQL = strSQL & "," & varNullChk(.GetRS![�X�e���X�g���l], 1) & " "
                strSQL = strSQL & ")"
                
                'Debug.Print strSQL
                
                If Not .ExecSQL(strSQL, strErrMsg) Then
                    Err.Raise 9999, , strErrMsg
                End If
                
                .GetRS.MoveNext
            Loop
        Else
            Err.Raise 9999, , "�D�f�[�^�i���[�J���j�I�[�v���G���[(Input)"
        End If
    End With
    
    SetBikouData = True
    
    GoTo Exit_SetBikouData
    
Err_SetBikouData:
    SetBikouData = False
    MsgBox Err.Description
    
Exit_SetBikouData:
     Set objLOCALdb = Nothing
     
End Function

Public Function varNullChk(in_Data As Variant, in_DBType As Integer) As Variant
'--------------------------------------------------------------------------------------------------------------------
'������Null�̏ꍇ�͕�����[Null]��Ԃ��B����ȊO�͂��̂܂ܕԂ�(DB�C���T�[�g�p)
'
'   :����
'       in_Data     Variant(�^�s�� ex�f�[�^�x�[�X�̃J�����j
'       in_DBType   1:Local(Jet) 2:SQLServer
'
'   :�߂�l
'       Variant�@   ������Null�̏ꍇ�͕�����[Null]�A����ȊO�͂��̂܂�(���t�A������͉��H����j
'
'1.10.7 K.Asayama ADD 20160108
'
'1.10.16 Change
'       ��  ������Ɠ��t�̐���������ւ�
'           ������ł����t�ϊ��\�Ȃ��͓̂��t�ŕϊ�����iSQLServer����_�C���N�g�ɗ���󂯎�����ꍇ�^�𕶎���ƔF�����Ă��܂����߁j
'           �󗓂̕������Null�Ƃ���
'2.0.0
'       ��  �f�[�^���Ɂu'�i�A�|�X�g���t�B�j�v���������ꍇ�u''�v�����ɒu����

'2.1.0
'       ���@�A�|�X�g���t�B�͑S�p�ɒu��������
'       ��  String�l�̓��t������Ĕ��f�����ꍇ������̂ŏC��
'--------------------------------------------------------------------------------------------------------------------
    Dim datDate As Date
    
    If IsNull(in_Data) Then
    
        varNullChk = "Null"
    
'    ElseIf VarType(in_Data) = vbDate Or (VarType(in_Data) = vbString And IsDate(in_Data)) Then
    ElseIf VarType(in_Data) = vbDate Or (VarType(in_Data) = vbString And (CStr(in_Data) Like "#*/#*/#*" Or CStr(in_Data) Like "#*-#*-#*")) Then
    
        If VarType(in_Data) = vbDate Then
            datDate = in_Data
            
            Select Case in_DBType
                Case 1
                    varNullChk = "#" & Format(datDate, "yyyy/mm/dd") & "#"
                Case Else
                    varNullChk = "'" & Format(datDate, "yyyy/mm/dd") & "'"
            End Select
            
        ElseIf IsDate(in_Data) Then
            datDate = CDate(in_Data)
            
            Select Case in_DBType
                Case 1
                    varNullChk = "#" & Format(datDate, "yyyy/mm/dd") & "#"
                Case Else
                    varNullChk = "'" & Format(datDate, "yyyy/mm/dd") & "'"
            End Select
            
        Else
            in_Data = Replace(in_Data, "'", "�f")
            varNullChk = "'" & in_Data & "'"
        End If
        
        
        
    ElseIf VarType(in_Data) = vbString Then
        '1.10.16
        'varNullChk = "'" & in_Data & "'"
        If in_Data = "" Then
            varNullChk = "Null"
        Else
            '��2.0.0 ADD
            in_Data = Replace(in_Data, "'", "�f")
            varNullChk = "'" & in_Data & "'"
        End If
    Else
        varNullChk = in_Data
    End If

End Function

Public Function bolfncTableCopyToLocal(in_RS As ADODB.Recordset, out_LocalTableName As String, Optional in_ADDMode As Boolean = False) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'�����[�gDB�̃��R�[�h�Z�b�g�����[�J���̃e�[�u���ɃR�s�[����
'   �i�����[�g�ƃ��[�J���̃J�������͓����ł���O��j
'
'   :����
'       in_RS                   �����[�g�f�[�^�x�[�X�̃��R�[�h�Z�b�g
'       out_LocalTableName      ���[�J���f�[�^�x�[�X�̃e�[�u����
'       in_ADDMode              True:�ǉ� False:Replace�i�ŏ��Ƀ��R�[�h��DELETE����j
'
'   :�߂�l
'       Boolean�@               True:����   False:���s
'
'1.11.2 ADD
'--------------------------------------------------------------------------------------------------------------------

    Dim objLOCALdb As New cls_LOCALDB
    Dim i As Integer
    Dim strErrMsg As String
    Dim varAutoNumber As Variant
    
    Dim DAODB As DAO.Database
    Dim DAORs As DAO.Recordset
    
    Set DAODB = CurrentDb
    Set DAORs = DAODB.OpenRecordset(out_LocalTableName)
    
    On Error GoTo Err_bolfncTableCopyToLocal
    
    bolfncTableCopyToLocal = False
    
    '�I�[�g�i���o�[�^�`�F�b�N
    '�I�[�g�i���o�[�͈ڑ����Ȃ�
    varAutoNumber = Null
    
    With DAORs
        For i = 0 To .Fields.Count - 1
            If (.Fields(i).Type = dbLong) And (.Fields(i).Attributes And dbAutoIncrField) Then
                varAutoNumber = .Fields(i).Name
                Exit For
            End If
        Next
    End With
    
    DAORs.Close
    DAODB.Close
    
    With objLOCALdb
    
        If Not in_ADDMode Then
            If Not .ExecSQL("delete * from " & out_LocalTableName & " ", strErrMsg) Then
                Err.Raise 9999, , strErrMsg
            Else
                '�I�[�g�i���o�[������
                If Not IsNull(varAutoNumber) Then
                    DoCmd.RunSQL "ALTER TABLE " & out_LocalTableName & " ALTER COLUMN " & varAutoNumber & " COUNTER(1, 1)"
                End If
                
            End If
        End If
        
        If .ExecSelect_Writable("select * from " & out_LocalTableName & " ") Then
        
            in_RS.MoveFirst
            
            Do While Not in_RS.EOF
                .GetRS.AddNew
                
                For i = 0 To .GetRS.Fields.Count - 1
                    If .GetRS.Fields(i).Name <> Nz(varAutoNumber, "") Then
                        .GetRS(.GetRS.Fields(i).Name) = in_RS(.GetRS.Fields(i).Name)
                    End If
                Next
                .GetRS.Update
               in_RS.MoveNext
            Loop
        End If
    End With
    
    bolfncTableCopyToLocal = True
    
    GoTo Exit_bolfncTableCopyToLocal
    
Err_bolfncTableCopyToLocal:
    MsgBox Err.Description
    
Exit_bolfncTableCopyToLocal:
    Set objLOCALdb = Nothing
    Set DAORs = Nothing
    Set DAODB = Nothing
End Function

Public Function bolfncMiseizoToExcel() As Boolean
'--------------------------------------------------------------------------------------------------------------------
'�������f�[�^Excel�փG�N�X�|�[�g
'1.12.2 ADD

'   :����

'   :�߂�l
'       True            :����
'       False           :���s

'2.2.0
'   ���E�H�[���X���[�������ǉ�
'2.8.0
'   ���N���[�b�g���o�גǉ�
'2.13.0
'   ���T�[�o�p�X�����ʕϐ��ɕύX
'--------------------------------------------------------------------------------------------------------------------

    Dim objApp As New cls_Excel
    Dim objREMOTEdb As New cls_BRAND_MASTER
    
    Dim xlsBookName As String
    Dim i As Integer
    Dim intSheetDel As Integer
    Dim strSQL As String
    Dim strSQLJ As String
    Dim strKBName(4) As String
    Dim strMidashiVal As String
    
    On Error GoTo Err_bolfncMiseizoToExcel
    intSheetDel = 0
    
    Screen.MousePointer = 11
    
    With objApp.getExcel

        .Workbooks.Add
        
        strKBName(0) = "����"
        strKBName(1) = "���n"
        strKBName(2) = "�g"
        strKBName(3) = "�E�H�[���X���[���o��"
        strKBName(4) = "�N���[�b�g���o��"
        
        strSQL = ""
        strSQL = strfncTextFileToString(conServerPath & "\SQL\subMISEIZO.sql")
        'strSQL = strfncTextFileToString("\\db\prog\�����Ǘ��V�X�e��\SQL\subMISEIZO.sql")
        If strSQL <> "" Then
            strSQL = Replace(strSQL, vbCrLf, " ")
        Else
            Err.Raise 9999, , "�������o�ُ͈�I��"
        End If
        
        strMidashiVal = "������Y�c " & Format(Now, "yyyy-MM-dd")
        
        If Not objREMOTEdb.ExecSelect(strSQL) Then
            Err.Raise 9999, , "�䒠�W�v�f�[�^�ُ�I��"
        End If
        
        objApp.WorkSheetADD strKBName(i)
                
        If Not bolfncexp_EXCELOBJECT(objREMOTEdb.GetRS, objApp.getExcel, True, strMidashiVal) Then
            Err.Raise 9999, , "Excel�G�N�X�|�[�g�ُ�I��"
        End If
        
        strSQL = ""
        'strSQL = strfncTextFileToString("\\db\prog\�����Ǘ��V�X�e��\SQL\subMISEIZOWaku.sql")
        strSQL = strfncTextFileToString(conServerPath & "\SQL\subMISEIZOWaku.sql")
        If strSQL <> "" Then
            strSQL = Replace(strSQL, vbCrLf, " ")
        Else
            Err.Raise 9999, , "�������o�ُ͈�I��"
        End If
        
        For i = 1 To 2
            strMidashiVal = "�g�����Y�c (" & strKBName(i) & ") "
            
            strSQLJ = Replace(strSQL, "@WakuKubun", "'" & strKBName(i) & "'")

            strMidashiVal = strMidashiVal & " " & Format(Now, "yyyy-MM-dd")
            
            objApp.WorkSheetADD strKBName(i)
            If Not objREMOTEdb.ExecSelect(strSQLJ) Then
                Err.Raise 9999, , "�䒠�W�v�f�[�^�ُ�I��"
            End If
            
            If Not bolfncexp_EXCELOBJECT(objREMOTEdb.GetRS, objApp.getExcel, True, strMidashiVal) Then
                Err.Raise 9999, , "Excel�G�N�X�|�[�g�ُ�I��"
            End If

        Next
        
        i = 3
        strSQL = ""
        'strSQL = strfncTextFileToString("\\db\prog\�����Ǘ��V�X�e��\SQL\subMISHUKKA_Wallthru.sql")
        strSQL = strfncTextFileToString(conServerPath & "\SQL\subMISHUKKA_Wallthru.sql")
        If strSQL <> "" Then
            strSQL = Replace(strSQL, vbCrLf, " ")
        Else
            Err.Raise 9999, , "�������o�ُ͈�I��"
        End If
        
        strMidashiVal = "�E�H�[���X���[���o�׎c " & Format(Now, "yyyy-MM-dd")
        
        If Not objREMOTEdb.ExecSelect(strSQL) Then
            Err.Raise 9999, , "�䒠�W�v�f�[�^�ُ�I��"
        End If
        
        objApp.WorkSheetADD strKBName(i)
                
        If Not bolfncexp_EXCELOBJECT(objREMOTEdb.GetRS, objApp.getExcel, True, strMidashiVal) Then
            Err.Raise 9999, , "Excel�G�N�X�|�[�g�ُ�I��"
        End If
        
        i = 4
        strSQL = ""
        strSQL = strfncTextFileToString("\\db\prog\�����Ǘ��V�X�e��\SQL\subMISHUKKA_Oredo.sql")
        strSQL = strfncTextFileToString(conServerPath & "\SQL\subMISHUKKA_Oredo.sql")
        If strSQL <> "" Then
            strSQL = Replace(strSQL, vbCrLf, " ")
        Else
            Err.Raise 9999, , "�������o�ُ͈�I��"
        End If
        
        strMidashiVal = "�N���[�b�g���o�׎c " & Format(Now, "yyyy-MM-dd")
        
        If Not objREMOTEdb.ExecSelect(strSQL) Then
            Err.Raise 9999, , "�䒠�W�v�f�[�^�ُ�I��"
        End If
        
        objApp.WorkSheetADD strKBName(i)
                
        If Not bolfncexp_EXCELOBJECT(objREMOTEdb.GetRS, objApp.getExcel, True, strMidashiVal) Then
            Err.Raise 9999, , "Excel�G�N�X�|�[�g�ُ�I��"
        End If
        
        '�s�v�ȃ��[�N�V�[�g�̍폜
        For i = 1 To .Worksheets.Count
            If .Worksheets(i - intSheetDel).Name Like "Sheet*" Then
                .Worksheets(i - intSheetDel).Delete
                intSheetDel = intSheetDel + 1
            End If
        Next
        
        .Worksheets(1).Activate
        
        objApp.ContinueOpen = True
        
    End With
    
    bolfncMiseizoToExcel = True
    
    GoTo Exit_bolfncMiseizoToExcel

Err_bolfncMiseizoToExcel:
    Screen.MousePointer = 0
    MsgBox Err.Description
    bolfncMiseizoToExcel = False
    
    
Exit_bolfncMiseizoToExcel:
    Screen.MousePointer = 0
    Set objApp = Nothing
    Set objREMOTEdb = Nothing
    
    
End Function