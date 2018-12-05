Option Compare Database
Option Explicit

Public Function intfncSeizoNissu(in_varHinban As Variant) As Integer
'   *************************************************************
'   ��������v�����m�F
'   �J�^���O�ɋL�ڂ���Ă���ŒZ�����\������Ԃ�
'
'   �߂�l:Integer
'                       ��  ���v����
'                           �i�ԕs���̏ꍇ��0��Ԃ�
'                           �N���[�[�b�g��0��Ԃ� (�ɐ������Y�ȊO)
'
'    Input����
'       in_strHinban        ����i��
'
'   1.10.7
'           �� ���i�֐��ɒu����
'   *************************************************************

    If Not in_varHinban Like "*-####*-*" Then
        intfncSeizoNissu = 0
        Exit Function
    End If
    
    'Caro(Flush����ɋL�ڂ���)
    If isCaro(in_varHinban) Then
    
        intfncSeizoNissu = 20
    '�۾ޯ�(Flush����ɋL�ڂ���)
    ElseIf in_varHinban Like "F*CME-####*-*" Then
    
        intfncSeizoNissu = 20
    '�۾ޯ�(SINA����ɋL�ڂ���)
    ElseIf in_varHinban Like "T*CME-####*-*" Then
    
        intfncSeizoNissu = 20
    '�۾ޯ�
    ElseIf in_varHinban Like "P*CSA-####*-*" Then
    
        intfncSeizoNissu = 20
    'Flush
    ElseIf in_varHinban Like "F*-####*-*" Then
    
        intfncSeizoNissu = 13
    'F/S
    ElseIf in_varHinban Like "S*-####*-*" Then
    
        intfncSeizoNissu = 13
    'LUCENTE
    ElseIf in_varHinban Like "P*-####*-*" Then
    
        intfncSeizoNissu = 20
    'SINA
    ElseIf in_varHinban Like "T*-####*-*" Then
    
        intfncSeizoNissu = 20
    'Air
    ElseIf IsAir(in_varHinban) Then
    
        intfncSeizoNissu = 20
    'MONSTER
    ElseIf IsMonster(in_varHinban) Then
    
        intfncSeizoNissu = 20
    'PALIO
    ElseIf IsPALIO(in_varHinban) Then
    
        intfncSeizoNissu = 23
    'REALART
    ElseIf IsREALART(in_varHinban) Then
    
        intfncSeizoNissu = 23
        
    Else
    
        intfncSeizoNissu = 23
    
    End If
    
End Function

Public Function fncbolCalender_Replace() As Boolean
'   *************************************************************
'   ���[�J���J�����_�[�u��������
'   �����[�g�f�[�^�x�[�X���烍�[�J���ɃJ�����_�[�f�[�^���R�s�[����
'
'   �߂�l:Boolean
'       ��True              �u������
'       ��False             �u�����s
'
'   1.10.6 K.Asayama ADD 20151211 �R�s�[�ς݂̏ꍇ(bolCalendarCopy=True�j�͏������Ȃ�
'   *************************************************************

    fncbolCalender_Replace = False
    
    If bolCalendarCopy Then
        fncbolCalender_Replace = True
        Exit Function
    End If
    
    Dim objREMOTEdb As New cls_BRAND_MASTER
    Dim objLOCALdb As New cls_LOCALDB
    
    Dim strSQL_Insert As String
    Dim strSQL As String
    
    '1.10.5 ADD By Asayama �G���[�ǉ� 20151209
    On Error GoTo Err_fncbolCalender_Replace
    
    strSQL_Insert = "Insert into WK_Calendar_�H��(�x��) values (#"
    
    '�H��p�R�s�[�iT_Calendar_�H��)
    If objLOCALdb.ExecSQL("delete from WK_Calendar_�H��") Then
        strSQL = "select �x�� from T_Calendar_�H�� "
        'strSQL = strSQL & "where convert(datetime,�x��) > '" & "2015/01/01" & "'"
        If objREMOTEdb.ExecSelect(strSQL) Then
            Do While Not objREMOTEdb.GetRS.EOF
                If Not objLOCALdb.ExecSQL(strSQL_Insert & objREMOTEdb.GetRS![�x��] & "#)") Then
                    Err.Raise 9999, , "�x���J�����_�[�i�H��j���[�J���R�s�[�G���["
                End If
                objREMOTEdb.GetRS.MoveNext
            Loop
        End If
    End If
    
    strSQL_Insert = "Insert into WK_Calendar_�Ɩ�(�x��) values (#"
    
    '�Ɩ��p�R�s�[�iT_Calendar)
    If objLOCALdb.ExecSQL("delete from WK_Calendar_�Ɩ�") Then
        strSQL = "select �x�� from T_Calendar "
        'strSQL = strSQL & "where convert(datetime,�x��) > '" & "2015/01/01" & "'"
        If objREMOTEdb.ExecSelect(strSQL) Then
            Do While Not objREMOTEdb.GetRS.EOF
                If Not objLOCALdb.ExecSQL(strSQL_Insert & objREMOTEdb.GetRS![�x��] & "#)") Then
                    Err.Raise 9999, , "�x���J�����_�[�i�Ɩ��j���[�J���R�s�[�G���["
                End If
                objREMOTEdb.GetRS.MoveNext
            Loop
            fncbolCalender_Replace = True
        End If
    End If
    
    '1.10.6 K.Asayama ADD 20151211 �R�s�[�����̏ꍇ���ʃt���O��True�ɂ���
    bolCalendarCopy = True
    
    GoTo Exit_fncbolCalender_Replace
    
Err_fncbolCalender_Replace:
    MsgBox Err.Description
    
Exit_fncbolCalender_Replace:
    Set objREMOTEdb = Nothing
    Set objLOCALdb = Nothing
End Function

Public Function bolfncCalc_DayOn(in_datNouhinDate As Variant, in_varHinban As Variant, in_intDays As Integer, out_datDay As Variant, out_datNextDay As Variant) As Boolean
'   *************************************************************
'   ����������t���Z����
'   �H��J�����_�[���Q�Ƃ�N����̓��t��Ԃ��iN�c�Ɠ���j
'
'   �߂�l:Boolean
'       ��True              ���t�擾����
'       ��False             ���t�擾�������s
'
'    Input����
'       in_datNouhinDate    Input�p���t
'       in_varHinban        �i��
'       in_intDays          ���Z���t
'    Output����
'       out_datDay          Input�p���t��in_intDays�����Z��̓��t
'       out_datNextDay      out_datDay��1�c�Ɠ���̓��t(F�y�ƋZ���������ȊO��Null�j
'   *************************************************************

    Dim objLOCALdb As New cls_LOCALDB
    
    Dim strSQL As String
    
    Dim datDayBefore As Date

    Dim datNextDay As Date
    
    Dim i As Integer, j As Integer
    
    bolfncCalc_DayOn = False
    
    '1.10.5 ADD By Asayama �G���[�ǉ� 20151209
    On Error GoTo Err_bolfncCalc_DayOn
    
    i = in_intDays
    j = 0
    out_datDay = Null
    out_datNextDay = Null
    
    If Not IsDate(in_datNouhinDate) Then GoTo Err_bolfncCalc_DayOn
    
    datDayBefore = DateDiff("d", -1, in_datNouhinDate)
 
    strSQL = ""
    strSQL = strSQL & "select �x�� from WK_Calendar_�H�� "
    strSQL = strSQL & "where �x�� > #" & in_datNouhinDate & "# "
    strSQL = strSQL & "order by �x�� "
    
    If objLOCALdb.ExecSelect(strSQL) Then
        Do While Not objLOCALdb.GetRS.EOF
            If datDayBefore = objLOCALdb.GetRS![�x��] Then
                objLOCALdb.GetRS.MoveNext
            Else
                i = i - 1
            End If
            
            If i = 0 Then Exit Do
            
            datDayBefore = DateDiff("d", -1, datDayBefore)
            
        Loop
        
        If i <> 0 Then Err.Raise 9999, , "�������擾�G���["
        
        out_datDay = datDayBefore
        
        '�Z��������
        If IsFkamachi(in_varHinban) Or IsGikan(in_varHinban) Then
                
            If Not bolfncNextDate(datDayBefore, out_datNextDay) Then
                Err.Raise 9999, , "�Z���i�y�j�������擾�G���["
            End If
        
'            strSQL = ""
'            strSQL = strSQL & "select �x�� from WK_Calendar_�H�� "
'            strSQL = strSQL & "where �x�� > #" & datDayBefore & "# "
'            strSQL = strSQL & "order by �x�� "
'
'            datNextDay = DateDiff("d", -1, datDayBefore)
'
'            If objLocalDB.ExecSelect(strSQL) Then
'                i = 1
'                Do While Not objLocalDB.GetRS.EOF
'
'                     If datNextDay = objLocalDB.GetRS![�x��] Then
'                         objLocalDB.GetRS.MoveNext
'                     Else
'                         i = i - 1
'                     End If
'
'                     If i = 0 Then Exit Do
'
'                     datNextDay = DateDiff("d", -1, datNextDay)
'
'                Loop
'
'                If i <> 0 Then Err.Raise 9999, , "�Z���i�y�j�������擾�G���["
'
'                out_datNextDay = datNextDay
'
'            Else
'                Err.Raise 9999, , "�x���J�����_�[�擾�G���["
'            End If
'
        End If
    Else
        Err.Raise 9999, , "�x���J�����_�[�擾�G���["
    End If
    
    
    bolfncCalc_DayOn = True
    GoTo Exit_bolfncCalc_DayOn
    
Err_bolfncCalc_DayOn:
    out_datDay = Null
    out_datNextDay = Null
    bolfncCalc_DayOn = False
    
Exit_bolfncCalc_DayOn:
    Set objLOCALdb = Nothing
    
End Function

Public Function bolfncCalc_DayOff(in_datNouhinDate As Variant, in_intDays As Integer, out_datDay As Variant, out_datNextDay As Variant) As Boolean
'   *************************************************************
'   ����������t���Z����
'   �H��J�����_�[���Q�Ƃ�N���O�̓��t��Ԃ��iN�c�Ɠ���j
'
'   �߂�l:Boolean
'       ��True              ���t�擾����
'       ��False             ���t�擾�������s
'
'    Input����
'       in_datNouhinDate    Input�p���t
'       in_intDays          ���Z���t
'    Output����
'       out_datDay          Input�p���t��in_intDays�����Z��̓��t
'       out_datNextDay      out_datDay��1�c�Ɠ���̓��t

'   *************************************************************

    Dim objLOCALdb As New cls_LOCALDB
    
    Dim strSQL As String
    
    Dim datDayBefore As Date

    Dim datNextDay As Date
    
    Dim i As Integer, j As Integer
    
    bolfncCalc_DayOff = False
    
    '1.10.5 ADD By Asayama �G���[�ǉ� 20151209
    On Error GoTo Err_bolfncCalc_DayOff
    
    i = in_intDays
    j = 0
    out_datDay = Null
    out_datNextDay = Null
    
    If Not IsDate(in_datNouhinDate) Then GoTo Err_bolfncCalc_DayOff
    
    datDayBefore = DateDiff("d", 1, in_datNouhinDate)

    strSQL = ""
    strSQL = strSQL & "select �x�� from WK_Calendar_�H�� "
    strSQL = strSQL & "where �x�� < #" & in_datNouhinDate & "# "
    strSQL = strSQL & "order by �x�� desc "
    
    If objLOCALdb.ExecSelect(strSQL) Then
        Do While Not objLOCALdb.GetRS.EOF
            If datDayBefore = objLOCALdb.GetRS![�x��] Then
                objLOCALdb.GetRS.MoveNext
            Else
                i = i - 1
            End If
            
            If i = 0 Then Exit Do
            
            datDayBefore = DateDiff("d", 1, datDayBefore)
            
        Loop
        
        If i <> 0 Then Err.Raise 9999, , "�������擾�G���["
        
        out_datDay = datDayBefore
        
        '�Z��������
        If Not bolfncNextDate(datDayBefore, out_datNextDay) Then
            Err.Raise 9999, , "�Z���i�y�j�������擾�G���["
        End If
        
'            strSQL = ""
'            strSQL = strSQL & "select �x�� from WK_Calendar_�H�� "
'            strSQL = strSQL & "where �x�� > #" & datDayBefore & "# "
'            strSQL = strSQL & "order by �x�� "
'
'            datNextDay = DateDiff("d", -1, datDayBefore)
'
'            If objLocalDB.ExecSelect(strSQL) Then
'                i = 1
'                Do While Not objLocalDB.GetRS.EOF
'
'                     If datNextDay = objLocalDB.GetRS![�x��] Then
'                         objLocalDB.GetRS.MoveNext
'                     Else
'                         i = i - 1
'                     End If
'
'                     If i = 0 Then Exit Do
'
'                     datNextDay = DateDiff("d", -1, datNextDay)
'
'                Loop
'
'                If i <> 0 Then Err.Raise 9999, , "�Z���i�y�j�������擾�G���["
'
'                out_datNextDay = datNextDay
'
'            Else
'                Err.Raise 9999, , "�x���J�����_�[�擾�G���["
'            End If

    Else
        Err.Raise 9999, , "�x���J�����_�[�擾�G���["
    End If
    
    
    bolfncCalc_DayOff = True
    GoTo Exit_bolfncCalc_DayOff
    
Err_bolfncCalc_DayOff:
    out_datDay = Null
    out_datNextDay = Null
    bolfncCalc_DayOff = False
    
Exit_bolfncCalc_DayOff:
    Set objLOCALdb = Nothing
    
End Function

Public Function bolfncNextDate(in_datStartDate As Variant, ByRef out_datNextDay As Variant) As Boolean
'   *************************************************************
'   ����������t���Z�����i�����j
'   input���t�̗��c�Ɠ����擾
'
'   �߂�l:Boolean
'       ��True              ���t�擾����
'       ��False             ���t�擾�������s
'
'    Input����
'       in_datStartDate     Input�p���t
'    Output����
'       out_datNextDay      Input�p���t��1�c�Ɠ���̓��t

'   *************************************************************
    Dim objLOCALdb As New cls_LOCALDB
    
    Dim strSQL As String
    Dim datNextDay As Date
    Dim i As Integer
    
    bolfncNextDate = False
    
    '1.10.5 ADD By Asayama �G���[�ǉ� 20151209
    On Error GoTo Err_bolfncNextDate
    
    strSQL = ""
    strSQL = strSQL & "select �x�� from WK_Calendar_�H�� "
    strSQL = strSQL & "where �x�� > #" & in_datStartDate & "# "
    strSQL = strSQL & "order by �x�� "
    
    datNextDay = DateDiff("d", -1, in_datStartDate)
    
    If objLOCALdb.ExecSelect(strSQL) Then
        i = 1
        Do While Not objLOCALdb.GetRS.EOF
        
             If datNextDay = objLOCALdb.GetRS![�x��] Then
                 objLOCALdb.GetRS.MoveNext
             Else
                 i = i - 1
             End If
             
             If i = 0 Then Exit Do
             
             datNextDay = DateDiff("d", -1, datNextDay)
        
        Loop
        
        If i <> 0 Then Err.Raise 9999, , "�Z���i�y�j�������擾�G���["
        
        out_datNextDay = datNextDay
        
    Else
        Err.Raise 9999, , "�x���J�����_�[�擾�G���[�i�Z���������j"
    End If
            
    bolfncNextDate = True
    GoTo Exit_bolfncNextDate
    
Err_bolfncNextDate:
    out_datNextDay = Null
    bolfncNextDate = False
    
Exit_bolfncNextDate:
    Set objLOCALdb = Nothing
    
End Function

Public Function fncbolSyukkaBiFromAddress(in_varAddress As Variant, in_varNouhinBi As Variant, ByRef out_SyukkaBi As Variant, ByRef out_MinusDay As Integer) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'�Z������o�ד��擾
'   ���[�i��Z������z�������������o���A�o�ד����쐬����
'
'-------------------------------------------------------
'20151021 K.Asayama �t�H�[�����W���[������ړ�
'-------------------------------------------------------
'
'   :����
'       in_varAddress       :�[�t��Z��
'       in_varNouhinBi      :�[�i��
'       out_SyukkaBi        :�o�ד��i�o�́j�@�擾�ł��Ȃ��ꍇ��Null
'       out_MinusDay        :�[�i��-�o�ד��i�c�Ɠ����j

'
'   :�߂�l
'       True            :�擾����
'       False           :�擾���s
'
'   1.10.8 K.Asayama Change 20160114
'           ���k�C���A����̓����ǉ�
'   1.10.13 K.Asayama Change 20170329
'           �����W���[����SQLServer���Ɉړ�
'--------------------------------------------------------------------------------------------------------------------
    '1.10.13
    Dim objREMOTEdb As New cls_BRAND_MASTER
    
    'Dim objLOCALDB As New cls_LOCALDB
    Dim intMinusDays As Integer
    Dim datTMPSyukkaBi As Date
    Dim datTMPKeisan As Date
    Dim i As Integer
    Dim strSQL As String
    
    fncbolSyukkaBiFromAddress = False
    strSQL = ""
    
    On Error GoTo Err_fncbolSyukkaBiFromAddress
    
    If IsNull(in_varAddress) Then
        Exit Function
    End If

'1.10.13 201603**

'    '�ȉ��ɊY������s���{���̏ꍇ��2��
'    If in_varAddress Like "�X��*" Or in_varAddress Like "��茧*" Or in_varAddress Like "�H�c��*" Or _
'        in_varAddress Like "�{�錧*" Or in_varAddress Like "������*" Or in_varAddress Like "�R�`��*" Or _
'        in_varAddress Like "�O�d��*" Or in_varAddress Like "���Ɍ�*" Or in_varAddress Like "�a�̎R��*" Or _
'        in_varAddress Like "������*" Or in_varAddress Like "���挧*" Or in_varAddress Like "�R����*" Or _
'        in_varAddress Like "�L����*" Or in_varAddress Like "���R��*" Or in_varAddress Like "���쌧*" Or _
'        in_varAddress Like "���Q��*" Or in_varAddress Like "������*" Or in_varAddress Like "���m��*" Or _
'        in_varAddress Like "������*" Or in_varAddress Like "�啪��*" Or in_varAddress Like "���ꌧ*" Or _
'        in_varAddress Like "���茧*" Or in_varAddress Like "�{�茧*" Or in_varAddress Like "�F�{��*" Or _
'        in_varAddress Like "��������*" _
'    Then
'
'        intMinusDays = 2
'
'    '1.10.8 ADD
'    ElseIf in_varAddress Like "�k�C��*" Then
'        intMinusDays = 3
'
'    ElseIf in_varAddress Like "���ꌧ*" Then
'        intMinusDays = 7
'    '1.10.8 ADD End
'    Else
'
'            intMinusDays = 1
'    End If
'
'    '��ʕ\���p
'    out_MinusDay = intMinusDays
'
'    '------------------------------------------------------------
'    '�o�ד��Ɣ[�i���̊Ԃɓ��A�j���܂܂�Ă���ꍇ�͂��̓��������Z
'    '�i�y�j�͔z�����Ɋ܂܂��j
'    datTMPKeisan = in_varNouhinBi
'
'    i = intMinusDays
'
'    While i <> 0
'        '�j���A���j�������ꍇ��1�����Z
'        If ktHolidayName(datTMPKeisan) <> "" Or Weekday(datTMPKeisan, vbSunday) = 1 Then '�j�����͓��j
'            intMinusDays = intMinusDays + 1
'        Else
'            i = i - 1
'
'        End If
'
'        '���t����1����
'        datTMPKeisan = DateDiff("d", 1, datTMPKeisan)
'    Wend
'    '------------------------------------------------------------
'
'    '�o�ד��擾
'    datTMPSyukkaBi = DateDiff("d", intMinusDays, in_varNouhinBi)
'
'    '�o�ד����y���j�łȂ����`�F�b�N�i�c�Ƃ̓y�j���ł��o�ׂ͂��Ȃ��j
'    Do
'        If ktHolidayName(datTMPSyukkaBi) = "" Then '�j���łȂ�
'            If Weekday(datTMPSyukkaBi, vbSunday) = 1 Or Weekday(datTMPSyukkaBi, vbSunday) = 7 Then '�����y
'
'            Else    '����
'                Exit Do
'            End If
'        End If
'
'        datTMPSyukkaBi = DateDiff("d", 1, datTMPSyukkaBi)
'
'    Loop
'
'    '��Ђ��x���̏ꍇ�͑O�c�Ɠ���Ԃ�
'    strSQL = ""
'    strSQL = strSQL & "select �x�� from WK_Calendar_�Ɩ� "
'    strSQL = strSQL & "where �x�� =< #" & datTMPSyukkaBi & "# "
'    strSQL = strSQL & "order by �x�� desc "
'
'    If objLOCALDB.ExecSelect(strSQL) Then
'        Do While Not objLOCALDB.GetRS.EOF
'            If datTMPSyukkaBi <> objLOCALDB.GetRS![�x��] Then
'                Exit Do
'            End If
'
'            datTMPSyukkaBi = DateDiff("d", 1, datTMPSyukkaBi)
'            objLOCALDB.GetRS.MoveNext
'
'        Loop
'    End If

    
    strSQL = ""
    strSQL = strSQL & "select dbo.fnc�o�׏��v�����擾('" & in_varAddress & "' ) AS �o�׏��v���� "
    If IsDate(in_varNouhinBi) Then
        strSQL = strSQL & ",dbo.fnc�o�ד��擾('" & in_varAddress & "','" & Format(in_varNouhinBi, "yyyy-mm-dd") & "') AS �o�ד� "
    Else
        strSQL = strSQL & ",Null AS �o�ד� "
    End If
    
    If objREMOTEdb.ExecSelect(strSQL) Then
        If Not objREMOTEdb.GetRS.EOF Then
            out_MinusDay = objREMOTEdb.GetRS("�o�׏��v����")
            '1.10.14 ���[�J�����t�^���ɕϊ�
            If IsNull(objREMOTEdb.GetRS("�o�ד�")) Then
                out_SyukkaBi = Null
            Else
                out_SyukkaBi = CDate(objREMOTEdb.GetRS("�o�ד�"))
            End If
        Else
            out_MinusDay = 0
            out_SyukkaBi = Null
        End If
    Else
        out_MinusDay = 0
        out_SyukkaBi = Null

    End If
    
    
    fncbolSyukkaBiFromAddress = True
    
    GoTo Exit_fncbolSyukkaBiFromAddress
    
Err_fncbolSyukkaBiFromAddress:

Exit_fncbolSyukkaBiFromAddress:
    'Set objLOCALDB = Nothing
    Set objREMOTEdb = Nothing
End Function

Public Function IsHoliday(ByVal in_date As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'   ��������x���m�F����
'   �������傪�x�����ǂ����m�F
'

'   Ver 1.01.* K.Asayama ADD 201510**
'
'   �߂�l:Boolean
'       ��True              �x��
'       ��False             �ғ���
'
'    Input����
'       in_Date     ���t�i������^���j

'--------------------------------------------------------------------------------------------------------------------

    Dim objLOCALdb As New cls_LOCALDB
    
    Dim strSQL As String
    
    On Error GoTo Err_IsHoliday
    
    If Not IsDate(in_date) Then GoTo Err_IsHoliday
    
    strSQL = ""
    strSQL = strSQL & "select �x�� from WK_Calendar_�H�� "
    strSQL = strSQL & "where �x�� = #" & in_date & "# "
    
    
    If objLOCALdb.ExecSelect(strSQL) Then
        If Not objLOCALdb.GetRS.EOF Then
            IsHoliday = True
        End If
    End If
        
    GoTo Exit_IsHoliday

Err_IsHoliday:
    IsHoliday = False
    
Exit_IsHoliday:
    Set objLOCALdb = Nothing
End Function

Public Function intfncSeizoNissu_FromSyukkaBi(in_varHinban As Variant, in_Kubun As Integer) As Integer
'   *************************************************************
'   ��������v�����m�F�i�o�ד����t�Z�j
'   �o�ד���萻���\�����v�Z����
'
'   1.10.7 ADD
'
'   �߂�l:Integer
'                       ��  ���v����
'                           �i�ԕs���̏ꍇ�͍ő�����i�h�����j��Ԃ�
'                           �N���[�[�b�g��0��Ԃ� (�ɐ������Y�ȊO)
'
'    Input����
'       in_strHinban        ����i��
'       in_intDefaultDays   �W���i(CUBE�����v�����j

'   1.10.11 K.Asayama Chenge
'           ���p���I�A���A���[�g��+9����+11��
'           ���N���[�b�g���f�t�H���g���t��
'   1.10.13 K.Asayama Change
'           �����W���[����SQLServer���Ɉړ�
'           �������ύX�@in_intDefaultDays��in_Kubun�i�����敪�j
'   *************************************************************

    Dim objREMOTEdb As New cls_BRAND_MASTER
    
    Dim strSQL As String
    
    intfncSeizoNissu_FromSyukkaBi = 0
    
    On Error GoTo Err_intfncSeizoNissu_FromSyukkaBi
    
    If IsNull(in_varHinban) Or in_Kubun = 0 Then
        Exit Function
    End If
    
    strSQL = ""
    strSQL = strSQL & "select dbo.fncSeizoNissu_FromSyukkaBi('" & in_varHinban & "'," & in_Kubun & ") AS �������� "
    
    If objREMOTEdb.ExecSelect(strSQL) Then
        If Not objREMOTEdb.GetRS.EOF Then
            intfncSeizoNissu_FromSyukkaBi = objREMOTEdb.GetRS("��������")
        End If
    End If
    
    GoTo Exit_intfncSeizoNissu_FromSyukkaBi
    
Err_intfncSeizoNissu_FromSyukkaBi:
    MsgBox Err.Description
    intfncSeizoNissu_FromSyukkaBi = 0
    
Exit_intfncSeizoNissu_FromSyukkaBi:
    Set objREMOTEdb = Nothing
    
'    If Not in_varHinban Like "*-####*-*" Then
'        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 11
'        Exit Function
'    End If
'
'    'Caro(Flush����ɋL�ڂ���)
'    If isCaro(in_varHinban) Then
'
'        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
'    '�۾ޯ�(Flush����ɋL�ڂ���)
'    ElseIf in_varHinban Like "F*CME-####*-*" Then
'
'        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays
'    '�۾ޯ�(SINA����ɋL�ڂ���)
'    ElseIf in_varHinban Like "T*CME-####*-*" Then
'
'        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays
'    '�۾ޯ�
'    ElseIf in_varHinban Like "P*CSA-####*-*" Then
'
'        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays
'    'Flush
'    ElseIf in_varHinban Like "F*-####*-*" Then
'
'        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays
'    'F/S
'    ElseIf in_varHinban Like "S*-####*-*" Then
'
'        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays
'    'LUCENTE
'    ElseIf in_varHinban Like "P*-####*-*" Then
'
'        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
'    'SINA
'    ElseIf in_varHinban Like "T*-####*-*" Then
'
'        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
'    'Air
'    ElseIf IsAir(in_varHinban) Then
'
'        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
'    'MONSTER
'    ElseIf IsMonster(in_varHinban) Then
'
'        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
'    'PALIO
'    ElseIf IsPALIO(in_varHinban) Then
'
'        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 11
'    'REALART
'    ElseIf IsREALART(in_varHinban) Then
'        If IsPainted(in_varHinban) Then
'            intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 11
'        Else
'            intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays
'        End If
'
'    Else
'
'        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 11
'
'    End If
    
End Function

Public Function datGetShukkaBi(in_KeiyakuNo As Variant, in_TouNo As Variant, in_HeyaNo As Variant, in_intKubun As Integer) As Variant
'   *************************************************************
'   �_��ԍ����̍ŏ��o�ד��擾
'
'   1.10.13 ADD
'
'   �߂�l:Variant(Date)
'          ��  �o�ד��i�擾�ł��Ȃ������ꍇ��Null�j
'
'    Input����
'       in_KeiyakuNo        �_��ԍ�
'       in_TouNo            ���ԍ�
'       in_HeyaNo           �����ԍ�
'       in_intKubun         �����敪

'1.10.16 K.Asayama ADD
'   ���W�v���@�ύX(BugFix)
'2.0.0
'   ���H��CD 10 �ǉ�
'2.5.0
'   ���o�ד������[�h�^�C����ɕύX
'   *************************************************************

    Dim objREMOTEdb As New cls_BRAND_MASTER
    
    Dim strSQL As String
    Dim intKubun As Integer
    Dim intNoukiKubun As Integer
    Dim strLTColumnName As String
    
    datGetShukkaBi = Null
    
    On Error GoTo Err_datGetShukkaBi
    
    If IsNull(in_KeiyakuNo) Or IsNull(in_TouNo) Or IsNull(in_HeyaNo) Or in_intKubun = 0 Then
        Exit Function
    End If
        
    Select Case in_intKubun
        Case 1, 2, 3
            intKubun = 1
            intNoukiKubun = 1
            strLTColumnName = "����LT"
        Case 4
            intKubun = 2
            intNoukiKubun = 2
            strLTColumnName = "�gLT"
        Case 5
            intKubun = 2
            intNoukiKubun = 5
            strLTColumnName = "�gLT"
        Case 6, 7
            intKubun = 3
            intNoukiKubun = 3
            strLTColumnName = "���nLT"
    End Select
    
    '�o�ד����L�ڍς݂̏ꍇ�͏o�ד��A�����łȂ��ꍇ�͔[������v�Z�����o�ד���}��
    
    strSQL = ""
    strSQL = strSQL & "select "
    strSQL = strSQL & "Format(Min(dbo.fncSeizoSyukkaDate(J.�_��ԍ�,J.���ԍ�,J.�����ԍ�,J.��," & intNoukiKubun & ")),'yyyy-MM-dd') AS �o�ד� "
'    strSQL = strSQL & ",Format(min(dbo.fnc�o�ד��擾(dbo.fncNohinAddress_DefaultGenba(J.�_��ԍ�,J.���ԍ�,J.�����ԍ�,J.��," & intNoukiKubun & ")"
'    strSQL = strSQL & ",(dbo.fncSeizoNohinDate(J.�_��ԍ�,J.���ԍ�,J.�����ԍ�,J.��," & intKubun & ")))),'yyyy-MM-dd') AS �v�Z�o�ד� "
    strSQL = strSQL & ",Format(min(dbo.fnc�o�ד��擾_LT�̂�(dbo.fncSeizoNohinDate(J.�_��ԍ�,J.���ԍ�,J.�����ԍ�,J.��," & intKubun & ")," & strLTColumnName & ")),'yyyy-MM-dd') AS �v�Z�o�ד� "
    
    strSQL = strSQL & "from T_�󒍖��� J "
    strSQL = strSQL & "inner join  T_��Ͻ�_2 JM2 "
    strSQL = strSQL & "on J.�_��ԍ� = JM2.�_��ԍ� and J.���ԍ� = JM2.���ԍ� and J.�����ԍ� = JM2.�����ԍ� "
    '1.10.16 Change
    'strSQL = strSQL & "left join T_�����w�� S "
    strSQL = strSQL & "left join (select * from T_�����w�� where �����敪 = " & in_intKubun & " "
    strSQL = strSQL & "and �_��ԍ� = '" & in_KeiyakuNo & "' and ���ԍ� = '" & in_TouNo & "' and �����ԍ� = '" & in_HeyaNo & "' "
    strSQL = strSQL & ") S "
    strSQL = strSQL & "on J.�_��ԍ� = S.�_��ԍ� and J.���ԍ� = S.���ԍ� and J.�����ԍ� = S.�����ԍ� and J.�� = S.�� "
    strSQL = strSQL & "where J.�_��ԍ� = '" & in_KeiyakuNo & "' and J.���ԍ� = '" & in_TouNo & "' and J.�����ԍ� = '" & in_HeyaNo & "' "
    '1.10.15
    'strSQL = strSQL & "and S.�����敪 = " & in_intKubun & " "
    '1.10.16 DEL
    'strSQL = strSQL & "and (S.�����敪 = " & in_intKubun & " or S.�����敪 is null) "
    strSQL = strSQL & "and (S.�m�� = 0 or S.�m�� is Null) "
    '1.10.16
    'strSQL = strSQL & "and J.��� = '�o����' "
    strSQL = strSQL & "and (J.��� = '�o����' or J.��� = '�۾ޯ�') "
    
    If intKubun = 1 Then
        
        strSQL = strSQL & "and J.�H��CD in (1,10) "

    End If
    
    
    If objREMOTEdb.ExecSelect(strSQL) Then
        If Not objREMOTEdb.GetRS.EOF Then
            If Not IsNull(objREMOTEdb.GetRS("�o�ד�")) Then
                datGetShukkaBi = CDate(objREMOTEdb.GetRS("�o�ד�"))
            ElseIf Not IsNull(objREMOTEdb.GetRS("�v�Z�o�ד�")) Then
                datGetShukkaBi = CDate(objREMOTEdb.GetRS("�v�Z�o�ד�"))
            End If
        End If
    End If
    
    
    GoTo Exit_datGetShukkaBi
    
Err_datGetShukkaBi:
    datGetShukkaBi = Null
    
Exit_datGetShukkaBi:

    Set objREMOTEdb = Nothing
    
End Function

Public Function bolfncDateCheck(ByVal inputMode As Byte, ByVal in_txtDate As String, ByRef out_txtDate As String) As Boolean
'   *************************************************************
'   ���t���̓`�F�b�N
'
'   1.11.0 ADD
'
'   �߂�l:Boolean
'           ��  True        ���t�`�F�b�NOK
'           ��  False       ���t�`�F�b�NNG
'
'    Input����
'       inputMode           ���̓��[�h 0���`�F�b�N�̂݁iout_txtDate�����������Ȃ��j 1���u����(out_txtDate����������)
'       in_txtDate          ���t �^�����R ��������"/"�i�X���b�V���j��؂�
'       out_txtDate         ���t yyyy/MM/dd

'   *************************************************************
    Dim i As Integer
    Dim j As Integer
    
    Dim strTxt As String
    
    Dim strYY As String
    Dim strMM As String
    Dim strDD As String
    
    Dim datNOW As Date
    
    On Error GoTo Err_bolfncDateCheck
    
    i = 1
    j = 0
    
    'input���󗓂̏ꍇ�͖���
    If in_txtDate = "" Then
        bolfncDateCheck = True
        Exit Function
    End If
    
    strTxt = in_txtDate
    
    Do Until InStr(strTxt, "/") = 0
        i = InStr(strTxt, "/")
        strTxt = Mid(strTxt, i + 1)
        If i <> 0 Then j = j + 1
    Loop

    Select Case j
        Case 1 '���Ɠ�
            i = InStr(in_txtDate, "/")
            strMM = left(in_txtDate, i - 1)
            strDD = Mid(in_txtDate, i + 1)
            
            '�N�������t��
            '�������O�̌��̏ꍇ�͗��N
            If CInt(strMM) < CInt(Month(Now())) Then
                strYY = CStr(CInt(Year(Now())) + 1)
                
                '�⊮�������ʂ��������5�����ȏ��̏ꍇ�͌x���\��
                If inputMode = 1 And DateDiff("M", CDate(Year(Now()) & "/" & Month(Now()) & "/01"), CDate(strYY & "/" & strMM & "/01")) > 4 Then
                    MsgBox "�N�����͂���Ă��Ȃ��̂ŗ��N(" & CStr(CInt(Year(Now()) + 1)) & ")��⊮���܂�" & vbCrLf & _
                            "�{�N�̏ꍇ�͔N�����������Ă�������" & vbCrLf & vbCrLf & _
                            "���{���b�Z�[�W�͔N���Ԃ������t���������5�����ȏ��ɂȂ����ꍇ�ɕ\������܂�", vbExclamation, "����!"
                End If
            Else
                strYY = CStr(CInt(Year(Now())))
            End If


        Case 2 '�N����
            i = InStr(in_txtDate, "/")
            strYY = left(in_txtDate, i - 1)
            j = InStr(i + 1, in_txtDate, "/")
            strMM = Mid(in_txtDate, i + 1, (j - 1) - i)
            strDD = Mid(in_txtDate, j + 1)

    End Select

'    MsgBox strYY & "/" & strMM & "/" & strDD
    
    If IsDate(strYY & "/" & strMM & "/" & strDD) Then
        out_txtDate = Format(strYY & "/" & strMM & "/" & strDD, "yyyy/MM/dd")
        If IsHoliday(out_txtDate) Then
            Err.Raise 9999, , "���̓��͋x���ł�"
        End If
        bolfncDateCheck = True
    Else
        Err.Raise 9999, , "���t���͌��"
        
    End If
    
    Exit Function
    
Err_bolfncDateCheck:
    out_txtDate = ""
    bolfncDateCheck = False
    
    If inputMode = 0 Then 'BeforeUpdate�̎��̂݃��b�Z�[�W�o��
        MsgBox Err.Description, vbCritical
    End If
    
End Function

Public Function fncbolSyukkaBiFromLeadTime(in_varLT As Variant, in_varNouhinBi As Variant, ByRef out_SyukkaBi As Variant, ByRef out_MinusDay As Integer) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'���[�h�^�C������o�ד��擾

'   :����
'       in_varLT            :���[�h�^�C��
'       in_varNouhinBi      :�[�i��
'       out_SyukkaBi        :�o�ד��i�o�́j�@�擾�ł��Ȃ��ꍇ��Null
'       out_MinusDay        :���[�h�^�C�������̂܂ܕԂ��i���֐��Ƃ̌݊����̂��߁j

'
'   :�߂�l
'       True            :�擾����
'       False           :�擾���s
'
'   2.5.0 ADD
'--------------------------------------------------------------------------------------------------------------------

    Dim objREMOTEdb As New cls_BRAND_MASTER

    Dim intMinusDays As Integer
    Dim datTMPSyukkaBi As Date
    Dim datTMPKeisan As Date
    Dim i As Integer
    Dim strSQL As String
    
    fncbolSyukkaBiFromLeadTime = False
    strSQL = ""
    
    On Error GoTo Err_fncbolSyukkaBiFromLeadTime
    
    If IsNull(in_varLT) Then
        Exit Function
    End If

    If IsNumeric(in_varLT) Then
        intMinusDays = in_varLT
    Else
        Exit Function
    End If
    
    strSQL = ""
    If IsDate(in_varNouhinBi) Then
        strSQL = strSQL & "select dbo.fnc�o�ד��擾_LT�̂�('" & Format(in_varNouhinBi, "yyyy-mm-dd") & "'," & intMinusDays & ") AS �o�ד� "

    
        If objREMOTEdb.ExecSelect(strSQL) Then
            If Not objREMOTEdb.GetRS.EOF Then
                If IsNull(objREMOTEdb.GetRS("�o�ד�")) Then
                    out_SyukkaBi = Null
                Else
                    out_SyukkaBi = CDate(objREMOTEdb.GetRS("�o�ד�"))
                End If
            Else
                out_SyukkaBi = Null
            End If
        Else
            out_SyukkaBi = Null
    
        End If
    
    Else
        out_SyukkaBi = Null
    
    End If
    
    out_MinusDay = intMinusDays
    
    fncbolSyukkaBiFromLeadTime = True
    
    GoTo Exit_fncbolSyukkaBiFromLeadTime
    
Err_fncbolSyukkaBiFromLeadTime:

Exit_fncbolSyukkaBiFromLeadTime:
    Set objREMOTEdb = Nothing
End Function