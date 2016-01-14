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
    
    Dim objREMOTEDB As New cls_BRAND_MASTER
    Dim objLOCALDB As New cls_LOCALDB
    
    Dim strSQL_Insert As String
    Dim strSQL As String
    
    '1.10.5 ADD By Asayama �G���[�ǉ� 20151209
    On Error GoTo Err_fncbolCalender_Replace
    
    strSQL_Insert = "Insert into WK_Calendar_�H��(�x��) values (#"
    
    '�H��p�R�s�[�iT_Calendar_�H��)
    If objLOCALDB.ExecSQL("delete from WK_Calendar_�H��") Then
        strSQL = "select �x�� from T_Calendar_�H�� "
        'strSQL = strSQL & "where convert(datetime,�x��) > '" & "2015/01/01" & "'"
        If objREMOTEDB.ExecSelect(strSQL) Then
            Do While Not objREMOTEDB.GetRS.EOF
                If Not objLOCALDB.ExecSQL(strSQL_Insert & objREMOTEDB.GetRS![�x��] & "#)") Then
                    Err.Raise 9999, , "�x���J�����_�[�i�H��j���[�J���R�s�[�G���["
                End If
                objREMOTEDB.GetRS.MoveNext
            Loop
        End If
    End If
    
    strSQL_Insert = "Insert into WK_Calendar_�Ɩ�(�x��) values (#"
    
    '�Ɩ��p�R�s�[�iT_Calendar)
    If objLOCALDB.ExecSQL("delete from WK_Calendar_�Ɩ�") Then
        strSQL = "select �x�� from T_Calendar "
        'strSQL = strSQL & "where convert(datetime,�x��) > '" & "2015/01/01" & "'"
        If objREMOTEDB.ExecSelect(strSQL) Then
            Do While Not objREMOTEDB.GetRS.EOF
                If Not objLOCALDB.ExecSQL(strSQL_Insert & objREMOTEDB.GetRS![�x��] & "#)") Then
                    Err.Raise 9999, , "�x���J�����_�[�i�Ɩ��j���[�J���R�s�[�G���["
                End If
                objREMOTEDB.GetRS.MoveNext
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
    Set objREMOTEDB = Nothing
    Set objLOCALDB = Nothing
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

    Dim objLOCALDB As New cls_LOCALDB
    
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
    
    If objLOCALDB.ExecSelect(strSQL) Then
        Do While Not objLOCALDB.GetRS.EOF
            If datDayBefore = objLOCALDB.GetRS![�x��] Then
                objLOCALDB.GetRS.MoveNext
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
    Set objLOCALDB = Nothing
    
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

    Dim objLOCALDB As New cls_LOCALDB
    
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
    
    If objLOCALDB.ExecSelect(strSQL) Then
        Do While Not objLOCALDB.GetRS.EOF
            If datDayBefore = objLOCALDB.GetRS![�x��] Then
                objLOCALDB.GetRS.MoveNext
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
    Set objLOCALDB = Nothing
    
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
    Dim objLOCALDB As New cls_LOCALDB
    
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
    
    If objLOCALDB.ExecSelect(strSQL) Then
        i = 1
        Do While Not objLOCALDB.GetRS.EOF
        
             If datNextDay = objLOCALDB.GetRS![�x��] Then
                 objLOCALDB.GetRS.MoveNext
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
    Set objLOCALDB = Nothing
    
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

'--------------------------------------------------------------------------------------------------------------------
    Dim objLOCALDB As New cls_LOCALDB
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
   
    '�ȉ��ɊY������s���{���̏ꍇ��2��
    If in_varAddress Like "�k�C��*" Or _
        in_varAddress Like "�X��*" Or in_varAddress Like "��茧*" Or in_varAddress Like "�H�c��*" Or _
        in_varAddress Like "�{�錧*" Or in_varAddress Like "������*" Or in_varAddress Like "�R�`��*" Or _
        in_varAddress Like "�O�d��*" Or in_varAddress Like "���Ɍ�*" Or in_varAddress Like "�a�̎R��*" Or _
        in_varAddress Like "������*" Or in_varAddress Like "���挧*" Or in_varAddress Like "�R����*" Or _
        in_varAddress Like "�L����*" Or in_varAddress Like "���R��*" Or in_varAddress Like "���쌧*" Or _
        in_varAddress Like "���Q��*" Or in_varAddress Like "������*" Or in_varAddress Like "���m��*" Or _
        in_varAddress Like "������*" Or in_varAddress Like "�啪��*" Or in_varAddress Like "���ꌧ*" Or _
        in_varAddress Like "���茧*" Or in_varAddress Like "�{�茧*" Or in_varAddress Like "�F�{��*" Or _
        in_varAddress Like "��������*" Or _
        in_varAddress Like "���ꌧ*" Then
       
            intMinusDays = 2
    Else
    
            intMinusDays = 1
    End If
    
    '��ʕ\���p
    out_MinusDay = intMinusDays
    
    '------------------------------------------------------------
    '�o�ד��Ɣ[�i���̊Ԃɓ��A�j���܂܂�Ă���ꍇ�͂��̓��������Z
    '�i�y�j�͔z�����Ɋ܂܂��j
    datTMPKeisan = in_varNouhinBi
        
    i = intMinusDays
    
    While i <> 0
        '�j���A���j�������ꍇ��1�����Z
        If ktHolidayName(datTMPKeisan) <> "" Or Weekday(datTMPKeisan, vbSunday) = 1 Then '�j�����͓��j
            intMinusDays = intMinusDays + 1
        Else
            i = i - 1
            
        End If
        
        '���t����1����
        datTMPKeisan = DateDiff("d", 1, datTMPKeisan)
    Wend
    '------------------------------------------------------------
    
    '�o�ד��擾
    datTMPSyukkaBi = DateDiff("d", intMinusDays, in_varNouhinBi)
    
    '�o�ד����y���j�łȂ����`�F�b�N�i�c�Ƃ̓y�j���ł��o�ׂ͂��Ȃ��j
    Do
        If ktHolidayName(datTMPSyukkaBi) = "" Then '�j���łȂ�
            If Weekday(datTMPSyukkaBi, vbSunday) = 1 Or Weekday(datTMPSyukkaBi, vbSunday) = 7 Then '�����y
                
            Else    '����
                Exit Do
            End If
        End If
        
        datTMPSyukkaBi = DateDiff("d", 1, datTMPSyukkaBi)
        
    Loop
    
    '��Ђ��x���̏ꍇ�͑O�c�Ɠ���Ԃ�
    strSQL = ""
    strSQL = strSQL & "select �x�� from WK_Calendar_�Ɩ� "
    strSQL = strSQL & "where �x�� =< #" & datTMPSyukkaBi & "# "
    strSQL = strSQL & "order by �x�� desc "
    
    If objLOCALDB.ExecSelect(strSQL) Then
        Do While Not objLOCALDB.GetRS.EOF
            If datTMPSyukkaBi <> objLOCALDB.GetRS![�x��] Then
                Exit Do
            End If
            
            datTMPSyukkaBi = DateDiff("d", 1, datTMPSyukkaBi)
            objLOCALDB.GetRS.MoveNext
            
        Loop
    End If
    
    out_SyukkaBi = datTMPSyukkaBi
    fncbolSyukkaBiFromAddress = True
    
    GoTo Exit_fncbolSyukkaBiFromAddress
    
Err_fncbolSyukkaBiFromAddress:

Exit_fncbolSyukkaBiFromAddress:
    Set objLOCALDB = Nothing
End Function

Public Function IsHoliday(ByVal in_Date As String) As Boolean
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

    Dim objLOCALDB As New cls_LOCALDB
    
    Dim strSQL As String
    
    On Error GoTo Err_IsHoliday
    
    If Not IsDate(in_Date) Then GoTo Err_IsHoliday
    
    strSQL = ""
    strSQL = strSQL & "select �x�� from WK_Calendar_�H�� "
    strSQL = strSQL & "where �x�� = #" & in_Date & "# "
    
    
    If objLOCALDB.ExecSelect(strSQL) Then
        If Not objLOCALDB.GetRS.EOF Then
            IsHoliday = True
        End If
    End If
        
    GoTo Exit_IsHoliday

Err_IsHoliday:
    IsHoliday = False
    
Exit_IsHoliday:
    Set objLOCALDB = Nothing
End Function

Public Function intfncSeizoNissu_FromSyukkaBi(in_varHinban As Variant, in_intDefaultDays As Integer) As Integer
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
'   *************************************************************

    If Not in_varHinban Like "*-####*-*" Then
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 13
        Exit Function
    End If
    
    'Caro(Flush����ɋL�ڂ���)
    If isCaro(in_varHinban) Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
    '�۾ޯ�(Flush����ɋL�ڂ���)
    ElseIf in_varHinban Like "F*CME-####*-*" Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
    '�۾ޯ�(SINA����ɋL�ڂ���)
    ElseIf in_varHinban Like "T*CME-####*-*" Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
    '�۾ޯ�
    ElseIf in_varHinban Like "P*CSA-####*-*" Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
    'Flush
    ElseIf in_varHinban Like "F*-####*-*" Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays
    'F/S
    ElseIf in_varHinban Like "S*-####*-*" Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays
    'LUCENTE
    ElseIf in_varHinban Like "P*-####*-*" Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
    'SINA
    ElseIf in_varHinban Like "T*-####*-*" Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
    'Air
    ElseIf IsAir(in_varHinban) Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
    'MONSTER
    ElseIf IsMonster(in_varHinban) Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
    'PALIO
    ElseIf IsPALIO(in_varHinban) Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 9
    'REALART
    ElseIf IsREALART(in_varHinban) Then
        If IsPainted(in_varHinban) Then
            intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 9
        Else
            intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays
        End If
        
    Else
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 9
    
    End If
    
End Function