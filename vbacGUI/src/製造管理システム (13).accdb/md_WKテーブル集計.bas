Attribute VB_Name = "md_WK�e�[�u���W�v"
Option Compare Database
Option Explicit

Public Function SetOrderData(ByVal inDate As Date, ByVal inDatekbn As Byte, inSeizoKbn As String) As Boolean
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
'--------------------------------------------------------------------------------------------------------------------
    Dim objREMOTEDB As New cls_BRAND_MASTER
    Dim objLOCALDB As New cls_LOCALDB
    
    Dim strSQL As String
    Dim bolTran As Boolean
    Dim strKeiyakuno As String
    Dim varCalcShukkaBi As Variant
    Dim intMinusDays As Integer
    Dim dblWindowTop As Double, dblWindowLeft As Double, dblWindowHight As Double, dblWindowWidth As Double
    Dim bolFormOpen As Boolean
    Dim strKubun As String
    
    bolFormOpen = False
    
    On Error GoTo Err_SetOrderData
    
'    Me.Painting = False
    Application.Echo False
    
    SetOrderData = False
    bolTran = False
    strKeiyakuno = ""
    
    Select Case inSeizoKbn
        Case "����"
            strKubun = "1,2,3"
        Case "�g"
            strKubun = "4,5"
        Case "���n"
            strKubun = "6,7"
        Case Else
            Err.Raise 9999, , "�����敪�]���G���["
    End Select
    
    strSQL = ""
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
    strSQL = strSQL & "from T_�����w�� s "
    strSQL = strSQL & "inner join T_��Ͻ� m "
    strSQL = strSQL & "on s.�_��ԍ� = m.�_��ԍ� and s.���ԍ� = m.���ԍ� and s.�����ԍ� = m.�����ԍ� "
    strSQL = strSQL & "left join T_�����\�� y "
    strSQL = strSQL & "on s.�_��ԍ� = y.�_��ԍ� and s.���ԍ� = y.���ԍ� and s.�����ԍ� = y.�����ԍ� and s.�����敪 = y.�����敪 "
    
    If inDatekbn = 1 Then
        strSQL = strSQL & "where s.�m��� = '" & Format(inDate, "yyyy/mm/dd") & "' "
    Else
        strSQL = strSQL & "where s.������ = '" & Format(inDate, "yyyy/mm/dd") & "' "
        strSQL = strSQL & " and �m�� > 0 "
    End If
    strSQL = strSQL & " and s.�����敪 in ( " & strKubun & ") "
    
    '�E�H�[���X���[�͐����������Ă��Ȃ��̂őΏۊO
    If inSeizoKbn = "���n" Then
        strSQL = strSQL & " and s.�o�^���i�� not like 'WS%' "
    End If
    
    
    If Not objLOCALDB.ExecSQL("delete from WK_�D�f�[�^") Then
        Err.Raise 9999, , "�����w���f�[�^���[�N�i���[�J���j�������G���["
    End If
    
    With objREMOTEDB
        If .ExecSelect(strSQL) Then
            If objLOCALDB.ExecSelect_Writable("select * from WK_�D�f�[�^") Then
            
                objLOCALDB.BeginTrans
                bolTran = True
                
                Do While Not .GetRS.EOF
                        objLOCALDB.GetRS.AddNew

                        objLOCALDB.GetRS![�_��ԍ�] = .GetRS![�_��ԍ�]
                        objLOCALDB.GetRS![���ԍ�] = .GetRS![���ԍ�]
                        objLOCALDB.GetRS![�����ԍ�] = .GetRS![�����ԍ�]
                        objLOCALDB.GetRS![������] = .GetRS![������]
                        objLOCALDB.GetRS![�{�H�X] = .GetRS![�{�H�X]
                        objLOCALDB.GetRS![�_��No] = .GetRS![�_��No]
                        objLOCALDB.GetRS![��] = .GetRS![��]
                        objLOCALDB.GetRS![�����敪] = .GetRS![�����敪]
                        objLOCALDB.GetRS![�m���] = .GetRS![�m���]
                        If IsNull(.GetRS![�o�ד�]) Then
                            objLOCALDB.GetRS![�o�ד��o�^] = False
                            If fncbolSyukkaBiFromAddress(.GetRS![�[�i�Z��], .GetRS![�m���], varCalcShukkaBi, intMinusDays) Then
                                objLOCALDB.GetRS![�o�ד�] = CDate(varCalcShukkaBi)
                            Else
                                objLOCALDB.GetRS![�o�ד�] = .GetRS![�o�ד�]
                            End If
                        Else
                            objLOCALDB.GetRS![�o�ד��o�^] = True
                            objLOCALDB.GetRS![�o�ד�] = .GetRS![�o�ד�]
                        End If
                        
                        objLOCALDB.GetRS![������] = .GetRS![������]
                        objLOCALDB.GetRS![�[�i�Z��] = .GetRS![�[�i�Z��]
                        
                        'If IsNull(.GetRS![�m��]) Or .GetRS![�m��] = 0 Then
                        '    objLocalDB.GetRS![�m��] = 0
                        'Else
                        '    objLocalDB.GetRS![�m��] = -1
                        'End If
                        
                        objLOCALDB.GetRS![�m��] = .GetRS![�m��]
                        objLOCALDB.GetRS![Flush��] = .GetRS![Flush��] + .GetRS![F�y��]
                        objLOCALDB.GetRS![F�y��] = .GetRS![F�y��]
                        objLOCALDB.GetRS![�y��] = .GetRS![�y��]
                        objLOCALDB.GetRS![�g��] = .GetRS![�g��]
                        objLOCALDB.GetRS![�O���g��] = .GetRS![�O���g��]
                        'objLOCALDB.GetRS![���n�g��] = .GetRS![���n�g��]
                        'objLOCALDB.GetRS![�X�e���X�g��] = .GetRS![�X�e���X�g��]
                        
                        If IsStealth_Seizo_TEMP(Nz(.GetRS![�o�^���i��], "nz")) Then
                            objLOCALDB.GetRS![���n�g��] = 0
                            objLOCALDB.GetRS![�X�e���X�g��] = .GetRS![���n�g��]
                        Else
                            objLOCALDB.GetRS![�X�e���X�g��] = 0
                            objLOCALDB.GetRS![���n�g��] = .GetRS![���n�g��]
                        End If
                        
                        If .GetRS![�����敪] >= 1 And .GetRS![�����敪] <= 3 Then
                            If IsThruGlass(.GetRS![�o�^���i��]) Then
                                objLOCALDB.GetRS![�X���[�K���X��] = .GetRS![Flush��]
                            Else
                                objLOCALDB.GetRS![�X���[�K���X��] = 0
                            End If
                            
                            If IsAir(.GetRS![�o�^���i��]) Then
                                objLOCALDB.GetRS![���[�o�[����] = .GetRS![Flush��]
                            Else
                                objLOCALDB.GetRS![���[�o�[����] = 0
                            End If
                            
                            If IsPainted(.GetRS![�o�^���i��]) Then
                                objLOCALDB.GetRS![�h������] = .GetRS![Flush��]
                            Else
                                objLOCALDB.GetRS![�h������] = 0
                            End If
                            
                            If IsMonster(.GetRS![�o�^���i��]) Then
                                objLOCALDB.GetRS![�����X�^�[��] = .GetRS![F�y��]
                            Else
                                objLOCALDB.GetRS![�����X�^�[��] = 0
                            End If
                        Else
                            objLOCALDB.GetRS![�X���[�K���X��] = 0
                            objLOCALDB.GetRS![���[�o�[����] = 0
                            objLOCALDB.GetRS![�h������] = 0
                            objLOCALDB.GetRS![�����X�^�[��] = 0
                        End If
                        
                        objLOCALDB.GetRS![���l] = .GetRS![���l]
                        
                    objLOCALDB.GetRS.Update
                    
                    .GetRS.MoveNext
                Loop
                
                If bolTran Then objLOCALDB.Commit
                bolTran = False
            Else
                Err.Raise 9999, , "�`�F�b�N���X�g���[�N�i���[�J���j�I�[�v���G���["
            
            End If
        Else
            Err.Raise 9999, , "�`�F�b�N���X�g���o�G���["
        End If
    End With
    
    DoCmd.SetWarnings False
    
    
    If Form_IsLoaded("F_�@��_����") Then
        bolFormOpen = True
    End If
    
    
    If Not bolFormOpen Then
        DoCmd.OpenForm "F_�@��_����", acNormal, , , , , inDatekbn
    Else
        If Not Form_F_�@��_����.bolfncData_Update(inSeizoKbn) Then
            DoCmd.Close acForm, "F_�@��_����", acSaveNo
        End If
    End If
    
    
    DoCmd.SetWarnings True
    
    SetOrderData = True
    GoTo Exit_SetOrderData
    
Err_SetOrderData:
    If bolTran Then objLOCALDB.Rollback
    bolTran = False
    MsgBox Err.Description

Exit_SetOrderData:
    Set objREMOTEDB = Nothing
    Set objLOCALDB = Nothing

    Application.Echo True
    'Me.Painting = True
    
End Function

Public Function SetOrderCount(ByVal inDatekbn As Byte, ByRef Captionctl() As cls_Labelset, ByRef Graphctl() As cls_Labelset, ByRef Graphctl_Kakutei() As cls_Labelset, ByRef Itemctl() As cls_Labelset, ByVal in_HinbanKubun As Integer, ByVal in_KojoCD As Integer)
'--------------------------------------------------------------------------------------------------------------------
'���ʏW�v����
'
'   :����
'       inDateKbn       :1:�[�i���x�[�X�W�v�A2:�������x�[�X�W�v
'       Captionctl      :���t�\�����x���i�R���g���[���z��j
'       Graphctl        :���ʕ\�����x���i�R���g���[���z��j
'       Graphctl_Kakutei:�m�萔�ʕ\�����x���i�R���g���[���z��j
'       Itemctl         :���i�\�����x���i2�����R���g���[���z��[���t,���i]�j
'       in_HinbanKubun  :1,����A2,�g�A3,���n
'       in_KojoCD       :�H��CD
'
'   :�߂�l
'       True            :����
'       False           :���s
'---------------------------
'   �ύX
'       20151110 K.Asayama ���n���A�X�e���X�������x���\���i�e�X�K���X���A�����X�^�[���𗬗p�j
'--------------------------------------------------------------------------------------------------------------------
    Dim objREMOTEDB As New cls_BRAND_MASTER
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
    
    On Error GoTo Err_SetOrderCount
    
    '���n�Ƌ��p���郉�x������U����ŏ�����
    For i = 0 To UBound(Itemctl)
        Itemctl(i, 1).CaptionSet ("�K���X")
        Itemctl(i, 4).CaptionSet ("Monster")
        
        If inDatekbn = 1 Then
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
        Case 3 'Shitaji
            strKubun = "6,7"
        Case Else
            strKubun = CStr(in_HinbanKubun)
    End Select
    
    strSQL_C = "select s.�o�^���i��, s.�����敪, s.�m��, s.���� as ���� from T_�����w�� s "
    strSQL_C = strSQL_C & "inner join T_�󒍖��� m "
    strSQL_C = strSQL_C & "on m.�_��ԍ� = s.�_��ԍ� and m.���ԍ� = s.���ԍ� and m.�����ԍ� = s.�����ԍ� and m.�� = s.�� "
    
    
    For i = 0 To UBound(Captionctl)
        strSQL = strSQL_C
        If inDatekbn = 1 Then
            strSQL = strSQL & " where s.�m��� = '" & Captionctl(i).GetTag & "'"
        Else
            strSQL = strSQL & " where s.������ = '" & Captionctl(i).GetTag & "'"
            strSQL = strSQL & " and �m�� > 0 "
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
    
        strSQL = strSQL & " and s.�H��CD = " & in_KojoCD
        strSQL = strSQL & " "
        
        With objREMOTEDB
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
                
                Do Until .GetRS.EOF
                    
                    intFlushM = intFlushM + .GetRS("����")
                    
                    Select Case .GetRS("�����敪")
                        Case 1, 2, 3
                            If .GetRS("�����敪") = 2 Then intFkamachiM = intFkamachiM + .GetRS("����")
                            If .GetRS("�����敪") = 3 Then intKamachiM = intKamachiM + .GetRS("����")
                            If IsThruGlass(.GetRS("�o�^���i��")) Then intThruM = intThruM + .GetRS("����")
                            If IsPainted(.GetRS("�o�^���i��")) Then intPaintM = intPaintM + .GetRS("����")
                            If IsAir(.GetRS("�o�^���i��")) Then intAirM = intAirM + .GetRS("����")
                            If IsMonster(.GetRS("�o�^���i��")) Then intMonsterM = intMonsterM + .GetRS("����")
                        Case 6
                            If IsStealth_Seizo_TEMP(.GetRS("�o�^���i��")) Then
                                intStealthM = intStealthM + .GetRS("����")
                            Else
                                intShitajiM = intShitajiM + .GetRS("����")
                            End If
                            
                                
                    End Select
                    
                    If .GetRS("�m��") <> 0 Then intKakuteiM = intKakuteiM + .GetRS("����")
        
                    .GetRS.MoveNext
                Loop
                
                Graphctl(i).SetTag (CStr(intFlushM))
                Graphctl(i).CaptionSet Graphctl(i).GetTag
                
                If intKakuteiM > 0 Then
                    Graphctl_Kakutei(i).SetTag (CStr(intKakuteiM))
                    Graphctl_Kakutei(i).myVisible (True)
                Else
                    Graphctl_Kakutei(i).SetTag "0"
                    Graphctl_Kakutei(i).myVisible (False)
                End If
                
                Graphctl_Kakutei(i).CaptionSet Graphctl_Kakutei(i).GetTag
                
                If intFkamachiM > 0 Then Itemctl(i, 0).myVisible (True): Itemctl(i, 0).SetControlTipText (intFkamachiM) Else Itemctl(i, 0).myVisible (False)
                If intThruM > 0 Then Itemctl(i, 1).myVisible (True): Itemctl(i, 1).SetControlTipText (intThruM) Else Itemctl(i, 1).myVisible (False)
                If intPaintM > 0 Then Itemctl(i, 2).myVisible (True): Itemctl(i, 2).SetControlTipText (intPaintM) Else Itemctl(i, 2).myVisible (False)
                If intAirM > 0 Then Itemctl(i, 3).myVisible (True): Itemctl(i, 3).SetControlTipText (intAirM) Else Itemctl(i, 3).myVisible (False)
                If intMonsterM > 0 Then Itemctl(i, 4).myVisible (True): Itemctl(i, 4).SetControlTipText (intMonsterM) Else Itemctl(i, 4).myVisible (False)
                If intKamachiM > 0 Then Itemctl(i, 5).myVisible (True): Itemctl(i, 5).SetControlTipText (intKamachiM) Else Itemctl(i, 5).myVisible (False)
                
                If intShitajiM > 0 Then Itemctl(i, 1).myVisible (True): Itemctl(i, 1).SetControlTipText ("���n��"): Itemctl(i, 1).CaptionSet (CStr(intShitajiM)): Itemctl(i, 1).SetWidth (240)
                If intStealthM > 0 Then Itemctl(i, 4).myVisible (True): Itemctl(i, 4).SetControlTipText ("�X�e���X��"): Itemctl(i, 4).CaptionSet (CStr(intStealthM)): Itemctl(i, 4).SetWidth (240)
                
           End If
        End With
    Next
    
    GoTo Exit_SetOrderCount
    
Err_SetOrderCount:
    MsgBox Err.Description

Exit_SetOrderCount:
    Set objREMOTEDB = Nothing


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
'--------------------------------------------------------------------------------------------------------------------
    On Error GoTo Err_fncbolSetComboKubun
    
    inCombobox.RowSourceType = "Value List"
    
    If inKubun = "�����敪" Then
        inCombobox.AddItem "����,1", 0
        inCombobox.AddItem "�g,2", 1
        inCombobox.AddItem "���n,3", 2
        inCombobox.value = inCombobox.ItemData(0)
    End If
    
    
    fncbolSetComboKubun = True
    
    GoTo Exit_fncbolSetComboKubun
    
Err_fncbolSetComboKubun:
    fncbolSetComboKubun = False
    MsgBox Err.Description
    
Exit_fncbolSetComboKubun:
    
End Function



