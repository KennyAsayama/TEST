Option Compare Database
Option Explicit
'2.1.0 ADD

Public Function bolfnc�����w���f�[�^���o(Optional SeizoDate As Date = #6/19/2017#) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'�w��̐������̎w���f�[�^�A���ޓW�J�����[�N�t�@�C���ɍ쐬����
'
'   :����
'       ������
'
'   :�߂�l
'       True            :����
'       False           :���s

'--------------------------------------------------------------------------------------------------------------------

    Dim objREMOTEDB As New cls_BRAND_MASTER
    Dim objLocalDB As New cls_LOCALDB
    Dim objTateguHinban As New cls_����i��
    
    Dim conSQL As String
    Dim strSQL As String
    Dim strErrMsg As String
    Dim intCnt As Integer
    Dim intMaisu As Integer
    Dim varBrandHinban As Variant
    Dim ChuuiFlg As Boolean
    
    bolfnc�����w���f�[�^���o = False
    
    On Error GoTo Err_bolfnc�����w���f�[�^���o
    
    conSQL = conSQL & "insert into WK_�����˗������� "
    conSQL = conSQL & "( "
    conSQL = conSQL & "�ʔ� "
    conSQL = conSQL & ",�_��ԍ� "
    conSQL = conSQL & ",���ԍ� "
    conSQL = conSQL & ",�����ԍ� "
    conSQL = conSQL & ",�_��No "
    conSQL = conSQL & ",������ "
    conSQL = conSQL & ",�{�H�X "
    conSQL = conSQL & ",�� "
    conSQL = conSQL & ",�ݒu�ꏊ "
    conSQL = conSQL & ",�i��1 "
    conSQL = conSQL & ",�q�� "
    conSQL = conSQL & ",���i�� "
    conSQL = conSQL & ",�F "
    conSQL = conSQL & ",�F�R�[�h "
    conSQL = conSQL & ",�݌� "
    conSQL = conSQL & ",�J�l�� "
    conSQL = conSQL & ",�{�� "
    conSQL = conSQL & ",���� "
    conSQL = conSQL & ",���� "
    conSQL = conSQL & ",�i�ԋ敪 "
    conSQL = conSQL & ",�����敪 "
    conSQL = conSQL & ",�H��CD "
    conSQL = conSQL & ",�ǉ� "
    conSQL = conSQL & ",�K "
    conSQL = conSQL & ",���q���W "
    conSQL = conSQL & ",���iFLG "
    conSQL = conSQL & ",��� "
    conSQL = conSQL & ",�H�� "
    conSQL = conSQL & ",���d�l "
    conSQL = conSQL & ",�o�ו��@ "
    conSQL = conSQL & ",�݌v���l "
    conSQL = conSQL & ",08�J�^���O "
    conSQL = conSQL & ",�W���n���h�� "
    conSQL = conSQL & ",DW "
    conSQL = conSQL & ",DH "
    conSQL = conSQL & ",CH "
    conSQL = conSQL & ",���葋 "
    conSQL = conSQL & ",��Spec "
    conSQL = conSQL & ",Spec "
    conSQL = conSQL & ",Style "
    conSQL = conSQL & ",�󒍖���Style "
    conSQL = conSQL & ",�V�ʍފ��t "
    conSQL = conSQL & ",����FLG "
    conSQL = conSQL & ",�����{�H�� "
    conSQL = conSQL & ",����o�ו��@ "
    conSQL = conSQL & ",�N���[���p���l "
    conSQL = conSQL & ",����m��� "
    conSQL = conSQL & ",������ "
    conSQL = conSQL & ") values ("
    conSQL = conSQL & "@�ʔ�@ "
    conSQL = conSQL & ",@�_��ԍ�@ "
    conSQL = conSQL & ",@���ԍ�@ "
    conSQL = conSQL & ",@�����ԍ�@ "
    conSQL = conSQL & ",@�_��No@ "
    conSQL = conSQL & ",@������@ "
    conSQL = conSQL & ",@�{�H�X@ "
    conSQL = conSQL & ",@��@ "
    conSQL = conSQL & ",@�ݒu�ꏊ@ "
    conSQL = conSQL & ",@�i��1@ "
    conSQL = conSQL & ",@�q��@ "
    conSQL = conSQL & ",@���i��@ "
    conSQL = conSQL & ",@�F@ "
    conSQL = conSQL & ",@�F�R�[�h@ "
    conSQL = conSQL & ",@�݌�@ "
    conSQL = conSQL & ",@�J�l��@ "
    conSQL = conSQL & ",@�{��@ "
    conSQL = conSQL & ",@����@ "
    conSQL = conSQL & ",@����@ "
    conSQL = conSQL & ",@�i�ԋ敪@ "
    conSQL = conSQL & ",@�����敪@ "
    conSQL = conSQL & ",@�H��CD@ "
    conSQL = conSQL & ",@�ǉ�@ "
    conSQL = conSQL & ",@�K@ "
    conSQL = conSQL & ",@���q���W@ "
    conSQL = conSQL & ",@���iFLG@ "
    conSQL = conSQL & ",@���@ "
    conSQL = conSQL & ",@�H��@ "
    conSQL = conSQL & ",@���d�l@ "
    conSQL = conSQL & ",@�o�ו��@@ "
    conSQL = conSQL & ",@�݌v���l@ "
    conSQL = conSQL & ",@08�J�^���O@ "
    conSQL = conSQL & ",@�W���n���h��@ "
    conSQL = conSQL & ",@DW@ "
    conSQL = conSQL & ",@DH@ "
    conSQL = conSQL & ",@CH@ "
    conSQL = conSQL & ",@���葋@ "
    conSQL = conSQL & ",@��Spec@ "
    conSQL = conSQL & ",@Spec@ "
    conSQL = conSQL & ",@Style@ "
    conSQL = conSQL & ",@�󒍖���Style@ "
    conSQL = conSQL & ",@�V�ʍފ��t@ "
    conSQL = conSQL & ",@����FLG@ "
    conSQL = conSQL & ",@�����{�H��@ "
    conSQL = conSQL & ",@����o�ו��@@ "
    conSQL = conSQL & ",@�N���[���p���l@ "
    conSQL = conSQL & ",@����m���@ "
    conSQL = conSQL & ",@������@ "
    conSQL = conSQL & ") "
    
    strSQL = strSQL & "select "
    strSQL = strSQL & "case "
    strSQL = strSQL & "when replace(dbo.fncgethinban(�i��1,��������i��),'�� ','') like 'S%(ZZ)' then 1 "
    strSQL = strSQL & "when  replace(dbo.fncgethinban(�i��1,��������i��),'�� ','') like '%SA-[0-9][0-9][0-9][0-9]%' then 3 "
    strSQL = strSQL & "else 2 "
    strSQL = strSQL & "end ���b�g��1 "
    strSQL = strSQL & ", "
    strSQL = strSQL & "case "
    strSQL = strSQL & "when right(b.��Spec,4) >= '1006' and b.�F = 'PW' then 0 "
    strSQL = strSQL & "else d.������ "
    strSQL = strSQL & "end ���b�g��2 "
    strSQL = strSQL & ",case "
    strSQL = strSQL & "when b.Style like 'S%' then 0 "
    strSQL = strSQL & "else 1 "
    strSQL = strSQL & "end ���b�g��3 "
    strSQL = strSQL & ",case "
    strSQL = strSQL & "when replace(dbo.fncgethinban(�i��1,��������i��),'�� ','') like '%-[0-9][0-9][0-9][0-9]C%' then 1 "
    strSQL = strSQL & "when replace(dbo.fncgethinban(�i��1,��������i��),'�� ','') like '%-[0-9][0-9][0-9][0-9]S[TSG]%' then 2 "
    strSQL = strSQL & "when replace(dbo.fncgethinban(�i��1,��������i��),'�� ','') like '%-[0-9][0-9][0-9][0-9]G%' then 3 "
    strSQL = strSQL & "when replace(dbo.fncgethinban(�i��1,��������i��),'�� ','') like '%-[0-9][0-9][0-9][0-9]MF%' then 3 "
    strSQL = strSQL & "when replace(dbo.fncgethinban(�i��1,��������i��),'�� ','') like '%-[0-9][0-9][0-9][0-9]D%' then 3 "
    strSQL = strSQL & "else 4 "
    strSQL = strSQL & "end ���b�g��4 "
    strSQL = strSQL & ",case when replace(dbo.fncgethinban(�i��1,��������i��),'�� ','') like '%DA-[0-9][0-9][0-9][0-9]%' then 1 "
    strSQL = strSQL & "when replace(dbo.fncgethinban(�i��1,��������i��),'�� ','') like '%DAS-[0-9][0-9][0-9][0-9]%' then 1 "
    strSQL = strSQL & "when replace(dbo.fncgethinban(�i��1,��������i��),'�� ','') like '%DO-[0-9][0-9][0-9][0-9]%' then 1 "
    strSQL = strSQL & "when replace(dbo.fncgethinban(�i��1,��������i��),'�� ','') like '%DOS-[0-9][0-9][0-9][0-9]%' then 1 "
    strSQL = strSQL & "when replace(dbo.fncgethinban(�i��1,��������i��),'�� ','') like '%DK-[0-9][0-9][0-9][0-9]%' then 1 "
    strSQL = strSQL & "when replace(dbo.fncgethinban(�i��1,��������i��),'�� ','') like '%DKS-[0-9][0-9][0-9][0-9]%' then 1 "
    strSQL = strSQL & "else 2 end ���b�g��5 "
    strSQL = strSQL & ",d.�F���� "
    strSQL = strSQL & ",a.�_��ԍ�,a.���ԍ�,a.�����ԍ� "
    strSQL = strSQL & ",a.�_��ԍ� + '-' + a.���ԍ� + '-' + a.�����ԍ� �_��No "
    strSQL = strSQL & ",������ "
    strSQL = strSQL & ",�{�H�X "
    strSQL = strSQL & ",������ "
    strSQL = strSQL & ",[2���≮] "
    strSQL = strSQL & ",c.�� "
    strSQL = strSQL & ",�ݒu�ꏊ "
    strSQL = strSQL & ",dbo.fncgethinban(�i��1,��������i��) �i��1 "
    strSQL = strSQL & ",dbo.fncgethinban(�q���i��,�����q���i��) �q���i�� "
    strSQL = strSQL & ",���i�� "
    strSQL = strSQL & ",�F "
    strSQL = strSQL & ",�݌� "
    strSQL = strSQL & ",�{�� "
    strSQL = strSQL & ",b.���� "
    strSQL = strSQL & ",���� "
    strSQL = strSQL & ",c.�i�ԋ敪 "
    strSQL = strSQL & ",c.�����敪 "
    strSQL = strSQL & ",b.�H��CD "
    strSQL = strSQL & ",b.�ǉ� "
    strSQL = strSQL & ",�K "
    strSQL = strSQL & ",���q���W "
    strSQL = strSQL & ",���iFLG "
    strSQL = strSQL & ",Null ��� "
    strSQL = strSQL & ",1 �H�� "
    strSQL = strSQL & ",���d�l "
    strSQL = strSQL & ",null �o�ו��@ "
    strSQL = strSQL & ",����݌v���l �݌v���l "
    strSQL = strSQL & ",[08�J�^���O] "
    strSQL = strSQL & ",�W���n���h�� "
    strSQL = strSQL & ",DW "
    strSQL = strSQL & ",DW�q�� "
    strSQL = strSQL & ",DH "
    strSQL = strSQL & ",���葋 "
    strSQL = strSQL & ",case  "
    strSQL = strSQL & " when replace(dbo.fncgethinban(�i��1,��������i��),'�� ','') like 'K_SD%' then FL��g���n�OH "
    strSQL = strSQL & " else FL�d���H "
    strSQL = strSQL & " end CH "
    strSQL = strSQL & ",��Spec "
    strSQL = strSQL & ",a.Spec "
    strSQL = strSQL & ",style �󒍖���style "
    strSQL = strSQL & ",�V�ʍފ��t "
    strSQL = strSQL & ",�����{�H��  "
    strSQL = strSQL & ",dbo.fncNohinHaisoCode(b.�_��ԍ�,b.���ԍ�,b.�����ԍ�,b.��,1) ����o�ו��@ "
    strSQL = strSQL & ",�N���[���p���l "
    strSQL = strSQL & ",����m��� "
    strSQL = strSQL & ",c.������ "
    strSQL = strSQL & "from T_��Ͻ� a "
    strSQL = strSQL & "inner join T_��Ͻ�_2 a2 "
    strSQL = strSQL & "on a.�_��ԍ� = a2.�_��ԍ� and a.���ԍ� = a2.���ԍ� and a.�����ԍ� = a2.�����ԍ� "
    strSQL = strSQL & "inner join T_�󒍖��� b "
    strSQL = strSQL & "on a.�_��ԍ� = b.�_��ԍ� and a.���ԍ� = b.���ԍ� and a.�����ԍ� = b.�����ԍ� "
    strSQL = strSQL & "inner join T_�����w�� c "
    strSQL = strSQL & "on c.�_��ԍ� = b.�_��ԍ� and c.���ԍ� = b.���ԍ� and c.�����ԍ� = b.�����ԍ� and c.�� = b.�� "
    strSQL = strSQL & "left join T_�F�L��Ͻ� d "
    strSQL = strSQL & "on b.�F = d.�F�L�� "
    strSQL = strSQL & "where  c.������ between " & varNullChk(SeizoDate, 2) & " and " & varNullChk(SeizoDate, 2) & " "
    strSQL = strSQL & "and c.�m�� = 2 "
    strSQL = strSQL & "and �����敪 in (1,2,3) "
    strSQL = strSQL & "order by ���b�g��1 "
    strSQL = strSQL & ",���b�g��2"
    strSQL = strSQL & ",���b�g��3"
    strSQL = strSQL & ",���b�g��4"
    strSQL = strSQL & ",���b�g��5"
    strSQL = strSQL & ",�i��1"
    strSQL = strSQL & ",�݌�"
    strSQL = strSQL & ",b.�_��ԍ�,b.���ԍ�,b.�����ԍ�,b.�� "
    
    If Not objLocalDB.ExecSQL("delete * from WK_�����˗������� ", strErrMsg) Then
        Err.Raise 9999, , strErrMsg
    End If
    
    With objREMOTEDB

        If .ExecSelect(strSQL) Then
            If Not .GetRS.EOF Then
            
                intCnt = 1
                
                Do Until .GetRS.EOF
                    
                    intMaisu = 0
                    varBrandHinban = ""
                    
                    If Not IsNull(.GetRS![�i��1]) Then
                        
                        If Not IsSxL(Nz(.GetRS![�i��1], ""), varBrandHinban) Then
                            varBrandHinban = .GetRS![�i��1]
                        End If
                        
                        If Not IsNull(.GetRS![�q���i��]) Then
                            intMaisu = .GetRS![����] / 2
                        Else
                            intMaisu = .GetRS![����]
                        End If
                        
                        If .GetRS![������] Like "*���ˌ��e*" Or .GetRS![2���≮] Like "*�`���l��*" Then
                            ChuuiFlg = True
                        Else
                            ChuuiFlg = False
                        End If
                        
                        strSQL = conSQL
                        
                        If .GetRS![�H��CD] = 10 Then
                            strSQL = Replace(strSQL, "@�ʔ�@", intCnt)
                            intCnt = intCnt + 1
                        Else
                            strSQL = Replace(strSQL, "@�ʔ�@", "Null")
                        End If
                        
                        strSQL = Replace(strSQL, "@�_��ԍ�@", varNullChk(.GetRS![�_��ԍ�], 1))
                        strSQL = Replace(strSQL, "@���ԍ�@", varNullChk(.GetRS![���ԍ�], 1))
                        strSQL = Replace(strSQL, "@�����ԍ�@", varNullChk(.GetRS![�����ԍ�], 1))
                        strSQL = Replace(strSQL, "@�_��No@", varNullChk(.GetRS![�_��No], 1))
                        strSQL = Replace(strSQL, "@������@", varNullChk(.GetRS![������], 1))
                        strSQL = Replace(strSQL, "@�{�H�X@", varNullChk(.GetRS![�{�H�X], 1))
                        strSQL = Replace(strSQL, "@��@", varNullChk(.GetRS![��], 1))
                        strSQL = Replace(strSQL, "@�ݒu�ꏊ@", varNullChk(.GetRS![�ݒu�ꏊ], 1))
                        strSQL = Replace(strSQL, "@�i��1@", varNullChk(.GetRS![�i��1], 1))
                        strSQL = Replace(strSQL, "@���i��@", varNullChk(.GetRS![���i��], 1))
                        strSQL = Replace(strSQL, "@�F@", varNullChk(.GetRS![�F����], 1))
                        strSQL = Replace(strSQL, "@�F�R�[�h@", varNullChk(.GetRS![�F], 1))
                        strSQL = Replace(strSQL, "@�݌�@", varNullChk(.GetRS![�݌�], 1))
                        strSQL = Replace(strSQL, "@�J�l��@", varNullChk(objTateguHinban.�J�l��(varBrandHinban), 1))
                        strSQL = Replace(strSQL, "@�{��@", varNullChk(.GetRS![�{��], 1))
                        strSQL = Replace(strSQL, "@����@", varNullChk(.GetRS![����], 1))
                        strSQL = Replace(strSQL, "@����@", varNullChk(intMaisu, 1))
                        strSQL = Replace(strSQL, "@�i�ԋ敪@", varNullChk(.GetRS![�i�ԋ敪], 1))
                        strSQL = Replace(strSQL, "@�����敪@", varNullChk(.GetRS![�����敪], 1))
                        strSQL = Replace(strSQL, "@�H��CD@", varNullChk(.GetRS![�H��CD], 1))
                        strSQL = Replace(strSQL, "@�ǉ�@", varNullChk(.GetRS![�ǉ�], 1))
                        strSQL = Replace(strSQL, "@�K@", varNullChk(.GetRS![�K], 1))
                        strSQL = Replace(strSQL, "@���q���W@", varNullChk(.GetRS![���q���W], 1))
                        strSQL = Replace(strSQL, "@���iFLG@", varNullChk(.GetRS![���iFLG], 1))
                        strSQL = Replace(strSQL, "@���@", varNullChk(.GetRS![���], 1))
                        strSQL = Replace(strSQL, "@�H��@", varNullChk(.GetRS![�H��], 1))
                        strSQL = Replace(strSQL, "@���d�l@", varNullChk(.GetRS![���d�l], 1))
                        strSQL = Replace(strSQL, "@�o�ו��@@", varNullChk(.GetRS![�o�ו��@], 1))
                        strSQL = Replace(strSQL, "@�݌v���l@", varNullChk(.GetRS![�݌v���l], 1))
                        strSQL = Replace(strSQL, "@08�J�^���O@", varNullChk(.GetRS![08�J�^���O], 1))
                        strSQL = Replace(strSQL, "@�W���n���h��@", varNullChk(.GetRS![�W���n���h��], 1))
                        strSQL = Replace(strSQL, "@DW@", varNullChk(.GetRS![DW], 1))
                        strSQL = Replace(strSQL, "@DH@", varNullChk(.GetRS![DH], 1))
                        strSQL = Replace(strSQL, "@CH@", varNullChk(.GetRS![CH], 1))
                        strSQL = Replace(strSQL, "@���葋@", varNullChk(.GetRS![���葋], 1))
                        strSQL = Replace(strSQL, "@��Spec@", varNullChk(.GetRS![��Spec], 1))
                        strSQL = Replace(strSQL, "@Spec@", varNullChk(.GetRS![Spec], 1))
                        strSQL = Replace(strSQL, "@�󒍖���Style@", varNullChk(.GetRS![�󒍖���Style], 1))
                        strSQL = Replace(strSQL, "@style@", varNullChk(objTateguHinban.Style(varBrandHinban), 1))
                        strSQL = Replace(strSQL, "@�V�ʍފ��t@", varNullChk(.GetRS![�V�ʍފ��t], 1))
                        strSQL = Replace(strSQL, "@����FLG@ ", varNullChk(ChuuiFlg, 1))
                        strSQL = Replace(strSQL, "@�����{�H��@", varNullChk(.GetRS![�����{�H��], 1))
                        strSQL = Replace(strSQL, "@����o�ו��@@", varNullChk(.GetRS![����o�ו��@], 1))
                        strSQL = Replace(strSQL, "@�N���[���p���l@", varNullChk(.GetRS![�N���[���p���l], 1))
                        strSQL = Replace(strSQL, "@����m���@", varNullChk(.GetRS![����m���], 1))
                        strSQL = Replace(strSQL, "@������@", varNullChk(.GetRS![������], 1))
                        strSQL = Replace(strSQL, "@�q��@", False)

                        If Not objLocalDB.ExecSQL(strSQL, strErrMsg) Then
                            Err.Raise 9999, , strErrMsg
                        End If
                    
                    End If
                    
                    If Not IsNull(.GetRS![�q���i��]) Then
                        
                        If Not IsSxL(.GetRS![�q���i��], varBrandHinban) Then
                            varBrandHinban = .GetRS![�q���i��]
                        End If
                        
                        If intMaisu = 0 Then
                            intMaisu = .GetRS![����]
                        End If
                        
                        strSQL = conSQL
                        
                        If .GetRS![�H��CD] = 10 Then
                            strSQL = Replace(strSQL, "@�ʔ�@", intCnt)
                            intCnt = intCnt + 1
                        Else
                            strSQL = Replace(strSQL, "@�ʔ�@", "Null")
                        End If
                        
                        strSQL = Replace(strSQL, "@�_��ԍ�@", varNullChk(.GetRS![�_��ԍ�], 1))
                        strSQL = Replace(strSQL, "@���ԍ�@", varNullChk(.GetRS![���ԍ�], 1))
                        strSQL = Replace(strSQL, "@�����ԍ�@", varNullChk(.GetRS![�����ԍ�], 1))
                        strSQL = Replace(strSQL, "@�_��No@", varNullChk(.GetRS![�_��No], 1))
                        strSQL = Replace(strSQL, "@������@", varNullChk(.GetRS![������], 1))
                        strSQL = Replace(strSQL, "@�{�H�X@", varNullChk(.GetRS![�{�H�X], 1))
                        strSQL = Replace(strSQL, "@��@", varNullChk(.GetRS![��], 1))
                        strSQL = Replace(strSQL, "@�ݒu�ꏊ@", varNullChk(.GetRS![�ݒu�ꏊ], 1))
                        strSQL = Replace(strSQL, "@�i��1@", varNullChk(.GetRS![�q���i��], 1))
                        strSQL = Replace(strSQL, "@���i��@", varNullChk(.GetRS![���i��], 1))
                        strSQL = Replace(strSQL, "@�F@", varNullChk(.GetRS![�F����], 1))
                        strSQL = Replace(strSQL, "@�F�R�[�h@", varNullChk(.GetRS![�F], 1))
                        strSQL = Replace(strSQL, "@�݌�@", varNullChk(.GetRS![�݌�], 1))
                        strSQL = Replace(strSQL, "@�J�l��@", varNullChk(objTateguHinban.�J�l��(varBrandHinban), 1))
                        strSQL = Replace(strSQL, "@�{��@", varNullChk(.GetRS![�{��], 1))
                        strSQL = Replace(strSQL, "@����@", varNullChk(.GetRS![����], 1))
                        strSQL = Replace(strSQL, "@����@", varNullChk(intMaisu, 1))
                        strSQL = Replace(strSQL, "@�i�ԋ敪@", varNullChk(.GetRS![�i�ԋ敪], 1))
                        strSQL = Replace(strSQL, "@�����敪@", varNullChk(.GetRS![�����敪], 1))
                        strSQL = Replace(strSQL, "@�H��CD@", varNullChk(.GetRS![�H��CD], 1))
                        strSQL = Replace(strSQL, "@�ǉ�@", varNullChk(.GetRS![�ǉ�], 1))
                        strSQL = Replace(strSQL, "@�K@", varNullChk(.GetRS![�K], 1))
                        strSQL = Replace(strSQL, "@���q���W@", varNullChk(.GetRS![���q���W], 1))
                        strSQL = Replace(strSQL, "@���iFLG@", varNullChk(.GetRS![���iFLG], 1))
                        strSQL = Replace(strSQL, "@���@", varNullChk(.GetRS![���], 1))
                        strSQL = Replace(strSQL, "@�H��@", varNullChk(.GetRS![�H��], 1))
                        strSQL = Replace(strSQL, "@���d�l@", varNullChk(.GetRS![���d�l], 1))
                        strSQL = Replace(strSQL, "@�o�ו��@@", varNullChk(.GetRS![�o�ו��@], 1))
                        strSQL = Replace(strSQL, "@�݌v���l@", varNullChk(.GetRS![�݌v���l], 1))
                        strSQL = Replace(strSQL, "@08�J�^���O@", varNullChk(.GetRS![08�J�^���O], 1))
                        strSQL = Replace(strSQL, "@�W���n���h��@", varNullChk(.GetRS![�W���n���h��], 1))
                        strSQL = Replace(strSQL, "@DW@", varNullChk(.GetRS![DW�q��], 1))
                        strSQL = Replace(strSQL, "@DH@", varNullChk(.GetRS![DH], 1))
                        strSQL = Replace(strSQL, "@CH@", varNullChk(.GetRS![CH], 1))
                        strSQL = Replace(strSQL, "@���葋@", varNullChk(.GetRS![���葋], 1))
                        strSQL = Replace(strSQL, "@��Spec@", varNullChk(.GetRS![��Spec], 1))
                        strSQL = Replace(strSQL, "@Spec@", varNullChk(.GetRS![Spec], 1))
                        strSQL = Replace(strSQL, "@�󒍖���Style@", varNullChk(.GetRS![�󒍖���Style], 1))
                        strSQL = Replace(strSQL, "@Style@", varNullChk(objTateguHinban.Style(varBrandHinban), 1))
                        strSQL = Replace(strSQL, "@�V�ʍފ��t@", varNullChk(.GetRS![�V�ʍފ��t], 1))
                        strSQL = Replace(strSQL, "@����FLG@ ", varNullChk(ChuuiFlg, 1))
                        strSQL = Replace(strSQL, "@�����{�H��@", varNullChk(.GetRS![�����{�H��], 1))
                        strSQL = Replace(strSQL, "@����o�ו��@@", varNullChk(.GetRS![����o�ו��@], 1))
                        strSQL = Replace(strSQL, "@�N���[���p���l@", varNullChk(.GetRS![�N���[���p���l], 1))
                        strSQL = Replace(strSQL, "@����m���@", varNullChk(.GetRS![����m���], 1))
                        strSQL = Replace(strSQL, "@������@", varNullChk(.GetRS![������], 1))
                        strSQL = Replace(strSQL, "@�q��@", True)
                        
                        If Not objLocalDB.ExecSQL(strSQL, strErrMsg) Then
                            Err.Raise 9999, , strErrMsg
                        End If
                    End If
                    
                    .GetRS.MoveNext
                Loop
            End If
        Else
            Err.Raise 9999, , "�����w���f�[�^������܂���"
        End If
    End With
    
    If DCount("*", "WK_�����˗�������") = 0 Then
        Err.Raise 9999, , "�����f�[�^������܂���"
    End If
    
    If Not ���ޓW�JWK�쐬(SeizoDate) Then
        Err.Raise 9999, , "�����f�[�^�쐬�ُ�I���i���ޓW�J���[�N�쐬�j"
    End If
    
    If Not ���ޓW�JWK�����X�V() Then
        Err.Raise 9999, , "�����f�[�^�쐬�ُ�I���i���ޓW�J�����ǉ��j"
    End If
    
    bolfnc�����w���f�[�^���o = True
    
    GoTo Exit_bolfnc�����w���f�[�^���o
    
Err_bolfnc�����w���f�[�^���o:
    MsgBox Err.Description
    Debug.Print strSQL
    
    'Resume
Exit_bolfnc�����w���f�[�^���o:
    Set objREMOTEDB = Nothing
    Set objLocalDB = Nothing
    Set objTateguHinban = Nothing

    
End Function

Private Function ���ޓW�JWK�쐬(ByVal inSeizoDate As Date) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'�w��̐������̕��ޓW�J�����[�N�t�@�C���ɍ쐬����
'
'
'   :�߂�l
'       True            :����
'       False           :���s

'--------------------------------------------------------------------------------------------------------------------
    Dim objREMOTEDB As New cls_BRAND_MASTER
    
    Dim strSQL As String
    
    On Error GoTo Err_���ޓW�JWK�쐬
    
    strSQL = ""
    strSQL = strSQL & "select b.�_��ԍ� "
    strSQL = strSQL & ",b.���ԍ� "
    strSQL = strSQL & ",b.�����ԍ� "
    strSQL = strSQL & ",b.�� "
    strSQL = strSQL & ",�@�� "
    strSQL = strSQL & ",�i�� "
    strSQL = strSQL & ",�����i�� "
    strSQL = strSQL & ",�i�� "
    strSQL = strSQL & ",���iCD "
    strSQL = strSQL & ",���ގ��CD "
    strSQL = strSQL & ",���ޖ� "
    strSQL = strSQL & ",���� "
    strSQL = strSQL & ",�����i "
    strSQL = strSQL & ",��t "
    strSQL = strSQL & ",b.���� "
    strSQL = strSQL & ",���ސ� "
    strSQL = strSQL & ",���ސ����v "
    strSQL = strSQL & ",���� "
    strSQL = strSQL & ",b.�i�ԋ敪 "
    strSQL = strSQL & ",b.�����敪 "
    strSQL = strSQL & ",�ǉ� "
    strSQL = strSQL & ",�L�����Z�� "
    strSQL = strSQL & ",�F "
    strSQL = strSQL & ",�P�� "
    strSQL = strSQL & ",�N���[�� "
    strSQL = strSQL & ",���[�J�[CD "
    strSQL = strSQL & ",�����F "
    strSQL = strSQL & ",���� "
    strSQL = strSQL & ",[PC��] "
    strSQL = strSQL & ",[No] "
    strSQL = strSQL & ",Null as ���� "
    strSQL = strSQL & ",case when dbo.IsKotobira(b.�i��) = 1 then 1 else 0 end �q�� "
    strSQL = strSQL & "from T_�����w�� a "
    strSQL = strSQL & "inner join BRAND_BOM.dbo.T_���ޓW�J b "
    strSQL = strSQL & "on a.�_��ԍ� = b.�_��ԍ� and a.���ԍ� = b.���ԍ� and a.�����ԍ� = b.�����ԍ� and a.�� = b.�� "
    strSQL = strSQL & "where a.������ between " & varNullChk(inSeizoDate, 2) & " and " & varNullChk(inSeizoDate, 2) & " "
    strSQL = strSQL & "and a.�m�� = 2 "
    strSQL = strSQL & "and a.�����敪 between 1 and 3 "
    strSQL = strSQL & "and b.�����敪 in (1,2,3) "
    
    With objREMOTEDB
    
        If .ExecSelect(strSQL) Then
            If Not bolfncTableCopyToLocal(.GetRS, "WK_���ޓW�J") Then
                Err.Raise 9999, , "���ޓW�J�R�s�[�ُ�I�� "
            End If
        Else
            Err.Raise 9999, , "SQL���s�G���[ SQL = " & strSQL
        End If
    End With
    
    ���ޓW�JWK�쐬 = True
    
        
    GoTo Exit_���ޓW�JWK�쐬

Err_���ޓW�JWK�쐬:
    MsgBox Err.Description

Exit_���ޓW�JWK�쐬:
    Set objREMOTEDB = Nothing
    
End Function

Private Function ���ޓW�JWK�����X�V() As Boolean
'--------------------------------------------------------------------------------------------------------------------
'���ޓW�J�����[�N�t�@�C���Ɏ��ނ�DB��蕄���R�[�h���X�V����
'
'
'   :�߂�l
'       True            :����
'       False           :���s

'--------------------------------------------------------------------------------------------------------------------
    Dim objLocalDB As New cls_LOCALDB
    Dim objSKAMIYADB As New cls_SKAMIYADB
    
    Dim strSQL As String
    
    On Error GoTo Err_���ޓW�JWK�����X�V
    
    strSQL = ""
    strSQL = strSQL & "select ���iCD,���� "
    strSQL = strSQL & "from WK_���ޓW�J "
    strSQL = strSQL & "where ���iCD is not null "
    
    
    With objLocalDB
        If .ExecSelect_Writable(strSQL) Then
            
            If Not .GetRS.EOF Then
                Do Until .GetRS.EOF
                    strSQL = ""
                    strSQL = strSQL & "select ���� from BRAND.T_KANAMONO_FUGO "
                    strSQL = strSQL & "where ����CD = 'B' and ���iCD = '" & .GetRS![���iCD] & "' "
                    
                    If objSKAMIYADB.ExecSelect(strSQL) Then
                        If Not objSKAMIYADB.GetRS.EOF Then
                            .GetRS![����] = objSKAMIYADB.GetRS![����]
                            .GetRS.Update
                        End If
                    Else
                        Err.Raise 9999, , "���ޓW�JWK�����X�V�G���[ SQL = " & strSQL
                    End If
                    
                    objSKAMIYADB.RecordSetClose
                
                    .GetRS.MoveNext
                Loop
                
            End If
        Else
            Err.Raise 9999, , "���ޓW�JWK�����X�V SQL���s�G���[ SQL = " & strSQL
        End If
    End With
    
    ���ޓW�JWK�����X�V = True
        
        
    GoTo Exit_���ޓW�JWK�����X�V

Err_���ޓW�JWK�����X�V:
    MsgBox Err.Description

Exit_���ޓW�JWK�����X�V:
    Set objSKAMIYADB = Nothing
    Set objLocalDB = Nothing
    
End Function

Public Function bolfnc�����w��_�t���n�C�g() As Boolean
'--------------------------------------------------------------------------------------------------------------------
'�����̐����w���f�[�^����t���n�C�g���C���p�̎w�������o�͂���
'
'
'   :�߂�l
'       True            :����
'       False           :���s

'--------------------------------------------------------------------------------------------------------------------
    Dim objREMOTEDB As cls_BRAND_MASTER
    Dim objLocalDB As New cls_LOCALDB
    Dim objTateguHinban As New cls_����i��
    Dim objFullHeight As New Cls_FullHeight
    
    Dim rsADO As New ADODB.Recordset
    
    Dim strSQL As String
    Dim conSQL As String
    Dim conSQLW As String
    
    Dim strErrMsg As String
    
    Dim inRecordCount As Long
    Dim inReadCount As Long
    Dim intCnt As Integer
    
    Dim strBrandHinban As String
    Dim strOEMHinban As String
    
    Dim strTateguTsurimoto() As String
    Dim intTateguShurui As Integer
    Dim intOyako As Integer
    
    On Error GoTo Err_bolfnc�����w��_�t���n�C�g
    
    conSQL = conSQL & "insert into WK_�����˗�������_FullHeight "
    conSQL = conSQL & "( "
    conSQL = conSQL & "�ʔ� "
    conSQL = conSQL & ",�_��ԍ� "
    conSQL = conSQL & ",���ԍ� "
    conSQL = conSQL & ",�����ԍ� "
    conSQL = conSQL & ",�_��No "
    conSQL = conSQL & ",������ "
    conSQL = conSQL & ",�{�H�X "
    conSQL = conSQL & ",�� "
    conSQL = conSQL & ",���ʒu "
    conSQL = conSQL & ",�ݒu�ꏊ "
    conSQL = conSQL & ",�i��1 "
    conSQL = conSQL & ",�J�l�� "
    conSQL = conSQL & ",���i�� "
    conSQL = conSQL & ",�F "
    conSQL = conSQL & ",�F�R�[�h "
    conSQL = conSQL & ",�݌� "
    conSQL = conSQL & ",�{�� "
    conSQL = conSQL & ",���� "
    conSQL = conSQL & ",���� "
    conSQL = conSQL & ",�󒍖��� "
    conSQL = conSQL & ",�i�ԋ敪 "
    conSQL = conSQL & ",�����敪 "
    conSQL = conSQL & ",�H��CD "
    conSQL = conSQL & ",�ǉ� "
    conSQL = conSQL & ",�K "
    conSQL = conSQL & ",���q���W "
    conSQL = conSQL & ",���iFLG "
    conSQL = conSQL & ",��� "
    conSQL = conSQL & ",�H�� "
    conSQL = conSQL & ",���d�l "
    conSQL = conSQL & ",�o�ו��@ "
    conSQL = conSQL & ",�݌v���l "
    conSQL = conSQL & ",08�J�^���O "
    conSQL = conSQL & ",�W���n���h�� "
    conSQL = conSQL & ",DW "
    conSQL = conSQL & ",DH "
    conSQL = conSQL & ",CH "
    conSQL = conSQL & ",���葋 "
    conSQL = conSQL & ",��Spec "
    conSQL = conSQL & ",Spec "
    conSQL = conSQL & ",Style "
    conSQL = conSQL & ",�V�ʍފ��t "
    conSQL = conSQL & ",����FLG "
    conSQL = conSQL & ",�����{�H�� "
    conSQL = conSQL & ",����o�ו��@ "
    conSQL = conSQL & ",�N���[���p���l "
    conSQL = conSQL & ",����m��� "
    conSQL = conSQL & ",������ "
    
    conSQL = conSQL & ") values ("
    conSQL = conSQL & "@�ʔ�@ "
    conSQL = conSQL & ",@�_��ԍ�@ "
    conSQL = conSQL & ",@���ԍ�@ "
    conSQL = conSQL & ",@�����ԍ�@ "
    conSQL = conSQL & ",@�_��No@ "
    conSQL = conSQL & ",@������@ "
    conSQL = conSQL & ",@�{�H�X@ "
    conSQL = conSQL & ",@��@ "
    conSQL = conSQL & ",@���ʒu@ "
    conSQL = conSQL & ",@�ݒu�ꏊ@ "
    conSQL = conSQL & ",@�i��1@ "
    conSQL = conSQL & ",@�J�l��@ "
    conSQL = conSQL & ",@���i��@ "
    conSQL = conSQL & ",@�F@ "
    conSQL = conSQL & ",@�F�R�[�h@ "
    conSQL = conSQL & ",@�݌�@ "
    conSQL = conSQL & ",@�{��@ "
    conSQL = conSQL & ",@����@ "
    conSQL = conSQL & ",@����@ "
    conSQL = conSQL & ",@�󒍖���@ "
    conSQL = conSQL & ",@�i�ԋ敪@ "
    conSQL = conSQL & ",@�����敪@ "
    conSQL = conSQL & ",@�H��CD@ "
    conSQL = conSQL & ",@�ǉ�@ "
    conSQL = conSQL & ",@�K@ "
    conSQL = conSQL & ",@���q���W@ "
    conSQL = conSQL & ",@���iFLG@ "
    conSQL = conSQL & ",@���@ "
    conSQL = conSQL & ",@�H��@ "
    conSQL = conSQL & ",@���d�l@ "
    conSQL = conSQL & ",@�o�ו��@@ "
    conSQL = conSQL & ",@�݌v���l@ "
    conSQL = conSQL & ",@08�J�^���O@ "
    conSQL = conSQL & ",@�W���n���h��@ "
    conSQL = conSQL & ",@DW@ "
    conSQL = conSQL & ",@DH@ "
    conSQL = conSQL & ",@CH@ "
    conSQL = conSQL & ",@���葋@ "
    conSQL = conSQL & ",@��Spec@ "
    conSQL = conSQL & ",@Spec@ "
    conSQL = conSQL & ",@Style@ "
    conSQL = conSQL & ",@�V�ʍފ��t@ "
    conSQL = conSQL & ",@����FLG@ "
    conSQL = conSQL & ",@�����{�H��@ "
    conSQL = conSQL & ",@����o�ו��@@ "
    conSQL = conSQL & ",@�N���[���p���l@ "
    conSQL = conSQL & ",@����m���@ "
    conSQL = conSQL & ",@������@ "
    conSQL = conSQL & ") "
    
    strSQL = "select * from WK_�����˗������� "
        
    conSQLW = conSQLW & "where �H��CD = 10 "
    conSQLW = conSQLW & "order by �ʔ� "
        
    
    With objLocalDB
        If Not .ExecSQL("delete * from WK_�����˗�������_Fullheight") Then Err.Raise 9999, , "��������X�g_���[�J���폜�G���["
        
        .CursorLocation = adUseClient
        
        strSQL = strSQL & conSQLW
        If Not .ExecSelect(strSQL) Then Err.Raise 9999, , "WK_�����˗������� ���o�G���["
        
        inRecordCount = .GetRS.RecordCount
        intCnt = 1
        
        If Not .GetRS.EOF Then
            Do Until .GetRS.EOF
            
                If .GetRS![�q��] Then
                    intOyako = 2
                Else
                    intOyako = 1
                End If
                
                Set rsADO = objFullHeight.Rs������(.GetRS![�_��ԍ�], .GetRS![���ԍ�], .GetRS![�����ԍ�], .GetRS![��], intOyako)
                
                If Not rsADO.EOF Then
                
                    Do Until rsADO.EOF
                        
                        intTateguShurui = rsADO![������]
                        
                        strSQL = conSQL

                        strSQL = Replace(strSQL, "@�ʔ�@", intCnt)
                        intCnt = intCnt + 1
                        
                        strSQL = Replace(strSQL, "@�_��ԍ�@", varNullChk(.GetRS![�_��ԍ�], 1))
                        strSQL = Replace(strSQL, "@���ԍ�@", varNullChk(.GetRS![���ԍ�], 1))
                        strSQL = Replace(strSQL, "@�����ԍ�@", varNullChk(.GetRS![�����ԍ�], 1))
                        strSQL = Replace(strSQL, "@�_��No@", varNullChk(.GetRS![�_��No], 1))
                        strSQL = Replace(strSQL, "@������@", varNullChk(.GetRS![������], 1))
                        strSQL = Replace(strSQL, "@�{�H�X@", varNullChk(.GetRS![�{�H�X], 1))
                        strSQL = Replace(strSQL, "@��@", varNullChk(.GetRS![��], 1))
                        strSQL = Replace(strSQL, "@���ʒu@", varNullChk(intTateguShurui, 1))
                        strSQL = Replace(strSQL, "@�ݒu�ꏊ@", varNullChk(.GetRS![�ݒu�ꏊ], 1))
                        strSQL = Replace(strSQL, "@�i��1@", varNullChk(.GetRS![�i��1], 1))
                        strSQL = Replace(strSQL, "@���i��@", varNullChk(.GetRS![���i��], 1))
                        strSQL = Replace(strSQL, "@�F@", varNullChk(.GetRS![�F], 1))
                        strSQL = Replace(strSQL, "@�F�R�[�h@", varNullChk(.GetRS![�F�R�[�h], 1))
                        strSQL = Replace(strSQL, "@�݌�@", varNullChk(.GetRS![�݌�], 1))
                        strSQL = Replace(strSQL, "@�{��@", varNullChk(.GetRS![�{��], 1))
                        strSQL = Replace(strSQL, "@����@", varNullChk(.GetRS![����], 1))
                        strSQL = Replace(strSQL, "@����@", varNullChk(.GetRS![����], 1))
                        strSQL = Replace(strSQL, "@�󒍖���@", varNullChk(.GetRS![����], 1))
                        strSQL = Replace(strSQL, "@�i�ԋ敪@", varNullChk(.GetRS![�i�ԋ敪], 1))
                        strSQL = Replace(strSQL, "@�����敪@", varNullChk(.GetRS![�����敪], 1))
                        strSQL = Replace(strSQL, "@�H��CD@", varNullChk(.GetRS![�H��CD], 1))
                        strSQL = Replace(strSQL, "@�ǉ�@", varNullChk(.GetRS![�ǉ�], 1))
                        strSQL = Replace(strSQL, "@�K@", varNullChk(.GetRS![�K], 1))
                        strSQL = Replace(strSQL, "@���q���W@", varNullChk(.GetRS![���q���W], 1))
                        strSQL = Replace(strSQL, "@���iFLG@", varNullChk(.GetRS![���iFLG], 1))
                        strSQL = Replace(strSQL, "@���@", varNullChk(.GetRS![���], 1))
                        strSQL = Replace(strSQL, "@�H��@", varNullChk(.GetRS![�H��], 1))
                        strSQL = Replace(strSQL, "@���d�l@", varNullChk(.GetRS![���d�l], 1))
                        strSQL = Replace(strSQL, "@�o�ו��@@", varNullChk(.GetRS![�o�ו��@], 1))
                        strSQL = Replace(strSQL, "@�݌v���l@", varNullChk(.GetRS![�݌v���l], 1))
                        strSQL = Replace(strSQL, "@08�J�^���O@", varNullChk(.GetRS![08�J�^���O], 1))
                        strSQL = Replace(strSQL, "@�W���n���h��@", varNullChk(.GetRS![�W���n���h��], 1))
                        strSQL = Replace(strSQL, "@DW@", varNullChk(.GetRS![DW], 1))
                        strSQL = Replace(strSQL, "@DH@", varNullChk(.GetRS![DH], 1))
                        strSQL = Replace(strSQL, "@CH@", varNullChk(.GetRS![CH], 1))
                        strSQL = Replace(strSQL, "@���葋@", varNullChk(.GetRS![���葋], 1))
                        strSQL = Replace(strSQL, "@��Spec@", varNullChk(.GetRS![��Spec], 1))
                        strSQL = Replace(strSQL, "@Spec@", varNullChk(.GetRS![Spec], 1))
                        strSQL = Replace(strSQL, "@Style@", varNullChk(.GetRS![Style], 1))
                        strSQL = Replace(strSQL, "@Style@", varNullChk(objTateguHinban.Style(.GetRS![�i��1]), 1))
                        strSQL = Replace(strSQL, "@�J�l��@", varNullChk(.GetRS![�J�l��], 1))
                        strSQL = Replace(strSQL, "@�V�ʍފ��t@", varNullChk(.GetRS![�V�ʍފ��t], 1))
                        strSQL = Replace(strSQL, "@����FLG@", varNullChk(.GetRS![����FLG], 1))
                        strSQL = Replace(strSQL, "@�����{�H��@", varNullChk(.GetRS![�����{�H��], 1))
                        strSQL = Replace(strSQL, "@����o�ו��@@", varNullChk(.GetRS![����o�ו��@], 1))
                        strSQL = Replace(strSQL, "@�N���[���p���l@", varNullChk(.GetRS![�N���[���p���l], 1))
                        strSQL = Replace(strSQL, "@����m���@", varNullChk(.GetRS![����m���], 1))
                        strSQL = Replace(strSQL, "@������@", varNullChk(.GetRS![������], 1))
                        
                        If Not objLocalDB.ExecSQL(strSQL, strErrMsg) Then
                            Err.Raise 9999, , strErrMsg
                        End If
                
                        inReadCount = inReadCount + 1

                        
                        SysCmd acSysCmdSetStatus, "���s��.... " & inReadCount & "/" & inRecordCount
                        If inReadCount Mod 10 = 0 Then
                            DoEvents
                        End If
                    
                        rsADO.MoveNext
                        
                    Loop
                    
                    If rsADO.State = adStateOpen Then
                        rsADO.Close
                    End If
                        
                Else
                    Err.Raise 9999, , "�t���n�C�g���C���p�̉��H�f�[�^�i�݌v�j������܂���B�_��ԍ� = " & .GetRS![�_��No] & " ,��No." & .GetRS![��]
                End If
            
                .GetRS.MoveNext
            Loop
        Else
            Err.Raise 9999, , "�����w���f�[�^������܂���"
        End If
        
    End With
    
    GoTo Exit_bolfnc�����w��_�t���n�C�g
    
Err_bolfnc�����w��_�t���n�C�g:
    MsgBox Err.Description
    
Exit_bolfnc�����w��_�t���n�C�g:
    
    Set objREMOTEDB = Nothing
    Set objLocalDB = Nothing
    Set objTateguHinban = Nothing
    Set objFullHeight = Nothing
    Set rsADO = Nothing
    
    SysCmd acSysCmdSetStatus, " "
    
End Function

Public Function bolfnc�����w���t���n�C�g���[�f�[�^() As Boolean

Dim objLocalDB As New cls_LOCALDB
    Dim objTateguKansu As New cls_������֐�
    Dim objCheckLabel As New cls_����ʃ��x��
    
    Dim strSQL As String
    Dim strErrMsg As String
    
    Dim inRecordCount As Long
    Dim inReadCount As Long
    
    Dim strBrandHinban As String
    Dim strOEMHinban As String
    
    Dim intHanWari As Integer
    Dim intTateguStyle As Integer
    Dim intTateguKeijo As Integer
    Dim intOyako As Integer
    Dim strTsurimoto As String
    Dim strLineKbn As String
    
    Dim strQRLabel As String
    
    On Error GoTo Err_bolfnc�����w���t���n�C�g���[�f�[�^
    
    SysCmd acSysCmdSetStatus, " "
    
    strSQL = "select * from WK_�����˗�������_FullHeight "
    strSQL = strSQL & "order by �ʔ� "
    
    With objLocalDB
        If Not .ExecSQL("delete * from WK_��������X�g_FullHeight") Then Err.Raise 9999, , "��������X�g_���[�J���폜�G���["
        
        .CursorLocation = adUseClient
        
        If Not .ExecSelect(strSQL) Then Err.Raise 9999, , "Input�t�@�C�����o�G���["
        
        inRecordCount = .GetRS.RecordCount
        
        If Not .GetRS.EOF Then
            SysCmd acSysCmdClearStatus
            Do Until .GetRS.EOF
                
                If objTateguKansu.Bind(.GetRS![�_��No], .GetRS![�i��1], .GetRS![����], .GetRS![�󒍖���], .GetRS![DW], Nz(.GetRS![DH], 0), .GetRS![CH], .GetRS![�{��], Nz(.GetRS![���葋], "-"), .GetRS![��Spec]) Then
                    
                    If Nz(objTateguKansu.W1, 0) > 0 And Nz(objTateguKansu.W2, 0) Then
                        intHanWari = 1
                        
                        If Replace(objTateguKansu.�u�����h�i��, "�� ", "") Like "??B*" Then
                        
                            intTateguKeijo = 2
                        Else
                            intTateguKeijo = 1
                        End If
                        
                    Else
                        intHanWari = 0
                        intTateguKeijo = 0
                    End If
                    
                    If IsHirakido(objTateguKansu.�u�����h�i��) Or IsOyatobira(objTateguKansu.�u�����h�i��) Or IsKotobira(objTateguKansu.�u�����h�i��) Then
                        intTateguStyle = 0
                    ElseIf IsHikido(objTateguKansu.�u�����h�i��) Then
                        intTateguStyle = 1
                    ElseIf IsCloset_Hikichigai(objTateguKansu.�u�����h�i��) Then
                        intTateguStyle = 1
                    ElseIf IsCloset_Slide(objTateguKansu.�u�����h�i��) Then
                        intTateguStyle = 3
                    Else
                        Err.Raise 9999, , "�J�l���G���[ �i�� = " & objTateguKansu.�u�����h�i��
                    End If
                    
                    If IsKotobira(objTateguKansu.�u�����h�i��) Then
                        intOyako = 2
                        
                        If Nz(.GetRS![�݌�], "") = "R" Then
                            strTsurimoto = "L"
                        ElseIf Nz(.GetRS![�݌�], "") = "L" Then
                            strTsurimoto = "R"
                        Else
                           Err.Raise 9999, , "�q���݌��G���[ �_��No = " & .GetRS![�_��No] & " �� = " & .GetRS![��]
                        End If
                    Else
                        intOyako = 1
                        strTsurimoto = Nz(.GetRS![�݌�], "Z")
                    End If
                    
                    strQRLabel = ""
                    strQRLabel = strQRLabel & RPAD(StrConv(.GetRS![�_��No], vbNarrow), " ", 20)
                    strQRLabel = strQRLabel & LPAD(.GetRS![��], "0", 3)
                    strQRLabel = strQRLabel & intOyako
                    If Nz(.GetRS![���ʒu], 0) > 3 Then
                        strQRLabel = strQRLabel & "03"
                    Else
                        strQRLabel = strQRLabel & LPAD(.GetRS![���ʒu], "0", 2)
                    End If
                    strQRLabel = strQRLabel & LPAD(CStr(CInt(Nz(objTateguKansu.����, 0) * 10)), "0", 3)
                    strQRLabel = strQRLabel & CStr(intHanWari)
                    
                    If intHanWari = 1 Then
                        strQRLabel = strQRLabel & "00000"
                    Else
                        strQRLabel = strQRLabel & LPAD(CStr(CInt(Nz(objTateguKansu.�e�m�iW, 0) * 10)), "0", 5)
                    End If
                    
                    If intHanWari = 0 Then
                        strQRLabel = strQRLabel & "0000"
                    Else
                        strQRLabel = strQRLabel & LPAD(CStr(CInt(Nz(objTateguKansu.W1, 0) * 10)), "0", 4)
                    End If
                    
                    If intHanWari = 0 Then
                        strQRLabel = strQRLabel & "00000"
                    Else
                        strQRLabel = strQRLabel & LPAD(CStr(CInt(Nz(objTateguKansu.W2, 0) * 10)), "0", 5)
                    End If
                    
                    strQRLabel = strQRLabel & LPAD(CStr(CInt(Nz(objTateguKansu.�e�m�iH, 0) * 10)), "0", 5)
                    
                    strQRLabel = strQRLabel & LPAD(CStr(CInt(Nz(objTateguKansu.�c��, 0) * 10)), "0", 3)
                    
                    strQRLabel = strQRLabel & LPAD(CStr(CInt(Nz(objTateguKansu.�\�ʍތ���, 0) * 10)), "0", 2)
                    
                    strQRLabel = strQRLabel & LPAD(.GetRS![�ʔ�], "0", 4)
                    
                    strOEMHinban = objTateguKansu.OEM�i��

                    If IsNull(.GetRS![�����敪]) Then
                        strLineKbn = "S"
                        
                    ElseIf .GetRS![�����敪] = 2 Or .GetRS![�����敪] = 3 Then
                        strLineKbn = "K"
                    
                    ElseIf IsGikan(objTateguKansu.�u�����h�i��) Then
                        '����������Ⴂ�i�K���X�j--���ʂ̏ꍇ�͓������C���A�Жʂ̏ꍇ�̓t���b�V�����C��
                        If IsCloset_Hikichigai(objTateguKansu.�u�����h�i��) Then
                        
                            If fncIntHalfGlassMirror_Maisu(objTateguKansu.�u�����h�i��, 2) = 2 Then
                                strLineKbn = "T"
                            Else
                                strLineKbn = "F"
                            End If
                        Else
                            strLineKbn = "T"
                        End If
                    Else
                        strLineKbn = "F"
                    End If
                        
                    strSQL = ""
                    strSQL = strSQL & "insert into WK_��������X�g_FullHeight( "
                    strSQL = strSQL & " �ʔ�,�_��No "
                    strSQL = strSQL & ",�_��ԍ�,���ԍ�,�����ԍ� "
                    strSQL = strSQL & ", ������, �ݒu�ꏊ, �{�H�X, ��, �u�����h�i��"
                    strSQL = strSQL & ", OEM�i��, �{��, �F�R�[�h, �F, ����"
                    strSQL = strSQL & ", ����, �󒍖���, �݌�, ���ʒu, �J�l��, ���葋, DW, DH, CH"
                    strSQL = strSQL & ", ������, ����X�^�C��, ����`��, �e�q"
                    strSQL = strSQL & ", ��Spec, ��{�}, �c�g�ڍא}, �K���X��"
                    strSQL = strSQL & ", �K���X���, �K���X����, �K���X���2"
                    strSQL = strSQL & ", �K���X����2, �e�m�iW, �e�m�iW2"
                    strSQL = strSQL & ", �e�m�iW����, �e�m�iW����2, W1, W2"
                    strSQL = strSQL & ", �K���XW, �e�m�iH, �e�m�iH2, �K���XH"
                    strSQL = strSQL & ", �e�m�iH����, �e�m�iH����2, H1, H2, H3"
                    strSQL = strSQL & ", �K���X��1, �K���XW1, �K���XH1, �c��"
                    strSQL = strSQL & ", ����, ������, �\�ʍތ���, �\�ʍތ���2"
                    strSQL = strSQL & ", CU_AS_AW�\�ʍ�W, CU_AS_AW�\�ʍ�W2, CU_AS_AW�\�ʍ�H"
                    strSQL = strSQL & ", CU_AS_AW�\�ʍ�H2, CU_AS_AW�\�ʍޖ���, CU_AS_AW�\�ʍޖ���2"
                    strSQL = strSQL & ", AS_AW����W, AS_AW����W1, AS_AW����W2, AS_AW����H, AS_AW����H1"
                    strSQL = strSQL & ", AS_AW����H2, AS_AW������, AS_AW������1, AS_AW������2"
                    strSQL = strSQL & ", ��p�l��W, ��p�l��H, ��p�l������, ��HFLG, ���p�l��H"
                    strSQL = strSQL & ", ���p�l������, ��HFLG, ���p�l��H, ���p�l������, ��HFLG"
                    strSQL = strSQL & ", �~���[�\H, �~���[�\����, �~���[��H, �~���[������"
                    strSQL = strSQL & ", ���^��W, ���^������, ��y���, ��y����, ��y�{��"
                    strSQL = strSQL & ", ���y���, ���y����, ���y�{��, �A�N�Z���g���C�� "
                    strSQL = strSQL & ", �_�{�s�b�`1, �_�{�s�b�`2, �_�{�s�b�`3, �_�{�s�b�`4, �_�{�s�b�`5 "
                    strSQL = strSQL & ", �_�{�s�b�`6, �_�{�s�b�`7, �_�{�s�b�`8, �_�{�s�b�`9, �_�{�s�b�`10 "
                    strSQL = strSQL & ", ���Ԉʒu��, ���Ԉʒu����, ���Ԉʒu����, ���Ԉʒu�� "
                    strSQL = strSQL & ", �a "
                    strSQL = strSQL & ", �n���h������Z���^�[, �n���h������BS "
                    strSQL = strSQL & ", �����Z���^�[, ����BS "
                    strSQL = strSQL & ", �������Z���^�[ "
                    strSQL = strSQL & ", �q�����b�`�󂯃Z���^�[, �q�������󂯃Z���^�[ "
                    strSQL = strSQL & ", �����O, ������ "
                    strSQL = strSQL & ", �J�Z�b�g���H�}, �ˊJ���} "
                    strSQL = strSQL & ", ���iFLG, QR�R�[�h, QR�R�[�h_2, Spec,�݌v���l "
                    strSQL = strSQL & ", �V�ʍފ��t "
                    strSQL = strSQL & ", ����FLG "
                    strSQL = strSQL & ", �����{�H��  "
                    strSQL = strSQL & ", ����o�ו��@ "
                    strSQL = strSQL & ", �N���[���p���l "
                    strSQL = strSQL & ", ����m��� "
                    strSQL = strSQL & ", 08�J�^���O "
                    strSQL = strSQL & ", �����敪 "
                    strSQL = strSQL & ", ���C���敪 "
                    strSQL = strSQL & ", ������ "
                    strSQL = strSQL & " ) "
                    strSQL = strSQL & "values "
                    strSQL = strSQL & "( "
                    strSQL = strSQL & varNullChk(.GetRS![�ʔ�], 1) & "," & varNullChk(.GetRS![�_��No], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![�_��ԍ�], 1) & "," & varNullChk(.GetRS![���ԍ�], 1) & "," & varNullChk(.GetRS![�����ԍ�], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![������], 1) & "," & varNullChk(.GetRS![�ݒu�ꏊ], 1) & "," & varNullChk(.GetRS![�{�H�X], 1) & "," & varNullChk(.GetRS![��], 1) & "," & varNullChk(objTateguKansu.�u�����h�i��, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.OEM�i��, 1) & "," & varNullChk(.GetRS![�{��], 1) & "," & varNullChk(.GetRS![�F�R�[�h], 1) & "," & varNullChk(.GetRS![�F], 1) & "," & varNullChk(.GetRS![����], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![����], 1) & "," & varNullChk(.GetRS![�󒍖���], 1) & "," & varNullChk(strTsurimoto, 1) & "," & varNullChk(.GetRS![���ʒu], 1) & "," & varNullChk(.GetRS![�J�l��], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![���葋], 1) & "," & varNullChk(.GetRS![DW], 1) & "," & varNullChk(.GetRS![DH], 1) & "," & varNullChk(.GetRS![CH], 1)
                    strSQL = strSQL & "," & varNullChk(intHanWari, 1) & "," & varNullChk(intTateguStyle, 1) & "," & varNullChk(intTateguKeijo, 1) & "," & varNullChk(intOyako, 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![��Spec], 1) & "," & varNullChk(objTateguKansu.��{�}, 1) & "," & varNullChk(objTateguKansu.�c�g�ڍא}, 1) & "," & varNullChk(objTateguKansu.�K���X��, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.�K���X��ޖ���("���"), 1) & "," & varNullChk(objTateguKansu.�K���X��ޖ���("����"), 1) & "," & varNullChk(objTateguKansu.�K���X��ޖ���2("���"), 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.�K���X��ޖ���2("����"), 1) & "," & varNullChk(objTateguKansu.�e�m�iW, 1) & "," & varNullChk(objTateguKansu.�e�m�iW2, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.�e�m�iW����, 1) & "," & varNullChk(objTateguKansu.�e�m�iW����2, 1) & "," & varNullChk(objTateguKansu.W1, 1) & "," & varNullChk(objTateguKansu.W2, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.�K���XW, 1) & "," & varNullChk(objTateguKansu.�e�m�iH, 1) & "," & varNullChk(objTateguKansu.�e�m�iH2, 1) & "," & varNullChk(objTateguKansu.�K���XH, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.�e�m�iH����, 1) & "," & varNullChk(objTateguKansu.�e�m�iH����2, 1) & "," & varNullChk(objTateguKansu.H1, 1) & "," & varNullChk(objTateguKansu.H2, 1) & "," & varNullChk(objTateguKansu.H3, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.�K���X��1, 1) & "," & varNullChk(objTateguKansu.�K���XW1, 1) & "," & varNullChk(objTateguKansu.�K���XH1, 1) & "," & varNullChk(objTateguKansu.�c��, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.����, 1) & "," & varNullChk(objTateguKansu.������, 1) & "," & varNullChk(objTateguKansu.�\�ʍތ���, 1) & "," & varNullChk(objTateguKansu.�\�ʍތ���2, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.CU_AS_AW�\�ʍ�W, 1) & "," & varNullChk(objTateguKansu.CU_AS_AW�\�ʍ�W2, 1) & "," & varNullChk(objTateguKansu.CU_AS_AW�\�ʍ�H, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.CU_AS_AW�\�ʍ�H2, 1) & "," & varNullChk(objTateguKansu.CU_AS_AW�\�ʍޖ���, 1) & "," & varNullChk(objTateguKansu.CU_AS_AW�\�ʍޖ���2, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.AS_AW����W, 1) & "," & varNullChk(objTateguKansu.AS_AW����W1, 1) & "," & varNullChk(objTateguKansu.AS_AW����W2, 1) & "," & varNullChk(objTateguKansu.AS_AW����H, 1) & "," & varNullChk(objTateguKansu.AS_AW����H1, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.AS_AW����H2, 1) & "," & varNullChk(objTateguKansu.AS_AW������, 1) & "," & varNullChk(objTateguKansu.AS_AW������1, 1) & "," & varNullChk(objTateguKansu.AS_AW������2, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.��p�l��W, 1) & "," & varNullChk(objTateguKansu.��p�l��H, 1) & "," & varNullChk(objTateguKansu.��p�l������, 1) & "," & varNullChk(objTateguKansu.��HFLG, 1) & "," & varNullChk(objTateguKansu.���p�l��H, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.���p�l������, 1) & "," & varNullChk(objTateguKansu.��HFLG, 1) & "," & varNullChk(objTateguKansu.���p�l��H, 1) & "," & varNullChk(objTateguKansu.���p�l������, 1) & "," & varNullChk(objTateguKansu.��HFLG, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.�~���[�\H, 1) & "," & varNullChk(objTateguKansu.�~���[�\����, 1) & "," & varNullChk(objTateguKansu.�~���[��H, 1) & "," & varNullChk(objTateguKansu.�~���[������, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.���^��W, 1) & "," & varNullChk(objTateguKansu.���^������, 1) & "," & varNullChk(objTateguKansu.��y���, 1) & "," & varNullChk(objTateguKansu.��y����, 1) & "," & varNullChk(objTateguKansu.��y�{��, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.���y���, 1) & "," & varNullChk(objTateguKansu.���y����, 1) & "," & varNullChk(objTateguKansu.���y�{��, 1) & "," & varNullChk(objTateguKansu.�A�N�Z���g���C��, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.�_�{�s�b�`1, 1) & "," & varNullChk(objTateguKansu.�_�{�s�b�`2, 1) & "," & varNullChk(objTateguKansu.�_�{�s�b�`3, 1) & "," & varNullChk(objTateguKansu.�_�{�s�b�`4, 1) & "," & varNullChk(objTateguKansu.�_�{�s�b�`5, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.�_�{�s�b�`6, 1) & "," & varNullChk(objTateguKansu.�_�{�s�b�`7, 1) & "," & varNullChk(objTateguKansu.�_�{�s�b�`8, 1) & "," & varNullChk(objTateguKansu.�_�{�s�b�`9, 1) & "," & varNullChk(objTateguKansu.�_�{�s�b�`10, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.���ԏ�ʒu, 1) & "," & varNullChk(objTateguKansu.���Ԓ���ʒu, 1) & "," & varNullChk(objTateguKansu.���Ԓ����ʒu, 1) & "," & varNullChk(objTateguKansu.���ԉ��ʒu, 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.�a(objTateguKansu.�u�����h�i��, .GetRS![�{��], .GetRS![��Spec], .GetRS![���ʒu]), 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.�n���h������Z���^�[(objTateguKansu.�u�����h�i��, .GetRS![�{��], .GetRS![��Spec]), 1) & "," & varNullChk(objCheckLabel.�n���h������BS(objTateguKansu.�u�����h�i��, .GetRS![�{��], .GetRS![��Spec]), 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.�����Z���^�[(objTateguKansu.�u�����h�i��, .GetRS![�{��], .GetRS![��Spec]), 1) & "," & varNullChk(objCheckLabel.����BS(objTateguKansu.�u�����h�i��, .GetRS![�{��], .GetRS![��Spec]), 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.�������Z���^�[(objTateguKansu.�u�����h�i��, .GetRS![�{��], .GetRS![��Spec]), 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.�q�����b�`�󂯃Z���^�[(objTateguKansu.�u�����h�i��, .GetRS![�{��], .GetRS![��Spec]), 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.�q�������󂯃Z���^�[(objTateguKansu.�u�����h�i��, .GetRS![�{��], .GetRS![��Spec]), 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.����_�O(.GetRS![�_��ԍ�], objTateguKansu.�u�����h�i��, .GetRS![�F�R�[�h]), 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.����_��(.GetRS![�_��ԍ�], objTateguKansu.�u�����h�i��, .GetRS![�F�R�[�h]), 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.�J�Z�b�g���H�}�p�X(objTateguKansu.�u�����h�i��, .GetRS![�{��], .GetRS![��Spec]), 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.�ˊJ���}�p�X(objTateguKansu.�u�����h�i��, .GetRS![�J�l��], .GetRS![�݌�], .GetRS![���ʒu]), 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![���iFLG], 1) & "," & varNullChk(strQRLabel, 1) & "," & varNullChk(strQRLabel, 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![Spec], 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.����݌v���l(objTateguKansu.�u�����h�i��, .GetRS![��Spec], .GetRS![�݌v���l]), 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![�V�ʍފ��t], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![����FLG], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![�����{�H��], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![����o�ו��@], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![�N���[���p���l], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![����m���], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![08�J�^���O], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![�����敪], 1)
                    strSQL = strSQL & "," & varNullChk(strLineKbn, 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![������], 1)
                    strSQL = strSQL & ") "
                    
                    'Debug.Print strSQL
                    If Not .ExecSQL(strSQL, strErrMsg) Then
                        Err.Raise 9999, , strErrMsg
                    End If
                    
                Else
                    If objTateguKansu.�u�����h�i�� <> "" Then
                        strBrandHinban = objTateguKansu.�u�����h�i��
                    Else
                        strBrandHinban = .GetRS![�i��1]
                    End If
                    
                    Err.Raise 9998, , "����֐������G���[ �i��=" & .GetRS![�i��1]

                    strSQL = ""
                    strSQL = strSQL & "insert into WK_��������X�g_FullHeight( "
                    strSQL = strSQL & "  �ʔ�,�_��No "
                    strSQL = strSQL & ",�_��ԍ�,���ԍ�,�����ԍ� "
                    strSQL = strSQL & ", ������, ��, �u�����h�i��, OEM�i�� "
                    strSQL = strSQL & ", �{��, �F�R�[�h, �F, ����"
                    strSQL = strSQL & ", ����, �󒍖���, �݌�, ���ʒu, ���葋,�J�l��"
                    strSQL = strSQL & ", DW, DH, CH"
                    strSQL = strSQL & ", ������, ����X�^�C��, ����`��, �e�q"
                    strSQL = strSQL & ", ��Spec "
                    strSQL = strSQL & ", ���iFLG, QR�R�[�h, QR�R�[�h_2, Spec,�݌v���l "
                    strSQL = strSQL & ", �V�ʍފ��t "
                    strSQL = strSQL & ", ����FLG "
                    strSQL = strSQL & ", �����{�H��  "
                    strSQL = strSQL & ", ����o�ו��@ "
                    strSQL = strSQL & ", �N���[���p���l "
                    strSQL = strSQL & ", ����m��� "
                    strSQL = strSQL & ", �����敪 "
                    strSQL = strSQL & ", ������ "
                    strSQL = strSQL & " ) "
                    strSQL = strSQL & "values "
                    strSQL = strSQL & "( "
                    strSQL = strSQL & varNullChk(.GetRS![�ʔ�], 1) & "," & varNullChk(.GetRS![�_��No], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![�_��ԍ�], 1) & "," & varNullChk(.GetRS![���ԍ�], 1) & "," & varNullChk(.GetRS![�����ԍ�], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![������], 1) & "," & varNullChk(.GetRS![��], 1) & "," & varNullChk(strBrandHinban, 1) & "," & varNullChk(strOEMHinban, 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![�{��], 1) & "," & varNullChk(.GetRS![�F�R�[�h], 1) & "," & varNullChk(.GetRS![�F], 1) & "," & varNullChk(.GetRS![����], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![����], 1) & "," & varNullChk(.GetRS![�󒍖���], 1) & "," & varNullChk(strTsurimoto, 1) & "," & varNullChk(.GetRS![���ʒu], 1) & "," & varNullChk(.GetRS![���葋], 1) & "," & varNullChk(.GetRS![�J�l��], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![DW], 1) & "," & varNullChk(.GetRS![DH], 1) & "," & varNullChk(.GetRS![CH], 1)
                    strSQL = strSQL & "," & varNullChk(intHanWari, 1) & "," & varNullChk(intTateguStyle, 1) & "," & varNullChk(intTateguKeijo, 1) & "," & varNullChk(intOyako, 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![��Spec], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![���iFLG], 1) & "," & varNullChk(.GetRS![QR�R�[�h], 1) & "," & varNullChk(.GetRS![QR�R�[�h_2], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![Spec], 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.����݌v���l(objTateguKansu.�u�����h�i��, .GetRS![��Spec], .GetRS![�݌v���l]), 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![�V�ʍފ��t], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![����FLG], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![�����{�H��], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![����o�ו��@], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![�N���[���p���l], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![����m���], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![�����敪], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![������], 1)
                    strSQL = strSQL & ") "
                    
                    'Debug.Print strSQL
                    If Not .ExecSQL(strSQL, strErrMsg) Then
                        Err.Raise 9999, , strErrMsg
                    End If
                    
                End If
                
                    inReadCount = inReadCount + 1

                    
                    SysCmd acSysCmdSetStatus, "���s��.... " & inReadCount & "/" & inRecordCount
                    If inReadCount Mod 10 = 0 Then
                        DoEvents
                    End If
                    .GetRS.MoveNext

            Loop
        End If
    
    End With
    
    If bolfnc�����w���t���n�C�g���[�f�[�^_�����X�V Then
        If bolfnc�����w���t���n�C�g���[�f�[�^_�����C���w�����擾 Then
            bolfnc�����w���t���n�C�g���[�f�[�^ = True
        End If
    End If
    

    GoTo Exit_bolfnc�����w���t���n�C�g���[�f�[�^
    
Err_bolfnc�����w���t���n�C�g���[�f�[�^:
    If Err.Number = 9998 Then
        Debug.Print Err.Description
        Resume Next
    Else
        MsgBox Err.Description
        
    End If
    'Resume
Exit_bolfnc�����w���t���n�C�g���[�f�[�^:
    SysCmd acSysCmdSetStatus, " "
    Set objLocalDB = Nothing
    Set objTateguKansu = Nothing
End Function

Public Function bolfnc�����w���t���n�C�g���[�f�[�^_�����X�V() As Boolean
    Dim objLocalDB As New cls_LOCALDB
    Dim objREMOTEDB As New cls_BRAND_MASTER
    
    Dim strSQL As String
    Dim conSQL As String
    
    Dim strBuzaimei As String
    Dim dblBuzaisuGoukei As Double
    
    Dim i As Integer
    
    bolfnc�����w���t���n�C�g���[�f�[�^_�����X�V = False
    
    On Error GoTo Err_bolfnc�����w���t���n�C�g���[�f�[�^_�����X�V
    
    strBuzaimei = ""
    dblBuzaisuGoukei = 0
    
    strSQL = "update WK_��������X�g_FullHeight "
    strSQL = strSQL & "set ��t����1 = null "
    strSQL = strSQL & ", ��t����1���� = null "
    strSQL = strSQL & ", ��t����2 = null "
    strSQL = strSQL & ", ��t����2���� = null "
    strSQL = strSQL & ", ��t����3 = null "
    strSQL = strSQL & ", ��t����3���� = null "
    strSQL = strSQL & ", ��t����4 = null "
    strSQL = strSQL & ", ��t����4���� = null "
    strSQL = strSQL & ", ��t����5 = null "
    strSQL = strSQL & ", ��t����5���� = null "
    strSQL = strSQL & ", ��t����6 = null "
    strSQL = strSQL & ", ��t����6���� = null "
    strSQL = strSQL & ", ��t����7 = null "
    strSQL = strSQL & ", ��t����7���� = null "
    strSQL = strSQL & ", ��t����8 = null "
    strSQL = strSQL & ", ��t����8���� = null "
    strSQL = strSQL & ", ��t����9 = null "
    strSQL = strSQL & ", ��t����9���� = null "
    strSQL = strSQL & ", ��t����10 = null "
    strSQL = strSQL & ", ��t����10���� = null "
    
    If Not objLocalDB.ExecSQL(strSQL) Then
        Err.Raise 9999, , "�������G���["
    End If
    
    conSQL = ""
    conSQL = conSQL & "select a.*,b.���� �󒍖��� from BRAND_BOM.dbo.T_���ޓW�J a "
    conSQL = conSQL & "inner join T_�󒍖��� b "
    conSQL = conSQL & "on a.�_��ԍ� = b.�_��ԍ� and a.���ԍ� = b.���ԍ� and a.�����ԍ� = b.�����ԍ� and a.�� = b.�� "
    conSQL = conSQL & "where a.�_��ԍ� = '@�_��ԍ�@' and a.���ԍ� = '@���ԍ�@' and a.�����ԍ� = '@�����ԍ�@' and a.�� = @��@ "
    conSQL = conSQL & "and a.�i�ԋ敪 = 1 "
    conSQL = conSQL & "and a.��t = '��' "
    
    strSQL = "select * from WK_��������X�g_FullHeight "
    
    With objLocalDB
        If .ExecSelect_Writable(strSQL) Then
            If Not .GetRS.EOF Then
                Do Until .GetRS.EOF
                    strSQL = conSQL
                    strSQL = Replace(strSQL, "@�_��ԍ�@", .GetRS![�_��ԍ�])
                    strSQL = Replace(strSQL, "@���ԍ�@", .GetRS![���ԍ�])
                    strSQL = Replace(strSQL, "@�����ԍ�@", .GetRS![�����ԍ�])
                    strSQL = Replace(strSQL, "@��@", .GetRS![��])

                    If objREMOTEDB.ExecSelect(strSQL) Then
                        If Not objREMOTEDB.GetRS.EOF Then
                            i = 1
                            Do Until objREMOTEDB.GetRS.EOF
                                If .GetRS![���ʒu] = 0 Then
                                    strBuzaimei = objREMOTEDB.GetRS![���ޖ�]
                                    dblBuzaisuGoukei = objREMOTEDB.GetRS![���ސ����v] / objREMOTEDB.GetRS![�󒍖���]
                                Else
                                    If (.GetRS![�J�l��] = "DF" And .GetRS![�u�����h�i��] Like "*-####*HY-*") Or (.GetRS![�J�l��] = "VF" And .GetRS![�u�����h�i��] Like "*-####*HF-*") Then
                                        If objREMOTEDB.GetRS![���ގ��CD] Like "*���޸����*" Then
                                            If .GetRS![���ʒu] <> 3 Then
                                                strBuzaimei = objREMOTEDB.GetRS![���ޖ�]
                                                dblBuzaisuGoukei = objREMOTEDB.GetRS![���ސ����v] / (objREMOTEDB.GetRS![�󒍖���] - 1)
                                            End If
                                        ElseIf objREMOTEDB.GetRS![���ގ��CD] Like "*�ˎ�*" Then
                                            If .GetRS![���ʒu] = 3 Then
                                                strBuzaimei = objREMOTEDB.GetRS![���ޖ�]
                                                dblBuzaisuGoukei = objREMOTEDB.GetRS![���ސ����v] / (objREMOTEDB.GetRS![�󒍖���] - 2)
                                            End If
                                        ElseIf objREMOTEDB.GetRS![���ގ��CD] Like "*�ر������*" And IsTateguInset(.GetRS![�u�����h�i��]) Then
                                            If .GetRS![���ʒu] <> 3 Then
                                                strBuzaimei = objREMOTEDB.GetRS![���ޖ�]
                                                dblBuzaisuGoukei = objREMOTEDB.GetRS![���ސ����v] / (objREMOTEDB.GetRS![�󒍖���] - 1)
                                            End If
                                        Else
                                            strBuzaimei = objREMOTEDB.GetRS![���ޖ�]
                                            dblBuzaisuGoukei = objREMOTEDB.GetRS![���ސ����v] / objREMOTEDB.GetRS![�󒍖���]
                                        End If
                                    Else
                                        strBuzaimei = objREMOTEDB.GetRS![���ޖ�]
                                        dblBuzaisuGoukei = objREMOTEDB.GetRS![���ސ����v] / objREMOTEDB.GetRS![�󒍖���]
                                    End If
                                End If
                                
                                If strBuzaimei <> "" Then
                                    .GetRS.Update "��t����" & i, strBuzaimei
                                    .GetRS.Update "��t����" & i & "����", dblBuzaisuGoukei
                                    i = i + 1
                                End If
                                
                                strBuzaimei = ""
                                dblBuzaisuGoukei = 0
                                
                                If i > 10 Then Exit Do
                                
                                objREMOTEDB.GetRS.MoveNext
                            Loop
                            objREMOTEDB.RecordSetClose
                        End If
                    Else
                        Err.Raise 9999, , "���ޓW�J���o���s�G���[ " & strSQL
                    End If
                    .GetRS.MoveNext
                Loop
            End If

        Else
            Err.Raise 9999, , "Input���s�G���[ " & strSQL
        End If
    End With
    
    bolfnc�����w���t���n�C�g���[�f�[�^_�����X�V = True
    
    GoTo Exit_bolfnc�����w���t���n�C�g���[�f�[�^_�����X�V
   
Err_bolfnc�����w���t���n�C�g���[�f�[�^_�����X�V:
    MsgBox Err.Description
    'Resume
Exit_bolfnc�����w���t���n�C�g���[�f�[�^_�����X�V:
    Set objREMOTEDB = Nothing
    Set objLocalDB = Nothing
    
End Function

Private Function bolfnc�����w���t���n�C�g���[�f�[�^_�����C���w�����擾() As Boolean
    Dim objLocalDB As New cls_LOCALDB
    Dim objREMOTEDB As New cls_BRAND_MASTER
    
    Dim strSQL As String
    Dim strSQLR As String
    Dim conSQL As String
    Dim consqlR As String
    
    On Error GoTo Err_bolfnc�����w���t���n�C�g���[�f�[�^_�����C���w�����擾
    
    conSQL = ""
    conSQL = conSQL & "update WK_��������X�g_FullHeight "
    conSQL = conSQL & "set ���C��ALL���� = @�����ALL@ "
    conSQL = conSQL & ",���C��2���� = @�����2@ "
    conSQL = conSQL & ",�܌˖��� = @�܌˖���@ "
    conSQL = conSQL & "where �_��No = '@�_��No@' "
    
    consqlR = ""
    consqlR = consqlR & "select sum(����) as �܌˖��� from T_�󒍖��� "
    consqlR = consqlR & "where �_��ԍ� = '@�_��ԍ�@' "
    consqlR = consqlR & "and ���ԍ� = '@���ԍ�@' "
    consqlR = consqlR & "and �����ԍ� = '@�����ԍ�@' "
    consqlR = consqlR & "and ��� = '�۾ޯ�' "
    consqlR = consqlR & "and dbo.IsCloset_Isehara(dbo.fncgetHinban(�i��1,��������i��)) = 0 "

    strSQL = ""
    strSQL = strSQL & "select ALLDATA.�_��ԍ�,ALLDATA.���ԍ�,ALLDATA.�����ԍ�,ALLDATA.�_��No,�����ALL,�����2 from "
    strSQL = strSQL & "(select �_��ԍ�,���ԍ�,�����ԍ�,�_��No,sum(����) as �����ALL from WK_�����˗������� group by �_��ԍ�,���ԍ�,�����ԍ�,�_��No) as ALLDATA "
    strSQL = strSQL & "left join "
    strSQL = strSQL & "(select �_��No,sum(����) as �����2 from WK_�����˗�������_FullHeight group by �_��No) as FillHeightLine "
    strSQL = strSQL & "on ALLDATA.�_��No = FillHeightLine.�_��No "
    
    With objLocalDB
        If .ExecSelect(strSQL) Then
            Do Until .GetRS.EOF
                                
                strSQLR = consqlR
                strSQLR = Replace(strSQLR, "@�_��ԍ�@", .GetRS![�_��ԍ�])
                strSQLR = Replace(strSQLR, "@���ԍ�@", .GetRS![���ԍ�])
                strSQLR = Replace(strSQLR, "@�����ԍ�@", .GetRS![�����ԍ�])
                
                strSQL = conSQL
                strSQL = Replace(strSQL, "@�����ALL@", Nz(.GetRS![�����ALL], 0))
                strSQL = Replace(strSQL, "@�����2@", Nz(.GetRS![�����2], 0))
                strSQL = Replace(strSQL, "@�_��No@", .GetRS![�_��No])
                
                With objREMOTEDB
                    If .ExecSelect(strSQLR) Then
                        If Not .GetRS.EOF Then
                            strSQL = Replace(strSQL, "@�܌˖���@", Nz(.GetRS![�܌˖���], 0))
                        Else
                            strSQL = Replace(strSQL, "@�܌˖���@", 0)
                        End If
                    Else
                        Err.Raise 9999, , "�܌˖��������G���[ SQL=" & strSQLR
                    End If
                End With
                'Debug.Print strSQL
                
                If Not .ExecSQL(strSQL) Then
                    Err.Raise 9999, , "SQL���s�G���[ SQL=" & strSQL
                End If
                
                .GetRS.MoveNext
            Loop
        Else
            Err.Raise 9999, , "�����W�v���s�G���[ "
        End If
    
    End With
    
    bolfnc�����w���t���n�C�g���[�f�[�^_�����C���w�����擾 = True
    
    GoTo Exit_bolfnc�����w���t���n�C�g���[�f�[�^_�����C���w�����擾

Err_bolfnc�����w���t���n�C�g���[�f�[�^_�����C���w�����擾:
    MsgBox Err.Description
Exit_bolfnc�����w���t���n�C�g���[�f�[�^_�����C���w�����擾:
    Set objLocalDB = Nothing
    Set objREMOTEDB = Nothing
    
End Function