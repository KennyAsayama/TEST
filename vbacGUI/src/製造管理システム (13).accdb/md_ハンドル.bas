Option Compare Database
Option Explicit

Public Function fncbol_Handle_����_��(�{�� As String, ��Spec As String) As Boolean
'--------------------------------------------------------------------------------
'           ���T�C�Y�̈�������m�F����
'
'����       �{��            : �{���R�[�h(2 or 3��)
'           ��Spec        : �d�l�R�[�h(7��)
'�߂�l     Boolean         : True  �����T�C�Y����  False   �����T�C�Y����ł͂Ȃ�
'--------------------------------------------------------------------------------
    Dim str��� As String
    Dim strSPEC As String
    
    str��� = left(��Spec, 3)
    strSPEC = right(��Spec, 4)
    
    Select Case strSPEC
        Case Is >= "1410"

            If �{�� Like "ZA?" Or �{�� Like "ZC?" Or �{�� Like "ZV?" Or �{�� Like "ZE?" Then
            
                fncbol_Handle_����_�� = True
            Else
                fncbol_Handle_����_�� = False
                
            End If
        
        Case Is >= "1404"
            If �{�� Like "A?" Or �{�� Like "C?" Or �{�� Like "V?" Then
               
                fncbol_Handle_����_�� = True
                
            Else
                fncbol_Handle_����_�� = False
                
            End If
        Case Is >= "1011"
            If �{�� Like "W?" Then
               
                fncbol_Handle_����_�� = True
                
            Else
                fncbol_Handle_����_�� = False
                
            End If
        Case "0911" To "1010"
            If �{�� Like "Z?" Or �{�� Like "Y?" Then
               
                fncbol_Handle_����_�� = True
                
            Else
                fncbol_Handle_����_�� = False
                
            End If
            
        Case Else
            fncbol_Handle_����_�� = False
            
    End Select

End Function

Public Function fncbol_Handle_����_�Z(�{�� As String, ��Spec As String) As Boolean
'--------------------------------------------------------------------------------
'           �Z�T�C�Y�̈�������m�F����
'
'����       �{��            : �{���R�[�h(2 or 3��)
'           ��Spec        : �d�l�R�[�h(7��)
'�߂�l     Boolean         : True  ���Z�T�C�Y����  False   ���Z�T�C�Y����ł͂Ȃ�
'--------------------------------------------------------------------------------

    Dim str��� As String
    Dim strSPEC As String
    
    str��� = left(��Spec, 3)
    strSPEC = right(��Spec, 4)
    
    Select Case strSPEC
        Case Is >= "1410"

            If �{�� Like "ZG?" Or �{�� Like "ZH?" Or �{�� Like "ZB?" Or �{�� Like "ZD?" Then
            
                fncbol_Handle_����_�Z = True
            Else
                fncbol_Handle_����_�Z = False
                
            End If
            
        Case Is >= "1307"
            If �{�� Like "G?" Or �{�� Like "H?" Or �{�� Like "B?" Then
               
                fncbol_Handle_����_�Z = True
                
            Else
                fncbol_Handle_����_�Z = False
                
            End If
            
        Case Else
            fncbol_Handle_����_�Z = False
    End Select
    
End Function

Public Function fncbol_Handle_EStyle(�{�� As String, ��Spec As String) As Boolean
'--------------------------------------------------------------------------------
'E-Style�̃f�t�H���g�̃n���h��(I,J,K)���m�F����
'
'����       �{��            : �{���R�[�h(2 or 3��)
'           ��Spec        : �d�l�R�[�h(7��)
'�߂�l     Boolean         : True  ��E-Style�n���h��  False   ��E-Style�n���h���ł͂Ȃ�
'--------------------------------------------------------------------------------

    Dim str��� As String
    Dim strSPEC As String
    
    str��� = left(��Spec, 3)
    strSPEC = right(��Spec, 4)
    

    Select Case strSPEC

        Case Is >= "1601"
            If �{�� Like "CI?" Or �{�� Like "CJ?" Or �{�� Like "CK?" Or �{�� Like "CG?" Or �{�� Like "CH?" Or �{�� Like "CB?" _
                Or �{�� Like "DI?" Or �{�� Like "DJ?" Or �{�� Like "DK?" Or �{�� Like "DG?" Or �{�� Like "DH?" Or �{�� Like "DB?" Then
                fncbol_Handle_EStyle = True
            Else
                fncbol_Handle_EStyle = False
            End If

        Case Is >= "1410"
            If �{�� Like "AI?" Or �{�� Like "AJ?" Or �{�� Like "AK?" Or �{�� Like "AA?" Or �{�� Like "AH?" Then
                fncbol_Handle_EStyle = True
            Else
                fncbol_Handle_EStyle = False
            End If
        Case Is >= "0911"
            If �{�� Like "I?" Or �{�� Like "J?" Or �{�� Like "K?" Then
                fncbol_Handle_EStyle = True
            Else
                fncbol_Handle_EStyle = False
            End If
        Case Else
            fncbol_Handle_EStyle = False
    End Select

                               
End Function

Public Function fncbol_Handle_WanNyan(�{�� As String, ��Spec As String) As Boolean
'--------------------------------------------------------------------------------

'           ���ɂ��n���h��(O,N,E,F)���m�F����
'
'����       �{��            : �{���R�[�h(2 or 3��)
'           ��Spec        : �d�l�R�[�h(7��)
'�߂�l     Boolean         : True  �����ɂ��n���h��  False   �����ɂ��n���h���ł͂Ȃ�
'--------------------------------------------------------------------------------
    Dim str��� As String
    Dim strSPEC As String
    
    str��� = left(��Spec, 3)
    strSPEC = right(��Spec, 4)
    

    Select Case strSPEC
        Case Is >= "1410"
            If �{�� Like "AE?" Or �{�� Like "AF?" Or �{�� Like "AN?" Or �{�� Like "AO?" Then
                fncbol_Handle_WanNyan = True
            Else
                fncbol_Handle_WanNyan = False
            End If
        Case Is >= "1404"
            If �{�� Like "E?" Or �{�� Like "F?" Or �{�� Like "N?" Or �{�� Like "O?" Then
                fncbol_Handle_WanNyan = True
            Else
                fncbol_Handle_WanNyan = False
            End If
        Case Else
            fncbol_Handle_WanNyan = False
    End Select

                               
End Function

Public Function fncstrHandle_Name(�{�� As String, ��Spec As String) As String
'--------------------------------------------------------------------------------

'           �n���h���R�[�h�𒊏o����
'
'����       �{��            : �{���R�[�h(2 or 3��)
'           ��Spec        : �d�l�R�[�h(7��)
'�߂�l     String(1~2��)   : �n���h���R�[�h
'--------------------------------------------------------------------------------

    Dim str��� As String
    Dim strSPEC As String
    
    str��� = left(��Spec, 3)
    strSPEC = right(��Spec, 4)
    
    If �{�� Like "��*" Then
        fncstrHandle_Name = "��"
        
    ElseIf str��� = "TSC" Then
        fncstrHandle_Name = left(�{��, 2)
    
    Else
        Select Case strSPEC
            Case Is >= "1410"
                If Len(�{��) > 2 Then '�۾ޯĂ�2��
                    fncstrHandle_Name = left(�{��, 2)
                Else
                    fncstrHandle_Name = left(�{��, 1)
                End If
            Case Else
                fncstrHandle_Name = left(�{��, 1)
        End Select
     End If
End Function

Public Function fncbol��(�{�� As String, ��Spec As String) As Boolean
'--------------------------------------------------------------------------------

'           �����薳���𔻒肷��
'
'����       �{��            : �{���R�[�h(2 or 3��)
'           ��Spec        : �d�l�R�[�h(7��)
'�߂�l     Boolean         : True  ��������  False   �����Ȃ�
'--------------------------------------------------------------------------------
    
    If �{�� Like "*C" Or �{�� Like "*M" Or �{�� Like "*K" Then
        fncbol�� = True
    Else
        fncbol�� = False
    End If
    
End Function

Public Function fncbol_Handle_Vertica(�{�� As String, ��Spec As String) As Boolean
'--------------------------------------------------------------------------------
'           ���F���`�J�i��背�X�j(-N,-B)���m�F����
'
'����       �{��            : �{���R�[�h(3��)
'           ��Spec        : �d�l�R�[�h(7��)
'�߂�l     Boolean         : True  �����F���`�J  False   �����F���`�J�ł͂Ȃ�
'--------------------------------------------------------------------------------
    Dim str��� As String
    Dim strSPEC As String
    
    str��� = left(��Spec, 3)
    strSPEC = right(��Spec, 4)
    

    Select Case strSPEC
        Case Is >= "1507"
            If �{�� Like "-N?" Or �{�� Like "-B?" Or �{�� Like "-Q?" Or �{�� Like "-K?" Then
                fncbol_Handle_Vertica = True
            Else
                fncbol_Handle_Vertica = False
            End If
        Case Else
            fncbol_Handle_Vertica = False
    End Select

                               
End Function

Public Function fncstrHandleKigoFileName(ByVal in_Hinban As String, ByVal in_Handle As String, ByVal in_spec As String) As String
'   *************************************************************
'   �n���h���L������L���}�̃t�@�C�������擾

'       ���V���C���p�f�[�^���o����܂ł̎b�藘�p
'
'   �߂�l:String
'       ���摜�t�@�C����
'
'    Input����
'       in_Hinban           ����i��
'       in_Handle           �{��
'       in_Spec             ��SPec

'2.1.0
'   ��DP,DQ�n���h���ǉ��i���O�j
'2.3.0
'   ��1801�d�l�ǉ�
'2.7.0
'   ��1808�d�l�ǉ�
'2.13.0
'   ��1901�d�l�ǉ�(HE-HI)
'   *************************************************************
    Dim Reg As Object
    '�f�B�N�V���i��
    Dim dictHandle As Object
    '�z��p�i�f�B�N�V���i���̃p�^�[���}���j
    Dim varPattern As Variant
    
    Dim i As Integer
    Dim strHandleHikiteName As String
    
    fncstrHandleKigoFileName = ""
    
    If in_Hinban = "" Or in_Handle = "" Or in_spec = "" Then Exit Function
    
    '�{���R�[�h��3���i���������j�ȊO�͑Ή����Ȃ��i�ߋ��̃R�[�h�͑Ή����Ȃ��j
    If Len(in_Handle) <> 3 Then Exit Function
    If in_Handle Like "*��*" Then Exit Function
    
    '���K�\��
    Set Reg = CreateObject("VBScript.RegExp")
    '�f�B�N�V���i��
    Set dictHandle = CreateObject("Scripting.Dictionary")
    
    On Error GoTo Err_fncstrHandleKigoFileName
    
    '�f�B�N�V���i���ɐ��K�\���̃p�^�[����}��
    '�i�L�[���p�^�[���A�A�C�e�����t�@�C�����j
    '���x����̂��߂ɓ�����₷���p�^�[��������ׂ�ׂ�
    
    With dictHandle
        If IsHirakido(in_Hinban) Or IsOyatobira(in_Hinban) Then
    
            
            .Add "^(C[B-KST]|D[B-FI-KST])", "SHIBUTANI"
            .Add "^(H[A-IPQR])", "KAWAJ_LJ"
            .Add "^(C[N-R]|D[NO])", "KAWAJUN"
            .Add "^(C[LM]|B[YZ]|B[ACDEFHIJLMNOPQRS]|D[PQ])", "KURAMAE"
            .Add "^A[EFON]", "NAGASAWA"
            .Add "^F[C-H]", "OLIVARI"

        ElseIf IsHikido(in_Hinban) Or IsCloset_Hikichigai(in_Hinban) Then
    
            .Add "^Z[ACEV]", "OKUDAIRATK"
            .Add "^Z[BDHG]", "OKUDAIRA"
            .Add "^-(?!-).(?!N)", "HIKITELESS"
           
        End If
        
        '�L�[�Ƀ}�b�`������t�@�C�������擾
        If .Count > 0 Then '�����˂ł��J���˂ł��Ȃ��ꍇ�̓J�E���g���[��
            varPattern = .keys
            
            For i = 0 To UBound(varPattern)
            
                Reg.Pattern = varPattern(i)
                
                If Reg.Test(in_Handle) Then
                    strHandleHikiteName = .Item(varPattern(i))
                    'Debug.Print "Match! " & in_Hinban & " key=" & strHandleHikiteName
                    Exit For
                End If
                
            Next
            
        End If
    
    End With
    
    If strHandleHikiteName = "" Then
        'Debug.Print "Nomatch!"
    Else
        '���t���̏ꍇ�̓t�@�C�����ɂ���ɏ��̖��O��t��
        If fncbol��(in_Handle, in_spec) Then
        
            strHandleHikiteName = strHandleHikiteName & "_Lock"
            
            '�����j�����̏ꍇ�̓V�����_�[���A�Ԏd�؂���̋�ʂ���
            If fncbol_Handle_WanNyan(in_Handle, in_spec) Then
                Select Case right(in_Handle, 1)
                    Case "C"
                        strHandleHikiteName = strHandleHikiteName & "_Cylinder"
                    Case "M", "K"
                        strHandleHikiteName = strHandleHikiteName & "_Majikiri"
                End Select
            
            '�G���h�g�����̓A�E�g�Z�b�g���ˏ�
            ElseIf IsEndWakunashi(in_Hinban) Then
            
                strHandleHikiteName = strHandleHikiteName & "_Outset"
                
            End If
        End If
    End If
    
    fncstrHandleKigoFileName = strHandleHikiteName
    
    GoTo Exit_fncstrHandleKigoFileName

Err_fncstrHandleKigoFileName:
    Debug.Print Err.Description
    fncstrHandleKigoFileName = ""
    
Exit_fncstrHandleKigoFileName:
    Set Reg = Nothing
    Set dictHandle = Nothing
        
End Function

Public Function fncstrHikiteColorName(ByVal in_Handle As String, ByVal in_spec As String) As String
'   *************************************************************
'   ������L�����������F���擾

'
'   �߂�l:String
'       ���F��
'
'    Input����
'       in_Handle           �{��
'       in_Spec             ��SPec

'   *************************************************************
    
    fncstrHikiteColorName = ""
    
    If in_Handle = "" Or in_spec = "" Then Exit Function
    
    '�{���R�[�h��3���i���������j�ȊO�͑Ή����Ȃ��i�ߋ��̃R�[�h�͑Ή����Ȃ��j
    If Len(in_Handle) <> 3 Then Exit Function
    If in_Handle Like "*��*" Then Exit Function
    
    If in_Handle Like "ZV*" Or in_Handle Like "ZH*" Then
        fncstrHikiteColorName = "�T�e���j�b�P��"
    ElseIf in_Handle Like "ZC*" Or in_Handle Like "ZG*" Then
        fncstrHikiteColorName = "�N���[��"
    ElseIf in_Handle Like "ZA*" Or in_Handle Like "ZB*" Then
        fncstrHikiteColorName = "�z���C�g"
    ElseIf in_Handle Like "ZE*" Or in_Handle Like "ZD*" Then
        fncstrHikiteColorName = "�u���b�N"
    End If
    
End Function

Public Function IsHikiteKako(ByVal in_varHinban As Variant, ByVal in_varTobiraichi As Variant, ByVal in_varTsurimoto As Variant, ByVal in_varSpec As Variant) As Boolean
'   *************************************************************
'   ��������H�����邩�m�F
'
'   �߂�l:Boolean
'       True                ������H����
'       False               ���H�Ȃ��i�Ȃ��̏����ȊO�͂��ׂĂ���ŕԂ��j
'
'    Input����
'       in_varHinban        ����i��
'       in_varTobiraichi    �ʒu�i1.���A2.�E�A3.���j
'       in_Spec             ��SPec(�쐬�����g�p�j

'2.14.0 ADD
'   *************************************************************
    
    Dim strHinban As String
    Dim intTobiraichi As Integer
    Dim strSPEC As String
    Dim strTsurimoto As String
    
    IsHikiteKako = True
    
    '�i�ԂȂ��̏ꍇ��True�ŕԂ�
    If IsNull(in_varHinban) Then
        Exit Function
    End If
    
    '�ʒu�Ȃ��̏ꍇ��True�ŕԂ�

    If IsNull(in_varTobiraichi) Then
        Exit Function
    End If
    
    If IsNull(in_varTsurimoto) Then
        strTsurimoto = "Z"
    Else
        strTsurimoto = in_varTsurimoto
    End If
    
    If IsNull(in_varSpec) Then
        strSPEC = ""
    Else
        strSPEC = in_varSpec
    End If
    
    strHinban = Replace(in_varHinban, "�� ", "")
    intTobiraichi = in_varTobiraichi
    
    If IsSynchro(strHinban) Then
        If IsHikichigai(strHinban) Then
            If intTobiraichi = 3 Then
                IsHikiteKako = False
            End If
        Else
            If strTsurimoto = "L" Then 'L�݌�
                If intTobiraichi <> 1 Then
                    IsHikiteKako = False
                End If
            ElseIf strTsurimoto = "R" Then 'R�݌�
                If intTobiraichi <> 2 Then
                    IsHikiteKako = False
                End If
            End If
        End If
    
    End If

End Function