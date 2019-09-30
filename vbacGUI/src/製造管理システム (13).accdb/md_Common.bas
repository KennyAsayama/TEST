Option Compare Database
Option Explicit
'--------------------------------------------------------------------------------------------------------------------
'���ʕϐ�

'2.6.0
'   ��bolShizaiUpdatable�@�ǉ�
'3.0.0
'   ��bolsekkei �ǉ�
'--------------------------------------------------------------------------------------------------------------------
'�{�ԃf�[�^�x�[�X��
Public Const strDBName As String = "DB02"

'�p�X���[�h����
Public Const constintPassWordLength As Integer = 5

'���[�U�[ID,������
Public strUserID As String
Public strUserName As String
Public bolUpdatable As Boolean
Public bolAdministrator As Boolean

Public bolShizaiUpdatable As Boolean

Public bolSekkei As Boolean

'1.10.6 K.Asayama 20151211 �ǉ�
'SxL���[�J���R�s�[,�J�����_�[�R�s�[
Public bolSxLCopy As Boolean
Public bolCalendarCopy As Boolean

'1.12.3 ADD
'�T�[�o�p�X
Public Const conServerPath As String = "\\db\prog\�����Ǘ��V�X�e��"
Public Const conUserPath As String = "\\db\prog\�����Ǘ��V�X�e��"

Public Sub UserINIT()
'--------------------------------------------------------------------------------------------------------------------
'���[�U�[�֘A�֐�������

''1.10.6 K.Asayama bolSxLCopy,bolCalendarCopy �������ǉ� 20151211 �ǉ�
'--------------------------------------------------------------------------------------------------------------------
    strUserID = ""
    strUserName = ""
    bolUpdatable = False
    bolAdministrator = False
    
    bolSxLCopy = False
    bolCalendarCopy = False
    
End Sub

Public Function Connection_DB() As String
'--------------------------------------------------------------------------------------------------------------------
'���ݎg�p���Ă���DB���ɃZ�b�g����Ă���ODBC�̐ڑ���DB����Ԃ�
'�߂�l : SQLSERVER��DB�� (String�`���j

'���ӁF32bit�ł�ACCESS�p�Ȃ̂�64bit�łɈڍs�����ꍇ���W�X�g���̃f�B���N�g�������킹��K�v����
'--------------------------------------------------------------------------------------------------------------------

    Dim ConnectDB As String
    
    Dim objWshShell As Object
    Dim OSBit As Byte
    Dim strConnDB As String
    
    ConnectDB = strDBName & "_BRAND_MASTER" 'strDBName��PUBLIC�ϐ�
    
    OSBit = OS_Architecture()
    strConnDB = ""
    
    Set objWshShell = CreateObject("WScript.Shell")
    
    If OSBit = 64 Then
        strConnDB = objWshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\ODBC\ODBC.INI\" & ConnectDB & "\Server")
    Else
        strConnDB = objWshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI\" & ConnectDB & "\Server")
    End If
    
    Connection_DB = strConnDB
    
    Set objWshShell = Nothing
    
End Function

Public Function OS_Architecture() As Byte
'--------------------------------------------------------------------------------------------------------------------
'OS��32bit�ł�64bit�ł����`�F�b�N���Đ����ŕԂ�
'�߂�l : 32 �܂��� 64 (byte�`���j
'--------------------------------------------------------------------------------------------------------------------

    Dim objWshShell As Object
    Dim strMode As String

    Set objWshShell = CreateObject("WScript.Shell")
    
    strMode = objWshShell.Environment("Process").Item("PROCESSOR_ARCHITECTURE")

    If UCase(strMode) = "X86" Then
         OS_Architecture = 32
    Else
         OS_Architecture = 64
    End If
    
    Set objWshShell = Nothing
    
End Function

Public Function to_Capital(intKeyASCII As Integer) As Integer
'--------------------------------------------------------------------------------------------------------------------
'���������啶���ϊ�
'�������̏ꍇ�͑啶���ɕϊ����ĕԂ��B����ȊO�͂��̂܂ܕԂ��B
'--------------------------------------------------------------------------------------------------------------------
        Select Case intKeyASCII
            'Case 48 To 57 '0�`9
            'Case 65 To 90 'A�`Z
            Case 97 To 122 'a�`z
                to_Capital = intKeyASCII - 32 '�啶���ɋ����ϊ�
            Case Else
                to_Capital = intKeyASCII
        End Select
End Function

Public Function RPAD(strValue As String, strCharactor As String, intKeta As Integer) As String
'--------------------------------------------------------------------------------------------------------------------
'��������
'string�̉E���Ɏw�肳�ꂽ�����𕶎������ɂȂ�悤���߂�
'--------------------------------------------------------------------------------------------------------------------
        RPAD = strValue & String(intKeta - Len(strValue), strCharactor)
        
End Function

Public Function IsNumber(intKeyASCII As Integer, Optional bolBackSpace As Variant) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'   �����m�F
'
'   �߂�l:Boolean
'       ��True              ����
'       ��False             �����ȊO
'
'    Input����
'       intKeyASCII         �L�[�{�[�h���͒l�iASCII�l)
'       bolBackSpace        True�̏ꍇ��BackSpace�L�[(Keyascii=8)�𐔎��Ƃ��ĔF������

'--------------------------------------------------------------------------------------------------------------------

        IsNumber = False
        
        Select Case intKeyASCII
            Case 48 To 57 '0�`9
                IsNumber = True
            Case Else
                If Not IsMissing(bolBackSpace) Then
                    If bolBackSpace And intKeyASCII = 8 Then
                       IsNumber = True
                    End If
                End If
        End Select
End Function

Public Function LPAD(strValue As String, strCharactor As String, intKeta As Integer) As String
'--------------------------------------------------------------------------------------------------------------------
'string�̍����Ɏw�肳�ꂽ�����𕶎������ɂȂ�悤���߂�
'--------------------------------------------------------------------------------------------------------------------
    
        LPAD = String(intKeta - Len(strValue), strCharactor) & strValue
        
End Function

Public Function IsKeiyakuNo(in_Text As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'���͂��ꂽ�����񂪌_��ԍ����m�F����(*-*-*�^���j
'   �߂�l:Boolean
'       ��True              �_��ԍ�
'       ��False             �_��ԍ��ȊO
'
'    Input����
'       in_Text             ���͒l
'--------------------------------------------------------------------------------------------------------------------

    If in_Text Like "SR*" Then
        If in_Text Like "SR####-###-####" Or in_Text Like "SR#####-###-####" Then
            IsKeiyakuNo = True '�Z�F
        End If
    Else
        If in_Text Like "??####-###-####" Then
            IsKeiyakuNo = True '�u�����h
        End If
    End If
End Function

Public Function bytfncCheckDigit_JAN(in_strCode As String) As Byte
'--------------------------------------------------------------------------------------------------------------------
'�`�F�b�N�f�B�W�b�g�v�Z�iJAN�R�[�h��p�j
'�v�Z���@��[���W�����X10/�E�F�C�g3]
'   �߂�l:Byte
'                           ���`�F�b�N�f�B�W�b�g
'                             �i�G���[�̎���99��Ԃ��j
'
'    Input����
'       in_strCode            JAN�R�[�h�i12���j
'--------------------------------------------------------------------------------------------------------------------
    Dim bytCode(11) As Byte
    Dim i As Byte
    Dim strDigit As String
    
    bytfncCheckDigit_JAN = 99
    
    On Error GoTo Err_bytfncCheckDigit_JAN
    
    
    If Not IsNumeric(in_strCode) Then Exit Function
    If Len(in_strCode) <> 12 Then Exit Function
    
    For i = 0 To 11
        bytCode(i) = Mid(in_strCode, i + 1, 1)
    Next
    
    strDigit = right(CStr(((bytCode(11) + bytCode(9) + bytCode(7) + bytCode(5) + bytCode(3) + bytCode(1)) * 3) + (bytCode(10) + bytCode(8) + bytCode(6) + bytCode(4) + bytCode(2) + bytCode(0))), 1)
    
    If strDigit = "0" Then
        bytfncCheckDigit_JAN = 0
    Else
        bytfncCheckDigit_JAN = 10 - CByte(strDigit)
    End If
    
    Exit Function
    
Err_bytfncCheckDigit_JAN:
    MsgBox Err.Description
    bytfncCheckDigit_JAN = 99

End Function

Public Function strfncGetVersion() As String
'--------------------------------------------------------------------------------------------------------------------
'�o�[�W�����擾����
'   ���o�[�W�������擾��������Ŗ߂�
'   ���擾�ł��Ȃ��ꍇ�͋󕶎��ŕԂ�

'3.0.0
'   ��Rev�ǉ�
'--------------------------------------------------------------------------------------------------------------------
    Dim objLOCALdb As New cls_LOCALDB
    Dim strRev As String
    
    On Error GoTo Err_strfncGetVersion
    
    If objLOCALdb.ExecSelect("select Version,Rev from T_Version�Ǘ� order by �X�V���� desc ") Then
        
        If Not objLOCALdb.GetRS.EOF Then
            strfncGetVersion = objLOCALdb.GetRS![Version]
            strRev = CStr(objLOCALdb.GetRS![Rev])
        Else
            Err.Raise 9999, , "�o�[�W�����擾�G���[�B���j���[���J�����Ƃ��ł��܂���"
        End If
        
    Else
    
        Err.Raise 9999, , "�o�[�W�����擾�G���[�B���j���[���J�����Ƃ��ł��܂���"
    
    End If
    
    strfncGetVersion = strfncGetVersion & "." & strRev
    
    GoTo Exit_strfncGetVersion

Err_strfncGetVersion:
    strfncGetVersion = ""
    MsgBox Err.Description
    
Exit_strfncGetVersion:
'�N���X�̃C���X�^���X��j��
    Set objLOCALdb = Nothing
    
End Function

Public Function strfncGetVersionOnly() As String
'--------------------------------------------------------------------------------------------------------------------
'�o�[�W����(���r�W��������)�擾����
'   ���o�[�W�������擾��������Ŗ߂�
'   ���擾�ł��Ȃ��ꍇ�͋󕶎��ŕԂ�

'3.0.0 ADD
'--------------------------------------------------------------------------------------------------------------------
    Dim objLOCALdb As New cls_LOCALDB
    
    On Error GoTo Err_strfncGetVersionOnly
    
    If objLOCALdb.ExecSelect("select Version from T_Version�Ǘ� order by �X�V���� desc ") Then
        
        If Not objLOCALdb.GetRS.EOF Then
            strfncGetVersionOnly = objLOCALdb.GetRS![Version]
        Else
            Err.Raise 9999, , "�o�[�W�����擾�G���[�B���j���[���J�����Ƃ��ł��܂���"
        End If
        
    Else
    
        Err.Raise 9999, , "�o�[�W�����擾�G���[�B���j���[���J�����Ƃ��ł��܂���"
    
    End If
    
    GoTo Exit_strfncGetVersionOnly

Err_strfncGetVersionOnly:
    strfncGetVersionOnly = ""
    MsgBox Err.Description
    
Exit_strfncGetVersionOnly:
'�N���X�̃C���X�^���X��j��
    Set objLOCALdb = Nothing
    
End Function

Public Function bolfncOpen_LogOnMenu(in_strMenuName As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'���O�I���ς݊m�F����
'   �����O�I����ʂ�\������
'   ��UserID���󗓂̏ꍇ�̓L�����Z���������ƂɂȂ�
'   :����
'       in_strMenuName  :���j���[��

'   :�߂�l
'       True            :���O�I���ς�
'       False           :�����O�I��
'--------------------------------------------------------------------------------------------------------------------
    bolfncOpen_LogOnMenu = False
    
    On Error GoTo Err_bolfncOpen_LogOnMenu
    
    DoCmd.OpenForm "F_���O�I��", acNormal, , , , acDialog, in_strMenuName
    
    If strUserID <> "" Then
        bolfncOpen_LogOnMenu = True
    End If
    
    GoTo Exit_bolfncOpen_LogOnMenu
    
Err_bolfncOpen_LogOnMenu:
    MsgBox Err.Description
    bolfncOpen_LogOnMenu = False
Exit_bolfncOpen_LogOnMenu:

End Function

Public Function bolfncEnableSystem() As Boolean
'--------------------------------------------------------------------------------------------------------------------
'���������I���m�F
'   ����ԃo�b�`�������I�����Ă��邩�m�F����
'
'   :�߂�l
'       True            :�V�X�e���g�p�\
'       False           :�����������I��
'--------------------------------------------------------------------------------------------------------------------
    Dim objREMOTEdb As New cls_BRAND_MASTER
    
    bolfncEnableSystem = False
    
    On Error GoTo Err_bolfncEnableSystem
    
    If objREMOTEdb.ExecSelect("select �l from T_Control where [key] = 4") Then
        
        If Not objREMOTEdb.GetRS.EOF Then
            If objREMOTEdb.GetRS![�l] <> Format(Date, "yyyymmdd") Then
                Err.Raise 9999, , "AM0:00�`���������I���܂ŃV�X�e���͎g�p�ł��܂���"
            End If
        Else
            Err.Raise 9999, , "�R���g���[���}�X�^�ɃL�[[4]�i�������j�����݂��܂���"
        End If
    Else
        Err.Raise 9999, , "�R���g���[���}�X�^�̎擾�G���["
    
    End If

    bolfncEnableSystem = True
    
    GoTo Exit_bolfncEnableSystem
    
Err_bolfncEnableSystem:
    MsgBox Err.Description
    bolfncEnableSystem = False
    
Exit_bolfncEnableSystem:

    Set objREMOTEdb = Nothing
End Function

Public Function bolfncTextCompare(in_BeforeUpdate As Variant, in_AfterUpdate As Variant, Optional in_CompareMode As Variant) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'�e�L�X�g��r����
'   ��2�̃e�L�X�g���r�������ł����True��Ԃ�
'
'   :����
'       in_BeforeUpdate     :�ύX�O
'       in_AfterUpdate      :�ύX�O
'       in_CompareMode      :��r���[�h
'                            0:�o�C�i�����[�h�i�S�p/���p�A�啶��/�������A�Ђ炪��/�J�^�J�i����ʂ���j�iDefault)
'                            1:�e�L�X�g���[�h�i�S�p/���p�A�啶��/�������A�Ђ炪��/�J�^�J�i����ʂ��Ȃ��j
'                            2:Access�̐ݒ�ɂ�������
'
'   :�߂�l
'       True            :�ύX����
'       False           :�ύX�Ȃ�
'--------------------------------------------------------------------------------------------------------------------
    Dim intComparemode As Byte
    Dim strBefore As String, strAfter As String
    
    On Error GoTo Err_bolfncTextCompare
    
    bolfncTextCompare = False
    
    If IsMissing(in_CompareMode) Then
        intComparemode = vbBinaryCompare
    Else
        intComparemode = in_CompareMode
    End If
    
    strBefore = Nz(in_BeforeUpdate, "")
    strAfter = Nz(in_AfterUpdate, "")
    
    If StrComp(strBefore, strAfter, intComparemode) Then
        '�ύX����
        bolfncTextCompare = True
    End If
        
    Exit Function
    
Err_bolfncTextCompare:
    MsgBox Err.Description, , "�e�L�X�g��r�G���["
    
End Function

Public Sub WindowSize_Restore()
'--------------------------------------------------------------------------------------------------------------------
'�A�v���P�[�V�����E�B���h�E�T�C�Y��W���ɖ߂�
'   Ver 1.01.1 K.Asayama ADD 20150910
'--------------------------------------------------------------------------------------------------------------------

    Dim lngRt As Long
    lngRt = ShowWindow(Application.hWndAccessApp, 1)
    
End Sub

Public Sub WindowSize_Minimize()
'--------------------------------------------------------------------------------------------------------------------
'�A�v���P�[�V�����E�B���h�E�T�C�Y���ŏ�������
'   Ver 1.01.1 K.Asayama ADD 20150910
'--------------------------------------------------------------------------------------------------------------------
    Dim lngRt As Long
    lngRt = ShowWindow(Application.hWndAccessApp, 2)
    
End Sub

Public Function fncMenuCall(ByVal strMenuName As String)
'--------------------------------------------------------------------------------------------------------------------
'���j���[���J��
'   Ver 1.01.1 K.Asayama ADD 20150910
'   Ver 1.01.* K.Asayama Change 201510**
'--------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo Err_fncMenuCall
''Window�ŏ���
    WindowSize_Minimize

'���j���[�R�[��
   'DoCmd.OpenForm strMenuName, acNormal, , , , acDialog
   DoCmd.OpenForm strMenuName, acNormal, , , , acWindowNormal
   
'Window�����ɖ߂�
    'WindowSize_Restore
    
    Exit Function
    
Err_fncMenuCall:
    Select Case Err.Number
        Case 2501 '�L�����Z���I���̏ꍇ
        
        Case Else
            MsgBox Err.Number
    End Select
    'Window�����ɖ߂�
    WindowSize_Restore
End Function

Public Function Form_IsLoaded(ByVal in_FormName As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'�t�H�[�����J���Ă��邩�m�F����
'   Ver 1.01.* K.Asayama ADD 201510**
'--------------------------------------------------------------------------------------------------------------------
    If CurrentProject.AllForms(in_FormName).IsLoaded Then
    
        Form_IsLoaded = True
    
    Else
    
        Form_IsLoaded = False
            
    End If

End Function

Public Function MainMenu_ReVisible()
'--------------------------------------------------------------------------------------------------------------------
'���C�����j���[���ĕ\������
'   Ver 1.01.* K.Asayama ADD 201510**
'--------------------------------------------------------------------------------------------------------------------
    If Form_IsLoaded("F_�H���Ǘ����j���[") Then
    
        Form_F_�H���Ǘ����j���[.Visible = True
    
    End If

End Function

Public Function TwipsToCm(ByVal value As Double) As Double
'--------------------------------------------------------------------------------------------------------------------
'   1 cm = 567 twips
'   1�C���` = 1440 twips = 2.54cm = 72 �|�C���g
'
'   twips ���� cm �ɕϊ�
'
'   :����
'       value               :twips�ł̒���
'
'   :�߂�l
'       Cm                  :�Z���`���[�g��
'--------------------------------------------------------------------------------------------------------------------

    TwipsToCm = value / 567

End Function

Public Function CmToTwips(ByVal value As Double) As Double
'--------------------------------------------------------------------------------------------------------------------
'
'   cm ���� twips �ɕϊ�
'
'   :����
'       value               :cm�ł̒���
'
'   :�߂�l
'       twips               :Twips
'--------------------------------------------------------------------------------------------------------------------
    CmToTwips = value * 567
    
End Function

Public Sub subAllbutton_Enabled(ByVal in_FormName As String, ByVal in_bolTF As Boolean)
'--------------------------------------------------------------------------------------------------------------------
'
'   �t�H�[���̃{�^���̎g�p�\�A�s�\�ꊇ�ύX
'
'   :����
'       in_FormName         :�t�H�[����
'       in_bolTF            :�g�p�\�iTrue�j/�s�\�iFalse�j
'
'--------------------------------------------------------------------------------------------------------------------
    Dim ctl As Access.Control
    Dim i As Byte
    i = 0
    
    On Error GoTo Err_subAllbutton_Enabled
    
    '���̃t�H�[�����̂��ׂẴR���g���[��������
    For Each ctl In Forms(in_FormName).Controls
        With ctl
            If .ControlType = acCommandButton Then
                   
                ctl.Enabled = in_bolTF

            End If
        End With
    Next ctl
        
    GoTo Exit_subAllbutton_Enabled
    
Err_subAllbutton_Enabled:

Exit_subAllbutton_Enabled:
    Set ctl = Nothing
End Sub

Public Function bolfncinputDate(ByVal in_MidashiText As String, ByRef out_Date As Variant) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'
'   �ėp���t���̓t�H�[���\��
'
'   :����
'       in_MidashiText      :���o�����i8�������炢���K���j
'       out_Date            :���͓��t
'
'   :�߂�l
'                           :���t���͍ς݁iTrue�j/�L�����Z���iFalse�j
'--------------------------------------------------------------------------------------------------------------------
Dim objLOCALdb As New cls_LOCALDB
Dim strErrMsg As String

On Error GoTo Err_bolfncinputDate

out_Date = Null

If Not objLOCALdb.ExecSQL("delete from WK_�Ώۓ��t", strErrMsg) Then
    Err.Raise 9999, , strErrMsg
End If

DoCmd.OpenForm "F_�ėp���t����", acNormal, , , , acDialog, in_MidashiText

If Not objLOCALdb.ExecSelect("select date1 from WK_�Ώۓ��t") Then
    Err.Raise 9999, , "���t�ǂݍ��݃G���["
Else
    If Not objLOCALdb.GetRS.EOF Then
        out_Date = objLOCALdb.GetRS!Date1
    End If
End If

If IsNull(out_Date) Then
    Err.Raise 9998, , "���t�̓��͂��L�����Z������܂���"
End If

bolfncinputDate = True

GoTo Exit_bolfncinputDate

Err_bolfncinputDate:
    bolfncinputDate = False
    MsgBox Err.Description
    
Exit_bolfncinputDate:
    Set objLOCALdb = Nothing
    
End Function

Public Sub subAllbutton_noPrintable(ByVal in_FormName As String)
'--------------------------------------------------------------------------------------------------------------------
'
'   �t�H�[���̃{�^���̈���s��
'
'   :����
'       in_FormName         :�t�H�[����
'
'1.10.9 K.Asayama ADD
'--------------------------------------------------------------------------------------------------------------------
    Dim ctl As Access.Control
    Dim i As Byte
    i = 0
    
    On Error GoTo Err_subAllbutton_noPrintable
    
    '���̃t�H�[�����̂��ׂẴR���g���[��������
    For Each ctl In Forms(in_FormName).Controls
        With ctl
            If .ControlType = acCommandButton Then
                   
                ctl.DisplayWhen = 2

            End If
        End With
    Next ctl
        
    GoTo Exit_subAllbutton_noPrintable
    
Err_subAllbutton_noPrintable:

Exit_subAllbutton_noPrintable:
    Set ctl = Nothing
End Sub

Public Sub subScreenShot_AllArea()
'--------------------------------------------------------------------------------------------------------------------
'
'   �S�X�N���[���̃X�N���[���V���b�g�擾
'
'
'1.10.14 K.Asayama ADD
'--------------------------------------------------------------------------------------------------------------------
    keybd_event vbKeySnapshot, 0&, &H1, 0&
    keybd_event vbKeySnapshot, 0&, &H1 Or &H2, 0&
    
End Sub

Public Sub subScreenShot_ActiveArea()
'--------------------------------------------------------------------------------------------------------------------
'
'   �A�N�e�B�u�X�N���[���̃X�N���[���V���b�g�擾
'
'
'1.10.14 K.Asayama ADD
'--------------------------------------------------------------------------------------------------------------------
    keybd_event &HA4, 0&, &H1, 0&
    keybd_event vbKeySnapshot, 0&, &H1, 0&
    keybd_event vbKeySnapshot, 0&, &H1 Or &H2, 0&
    keybd_event &HA4, 0&, &H1 Or &H2, 0&
End Sub

Public Function bolfncinputDate_FromTo(ByVal in_MidashiText As String, ByVal in_DateDetail As String, ByRef out_DateFrom As Variant, ByRef out_DateTo As Variant) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'
'   �ėp���t���̓t�H�[��(FromTo)�\��
'
'   :����
'       in_MidashiText      :���o�����i8�������炢���K���j
'       in_DateDetail       :���t�ڍׁi�������A�[�i������ʕ\���p�j
'       out_DateFrom        :���͓��t(From)
'       out_DateTo          :���͓��t(To)
'
'   :�߂�l
'                           :���t���͍ς݁iTrue�j/�L�����Z���iFalse�j
'1.10.15 ADD
'--------------------------------------------------------------------------------------------------------------------
Dim objLOCALdb As New cls_LOCALDB
Dim strErrMsg As String

On Error GoTo Err_bolfncinputDate_FromTo

out_DateFrom = Null
out_DateTo = Null

If Not objLOCALdb.ExecSQL("delete from WK_�Ώۓ��t", strErrMsg) Then
    Err.Raise 9999, , strErrMsg
End If

DoCmd.OpenForm "F_�ėp���t����_FromTo", acNormal, , , , acDialog, in_MidashiText & vbTab & in_DateDetail

If Not objLOCALdb.ExecSelect("select date1,date2 from WK_�Ώۓ��t") Then
    Err.Raise 9999, , "���t�ǂݍ��݃G���["
Else
    If Not objLOCALdb.GetRS.EOF Then
        out_DateFrom = objLOCALdb.GetRS!Date1
        out_DateTo = objLOCALdb.GetRS!Date2
    End If
End If

If IsNull(out_DateFrom) Or IsNull(out_DateTo) Then
    Err.Raise 9998, , "���t�̓��͂��L�����Z������܂���"
End If

bolfncinputDate_FromTo = True

GoTo Exit_bolfncinputDate_FromTo

Err_bolfncinputDate_FromTo:
    bolfncinputDate_FromTo = False
    MsgBox Err.Description
    
Exit_bolfncinputDate_FromTo:
    Set objLOCALdb = Nothing
    
End Function

Public Function bolfncReport(in_ReportName As String, in_Preview As Boolean, Optional in_Message As Boolean) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'
'   ���|�[�g�o��
'
'   :����
'       in_ReportName       :���|�[�g��
'       in_Preview          :True���v���r���[ False���v�����^�o��
'       in_Message(Option)  :True���f�[�^0���̍ۃ��b�Z�[�W���o�͂���   False�����Ȃ�
'
'   :�߂�l
'       True            :����
'       False           :���s
'
'   1.11.0 ADD
'--------------------------------------------------------------------------------------------------------------------
    Dim bytPrintmode As Byte

    On Error GoTo Err_bolfncReport

    If in_Preview Then
        bytPrintmode = 2
    Else
        bytPrintmode = 0
    End If
    
    DoCmd.OpenReport in_ReportName, bytPrintmode
    
    bolfncReport = True
    
    Exit Function

Err_bolfncReport:
    
    If Err.Number = 2501 Then
        If in_Message Then
            MsgBox in_ReportName & vbCrLf & "�f�[�^������܂���"
        End If
        Resume Next
    Else
        MsgBox Err.Description
    End If
    
    bolfncReport = False
    

End Function

Public Function strfncTextFileToString(strFileFullpath As String) As String
'--------------------------------------------------------------------------------------------------------------------
'Text����String�փt���R�s�[
'   ���t�@�C���i�t���p�X�j��ǂݍ���ł��̂܂�String�ϐ��ɃC���|�[�g
'
'   :����
'       strFileFullpath     �t�@�C�����i�t���p�X�j

'1.11.1 ADD
'--------------------------------------------------------------------------------------------------------------------
    Dim strTxt As String
    
    strfncTextFileToString = ""
    strTxt = ""
    
    On Error GoTo Err_strfncTextFileToString
    
    If Dir(strFileFullpath) <> "" Then
        With CreateObject("Scripting.FileSystemObject")
            With .GetFile(strFileFullpath).OpenAsTextStream
                strTxt = .ReadAll
                .Close
            End With
        End With
        
    Else
        Err.Raise 9999, , "�ϊ��p�t�@�C�������݂��܂���B�Ǘ��҂ɘA�����Ă�������"
    End If
    
    strfncTextFileToString = strTxt
    
    Exit Function
    
Err_strfncTextFileToString:
    Close
    MsgBox Err.Description
    
End Function

Public Sub subAllOption_Enabled(ByVal in_FormName As String, ByVal in_bolTF As Boolean)
'--------------------------------------------------------------------------------------------------------------------
'
'   �t�H�[����Option�R���g���[���̎g�p�\�A�s�\�ꊇ�ύX
'
'   :����
'       in_FormName         :�t�H�[����
'       in_bolTF            :�g�p�\�iTrue�j/�s�\�iFalse�j
'
'1.11.1 ADD
'1.12.0
'   ��  �R���{�{�b�N�X�ƃ`�F�b�N�{�b�N�X�ǉ�
'--------------------------------------------------------------------------------------------------------------------
    Dim ctl As Access.Control
    Dim i As Byte
    i = 0
    
    On Error GoTo Err_subAllOption_Enabled_Enabled
    
    '���̃t�H�[�����̂��ׂẴR���g���[��������
    For Each ctl In Forms(in_FormName).Controls
        With ctl
            If .ControlType = acOptionGroup Or .ControlType = acComboBox Or .ControlType = acCheckBox Then
                   
                ctl.Enabled = in_bolTF

            End If
        End With
    Next ctl
        
    GoTo Exit_subAllOption_Enabled_Enabled
    
Err_subAllOption_Enabled_Enabled:

Exit_subAllOption_Enabled_Enabled:
    Set ctl = Nothing
End Sub

Public Sub subAllTextBox_Enabled(ByVal in_FormName As String, ByVal in_bolTF As Boolean)
'--------------------------------------------------------------------------------------------------------------------
'
'   �t�H�[����TextBox�̎g�p�\�A�s�\�ꊇ�ύX
'
'   :����
'       in_FormName         :�t�H�[����
'       in_bolTF            :�g�p�\�iTrue�j/�s�\�iFalse�j
'
'2.0.0 ADD
'--------------------------------------------------------------------------------------------------------------------
    Dim ctl As Access.Control
    Dim i As Byte
    i = 0
    
    On Error GoTo Err_subAllOption_Enabled_Enabled
    
    '���̃t�H�[�����̂��ׂẴR���g���[��������
    For Each ctl In Forms(in_FormName).Controls
        With ctl
            If .ControlType = acTextBox Then
                   
                ctl.Enabled = in_bolTF

            End If
        End With
    Next ctl
        
    GoTo Exit_subAllOption_Enabled_Enabled
    
Err_subAllOption_Enabled_Enabled:

Exit_subAllOption_Enabled_Enabled:
    Set ctl = Nothing
End Sub
Public Function Report_IsLoaded(ByVal in_ReportName As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'���|�[�g���J���Ă��邩�m�F����
'   Ver 1.11.2 ADD
'--------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo Err_Report_IsLoaded
    
    If CurrentProject.AllReports(in_ReportName).IsLoaded Then
    
        Report_IsLoaded = True
    
    Else
    
        Report_IsLoaded = False
            
    End If
    
    Exit Function
    
Err_Report_IsLoaded:
'    If Err.Number = 2467 Then
'        Resume Next
'    End If
    Report_IsLoaded = False
    
End Function

Public Function bolfncinFlieGet(ByVal in_KeyName As String, ByRef out_iniData As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'ini�t�@�C��������w��̃L�[�𒊏o

'   :����
'       in_KeyName             :ini�t�@�C���L�[��
'       out_iniData            :�ϐ���
'
'   :�߂�l
'       True            :����
'       False           :���s

'   Ver 1.11.2 ADD
'   1.11.3  Change �e�X�g�����ʒǉ��i���[�J��(C:\kamiya_Brand��ini�t�@�C��������ꍇ�͂������D�悷��
'   1.12.3
'           ���T�[�o�p�X���ʉ�
'2.13.0
'   �����[�J���i�e�X�g�p�j���g�p����ꍇ�̓��b�Z�[�W��\������
'--------------------------------------------------------------------------------------------------------------------
       
    Const strIniPath As String = conServerPath & "\�����Ǘ��V�X�e��.ini"
    
    Const strTestPath As String = "C:\Kamiya_Brand\�����Ǘ��V�X�e��.ini"
    
    Dim strBuf As String
    Dim varText As Variant
    Dim varPath As Variant
    
    Dim strInputPath As String
    
    Dim i As Integer
    
    bolfncinFlieGet = False
    
    On Error GoTo Err_bolfncinFlieGet
    
    varPath = Null
    
    'ini�t�@�C�������[�J���ɂ���ꍇ�͂������D��
    If Dir(strTestPath) <> "" Then
        MsgBox "�e�X�g�p��Ini�t�@�C�����g�p���Ă��܂�", vbInformation
        strInputPath = strTestPath
    Else
        strInputPath = strIniPath
    End If
    
    'ini�t�@�C�����o�b�t�@�ɓǂݍ���
    With CreateObject("Scripting.FileSystemObject")
        With .GetFile(strInputPath).OpenAsTextStream
            strBuf = .ReadAll
            .Close
        End With
    End With
    
    varText = Split(strBuf, vbCrLf)
    
    If VarType(varText) > vbArray Then
        For i = LBound(varText) To UBound(varText)
            If varText(i) Like in_KeyName & vbTab & "*" Then
                varPath = Split(varText(i), vbTab)
                Exit For
            End If
        Next
    Else
        'Debug.Print varText
    End If
        
    If VarType(varPath) > vbArray Then
       out_iniData = varPath(1)
       bolfncinFlieGet = True
    End If
    
    GoTo Exit_bolfncinFlieGet
    
Err_bolfncinFlieGet:
    MsgBox Err.Description
    Close
Exit_bolfncinFlieGet:

End Function

Public Sub OpenExplorer(in_Path As String)
'--------------------------------------------------------------------------------------------------------------------
'�w�肵���t�@�C���ʒu���G�N�X�v���[���ŊJ��

'   :����
'       in_Path             :�t���p�X��

'1.12.0 ADD
'--------------------------------------------------------------------------------------------------------------------
    Call Shell("explorer.exe /select," & in_Path, vbNormalFocus)
End Sub

Public Function NullIFNothing(ByVal InputData As Variant) As Variant
'***********************************************************
'0�܂��͋󗓁iEmpty)�̎�Null�ɒu��������
'
'   �l��0���͋󗓂̎�Null�ɒu��������
'       ���DB�ւ̒l�o�^��e�L�X�g�{�b�N�X���̃I�u�W�F�N�g��
'       �l��߂��Ƃ��Ɏg�p
'
'   �߂�l�iVariant�^)
'
'2.1.0 ADD
'***********************************************************
    Select Case VarType(InputData)
    
        Case Is <= 1 '�� ���� Null
            
            NullIFNothing = Null
            
        Case Is <= 6 '���l�n
            
            If InputData = 0 Then
                NullIFNothing = Null
            Else
                NullIFNothing = InputData
            End If
        
        Case 8  'String�^
            
            If InputData = "" Then
                NullIFNothing = Null
            ElseIf IsNumeric(InputData) Then
                If CDbl(InputData) = 0 Then
                    NullIFNothing = Null
                Else
                    NullIFNothing = InputData
                End If
            Else
                NullIFNothing = InputData
            End If
            
        Case Else
            
            NullIFNothing = InputData
    
    End Select
    
End Function

Public Function RPADB(strValue As String, strCharactor As String, intKeta As Integer) As String
'***********************************************************
'RPAD�i�����̌����֕����𖄂߂�j�̃o�C�g����
'
'   ����
'       :strValue               ���ɂȂ�l
'       :strCharactor           ���߂镶���i0�̏ꍇ�E����0�Ŗ��߂�j
'       :intKeta                ���߂���̌���
'
'   �߂�l�iString�^)
'
'2.1.0 ADD
'***********************************************************
    'string�̉E���Ɏw�肳�ꂽ�����𕶎������ɂȂ�悤���߂�i�o�C�g���j
    Dim strSJIS As String
    
    'Unicode��Shift-JIS�ɕϊ�
    strSJIS = StrConv(strValue, vbFromUnicode)
    RPADB = strValue & String(intKeta - LenB(strSJIS), strCharactor)
        
End Function

Public Function fncFileSelector(ByVal inFolder As String, ByVal inShurui As Integer) As String
'***********************************************************
'�_�C�A���O��\�����ăt�@�C����I������

'
'   ����
'       :inFolder               �����p�X
'       :inShurui               �t�@�C����� 0:xlsx

'   �߂�l�iString�^)
'
'2.8.0 ADD
'***********************************************************

    Dim strFile As String
    Dim intRet As Integer
    
    On Error GoTo Err_fncFileSelector
    
    fncFileSelector = ""
    With Application.FileDialog(msoFileDialogOpen)
    
        '�_�C�A���O�̃^�C�g����ݒ�
        '.Title = "�_�C�A���O"
        
        '�t�@�C���̎�ނ�ݒ�
        .Filters.Clear
        
        Select Case inShurui
            Case 0
                .Filters.Add "Microsoft Office Excel�t�@�C��", "*.xlsx"
        End Select
        
        .FilterIndex = 1
        
        '�����t�@�C���I���������Ȃ�
        .AllowMultiSelect = False
        '�����p�X��ݒ�
        
        .InitialFileName = inFolder
        '�_�C�A���O��\��
        intRet = .Show
        
        If intRet <> 0 Then
          '�t�@�C�����I�����ꂽ�Ƃ�
          '���̃t���p�X��Ԃ�l�ɐݒ�
          strFile = Trim(.SelectedItems.Item(1))
        Else
          '�t�@�C�����I������Ȃ���΃u�����N
          strFile = ""
        End If
        
    End With

    If strFile <> "" Then
        fncFileSelector = strFile
    End If
    
    Exit Function
    
Err_fncFileSelector:
    MsgBox Err.Description
    
    
End Function

Public Function isTableExist(ByVal strTableName As String) As Boolean
'   *************************************************************
'   isTableExist
'   �e�[�u�����݊m�F

'
'   �߂�l:Boolean
'       ��True              �e�[�u���L
'       ��False             �e�[�u������
'
'    Input����
'       strTableName        �e�[�u����

'3.0.0 ADD
'   *************************************************************

    Dim daoDB As DAO.Database
    Dim daoTableDef As DAO.TableDef
    Set daoDB = CurrentDb
    
    isTableExist = False
    
    On Error GoTo Err_isTableExist
    
    For Each daoTableDef In CurrentDb.TableDefs
        If daoTableDef.Name = strTableName Then
            isTableExist = True
            Exit For
        End If
    Next
    
    GoTo Exit_isTableExist
    
Err_isTableExist:

Exit_isTableExist:
    Set daoTableDef = Nothing
    Set daoDB = Nothing
End Function

Public Function UrlEncodeUtf8(ByVal strSource As String) As String
'   *************************************************************
'   UrlEncodeUtf8
'   �������UTF8�G���R�[�h���Ė߂�

'
'   �߂�l:String
'       ���G���R�[�h��̕�����
'
'    Input����
'       strSource        �G���R�[�h�O�̕�����

'3.0.0 ADD
'   *************************************************************
    Dim objSC As Object
    Set objSC = CreateObject("ScriptControl")
    objSC.Language = "Jscript"
    UrlEncodeUtf8 = objSC.CodeObject.encodeURIComponent(strSource)
    Set objSC = Nothing
End Function

Public Function bolint40mmOrder() As Integer
'--------------------------------------------------------------------------------------------------------------------
'40mm�\�[�g��
'   ���\�[�g����40mm��D�悷�邩36mm��D�悷�邩�m�F
'
'   :�߂�l
'       0       :40mm,36mm�̏�
'       1       :36mm,40mm�̏�
'--------------------------------------------------------------------------------------------------------------------
    Dim objREMOTEdb As New cls_BRAND_MASTER
    
    bolint40mmOrder = 1
    
    On Error GoTo Err_bolint40mmOrder
    
    If objREMOTEdb.ExecSelect("select �l from T_Control where [key] = 14") Then
        
        If Not objREMOTEdb.GetRS.EOF Then
            If IsNull(objREMOTEdb.GetRS![�l]) Then
                Err.Raise 9999, , "40mm�\�[�g�����擾�ł��Ȃ��̂�36mm��40mm�̏��Ń\�[�g���܂�"
                If Not IsNumeric(objREMOTEdb.GetRS![�l]) Then
                    Err.Raise 9999, , "�R���g���[���}�X�^�̃\�[�g���̒l���ُ�ł�"
                End If
            End If
        Else
            Err.Raise 9999, , "�R���g���[���}�X�^�ɃL�[[14]�i40mm�\�[�g���j�����݂��܂���"
        End If
    Else
        Err.Raise 9999, , "�R���g���[���}�X�^�̎擾�G���["
    
    End If

    bolint40mmOrder = objREMOTEdb.GetRS![�l]
    
    GoTo Exit_bolint40mmOrder
    
Err_bolint40mmOrder:
    MsgBox Err.Description
    
Exit_bolint40mmOrder:

    Set objREMOTEdb = Nothing
End Function

Public Function fncUserGroup_Belongs(in_strGroup As String) As Boolean
'   *************************************************************
'   ���[�U�[�O���[�v�����m�F
'       ���O�C�����[�U�[��AD�A�J�E���g��
'       �����̃O���[�v�ɓo�^����Ă����True��Ԃ�

'       ���h���C���ɎQ�����Ă��Ȃ��ꍇ�APC���l�b�g���[�N�ɐڑ�����Ă��Ȃ��ꍇ
'       �@���s���G���[�ɂȂ�

'   �߂�l:Boolean
'       ��True              �O���[�v�ɏ������Ă���
'       ��False             �O���[�v�ɏ������Ă��Ȃ�
'
'   ����
'       in_strGroup        ActiveDirectory���[�U�[�O���[�v��

'  3.0.0 K.Asayama ADD
'   *************************************************************

    Dim objADSysInfo As Object
    Dim objUser As Object
    Dim objGroup As Object
    Dim varGroup As Variant
    Dim strUser As String
    
    
    Set objADSysInfo = CreateObject("ADSystemInfo")
    strUser = objADSysInfo.UserName
    Set objUser = GetObject("LDAP://" & strUser)
    
    On Error GoTo Err_fncUserGroup_Belongs
        
    strUser = objADSysInfo.UserName

    'memberOf�����͕�������ꍇ�̂݃I�u�W�F�N�g�z��ɂȂ邽��
    '�ȉ��Ŕz�񂩊m�F���ď����𕪊�

    If IsArray(objUser.memberOf) Then
    
        For Each varGroup In objUser.memberOf
            Set objGroup = GetObject("LDAP://" & varGroup)
            If objGroup.cn = in_strGroup Then
                fncUserGroup_Belongs = True
                Exit For
            End If
        Next
        
        Set objGroup = Nothing

    Else
        '���O���[�v�ɂP�������ĂȂ���GetObject�ŃG���[�ɂȂ邽�߃g���b�v
        On Error Resume Next
        
        Set objGroup = GetObject("LDAP://" & objUser.memberOf)
        
        If Err.Number = 0 Then
            
            If objGroup.cn = in_strGroup Then
                fncUserGroup_Belongs = True
            End If
             
            Set objGroup = Nothing
        End If

    End If
    
    GoTo Exit_fncUserGroup_Belongs
    
Err_fncUserGroup_Belongs:
    '�G���[�̏ꍇ��False��Ԃ�
    fncUserGroup_Belongs = False

Exit_fncUserGroup_Belongs:
    Set objADSysInfo = Nothing
    Set objUser = Nothing
    
End Function