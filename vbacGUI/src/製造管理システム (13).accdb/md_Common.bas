Option Compare Database
Option Explicit
'--------------------------------------------------------------------------------------------------------------------
'���ʕϐ�
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

'1.10.6 K.Asayama 20151211 �ǉ�
'SxL���[�J���R�s�[,�J�����_�[�R�s�[
Public bolSxLCopy As Boolean
Public bolCalendarCopy As Boolean


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
'--------------------------------------------------------------------------------------------------------------------
    Dim objLocalDB As New cls_LOCALDB

    On Error GoTo Err_strfncGetVersion
    
    If objLocalDB.ExecSelect("select Version from T_Version�Ǘ� order by �X�V���� desc ") Then
        
        If Not objLocalDB.GetRS.EOF Then
            strfncGetVersion = objLocalDB.GetRS![Version]
        Else
            Err.Raise 9999, , "�o�[�W�����擾�G���[�B���j���[���J�����Ƃ��ł��܂���"
        End If
        
    Else
    
        Err.Raise 9999, , "�o�[�W�����擾�G���[�B���j���[���J�����Ƃ��ł��܂���"
    
    End If
    
    GoTo Exit_strfncGetVersion

Err_strfncGetVersion:
    strfncGetVersion = ""
    MsgBox Err.Description
    
Exit_strfncGetVersion:
'�N���X�̃C���X�^���X��j��
    Set objLocalDB = Nothing
    
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
    Dim objREMOTEDB As New cls_BRAND_MASTER
    
    bolfncEnableSystem = False
    
    On Error GoTo Err_bolfncEnableSystem
    
    If objREMOTEDB.ExecSelect("select �l from T_Control where [key] = 4") Then
        
        If Not objREMOTEDB.GetRS.EOF Then
            If objREMOTEDB.GetRS![�l] <> Format(Date, "yyyymmdd") Then
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

    Set objREMOTEDB = Nothing
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
Dim objLocalDB As New cls_LOCALDB
Dim strErrMsg As String

On Error GoTo Err_bolfncinputDate

out_Date = Null

If Not objLocalDB.ExecSQL("delete from WK_�Ώۓ��t", strErrMsg) Then
    Err.Raise 9999, , strErrMsg
End If

DoCmd.OpenForm "F_�ėp���t����", acNormal, , , , acDialog, in_MidashiText

If Not objLocalDB.ExecSelect("select date1 from WK_�Ώۓ��t") Then
    Err.Raise 9999, , "���t�ǂݍ��݃G���["
Else
    If Not objLocalDB.GetRS.EOF Then
        out_Date = objLocalDB.GetRS!Date1
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
    Set objLocalDB = Nothing
    
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
Dim objLocalDB As New cls_LOCALDB
Dim strErrMsg As String

On Error GoTo Err_bolfncinputDate_FromTo

out_DateFrom = Null
out_DateTo = Null

If Not objLocalDB.ExecSQL("delete from WK_�Ώۓ��t", strErrMsg) Then
    Err.Raise 9999, , strErrMsg
End If

DoCmd.OpenForm "F_�ėp���t����_FromTo", acNormal, , , , acDialog, in_MidashiText & vbTab & in_DateDetail

If Not objLocalDB.ExecSelect("select date1,date2 from WK_�Ώۓ��t") Then
    Err.Raise 9999, , "���t�ǂݍ��݃G���["
Else
    If Not objLocalDB.GetRS.EOF Then
        out_DateFrom = objLocalDB.GetRS!Date1
        out_DateTo = objLocalDB.GetRS!Date2
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
    Set objLocalDB = Nothing
    
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
'--------------------------------------------------------------------------------------------------------------------
    Dim ctl As Access.Control
    Dim i As Byte
    i = 0
    
    On Error GoTo Err_subAllOption_Enabled_Enabled
    
    '���̃t�H�[�����̂��ׂẴR���g���[��������
    For Each ctl In Forms(in_FormName).Controls
        With ctl
            If .ControlType = acOptionGroup Then
                   
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
'--------------------------------------------------------------------------------------------------------------------
       
    Const strIniPath As String = "\\db\Prog\�����Ǘ��V�X�e��\�����Ǘ��V�X�e��.ini"
    
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