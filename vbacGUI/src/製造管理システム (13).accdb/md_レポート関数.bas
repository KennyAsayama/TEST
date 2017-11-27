Option Compare Database
Option Explicit
'2.1.0 ADD

Public Sub AutoFontSize(ByRef Ctr As Control, ByVal IniFontSize As Integer, Optional ByVal MinFontSize = 8)
'   *************************************************************
'   AutoFontSize
'    '�u�����h�V�X�e������R�s�[
'    '���|�[�g���̃e�L�X�g�̃t�H���g�T�C�Y��
'    '�e�L�X�g�{�b�N�X�A���x���̕��ɍ��킹��

'    Input����
'           Ctr         : �R���g���[��
'           IniFontSize : �t�H���g�T�C�Y�����l
'           MinFontSize : �ŏ��t�H���g�T�C�Y�i�f�t�H���g�l:8�j

'   *************************************************************

    'Const MinFontSize = 4 '�ŏ��̃t�H���g�T�C�Y
    Const d = 40 '���܂����܂炸�ɉ��s����Ă��܂��ꍇ�͂����̐��l�𑝂₷
    Dim rpt As Report, str As String, W As Long
    Dim arStr, i As Integer, H As Long
    Set rpt = CodeContextObject
    
With rpt
    If Ctr.ControlType = acTextBox Then
        str = Nz(Ctr.Text, "")
    ElseIf Ctr.ControlType = acLabel Then
        str = Ctr.Caption
    Else
        Exit Sub
    End If
    If str = "" Then Exit Sub
    
    .FontName = Ctr.FontName
    If Ctr.Vertical Then
        W = Ctr.Height - d
        H = Ctr.Width - d
        If InStr(1, .FontName, "@") = 0 Then
            .FontName = "@" & .FontName
        Else
            .FontName = Mid(.FontName, 2)
        End If
    Else
        W = Ctr.Width - d
        H = Ctr.Height - d
    End If
    
    arStr = Split(str, vbCrLf, -1, vbBinaryCompare)
    str = arStr(0)
    For i = 1 To UBound(arStr)
        If .TextWidth(arStr(i)) > .TextWidth(str) Then str = arStr(i)
    Next
    
    .ScaleMode = 1
    If Ctr.FontBold = 1 Then .FontBold = True
    .FontSize = IniFontSize
    Do Until rpt.FontSize = MinFontSize
        If W > .TextWidth(str) Then
            Exit Do
        End If
        .FontSize = .FontSize - 1
    Loop
    
    Do Until rpt.FontSize = MinFontSize
        If H > .TextHeight("A") * (UBound(arStr) + 1) _
            + Ctr.LineSpacing * UBound(arStr) Then
            Exit Do
        End If
        .FontSize = .FontSize - 1
    Loop
    Ctr.FontSize = .FontSize
End With

End Sub

Public Function fncintReport_PrintOut(ByVal strReportName As String, ByVal blnPreView As Boolean, ByVal varWhereCondition As Variant, Optional ByVal intCopies As Integer = 1, Optional bolVisible As Boolean = True) As Integer
'   *************************************************************
'   fncintReport_PrintOut

'   Access���|�[�g�o�͊֐�

'    Input����
'           strReportName       : ���|�[�g��
'           blnPreView          : preview���[�h
'           varWhereCondition   : �p�����[�^
'           intCopies(Option)   : ��������i�f�t�H���g�l:1�j
'           bolVisible(Option)  : True�����@Flse���s���i�f�t�H���g�l:True�j

'   *************************************************************
    
    Dim intHiddenMode As Integer
    Dim intPreviewMode As Integer
    
    On Error GoTo Err_fncintReport_PrintOut
    
    If blnPreView Then
        intPreviewMode = acViewPreview
    Else
        intPreviewMode = acViewNormal
    End If
    
    If bolVisible Then
        intHiddenMode = acWindowNormal
    Else
        intHiddenMode = acHidden
    End If
    
    If IsNull(varWhereCondition) Then
        DoCmd.OpenReport strReportName, intPreviewMode, , , intHiddenMode
    Else
        DoCmd.OpenReport strReportName, intPreviewMode, , varWhereCondition, intHiddenMode
    End If
        
    Reports(strReportName).Printer.Copies = intCopies '�����w��
    
    fncintReport_PrintOut = Reports(strReportName).Pages
    
    Exit Function
    
Err_fncintReport_PrintOut:
        Select Case Err.Number
            Case 2501 'Nodata
                fncintReport_PrintOut = 0
            Case Else
                MsgBox Err.Number & vbCrLf & Err.Description
                fncintReport_PrintOut = -1
        End Select
    
End Function

Public Sub Close_AllReports()
'   *************************************************************
'   �S�Ẵ��|�[�g�����

'   *************************************************************
        
    Dim i As Integer
    
    On Error Resume Next
    
    For i = Reports.Count - 1 To 0 Step -1
        DoCmd.Close acReport, Reports(i).Name
    Next i
  
End Sub

Public Sub Visible_ALLReports()
'   *************************************************************
'   �S�Ẵ��|�[�g����������

'   *************************************************************
        
    Dim i As Integer
    
    On Error Resume Next
    
    For i = Reports.Count - 1 To 0 Step -1
        Reports(i).Visible = True
    Next i
  
End Sub