Option Compare Database
Option Explicit '1.10.15 ADD

Public Sub exp_EXCEL(strSQL As String, Optional boolFilter As Boolean, Optional strMIDASHI As String)
'--------------------------------------------------------------------------------------------------------------------
'EXCEL�G�N�X�|�[�g
'   ��Input��SQL���G�N�Z���̐V�K�u�b�N�ɏo�͂���
'
'   :����
'       strSQL          SQL
'       boolFilter      True�̏ꍇ�͊J�n�s�ɃI�[�g�t�B���^���|����
'       strMIDASHI      True�̏ꍇ��1�s�ڂɌ��o����\������

'   1.10.9 K.Asayama Bug Fix
'   1.10.14 K.Asayama �X�N���[���V���b�g�\��t���p�T�u���[�`���ǉ�
'--------------------------------------------------------------------------------------------------------------------
'

    Dim objREMOTEDB As New cls_BRAND_MASTER
    
    Dim objApp As Object 'Excel
    
    Dim rsADO As New ADODB.Recordset '1.10.15
    
    Dim xlsBookName As String
    
    Dim i, j As Integer
    
    Dim intCount As Integer
    '---------------------------------------
    ' LineStyle
    '---------------------------------------
    Const xlContinuous   As Integer = 1
    Const xlDashDot      As Integer = 4
    Const xlDashDotDot   As Integer = 5
    Const xlSlantDashDot As Integer = 13
    Const xlDash         As Integer = -4115
    Const xlDot          As Integer = -4118
    Const xlDouble       As Integer = -4119
    Const xlLineStyleNone As Integer = -4142
    '---------------------------------------

    '---------------------------------------
    ' Borders
    '---------------------------------------
    Const xlDiagonalDown  As Integer = 5
    Const xlDiagonalUp    As Integer = 6
    Const xlEdgeLeft      As Integer = 7
    Const xlEdgeTop       As Integer = 8
    Const xlEdgeBottom    As Integer = 9
    Const xlEdgeRight     As Integer = 10
    Const xlInsideVertical   As Integer = 11
    Const xlInsideHorizontal As Integer = 12
    '---------------------------------------
    
    '---------------------------------------
    ' Others
    '---------------------------------------
    Const xlDown  As Integer = -4121
    
    
    On Error GoTo Err_exp_EXCEL
   
    
    Set objApp = CreateObject("Excel.Application")
    
    objApp.Visible = False
    objApp.Workbooks.Add
    
    xlsBookName = objApp.ActiveWorkBook.Name
    
    If strMIDASHI = "" Then
        j = 1
    Else
        j = 2
        objApp.Activesheet.Cells(1, 1).value = strMIDASHI
    End If
    
    
    If objREMOTEDB.ExecSelect(strSQL) Then
        Set rsADO = objREMOTEDB.GetRS
        
        
        With objApp.Activesheet
            For i = 0 To rsADO.Fields.Count - 1
                .Cells(j, i + 1).value = rsADO.Fields(i).Name
                .Cells(j, i + 1).Interior.ColorIndex = 15 'Gray
                .Cells(j, i + 1).Borders(xlEdgeTop).LineStyle = xlContinuous
                .Cells(j, i + 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Cells(j, i + 1).Borders(xlEdgeRight).LineStyle = xlContinuous
                .Cells(j, i + 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Next i
            
            '1.10.15 ���o���}�[�W
            If j = 2 Then
                .Range(.Cells(1, 1), .Cells(1, i)).Merge
            End If
            
            .Cells(j + 1, 1).CopyFromRecordset rsADO
            
            '1.10.9 K.Asayama Change Bug Fix
            .Range(.Cells(j, 1), .Cells(.Cells(j, 1).end(xlDown).Row, i)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(j, 1), .Cells(.Cells(j, 1).end(xlDown).Row, i)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(j, 1), .Cells(.Cells(j, 1).end(xlDown).Row, i)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(j, 1), .Cells(.Cells(j, 1).end(xlDown).Row, i)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(j, 1), .Cells(.Cells(j, 1).end(xlDown).Row, i)).Borders(xlInsideVertical).LineStyle = xlContinuous
            .Range(.Cells(j, 1), .Cells(.Cells(j, 1).end(xlDown).Row, i)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
            '1.10.9 K.Asayama Change End
            
            .PageSetup.CenterFooter = "&P / &N �y�[�W"
            .PageSetup.PrintTitleRows = "$" & j & ":$" & j
    
        End With
        
                        
        rsADO.Close
        
        If boolFilter Then
            objApp.Rows(j & ":" & j).AutoFilter       '1��ڂɃI�[�g�t�B���^�[
        End If
        
        objApp.Cells.EntireColumn.AutoFit   '�Z����������
        
    '1.10.15
    Else
        Err.Raise 9999, , "SQL���s�G���[ SQL���m�F���Ă�������"
        
    End If
    
    Beep
    'MsgBox "EXCEL�f�[�^���쐬���܂���"
    
    objApp.Visible = True
    
    GoTo Exit_exp_EXCEL
    
Err_exp_EXCEL:
    MsgBox Err.Number & " " & Err.Description
    
    On Error Resume Next
    objApp.ActiveWorkBook.Close Savechanges:=False
    
Exit_exp_EXCEL:
    
    Set objREMOTEDB = Nothing
    Set rsADO = Nothing
    Set objApp = Nothing
    
End Sub


Public Sub exp_EXCEL_LOCAL(strSQL As String, Optional boolFilter As Boolean, Optional strMIDASHI As String)
'--------------------------------------------------------------------------------------------------------------------
'EXCEL�G�N�X�|�[�g
'   ��Input��SQL���G�N�Z���̐V�K�u�b�N�ɏo�͂���i���[�J��DB��p)
'
'   :����
'       strSQL          SQL
'       boolFilter      True�̏ꍇ�͊J�n�s�ɃI�[�g�t�B���^���|����
'       strMIDASHI      True�̏ꍇ��1�s�ڂɌ��o����\������

'--------------------------------------------------------------------------------------------------------------------
'

    Dim objLOCALDB As New cls_LOCALDB
    
    Dim rsADO As New ADODB.Recordset '1.10.15
    
    Dim objApp As Object 'Excel
    
    Dim xlsBookName As String
    
    Dim i, j As Integer
    
    Dim intCount As Integer
    '---------------------------------------
    ' LineStyle
    '---------------------------------------
    Const xlContinuous   As Integer = 1
    Const xlDashDot      As Integer = 4
    Const xlDashDotDot   As Integer = 5
    Const xlSlantDashDot As Integer = 13
    Const xlDash         As Integer = -4115
    Const xlDot          As Integer = -4118
    Const xlDouble       As Integer = -4119
    Const xlLineStyleNone As Integer = -4142
    '---------------------------------------

    '---------------------------------------
    ' Borders
    '---------------------------------------
    Const xlDiagonalDown  As Integer = 5
    Const xlDiagonalUp    As Integer = 6
    Const xlEdgeLeft      As Integer = 7
    Const xlEdgeTop       As Integer = 8
    Const xlEdgeBottom    As Integer = 9
    Const xlEdgeRight     As Integer = 10
    Const xlInsideVertical   As Integer = 11
    Const xlInsideHorizontal As Integer = 12
    '---------------------------------------
    
    '---------------------------------------
    ' Others
    '---------------------------------------
    Const xlDown  As Integer = -4121
    
    
    On Error GoTo Err_exp_EXCEL_LOCAL
   
    
    Set objApp = CreateObject("Excel.Application")
    
    objApp.Visible = False
    objApp.Workbooks.Add
    
    xlsBookName = objApp.ActiveWorkBook.Name
    
    If strMIDASHI = "" Then
        j = 1
    Else
        j = 2
        objApp.Activesheet.Cells(1, 1).value = strMIDASHI
    End If
    
    
    If objLOCALDB.ExecSelect(strSQL) Then
        Set rsADO = objLOCALDB.GetRS
        
        
        With objApp.Activesheet
            For i = 0 To rsADO.Fields.Count - 1
                .Cells(j, i + 1).value = rsADO.Fields(i).Name
                .Cells(j, i + 1).Interior.ColorIndex = 15 'Gray
                .Cells(j, i + 1).Borders(xlEdgeTop).LineStyle = xlContinuous
                .Cells(j, i + 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Cells(j, i + 1).Borders(xlEdgeRight).LineStyle = xlContinuous
                .Cells(j, i + 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Next i
            
            '1.10.18 ���o���}�[�W
            If j = 2 Then
                .Range(.Cells(1, 1), .Cells(1, i)).Merge
            End If
            
            .Cells(j + 1, 1).CopyFromRecordset rsADO
            
            .Range(.Cells(j + 1, 1), .Cells(.Cells(j + 1, 1).end(xlDown).Row, i)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(j + 1, 1), .Cells(.Cells(j + 1, 1).end(xlDown).Row, i)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(j + 1, 1), .Cells(.Cells(j + 1, 1).end(xlDown).Row, i)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(j + 1, 1), .Cells(.Cells(j + 1, 1).end(xlDown).Row, i)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(j + 1, 1), .Cells(.Cells(j + 1, 1).end(xlDown).Row, i)).Borders(xlInsideVertical).LineStyle = xlContinuous
            .Range(.Cells(j + 1, 1), .Cells(.Cells(j + 1, 1).end(xlDown).Row, i)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
            .PageSetup.CenterFooter = "&P / &N �y�[�W"
            .PageSetup.PrintTitleRows = "$" & j & ":$" & j
    
        End With
        
                        
        rsADO.Close
        
        If boolFilter Then
            objApp.Rows(j & ":" & j).AutoFilter       '1��ڂɃI�[�g�t�B���^�[
        End If
        
        objApp.Cells.EntireColumn.AutoFit   '�Z����������
        
    End If
    
    Beep
    'MsgBox "EXCEL�f�[�^���쐬���܂���"
    
    objApp.Visible = True
    
    GoTo Exit_exp_EXCEL_LOCAL
    
Err_exp_EXCEL_LOCAL:
    MsgBox Err.Number & " " & Err.Description
    
    On Error Resume Next
    objApp.ActiveWorkBook.Close Savechanges:=False
    
Exit_exp_EXCEL_LOCAL:
    
    Set objLOCALDB = Nothing
    Set rsADO = Nothing
    Set objApp = Nothing
    
End Sub

Public Sub sub_ClipBord_Paste_to_Excel()
'--------------------------------------------------------------------------------------------------------------------
'EXCEL�G�N�X�|�[�g�i�N���b�v�{�[�h�j
'   ��Excel�̐V�KBook���J���ăN���b�v�{�[�h��Paste
'
'1.10.14 ADD
'--------------------------------------------------------------------------------------------------------------------
    Dim objApp As Object 'Excel
    
    Set objApp = CreateObject("Excel.Application")
    
    On Error GoTo Err_sub_ClipBord_Paste_to_Excel
    
    objApp.Visible = False
    objApp.Workbooks.Add
    
    objApp.Activesheet.Paste
    objApp.CutCopyMode = False
    
    objApp.Visible = True
    
    GoTo Exit_sub_ClipBord_Paste_to_Excel
    
Err_sub_ClipBord_Paste_to_Excel:
    MsgBox Err.Number & " " & Err.Description
    
    On Error Resume Next
    objApp.ActiveWorkBook.Close Savechanges:=False
    
Exit_sub_ClipBord_Paste_to_Excel:
    Set objApp = Nothing
End Sub

Public Function fncbolFileToExcel(strFileFullpath As String, byteConnectionDB As Byte, Optional boolFilter As Boolean, Optional strMIDASHI As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'�ėpEXCEL�G�N�X�|�[�g
'   ���t�@�C���i�t���p�X�j��ǂݍ����SQL�����s�����ʂ�Excel�ɓ]��
'
'   :����
'       strFileFullpath     �t�@�C�����i�t���p�X�j
'       byteConnectionDB    0:�����[�g(SQLSERVER) 1:���[�J��:(ACCESS)
'       boolFilter          True�̏ꍇ�͊J�n�s�ɃI�[�g�t�B���^���|����
'       strMIDASHI          True�̏ꍇ��1�s�ڂɌ��o����\������
'
'1.10.18 ADD
'--------------------------------------------------------------------------------------------------------------------
    
    Dim strSQL As String
    
    On Error GoTo Err_fncbolFileToExcel
    

    strSQL = ""
    
    If Dir(strFileFullpath) <> "" Then
        With CreateObject("Scripting.FileSystemObject")
            With .GetFile(strFileFullpath).OpenAsTextStream
                strSQL = .ReadAll
                .Close
            End With
        End With
        
        '���s�폜
        strSQL = Replace(Replace(strSQL, vbCrLf, " "), vbLf, " ")

        Screen.MousePointer = 11
        
        If byteConnectionDB = 0 Then
            exp_EXCEL strSQL, boolFilter, strMIDASHI
        Else
            exp_EXCEL_LOCAL strSQL, boolFilter, strMIDASHI
        End If

    Else
        Err.Raise 9999, , "SQL�t�@�C�������݂��܂���B�Ǘ��҂ɘA�����Ă�������"
    End If
    
    fncbolFileToExcel = True
    
    GoTo Exit_fncbolFileToExcel

Err_fncbolFileToExcel:
    Close
    MsgBox Err.Description
Exit_fncbolFileToExcel:
    Screen.MousePointer = 0
    
End Function

Public Function bolfncexp_EXCELOBJECT(in_objRS As ADODB.Recordset, in_ExcelObj As Object, Optional boolFilter As Boolean, Optional strMIDASHI As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'EXCEL�G�N�X�|�[�g
'   ��Input�̃��R�[�h�Z�b�g�������Ŏ󂯎����Excel���[�N�V�[�g�ɏo��
'       �i�\��t����Ɏ󂯓n������Excel�𑀍삵�����ꍇ�Ɏg�p�j

'       ���\��t���郏�[�N�V�[�g���A�N�e�B���ɂ��Ă���n������
'
'   :����
'       in_objRS        ���R�[�h�Z�b�g
'       in_ExcelObj     ���R�[�h�Z�b�g
'       boolFilter      True�̏ꍇ�͊J�n�s�ɃI�[�g�t�B���^���|����
'       strMIDASHI      True�̏ꍇ��1�s�ڂɌ��o����\������

'   1.11.0 K.Asayama ADD

'2.3.0
'   ���r����xlDown����xlUp�֕ύX
'--------------------------------------------------------------------------------------------------------------------
    
    Dim xlsBookName As String
    
    Dim i, j As Integer
    
    Dim intCount As Integer
    '---------------------------------------
    ' LineStyle
    '---------------------------------------
    Const xlContinuous   As Integer = 1
    Const xlDashDot      As Integer = 4
    Const xlDashDotDot   As Integer = 5
    Const xlSlantDashDot As Integer = 13
    Const xlDash         As Integer = -4115
    Const xlDot          As Integer = -4118
    Const xlDouble       As Integer = -4119
    Const xlLineStyleNone As Integer = -4142
    '---------------------------------------

    '---------------------------------------
    ' Borders
    '---------------------------------------
    Const xlDiagonalDown  As Integer = 5
    Const xlDiagonalUp    As Integer = 6
    Const xlEdgeLeft      As Integer = 7
    Const xlEdgeTop       As Integer = 8
    Const xlEdgeBottom    As Integer = 9
    Const xlEdgeRight     As Integer = 10
    Const xlInsideVertical   As Integer = 11
    Const xlInsideHorizontal As Integer = 12
    '---------------------------------------
    
    '---------------------------------------
    ' Others
    '---------------------------------------
    Const xlDown  As Integer = -4121
    Const xlUp  As Integer = -4162
    
    bolfncexp_EXCELOBJECT = False
    
    On Error GoTo Err_bolfncexp_EXCELOBJECT
   
    xlsBookName = in_ExcelObj.ActiveWorkBook.Name
    
    If strMIDASHI = "" Then
        j = 1
    Else
        j = 2
        in_ExcelObj.Activesheet.Cells(1, 1).value = strMIDASHI
    End If
    
    
    With in_ExcelObj.Activesheet
        For i = 0 To in_objRS.Fields.Count - 1
            .Cells(j, i + 1).value = in_objRS.Fields(i).Name
            .Cells(j, i + 1).Interior.ColorIndex = 15 'Gray
            .Cells(j, i + 1).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Cells(j, i + 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Cells(j, i + 1).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Cells(j, i + 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Next i
        
        If j = 2 Then
            .Range(.Cells(1, 1), .Cells(1, i)).Merge
        End If
        
        .Cells(j + 1, 1).CopyFromRecordset in_objRS
        
        .Range(.Cells(j, 1), .Cells(.Cells(.Rows.Count, 1).end(xlUp).Row, i)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range(.Cells(j, 1), .Cells(.Cells(.Rows.Count, 1).end(xlUp).Row, i)).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(.Cells(j, 1), .Cells(.Cells(.Rows.Count, 1).end(xlUp).Row, i)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(.Cells(j, 1), .Cells(.Cells(.Rows.Count, 1).end(xlUp).Row, i)).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range(.Cells(j, 1), .Cells(.Cells(.Rows.Count, 1).end(xlUp).Row, i)).Borders(xlInsideVertical).LineStyle = xlContinuous
        .Range(.Cells(j, 1), .Cells(.Cells(.Rows.Count, 1).end(xlUp).Row, i)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
        .PageSetup.CenterFooter = "&P / &N �y�[�W"
        .PageSetup.PrintTitleRows = "$" & j & ":$" & j

    End With
        
        
    If boolFilter Then
        in_ExcelObj.Rows(j & ":" & j).AutoFilter       '1��ڂɃI�[�g�t�B���^�[
    End If
        
    in_ExcelObj.Cells.EntireColumn.AutoFit   '�Z����������

    
    Beep
    'MsgBox "EXCEL�f�[�^���쐬���܂���"
    
    bolfncexp_EXCELOBJECT = True
    
    GoTo Exit_bolfncexp_EXCELOBJECT
    
Err_bolfncexp_EXCELOBJECT:
    bolfncexp_EXCELOBJECT = False
    MsgBox Err.Number & " " & Err.Description
    
Exit_bolfncexp_EXCELOBJECT:

End Function