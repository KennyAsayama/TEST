Option Compare Database
Option Explicit
'2.1.0 ADD

Public Sub AutoFontSize(ByRef Ctr As Control, ByVal IniFontSize As Integer, Optional ByVal MinFontSize = 8)
'   *************************************************************
'   AutoFontSize
'    'ブランドシステムからコピー
'    'レポート内のテキストのフォントサイズを
'    'テキストボックス、ラベルの幅に合わせる

'    Input項目
'           Ctr         : コントロール
'           IniFontSize : フォントサイズ初期値
'           MinFontSize : 最小フォントサイズ（デフォルト値:8）

'   *************************************************************

    'Const MinFontSize = 4 '最小のフォントサイズ
    Const d = 40 'うまく収まらずに改行されてしまう場合はここの数値を増やす
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

'   Accessレポート出力関数

'    Input項目
'           strReportName       : レポート名
'           blnPreView          : previewモード
'           varWhereCondition   : パラメータ
'           intCopies(Option)   : 印刷部数（デフォルト値:1）
'           bolVisible(Option)  : True→可視　Flse→不可視（デフォルト値:True）

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
        
    Reports(strReportName).Printer.Copies = intCopies '部数指定
    
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
'   全てのレポートを閉じる

'   *************************************************************
        
    Dim i As Integer
    
    On Error Resume Next
    
    For i = Reports.Count - 1 To 0 Step -1
        DoCmd.Close acReport, Reports(i).Name
    Next i
  
End Sub

Public Sub Visible_ALLReports()
'   *************************************************************
'   全てのレポートを可視化する

'   *************************************************************
        
    Dim i As Integer
    
    On Error Resume Next
    
    For i = Reports.Count - 1 To 0 Step -1
        Reports(i).Visible = True
    Next i
  
End Sub