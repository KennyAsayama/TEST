Attribute VB_Name = "md_Excel"
Option Compare Database

Sub exp_EXCEL(strSQL As String, Optional boolFilter As Boolean, Optional strMIDASHI As String)
'--------------------------------------------------------------------------------------------------------------------
'EXCELエクスポート
'   →InputのSQLをエクセルの新規ブックに出力する
'
'   :引数
'       strSQL          SQL
'       boolFilter      Trueの場合は開始行にオートフィルタを掛ける
'       strMIDASHI      Trueの場合は1行目に見出しを表示する

'--------------------------------------------------------------------------------------------------------------------
'

    Dim objREMOTEDB As New cls_BRAND_MASTER
    
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
    
    
    On Error GoTo Err_exp_EXCEL
   
    
    Set objApp = CreateObject("Excel.Application")
    
    objApp.Visible = False
    objApp.workbooks.Add
    
    xlsBookName = objApp.ActiveWorkBook.Name
    
    If strMIDASHI = "" Then
        j = 1
    Else
        j = 2
        objApp.ActiveSheet.cells(1, 1).value = strMIDASHI
    End If
    
    
    If objREMOTEDB.ExecSelect(strSQL) Then
        Set rsADO = objREMOTEDB.GetRS
        
        
        With objApp.ActiveSheet
            For i = 0 To rsADO.Fields.Count - 1
                .cells(j, i + 1).value = rsADO.Fields(i).Name
                .cells(j, i + 1).Interior.ColorIndex = 15 'Gray
                .cells(j, i + 1).Borders(xlEdgeTop).LineStyle = xlContinuous
                .cells(j, i + 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
                .cells(j, i + 1).Borders(xlEdgeRight).LineStyle = xlContinuous
                .cells(j, i + 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Next i
    
            .cells(j + 1, 1).CopyFromRecordset rsADO
            
            .Range(.cells(j + 1, 1), .cells(.cells(j + 1, 1).End(xlDown).Row, i)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.cells(j + 1, 1), .cells(.cells(j + 1, 1).End(xlDown).Row, i)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.cells(j + 1, 1), .cells(.cells(j + 1, 1).End(xlDown).Row, i)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.cells(j + 1, 1), .cells(.cells(j + 1, 1).End(xlDown).Row, i)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.cells(j + 1, 1), .cells(.cells(j + 1, 1).End(xlDown).Row, i)).Borders(xlInsideVertical).LineStyle = xlContinuous
            .Range(.cells(j + 1, 1), .cells(.cells(j + 1, 1).End(xlDown).Row, i)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
            .PageSetup.CenterFooter = "&P / &N ページ"
            .PageSetup.PrintTitleRows = "$" & j & ":$" & j
    
        End With
        
                        
        rsADO.Close
        
        If boolFilter Then
            objApp.Rows(j & ":" & j).AutoFilter       '1列目にオートフィルター
        End If
        
        objApp.cells.EntireColumn.AutoFit   'セル自動調整
        
    End If
    
    Beep
    'MsgBox "EXCELデータを作成しました"
    
    objApp.Visible = True
    
    GoTo Exit_exp_EXCEL
    
Err_exp_EXCEL:
    MsgBox Err.Number & " " & Err.Description
    
    On Error Resume Next
    objApp.ActiveWorkBook.Close savechanges:=False
    
Exit_exp_EXCEL:
    
    Set objREMOTEDB = Nothing
    Set rsADO = Nothing
    Set objApp = Nothing
    
End Sub


Sub exp_EXCEL_LOCAL(strSQL As String, Optional boolFilter As Boolean, Optional strMIDASHI As String)
'--------------------------------------------------------------------------------------------------------------------
'EXCELエクスポート
'   →InputのSQLをエクセルの新規ブックに出力する（ローカルDB専用)
'
'   :引数
'       strSQL          SQL
'       boolFilter      Trueの場合は開始行にオートフィルタを掛ける
'       strMIDASHI      Trueの場合は1行目に見出しを表示する

'--------------------------------------------------------------------------------------------------------------------
'

    Dim objLOCALDB As New cls_LOCALDB
    
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
    objApp.workbooks.Add
    
    xlsBookName = objApp.ActiveWorkBook.Name
    
    If strMIDASHI = "" Then
        j = 1
    Else
        j = 2
        objApp.ActiveSheet.cells(1, 1).value = strMIDASHI
    End If
    
    
    If objLOCALDB.ExecSelect(strSQL) Then
        Set rsADO = objLOCALDB.GetRS
        
        
        With objApp.ActiveSheet
            For i = 0 To rsADO.Fields.Count - 1
                .cells(j, i + 1).value = rsADO.Fields(i).Name
                .cells(j, i + 1).Interior.ColorIndex = 15 'Gray
                .cells(j, i + 1).Borders(xlEdgeTop).LineStyle = xlContinuous
                .cells(j, i + 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
                .cells(j, i + 1).Borders(xlEdgeRight).LineStyle = xlContinuous
                .cells(j, i + 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Next i
    
            .cells(j + 1, 1).CopyFromRecordset rsADO
            
            .Range(.cells(j + 1, 1), .cells(.cells(j + 1, 1).End(xlDown).Row, i)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.cells(j + 1, 1), .cells(.cells(j + 1, 1).End(xlDown).Row, i)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.cells(j + 1, 1), .cells(.cells(j + 1, 1).End(xlDown).Row, i)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.cells(j + 1, 1), .cells(.cells(j + 1, 1).End(xlDown).Row, i)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.cells(j + 1, 1), .cells(.cells(j + 1, 1).End(xlDown).Row, i)).Borders(xlInsideVertical).LineStyle = xlContinuous
            .Range(.cells(j + 1, 1), .cells(.cells(j + 1, 1).End(xlDown).Row, i)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
            .PageSetup.CenterFooter = "&P / &N ページ"
            .PageSetup.PrintTitleRows = "$" & j & ":$" & j
    
        End With
        
                        
        rsADO.Close
        
        If boolFilter Then
            objApp.Rows(j & ":" & j).AutoFilter       '1列目にオートフィルター
        End If
        
        objApp.cells.EntireColumn.AutoFit   'セル自動調整
        
    End If
    
    Beep
    'MsgBox "EXCELデータを作成しました"
    
    objApp.Visible = True
    
    GoTo Exit_exp_EXCEL_LOCAL
    
Err_exp_EXCEL_LOCAL:
    MsgBox Err.Number & " " & Err.Description
    
    On Error Resume Next
    objApp.ActiveWorkBook.Close savechanges:=False
    
Exit_exp_EXCEL_LOCAL:
    
    Set objLOCALDB = Nothing
    Set rsADO = Nothing
    Set objApp = Nothing
    
End Sub

