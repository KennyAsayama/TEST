Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'--------------------------------------------------------------------------------------------------------------------
'エクセル操作クラス
'   インスタンス化でエクセル起動
'   インスタンス破棄時bolQuit変数を[False]にしておくとExcelは起動したまま残る

'1.12.0 ADD
'2.0.0
'   →不要シート削除メソッド化
'--------------------------------------------------------------------------------------------------------------------

    'Private objExcel As Excel.Application  'Excel
    Private objExcel As Object
    Private bolQuit As Boolean
    Private BookName As String
    
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
    Const xlUp    As Integer = -4162
    Const xlToLeft As Integer = -4159
    Const xlToRight As Integer = -4161

Private Sub Class_Initialize()
'--------------------------------------------------------------------------------------------------------------------
'コンストラクタ
'--------------------------------------------------------------------------------------------------------------------
    Set objExcel = CreateObject("Excel.Application")
    bolQuit = True
    UserAddin
End Sub

Private Sub Class_Terminate()
'--------------------------------------------------------------------------------------------------------------------
'ディストラクタ
'--------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    'bolQuit=Trueの時Excelを閉じる
    If bolQuit Then
        AllClose
        objExcel.Quit
    Else
        objExcel.Visible = True
    End If
    
    Set objExcel = Nothing
End Sub

Private Function AllClose() As Boolean
'--------------------------------------------------------------------------------------------------------------------
'Excelのシートをすべて保存しないで閉じる
'   クラス内からのみ呼び出し可
'--------------------------------------------------------------------------------------------------------------------
    Dim wb As Object
    AllClose = False
    
    On Error GoTo Err_AllClose
    
    For Each wb In objExcel.Workbooks            '開いている全てのブックに対して
        If wb.Saved = False Then
            wb.Close Savechanges:=False
        End If
    Next wb
    
    AllClose = True
    
    GoTo Exit_AllClose
    
Err_AllClose:
    AllClose = False
    
Exit_AllClose:
    Set wb = Nothing
    
End Function
Public Function WorkSheetADD(Optional byName As String = "") As Boolean
'--------------------------------------------------------------------------------------------------------------------
'ワークシートを引数の名称で末に追加
'--------------------------------------------------------------------------------------------------------------------
    On Error GoTo Err_WorkSheetADD
        
    With objExcel
        .Worksheets.Add After:=.Worksheets(.Worksheets.Count)
        If Not IsMissing(byName) Then
            .Activesheet.Name = byName
        End If
        
    End With
    
Err_WorkSheetADD:

End Function
Public Function WorkSheetDelete(WorkSheetName As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'引数の名称のワークシートを削除
'--------------------------------------------------------------------------------------------------------------------
    Dim ws As Object
    
    On Error GoTo Err_WorkSheetDelete
    
    objExcel.Application.displayalerts = False
    
    For Each ws In objExcel.Worksheets
        If ws.Name = WorkSheetName Then
            objExcel.Worksheets(WorkSheetName).Delete
        End If
    Next ws
    
    WorkSheetDelete = True
    
    GoTo Exit_WorkSheetDelete
    
Err_WorkSheetDelete:
    WorkSheetDelete = False
    
Exit_WorkSheetDelete:
    objExcel.Application.displayalerts = True
    Set ws = Nothing
End Function

Public Function WorkSheetExists(WorkSheetName As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'引数の名称のワークシートが存在するか確認
'--------------------------------------------------------------------------------------------------------------------
    Dim ws As Object
    
    On Error GoTo Err_WorkSheetExists
    
    WorkSheetExists = False
    
    For Each ws In objExcel.Worksheets
        If ws.Name = WorkSheetName Then
            WorkSheetExists = True
            Exit For
        End If
    Next ws
    
    GoTo Exit_WorkSheetExists
    
Err_WorkSheetExists:
    WorkSheetExists = False
    
Exit_WorkSheetExists:
    Set ws = Nothing
    
End Function
Public Property Let ContinueOpen(x As Boolean)
'--------------------------------------------------------------------------------------------------------------------
'インスタンス終了後もExcelを開いたままにしたい場合はTrueにする
'--------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    If x = True Then
        bolQuit = False
    Else
        bolQuit = True
    End If
    
End Property

Public Property Get getExcel() As Object
'--------------------------------------------------------------------------------------------------------------------
'Excelオブジェクト取得
'--------------------------------------------------------------------------------------------------------------------
    Set getExcel = objExcel
End Property

Private Function UserAddin()
'--------------------------------------------------------------------------------------------------------------------
'ユーザのAddInとスタートアップマクロ起動
'1.12.2 ADD
'--------------------------------------------------------------------------------------------------------------------
    Dim strAddinPath As String
    Dim strxlsSTARTPath As String
    Dim strFName As String
    Dim i As Integer
    
    strAddinPath = objExcel.Application.UserLibraryPath
    strxlsSTARTPath = objExcel.Application.StartupPath
    
        i = 0
    'ファイル探索
    
    strFName = Dir(strAddinPath & "\*.xlam", vbNormal)
    
    Do While strFName <> ""
        
        i = i + 1
        
        'Debug.Print strFName
        objExcel.Workbooks.Open strAddinPath & "\" & strFName
        strFName = Dir()
        
    Loop
    
    strFName = ""
    strFName = Dir(strxlsSTARTPath & "\*.xlsm", vbNormal)
    
    Do While strFName <> ""
        
        i = i + 1
        
        'Debug.Print strFName
        objExcel.Workbooks.Open strAddinPath & "\" & strFName
        strFName = Dir()
        
    Loop
End Function

Public Function UnnecessarySheetDel(Optional inSheetName As String = "Sheet*") As Boolean
'--------------------------------------------------------------------------------------------------------------------
'不要シート削除
'2.0.0 ADD
'   Excelの不要ワークシートを削除
'       引数:   inSheetNameを受け取った場合は該当シート名を削除、
'               文字列にアスタリスク(*)が含まれている場合はあいまい検索指定
'--------------------------------------------------------------------------------------------------------------------

    Dim intSheetDel As Integer
    Dim i As Integer
    Dim bolAmbiguous As Boolean
    
    On Error GoTo Err_UnnecessarySheetDel
    
    intSheetDel = 0
    i = 0
    
    '曖昧検索指定
    If InStr(1, inSheetName, "*") > 0 Then
        bolAmbiguous = True
    Else
        bolAmbiguous = False
    End If
    
    With Me.getExcel
        
        If bolAmbiguous Then
            For i = 1 To .Worksheets.Count
                If .Worksheets(i - intSheetDel).Name Like inSheetName Then
                    .Worksheets(i - intSheetDel).Delete
                    intSheetDel = intSheetDel + 1
                End If
            Next
        Else
            For i = 1 To .Worksheets.Count
                If .Worksheets(i - intSheetDel).Name = inSheetName Then
                    .Worksheets(i - intSheetDel).Delete
                    intSheetDel = intSheetDel + 1
                End If
            Next
        End If
            
        .Worksheets(1).Activate
    
    End With
    
    UnnecessarySheetDel = True
    
    Exit Function
    
Err_UnnecessarySheetDel:
    MsgBox Err.Description

End Function

Public Function WorkBookSave(ByVal inPath As String, ByVal inFileName As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'開いたブックに名前を付けて保存する

'   開いたExcelオブジェクトに名前を付けて保存する
'   保存に成功した場合とはTrueを返す
'   保存に失敗した場合はFalseを返す
'   （ただし同じファイル名がすでに存在している場合は上書きをするかメッセージを出す
'     その場合は保存するしないにかかわらずTrueを返す）

'       引数:   inPath      保存するパス名
'               inFileName  保存するファイル名（拡張子付き）

'2.5.3 ADD
'--------------------------------------------------------------------------------------------------------------------
    WorkBookSave = False
    
    With Me.getExcel
        If Dir(inPath & "\" & inFileName) <> "" Then
        
            If MsgBox("同じファイル名のファイルが既にあります。" & vbCrLf & "上書きしますか？", vbExclamation + vbYesNo, "ファイル上書き確認") = vbYes Then
                .DisplayAlarts = False
                .ActiveWorkBook.SaveAs FileName:=inPath & "\" & inFileName
                .DisplayAlarts = True
            End If
            
        Else
            .ActiveWorkBook.SaveAs FileName:=inPath & "\" & inFileName
            
        End If
        
        WorkBookSave = True
        
    End With
    
End Function