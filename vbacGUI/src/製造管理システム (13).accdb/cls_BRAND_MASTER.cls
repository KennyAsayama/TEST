Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private cnADO As ADODB.Connection
Private rsADO  As ADODB.Recordset

Private Sub Class_Initialize()
'アクセスの場合
    'Set cnADO = Application.CurrentProject.Connection
    
'SQLServerの場合
    Set cnADO = New ADODB.Connection

'ODBCの場合
    cnADO.Open "Driver={SQL Server}; " & _
    "server=" & Connection_DB() & "; " & _
    "database=" & "BRAND_MASTER" & "; " & _
    "Trusted_Connection=yes"
'    "uid=" & "sa" & "; " & _
'    "pwd=" & "kamiya" & ";"

'OLEProviderの場合
'    cnADO.Open "Provider=SQLOLEDB " & _
'                    ";Integrated Security=SSPI " & _
'                    ";Data Source=" & Connection_DB() & _
'                    ";Initial Catalog=" & "BRAND_MASTER" & "; "
                    
    'Set rsADO = New ADODB.Recordset


End Sub

Private Sub Class_Terminate()
    On Error Resume Next

    rsADO.Close
    cnADO.Close
    
    Set cnADO = Nothing
    Set rsADO = Nothing
    
End Sub

Public Function ExecSelect(SQL As String) As Boolean

    ExecSelect = False

    On Error Resume Next

    Set rsADO = cnADO.Execute(SQL)

    'SQLを実行時の結果を判定、実行エラーならFalse
    If cnADO.Errors.Count = 0 Then
       ExecSelect = True
    End If

    On Error GoTo 0

End Function
Public Function ExecSelect_ExternalRS(rsIO As ADODB.Recordset, SQL As String) As Boolean

    ExecSelect_ExternalRS = False

    On Error Resume Next
    
    Set rsIO = New ADODB.Recordset
    
    rsIO.Open SQL, cnADO, adOpenKeyset
    
    'SQLを実行時の結果を判定、実行エラーならFalse
    If cnADO.Errors.Count = 0 Then
       ExecSelect_ExternalRS = True
    End If

    On Error GoTo 0

End Function

Public Function ExecSelect_Keyset(SQL As String) As Boolean

    ExecSelect_Keyset = False
    
    On Error Resume Next

    Set rsADO = New ADODB.Recordset
    
    rsADO.Open SQL, cnADO, adOpenKeyset
    
    'SQLを実行時の結果を判定、実行エラーならFalse
    If cnADO.Errors.Count = 0 Then
        ExecSelect_Keyset = True
    End If

    On Error GoTo 0

End Function

Public Function ExecSelect_Writable(SQL As String) As Boolean

    ExecSelect_Writable = False
    
    On Error Resume Next

    Set rsADO = New ADODB.Recordset
    
    rsADO.Open SQL, cnADO, adOpenKeyset, adLockOptimistic
    
    'SQLを実行時の結果を判定、実行エラーならFalse
    If cnADO.Errors.Count = 0 Then
        ExecSelect_Writable = True
    End If

    On Error GoTo 0

End Function

Public Function ExecSQL(SQL As String, Optional ByRef strErrMsg As Variant) As Boolean
    
    ExecSQL = False

    On Error Resume Next

    Call cnADO.Execute(SQL)

    'SQLを実行時の結果を判定、実行エラーならFalse
    If cnADO.Errors.Count = 0 Then
       ExecSQL = True
    Else
        If Not IsMissing(strErrMsg) Then
            strErrMsg = cnADO.Errors.Item(0).Description
        End If
    End If
'    Debug.Print strErrMSG
    
    On Error GoTo 0

End Function

Public Sub BeginTrans()

    cnADO.BeginTrans

End Sub

Public Sub Commit()

    cnADO.CommitTrans

End Sub

Public Sub Rollback()

    cnADO.RollbackTrans

End Sub

Public Property Get GetRS() As ADODB.Recordset

    Set GetRS = rsADO

End Property

Public Sub TimeOutValue(intVal As Integer)
'1.11.0 ADD
    cnADO.CommandTimeout = intVal
    
End Sub

Public Function AsyncState() As Integer
'1.11.1 ADD 非同期処理

    On Error GoTo Err_AsyncState
    
    AsyncState = cnADO.State
    
    Exit Function
    
Err_AsyncState:
    AsyncState = 999
End Function

Public Function ExecSQL_Async(SQL As String, Optional ByRef strErrMsg As Variant) As Boolean
'1.11.1 ADD 非同期実行
    ExecSQL_Async = False

    On Error Resume Next

    Call cnADO.Execute(SQL, , adAsyncConnect)

    'SQLを実行時の結果を判定、実行エラーならFalse
    If cnADO.Errors.Count = 0 Then
       ExecSQL_Async = True
    Else
        If Not IsMissing(strErrMsg) Then
            strErrMsg = cnADO.Errors.Item(0).Description
        End If
    End If
'    Debug.Print strErrMSG
    
    On Error GoTo 0

End Function

Public Sub RecordSetClose()
'1.11.1 ADD レコードセットを閉じる
    On Error Resume Next

    rsADO.Close
    
End Sub

Public Property Get GetClerkName(ByVal ClerkID As Variant) As String
'社員コードから社員名取得
'1.12.2 ADD
    Dim RsClerk As New ADODB.Recordset
    Dim strSQL As String
    
    GetClerkName = ""
    On Error GoTo Err_GetClerkName
    
    If IsNull(ClerkID) Then Exit Property
    
    strSQL = "select 担当者氏名 from T_社員ﾏｽﾀ_工程管理 where 担当者ID = '" & ClerkID & "' "
    
    Set RsClerk = cnADO.Execute(strSQL)
    
    If Not RsClerk.EOF Then
        GetClerkName = RsClerk![担当者氏名]
    End If
    
    RsClerk.Close
    
    GoTo Exit_GetClerkName

Err_GetClerkName:
    GetClerkName = ""
Exit_GetClerkName:
    Set RsClerk = Nothing
    
End Property

Public Property Get GetSystemDate() As Date
'サーバの日付取得
'1.12.2 ADD
    Dim RsDate As New ADODB.Recordset
    Dim strSQL As String
    
    GetSystemDate = Now()
    On Error GoTo Err_GetSystemDate

    strSQL = "select getdate() システム日付 "
    
    Set RsDate = cnADO.Execute(strSQL)
    
    If Not RsDate.EOF Then
        GetSystemDate = RsDate![システム日付]
    End If
    
    RsDate.Close
    
    GoTo Exit_GetSystemDate

Err_GetSystemDate:
    GetSystemDate = Now()
Exit_GetSystemDate:
    Set RsDate = Nothing
End Property