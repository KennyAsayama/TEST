Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'2.1.0 ADD

Private cnADO As ADODB.Connection
Private rsADO  As ADODB.Recordset

Private Sub Class_Initialize()
'アクセスの場合
    'Set cnADO = Application.CurrentProject.Connection
                    
'OracleOLEProviderの場合
    Set cnADO = New ADODB.Connection
        
    cnADO.Open "Data Source=skamiyadb " & _
                ";User ID=" & "skamiya" & " " & _
                ";Password=" & "skamiya" & " " & _
                    ";Pooling=False; "


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
    Else
        MsgBox cnADO.Errors.Item(0).Description
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
    
    cnADO.CommandTimeout = intVal
End Sub

Public Sub RecordSetClose()
'レコードセットを閉じる
    On Error Resume Next

    rsADO.Close
    
End Sub