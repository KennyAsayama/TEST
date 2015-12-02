VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_LOCALDB"
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
    Set cnADO = Application.CurrentProject.Connection
    
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

    'SQLを実行時の結果を判定
    If cnADO.Errors.Count = 0 Then
       ExecSelect = True
    End If

    On Error GoTo 0

End Function

Public Function ExecSelect_Writable(SQL As String) As Boolean

    ExecSelect_Writable = False
    
    On Error Resume Next

    Set rsADO = New ADODB.Recordset
    
    rsADO.Open SQL, cnADO, adOpenKeyset, adLockPessimistic
    
    'SQLを実行時の結果を判定、実行エラーならFalse
    If cnADO.Errors.Count = 0 Then
        ExecSelect_Writable = True
    End If

    On Error GoTo 0

End Function

Public Function ExecSQL(SQL As String, Optional ByRef strErrMsg As String) As Boolean

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
