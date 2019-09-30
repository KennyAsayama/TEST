Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'--------------------------------------------------------------------------------------------------------------------
'本クラスにバインドする際のラベル名は"_Sort"で終わる名前にして下さい
'--------------------------------------------------------------------------------------------------------------------
 
Private mctlName As String
Private mctlName_Parent As String
Private WithEvents mctl As Access.Label
Attribute mctl.VB_VarHelpID = -1

Public Sub Bind(objCtl As Access.Control)
    If TypeOf objCtl Is Access.Label Then
        Set mctl = objCtl
        mctlName = mctl.Name
        mctl.OnDblClick = "[EVENT PROCEDURE]"
    End If

End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set mctl = Nothing
End Sub

Private Sub mctl_DblClick(Cancel As Integer)
    Dim strSortName As String
    Dim bolASC As Boolean

'1.12.2
'   →F_製造先行受付を追加
'2.0.0
'   →F_フルハイトライン_数量を追加
'3.0.0
'   →対応できないオブジェクト名を想定してカッコ[]をセット
'--------------------------------------------------------------------------------------------------------------------
'フォーム別処理をここにセット
'--------------------------------------------------------------------------------------------------------------------

    'フォーム（F_邸別_数量）用
    If mctl.Parent.Name Like "F_邸別_数量*" Or mctl.Parent.Name Like "F_製造先行受付*" Or mctl.Parent.Name Like "F_フルハイトライン_数量*" Then
        If Forms(mctl.Parent.Name).OrderByOn = True Then
        
            If mctl.ForeColor = vbGreen Then
                bolASC = False
            Else
                bolASC = True
            End If
            
            Call Sort_labelColor(mctl)
        End If
        
        strSortName = Replace(Replace(mctl.Name, "lbl", ""), "_Sort", "")
        'MsgBox strSortName
        
        If bolASC Then
            Forms(mctl.Parent.Name).OrderBy = "[" & strSortName & "]"
            mctl.ForeColor = vbGreen
        Else
            Forms(mctl.Parent.Name).OrderBy = "[" & strSortName & "] DESC"
            mctl.ForeColor = RGB(&HF5, &H9D, &H56)
        End If
        
        Forms(mctl.Parent.Name).OrderByOn = True
    End If
End Sub

Private Sub Sort_labelColor(ByVal in_ctrl As Access.Control)
'--------------------------------------------------------------------------------------------------------------------
'ラベル表示色初期化
'--------------------------------------------------------------------------------------------------------------------
    Dim ctl As Access.Control
    Dim i As Byte
    i = 0
    
    On Error GoTo Err_Sort_labelColor
    
    'このフォーム内のすべてのコントロールを検索
    For Each ctl In Forms(in_ctrl.Parent.Name).Controls
        With ctl
            If .ControlType = acLabel And .Name Like "*_Sort" Then
                   
                ctl.ForeColor = vbWhite

            End If
        End With
    Next ctl
        
    GoTo Exit_Sort_labelColor
    
Err_Sort_labelColor:

Exit_Sort_labelColor:
    Set ctl = Nothing
    
End Sub