VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Labelset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mctl As Access.Label
Private mctlName As String
Private mctlName_Parent As String

Public Sub Bind(objCtl As Access.Control)
    If TypeOf objCtl Is Access.Label Then

        Set mctl = objCtl
        mctlName = mctl.Name
        '日付ラベル、確定ラベルはParentに総数ラベル名(lbl_D*)をセットする
        If mctlName Like "lbl_Disp*" Then
            mctlName_Parent = Replace(mctlName, "Disp", "D")
        ElseIf mctlName Like "lbl_K*" Then
            mctlName_Parent = Replace(mctlName, "lbl_K", "lbl_D")
        Else
            mctlName_Parent = mctlName
        End If

    End If

End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    Set mctl = Nothing

End Sub

Public Sub CaptionSet(in_Text As String)
    mctl.Caption = in_Text
End Sub

Public Sub SetTag(in_Text As String)
    mctl.Tag = in_Text
End Sub

Public Function GetTag() As String
    GetTag = mctl.Tag
End Function

Public Function GetName() As String
    GetName = mctlName
End Function

Public Function myVisible(bolVisible As Boolean) As String
    mctl.Visible = bolVisible
End Function

Public Sub SetColor(in_color As Long)
    mctl.ForeColor = in_color
End Sub

Public Sub SetBackColor(in_color As Long)
    mctl.BackColor = in_color
End Sub

Public Sub SetControlTipText(in_String As String)
    mctl.ControlTipText = in_String
End Sub

Public Function myWidth() As Double
    myWidth = mctl.Width
End Function

Public Sub SetWidth(in_Width As Double)
    mctl.Width = in_Width
End Sub
