Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Reg As Object
'2.3.0 パターンをブランドシステムと統合
'Const reTateguPattern As String = "^(特 )?(HE)?([A-Z])([A-Z])([A-Z])([A-Z]{2}S?)-(\d{2})(\d{2})([A-Z][A-Z]?[A-Z]?[A-Z]?)-([NMKC])([LR]?)\((.+)\)$"
Private Const reTateguPattern As String = "^(特 )?(HE|TY)?([A-Z])([A-Z])([A-Z])([A-Z]{2,3})-(\d{2})(\d{2})([A-Z]{1,5})-([NMKC])([LR]?)\((.+)\)$"

Private Sub Class_Initialize()
    Set Reg = CreateObject("VBScript.RegExp")
    Reg.Pattern = reTateguPattern
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    Set Reg = Nothing
    
End Sub
Public Property Get IsTateguHinban(ByVal in_Hinban As Variant) As Boolean

    On Error GoTo Err_IsTateguHinban
    
    If IsNull(in_Hinban) Then Exit Property

    If Reg.Test(in_Hinban) Then
        IsTateguHinban = True
    End If
        
    Exit Property
    
Err_IsTateguHinban:
    IsTateguHinban = False
    
End Property

Public Property Get シリーズ(in_Hinban As Variant) As Variant
    
    If IsTateguHinban(in_Hinban) Then
        シリーズ = varfncGetDetail(in_Hinban, 3)
    End If
    
End Property

Public Property Get 枠仕様(in_Hinban As Variant) As Variant

    If IsTateguHinban(in_Hinban) Then
        枠仕様 = varfncGetDetail(in_Hinban, 4)
    End If
    
End Property

Public Property Get 建具スタイル(in_Hinban As Variant) As Variant

    If IsTateguHinban(in_Hinban) Then
        建具スタイル = varfncGetDetail(in_Hinban, 5)
    End If
    
End Property

Public Property Get 開閉様式(in_Hinban As Variant) As Variant

    If IsTateguHinban(in_Hinban) Then
        開閉様式 = varfncGetDetail(in_Hinban, 6)
    End If
    
End Property

Public Property Get DW(in_Hinban As Variant) As Variant
    Dim strDW As String

    If IsTateguHinban(in_Hinban) Then
        strDW = varfncGetDetail(in_Hinban, 7)
    
        DW = CInt(strDW)
    End If
    
End Property

Public Property Get DH(in_Hinban As Variant) As Variant
    Dim strDH As String
    
    If IsTateguHinban(in_Hinban) Then
        strDH = varfncGetDetail(in_Hinban, 8)
    
        DH = CInt(strDH)
    End If
    
End Property

Public Property Get パネルガラス(in_Hinban As Variant) As Variant

    If IsTateguHinban(in_Hinban) Then
        パネルガラス = left(varfncGetDetail(in_Hinban, 9), 1)
    End If
    
End Property

Public Property Get ガラスレザー種類(in_Hinban As Variant) As Variant
    
    Dim strAll As String

    If IsTateguHinban(in_Hinban) Then
        strAll = varfncGetDetail(in_Hinban, 9)
        
        Select Case Len(strAll)
            Case 2
                ガラスレザー種類 = right(strAll, 1)
            Case 4
                ガラスレザー種類 = Mid(strAll, 2, 1)
        End Select
    End If
    
End Property

Public Property Get 引戸仕様ルーバー色(in_Hinban As Variant) As Variant
    
    Dim strAll As String

    If IsTateguHinban(in_Hinban) Then
        strAll = varfncGetDetail(in_Hinban, 9)
        
        Select Case Len(strAll)
            Case Is >= 3
                引戸仕様ルーバー色 = right(varfncGetDetail(in_Hinban, 8), 2)
        End Select
    End If
    
End Property

Public Property Get 錠(in_Hinban As Variant) As Variant

    If IsTateguHinban(in_Hinban) Then
        錠 = varfncGetDetail(in_Hinban, 10)
    End If
    
End Property

Public Property Get 吊元(in_Hinban As Variant) As Variant

    If IsTateguHinban(in_Hinban) Then
        吊元 = varfncGetDetail(in_Hinban, 11)
    End If
    
End Property

Public Property Get 色(in_Hinban As Variant) As Variant

    If IsTateguHinban(in_Hinban) Then
        色 = varfncGetDetail(in_Hinban, 12)
    End If
    
End Property

Public Property Get Style(in_Hinban As Variant) As Variant
    
'2.3.0
'   →タモシリーズ追加

    Dim strStyle1 As String
    Dim strStyle2 As String
    Dim strStyle3 As String
    
    Dim strPanelGlass As String
    Dim strGlassLeather As String
    Dim strColor As String
    Dim strHinban As String
    
    Style = Null
    
    strPanelGlass = Nz(Me.パネルガラス(in_Hinban), "")
    strGlassLeather = Nz(Me.ガラスレザー種類(in_Hinban))
    strColor = Nz(Me.色(in_Hinban))
    
    If Not IsTateguHinban(in_Hinban) Then
        Exit Function
    Else
        strHinban = Replace(in_Hinban, "特 ", "")
    End If
    
    'トヨタ品番の時頭のHEを除去
    If strHinban Like "HE*-####*-*" Then
        strHinban = Mid(strHinban, 3)
    End If
    
    If IsGranArt(CStr(strHinban)) Then
        strStyle1 = "G"
    ElseIf IsTerrace(CStr(strHinban)) Then
        strStyle1 = "Y"
    ElseIf IsLUCENTE(CStr(strHinban)) Then
        If IsTamo(CStr(strHinban)) Then
            strStyle1 = "J"
        Else
            strStyle1 = "P"
        End If
    ElseIf IsPALIO(CStr(strHinban)) Then
        strStyle1 = "B"
    ElseIf IsSINA(CStr(strHinban)) Then
        strStyle1 = "T"
    ElseIf IsREALART(CStr(strHinban)) Then
        strStyle1 = "R"
    ElseIf IsAir(CStr(strHinban)) Then
        strStyle1 = "F"
    ElseIf isCaro(strHinban) Then
        strStyle1 = "A"
    ElseIf IsFs(CStr(strHinban)) Then
        strStyle1 = "K"
    ElseIf IsMonster(CStr(strHinban)) Then
        strStyle1 = "O"
    ElseIf IsCloset_Hikichigai(CStr(strHinban)) Then
        strStyle1 = "M"
        If IsSINAColor(CStr(strHinban)) Then
            strStyle2 = "S"
        ElseIf Me.IsEstyle(strHinban) Then
            strStyle2 = "E"
        Else
            strStyle2 = "C"
        End If
        strStyle3 = "4"
    ElseIf IsCloset_Slide(strHinban) Then
        strStyle1 = "M"
        strStyle2 = "P"
        strStyle3 = "3"
    ElseIf IsVertica(CStr(strHinban)) Then
        If IsEstyle(strHinban) Then
            strStyle1 = "E"
        Else
            strStyle1 = "C"
        End If
        
    ElseIf IsEstyle(strHinban) Then
        strStyle1 = "E"
        
    ElseIf strHinban Like "F*" Or strHinban Like "V*" Then
        strStyle1 = "C"
    Else
        Exit Function
    End If
    
    If strStyle2 = "" Then
        
        Select Case strPanelGlass
            Case "S", "C", "G", "M", "D", "V", "T", "E", "Q"
               strStyle2 = "G"
                
                Select Case strPanelGlass
                    Case "S", "Q"
                        strStyle3 = "1"
                    Case "C"
                        strStyle3 = "2"
                    Case "G"
                        strStyle3 = "3"
                    Case "M"
                        If strGlassLeather = "F" Then
                            strStyle3 = "7"
                        End If
                    Case "D"
                        strStyle3 = "8"
                    Case "V"
                        strStyle3 = "4"
                    Case "T"
                        strStyle3 = "6"
                    Case "E"
                        strStyle3 = "9"
                End Select
                
            Case "W"
                If strGlassLeather = "F" Then
                    strStyle2 = "F"
                Else
                    strStyle2 = "G"
                End If
                strStyle3 = "5"
                
            Case "A", "B", "O"
                If Nz(Me.建具スタイル(strHinban), "") = "B" Then
                    strStyle2 = "F"
                Else
                    strStyle2 = "G"
                End If
                
                Select Case strPanelGlass
                    Case "A"
                        strStyle3 = "1"
                    Case "B"
                        strStyle3 = "2"
                    Case "O"
                        strStyle3 = "3"
                End Select
                
            Case "S"
                strStyle2 = "A"
                
                If strGlassLeather = "C" Then
                    strStyle3 = "1"
                Else
                    strStyle3 = "2"
                End If
               
            Case "Z"
                strStyle2 = "F"
                strStyle3 = "1"
                
            Case Else
                strStyle2 = "F"
                
                If IsVertica(CStr(strHinban)) Then
                    strStyle3 = "4"
                ElseIf strStyle1 = "K" Then
                    strStyle3 = "7"
                ElseIf strGlassLeather <> "" Then
                    strStyle3 = "6"
                Else
                    strStyle3 = "1"
                End If
        End Select
                
    End If

    Style = strStyle1 & strStyle2 & strStyle3
    
    If Len(Style) <> 3 Then
        Style = Null
    End If
    
End Property

Public Function IsEstyle(in_Hinban As Variant) As Boolean
'Estyle色（ただし1608以降）
    Dim strColor As String
    
    IsEstyle = False
    
    strColor = Nz(Me.色(in_Hinban), "")
    
    Select Case strColor
        
        Case "QB", "EB", "EW", "MO", "BA", "FW", "PH"
        
            IsEstyle = True
            
    End Select
    
End Function
Private Function varfncGetDetail(in_Hinban As Variant, in_Place As Byte) As Variant
    
    Dim strReplace As String
    Dim strReplaceVar As String
    
    varfncGetDetail = Null
    If IsNull(in_Hinban) Then Exit Function
    
    strReplaceVar = "$" & CStr(in_Place)
    
    If IsTateguHinban(in_Hinban) Then
        strReplace = Reg.Replace(in_Hinban, strReplaceVar)
        If strReplace <> "" Then varfncGetDetail = strReplace
    End If
    
End Function