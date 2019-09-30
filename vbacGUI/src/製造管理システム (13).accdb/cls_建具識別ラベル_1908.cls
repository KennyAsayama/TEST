Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const constPicturePath As String = "\\db\Prog\画像\R_BRAND045\"

Private Function fncGetTateguHinban(ByVal in_Hinban As String) As String
    
    fncGetTateguHinban = Trim(Replace(in_Hinban, "特 ", ""))
    
End Function

Public Property Get 中パネル(ByVal 建具品番 As String, ByVal 個別Spec As Variant, DW As Variant) As Variant
    
    Dim strHinban As String
    
    中パネル = Null
    
    If IsNull(個別Spec) Or IsNull(DW) Then Exit Property
    
    strHinban = fncGetTateguHinban(建具品番)
    
    If right(個別Spec, 4) > "1110" Then
        If strHinban Like "F?D*" Then
            中パネル = "横帯 42×" & DW + 4
        ElseIf IsLUCENTE(strHinban) Then
            If strHinban Like "P?D??-####*-*" Or strHinban Like "P?D???-####*-*" Then
                中パネル = "横帯 42×" & DW + 6
            Else
                中パネル = Null
            End If
        Else
            中パネル = Null
        End If
    Else
        If strHinban Like "F?D*" Then
            中パネル = "ﾒﾀﾙ 36×" & [DW] + 4
        ElseIf IsLUCENTE(strHinban) Then
            If strHinban Like "P?D??-####*-*" Or strHinban Like "P?D???-####*-*" Then
                中パネル = "ﾒﾀﾙ 36×" & [DW] + 6
            Else
                中パネル = Null
            End If
        Else
            中パネル = Null
        End If
    End If
    
End Property

Public Property Get 縁張_外(ByVal 契約番号 As String, ByVal 建具品番 As String, ByVal 色 As Variant, ByVal 個別Spec As Variant) As Variant


    Dim strHinban As String
    
    縁張_外 = Null
    
    strHinban = fncGetTateguHinban(建具品番)
    
    'トヨタ
    If strHinban Like "TY*" And Not IsSINAColor(strHinban) Then
        縁張_外 = "ローカン"
        Exit Property
    End If
    
    'F/S
    If IsFs(strHinban) Then
        If Nz(色, "") = "ZZ" Then
            縁張_外 = "1C"
        Else
            縁張_外 = "2C"
        End If
        Exit Property
    End If
    
    'GranArt
    If IsGranArt(strHinban) Then
        縁張_外 = "木大手"
        Exit Property
    End If

    If IsREALART(strHinban) Or IsLUCENTE(strHinban) Or IsCloset_Slide(strHinban) Or IsZENstyle(strHinban) Then
        縁張_外 = "1mm左右"
        Exit Property
    End If

    If IsSINA(strHinban) Or IsPALIO(strHinban) Or IsCloset_Hikichigai(strHinban) Or IsMonster(strHinban) Or IsSINAColor(strHinban) Then
        縁張_外 = "エッジ"
        Exit Property
    End If
    
    縁張_外 = "ローカン"
    
    'カルロ/ジュリア
    If IsCarloGiulia(strHinban) Then
        縁張_外 = 縁張_外 & "(" & strFncFuchibariColor(strHinban, Nz(色, ""), 個別Spec) & ")"
    End If
    
End Property

Public Property Get 縁張_内(ByVal 契約番号 As String, ByVal 建具品番 As String, ByVal 色 As Variant) As Variant
    

    Dim strHinban As String
    
    縁張_内 = Null
    
    strHinban = fncGetTateguHinban(建具品番)
    
    'トヨタ
    If strHinban Like "TY*" And Not IsSINAColor(strHinban) Then
        If IsSideThru(strHinban) Or IsCenterThru(strHinban) Then
            縁張_内 = "エッジ"
        End If
        Exit Property
    End If
    
    'F框
    If strHinban Like "??A*" Or strHinban Like "??B*" Or strHinban Like "??E*" Or strHinban Like "??F*" Or strHinban Like "??W*" Then
        縁張_内 = "エッジ"
        Exit Property
    End If
    
    'ガラス
    If IsSideThru(strHinban) Or IsCenterThru(strHinban) Or IsWideThru(strHinban) Or isCaro(strHinban) Or IsG7_Flush(strHinban) Then
        '1mm
        If IsREALART(strHinban) Or IsLUCENTE(strHinban) Then
        
            縁張_内 = "1mm左右"

        Else
            縁張_内 = "エッジ"
        End If
        
        Exit Property
    End If
    
    '引き手レス
    If IsMonster(建具品番) Or IsVertica(建具品番) Then
        縁張_内 = "エッジ"
        Exit Property
    End If
    
    'F/S(KF1),物入引き違い
    '↓本来は「1C」表示だがテープ色を識別するため色を表示
    If (IsFs(建具品番) And Nz(色, "") <> "ZZ") Or IsCloset_Hikichigai(建具品番) Then
        縁張_内 = 色
        Exit Property
    End If
End Property

Public Property Get ハンドル引手センター(ByVal 建具品番 As String, ByVal 施錠 As Variant, ByVal 個別Spec As Variant, ByVal 吊元 As Variant, ByVal 扉位置 As Variant) As Variant


    Dim strHinban As String
    Dim strHandle As String
    Dim strHandleMaker As String
    
    ハンドル引手センター = Null
    
    If IsNull(個別Spec) Or IsNull(施錠) Then Exit Property
    
    If 施錠 Like "*特*" Then Exit Property
    
    strHinban = fncGetTateguHinban(建具品番)
    strHandleMaker = fncstrHandleKigoFileName(strHinban, 施錠, 個別Spec)
    
    '引き手レス,子扉は除外
    If IsMonster(strHinban) Or IsVertica(strHinban) Or IsKotobira(strHinban) Then Exit Property
        
    '引戸&物入れ引違い
    If IsHikido(strHinban) Or IsCloset_Hikichigai(strHinban) Then
        
        '自動ドアは対象外
        If 施錠 <> "ZZN" Then
            'シンクロ引手なし対応
            If IsHikiteKako(strHinban, 扉位置, 吊元, 個別Spec) Then
                If strHinban Like "???VL-*" Or strHinban Like "???VM-*" Or strHinban Like "???VN-*" Then
                    ハンドル引手センター = 1046
                ElseIf IsKabetsukeGuide(strHinban) Then
                    ハンドル引手センター = 1026
                Else
                    ハンドル引手センター = 1040
                End If
            End If
        End If
        
        Exit Property
    End If
    
    '開き戸（含親扉）
    If IsHirakido(strHinban) Or IsOyatobira(strHinban) Then

        
        If strHandleMaker Like "KAWAJ_LJ*" Then

                ハンドル引手センター = 1039
            
        ElseIf strHandleMaker Like "OLIVARI*" Or strHandleMaker Like "NAGASAWA*" Then

                ハンドル引手センター = 1037
                

        End If
        
        Exit Property

    End If
    
    'スライド収納
    If IsCloset_Slide(strHinban) Then
        ハンドル引手センター = 990
        Exit Property
    End If
    
End Property

Public Property Get ハンドル引手BS(ByVal 建具品番 As String, ByVal 施錠 As Variant, ByVal 個別Spec As Variant, ByVal 吊元 As Variant, ByVal 扉位置 As Variant) As Variant


    Dim strHinban As String
    Dim strHandle As String
    Dim strHandleMaker As String
    
    ハンドル引手BS = Null
    
    If IsNull(個別Spec) Or IsNull(施錠) Then Exit Property
    If 施錠 Like "*特*" Then Exit Property
    
    strHinban = fncGetTateguHinban(建具品番)
    strHandleMaker = fncstrHandleKigoFileName(strHinban, 施錠, 個別Spec)
    
    '引き手レス,子扉は除外
    If IsMonster(strHinban) Or IsVertica(strHinban) Or IsKotobira(strHinban) Then Exit Property
    
    'スライド収納
    If IsCloset_Slide(strHinban) Then
        ハンドル引手BS = 150
        Exit Property
    End If
    
    '引戸
    If IsHikido(strHinban) Then
    
        '自動ドア除外
        If 施錠 <> "ZZN" Then
            'シンクロ引手なし対応
            If IsHikiteKako(strHinban, 扉位置, 吊元, 個別Spec) Then
                
                If IsEndWakunashi(strHinban) Then
                        ハンドル引手BS = 61
                Else
                    If strHandleMaker Like "PC422*" Or strHandleMaker Like "BEST*" Or strHandleMaker Like "OKUDAIRATK*" Then
                        ハンドル引手BS = 38
                        
                    ElseIf strHandleMaker Like "PC210*" Then
                        ハンドル引手BS = 42
                        
                    End If
                End If
            End If
        End If
        
        Exit Property
    End If
    
    '開き戸（含親扉）
    If IsHirakido(strHinban) Or IsOyatobira(strHinban) Then

        strHandle = left(施錠, 2)
        
        Select Case fncstrHandleKigoFileName(strHinban, strHandle & "N", Nz(個別Spec, ""))
            Case "KAWAJ_LJ", "KURAMAE", "OLIVARI", "NAGASAWA"
                ハンドル引手BS = 51
        End Select
            
            
        Exit Property

        
    End If

End Property

Public Property Get 鎌錠センター(ByVal 建具品番 As String, ByVal 施錠 As Variant, ByVal 個別Spec As Variant, ByVal 吊元 As Variant, ByVal 扉位置 As Variant) As Variant


    Dim strHinban As String
    Dim strHandle As String
    Dim strLock As String
    Dim strHandleMaker As String
    
    鎌錠センター = Null
    
    If IsNull(個別Spec) Or IsNull(施錠) Then Exit Property
    
    If Len(施錠) < 2 Then Exit Property
    
    If 施錠 Like "*特*" Then Exit Property
    
    strHinban = fncGetTateguHinban(建具品番)
    strLock = right(施錠, 1)
    strHandleMaker = fncstrHandleKigoFileName(strHinban, 施錠, 個別Spec)
    
    '錠無しは除外
    If strLock = "N" Then Exit Property
    
    '子扉は除外
    If IsKotobira(strHinban) Then Exit Property
    
    '引戸
    If IsHikido(strHinban) Then
        'アウトセット引戸錠
        If IsEndWakunashi(strHinban) Then
            If strHinban Like "???VN-*" Then
                鎌錠センター = 1046
            Else
                鎌錠センター = 1040
            End If
        ElseIf strHinban Like "???VL-*" Or strHinban Like "???VM-*" Then
            鎌錠センター = 846
        Else
            鎌錠センター = 840
        End If

        
        Exit Property
        
    'ワンニャンハンドル
    ElseIf strHandleMaker Like "NAGASAWA*" And right(CStr(施錠), 1) <> "N" Then
        
        '現行品
        If right(CStr(施錠), 1) = "C" Then
            鎌錠センター = 937
        Else
            鎌錠センター = 917
        End If

    'シリンダー錠
    ElseIf strHandleMaker = "KAWAJ_LJ_Lock_Cylinder" Then
    
        鎌錠センター = 937
        
    End If
    
End Property

Public Property Get 鎌錠BS(ByVal 建具品番 As String, ByVal 施錠 As Variant, ByVal 個別Spec As Variant) As Variant


    Dim strHinban As String
    Dim strHandle As String
    Dim strLock As String
    Dim strHandleMaker As String
    
    鎌錠BS = Null
    
    If IsNull(施錠) Then Exit Property
    If Len(施錠) < 2 Then Exit Property
    
    If 施錠 Like "*特*" Then Exit Property
    
    strHinban = fncGetTateguHinban(建具品番)
    strLock = right(施錠, 1)
    strHandleMaker = fncstrHandleKigoFileName(strHinban, 施錠, 個別Spec)
    
    '錠無しは除外
    If strLock = "N" Then Exit Property
    
    '子扉は除外
    If IsKotobira(strHinban) Then Exit Property
    
    '引戸
    If IsHikido(strHinban) Then
        'アウトセット引戸錠
        If IsEndWakunashi(strHinban) Then
            If strHinban Like "???DU-*" Or strHinban Like "???KU-*" Then
                鎌錠BS = 48
            ElseIf strHinban Like "???DN-*" Or strHinban Like "???VN-*" Then
                鎌錠BS = 101
            End If
        Else
            鎌錠BS = 38
        End If
       
        Exit Property
        
    'ワンニャンハンドル
    ElseIf fncbol_Handle_WanNyan(CStr(施錠), CStr(個別Spec)) And right(CStr(施錠), 1) <> "N" Then
        鎌錠BS = 51
        
    'シリンダー錠
    ElseIf strHandleMaker = "KAWAJ_LJ_Lock_Cylinder" Then
    
        鎌錠BS = 51
        
    End If

End Property

Public Property Get 溝(ByVal 建具品番 As String, ByVal 施錠 As Variant, ByVal 個別Spec As Variant, ByVal 扉位置 As Variant, ByVal 吊元 As Variant) As Variant


    Dim strHinban As String
    
    溝 = Null
    
    If IsNull(個別Spec) Or IsNull(施錠) Then Exit Property
    
    strHinban = fncGetTateguHinban(建具品番)
    
    'DQはレール無し
    If strHinban Like "*DQ-*" Or strHinban Like "*VQ-*" Then
        Exit Property
    End If
    
    If IsYukazukeRail(strHinban) Then

        溝 = "上5.5"

'    ElseIf strHinban Like "???DH-*" Or strHinban Like "???DJ-*" Then
'
'        溝 = "下4"
        
    ElseIf strHinban Like "???SH-*" Or strHinban Like "???SJ-*" Then

        If Nz(扉位置, 0) = 3 Then
            溝 = "20.8"
            
        ElseIf Nz(吊元, "") = "L" And Nz(扉位置, 0) = 2 Then
            溝 = "20.8"
            
        ElseIf Nz(吊元, "") = "R" And Nz(扉位置, 0) = 1 Then
            溝 = "20.8"
            
        End If
        
'    ElseIf strHinban Like "???DF-*" And Not Nz(扉位置, 0) = 3 Then
'
'        溝 = "下4"
'
'    ElseIf strHinban Like "???VF-*" And Not Nz(扉位置, 0) = 3 Then
'
'        溝 = "下4"
    
    ElseIf strHinban Like "???SF-*" And Nz(扉位置, 0) = 3 Then

        溝 = "20.8"

    ElseIf strHinban Like "???TF-*" And Nz(扉位置, 0) = 3 Then

        溝 = "20.8"
    
    ElseIf IsVertica(strHinban) Then

        溝 = "5.5"

    ElseIf strHinban Like "???DC-*" Or strHinban Like "???DE-*" Or strHinban Like "???DP-*" Or strHinban Like "???DT-*" Or strHinban Like "???DU-*" Then

        If IsG7_Flush(strHinban) Then
            溝 = "5.5"
        ElseIf isCaro(strHinban) Then
            溝 = "12Φ"
        Else
            '溝共通化
            If left(Nz(個別Spec, ""), 3) = "BRD" Then
                'BRD1908以降
                溝 = "5.5"
                
            ElseIf left(Nz(個別Spec, ""), 3) = "TSC" And right(Nz(個別Spec, "XXX0000"), 4) >= "1801" Then
                'TSC1801以降
                溝 = "4.5"
            Else
                'BRD以外
                溝 = ""
            End If
        End If

    'MEクロゼットは[4.5]固定
'    ElseIf IsCloset_Hikichigai(strHinban) Then
'        溝 = "4.5"
    End If

End Property

Public Property Get 引込取手センター(ByVal 建具品番 As String, ByVal 施錠 As Variant, ByVal 個別Spec As Variant, ByVal 吊元 As Variant, ByVal 扉位置 As Variant) As Variant

    Dim strHinban As String
    
    引込取手センター = Null
    
    If IsNull(個別Spec) Or IsNull(施錠) Then Exit Property
    
    If 施錠 Like "*特*" Then Exit Property
    
    strHinban = fncGetTateguHinban(建具品番)
    
    'ポケット戸のみ
    If strHinban Like "???DP-*" Then
        If right(個別Spec, 4) >= "1307" Then
            引込取手センター = 1040
        Else
            If 施錠 Like "W*" Then
                引込取手センター = Nz(Me.ハンドル引手センター(strHinban, 施錠, 個別Spec, 吊元, 扉位置), 0) - 100
            ElseIf 施錠 Like "G*" Or 施錠 Like "H*" Then
                引込取手センター = Nz(Me.ハンドル引手センター(strHinban, 施錠, 個別Spec, 吊元, 扉位置), 0) - 60
            End If
        End If
    End If
    
    '異常データはNullに置換え
    If 引込取手センター < 0 Then 引込取手センター = Null
    
End Property

Public Property Get 子扉ラッチ受けセンター(ByVal 建具品番 As String, ByVal 施錠 As Variant, ByVal 個別Spec As Variant) As Variant


    Dim strHandle As String
    Dim strHinban As String
    
    strHinban = fncGetTateguHinban(建具品番)
    
    子扉ラッチ受けセンター = Null
    
    If IsNull(個別Spec) Or IsNull(施錠) Then Exit Property
    
    If 施錠 Like "*特*" Then Exit Property
    
    'コード3桁化以降
    If right(個別Spec, 4) >= "1410" Then
        strHandle = left(施錠, 2)
        
        '蔵前のみ1011,カワジュンU 1046 その他は1052
        If IsKotobira(strHinban) Then
            Select Case left(strHandle, 1)
                Case "A"
                    Select Case strHandle
                        Case "AL", "AM"
                            子扉ラッチ受けセンター = 1011
                        Case "AU"
                            子扉ラッチ受けセンター = 1046
                        Case Else
                            子扉ラッチ受けセンター = 1052
                    End Select

                Case "B"
                    Select Case strHandle
                        Case "BB", "BK", "BG"
                            子扉ラッチ受けセンター = 1052
                        Case Else
                            子扉ラッチ受けセンター = 1011
                    End Select

                Case "C"
                    Select Case strHandle
                        Case "CL", "CM"
                            子扉ラッチ受けセンター = 1011
                        Case Else
                            子扉ラッチ受けセンター = 1052
                    End Select
                
                Case "D"
                    Select Case strHandle
                        Case "DP", "DQ"
                            子扉ラッチ受けセンター = 1011
                        Case Else
                            子扉ラッチ受けセンター = 1052
                    End Select
                   
                Case "H"
                    子扉ラッチ受けセンター = 1059
                    
                Case Else
                    子扉ラッチ受けセンター = 1052
                            
                    
            End Select
        End If
        
    '1404以前（2桁)
    Else
        strHandle = left(施錠, 1)
        
        If IsKotobira(strHinban) Then
            Select Case strHandle
                Case "L", "M"
                    子扉ラッチ受けセンター = 1011
                Case "U"
                    子扉ラッチ受けセンター = 1046
                Case Else
                    子扉ラッチ受けセンター = 1052
            End Select
        End If
    End If

End Property

Public Property Get 子扉鎌錠受けセンター(ByVal 建具品番 As String, ByVal 施錠 As Variant, ByVal 個別Spec As Variant) As Variant
    
    Dim strHandle As String
    Dim strLock As String
    Dim strHinban As String
    
    子扉鎌錠受けセンター = Null
    
    strHinban = fncGetTateguHinban(建具品番)
    
    If IsNull(個別Spec) Or IsNull(施錠) Then Exit Property
    
    If 施錠 Like "*特*" Then Exit Property
    
        strLock = right(施錠, 1)
    
    'コード3桁化以降
    If right(個別Spec, 4) >= "1410" Then
        strHandle = left(施錠, 2)
        
        '長沢（ペット）、カワジュンUハンドルのみ
        If IsKotobira(strHinban) Then
            Select Case strHandle
                Case "AE", "AF", "AO", "AN"
                    子扉鎌錠受けセンター = 926
                Case "AU"
                    子扉鎌錠受けセンター = 1157
            End Select
        End If
    '1404以前（2桁)
    Else
        strHandle = left(施錠, 1)
        
        '長沢（ペット）、カワジュンUハンドルのみ
        If IsKotobira(strHinban) Then
            Select Case strHandle
                Case "E", "F", "O", "N"
                    If right(個別Spec, 4) >= "1404" Then
                        子扉鎌錠受けセンター = 926
                    End If
                Case "U"
                    子扉鎌錠受けセンター = 1157
            End Select
        End If
    End If

End Property

Public Property Get カセット加工図パス(ByVal 建具品番 As String, ByVal 施錠 As Variant, ByVal 個別Spec As Variant, ByVal 吊元 As Variant, ByVal 扉位置 As Variant) As Variant


    Dim strPictureName As String
    Dim strHinban As String
        
    Const constPath_Handle As String = "Handle\40mm\"
    Const constFileExtension As String = ".png"
    
    カセット加工図パス = Null
    strPictureName = ""
    
    strHinban = fncGetTateguHinban(建具品番)
    
    If IsNull(個別Spec) Or IsNull(施錠) Then Exit Property
    
    If 施錠 Like "*特*" Then Exit Property
    
    'シンクロ引手なし対応
    If Not IsHikiteKako(strHinban, 扉位置, 吊元, 個別Spec) Then Exit Property
        
    strPictureName = fncstrHandleKigoFileName(strHinban, 施錠, 個別Spec)
    
    If strPictureName <> "" Then
        カセット加工図パス = constPicturePath & constPath_Handle & strPictureName & constFileExtension
    End If
    
    If Not IsNull(カセット加工図パス) Then
        If Dir(カセット加工図パス) = "" Then
           ' Debug.Print カセット加工図パス
        End If
    End If
    
End Property

Public Property Get 戸開き図パス(ByVal 建具品番 As String, ByVal 開閉様式 As Variant, ByVal 吊元 As Variant, ByVal 扉位置 As Integer) As Variant
'3.0.0
'   →扉位置RCが間違っていたため訂正

    Dim strTobiraichi As String
    Dim strPath As String
    Dim strTsurimoto As String
    Dim strHinban As String
    
    Const constPicturePathtealth As String = "stealth\"
    Const strPath_Inset As String = "inset\"
    Const strPath_HikiteLess As String = "HikiteLess\"
    Const strPath_Monster As String = "Monster\"
    
    Const txt_Extension As String = ".png"
    
    戸開き図パス = Null
    strHinban = fncGetTateguHinban(建具品番)
    
    If strHinban Like "?Z???-*" Or strHinban Like "?Y???-*" Or strHinban Like "?T???-*" Then 'インセット品番
        strPath = constPicturePath & strPath_Inset
    Else
        strPath = constPicturePath & constPicturePathtealth
    End If

    If IsVertica(strHinban) Or IsMonster(strHinban) Then
        strPath = strPath & strPath_HikiteLess
        
        If IsMonster(strHinban) And IsSynchro(strHinban) Then
            strPath = strPath & strPath_Monster
        End If
        
    End If

    strTobiraichi = ""
    
    If IsNull(吊元) Then
        strTsurimoto = "Z"
    Else
        strTsurimoto = 吊元
    End If
    
    If Not IsNull(扉位置) Then
        If 扉位置 = 0 Then
            strTobiraichi = 吊元
        ElseIf 扉位置 = 1 Then
            strTobiraichi = "L"
        ElseIf 扉位置 = 2 Then
            strTobiraichi = "R"
        ElseIf 扉位置 Like 3 Then
            strTobiraichi = "C"
        ElseIf 扉位置 Like 4 Then
            strTobiraichi = "LC"
        ElseIf 扉位置 Like 5 Then
            strTobiraichi = "RC"
        End If
    End If
    
    '親子扉（子扉は引数渡しの時点でLRが反転しているシステムがあるので注意。ここで反転させないこと
    If IsOyatobira(strHinban) Or IsKotobira(strHinban) Then

        戸開き図パス = strPath & "DO-" & strTsurimoto & txt_Extension
        
    '開き戸（IsHirakido関数は親子を含むので親子はこれより上で条件処理を行うこと）
    ElseIf IsHirakido(strHinban) Then

        戸開き図パス = strPath & "DA-" & strTsurimoto & txt_Extension
    
    '1枚引戸 吊元(L,R)あり
    ElseIf IsHikido(strHinban) And 扉位置 = 0 Then
        
        戸開き図パス = strPath & 開閉様式 & "-" & strTsurimoto & txt_Extension

    '2枚以上引戸吊元無し
    ElseIf IsHikido(strHinban) And 扉位置 <> 0 And strTsurimoto = "Z" Then

        If strTobiraichi <> "" Then 戸開き図パス = strPath & 開閉様式 & "_" & strTobiraichi & txt_Extension
    
    '2枚以上引戸 吊元(L,R)あり
    ElseIf IsHikido(strHinban) And 扉位置 <> 0 And strTsurimoto <> "Z" Then

        If strTobiraichi <> "" Then 戸開き図パス = strPath & 開閉様式 & "-" & strTsurimoto & "_" & strTobiraichi & txt_Extension
    '?
    ElseIf strHinban Like "??CME-*" Or strHinban Like "??CSA-*" Then
    
        戸開き図パス = strPath & Mid(strHinban, 3, 3) & txt_Extension
    End If

    If Not IsNull(戸開き図パス) Then
        If Dir(戸開き図パス) = "" Then
            Debug.Print 戸開き図パス
        End If
    End If

End Property

Public Function 建具設計備考(ByVal 建具品番 As String, ByVal 個別Spec As Variant, ByVal 設計備考 As Variant) As Variant


    Dim strBikou As String
    
    If IsCloset_Hikichigai(建具品番) Or IsCloset_Slide(建具品番) Then
    
        If Not IsNull(個別Spec) Then
        
            If right(個別Spec, 4) <= "1608" Then
                strBikou = " 裏面はPH色"
            Else
                If IsCloset_Slide(建具品番) Or right(個別Spec, 4) < "1808" Then
                    strBikou = " 裏面はSH色"
                End If
            End If
        End If
        
    ElseIf IsG7_Flush(建具品番) Then
        strBikou = " 中板ﾐﾗｰ側はSH色"
        
    ElseIf IsCaro_Panel(建具品番) Then
        strBikou = " 中板②はMDF素地"
        
    End If
    
    If strBikou <> "" Then
        建具設計備考 = Nz(設計備考, "") & strBikou
    Else
        建具設計備考 = 設計備考
    End If

End Function

Public Property Get ルーバー溝加工位置(ByVal 建具品番 As String, ByVal 開閉様式 As String, ByVal ガラスレザー区分 As String, ByVal DH As Variant, ByVal 個別Spec As Variant) As Variant
'3.0.0 ADD
    
    Dim strHinban As String
    
    ルーバー溝加工位置 = Null
    
    strHinban = fncGetTateguHinban(建具品番)
    
    If IsAir(strHinban) Then
        
        If ガラスレザー区分 = "L" Then 'FA2型のみ
            
            '開き戸
            If IsHirakido(strHinban) Then
                
                'ステルス下がり壁
                If 開閉様式 Like "??S" Then
                    Select Case DH
                        Case 1381
                            ルーバー溝加工位置 = "828"
                        Case 1381.5 To 1531
                            ルーバー溝加工位置 = "932"
                        Case 1531.5 To 2681
                            ルーバー溝加工位置 = "1062"
                        Case Else
                            ルーバー溝加工位置 = "?"
                    End Select
                Else
                    Select Case DH
                        Case 1374
                            ルーバー溝加工位置 = "828"
                        Case 1374.5 To 1524
                            ルーバー溝加工位置 = "932"
                        Case 1524.5 To 2674
                            ルーバー溝加工位置 = "1062"
                        Case Else
                            ルーバー溝加工位置 = "?"
                    End Select
                End If
                
            ElseIf IsHikido(strHinban) Then
                Select Case DH
                        Case 1367 To 1381
                            ルーバー溝加工位置 = "849"
                        Case 1381.5 To 1531
                            ルーバー溝加工位置 = "953"
                        Case 1531.5 To 2681
                            ルーバー溝加工位置 = "1057"
                        Case Else
                            ルーバー溝加工位置 = "?"
                    End Select
            End If
            
        End If
    End If
    
End Property