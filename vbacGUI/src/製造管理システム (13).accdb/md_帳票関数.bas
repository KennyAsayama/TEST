Option Compare Database
Option Explicit

'20160825 K.Asayama ADD 引数が多くなったためユーザ定義型に変更
Type KidoriData
'   寸法
    out_dblShinAtsu As Variant      '芯厚
    out_dblsan As Variant           '上下桟
    out_dblgakuyoko1 As Variant     '額横（１）
    out_dblgakuYoko2 As Variant     '額横（２）
    '20180201 K.Asayama ADD
    out_dblGakuyokoLVL30 As Variant '額横（LVL30）
    '20180201 K.Asayama ADD END
    out_dblhashira As Variant       '柱
    '20170517 K.Asayama ADD
    out_dblhashira2 As Variant      '柱（２）
    '20170517 K.Asayama ADD END
    out_dblgakutate1 As Variant     '額縦（１）
    out_dblgakutate2 As Variant     '額縦（２）
    out_dblgakutate3 As Variant     '額縦（３）
    out_dbltegakeShurui As Variant  '手掛け種類
    out_dbltegake As Variant        '手賭け
    out_dblsode1 As Variant         '袖（１）
    out_dblsode2 As Variant         '袖（２）
    out_dbldaboshitaji As Variant   'ダボ下地
    out_dblCupShitaji As Variant    'カップ下地
'   本数
    out_intsan As Variant           '上下桟
    out_intgakuyoko1 As Variant     '額横（1）
    out_intgakuyoko2 As Variant     '額横（２）
    '20180201 K.Asayama ADD
    out_intgakuyokoLVL30 As Variant '額横（LVL30）
    '20180201 K.Asayama ADD END
    out_inthashira As Variant       '柱
    '20170517 K.Asayama ADD
    out_inthashira2 As Variant      '柱（２）
    '20170517 K.Asayama ADD END
    out_intgakutate1 As Variant     '額縦（１）
    out_intgakutate2 As Variant     '額縦（２）
    out_intgakutate3 As Variant     '額縦（３）
    out_inttegake As Variant        '手掛け
    out_intsode1 As Variant         '袖（１）
    out_intsode2 As Variant         '袖（２）
    out_intdaboshitaji As Variant   'ダボ下地
    out_intcupshitaji As Variant    'カップ下地
'   中板
    out_dblShinAtsu_N As Variant    '芯厚
    out_dblsan_N As Variant         '上下桟
    out_dblgakuyoko1_N As Variant   '額横
    out_dblhashira_N As Variant     '柱
    out_dblhashira2_N As Variant    '柱（２）
    out_dblhashiraSt_N As Variant   '柱(下)
    out_dblYokosan_N As Variant     '横桟
    '20180201 K.Asayama ADD
    out_dblsanH2_N As Variant       '上下桟(LVL45)
'   中板本数
    out_intsanh_N As Variant        '上下桟
    '20180201 K.Asayama ADD
    out_intsanh2_N As Variant   '上下桟(LVL45)
    out_inthashiraH2_N As Variant   '柱（２）
    '20180201 K.Asayama ADD END
    out_intgakuyokoH1_N As Variant  '額横
    out_inthashiraH_N As Variant    '柱
    out_inthashiraStH_N As Variant  '柱(下)
    out_intYokosanh_N As Variant    '横桟
'   図面
    out_strShingumizu As Variant    '芯組詳細図
'20160825 K.Asayama ADD
End Type

Public Function bolFncKidoriData(ByVal varSpec As Variant, ByVal in_strHinban As String, intMaisu As Integer, ByVal dblDW As Double, ByVal dblDH As Double, ByVal strAkarimado As Variant, ByVal varHandle As Variant _
                            , ByRef KidoriSunpo As KidoriData) As Boolean
                            
'   *************************************************************
'   木取り寸法表関数
'   'ADD by Asayama 20150917
'   '20160308 K.Asayama ADD 引数追加(柱(下),横桟）
'   '20160825 K.Asayama Change
'       →引数をTYPE型に変更
'       →1608仕様対応
'   '20170517 K.Asayama Change
'       →Terrace用柱追加
'   '20180201 K.Asayama Change
'       →1801仕様対応

'   戻り値:Boolean
'       →True              照合OK　数値戻し
'       →True              照合NG　数値なし
'
'    Input項目
'       varspec             個別Spec
'       strHinban           建具品番
'       intMaisu            枚数
'       dblDW               DW
'       dblDH               DH
'       strAkariMado        明り窓
'       varHandle           施錠
    
'    Output項目
'      寸法
'       out_dblShinAtsu     芯厚
'       out_dblsan          上下桟
'       out_dblgakuyoko1    額横（１）
'       out_dblgakuyoko2    額横（２）
'       out_dblgakuyokoLVL30額横（LVL30）
'       out_dblhashira      柱
'       out_dblHashira2     柱（２）
'       out_dblgakutate1    額縦（１）
'       out_dblgakutate2    額縦（２）
'       out_dblgakutate3    額縦（３）
'       out_dbltegakeShurui 手掛け種類
'       out_dbltegake       手賭け
'       out_dblsode1        袖（１）
'       out_dblsode2        袖（２）
'       out_dbldaboshitaji  ダボ下地
'       out_dblcupshitaji   カップ下地
'      本数
'       out_intsan          上下桟
'       out_intgakuyoko1    額横（1）
'       out_intgakuyoko2    額横（２）
'       out_intgakuyokoLVL30額横（LVL30）
'       out_inthashira      柱
'       out_intHashira2     柱（２）
'       out_intgakutate1    額縦（１）
'       out_intgakutate2    額縦（２）
'       out_intgakutate3    額縦（３）
'       out_inttegake       手掛け
'       out_intsode1        袖（１）
'       out_intsode2        袖（２）
'       out_intdaboshitaji  ダボ下地
'       out_intcupshitaji   カップ下地
'      中板
'       out_dblShinAtsu_N   芯厚
'       out_dblsan_N        上下桟
'       out_dblgakuyoko1_N  額横
'       out_dblhashira_N    柱
'       out_dblhashiraSt_N  柱(下)
'       out_dblYokosan_N    横桟
'      中板本数
'       out_intsanh_N       上下桟
'       out_intgakuyokoh1_N 額横
'       out_inthashirah_N   柱
'       out_inthashiraSth_N 柱(下)
'       out_intYokosanh_N   横桟
'      図面
'       out_strShingumizu   芯組詳細図
'   *************************************************************


    Dim dblSan As Double, dblGakuYoko1 As Double, dblGakuYoko2 As Double, dblHashira As Double, dblDaboShitaji As Double, dblCupShitaji As Double, dblSode1 As Double, dblSode2 As Double
    Dim dblGakutate1 As Double, dblGakutate2 As Double, dblGakutate3 As Double
    Dim dblTegake As Double, dblTegakeShurui As Double
    
    '20170517 K.Asayama ADD
    Dim dblHashira2 As Double
    
    Dim dblSan_N As Double, dblGakuYoko1_N As Double, dblHashira_N As Double
    '20151211 K.Asayama ADD
    Dim dblHashiraShita_N As Double, dblYokoSan_N As Double
    
    Dim dblShinAtsu As Double, dblShinAtsu_N As Double
    
    '20180205 K.Asayama ADD
    Dim dblGakuYokoLVL30 As Double
    Dim dblhashira2_N As Double
    Dim dblsanH2_N As Double
    
    Dim intSanH As Integer, intGakuYokoH1 As Integer, intGakuYokoH2 As Integer, intHashiraH As Integer, intDaboShitajiH As Integer, intCupShitajiH As Integer, intSode1H As Integer, intSode2H As Integer
    Dim intGakutateH1 As Integer, intGakutateH2 As Integer, intGakutateH3 As Integer
    Dim intTegakeH As Integer
    
    '20170517 K.Asayama ADD
    Dim intHashiraH2 As Integer
    
    Dim intSanH_N As Integer, intGakuYokoH1_N As Integer, intHashiraH_N As Integer
    '20151211 K.Asayama ADD
    Dim intHashiraShitaH_N As Integer, intYokoSanH_N As Integer
    
    '20180205 K.Asayama ADD
    Dim intGakuYokoLVL30 As Integer
    Dim intsanh2_N As Integer
    Dim inthashiraH2_N As Integer
    
    Dim strShingumizu As String
    
    '20180205 K.Asayama ADD
    Dim strHinban As String
    strHinban = Replace(in_strHinban, "特 ", "")
    
'   *************************************************************
'   共通項目の挿入
'       手掛けの長さ、種類、本数
'       クローゼットとAUﾊﾝﾄﾞﾙはさらに下のモジュールで情報が上書きされる場合があり
'       20170517 K.Asayama テラスの框は品番毎に違うので下のモジュールで情報が上書きされる
'   *************************************************************
'   20160308 K.Asayama Change 手掛けの誤り修正（寸法、モンスター、グランアート、Air[シナ])
    If IsKotobira(strHinban) Then
        
        dblTegake = 0
        
    '引き手レス
    ElseIf IsVertica(strHinban) Then
        '20160825 K.Asayama Change
        'dblTegake = 0
        dblTegake = 60
        '20160825 K.Asayama Cahnge END
    'モンスター
    ElseIf IsMonster(strHinban) Then
    
        dblTegake = 0
        
    
    'U(AU)ハンドル
    ElseIf fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" _
        Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
        
        dblTegake = 110
        
    
    ElseIf IsSideThru(strHinban) Then
        '20160825 K.Asayama ADD
        'If IsREALART(strHinban) Or IsLUCENTE(strHinban) Then
        If IsHikido(strHinban) Then
            dblTegake = 60
    
        ElseIf IsREALART(strHinban) Or IsLUCENTE(strHinban) Then
        '20160825 K.Asayama ADD END
            'dblTegake = 87.5
            dblTegake = 90.5
        'ElseIf IsSINA(strHinban) Or IsPALIO(strHinban) Then
        ElseIf IsSINA(strHinban) Or IsSINAColor(strHinban) Or IsPALIO(strHinban) Then
            'dblTegake = 88.5
            dblTegake = 91.5
        ElseIf IsGranArt(strHinban) Then
            dblTegake = 73.5
        Else
            dblTegake = 90
        End If
        
    '20160825 K.Asayama ADD
    ElseIf IsAir(strHinban) Then
        If IsHikido(strHinban) Then
            dblTegake = 60
        Else
            If IsSINAColor(strHinban) Then
                dblTegake = 91.5
            Else
                dblTegake = 90
            End If
        End If
    '20160825 K.Asayama ADD END
    Else
    
        dblTegake = 60
    
    End If
    
    '20160825 K.Asayama Change
    'dblTegakeShurui = dblfncTekake_Shurui(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, "")))
    dblTegakeShurui = dblfncTekake_Shurui(strHinban, CStr(Nz(varHandle, "")), CStr(Nz(varSpec, "")))
    '20160825 K.Asayama Change End
    
'    20161121 K.Asayama Change 手掛け本数改定
'    引戸は開き戸と同じに戻す
'    以下は例外
'       1.アウトセットエンド枠無し引戸上吊り(DU,KU)の錠付
'       2.アウトセットエンド枠無し引戸床付けレール(DN)
'       3.2枚引き込み(DH)
'       4.3枚扉
'       5.ヴェルチカ

'    If IsKotobira(strHinban) Then
'
'        intTegakeH = 0
'
'    '引き手レス
'    ElseIf IsVertica(strHinban) Then
'
'        '20160825 K.Asayama Change
'        'intTegakeH = 0
'        intTegakeH = intMaisu
'        '20160825 K.Asayama Change END
'
'    '20160825 K.Asayama Change 新仕様対応
'    'ElseIf strHinban = "*DH-####*-*" Or strHinban = "*DF-####*-*" Or strHinban = "*DJ-####*-*" Or strHinban = "*DQ-####*-*" Then
'    ElseIf IsHikido(strHinban) Then
'
'        intTegakeH = intMaisu * 2
'    '20160825 K.Asayama Change END
'
'    Else
'
'        intTegakeH = intMaisu
'
'    End If
    
    '子扉はなし
    If IsKotobira(strHinban) Then
        intTegakeH = 0
    
    'アウトセット床付けレールエンド枠なし(DN)は2本
    '20170105 K.Asayama Change
'    ElseIf strHinban Like "*DN-####*-*" Then
    ElseIf strHinban Like "*DN-####*-*" Or strHinban Like "*VN-####*-*" Then
    '20170105 K.Asayama Change END
    
        '20170517 K.Asayama Change
        'intTegakeH = 2 * intMaisu
        If IsTerrace(strHinban) Then
            intTegakeH = 1 * intMaisu
        Else
            intTegakeH = 2 * intMaisu
        End If
        '20170517 K.Asayama Change End
    
    'アウトセットエンド枠なしで錠付は2本
    ElseIf IsEndWakunashi_Jou(strHinban) Then
        intTegakeH = 2 * intMaisu
    
    '3枚扉と2枚引き込み(DH)
    '20170105 K.Asayama Change
'    ElseIf strHinban Like "*DH-####*-*" Or strHinban Like "*DF-####*-*" Or strHinban Like "*DJ-####*-*" Or strHinban Like "*DQ-####*-*" Then
    ElseIf strHinban Like "*DH-####*-*" Or strHinban Like "*DF-####*-*" Or strHinban Like "*DJ-####*-*" Or strHinban Like "*DQ-####*-*" Or strHinban Like "*VF-####*-*" Or strHinban Like "*VQ-####*-*" Then
    '20170105 K.Asayama Change END
        intTegakeH = 2 * intMaisu
    
    Else
    
        intTegakeH = intMaisu
        
    End If
    
    'ヴェルチカは結果から1*枚数を引く
    If intTegakeH > 0 And IsVertica(strHinban) Then
    
        intTegakeH = intTegakeH - (1 * intMaisu)
    End If
    
    '20161121 K.Asayama Change END
    
'   *************************************************************
'   品番別データの挿入
'   （クローゼットと建具で品番が被っているのでクローゼットを先に処理）
'   *************************************************************
    
'   *MC1/ME1/MZ1*************************************************
    If strHinban Like "F?CME-####F*-*" Then
        
        dblShinAtsu = 30.2
        dblSan = dblDW + 5
        dblGakuYoko1 = dblDW - 245
        intHashiraH = 2 * intMaisu
        dblSode1 = 60
        intSode1H = 10

        
        Select Case dblDH
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                dblGakuYokoLVL30 = dblDW - 55
                intGakuYokoLVL30 = 2 * intMaisu
                dblGakutate3 = 150
                intGakutateH3 = 8 * intMaisu

            Case 2530 To 2589
                intSanH = 6 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 174
                'strShingumizu = "SS-39"
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                'strShingumizu = "SS-37"
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 1 * intMaisu
                dblHashira = dblDH - 114
                'strShingumizu = "SS-37"
                
        End Select
        
        If dblDH > 2589 Then
            dblCupShitaji = 60
            intCupShitajiH = 6 * intMaisu
        ElseIf dblDH <= 2529 Then
            dblCupShitaji = 35
            intCupShitajiH = 2 * intMaisu
        End If

        
'   *MC1/ME1/MZ1(ミラー)*****************************************
    ElseIf strHinban Like "F?CME-####M*-*" Then
        
        '両側ミラー
        If strHinban Like "*-####MM*-*" Then
            
            dblTegake = 56.5
            intTegakeH = 1 * intMaisu
            
            dblShinAtsu = 30.2
            dblSan = dblDW - 320
            '20160115 K.Asayama Change 寸法間違い修正（笹原課長より）
            'dblGakuYoko1 = dblDW - 596.5
            dblGakuYoko1 = dblDW - 631.5
            '20160115 K.Asayama Change End
            intHashiraH = 2 * intMaisu
            dblSode1 = 56.5
            intSode1H = 2 * intMaisu
            
            dblShinAtsu_N = 14.8
            dblSan_N = 388
            dblGakuYoko1_N = 328
            intHashiraH_N = 2 * intMaisu
            
            dblsanH2_N = dblSan_N
            intsanh2_N = 1 * intMaisu
            
            Select Case dblDH
                '20180205 K.Asayama ADD
                Case 2589.5 To 2689
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 2 * intMaisu
                    dblHashira = dblDH - 114
                    dblGakutate3 = 150
                    intGakutateH3 = 4 * intMaisu
                    
                    intSanH_N = 2 * intMaisu
                    intGakuYokoH1_N = 7 * intMaisu
                    dblHashira_N = dblDH - 100
                    
                Case 2530 To 2589
                    intSanH = 6 * intMaisu
                    intGakuYokoH1 = 2 * intMaisu
                    dblHashira = dblDH - 174
                    
                    intSanH_N = 4 * intMaisu
                    intGakuYokoH1_N = 7 * intMaisu
                    dblHashira_N = dblDH - 160
                    
                    'strShingumizu = "SS-40"
                    
                Case 1801 To 2529
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 2 * intMaisu
                    dblHashira = dblDH - 114
                    
                    intSanH_N = 2 * intMaisu
                    intGakuYokoH1_N = 6 * intMaisu
                    dblHashira_N = dblDH - 100
                    
                    'strShingumizu = "SS-38"
                    
                Case Is <= 1800
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 1 * intMaisu
                    dblHashira = dblDH - 114
                    
                    intSanH_N = 2 * intMaisu
                    intGakuYokoH1_N = 5 * intMaisu
                    dblHashira_N = dblDH - 100
                    
                    'strShingumizu = "SS-38"
                    
            End Select
            
            '20180205 K.Asayama Change
            If dblDH > 2589 Then
                dblCupShitaji = 60
                intCupShitajiH = 2 * intMaisu
            
            ElseIf dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If
        
        '片側ミラー（帳票レイアウトが対応していないため見送り）
        Else
        
            dblTegake = 0: intTegakeH = 0
            
            Select Case dblDH
                Case 2530 To 2589
                    'strShingumizu = "SS-39/40"
                Case 1801 To 2529
                    'strShingumizu = "SS-37/38"
                Case Is <= 1800
                    'strShingumizu = "SS-37/38"
            End Select
            
        End If
                
'   *MS1*********************************************************
    ElseIf strHinban Like "T?CME-####F*-*" Then

        dblShinAtsu = 30.2
        dblSan = dblDW + 5
        dblGakuYoko1 = dblDW - 245
        intHashiraH = 2 * intMaisu
        dblSode1 = 60
        intSode1H = 10
        
        Select Case dblDH
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 174
                
                'strShingumizu = "SS-39"
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-38"
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 1 * intMaisu
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-38"
                
        End Select
        
        If dblDH <= 2529 Then
            dblCupShitaji = 35
            intCupShitajiH = 2 * intMaisu
        End If
        
'   *MS1(ミラー)*************************************************
    ElseIf strHinban Like "T?CME-####M*-*" Then
        '両側ミラー
        If strHinban Like "*-####MM*-*" Then
            
            dblTegake = 56.5
            intTegakeH = 1 * intMaisu
            
            dblShinAtsu = 30.2
            dblSan = dblDW - 320
            '20160115 K.Asayama Change 寸法間違い修正（笹原課長より）
            'dblGakuYoko1 = dblDW - 596.5
            dblGakuYoko1 = dblDW - 631.5
            '20160115 K.Asayama Change End
            intHashiraH = 2 * intMaisu
            dblSode1 = 56.5
            intSode1H = 2 * intMaisu
            
            dblShinAtsu_N = 14.8
            dblSan_N = 388
            dblGakuYoko1_N = 328
            intHashiraH_N = 2 * intMaisu
            
            Select Case dblDH
                Case 2530 To 2589
                    intSanH = 6 * intMaisu
                    intGakuYokoH1 = 2 * intMaisu
                    dblHashira = dblDH - 174
                    
                    intSanH_N = 6 * intMaisu
                    intGakuYokoH1_N = 7 * intMaisu
                    dblHashira_N = dblDH - 175
                    
                    'strShingumizu = "SS-40"
                    
                Case 1801 To 2529
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 2 * intMaisu
                    dblHashira = dblDH - 114
                    
                    intSanH_N = 4 * intMaisu
                    intGakuYokoH1_N = 6 * intMaisu
                    dblHashira_N = dblDH - 115
                    
                    'strShingumizu = "SS-38"
                    
                Case Is <= 1800
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 1 * intMaisu
                    dblHashira = dblDH - 114
                    
                    intSanH_N = 4 * intMaisu
                    intGakuYokoH1_N = 5 * intMaisu
                    dblHashira_N = dblDH - 115
                    
                    'strShingumizu = "SS-38"
                    
            End Select
            
            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If
        
        '片側ミラー
        Else
        
            dblTegake = 0: intTegakeH = 0
            
            Select Case dblDH
                Case 2530 To 2589
                    'strShingumizu = "SS-39/40"
                Case 1801 To 2529
                    'strShingumizu = "SS-37/38"
                Case Is <= 1800
                    'strShingumizu = "SS-37/38"
            End Select
            
        End If
        
'   *MP3*********************************************************
'20161108 K.Asayama Change 品番間違い修正
    'ElseIf strHinban Like "*F?CSA-####F*-*" Then
    ElseIf strHinban Like "P?CSA-####F*-*" Then
'20161108 K.Asayama Change END
    
        dblShinAtsu = 18
        dblSan = dblDW + 4
        dblGakuYoko1 = dblDW - 442
        intGakuYokoH1 = 2 * intMaisu
        dblGakuYoko2 = 218
        '20161116 K.Asayama Change 額横2本数漏れ追記
        intGakuYokoH2 = 1 * intMaisu
        '20161116 K.Asayama Change END
        intHashiraH = 6 * intMaisu
        intGakutateH1 = 1 * intMaisu
        
        dblTegakeShurui = 20
        dblTegake = 218
        intTegakeH = 3
        
        Select Case dblDH
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                dblGakutate1 = dblDH - 174
                dblHashira = dblDH - 174
                
                'strShingumizu = "SS-42"
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                dblGakutate1 = dblDH - 114
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-41"
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                dblGakutate1 = dblDH - 114
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-41"
                
        End Select
        
        '20161116 K.Asayama DEL 不要→削除
'        If dblDH <= 2529 Then
'            dblCupShitaji = 35
'            intCupShitajiH = 2 * intMaisu
'        End If
'   *CF1/EF1/ZF1*************************************************
    
    ElseIf strHinban Like "F?C??*-####F*-*" Then
        
        dblShinAtsu = 30.2
        dblSan = dblDW + 2
        'dblGakuYoko1 = dblDW - 248
        intHashiraH = 2 * intMaisu
        dblSode1 = 60
        

        
        '20160825 K.Asayama Change
'        If strHinban Like "*DH-####*" Then
'            intSode1H = 8
'        ElseIf strHinban Like "*DE-####*" Then
'            intSode1H = 10
'        ElseIf strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Then
'            intSode1H = 12
'        Else
'            intSode1H = 5
'        End If

'20161121 K.Asayama Change
'        If IsHikido(strHinban) Then
'            If strHinban Like "*DN-####*-*" Then
'                intSode1H = 2 * intMaisu
'            Else
'                intSode1H = 4 * intMaisu
'            End If
'        Else
'            intSode1H = 5 * intMaisu
'        End If
        
        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
'20161121 K.Asayama Change END
        
        '20170105 K.Asayama Change
'        If strHinban Like "*DN-####*" Then
        If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
        '20170105 K.Asayama Change END
        
            dblGakuYoko1 = dblDW - 188
        Else
            dblGakuYoko1 = dblDW - 248
        End If
        '20160825 K.Asayama Change END
        
        Select Case dblDH
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                dblGakuYokoLVL30 = dblDW - 58
                intGakuYokoLVL30 = 2 * intMaisu
                dblGakutate3 = 150
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                    intGakutateH3 = 4 * intMaisu
                Else
                    intGakutateH3 = 8 * intMaisu
                End If
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 174
                
                'strShingumizu = "SS-2"
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-1"
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 1 * intMaisu
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-1"
        End Select
        
        '20180205 K.Asayama Change
        If dblDH > 2589 Then
            dblCupShitaji = 60
            intCupShitajiH = 6 * intMaisu
        
        '20170105 K.Asayama Change 1701品番追加
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
        Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" _
        Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" _
        Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Then
        '20170105 K.Asayama Change END
        
            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If
        End If
        
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
            intSode1H = 3 * intMaisu
            dblSode2 = 110
            intSode2H = 2 * intMaisu
            dblGakuYoko1 = dblDW - 298
        End If
        '******************************************************
        
        '20180205 K.Asayama Change
        If dblDH <= 2589 Then
        '20151211 K.Asayama Change 1601仕様追加
            If IsHidden_Hinge(strHinban) Then
                dblGakutate1 = 210
                intGakutateH1 = 2
            End If
        '20151211 K.Asayama Change End
        End If
        
'   *CG2/EG2/ZG2*************************************************

    ElseIf strHinban Like "F?C??*-####C*-*" Then
    
        dblShinAtsu = 30.2
        dblSan = dblDW - 80
        dblGakuYoko1 = (dblDW / 2) - 210
        intHashiraH = 5 * intMaisu
        dblSode1 = 60
        
        '20161121 K.Asayama Change
        '20160825 K.Asayama Change
        'intSode1H = 5 * intMaisu
'        If IsHikido(strHinban) Then
'            If strHinban Like "*DN-####*-*" Then
'                intSode1H = 2 * intMaisu
'            Else
'                intSode1H = 4 * intMaisu
'            End If
'        Else
'            intSode1H = 5 * intMaisu
'        End If
        
        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END

        '20170105 K.Asayama Change
'        If strHinban Like "*DN-####*" Then
        If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
        '20170105 K.Asayama Change END
        
            dblGakuYoko2 = dblDW - 188
        End If
        '20160825 K.Asayama Change END
        
        Select Case dblDH
            
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intSanH = 4 * intMaisu
                
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                    intGakuYokoH1 = 2 * intMaisu
                    dblGakuYoko2 = (dblDW / 2) - 150
                    intGakuYokoH2 = 2 * intMaisu
                Else
                    intGakuYokoH1 = 4 * intMaisu
                End If

                dblHashira = dblDH - 114
                dblGakuYokoLVL30 = (dblDW / 2) - 115
                intGakuYokoLVL30 = 4 * intMaisu
                dblGakutate3 = 150
                intGakutateH3 = 8 * intMaisu
                
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                '20160825 K.Asayama Change
                'intGakuYokoH1 = 4 * intMaisu
                
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END

                    intGakuYokoH1 = 2 * intMaisu
                    dblGakuYoko2 = (dblDW / 2) - 150
                    intGakuYokoH2 = 2 * intMaisu
                Else
                    intGakuYokoH1 = 4 * intMaisu
                End If
                '20160825 K.Asayama Change END
                dblHashira = dblDH - 174
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-8"
                Else
                    'strShingumizu = "SS-6"
                End If
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                '20160825 K.Asayama Change
                'intGakuYokoH1 = 4 * intMaisu
                
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END

                    intGakuYokoH1 = 2 * intMaisu
                    dblGakuYoko2 = (dblDW / 2) - 150
                    intGakuYokoH2 = 2 * intMaisu
                Else
                    intGakuYokoH1 = 4 * intMaisu
                End If
                '20160825 K.Asayama Change END
                dblHashira = dblDH - 114
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-7"
                Else
                    'strShingumizu = "SS-5"
                End If
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                '20160825 K.Asayama Change
                'intGakuYokoH1 = 2 * intMaisu
                
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END

                    intGakuYokoH1 = 1 * intMaisu
                    dblGakuYoko2 = (dblDW / 2) - 150
                    intGakuYokoH2 = 1 * intMaisu
                Else
                    intGakuYokoH1 = 2 * intMaisu
                End If
                '20160825 K.Asayama Change END
                dblHashira = dblDH - 114
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-7"
                Else
                    'strShingumizu = "SS-5"
                End If
                
        End Select
        
        If IsKotobira(strHinban) Then
            dblSan = dblDW - 64
            intGakuYokoH1 = 4
            dblGakuYoko1 = (dblDW / 2) - 142
            dblSode1 = 0
            intSode1H = 0
            
            '20180205 K.Asayama ADD
            dblGakuYokoLVL30 = 0
            intGakuYokoLVL30 = 0
            dblGakutate3 = 0
            intGakutateH3 = 0
            
        '20180205 K.Asayama ADD
        ElseIf dblDH > 2589 Then
                dblCupShitaji = 60
                intCupShitajiH = 8 * intMaisu
                
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then

            dblDaboShitaji = 150
            intDaboShitajiH = 2 * intMaisu
            
            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If
        
        End If
        
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
        
            intSode1H = 3 * intMaisu
            dblSode2 = 110
            intSode2H = 2 * intMaisu
            intGakuYokoH1 = intGakuYokoH1 / 2
            dblGakuYoko2 = dblDW - 260
            intGakuYokoH2 = intGakuYokoH1
            
        End If
        '******************************************************
        
        '20151211 K.Asayama Change 1601仕様追加
        If IsHidden_Hinge(strHinban) Then
            If IsKotobira(strHinban) Then
                dblGakuYoko1 = (dblDW / 2) - 142.5
                intGakuYokoH1 = 2
                dblGakuYoko2 = (dblDW / 2) - 162.5
                intGakuYokoH2 = 2
                intGakutateH1 = 1
                Select Case dblDH
                '20180205 K.Asayama Change
                    Case Is > 2589
                        dblGakutate1 = dblDH - 114
                    Case 2530 To 2589
                        dblGakutate1 = dblDH - 174
                    Case Is <= 2529
                        dblGakutate1 = dblDH - 114
                End Select
            Else
            '20180205 K.Asayama Change
                If dblDH <= 2589 Then
                    dblGakutate1 = 210
                    intGakutateH1 = 2
                End If
            End If
        End If
        '20151211 K.Asayama Change End

'   *RG1*********************************************************

    ElseIf strHinban Like "R?C??*-####S*-*" Then
        
        dblShinAtsu = 30.2
        dblSan = dblDW - 63
'        dblGakuYoko1 = dblDW - 493.5
        intHashiraH = 5 * intMaisu
        dblSode1 = 90.5
        dblSode2 = 60

        '20160825 K.Asayama Change
'        intSode1H = 2 * intMaisu
'        intSode2H = 6 * intMaisu
        
        '20161121 K.Asayama Change
        If IsHikido(strHinban) Then
            intSode1H = 4 * intMaisu
'            If strHinban Like "*DN-####*-*" Then
'                intSode2H = 3 * intMaisu
'            Else
'                intSode2H = 5 * intMaisu
'            End If
        Else
            intSode1H = 2 * intMaisu
'            intSode2H = 6 * intMaisu
        End If
        
        intSode2H = intFncSode2Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        
        '20170105 K.Asayama Change
'        If strHinban Like "*DN-####*" Then
        If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
        '20170105 K.Asayama Change END

            dblGakuYoko1 = dblDW - 433.5
        Else
            dblGakuYoko1 = dblDW - 493.5
        End If
        '20160825 K.Asayama Change END
        
        
        Select Case dblDH
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intSanH = 4 * intMaisu

                intGakuYokoH1 = 2 * intMaisu

                dblHashira = dblDH - 114
                dblGakuYokoLVL30 = dblDW - 303.5
                intGakuYokoLVL30 = 2 * intMaisu
                dblGakutate3 = 150
                If IsHirakido(strHinban) Or IsOyatobira(strHinban) Then
                    If IsHidden_Hinge(strHinban) Then
                        intGakutateH3 = 9 * intMaisu
                    Else
                        intGakutateH3 = 8 * intMaisu
                    End If
                    
                ElseIf strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                    intGakutateH3 = 5 * intMaisu
                Else
                    intGakutateH3 = 9 * intMaisu
                End If
                
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 174
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-8"
                Else
                    'strShingumizu = "SS-4"
                End If
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-7"
                Else
                    'strShingumizu = "SS-3"
                End If
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 1 * intMaisu
                dblHashira = dblDH - 114
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-7"
                Else
                    'strShingumizu = "SS-3"
                End If
                
        End Select
        
        If IsKotobira(strHinban) Then
        
            intGakuYokoH1 = 4
            dblGakuYoko1 = (dblDW / 2) - 141.5
            dblSode1 = 0
            intSode1H = 0
            dblSode2 = 0
            intSode2H = 0
            
            '20180205 K.Asayama ADD
            dblGakuYokoLVL30 = 0
            intGakuYokoLVL30 = 0
            dblGakutate3 = 0
            intGakutateH3 = 0
            
        '20180205 K.Asayama ADD
        ElseIf dblDH > 2589 Then
            dblCupShitaji = 60
            If IsHirakido(strHinban) Or IsOyatobira(strHinban) Then
                intCupShitajiH = 4 * intMaisu
            Else
                intCupShitajiH = 5 * intMaisu
            End If
            
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then

            dblDaboShitaji = 150
            intDaboShitajiH = 2 * intMaisu
            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If
            '20180205 K.Asayama ADD
            If dblDH > 2589 Then
                dblCupShitaji = 60
                intCupShitajiH = 5 * intMaisu
            End If

        End If
        
        '20151211 K.Asayama Change 1601仕様追加
        If IsHidden_Hinge(strHinban) Then
            If IsKotobira(strHinban) Then
                dblGakuYoko1 = (dblDW / 2) - 141.5
                intGakuYokoH1 = 2
                dblGakuYoko2 = (dblDW / 2) - 161.5
                intGakuYokoH2 = 2
                intGakutateH1 = 1
                Select Case dblDH
                    '20180205 K.Asayama ADD
                    Case Is > 2589
                        dblGakutate1 = dblDH - 114
                    Case 2530 To 2589
                        dblGakutate1 = dblDH - 174
                    Case Is <= 2529
                        dblGakutate1 = dblDH - 114
                End Select
            '20180205 K.Asayama ADD
            ElseIf dblDH <= 2589 Then
                dblGakutate1 = 210
                intGakutateH1 = 2
            End If
        End If
        '20151211 K.Asayama Change End

'   *RG2*********************************************************

    ElseIf strHinban Like "R?C??*-####C*-*" Then
        
        dblShinAtsu = 30.2
        dblSan = dblDW - 79
        '20160819 K.Asayama 誤り訂正
        'dblGakuYoko1 = (dblDW / 2) - 209
        dblGakuYoko1 = (dblDW / 2) - 209.5
        '20160819 K.Asayama Change End
        
        intHashiraH = 5 * intMaisu
        dblSode1 = 60
        
        '20160825 K.Asayama Change
'        intSode1H = 5 * intMaisu
        
        '20161121 K.Asayama Change
'        If IsHikido(strHinban) Then
'            If strHinban Like "*DN-####*-*" Then
'                intSode1H = 2 * intMaisu
'            Else
'                intSode1H = 4 * intMaisu
'            End If
'        Else
'            intSode1H = 5 * intMaisu
'        End If

        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        
        '20170105 K.Asayama Change
'        If strHinban Like "*DN-####*" Then
        If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
        '20170105 K.Asayama Change END

            dblGakuYoko2 = dblDW - 149.5
        End If
        '20160825 K.Asayama Change END
        
        Select Case dblDH
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intSanH = 4 * intMaisu
                dblHashira = dblDH - 114
                dblGakuYokoLVL30 = (dblDW / 2) - 114.5
                intGakuYokoLVL30 = 4 * intMaisu
                dblGakutate3 = 150
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                    intGakuYokoH1 = 2 * intMaisu
                    intGakutateH3 = 4 * intMaisu
                Else
                    intGakuYokoH1 = 4 * intMaisu
                    intGakutateH3 = 8 * intMaisu
                End If
                
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                '20160825 K.Asayama Change
                'intGakuYokoH1 = 4 * intMaisu
                
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END

                    intGakuYokoH1 = 2 * intMaisu
                    intGakuYokoH2 = 2 * intMaisu
                Else
                    intGakuYokoH1 = 4 * intMaisu
                End If
                '20160825 K.Asayama Change END
                dblHashira = dblDH - 174
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-8"
                Else
                    'strShingumizu = "SS-6"
                End If
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                '20160825 K.Asayama Change
                'intGakuYokoH1 = 4 * intMaisu
            
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END
                
                    intGakuYokoH1 = 2 * intMaisu
                    intGakuYokoH2 = 2 * intMaisu
                Else
                    intGakuYokoH1 = 4 * intMaisu
                End If
                '20160825 K.Asayama Change END
                dblHashira = dblDH - 114
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-7"
                Else
                    'strShingumizu = "SS-5"
                End If
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                '20160825 K.Asayama Change
                'intGakuYokoH1 = 2 * intMaisu
                
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END
                
                    intGakuYokoH1 = 1 * intMaisu
                    intGakuYokoH2 = 1 * intMaisu
                Else
                    intGakuYokoH1 = 2 * intMaisu
                End If
                '20160825 K.Asayama Change END
                dblHashira = dblDH - 114
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-7"
                Else
                    'strShingumizu = "SS-5"
                End If
                
        End Select
        
        If IsKotobira(strHinban) Then
            dblSan = dblDW - 63
            intGakuYokoH1 = 4
            dblGakuYoko1 = (dblDW / 2) - 141.5
            dblSode1 = 0
            intSode1H = 0
            
            '20180205 K.Asayama ADD
            dblGakuYokoLVL30 = 0
            intGakuYokoLVL30 = 0
            dblGakutate3 = 0
            intGakutateH3 = 0
            
        '20180205 K.Asayama Change
        ElseIf dblDH > 2589 Then
            dblCupShitaji = 60
            intCupShitajiH = 8 * intMaisu
            
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then

            dblDaboShitaji = 150
            intDaboShitajiH = 2 * intMaisu
            
            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If

        End If
        
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
        
            intSode1H = 3 * intMaisu
            dblSode2 = 110
            intSode2H = 2 * intMaisu
            intGakuYokoH1 = intGakuYokoH1 / 2
            dblGakuYoko2 = dblDW - 259.5
            intGakuYokoH2 = intGakuYokoH1
            
        End If
        '******************************************************
        
        '20151211 K.Asayama Change 1601仕様追加
        If IsHidden_Hinge(strHinban) Then
            If IsKotobira(strHinban) Then
                dblGakuYoko1 = (dblDW / 2) - 141.5
                intGakuYokoH1 = 2
                dblGakuYoko2 = (dblDW / 2) - 161.5
                intGakuYokoH2 = 2
                intGakutateH1 = 1
                Select Case dblDH
                    '20180205 K.Asayama ADD
                    Case Is > 2589
                        dblGakutate1 = dblDH - 114
                    Case 2530 To 2589
                        dblGakutate1 = dblDH - 174
                    Case Is <= 2529
                        dblGakutate1 = dblDH - 114
                End Select
            Else
                dblGakutate1 = 210
                intGakutateH1 = 2
            End If
        End If
        '20151211 K.Asayama Change End

'   *RF1/JF1*****************************************************

    ElseIf strHinban Like "R?C??*-####F*-*" Then
        
        dblShinAtsu = 30.2
        dblSan = dblDW + 4
        'dblGakuYoko1 = dblDW - 246
        intHashiraH = 2 * intMaisu
        dblSode1 = 60
         
        '20160825 K.Asayama Change
'        If strHinban Like "*DH-####*" Then
'            intSode1H = 8
'        ElseIf strHinban Like "*DE-####*" Then
'            intSode1H = 10
'        ElseIf strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Then
'            intSode1H = 12
'        Else
'            intSode1H = 5
'        End If
               
        '20161121 K.Asayama Change
'        If IsHikido(strHinban) Then
'            If strHinban Like "*DN-####*-*" Then
'                intSode1H = 2 * intMaisu
'            Else
'                intSode1H = 4 * intMaisu
'            End If
'        Else
'            intSode1H = 5 * intMaisu
'        End If

        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
    
        '20170105 K.Asayama Change
'        If strHinban Like "*DN-####*" Then
        If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
        '20170105 K.Asayama Change END

            dblGakuYoko1 = dblDW - 186
        Else
            dblGakuYoko1 = dblDW - 246
        End If
        '20160825 K.Asayama Change END
        
        Select Case dblDH
        
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                dblGakuYokoLVL30 = dblDW - 56
                intGakuYokoLVL30 = 2 * intMaisu
                dblGakutate3 = 150
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                    intGakutateH3 = 4 * intMaisu
                Else
                    intGakutateH3 = 8 * intMaisu
                End If
                
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 174
                
                'strShingumizu = "SS-2"
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-1"
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 1 * intMaisu
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-1"
                
        End Select
        
        '20180205 K.Asayama Change
        If dblDH > 2589 Then
            dblCupShitaji = 60
            intCupShitajiH = 6 * intMaisu
            
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
            
            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If
            
        '20170105 K.Asayama Change
'        ElseIf strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Then
        ElseIf strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Then
        '20170105 K.Asayama Change END
        
            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 4
            End If
        End If
        
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
            intSode1H = 3 * intMaisu
            dblSode2 = 110
            intSode2H = 2 * intMaisu
            dblGakuYoko1 = dblDW - 296
        End If
        '******************************************************
        
        '20160822 K.Asayama ADD 1601仕様漏れ追加
        '20180205 K.Asayama Change
        If dblDH <= 2589 Then
            If IsHidden_Hinge(strHinban) Then
                dblGakutate1 = 210
                intGakutateH1 = 2
            End If
        End If
        '20160822 K.Asayama ADD END
        
'   *CF6/EF6/ZF6*************************************************

    ElseIf strHinban Like "F?D??*-####F*-*" Then
        
        dblShinAtsu = 30.2
        dblSan = dblDW + 2
'        dblGakuYoko1 = dblDW - 248
        intHashiraH = 2 * intMaisu
        dblSode1 = 60
        
        '20160825 K.Asayama Change
'        If strHinban Like "*DH-####*" Then
'            intSode1H = 8
'        ElseIf strHinban Like "*DE-####*" Then
'            intSode1H = 10
'        ElseIf strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Then
'            intSode1H = 12
'        Else
'            intSode1H = 5
'        End If
        
        '20161121 K.Asayama Change
'        If IsHikido(strHinban) Then
'            If strHinban Like "*DN-####*-*" Then
'                intSode1H = 2 * intMaisu
'            Else
'                intSode1H = 4 * intMaisu
'            End If
'        Else
'            intSode1H = 5 * intMaisu
'        End If

        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        
        '20170105 K.Asayama Change
'        If strHinban Like "*DN-####*" Then
        If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
        '20170105 K.Asayama Change END

            dblGakuYoko1 = dblDW - 188
        Else
            dblGakuYoko1 = dblDW - 248
        End If
        '20160825 K.Asayama Change END
        
        Select Case dblDH
        
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intSanH = 4 * intMaisu
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                    intGakuYokoH1 = 7 * intMaisu
                Else
                    intGakuYokoH1 = 6 * intMaisu
                End If
                dblHashira = dblDH - 114
                dblGakuYokoLVL30 = dblDW - 58
                intGakuYokoLVL30 = 2 * intMaisu
                dblGakutate3 = 150
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                    intGakutateH3 = 4 * intMaisu
                Else
                    intGakutateH3 = 8 * intMaisu
                End If
                
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                '20160825 K.Asayama Change
                'intGakuYokoH1 = 6 * intMaisu
                
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END
                
                    intGakuYokoH1 = 7 * intMaisu
                Else
                    intGakuYokoH1 = 6 * intMaisu
                End If
                '20160825 K.Asayama Change END
                
                dblHashira = dblDH - 174
                
                'strShingumizu = "SS-14"
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                '20160825 K.Asayama Change
                'intGakuYokoH1 = 6 * intMaisu
                
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END
                
                    intGakuYokoH1 = 7 * intMaisu
                Else
                    intGakuYokoH1 = 6 * intMaisu
                End If
                '20160825 K.Asayama Change END
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-13"
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                '20160825 K.Asayama Change
                'intGakuYokoH1 = 4 * intMaisu
                
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END
                
                    intGakuYokoH1 = 5 * intMaisu
                Else
                    intGakuYokoH1 = 4 * intMaisu
                End If
                '20160825 K.Asayama Change END
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-13"
                
        End Select
        
        '20180205 K.Asayama Change
        If dblDH > 2589 Then
            dblCupShitaji = 60
            intCupShitajiH = 6 * intMaisu
            
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
            
            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If
        End If
        
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
            intSode1H = 3 * intMaisu
            dblSode2 = 110
            intSode2H = 2 * intMaisu
            dblGakuYoko1 = dblDW - 298
        End If
        '******************************************************
        
        '20151211 K.Asayama Change 1601仕様追加
        '20180205 K.Asayama Change
        'If IsHidden_Hinge(strHinban) Then
        If IsHidden_Hinge(strHinban) And dblDH <= 2589 Then
            dblGakutate1 = 210
            intGakutateH1 = 2
        End If
        '20151211 K.Asayama Change End

'   *CG8/EG8/ZG8*************************************************
    
    ElseIf strHinban Like "F?C??*-####D*-*" Then
        
        dblShinAtsu = 30.2
        dblSan = dblDW - 290
        dblGakuYoko1 = (dblDW / 2) - 315
        intHashiraH = 5 * intMaisu
        dblSode1 = 60
        '20160825 K.Asayama Change
'        intSode1H = 5
        
        '20161121 K.Asayama Change
'        If IsHikido(strHinban) Then
'            If strHinban Like "*DN-####*-*" Then
'                intSode1H = 2 * intMaisu
'            Else
'                intSode1H = 4 * intMaisu
'            End If
'        Else
'            intSode1H = 5 * intMaisu
'        End If

        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        
        '20160825 K.Asayama Change END
        
        intSanH = 4 * intMaisu
        dblHashira = dblDH - 114
        
        Select Case dblDH
 
            Case 1801 To 2529
                '20160825 K.Asayama Change
                'intGakuYokoH1 = 4 * intMaisu
                
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END

                    intGakuYokoH1 = 2 * intMaisu
                    intGakuYokoH2 = 2 * intMaisu
                    dblGakuYoko2 = (dblDW / 2) - 255
                Else
                    intGakuYokoH1 = 4 * intMaisu
                End If
                '20160825 K.Asayama Change END
              
            
            Case Is <= 1800
                '20160825 K.Asayama Change
                'intGakuYokoH1 = 2 * intMaisu
                
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END

                    intGakuYokoH1 = 1 * intMaisu
                    intGakuYokoH2 = 1 * intMaisu
                    dblGakuYoko2 = (dblDW / 2) - 255
                Else
                    intGakuYokoH1 = 2 * intMaisu
                End If
                '20160825 K.Asayama Change END
                
        End Select
        
        If IsKotobira(strHinban) Then
            'strShingumizu = "SS-7"
        Else
            'strShingumizu = "SS-5"
        End If
                
        If IsKotobira(strHinban) Then
            dblSan = dblDW - 64
            intGakuYokoH1 = 4
            dblGakuYoko1 = (dblDW / 2) - 142
            dblSode1 = 0
            intSode1H = 0
            
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then

            dblDaboShitaji = 150
            intDaboShitajiH = 2 * intMaisu
            dblCupShitaji = 35
            intCupShitajiH = 2 * intMaisu
 
        End If
    
        '20151211 K.Asayama Change 1601仕様追加
        If IsHidden_Hinge(strHinban) Then
            If IsKotobira(strHinban) Then
                intGakuYokoH1 = 2
                dblGakuYoko2 = (dblDW / 2) - 162
                intGakuYokoH2 = 2
                dblGakutate1 = dblDH - 114
                intGakutateH1 = 1
            Else
                dblGakutate1 = 210
                intGakutateH1 = 2
            End If
        End If
        '20151211 K.Asayama Change End

'   *CG1/EG1/ZG1*************************************************

    ElseIf strHinban Like "F?C??*-####S*-*" Then
        
        dblShinAtsu = 30.2
        dblSan = dblDW - 64
        
        intHashiraH = 5 * intMaisu
        dblSode1 = 90
        dblSode2 = 60
        
        '20160825 K.Asayama Change
'        intSode1H = 2 * intMaisu
'        intSode2H = 6 * intMaisu

        '20161121 K.Asayama Change
        If IsHikido(strHinban) Then
            intSode1H = 4 * intMaisu
'            If strHinban Like "*DN-####*-*" Then
'                intSode2H = 3 * intMaisu
'            Else
'                intSode2H = 5 * intMaisu
'            End If
        Else
            intSode1H = 2 * intMaisu
'            intSode2H = 6 * intMaisu
        End If
        
        intSode2H = intFncSode2Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        
        '20170105 K.Asayama Change
'        If strHinban Like "*DN-####*" Then
        If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
        '20170105 K.Asayama Change END

            dblGakuYoko1 = dblDW - 434
        Else
            dblGakuYoko1 = dblDW - 494
        End If
        '20160825 K.Asayama Change END
        
        
        Select Case dblDH
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu

                dblHashira = dblDH - 114
                dblGakuYokoLVL30 = dblDW - 303.5
                intGakuYokoLVL30 = 2 * intMaisu
                dblGakutate3 = 150
                
                If IsHirakido(strHinban) Or IsOyatobira(strHinban) Then
                    If IsHidden_Hinge(strHinban) Then
                        intGakutateH3 = 9 * intMaisu
                    Else
                        intGakutateH3 = 8 * intMaisu
                    End If
                    
                ElseIf strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                    intGakutateH3 = 5 * intMaisu
                Else
                    intGakutateH3 = 9 * intMaisu
                End If
                
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 174
                'dblGakuYoko1 = dblDW - 494
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-8"
                Else
                    'strShingumizu = "SS-4"
                End If
        
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                'dblGakuYoko1 = dblDW - 494
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-7"
                Else
                    'strShingumizu = "SS-3"
                End If
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 1 * intMaisu
                dblHashira = dblDH - 114
                'dblGakuYoko1 = dblDW - 494
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-7"
                Else
                    'strShingumizu = "SS-3"
                End If
                
        End Select
        
        If IsKotobira(strHinban) Then
            intGakuYokoH1 = 4
            dblGakuYoko1 = (dblDW / 2) - 142
            dblSode1 = 0
            intSode1H = 0
            dblSode2 = 0
            intSode2H = 0
            
            '20180205 K.Asayama ADD
            dblGakuYokoLVL30 = 0
            intGakuYokoLVL30 = 0
            dblGakutate3 = 0
            intGakutateH3 = 0
            
        '20180205 K.Asayama Change
        ElseIf dblDH > 2589 Then
            dblCupShitaji = 60
            If IsHirakido(strHinban) Or IsOyatobira(strHinban) Then
                intCupShitajiH = 4 * intMaisu
            Else
                intCupShitajiH = 5 * intMaisu
            End If
        '20180111 K.Asayama Change 品番漏れ修正
'        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DE-####*" _
'            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Then
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
            
            dblDaboShitaji = 150
            intDaboShitajiH = 2 * intMaisu
                
            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If

        End If
    
        '20151211 K.Asayama Change 1601仕様追加
        If IsHidden_Hinge(strHinban) Then
            If IsKotobira(strHinban) Then
                '20160819 K.Asayama 誤り訂正
                'dblGakuYoko1 = (dblDW / 2) - 142.5
                dblGakuYoko1 = (dblDW / 2) - 142
                intGakuYokoH1 = 2
                'dblGakuYoko2 = (dblDW / 2) - 162.5
                dblGakuYoko2 = (dblDW / 2) - 162
                '20160819 K.Asayama Change End
                
                intGakuYokoH2 = 2
                intGakutateH1 = 1
                Select Case dblDH
                    '20180205 K.Asayama Change
                    Case Is > 2589
                        dblGakutate1 = dblDH - 114
                    Case 2530 To 2589
                        dblGakutate1 = dblDH - 174
                    Case Is <= 2529
                        dblGakutate1 = dblDH - 114
                End Select
            Else
                '20180205 K.Asayama Change
                If dblDH <= 2589 Then
                    dblGakutate1 = 210
                    intGakutateH1 = 2
                End If
            End If
        End If
        '20151211 K.Asayama Change End

'   *KF7*********************************************************
    '20151211 K.Asayama Change KF1,KF7隠し丁番統一
    'ElseIf strHinban Like "*S?CD?*-####F*-*" Then
    ElseIf strHinban Like "S?C??*-####F*-*" Then
    '20151211 K.Asayama Change End
    
        dblShinAtsu = 30.2
        dblSan = dblDW + 2
        'dblGakuYoko1 = dblDW - 248
        intHashiraH = 2 * intMaisu
        dblSode1 = 60
        
        '20160825 K.Asayama Change
'        If strHinban Like "*DH-####*" Then
'            intSode1H = 8
'        ElseIf strHinban Like "*DE-####*" Then
'            intSode1H = 10
'        ElseIf strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Then
'            intSode1H = 12
'        Else
'            intSode1H = 5
'        End If
        
        '20161121 K.Asayama Change
'        If IsHikido(strHinban) Then
'            If strHinban Like "*DN-####*-*" Then
'                intSode1H = 2 * intMaisu
'            Else
'                intSode1H = 4 * intMaisu
'            End If
'        Else
'            intSode1H = 5 * intMaisu
'        End If
        
        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        
        '20170105 K.Asayama Change
'        If strHinban Like "*DN-####*" Then
        If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
        '20170105 K.Asayama Change END

            dblGakuYoko1 = dblDW - 188
        Else
            dblGakuYoko1 = dblDW - 248
        End If
        '20160825 K.Asayama Change END
        
        Select Case dblDH
            
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                dblGakutate1 = (dblDH - 274) / 3
                intGakutateH1 = 3 * intMaisu
                
                dblGakuYokoLVL30 = dblDW - 58
                intGakuYokoLVL30 = 2 * intMaisu
                dblGakutate3 = 150
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                    intGakutateH3 = 4 * intMaisu
                Else
                    intGakutateH3 = 8 * intMaisu
                End If
                
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 174
                dblGakutate1 = (dblDH - 214) / 3
                intGakutateH1 = 3 * intMaisu
                
                'strShingumizu = "SS-31"
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                dblGakutate1 = (dblDH - 154) / 3
                intGakutateH1 = 3 * intMaisu
                
                'strShingumizu = "SS-30"
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 1 * intMaisu
                dblHashira = dblDH - 114
                dblGakutate1 = (dblDH - 134) / 2
                intGakutateH1 = 2 * intMaisu
                
                'strShingumizu = "SS-30"
                
        End Select
        
        '20180205 K.Asayama Change
        If dblDH > 2589 Then
            dblCupShitaji = 60
            intCupShitajiH = 6 * intMaisu
            
        '20151211 K.Asayama Change
        
        '20170105 K.Asayama Change
'        If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
'            Or strHinban Like "*DH-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" _
'            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
            Or strHinban Like "*DH-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" _
            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" _
            Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Then
        '20170105 K.Asayama Change
            
            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If
        End If
        '20151211 K.Asayama Change End
        
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
            intSode1H = 3 * intMaisu
            dblSode2 = 110
            intSode2H = 2 * intMaisu
            dblGakuYoko1 = dblDW - 298
        End If
        '******************************************************
        
        '20151211 K.Asayama Change 1601仕様追加
        '20180205 K.Asayama Change
        'If IsHidden_Hinge(strHinban) Then
        If IsHidden_Hinge(strHinban) And dblDH <= 2589 Then
            '20160823 K.Asayama Change 誤り訂正
            'dblGakutate1 = 210
            'intGakutateH1 = 2
            dblGakutate2 = 210
            intGakutateH2 = 2
            '20160823 K.Asayama Change End
        End If
        '20151211 K.Asayama Change End

'20151211 K.Asayama Del KF1,KF7隠し丁番統一
''   *KF7隠し丁番*************************************************
'
'    ElseIf strHinban Like "*S?CK?*-####Z*-*" Then
'
'        dblShinAtsu = 30.2
'        dblSan = dblDW + 2
'        dblGakuYoko1 = dblDW - 268
'        intHashiraH = 2
'        dblSode1 = 60
'        intSode1H = 5
'
'        Select Case dblDH
'            Case 2530 To 2589
'                intSanH = 6 * intMaisu
'                intGakuYokoH1 = 2 * intMaisu
'                dblHashira = dblDH - 174
'                dblGakutate1 = (dblDH - 214) / 3
'                intGakutateH1 = 3
'                dblGakutate2 = dblDH - 174
'                intGakutateH2 = 1
'
'                'strShingumizu = "SS-31"
'
'            Case 1801 To 2529
'                intSanH = 4 * intMaisu
'                intGakuYokoH1 = 2 * intMaisu
'                dblHashira = dblDH - 114
'                dblGakutate1 = (dblDH - 154) / 3
'                intGakutateH1 = 3
'                dblGakutate2 = dblDH - 114
'                intGakutateH2 = 1
'
'                'strShingumizu = "SS-30"
'
'            Case Is <= 1800
'                intSanH = 4 * intMaisu
'                intGakuYokoH1 = 1 * intMaisu
'                dblHashira = dblDH - 114
'                dblGakutate1 = (dblDH - 134) / 2
'                intGakutateH1 = 2
'                dblGakutate2 = dblDH - 114
'                intGakutateH2 = 1
'
'                'strShingumizu = "SS-30"
'
'        End Select
'
'        'AUハンドル例外処理************************************
'        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
'            intSode1H = 3 * intMaisu
'            dblSode2 = 110
'            intSode2H = 2 * intMaisu
'            dblGakuYoko1 = dblDW - 318
'        End If
'        '******************************************************
'20151211 K.Asayama Del End

'   *XG1*********************************************************
    '20160909 K.Asayama Change PG1型の130/70ステルス(X)と品番が被るため条件修正
    'ElseIf strHinban Like "*X?C??*-####S*-*" Then
    ElseIf strHinban Like "X?C??*-####S*-*" Or strHinban Like "特 X?C??*-####S*-*" Then
    '20160909 K.Asayama Change END
    
        dblShinAtsu = 30.2
        dblSan = dblDW - 64
        dblGakuYoko1 = dblDW - 494
        intHashiraH = 5
        dblSode1 = 90
        'intSode1H = 2
        dblSode2 = 60
        'intSode2H = 6
        
        '20160825 K.Asayama Change
        
        '20161121 K.Asayama Change
        If IsHikido(strHinban) Then
            intSode1H = 4
'            intSode2H = 5
        Else
            intSode1H = 2
'            intSode2H = 6
        End If
        '20160825 K.Asayama Change END
        
        intSode2H = intFncSode2Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        
        Select Case dblDH
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 174
                
                'strShingumizu = "SS-4"
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-3"
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 1 * intMaisu
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-3"
                
        End Select
        
        If strHinban Like "*DC-####*" Or strHinban Like "*KC-####*" Then
        
            dblDaboShitaji = 150
            intDaboShitajiH = 2
            
            If dblDH < 2530 Then
                dblCupShitaji = 35
                intCupShitajiH = 2
            End If
        End If

'   *XG2*********************************************************
    
    '20160909 K.Asayama Change PG1型の130/70ステルス(X)と品番が被るため条件修正
    'ElseIf strHinban Like "*X?C??*-####C*-*" Then
    ElseIf strHinban Like "X?C??*-####C*-*" Or strHinban Like "特 X?C??*-####S*-*" Then
    '20160909 K.Asayama Change END
    
    
    
        dblShinAtsu = 30.2
        dblSan = dblDW - 80
        dblGakuYoko1 = (dblDW / 2) - 80
        intHashiraH = 5
        dblSode1 = 60
        intSode1H = 5
        
        '20160825 K.Asayama Change
        
        '20161121 K.Asayama Change
        
'        If IsHikido(strHinban) Then
'            intSode1H = 4
'        Else
'            intSode1H = 5
'        End If

        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        
        '20160825 K.Asayama Change END
        
        Select Case dblDH
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                intGakuYokoH1 = 4 * intMaisu
                dblHashira = dblDH - 174
                
                'strShingumizu = "SS-6"
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 4 * intMaisu
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-5"
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-5"
                
        End Select
        
        If strHinban Like "*DC-####*" Or strHinban Like "*KC-####*" Then
        
            dblDaboShitaji = 150
            intDaboShitajiH = 2
            
            If dblDH < 2530 Then
                dblCupShitaji = 35
                intCupShitajiH = 2
            End If
        End If
                  
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
        
            intSode1H = 3 * intMaisu
            dblSode2 = 110
            intSode2H = 2 * intMaisu
            intGakuYokoH1 = intGakuYokoH1 / 2
            dblGakuYoko2 = dblDW - 260
            intGakuYokoH2 = intGakuYokoH1
            
        End If
        '******************************************************
        
'20151211 K.Asayama Del KF1,KF7隠し丁番統一
''   *KF1（隠し丁番）*********************************************
'
'    ElseIf strHinban Like "*S?CKA-####Z*-*" Or strHinban Like "*S?CKAS-####Z*-*" Then
'
'        dblShinAtsu = 30.2
'        dblSan = dblDW + 6
'        dblGakuYoko1 = dblDW - 264
'        intHashiraH = 2
'        dblSode1 = 60
'        intSode1H = 5
'
'        Select Case dblDH
'            Case 2530 To 2589
'                intSanH = 6 * intMaisu
'                intGakuYokoH1 = 2 * intMaisu
'                dblHashira = dblDH - 174
'                dblGakutate1 = (dblDH - 214) / 3
'                intGakutateH1 = 3
'                dblGakutate2 = dblDH - 174
'                intGakutateH2 = 1
'
'                'strShingumizu = "SS-31"
'
'            Case 1801 To 2529
'                intSanH = 4 * intMaisu
'                intGakuYokoH1 = 2 * intMaisu
'                dblHashira = dblDH - 114
'                dblGakutate1 = (dblDH - 154) / 3
'                intGakutateH1 = 3
'                dblGakutate2 = dblDH - 114
'                intGakutateH2 = 1
'
'                'strShingumizu = "SS-30"
'
'            Case Is <= 1800
'                intSanH = 4 * intMaisu
'                intGakuYokoH1 = 1 * intMaisu
'                dblHashira = dblDH - 114
'                dblGakutate1 = (dblDH - 134) / 2
'                intGakutateH1 = 2
'                dblGakutate2 = dblDH - 114
'                intGakutateH2 = 1
'
'                'strShingumizu = "SS-30"
'
'        End Select
'
'        'AUハンドル例外処理************************************
'        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
'            intSode1H = 3 * intMaisu
'            dblSode2 = 110
'            intSode2H = 2 * intMaisu
'            dblGakuYoko1 = dblDW - 314
'        End If
'        '******************************************************
'20151211 K.Asayama Del End
        
'   *KF1*********************************************************

    ElseIf strHinban Like "S?C??*-####Z*-*" Then

        dblShinAtsu = 30.2
        dblSan = dblDW + 6
        'dblGakuYoko1 = dblDW - 244
        intHashiraH = 2 * intMaisu
        dblSode1 = 60
        
        '20160825 K.Asayama Change
'        If strHinban Like "*DH-####*" Then
'            intSode1H = 8
'        ElseIf strHinban Like "*DE-####*" Then
'            intSode1H = 10
'        ElseIf strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Then
'            intSode1H = 12
'        Else
'            intSode1H = 5
'        End If
        
        
        '20161121 K.Asayama Change
        
'        If IsHikido(strHinban) Then
'            If strHinban Like "*DN-####*-*" Then
'                intSode1H = 2 * intMaisu
'            Else
'                intSode1H = 4 * intMaisu
'            End If
'        Else
'            intSode1H = 5 * intMaisu
'        End If
        
        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        
        '20170105 K.Asayama Change
'        If strHinban Like "*DN-####*" Then
        If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
        '20170105 K.Asayama Change END

            dblGakuYoko1 = dblDW - 184
        Else
            dblGakuYoko1 = dblDW - 244
        End If
        '20160825 K.Asayama Change END
        
        Select Case dblDH
        
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                dblGakutate1 = (dblDH - 274) / 3
                intGakutateH1 = 3 * intMaisu
                
                dblGakuYokoLVL30 = dblDW - 54
                intGakuYokoLVL30 = 2 * intMaisu
                dblGakutate3 = 150
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                    intGakutateH3 = 4 * intMaisu
                Else
                    intGakutateH3 = 8 * intMaisu
                End If
                
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 174
                dblGakutate1 = (dblDH - 214) / 3
                intGakutateH1 = 3 * intMaisu

                'strShingumizu = "SS-31"

            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                dblGakutate1 = (dblDH - 154) / 3
                intGakutateH1 = 3 * intMaisu

                'strShingumizu = "SS-30"

            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 1 * intMaisu
                dblHashira = dblDH - 114
                dblGakutate1 = (dblDH - 134) / 2
                intGakutateH1 = 2 * intMaisu

                'strShingumizu = "SS-30"

        End Select
        
        '20180205 K.Asayama Change
        If dblDH > 2589 Then
            dblCupShitaji = 60
            intCupShitajiH = 6 * intMaisu
            
        '20170105 K.Asayama Change
'        If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
'        Or strHinban Like "*DM-####*" Or strHinban Like "*DL-####*" Or strHinban Like "*DN-####*" _
'        Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" _
'        Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
        Or strHinban Like "*DM-####*" Or strHinban Like "*DL-####*" Or strHinban Like "*DN-####*" _
        Or strHinban Like "*VM-####*" Or strHinban Like "*VL-####*" Or strHinban Like "*VN-####*" _
        Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" _
        Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" _
        Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
        '20170105 K.Asayama Change END
            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If
        End If

        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
            intSode1H = 3 * intMaisu
            dblSode2 = 110
            intSode2H = 2 * intMaisu
            dblGakuYoko1 = dblDW - 294
        End If
        '******************************************************
        
        '20151211 K.Asayama Change 1601仕様追加
        '20180205 K.Asayama Change
        'If IsHidden_Hinge(strHinban) Then
        If IsHidden_Hinge(strHinban) And dblDH <= 2589 Then
            '20160823 K.Asayama Change 誤り訂正
            'dblGakutate1 = 210
            'intGakutateH1 = 2
            dblGakutate2 = 210
            intGakutateH2 = 2
            '20160823 K.Asayama Change End
        End If
        
        '20151211 K.Asayama Change End
        
'   *TF1*********************************************************

    ElseIf strHinban Like "T?C??*-####F*-*" Then
        
        dblShinAtsu = 30.2
        dblSan = dblDW + 5
'        dblGakuYoko1 = dblDW - 245
        intHashiraH = 2 * intMaisu
        dblSode1 = 60
        
        '20160825 K.Asayama Change
        
'        If strHinban Like "*DH-####*" Then
'            intSode1H = 8
'        ElseIf strHinban Like "*DE-####*" Then
'            intSode1H = 10
'        ElseIf strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Then
'            intSode1H = 12
'        Else
'            intSode1H = 5
'        End If
        
        
        '20161121 K.Asayama Change
'        If IsHikido(strHinban) Then
'            If strHinban Like "*DN-####*-*" Then
'                intSode1H = 2 * intMaisu
'            Else
'                intSode1H = 4 * intMaisu
'            End If
'        Else
'            intSode1H = 5 * intMaisu
'        End If

        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        
        '20170105 K.Asayama Change
'        If strHinban Like "*DN-####*" Then
        If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
        '20170105 K.Asayama Change END

            dblGakuYoko1 = dblDW - 185
        Else
            dblGakuYoko1 = dblDW - 245
        End If
        '20160825 K.Asayama Change END
        
        Select Case dblDH
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                dblGakuYokoLVL30 = dblDW - 55
                intGakuYokoLVL30 = 2 * intMaisu
                dblGakutate3 = 150
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                    intGakutateH3 = 4 * intMaisu
                Else
                    intGakutateH3 = 8 * intMaisu
                End If
                
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 174
                
                'strShingumizu = "SS-2"
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-1"
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 1 * intMaisu
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-1"
        End Select
        
        '20180205 K.Asayama ADD
        If dblDH > 2589.5 Then
            dblCupShitaji = 60
            intCupShitajiH = 6 * intMaisu
            
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
            
            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If
            
        '20170105 K.Asayama Change
'        ElseIf strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Then
        ElseIf strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Then
        '20170105 K.Asayama Change END
        
            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 6
            End If
        End If
        
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
            intSode1H = 3 * intMaisu
            dblSode2 = 110
            intSode2H = 2 * intMaisu
            dblGakuYoko1 = dblDW - 295
        End If
        '******************************************************
        
        '20151211 K.Asayama Change 1601仕様追加
        If IsHidden_Hinge(strHinban) Then
            '20180205 K.Asayama Chnage
            If dblDH <= 2589 Then
                dblGakutate1 = 210
                intGakutateH1 = 2
            End If
        End If
        '20151211 K.Asayama Change End
'   *TG1*********************************************************

    ElseIf strHinban Like "T?C??*-####S*-*" Then
        
        dblShinAtsu = 30.2
        dblSan = dblDW - 61
        
        intHashiraH = 5 * intMaisu
        dblSode1 = 91.5
'        intSode1H = 2 * intMaisu
        dblSode2 = 60
'        intSode2H = 6 * intMaisu
'        dblGakuYoko1 = dblDW - 492.5
        
        '20160825 K.Asayama Change
        '20161121 K.Asayama Change
        If IsHikido(strHinban) Then
            '20170105 K.Asayama Change
    '        If strHinban Like "*DN-####*" Then
            If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
            '20170105 K.Asayama Change END

                intSode1H = 4 * intMaisu
                'intSode2H = 3 * intMaisu
            Else
                intSode1H = 4 * intMaisu
                'intSode2H = 5 * intMaisu
            End If
        Else
            intSode1H = 2 * intMaisu
            'intSode2H = 6 * intMaisu
        End If

        intSode2H = intFncSode2Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        
        '20170105 K.Asayama Change
'        If strHinban Like "*DN-####*" Then
        If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
        '20170105 K.Asayama Change END

            dblGakuYoko1 = dblDW - 492.5
        Else
            dblGakuYoko1 = dblDW - 432.5
        End If
        '20160825 K.Asayama Change END
        
        Select Case dblDH
        
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                dblGakuYokoLVL30 = dblDW - 302
                intGakuYokoLVL30 = 2 * intMaisu
                dblGakutate3 = 150
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                    intGakutateH3 = 5 * intMaisu
                ElseIf IsHirakido(strHinban) Or IsOyatobira(strHinban) Then
                    intGakutateH3 = 8 * intMaisu
                Else
                    intGakutateH3 = 9 * intMaisu
                End If
                
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 174
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-8"
                Else
                    'strShingumizu = "SS-4"
                End If
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-7"
                Else
                    'strShingumizu = "SS-3"
                End If
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 1 * intMaisu
                dblHashira = dblDH - 114
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-7"
                Else
                    'strShingumizu = "SS-3"
                End If
                
        End Select
        
        If IsKotobira(strHinban) Then
            
            dblSan = dblDW - 61
            intGakuYokoH1 = 4
            dblGakuYoko1 = (dblDW / 2) - 140.5
            dblSode1 = 0
            intSode1H = 0
            dblSode2 = 0
            intSode2H = 0
            '20180205 K.Asayama ADD
            dblGakuYokoLVL30 = 0
            intGakuYokoLVL30 = 0
            dblGakutate3 = 0
            intGakutateH3 = 0
            
        '20180205 K.Asayama ADD
        ElseIf dblDH > 2589 Then
            dblCupShitaji = 60
            If IsHirakido(strHinban) Or IsOyatobira(strHinban) Then
                intCupShitajiH = 4 * intMaisu
            Else
                intCupShitajiH = 5 * intMaisu
            End If
            
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
                
            Select Case dblDH
                    
                Case 2530 To 2589
                    dblDaboShitaji = 150
                    intDaboShitajiH = 2 * intMaisu
                Case Is <= 2529
                    dblDaboShitaji = 150
                    intDaboShitajiH = 2 * intMaisu
                    dblCupShitaji = 35
                    intCupShitajiH = 2 * intMaisu
            End Select
        
        
        End If
        
        '20151211 K.Asayama Change 1601仕様追加
        If IsHidden_Hinge(strHinban) Then
            If IsKotobira(strHinban) Then
                dblGakuYoko1 = (dblDW / 2) - 160.5
                intGakuYokoH1 = 2
                dblGakuYoko2 = (dblDW / 2) - 140.5
                intGakuYokoH2 = 2
                intGakutateH1 = 1
                Select Case dblDH
                    Case 2530 To 2589
                        dblGakutate1 = dblDH - 174
                    Case Is <= 2529
                        dblGakutate1 = dblDH - 114
                End Select
            Else
                '20180205 K.Asayama ADD
                If dblDH <= 2589 Then
                    dblGakutate1 = 210
                    intGakutateH1 = 2
                End If
            End If
        End If
        '20151211 K.Asayama Change End
        
'   *TG2*********************************************************

    ElseIf strHinban Like "T?C??*-####C*-*" Then
        
        dblShinAtsu = 30.2
        dblSan = dblDW - 77
        
        intHashiraH = 5 * intMaisu
        dblSode1 = 60
        dblGakuYoko1 = (dblDW / 2) - 208.5
        
        '20160825 K.Asayama Change
        'intSode1H = 5 * intMaisu
        '20161121 K.Asayama Change
'        If IsHikido(strHinban) Then
'            If strHinban Like "*DN-####*-*" Then
'                intSode1H = 2 * intMaisu
'            Else
'                intSode1H = 4 * intMaisu
'            End If
'        Else
'            intSode1H = 5 * intMaisu
'        End If

        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        
        '20160825 K.Asayama Change END
        
        
        Select Case dblDH
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 4 * intMaisu
                dblHashira = dblDH - 114
                dblGakuYokoLVL30 = (dblDW / 2) - 112.5
                intGakuYokoLVL30 = 4 * intMaisu
                dblGakutate3 = 150
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                    intGakutateH3 = 4 * intMaisu
                Else
                    intGakutateH3 = 8 * intMaisu
                End If
                
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 174
                
                '20160825 K.Asayama Change
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END

                    dblGakuYoko2 = (dblDW / 2) - 148.5
                    intGakuYokoH2 = 2
                End If
                '20160825 K.Asayama Change END
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-8"
                Else
                    'strShingumizu = "SS-6"
                End If
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                
                '20160825 K.Asayama Change
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END

                    dblGakuYoko2 = (dblDW / 2) - 148.5
                    intGakuYokoH2 = 2
                End If
                '20160825 K.Asayama Change END
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-7"
                Else
                    'strShingumizu = "SS-5"
                End If
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 1 * intMaisu
                dblHashira = dblDH - 114
                
                '20160825 K.Asayama Change
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END

                    dblGakuYoko2 = (dblDW / 2) - 148.5
                    intGakuYokoH2 = 1
                End If
                '20160825 K.Asayama Change END
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-7"
                Else
                    'strShingumizu = "SS-5"
                End If
                
        End Select
        
        If IsKotobira(strHinban) Then
            
            dblSan = dblDW - 61
            intGakuYokoH1 = 4
            dblGakuYoko1 = (dblDW / 2) - 140.5
            dblSode1 = 0
            intSode1H = 0
            
            '20180205 K.Asayama ADD
            dblGakuYokoLVL30 = 0
            intGakuYokoLVL30 = 0
            dblGakutate3 = 0
            intGakutateH3 = 0
            
        '20180205 K.Asayama ADD
        ElseIf dblDH > 2589 Then
            dblCupShitaji = 60
            intCupShitajiH = 8 * intMaisu
            
        '20160825 K.Asayama Change
        'ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DE-####*" Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Then
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
        '20160825 K.Asayama Change END
        
            Select Case dblDH
                Case 2530 To 2589
                    dblDaboShitaji = 150
                    intDaboShitajiH = 2 * intMaisu
                Case Is <= 2529
                    dblDaboShitaji = 150
                    intDaboShitajiH = 2 * intMaisu
                    dblCupShitaji = 35
                    intCupShitajiH = 2 * intMaisu
            End Select
        End If
        
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
        
            intSode1H = 3 * intMaisu
            dblSode2 = 110
            intSode2H = 2 * intMaisu
            intGakuYokoH1 = intGakuYokoH1 / 2
            dblGakuYoko2 = dblDW - 258.5
            intGakuYokoH2 = intGakuYokoH1
            
        End If
        '******************************************************
        
        '20151211 K.Asayama Change 1601仕様追加
            If IsHidden_Hinge(strHinban) Then
                If IsKotobira(strHinban) Then
                    dblGakuYoko1 = (dblDW / 2) - 160.5
                    intGakuYokoH1 = 2
                    dblGakuYoko2 = (dblDW / 2) - 140.5
                    intGakuYokoH2 = 2
                    intGakutateH1 = 1
                    Select Case dblDH
                        Case 2530 To 2589
                            dblGakutate1 = dblDH - 174
                        Case Is <= 2529
                            dblGakutate1 = dblDH - 114
                    End Select
                Else
                    '20180205 K.Asayama Chnage
                    If dblDH <= 2589 Then
                        dblGakutate1 = 210
                        intGakutateH1 = 2
                    End If
                End If
            End If
        '20151211 K.Asayama Change End

'   *SG1*********************************************************

    ElseIf strHinban Like "F?S??*-####S*-*" Then
    
        dblShinAtsu = 28
        dblSan = dblDW - 64
'        dblGakuYoko1 = dblDW - 380
        intGakuYokoH1 = 1 * intMaisu
        intHashiraH = 5 * intMaisu
        intGakutateH1 = 2 * intMaisu
        dblSode1 = 90
'        intSode1H = 2 * intMaisu
        
        '20160825 K.Asayama Change
        If IsHikido(strHinban) Then
            dblSode2 = 60
            intSode1H = 4 * intMaisu
            '20161121 K.Asayama Change
            'intSode2H = 2 * intMaisu
            intSode2H = intFncSode2Honsu_Group2(strHinban, intMaisu)
            '20161121 K.Asayama Change END
            '20170105 K.Asayama Change
    '        If strHinban Like "*DN-####*" Then
            If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
            '20170105 K.Asayama Change END

                dblGakuYoko1 = dblDW - 380
            Else
                dblGakuYoko1 = dblDW - 440
            End If
        Else
            intSode1H = 2 * intMaisu
            dblGakuYoko1 = dblDW - 380
        End If
        
        '20160825 K.Asayama Change END
        
        Select Case dblDH
            
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                dblHashira = dblDH - 174
                dblGakutate1 = dblDH - 174
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-22"
                Else
                    'strShingumizu = "SS-18"
                End If
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                dblHashira = dblDH - 114
                dblGakutate1 = dblDH - 114
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-21"
                Else
                    'strShingumizu = "SS-17"
                End If
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                dblHashira = dblDH - 114
                dblGakutate1 = dblDH - 114
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-21"
                Else
                    'strShingumizu = "SS-17"
                End If
                
                
        End Select
        
        If IsKotobira(strHinban) Then
            
            '20160510 K.Asayama 本数変更
            intGakuYokoH1 = 2 'intGakuYokoH1 = 4
            
            dblGakuYoko1 = (dblDW / 2) - 145
            dblSode1 = 0
            intSode1H = 0
            
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then

            dblDaboShitaji = 150
            intDaboShitajiH = 2 * intMaisu
            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If
 
        End If
        
'   *SG2*********************************************************

    ElseIf strHinban Like "F?S??*-####C*-*" Then
        
        dblShinAtsu = 28
        dblSan = dblDW - 80
        
        intHashiraH = 5 * intMaisu
        dblSode1 = 60
        'intSode1H = 2 * intMaisu
        dblGakuYoko1 = (dblDW / 2) - 213
        'intGakuYokoH1 = 1 * intMaisu
        'dblGakuYoko2 = (dblDW / 2) - 153
        'intGakuYokoH2 = 1 * intMaisu
                
        '20160825 K.Asayama Change
        If IsHikido(strHinban) Then
            '20161121 K.Asayama Change
            '20170105 K.Asayama Change
    '        If strHinban Like "*DN-####*" Then
            If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
            '20170105 K.Asayama Change END

'                intSode1H = 2 * intMaisu
                intGakuYokoH1 = 1 * intMaisu
            Else
'                intSode1H = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
            End If
            
            intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
            '20161121 K.Asayama Change END
        Else
            intSode1H = 2 * intMaisu
            intGakuYokoH1 = 1 * intMaisu
            dblGakuYoko2 = (dblDW / 2) - 153
            intGakuYokoH2 = 1 * intMaisu
        End If

        '20160825 K.Asayama Change END
        
        Select Case dblDH
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                dblHashira = dblDH - 174
                dblGakutate1 = dblDH - 174
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-22"
                Else
                    'strShingumizu = "SS-20"
                End If
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                dblHashira = dblDH - 114
                dblGakutate1 = dblDH - 114
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-21"
                Else
                    'strShingumizu = "SS-19"
                End If
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                dblHashira = dblDH - 114
                dblGakutate1 = dblDH - 114
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-21"
                Else
                    'strShingumizu = "SS-19"
                End If
                
        End Select
        
        If IsKotobira(strHinban) Then
            
            dblSan = dblDW - 64
            '20160510 K.Asayama 本数変更
            intGakuYokoH1 = 2 'intGakuYokoH1 = 4
            
            dblGakuYoko1 = (dblDW / 2) - 145
            dblGakuYoko2 = 0
            intGakuYokoH2 = 0
            dblSode1 = 0
            intSode1H = 0
            
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DE-####*" Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Then
            Select Case dblDH
                Case 2530 To 2589
                    dblDaboShitaji = 150
                    intDaboShitajiH = 2 * intMaisu
                Case Is <= 2529
                    dblDaboShitaji = 150
                    intDaboShitajiH = 2 * intMaisu
                    dblCupShitaji = 35
                    intCupShitajiH = 2 * intMaisu
            End Select
            
            If strHinban Like "*DE-####*" Then
                intGakutateH1 = 3
            Else
                intGakutateH1 = 2
            End If
        End If
        
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
            
            dblSode1 = 110
            dblGakuYoko1 = dblDW - 263
            
        End If
        '******************************************************
        
'   *XF1*********************************************************
    
    '20160909 K.Asayama Change PF1型の130/70ステルス(X)と品番が被るため条件修正
    'ElseIf strHinban Like "*X?C??*-####F*-*" Then
    ElseIf strHinban Like "X?C??*-####F*-*" Or strHinban Like "特 X?C??*-####F*-*" Then
    '20160909 K.Asayama Change END
    
        dblShinAtsu = 30.2
        dblSan = dblDW + 2
        intSanH = 6
        dblGakuYoko1 = dblDW - 248
        intHashiraH = 2
        dblSode1 = 60
        'intSode1H = 5
        
        '20160825 K.Asayama Change
        '20161121 K.Asayama Change
'        If IsHikido(strHinban) Then
'            intSode1H = 4
'        Else
'            intSode1H = 5
'        End If

        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        '20160825 K.Asayama Change END
        
        Select Case dblDH
            Case 2530 To 2589
                intGakuYokoH1 = 2
                dblHashira = dblDH - 174
                dblGakutate1 = (dblDH - 214) / 3
                intGakutateH1 = 3
                
                'strShingumizu = "SS-2"
                
            Case 1801 To 2529
                intGakuYokoH1 = 2
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-1"
                
            Case Is <= 1800
                intGakuYokoH1 = 1
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-1"
                
        End Select
        
        If strHinban Like "*DC-####*" Or strHinban Like "*KC-####*" Then
        
            Select Case dblDH
                Case Is <= 2529
                    dblCupShitaji = 35
                    intCupShitajiH = 2
            End Select
        End If
        
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
            intSode1H = 3 * intMaisu
            dblSode2 = 110
            intSode2H = 2 * intMaisu
            dblGakuYoko1 = dblDW - 298
        End If
        '******************************************************
        
'   *OF1(旧)*******************************************************
    '20170207 K.Asayama Change
    
'    ElseIf strHinban Like "*O?C??*-####P*-*" Then
'        If strHinban Like "*(PH)*" Then
    '20170412 K.Asayama 1701新品番は内容が違うため下へ移動
    'ElseIf strHinban Like "*O?C??*-####P*-*" Or strHinban Like "*O?C??*-####N*-*" Then
    ElseIf strHinban Like "O?C??*-####P*-*" Then
    '20170412 K.Asayama Change END
    
        'PH,SH色****************************************************
        If strHinban Like "*(PH)*" Or strHinban Like "*(SH)*" Then
    '20170207 K.Asayama Change END
    
            dblShinAtsu = 30.2
            dblShinAtsu_N = 30.2
            dblSan = dblDW - 241.5
            dblGakuYoko1 = (dblDW - 671.5) / 3
            intHashiraH = 3 * intMaisu
            dblSode1 = 100
            '20161121 K.Asayama Change
            'intSode1H = 6 * intMaisu
            intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
            '20161121 K.Asayama Change END
            
            dblCupShitaji = 100
            intCupShitajiH = 20 * intMaisu
            
            dblSan_N = 215.5
            dblGakuYoko1_N = 90.5
            intHashiraH_N = 3 * intMaisu
            
            Select Case dblDH
                Case 2530 To 2589
                    intSanH = 6 * intMaisu
                    intGakuYokoH1 = 6 * intMaisu
                    dblHashira = dblDH - 174
                    
                    intSanH_N = 6 * intMaisu
                    intGakuYokoH1_N = 2 * intMaisu
                    dblHashira_N = dblDH - 174
                    
                    'strShingumizu = "SS-24"
                    
                Case 1801 To 2529
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 6 * intMaisu
                    dblHashira = dblDH - 114
                    
                    intSanH_N = 4 * intMaisu
                    intGakuYokoH1_N = 2 * intMaisu
                    dblHashira_N = dblDH - 114
                    
                    'strShingumizu = "SS-23"
                    
                Case Is <= 1800
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 3 * intMaisu
                    dblHashira = dblDH - 114
                    
                    intSanH_N = 4 * intMaisu
                    intGakuYokoH1_N = 1 * intMaisu
                    dblHashira_N = dblDH - 114
                    
                    'strShingumizu = "SS-23"
                    
            End Select
              
        'PH色以外************************************************
        Else
            
            dblShinAtsu = 30.2
            dblShinAtsu_N = 30.2
            dblSan = dblDW - 240
            dblGakuYoko1 = (dblDW - 670) / 3
            intHashiraH = 3 * intMaisu
            dblSode1 = 100
            '20161121 K.Asayama Change
            'intSode1H = 6 * intMaisu
            intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
            '20161121 K.Asayama Change END
            dblCupShitaji = 100
            intCupShitajiH = 20 * intMaisu
            
            dblGakuYoko1_N = 92
            intHashiraH_N = 3 * intMaisu
            
            Select Case dblDH
                Case 2530 To 2589
                    intSanH = 6 * intMaisu
                    intGakuYokoH1 = 6 * intMaisu
                    dblHashira = dblDH - 174
                    
                    intSanH_N = 6 * intMaisu
                    intGakuYokoH1_N = 2 * intMaisu
                    dblHashira_N = dblDH - 174
                    
                    'strShingumizu = "SS-24"
                    
                Case 1801 To 2529
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 6 * intMaisu
                    dblHashira = dblDH - 114
                    
                    intSanH_N = 4 * intMaisu
                    intGakuYokoH1_N = 2 * intMaisu
                    dblHashira_N = dblDH - 114
                    
                    'strShingumizu = "SS-23"
                    
                Case Is <= 1800
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 3 * intMaisu
                    dblHashira = dblDH - 114
                    
                    intSanH_N = 4 * intMaisu
                    intGakuYokoH1_N = 1 * intMaisu
                    dblHashira_N = dblDH - 114
                    
                    'strShingumizu = "SS-23"
                    
            End Select
              
            If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" Or strHinban Like "*DH-####*" _
                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
                
                dblSan_N = 217
            
            '20170105 K.Asayama Change
'            ElseIf strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Or strHinban Like "*DM-####*" Or strHinban Like "*DL-####*" Or strHinban Like "*DN-####*" Then
            ElseIf strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Or strHinban Like "*DM-####*" Or strHinban Like "*DL-####*" Or strHinban Like "*DN-####*" _
                Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Or strHinban Like "*VM-####*" Or strHinban Like "*VL-####*" Or strHinban Like "*VN-####*" _
            Then
            '20170105 K.Asayama Change END
            
                dblSan_N = 215.5
                
            End If
        End If
        
    '20170412 K.Asayama ADD
'   *OF1(新)*******************************************************
'   *OG1***********************************************************
'   20180205 K.Asayama 柱LVL45追加対応

    '20180205 K.Asayama Change
    'ElseIf strHinban Like "*O?C??*-####N*-*" Then
    ElseIf strHinban Like "O?C??*-####N*-*" Or strHinban Like "O?C??*-####Q*-*" Then
        'SH色****************************************************
        If strHinban Like "*(SH)*" Then
    
            dblShinAtsu = 30.2
            dblShinAtsu_N = 30.2
            dblSan = dblDW - 248.5
            '20180205 K.Asayama Change
            'dblGakuYoko1 = (dblDW - 678.5) / 3
            dblGakuYoko1 = (dblDW - 663.5) / 3
            
            '20180205 K.Asayama Change
            'intHashiraH = 3 * intMaisu
            intHashiraH = 1 * intMaisu
            intHashiraH2 = 1 * intMaisu
            inthashiraH2_N = 1 * intMaisu
            
            dblSode1 = 100
            intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
            
            dblCupShitaji = 100
            intCupShitajiH = 20 * intMaisu
            
            dblSan_N = 215.5
            '20180205 K.Asayama Change
            'dblGakuYoko1_N = 90.5
            dblGakuYoko1_N = 105.5
            
            intHashiraH_N = 3 * intMaisu
            
            Select Case dblDH
                Case 2530 To 2589
                    intSanH = 6 * intMaisu
                    intGakuYokoH1 = 6 * intMaisu
                    dblHashira = dblDH - 174
                    
                    intSanH_N = 6 * intMaisu
                    intGakuYokoH1_N = 2 * intMaisu
                    dblHashira_N = dblDH - 174
                    
                    
                Case 1801 To 2529
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 6 * intMaisu
                    dblHashira = dblDH - 114
                    
                    intSanH_N = 4 * intMaisu
                    intGakuYokoH1_N = 2 * intMaisu
                    dblHashira_N = dblDH - 114

                    
                Case Is <= 1800
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 3 * intMaisu
                    dblHashira = dblDH - 114
                    
                    intSanH_N = 4 * intMaisu
                    intGakuYokoH1_N = 1 * intMaisu
                    dblHashira_N = dblDH - 114

                    
            End Select
              
        'SH色以外************************************************
        Else
            
            dblShinAtsu = 30.2
            dblShinAtsu_N = 30.2
            dblSan = dblDW - 247
            '20180205 K.Asayama Change
            dblGakuYoko1 = (dblDW - 663.5) / 3
            dblGakuYoko1 = (dblDW - 662) / 3
            
            '20180205 K.Asayama Change
            'intHashiraH = 3 * intMaisu
            intHashiraH = 1 * intMaisu
            intHashiraH2 = 1 * intMaisu
            inthashiraH2_N = 1 * intMaisu
            
            dblSode1 = 100

            intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)

            dblCupShitaji = 100
            intCupShitajiH = 20 * intMaisu
            
            '20180205 K.Asayama Change
            'dblGakuYoko1_N = 92
            dblGakuYoko1_N = 107
            
            intHashiraH_N = 3 * intMaisu
            
            Select Case dblDH
                Case 2530 To 2589
                    intSanH = 6 * intMaisu
                    intGakuYokoH1 = 6 * intMaisu
                    dblHashira = dblDH - 174
                    
                    intSanH_N = 6 * intMaisu
                    intGakuYokoH1_N = 2 * intMaisu
                    dblHashira_N = dblDH - 174

                    
                Case 1801 To 2529
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 6 * intMaisu
                    dblHashira = dblDH - 114
                    
                    intSanH_N = 4 * intMaisu
                    intGakuYokoH1_N = 2 * intMaisu
                    dblHashira_N = dblDH - 114

                    
                Case Is <= 1800
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 3 * intMaisu
                    dblHashira = dblDH - 114
                    
                    intSanH_N = 4 * intMaisu
                    intGakuYokoH1_N = 1 * intMaisu
                    dblHashira_N = dblDH - 114

                    
            End Select
              
            If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" Or strHinban Like "*DH-####*" _
                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
                
                dblSan_N = 217
            
            
            ElseIf strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Or strHinban Like "*DM-####*" Or strHinban Like "*DL-####*" Or strHinban Like "*DN-####*" _
                Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Or strHinban Like "*VM-####*" Or strHinban Like "*VL-####*" Or strHinban Like "*VN-####*" _
            Then
            
                dblSan_N = 215.5
                
            End If
            
        End If
        
        '20180205 K.Asayama ADD
        dblHashira2 = dblHashira
        dblhashira2_N = dblHashira_N
            
    '20170412 K.Asayama ADD END
'   *SF1*********************************************************

    ElseIf strHinban Like "F?S??*-####F*-*" Then
        
        dblShinAtsu = 28
        dblSan = dblDW + 2
        intGakuYokoH1 = 1 * intMaisu
        intHashiraH = 2 * intMaisu
        dblSode1 = 60
        intGakutateH1 = 2 * intMaisu
        

        '20160825 K.Asayama Change
'        If strHinban Like "*DH-####*" Then
'            intSode1H = 8
'            dblGakuYoko1 = dblDW - 254
'        ElseIf strHinban Like "*DE-####*" Then
'            intSode1H = 4
'            dblGakuYoko1 = dblDW - 194
'        ElseIf strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Then
'            intSode1H = 12
'            dblGakuYoko1 = dblDW - 254
'        Else
'            intSode1H = 2
'            dblGakuYoko1 = dblDW - 194
'        End If

        If IsHikido(strHinban) Then
            '20161121 K.Asayama Change
            '20170105 K.Asayama Change
    '        If strHinban Like "*DN-####*" Then
            If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
            '20170105 K.Asayama Change END

                'intSode1H = 2 * intMaisu
                dblGakuYoko1 = dblDW - 194
            Else
                'intSode1H = 4 * intMaisu
                dblGakuYoko1 = dblDW - 254
            End If
            
            intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
            '20161121 K.Asayama Change END
        
        Else
            intSode1H = 2 * intMaisu
            dblGakuYoko1 = dblDW - 194
        End If

        '20160825 K.Asayama Change END
        
        Select Case dblDH
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                dblGakutate1 = dblDH - 174
                dblHashira = dblDH - 174
                
                'strShingumizu = "SS-16"
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                dblGakutate1 = dblDH - 114
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-15"
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                dblGakutate1 = dblDH - 114
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-15"
                
        End Select
        
        '20170105 K.Asayama Change
'        If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" Or strHinban Like "*DH-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" _
'            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
        If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" Or strHinban Like "*DH-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" _
            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Or strHinban Like "*VF-####*" Then
        '20170105 K.Asayama Change END
            
            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If
        End If
        
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
            
            dblSode1 = 110
            dblGakuYoko1 = dblDW - 244
            
        End If
        '******************************************************
        
'   *PF6*********************************************************

    ElseIf strHinban Like "P?D??*-####F*-*" Then
    
        dblShinAtsu = 26.6
        dblSan = dblDW + 4
        'dblGakuYoko1 = dblDW - 256
        intHashiraH = 2 * intMaisu
        dblSode1 = 60
        
        '20160825 K.Asayama Change
        If strHinban Like "*DH-####*" Then
'            intSode1H = 8
            intGakutateH1 = 6
        ElseIf strHinban Like "*DE-####*" Then
'            intSode1H = 4
            intGakutateH1 = 8
        '20170105 K.Asayama Change
'        ElseIf strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Then
        ElseIf strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*VF-####*" Then
        '20170105 K.Asayama Change END
        
'            intSode1H = 12
            intGakutateH1 = 12
        Else
'            intSode1H = 2
            intGakutateH1 = 3
        End If
        '20160825 K.Asayama Change END
        
        '20160825 K.Asayama Change
        If IsHikido(strHinban) Then
            '20161121 K.Asayama Change
            
            '20170105 K.Asayama Change
    '        If strHinban Like "*DN-####*" Then
            If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
            '20170105 K.Asayama Change END
            
                'intSode1H = 2 * intMaisu
                dblGakuYoko1 = dblDW - 256
                intGakuYokoH1 = 5 * intMaisu
            Else
                'intSode1H = 4 * intMaisu
                dblGakuYoko1 = dblDW - 336
                intGakuYokoH1 = 4 * intMaisu
            End If
            
            intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
            '20161121 K.Asayama Change END
        
        Else
            intSode1H = 2 * intMaisu
            dblGakuYoko1 = dblDW - 256
            intGakuYokoH1 = 4 * intMaisu
        End If

        '20160825 K.Asayama Change END
        
        intSanH = 4 * intMaisu
'        intGakuYokoH1 = 4 * intMaisu
        dblHashira = dblDH - 114
        dblGakutate1 = dblDH - 114
        '20160510 K.Asayama Change 額縦2と3廃止
        'intGakutateH2 = 1 * intMaisu
        'intGakutateH3 = 1 * intMaisu
        
        'strShingumizu = "SS-29"
        
        If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
            
            dblCupShitaji = 35
            intCupShitajiH = 2 * intMaisu
            '20160510 K.Asayama Change 額縦2と3廃止
            'dblGakutate2 = 943
            'dblGakutate3 = dblDH - 1137
            
        '20170105 K.Asayama Change
'        ElseIf strHinban Like "*DM-####*" Or strHinban Like "*DL-####*" Or strHinban Like "*DN-####*" Then
        ElseIf strHinban Like "*DM-####*" Or strHinban Like "*DL-####*" Or strHinban Like "*DN-####*" Or strHinban Like "*VM-####*" Or strHinban Like "*VL-####*" Or strHinban Like "*VN-####*" Then
        '20170105 K.Asayama Change END
        
            dblCupShitaji = 35
            intCupShitajiH = 2
            '20160510 K.Asayama Change 額縦2と3廃止
            'dblGakutate2 = 943
            'dblGakutate3 = dblDH - 1137
            
        Else '開き
            '20160510 K.Asayama Change 額縦2と3廃止
            'dblGakutate2 = 940
            'dblGakutate3 = dblDH - 1134
        End If
        
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
            
            dblSode1 = 110
            dblGakuYoko1 = dblDW - 306
            
        End If
        '******************************************************
        
        '20151211 K.Asayama Change 1601仕様追加
        If IsHidden_Hinge(strHinban) Then
            dblGakuYoko1 = dblDW - 276
            intGakutateH1 = 4
        End If
        '20151211 K.Asayama Change End
        
'   *GG2*********************************************************

    ElseIf strHinban Like "G?C??*-####C*-*" Then
        
        dblShinAtsu = 35.5
        dblSan = dblDW - 113
        intSanH = 6
        
        '20160706 K.Asayama Change
        'dblGakuYoko1 = (dblDW / 2) - 281.5
        '20160819 K.Asayama 誤り訂正
        'dblGakuYoko1 = dblDW - 280.5
        dblGakuYoko1 = (dblDW / 2) - 280.5
        '20160819 K.Asayama Change End
        '20160706 K.Asayama Change END
        
        dblHashira = dblDH - 174
        intHashiraH = 9
        dblSode1 = 60
        intSode1H = 2
        
        'strShingumizu = "SS-43"
        
        Select Case dblDH
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intGakuYokoH1 = 4
                dblGakutate3 = 150
                intGakutateH3 = 4 * intMaisu

            Case 2530 To 2589
                intGakuYokoH1 = 4
                
            Case 1801 To 2529
                intGakuYokoH1 = 4
                
            Case Is <= 1800
                intGakuYokoH1 = 2
                
        End Select
        
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
        
            dblTegake = 90
            dblSode1 = 90
            intGakuYokoH1 = intGakuYokoH1 / 2
            dblGakuYoko2 = (dblDW / 2) - 311.5
            intGakuYokoH2 = intGakuYokoH1
            
        End If
        '******************************************************
        
'   *PG2*********************************************************

    ElseIf strHinban Like "P?C??*-####C*-*" Then
        
        dblShinAtsu = 26.6
        dblSan = dblDW - 79
        intSanH = 4 * intMaisu
        
        dblHashira = dblDH - 114
        intHashiraH = 5 * intMaisu
        
        dblSode1 = 60
        dblGakuYoko1 = (dblDW / 2) - 254.5
        
        '20160825 K.Asayama Change
'        intSode1H = 2 * intMaisu
'        dblGakuYoko2 = (dblDW / 2) - 174.5
'        intGakutateH1 = 3 * intMaisu

        '   20161121 K.Asayama Change
        If IsHikido(strHinban) Then
        
            '20170105 K.Asayama Change
    '        If strHinban Like "*DN-####*" Then
            If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
            '20170105 K.Asayama Change END

                'intSode1H = 2 * intMaisu
                dblGakuYoko2 = (dblDW / 2) - 174.5
                intGakutateH1 = 3 * intMaisu
            Else
                'intSode1H = 4 * intMaisu
                dblGakuYoko2 = (dblDW / 2) - 254.5
                intGakutateH1 = 4 * intMaisu
            End If
            intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '   20161121 K.Asayama Change END
        Else
            intSode1H = 2 * intMaisu
            dblGakuYoko2 = (dblDW / 2) - 174.5
            intGakutateH1 = 3 * intMaisu
        End If

        '20160825 K.Asayama Change END
        
        intGakuYokoH1 = 1 * intMaisu
        intGakuYokoH2 = 1 * intMaisu
        
        dblGakutate1 = dblDH - 114
        
        If IsKotobira(strHinban) Then
            'strShingumizu = "SS-28"
        Else
            'strShingumizu = "SS-27"
        End If
        
        If IsKotobira(strHinban) Then
            
            dblSan = dblDW - 63
            intGakuYokoH1 = 4
            dblGakuYoko1 = (dblDW / 2) - 146.5
            dblSode1 = 0
            intSode1H = 0
            
            dblGakuYoko2 = 0
            intGakuYokoH2 = 0
            
            intGakutateH1 = 2
            
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then

            dblDaboShitaji = 150
            intDaboShitajiH = 2 * intMaisu
            dblCupShitaji = 35
            intCupShitajiH = 2 * intMaisu

        End If
        
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
            
            dblSode1 = 110
            dblGakuYoko1 = dblDW - 304.5
            
        End If
        '******************************************************
        
        '20151211 K.Asayama Change 1601仕様追加
            If IsHidden_Hinge(strHinban) Then
                intGakutateH1 = 4
                
                If IsKotobira(strHinban) Then
                    dblGakuYoko1 = (dblDW / 2) - 186.5
                    intGakuYokoH1 = 2
                    dblGakuYoko2 = (dblDW / 2) - 166.5
                    intGakuYokoH2 = 2
                '20160819 K.Asayama 誤り訂正
                Else
                    dblGakuYoko2 = (dblDW / 2) - 194.5
                '20160819 K.Asayama ADD End
                End If
            End If
        '20151211 K.Asayama Change End
        
'   *GG1*********************************************************

    ElseIf strHinban Like "G?C??*-####S*-*" Then
        
        dblShinAtsu = 35.5
        dblSan = dblDW - 97
        intSanH = 6
        
        '20160706 K.Asayama Change
        'dblGakuYoko1 = dblDW - 500.5
        dblGakuYoko1 = dblDW - 499
        '20160706 K.Asayama Change END
        dblHashira = dblDH - 174
        intHashiraH = 9
        dblSode1 = 73.5
        intSode1H = 2
        
        'strShingumizu = "SS-43"
        
        Select Case dblDH
            
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intGakuYokoH1 = 2
                
            Case 2530 To 2589
                intGakuYokoH1 = 2
                
            Case 1801 To 2529
                intGakuYokoH1 = 2
                
            Case Is <= 1800
                intGakuYokoH1 = 1
                
        End Select
        
'   *PG1*********************************************************

    ElseIf strHinban Like "P?C??*-####S*-*" Then
        
        dblShinAtsu = 26.6
        dblSan = dblDW - 63
        intSanH = 4 * intMaisu
        
        dblHashira = dblDH - 114
        intHashiraH = 5 * intMaisu
        dblSode1 = 90.5
        
        '20160825 K.Asayama Change
'        intSode1H = 2 * intMaisu
        '20151019 K.Asayama Change 誤り訂正
        'dblGakuYoko1 = (dblDW / 2) - 432.5
'        dblGakuYoko1 = dblDW - 423.5
        '20151019 K.Asayama Change End
        intGakuYokoH1 = 1 * intMaisu
        If IsHikido(strHinban) Then
            intSode1H = 4 * intMaisu
            
            '20170105 K.Asayama Change
    '        If strHinban Like "*DN-####*" Then
            If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
            '20170105 K.Asayama Change END

                dblGakuYoko1 = dblDW - 423.5
            Else
                dblSode2 = 60
                '20161121 K.Asayama Change
                'intSode2H = 2 * intMaisu
                intSode2H = intFncSode2Honsu_Group2(strHinban, intMaisu)
                '20161121 K.Asayama Change END
                '20170105 K.Asayama Change
'                If strHinban Like "*DM-####*-*" Or strHinban Like "*DL-####*-*" Then
                If strHinban Like "*DM-####*-*" Or strHinban Like "*DL-####*-*" Or strHinban Like "*VM-####*-*" Or strHinban Like "*VL-####*-*" Then
                '20170105 K.Asayama Change
                
                    dblGakuYoko1 = dblDW - 483.5
                Else
                    dblGakuYoko1 = dblDW - 503.5
                End If
            End If
        Else
            intSode1H = 2 * intMaisu
            dblGakuYoko1 = dblDW - 423.5
        End If

        '20160825 K.Asayama Change END
        

        
        dblGakutate1 = dblDH - 114
        intGakutateH1 = 2 * intMaisu
        
        If IsKotobira(strHinban) Then
            'strShingumizu = "SS-28"
        Else
            'strShingumizu = "SS-26"
        End If
                
        If IsKotobira(strHinban) Then
            
            intGakuYokoH1 = 4
            dblGakuYoko1 = (dblDW / 2) - 146.5
            dblSode1 = 0
            intSode1H = 0
            
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then

            dblDaboShitaji = 150
            intDaboShitajiH = 2 * intMaisu
            dblCupShitaji = 35
            intCupShitajiH = 2 * intMaisu

        End If
        
        '20151211 K.Asayama Change 1601仕様追加
        If IsHidden_Hinge(strHinban) Then
        
            intGakutateH1 = 3
    
            If IsKotobira(strHinban) Then
                dblGakuYoko1 = (dblDW / 2) - 186.5
                intGakuYokoH1 = 2
            Else
                dblGakuYoko1 = dblDW - 443.5
            End If
            
        End If
        '20151211 K.Asayama Change End

'   *GF1*********************************************************

    ElseIf strHinban Like "G?C??*-####F*-*" Then
        
        dblShinAtsu = 35.5
        dblSan = dblDW - 30
        intSanH = 6
        '20160706 K.Asayama Change
        'dblGakuYoko1 = dblDW - 330
        dblGakuYoko1 = dblDW - 328.5
        '20160706 K.Asayama Change END
        
        dblHashira = dblDH - 174
        intHashiraH = 6
        dblSode1 = 60
        intSode1H = 2
        
        'strShingumizu = "SS-43"
        
        Select Case dblDH
            
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intGakuYokoH1 = 2
                dblGakutate3 = 150
                intGakutateH3 = 4 * intMaisu
                
            Case 2530 To 2589
                intGakuYokoH1 = 2
                
            Case 1801 To 2529
                intGakuYokoH1 = 2
                
            Case Is <= 1800
                intGakuYokoH1 = 1
                
        End Select
        
        
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
            
            dblTegake = 90
            dblSode1 = 90
            dblGakuYoko1 = dblDW - 360
            
        End If
        '******************************************************
        
'   *PF1*********************************************************

    ElseIf strHinban Like "P?C??*-####F*-*" Then
    
        dblShinAtsu = 26.6
        dblSan = dblDW + 4
'        dblGakuYoko1 = dblDW - 256
        intHashiraH = 2 * intMaisu
        dblSode1 = 60
        
        '20160825 K.Asayama Change
'        If strHinban Like "*DH-####*" Then
'            intSode1H = 8
'            intGakutateH1 = 8
'            dblGakuYoko1 = dblDW - 336
'
'        ElseIf strHinban Like "*DE-####*" Then
'            intSode1H = 4
'            intGakutateH1 = 6
'        '20160609 DQ追加
'        'ElseIf strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Then
'        ElseIf strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Then
'            intSode1H = 12
'            intGakutateH1 = 12
'            dblGakuYoko1 = dblDW - 336
'        Else
'            intSode1H = 2
'            intGakutateH1 = 3
'        End If
        
        If IsHikido(strHinban) Then
            '20161121 K.Asayama Change
            '20170105 K.Asayama Change
    '        If strHinban Like "*DN-####*" Then
            If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
            '20170105 K.Asayama Change END

                'intSode1H = 2 * intMaisu
                intGakutateH1 = 3 * intMaisu
                dblGakuYoko1 = dblDW - 256
            Else
                'intSode1H = 4 * intMaisu
                intGakutateH1 = 4 * intMaisu
                dblGakuYoko1 = dblDW - 336
            End If
            
            intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
            '20161121 K.Asayama Change END
        Else
            intSode1H = 2 * intMaisu
            intGakutateH1 = 3 * intMaisu
            dblGakuYoko1 = dblDW - 256
        End If
            
            
        
        intSanH = 4 * intMaisu
        intGakuYokoH1 = 1 * intMaisu
        dblHashira = dblDH - 114
        dblGakutate1 = dblDH - 114
        '20160510 K.Asayama Change 額縦2廃止
        'dblGakutate2 = (dblDH / 2) - 67
        'intGakutateH2 = 2 * intMaisu
        
        'strShingumizu = "SS-25"
        
        '20170105 K.Asayama Change
'        If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" _
'            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
        If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" _
            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Then
        '20170105 K.Asayama Change END
        
            dblCupShitaji = 35
            intCupShitajiH = 2 * intMaisu

        End If
        
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
            intSode1H = 2 * intMaisu
            dblSode1 = 110
            dblGakuYoko1 = dblDW - 306
        End If
        '******************************************************
        
        '20151211 K.Asayama Change 1601仕様追加
            If IsHidden_Hinge(strHinban) Then
                dblGakuYoko1 = dblDW - 276
                intGakutateH1 = 4
            End If
        '20151211 K.Asayama Change End

'   *FA2*********************************************************

    ElseIf strHinban Like "A?C??*-####SL*-*" Then
    
        dblShinAtsu = 30.2
        
        'シナ****************************************************
        If IsSINAColor(strHinban) Then
        
            dblSan = dblDW - 61
            intSanH = 6
            dblGakuYoko1 = (dblDW / 2) - 522.5
            intGakuYokoH1 = 2
            
            intHashiraH = 6
            dblSode1 = 91.5
            dblSode2 = 60
            
            '20160825 K.Asayama Change
            'intSode1H = 2
            'intSode2H = 6
            
            '20161121 K.Asayama Change
            If IsHikido(strHinban) Then
                intSode1H = 4 * intMaisu
                'intSode2H = 5 * intMaisu
            Else
                intSode1H = 2 * intMaisu
                'intSode2H = 6 * intMaisu
            End If
            
            
            intSode2H = intFncSode2Honsu_Group1(strHinban, intMaisu)
            '20161121 K.Asayama Change END
        
            '20160825 K.Asayama Change End
            
            Select Case dblDH
            
                Case 2530 To 2589
                    dblHashira = dblDH - 174
                    
                    'strShingumizu = "SS-36"
                    
                Case 1801 To 2529
                    dblHashira = dblDH - 114
                    
                    'strShingumizu = "SS-35"
                    
                Case Is <= 1800
                    dblHashira = dblDH - 114
                    
                    'strShingumizu = "SS-35"
                    
            End Select
            
            If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DU-####*" _
                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
                
                '20151211 K.Asayama Change 1601仕様追加
                dblDaboShitaji = 150
                intDaboShitajiH = 1
                '20151211 K.Asayama Change End
                
                If dblDH <= 2529 Then
                    dblCupShitaji = 35
                    intCupShitajiH = 2
                End If
 
            End If
        
        'シナ以外************************************************
        Else
        
            dblSan = dblDW - 64
            dblGakuYoko1 = dblDW - 524
            
            
            intHashiraH = 6
            dblSode1 = 90
            dblSode2 = 60
            
            '20160825 K.Asayama Change
            'intSode1H = 2
            'intSode2H = 6
            
            '20161121 K.Asayama Change
            If IsHikido(strHinban) Then
                intSode1H = 4 * intMaisu
                'intSode2H = 5 * intMaisu
            Else
                intSode1H = 2 * intMaisu
                'intSode2H = 6 * intMaisu
            End If
            
            
            intSode2H = intFncSode2Honsu_Group1(strHinban, intMaisu)
            '20161121 K.Asayama Change END
        
            '20160825 K.Asayama Change End
            
            Select Case dblDH
            
                Case 2530 To 2589
                    dblHashira = dblDH - 174
                    
                    If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DU-####*" _
                        Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
                        
                        intSanH = 4
                        intGakuYokoH1 = 2
                    Else
                        intSanH = 6
                        intGakuYokoH1 = 2
                    End If
                    
                    'strShingumizu = "SS-36"
                    
                Case 1801 To 2529
                    dblHashira = dblDH - 114
                    
                    If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DU-####*" _
                        Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
                        
                        intSanH = 4
                        intGakuYokoH1 = 1
                        dblCupShitaji = 35
                        intCupShitajiH = 2
                    Else
                        intSanH = 4
                        intGakuYokoH1 = 2
                    End If
                    
                    'strShingumizu = "SS-35"
                    
                Case Is <= 1800
                    dblHashira = dblDH - 114
                    
                    If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DU-####*" _
                        Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
                        
                        intSanH = 6
                        intGakuYokoH1 = 2
                        dblCupShitaji = 35
                        intCupShitajiH = 2
                    Else
                        intSanH = 4
                        intGakuYokoH1 = 1
                    End If
                    
                    'strShingumizu = "SS-35"
                    
            End Select
            
        End If
        
        '20151211 K.Asayama Change 1601仕様追加
        If IsHidden_Hinge(strHinban) Then
            dblGakutate1 = 210
            intGakutateH1 = 2
        End If
        '20151211 K.Asayama Change End
        
'   *CG7/EG7/ZG7*************************************************
    '1608以降 20160923 K.Asayama ADD
    '20180214 K.Asayama Change 中板上下桟全面改定
    ElseIf strHinban Like "F?C??*-####M*-*" Then
        dblShinAtsu = 30.2
        intHashiraH = 5 * intMaisu
        dblSode1 = 60
        
        dblShinAtsu_N = 14.8
        
        '20161121 K.Asayama Change
'        If IsHikido(strHinban) Then
'            If strHinban Like "*DN-####*-*" Then
'                intSode1H = 2 * intMaisu
'            Else
'                intSode1H = 4 * intMaisu
'            End If
'        Else
'            intSode1H = 5 * intMaisu
'        End If
        
        
        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        
        Select Case dblDH
            
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intSanH = 4 * intMaisu
                
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                    intGakuYokoH1 = 2 * intMaisu
                    intGakuYokoH2 = 2 * intMaisu
                Else
                    intGakuYokoH1 = 4 * intMaisu
                End If
                
                dblHashira = dblDH - 114
                
                'dblHashira_N = dblDH - 100
                dblHashira_N = dblDH - 85
                intHashiraH_N = 2 * intMaisu
                
                'intSanH_N = 2 * intMaisu
                'intsanh2_N = 1 * intMaisu
                intsanh2_N = 2 * intMaisu
                intGakuYokoH1_N = 7 * intMaisu
                
                intGakuYokoLVL30 = 4 * intMaisu
                dblGakutate3 = 150
                intGakutateH3 = 8 * intMaisu
                
                
                '20180205 K.Asayama ADD
                '2700固定値取得
                bolFncLVL30_Koteichi dblDW, dblDH, strHinban, dblGakuYokoLVL30, dblCupShitaji

            Case 2530 To 2589
                intSanH = 6 * intMaisu
                
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END

                    intGakuYokoH1 = 2 * intMaisu
                Else
                    intGakuYokoH1 = 4 * intMaisu
                End If
                
                dblHashira = dblDH - 174
                
                '20180205 K.Asayama Change
                'dblHashira_N = dblDH - 175
                'dblHashira_N = dblDH - 160
                dblHashira_N = dblDH - 175
                intHashiraH_N = 2 * intMaisu
                
                '20180205 K.Asayama Change
                'intSanH_N = 6 * intMaisu
                'intSanH_N = 4 * intMaisu
                'intsanh2_N = 1 * intMaisu
                intsanh2_N = 4 * intMaisu
                
                intGakuYokoH1_N = 7 * intMaisu
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END

                    intGakuYokoH1 = 2 * intMaisu
                Else
                    intGakuYokoH1 = 4 * intMaisu
                End If
                
                dblHashira = dblDH - 114
                
                '20180205 K.Asayama Change
                'dblHashira_N = dblDH - 115
'                dblHashira_N = dblDH - 100
                dblHashira_N = dblDH - 85
                intHashiraH_N = 2 * intMaisu
                
                '20180205 K.Asayama Change
                'intSanH_N = 4 * intMaisu
'                intSanH_N = 2 * intMaisu
'                intsanh2_N = 1 * intMaisu
                intsanh2_N = 2 * intMaisu
                intGakuYokoH1_N = 6 * intMaisu
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END

                    intGakuYokoH1 = intMaisu
                Else
                    intGakuYokoH1 = 2 * intMaisu
                End If
                
                dblHashira = dblDH - 114
                
                '20180205 K.Asayama Change
                'dblHashira_N = dblDH - 115
'                dblHashira_N = dblDH - 100
                dblHashira_N = dblDH - 85
                intHashiraH_N = 2 * intMaisu
                
                '20180205 K.Asayama Change
                'intSanH_N = 4 * intMaisu
'                intSanH_N = 2 * intMaisu
'                intsanh2_N = 1 * intMaisu
                intsanh2_N = 2 * intMaisu
                intGakuYokoH1_N = 5 * intMaisu
                
        End Select
        
        '縦框関数
        If bolFncSan_Koteichi(dblDW, dblDH, strHinban, dblSan, dblGakuYoko1, strShingumizu) Then
            
            '20170105 K.Asayama Change
    '        If strHinban Like "*DN-####*" Then
            If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
            '20170105 K.Asayama Change END

                dblGakuYoko2 = dblGakuYoko1 + 60
            End If
        
        Else
            'エラー
        End If
        
        '中板関数
        '20180214 K.Asayama Change
       ' If bolFncSan_Koteichi_Nakaita(dblDW, dblDH, strHinban, dblSan_N, dblGakuYoko1_N) Then
        If bolFncSan_Koteichi_Nakaita(dblDW, dblDH, strHinban, dblsanH2_N, dblGakuYoko1_N) Then
            '20180205 K.Asayama ADD
'            dblsanH2_N = dblSan_N
        Else
            'エラー（0を送る)
        End If
        
        '20180205 K.Asayama ADD
        If dblDH > 2589 Then
            If dblCupShitaji > 0 Then
                If dblDW < 571 Then
                    intCupShitajiH = 4 * intMaisu
                Else
                    intCupShitajiH = 8 * intMaisu
                End If
            End If
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then

            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If
            
        End If
 
            
        If IsHidden_Hinge(strHinban) Then
            '20180205 K.Asayama ADD
            If dblDH <= 2589 Then
                dblGakutate1 = 210
                intGakutateH1 = 2
            End If
        End If
    '20160923 K.Asayama ADD END
    
'   *（旧）CG7/EG7/ZG7*************************************************
    '1608にて廃盤（1601まで）
    ElseIf strHinban Like "F?B??*-####M*-*" Then

        dblShinAtsu = 30.2
        intHashiraH = 5 * intMaisu
        dblSode1 = 60
        intSode1H = 5 * intMaisu
        
        dblDaboShitaji = 150
        intDaboShitajiH = 4 * intMaisu
        
        Select Case dblDH
            
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                intGakuYokoH1 = 4 * intMaisu
                dblHashira = dblDH - 174
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-12"
                Else
                    'strShingumizu = "SS-10"
                End If

            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 4 * intMaisu
                dblHashira = dblDH - 114
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-11"
                Else
                    'strShingumizu = "SS-9"
                End If
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-11"
                Else
                    'strShingumizu = "SS-9"
                End If
                
        End Select
        
        If IsKotobira(strHinban) Then
        
            intGakuYokoH1 = 4
            dblSan = 339
            dblGakuYoko1 = 59.5
            dblSode1 = 0
            intSode1H = 0
            
        Else
            '20160903 K.Asayama Change 廃盤のため固定値化
            'If bolFncSan_Koteichi(dblDW, dblDH, strHinban, dblSan, dblGakuYoko1, strShingumizu) Then
                
            'Else
            '    'エラー
            'End If
            
            intTegakeH = 1 * intMaisu
            
            Select Case dblDW
                Case 426 To 570.9
                    dblSan = 346: dblGakuYoko1 = 3
                    
                Case 571 To 618.9
                    dblSan = 390: dblGakuYoko1 = 25
                
                Case 619 To 669.9
                    dblSan = 422: dblGakuYoko1 = 41
                
                Case 670 To 717.9
                    dblSan = 454: dblGakuYoko1 = 57
                
                Case 718 To 750.9
                    dblSan = 488: dblGakuYoko1 = 74
                            
                Case 751 To 780.9
                    dblSan = 502: dblGakuYoko1 = 81
                
                Case 781 To 819.9
                    dblSan = 526: dblGakuYoko1 = 93
                    
                Case 820 To 862.9
                    dblSan = 552: dblGakuYoko1 = 106
                            
                Case 863 To 900.9
                    dblSan = 576: dblGakuYoko1 = 118
                            
                Case 901 To 944.9
                    dblSan = 610: dblGakuYoko1 = 135
                            
                Case 945 To 985.9
                    dblSan = 638: dblGakuYoko1 = 149
                
                Case 986 To 1022.9
                    dblSan = 662: dblGakuYoko1 = 161
                            
                Case 1023 To 1061.9
                    dblSan = 688: dblGakuYoko1 = 174
                    
                Case 1062 To 1100
                    dblSan = 710: dblGakuYoko1 = 185
                    
                Case Else
                    dblSan = 0: dblGakuYoko1 = 0
            End Select
            
            '20160903 K.Asayama Change
            
            If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then

                If dblDH <= 2529 Then
                    dblCupShitaji = 35
                    intCupShitajiH = 2 * intMaisu
                End If
                
            End If
 
        End If
        
        '20151211 K.Asayama Change 1601仕様追加
        If IsHidden_Hinge(strHinban) Then
            If IsKotobira(strHinban) Then
                intGakuYokoH1 = 2
                dblGakuYoko2 = 39.5
                intGakuYokoH2 = 2
                intGakutateH1 = 1
                Select Case dblDH
                    Case 2530 To 2589
                        dblGakutate1 = dblDH - 174
                    Case Is <= 2529
                        dblGakutate1 = dblDH - 114
                End Select
            Else
                dblGakutate1 = 210
                intGakutateH1 = 2
            End If
        End If
        '20151211 K.Asayama Change End
        
 '   *CG3/EG3/ZG3*************************************************

    ElseIf strHinban Like "F?B??*-####G*-*" Then
    
        dblShinAtsu = 30.2
        intHashiraH = 5 * intMaisu
        dblSode1 = 60
        '20160825 K.Asayama Change
        'intSode1H = 5 * intMaisu
        
        '20161121 K.Asayama Change
'        If IsHikido(strHinban) Then
'            If strHinban Like "*DN-####*-*" Then
'                intSode1H = 2 * intMaisu
'            Else
'                intSode1H = 4 * intMaisu
'            End If
'        Else
'            intSode1H = 5 * intMaisu
'        End If
        '20160825 K.Asayama Change END

        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        
        '20801** K.Asayama ADD
        If dblDH > 2589 Then
            If IsKotobira(strHinban) Then
                dblDaboShitaji = 150
                intDaboShitajiH = 4 * intMaisu
            End If
        Else
            dblDaboShitaji = 150
            '20160819 K.Asayama 誤り訂正
            'intDaboShitajiH = 2 * intMaisu
            intDaboShitajiH = 4 * intMaisu
            '20160819 K.Asayama Change End
        End If
        
        Select Case dblDH
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intSanH = 4 * intMaisu
                
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                    intGakuYokoH1 = 2 * intMaisu
                    intGakuYokoH2 = 2 * intMaisu
                    intGakutateH3 = 4 * intMaisu
                Else
                    intGakuYokoH1 = 4 * intMaisu
                    intGakutateH3 = 8 * intMaisu
                End If

                dblHashira = dblDH - 114
                intGakuYokoLVL30 = 4 * intMaisu
                dblGakutate3 = 150
                
                '20180205 K.Asayama ADD
                '2700固定値取得
                bolFncLVL30_Koteichi dblDW, dblDH, strHinban, dblGakuYokoLVL30, dblCupShitaji
                
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                '20160825 K.Asayama Change
                'intGakuYokoH1 = 4 * intMaisu
                
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END

                    intGakuYokoH1 = 2 * intMaisu
                    intGakuYokoH2 = 2 * intMaisu
                Else
                    intGakuYokoH1 = 4 * intMaisu
                End If
                '20160825 K.Asayama Change END
                
                dblHashira = dblDH - 174
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-12"
                Else
                    'strShingumizu = "SS-10"
                End If
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                '20160825 K.Asayama Change
                'intGakuYokoH1 = 4 * intMaisu
                
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END

                    intGakuYokoH1 = 2 * intMaisu
                    intGakuYokoH2 = 2 * intMaisu
                Else
                    intGakuYokoH1 = 4 * intMaisu
                End If
                '20160825 K.Asayama Change END
                
                dblHashira = dblDH - 114
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-11"
                Else
                    'strShingumizu = "SS-9"
                End If
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                '20160825 K.Asayama Change
                'intGakuYokoH1 = 2 * intMaisu
                
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END
                
                    intGakuYokoH1 = intMaisu
                    intGakuYokoH2 = intMaisu
                Else
                    intGakuYokoH1 = 2 * intMaisu
                End If
                '20160825 K.Asayama Change END
                
                dblHashira = dblDH - 114
                
                If IsKotobira(strHinban) Then
                    'strShingumizu = "SS-11"
                Else
                    'strShingumizu = "SS-9"
                End If
                
        End Select
        
        If IsKotobira(strHinban) Then
        
            intGakuYokoH1 = 4
            dblSan = 339
            dblGakuYoko1 = 59.5
            dblSode1 = 0
            intSode1H = 0
            '20180205 K.Asayama ADD
            dblGakuYokoLVL30 = 0
            intGakuYokoLVL30 = 0
            dblGakutate3 = 0
            intGakutateH3 = 0
            
        Else
            If bolFncSan_Koteichi(dblDW, dblDH, strHinban, dblSan, dblGakuYoko1, strShingumizu) Then
            '20160825 K.Asayama ADD
                '20170105 K.Asayama Change
        '        If strHinban Like "*DN-####*" Then
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                '20170105 K.Asayama Change END

                    dblGakuYoko2 = dblGakuYoko1 + 60
                End If
            '20160825 K.Asayama ADD END
            
            Else
                'エラー
            End If
            '20180205 K.Asayama Change
            If dblDH > 2589 Then
                If dblCupShitaji > 0 Then
                    intCupShitajiH = 8 * intMaisu
                End If
                
            ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then

                If dblDH <= 2529 Then
                    dblCupShitaji = 35
                    intCupShitajiH = 2 * intMaisu
                End If
                
            End If
            
            
            
        End If
        
        '20151211 K.Asayama Change 1601仕様追加
        If IsHidden_Hinge(strHinban) Then
            If IsKotobira(strHinban) Then
                intGakuYokoH1 = 2
                dblGakuYoko2 = 39.5
                intGakuYokoH2 = 2
                intGakutateH1 = 1
                Select Case dblDH
                    '20180205 K.Asayama ADD
                    Case Is > 2589
                        dblGakutate1 = dblDH - 114
                    Case 2530 To 2589
                        dblGakutate1 = dblDH - 174
                    Case Is <= 2529
                        dblGakutate1 = dblDH - 114
                End Select
            Else
                '20180205 K.Asayama Change
                If dblDH <= 2589 Then
                    dblGakutate1 = 210
                    intGakutateH1 = 2
                End If
            End If
        End If
        '20151211 K.Asayama Change End

'   *BG2*********************************************************

    ElseIf strHinban Like "B?C??*-####C*-*" Then
        
        
        If IsPALIOBlack(strHinban) Then
        
        Else
            dblShinAtsu = 30.2
            dblSan = dblDW - 77
            dblGakuYoko1 = (dblDW / 2) - 208.5
            intHashiraH = 5 * intMaisu
            dblSode1 = 60
            
            '20161121 K.Asayama Change
            'intSode1H = 5 * intMaisu
            
            intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
            '20161121 K.Asayama Change END
            
            Select Case dblDH
                
                '20180205 K.Asayama ADD
                Case 2589.5 To 2689
                    intSanH = 4 * intMaisu
    
                    dblHashira = dblDH - 114
                    dblGakuYokoLVL30 = (dblDW / 2) - 112
                    intGakuYokoLVL30 = 4 * intMaisu
                    dblGakutate3 = 150
                    
                    If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                        intGakuYokoH1 = 2 * intMaisu
                        intGakutateH3 = 4 * intMaisu
                    Else
                        intGakuYokoH1 = 4 * intMaisu
                        intGakutateH3 = 8 * intMaisu
                    End If
                    
                Case 2530 To 2589
                    intSanH = 6 * intMaisu
                    intGakuYokoH1 = 4 * intMaisu
                    dblHashira = dblDH - 174
                    
                    If IsKotobira(strHinban) Then
                        'strShingumizu = "SS-8"
                    Else
                        'strShingumizu = "SS-6"
                    End If
                
                Case 1801 To 2529
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 4 * intMaisu
                    dblHashira = dblDH - 114
                    
                    If IsKotobira(strHinban) Then
                        'strShingumizu = "SS-7"
                    Else
                        'strShingumizu = "SS-5"
                    End If
                    
                Case Is <= 1800
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 2 * intMaisu
                    dblHashira = dblDH - 114
                    
                    If IsKotobira(strHinban) Then
                        'strShingumizu = "SS-7"
                    Else
                        'strShingumizu = "SS-5"
                    End If
                    
            End Select
            
            If IsKotobira(strHinban) Then
                
                dblSan = dblDW - 61
                intGakuYokoH1 = 4
                dblGakuYoko1 = (dblDW / 2) - 140.5
                dblSode1 = 0
                intSode1H = 0
                '20180205 K.Asayama ADD
                dblGakuYokoLVL30 = 0
                intGakuYokoLVL30 = 0
                dblGakutate3 = 0
                intGakutateH3 = 0
                
            '20180205 K.Asayama Change
            ElseIf dblDH > 2589 Then
                dblCupShitaji = 60
                intCupShitajiH = 8 * intMaisu
                
            ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
    
                dblDaboShitaji = 150
                intDaboShitajiH = 2 * intMaisu
                If dblDH <= 2529 Then
                    dblCupShitaji = 35
                    intCupShitajiH = 2 * intMaisu
                End If
     
            End If
            
        End If
        
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
        
            intSode1H = 3 * intMaisu
            dblSode2 = 110
            intSode2H = 2 * intMaisu
            intGakuYokoH1 = intGakuYokoH1 / 2
            dblGakuYoko2 = (dblDW / 2) - 258.5
            intGakuYokoH2 = intGakuYokoH1
            
        End If
        '******************************************************
        
        '20151211 K.Asayama Change 1601仕様追加
        If IsHidden_Hinge(strHinban) Then
            If IsKotobira(strHinban) Then
                dblGakuYoko1 = (dblDW / 2) - 140.5
                intGakuYokoH1 = 2
                dblGakuYoko2 = (dblDW / 2) - 160.5
                intGakuYokoH2 = 2
                intGakutateH1 = 1
                Select Case dblDH
                    '20180205 K.Asayama ADD
                    Case Is > 2589
                        dblGakutate1 = dblDH - 114
                    Case 2530 To 2589
                        dblGakutate1 = dblDH - 174
                    Case Is <= 2529
                        dblGakutate1 = dblDH - 114
                End Select
            Else
                '20180205 K.Asayama ADD
                If dblDH <= 2589 Then
                    dblGakutate1 = 210
                    intGakutateH1 = 2
                End If
            End If
        End If
        '20151211 K.Asayama Change End
'   *BG1*********************************************************

    ElseIf strHinban Like "B?C??*-####S*-*" Then
    
        If IsPALIOBlack(strHinban) Then
        
        Else
            dblShinAtsu = 30.2
            dblSan = dblDW - 61
            dblGakuYoko1 = dblDW - 492.5
            intHashiraH = 5 * intMaisu
            dblSode1 = 91.5
            dblSode2 = 60
            '20161121 K.Asayama Change

            'intSode1H = 2 * intMaisu
            'intSode2H = 6 * intMaisu
            If IsHikido(strHinban) Then
                intSode1H = 4 * intMaisu
            Else
                intSode1H = 2 * intMaisu
            End If
            
            intSode2H = intFncSode2Honsu_Group1(strHinban, intMaisu)
        
            '20161121 K.Asayama Change END
            
            Select Case dblDH
                
                '20180205 K.Asayama ADD
                Case 2589.5 To 2689
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 2 * intMaisu
    
                    dblHashira = dblDH - 114
                    dblGakuYokoLVL30 = dblDW - 302
                    intGakuYokoLVL30 = 2 * intMaisu
                    dblGakutate3 = 150
                    
                    If IsHirakido(strHinban) Or IsOyatobira(strHinban) Then
                        If IsHidden_Hinge(strHinban) Then
                            intGakutateH3 = 9 * intMaisu
                        Else
                            intGakutateH3 = 8 * intMaisu
                        End If
                        
                    ElseIf strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                        intGakutateH3 = 5 * intMaisu
                    Else
                        intGakutateH3 = 9 * intMaisu
                    End If
                
                Case 2530 To 2589
                    intSanH = 6 * intMaisu
                    intGakuYokoH1 = 2 * intMaisu
                    dblHashira = dblDH - 174
                    
                    If IsKotobira(strHinban) Then
                        'strShingumizu = "SS-8"
                    Else
                        'strShingumizu = "SS-4"
                    End If
                    
                Case 1801 To 2529
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 2 * intMaisu
                    dblHashira = dblDH - 114
                    
                    If IsKotobira(strHinban) Then
                        'strShingumizu = "SS-7"
                    Else
                        'strShingumizu = "SS-3"
                    End If
                Case Is <= 1800
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 1 * intMaisu
                    dblHashira = dblDH - 114
                    
                    If IsKotobira(strHinban) Then
                        'strShingumizu = "SS-7"
                    Else
                        'strShingumizu = "SS-3"
                    End If
                    
            End Select
            
            If IsKotobira(strHinban) Then
            
                intGakuYokoH1 = 4
                dblGakuYoko1 = (dblDW / 2) - 140.5
                dblSode1 = 0
                intSode1H = 0
                dblSode2 = 0
                intSode2H = 0
                
                '20180205 K.Asayama ADD
                dblGakuYokoLVL30 = 0
                intGakuYokoLVL30 = 0
                dblGakutate3 = 0
                intGakutateH3 = 0
            
            '20180205 K.Asayama Change
            ElseIf dblDH > 2589 Then
                dblCupShitaji = 60
                If IsHirakido(strHinban) Or IsOyatobira(strHinban) Then
                    intCupShitajiH = 4 * intMaisu
                Else
                    intCupShitajiH = 5 * intMaisu
                End If
                
            ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then

                dblDaboShitaji = 150
                intDaboShitajiH = 2 * intMaisu
                If dblDH <= 2529 Then
                    dblCupShitaji = 35
                    intCupShitajiH = 2 * intMaisu
                End If
     
            End If
        
        End If
        
        '20151211 K.Asayama Change 1601仕様追加
        If IsHidden_Hinge(strHinban) Then
            If IsKotobira(strHinban) Then
                dblGakuYoko1 = (dblDW / 2) - 140.5
                intGakuYokoH1 = 2
                dblGakuYoko2 = (dblDW / 2) - 160.5
                intGakuYokoH2 = 2
                intGakutateH1 = 1
                Select Case dblDH
                    '20180205 K.Asayama ADD
                    Case Is > 2589
                        dblGakutate1 = dblDH - 114
                    Case 2530 To 2589
                        dblGakutate1 = dblDH - 174
                    Case Is <= 2529
                        dblGakutate1 = dblDH - 114
                End Select
            Else
                '20180205 K.Asayama ADD
                If dblDH <= 2589 Then
                    dblGakutate1 = 210
                    intGakutateH1 = 2
                End If
            End If
        End If
        '20151211 K.Asayama Change End

'   *FA1*********************************************************

    ElseIf strHinban Like "A?C??*-####SC*-*" Then
        'シナ色
        If IsSINAColor(strHinban) Then
        
            dblShinAtsu = 30.2
            dblShinAtsu_N = 15
            dblSan = dblDW - 61
            dblGakuYoko1 = dblDW - 522.5
            intHashiraH = 6 * intMaisu
            dblSode1 = 91.5
            'intSode1H = 2 * intMaisu
            dblSode2 = 60
            'intSode2H = 6 * intMaisu
            
            '20160825 K.Asayama Change
            '20161121 K.Asayama Change
            If IsHikido(strHinban) Then
                intSode1H = 4 * intMaisu
                'intSode2H = 5 * intMaisu
            Else
                intSode1H = 2 * intMaisu
                'intSode2H = 6 * intMaisu
            End If
            
            intSode2H = intFncSode2Honsu_Group1(strHinban, intMaisu)
            '20161121 K.Asayama Change End
            '20160825 K.Asayama Change End
            
            Select Case dblDH
            
                Case 2530 To 2589
                    intSanH = 6 * intMaisu
                    intGakuYokoH1 = 2 * intMaisu
                    dblHashira = dblDH - 174
                    
                    'strShingumizu = "SS-34"
                    
                Case 1801 To 2529
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 2 * intMaisu
                    dblHashira = dblDH - 114
                    
                    'strShingumizu = "SS-33"
                    
                Case Is <= 1800
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 1 * intMaisu
                    dblHashira = dblDH - 114
                    
                    'strShingumizu = "SS-33"
                    
            End Select
            
            If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DU-####*" _
                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
                '20151211 K.Asayama Change 1601仕様追加
                dblDaboShitaji = 150
                intDaboShitajiH = 1
                '20151211 K.Asayama Change End

                If dblDH <= 2529 Then
                    dblCupShitaji = 35
                    intCupShitajiH = 2 * intMaisu
                End If
            End If
            
        'それ以外
        Else
            dblShinAtsu = 30.2
            dblShinAtsu_N = 15
            dblSan = dblDW - 64
            dblGakuYoko1 = dblDW - 524
            intHashiraH = 6 * intMaisu
            dblSode1 = 90
            dblSode2 = 60
            
            '20160825 K.Asayama Change
            '20161121 K.Asayama Change
            If IsHikido(strHinban) Then
                intSode1H = 4 * intMaisu
                'intSode2H = 5 * intMaisu
            Else
                intSode1H = 2 * intMaisu
                'intSode2H = 6 * intMaisu
            End If
            
            intSode2H = intFncSode2Honsu_Group1(strHinban, intMaisu)
            '20161121 K.Asayama Change END
        
            '20160825 K.Asayama Change End
            
            Select Case dblDH
            
                Case 2530 To 2589
                    intSanH = 6 * intMaisu
                    intGakuYokoH1 = 2 * intMaisu
                    dblHashira = dblDH - 174
                    
                    'strShingumizu = "SS-34"
                    
                Case 1801 To 2529
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 2 * intMaisu
                    dblHashira = dblDH - 114
                    
                    'strShingumizu = "SS-33"
                    
                Case Is <= 1800
                    intSanH = 4 * intMaisu
                    intGakuYokoH1 = 1 * intMaisu
                    dblHashira = dblDH - 114
                    
                    'strShingumizu = "SS-33"
                    
            End Select
            
            If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DU-####*" _
                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
                
                '20151211 K.Asayama Change 1601仕様追加
                dblDaboShitaji = 150
                intDaboShitajiH = 1
                '20151211 K.Asayama Change End
                
                If dblDH <= 2529 Then
                    dblCupShitaji = 35
                    intCupShitajiH = 2 * intMaisu
                End If
            End If
            
        End If

        '20151211 K.Asayama Change 1601仕様追加
        If IsHidden_Hinge(strHinban) Then
            dblGakutate1 = 210
            intGakutateH1 = 2
        End If
        
        dblSan_N = 72
        dblYokoSan_N = 72
        intHashiraShitaH_N = 2
        
        '20160308 改定により全面見直し(扉種類毎にDHの範囲が違う）
        
'        If IsStealth_SagariKabe(strHinban) Then
'
'            Select Case dblDH
'                Case 2132 To 2589
'                    intSanH_N = 4
'                    dblHashira_N = dblDH - 2068
'                    intHashiraH_N = 2
'                    dblHashiraShita_N = 1765
'                    intYokoSanH_N = 6
'
'                Case Is <= 2131
'                    intSanH_N = 2
'                    dblHashiraShita_N = dblDH - 309
'                    intYokoSanH_N = 5
'            End Select
'        Else
'            Select Case dblDH
'                Case 2125 To 2589
'                    intSanH_N = 4
'                    dblHashira_N = dblDH - 2068
'                    intHashiraH_N = 2
'
'                    ' 20151221 katsumata change start
'                    ' 条件文の誤りを修正
''                    If strHinban Like "*DC-####*#" Or "*DT-####*#" Or "*DU-####*#" Or "*KC-####*#" Or "*KT-####*#" Or "*KU-####*#" Then
'                    If strHinban Like "*DC-####*#" Or strHinban Like "*DT-####*#" Or strHinban Like "*DU-####*#" Or strHinban Like "*KC-####*#" Or strHinban Like "*KT-####*#" Or strHinban Like "*KU-####*#" Then
'                    ' 20151221 katsumata change end
'                        dblHashiraShita_N = 1637
'                    Else
'                        dblHashiraShita_N = 1765
'                    End If
'                    intYokoSanH_N = 6
'
'                Case Is <= 2124
'                    intSanH_N = 2
'                    ' 20151221 katsumata change start
'                    ' 条件文の誤りを修正
''                    If strHinban Like "*DC-####*#" Or "*DT-####*#" Or "*DU-####*#" Or "*KC-####*#" Or "*KT-####*#" Or "*KU-####*#" Then
'                    If strHinban Like "*DC-####*#" Or strHinban Like "*DT-####*#" Or strHinban Like "*DU-####*#" Or strHinban Like "*KC-####*#" Or strHinban Like "*KT-####*#" Or strHinban Like "*KU-####*#" Then
'                    ' 20151221 katsumata change end
'                        dblHashiraShita_N = dblDH - 438
'                    Else
'                        dblHashiraShita_N = dblDH - 309
'                    End If
'
'                    intYokoSanH_N = 5
'            End Select
'        End If
'        '20151211 K.Asayama Change End

        Dim dblDHRange As Double
        
        If IsHirakido(strHinban) Then
        
            If IsHidden_Hinge(strHinban) Then
                dblDHRange = 2132
            Else
                dblDHRange = 2125
            End If
            
            Select Case dblDH
                Case dblDHRange To 2589
                    intSanH_N = 4
                    dblHashira_N = dblDH - 2067
                    intHashiraH_N = 2
                    dblHashiraShita_N = 1766
                    intYokoSanH_N = 6

                Case Is <= dblDHRange - 1
                    intSanH_N = 2
                    dblHashiraShita_N = dblDH - 309
                    intYokoSanH_N = 5
            End Select
            
        ElseIf IsKabetsukeGuide(strHinban) Then
            dblDHRange = 2118
        Else
            dblDHRange = 2132
            
            Select Case dblDH
                Case dblDHRange To 2589
                    intSanH_N = 4
                    dblHashira_N = dblDH - 2067
                    intHashiraH_N = 2
                    dblHashiraShita_N = 1638
                    intYokoSanH_N = 6

                Case Is <= dblDHRange - 1
                    intSanH_N = 2
                    dblHashiraShita_N = dblDH - 437
                    intYokoSanH_N = 5
            End Select
            
        End If
'   *AG3*********************************************************

    ElseIf strHinban Like "F?C??*-####O*-*" Then
    
        dblShinAtsu = 30.2
        dblSan = dblDW - 290
        dblGakuYoko1 = (dblDW / 2) - 315
        intHashiraH = 5 * intMaisu
        
         '20160825 K.Asayama Change
        dblSode1 = 60
'        intSode1H = 2 * intMaisu
'        dblSode2 = 60
'        intSode2H = 3 * intMaisu
        
        '20161121 K.Asayama Change
'        If IsHikido(strHinban) Then
'            intSode1H = 4
'        Else
'            intSode1H = 5
'        End If
        
        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        
        '20160825 K.Asayama Change END
        
        '20151211 K.Asayama Change
        'dblDaboShitaji = 60
        dblDaboShitaji = 150
        '20151211 K.Asayama Change End
        intDaboShitajiH = 2
        
        'strShingumizu = "SS-32"
        
        Select Case dblDH
        
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 4 * intMaisu
                dblHashira = dblDH - 114
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                
        End Select
        
        If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DU-####*" _
            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
            
            dblCupShitaji = 35
            intCupShitajiH = 2 * intMaisu
        End If
           
        '20160822 K.Asayama ADD 1601仕様漏れ追加
        If IsHidden_Hinge(strHinban) Then
            dblGakutate1 = 210
            intGakutateH1 = 2
        End If
        '20160822 K.Asayama ADD END
        
'   *AG2*********************************************************

    ElseIf strHinban Like "F?C??*-####B*-*" Then
    
        dblShinAtsu = 30.2
        dblSan = dblDW - 240
        dblGakuYoko1 = (dblDW / 2) - 290
        intHashiraH = 5 * intMaisu
        
        '20160825 K.Asayama Change
        dblSode1 = 60
'        intSode1H = 2 * intMaisu
'        dblSode2 = 60
'        intSode2H = 3 * intMaisu
        
        '20161121 K.Asayama Change
        'If IsHikido(strHinban) Then
        '    intSode1H = 4
        'Else
        '    intSode1H = 5
        'End If
        
        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        '20160825 K.Asayama Change END
        
        'strShingumizu = "SS-5"
        
        Select Case dblDH
        
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 4 * intMaisu
                dblHashira = dblDH - 114
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                
        End Select
        
        If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DU-####*" _
            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
            
            dblCupShitaji = 35
            intCupShitajiH = 2 * intMaisu
        End If
        
        '20160822 K.Asayama ADD 1601仕様漏れ追加
        If IsHidden_Hinge(strHinban) Then
            dblGakutate1 = 210
            intGakutateH1 = 2
        End If
        '20160822 K.Asayama ADD END
        
'   *AG1*********************************************************

    ElseIf strHinban Like "F?C??*-####A*-*" Then
        
        dblShinAtsu = 30.2
        dblSan = dblDW - 200
        dblGakuYoko1 = (dblDW / 2) - 270
        intHashiraH = 5 * intMaisu
        
        '20160825 K.Asayama Change
        dblSode1 = 60
'        intSode1H = 2 * intMaisu
'        dblSode2 = 60
'        intSode2H = 3 * intMaisu
        
        
        '20161121 K.Asayama Change
        'If IsHikido(strHinban) Then
        '    intSode1H = 4
        'Else
        '    intSode1H = 5
        'End If
        
        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        '20160825 K.Asayama Change END
        
        'strShingumizu = "SS-5"
        
        Select Case dblDH
        
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 4 * intMaisu
                dblHashira = dblDH - 114
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                
        End Select
        
        If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DU-####*" _
            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
            
            dblCupShitaji = 35
            intCupShitajiH = 2 * intMaisu
        End If
        
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
        
            intSode1H = 3 * intMaisu
            dblSode2 = 110
            intSode2H = 2 * intMaisu
            intGakuYokoH1 = intGakuYokoH1 / 2
            dblGakuYoko2 = (dblDW / 2) - 310
            intGakuYokoH2 = intGakuYokoH1
            
        End If
        '******************************************************
        
        '20160822 K.Asayama ADD 1601仕様漏れ追加
        If IsHidden_Hinge(strHinban) Then
            dblGakutate1 = 210
            intGakutateH1 = 2
        End If
        '20160822 K.Asayama ADD END
        
'   *BF1*********************************************************

    ElseIf strHinban Like "B?C??*-####F*-*" Then
        
        If IsPALIOBlack(strHinban) Then
        
        Else
            dblShinAtsu = 30.2
            dblSan = dblDW + 5
            dblGakuYoko1 = dblDW - 245
            intGakuYokoH1 = 2 * intMaisu
            intHashiraH = 2 * intMaisu
            dblSode1 = 60
            
            '20161121 K.Asayama Change
'            If strHinban Like "*DH-####*" Then
'                intSode1H = 8
'            ElseIf strHinban Like "*DE-####*" Then
'                intSode1H = 10
'            ElseIf strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Then
'                intSode1H = 12
'            Else
'                intSode1H = 5
'            End If
            
            intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
            '20161121 K.Asayama Change END
        
            Select Case dblDH
                
                '20180205 K.Asayama ADD
                Case 2589.5 To 2689
                    intSanH = 4 * intMaisu
                    dblHashira = dblDH - 114
                    dblGakuYokoLVL30 = dblDW - 55
                    intGakuYokoLVL30 = 2 * intMaisu
                    dblGakutate3 = 150
                    If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                        intGakutateH3 = 4 * intMaisu
                    Else
                        intGakutateH3 = 8 * intMaisu
                    End If
                
                Case 2530 To 2589
                    intSanH = 6 * intMaisu
                    dblHashira = dblDH - 174
                    
                    'strShingumizu = "SS-2"
                    
                Case 1801 To 2529
                    intSanH = 4 * intMaisu
                    dblHashira = dblDH - 114
                    
                    'strShingumizu = "SS-1"
                    
                Case Is <= 1800
                    intSanH = 4 * intMaisu
                    dblHashira = dblDH - 114
                    
                    'strShingumizu = "SS-1"
                    
            End Select
            
            '20170105 K.Asayama Change
'            If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" Or strHinban Like "*DH-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" _
'                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
             
             '20180205 K.Asayama Change
            If dblDH > 2589 Then
                dblCupShitaji = 60
                intCupShitajiH = 6 * intMaisu
            
            ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" Or strHinban Like "*DH-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" _
                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Then
            '20170105 K.Asayama Change END
                
                If dblDH <= 2529 Then
                    dblCupShitaji = 35
                    intCupShitajiH = 2 * intMaisu
                End If
            End If
            
        End If
        
        'AUハンドル例外処理************************************
        If (IsHirakido(strHinban) Or IsOyatobira(strHinban)) And fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "AU" Or fncstrHandle_Name(CStr(Nz(varHandle, "")), CStr(Nz(varSpec, ""))) = "U" Then
            intSode1H = 3 * intMaisu
            dblSode2 = 110
            intSode2H = 2 * intMaisu
            dblGakuYoko1 = dblDW - 295
        End If
        '******************************************************
        
        '20151211 K.Asayama Change 1601仕様追加

        '20180205 K.Asayama Change
        'If IsHidden_Hinge(strHinban)
        If IsHidden_Hinge(strHinban) And dblDH <= 2589 Then
            dblGakutate1 = 210
            intGakutateH1 = 2
        End If
        '20151211 K.Asayama Change End

'   *XG3*********************************************************
    '20151211 K.Asayama Change 1601仕様追加
    'ElseIf strHinban Like "*F?B??*-####G*-*" Then
    ElseIf strHinban Like "X?B??*-####G*-*" Then
    '20151211 K.Asayama Change End

        dblShinAtsu = 30.2
        intHashiraH = 5 * intMaisu
        dblSode1 = 60
        dblDaboShitaji = 150
        intDaboShitajiH = 4
        
        '20160825 K.Asayama ADD
        '20161121 K.Asayama Change
'        If IsHikido(strHinban) Then
'            intSode1H = 4
'        Else
'            intSode1H = 5
'        End If

        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        
        '20160825 K.Asayama ADD END
        
        Select Case dblDH
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                intGakuYokoH1 = 4 * intMaisu
                dblHashira = dblDH - 174
                
                'strShingumizu = "SS-10"
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 4 * intMaisu
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-9"
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-9"
                
        End Select
        
        If bolFncSan_Koteichi(dblDW, dblDH, strHinban, dblSan, dblGakuYoko1, strShingumizu) Then
                
        Else
            'エラー
        End If
        
        If strHinban Like "*DC-####*" Or strHinban Like "*KC-####*" Then
        
            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If
        End If

'   *CF4/EF4(旧VF1)**********************************************
'   '20150902 K.Asayama ADD
'   '20180205 K.Asayama 柱LVL45追加対応

    ElseIf strHinban Like "F?V??*-####P*-*" Then
    
        dblShinAtsu = 30.2
        dblSan = dblDW - 166.5
        'dblGakuYoko1 = dblDW - 446.5
        'intHashiraH = 3 * intMaisu
        intHashiraH = 1 * intMaisu
        intHashiraH2 = 1 * intMaisu
        dblSode1 = 60
        'intSode1H = 6 * intMaisu
        
        '20160825 K.Asayama Change
        '20161121 K.Asayama Change
'        If strHinban Like "*DN-####*-*" Then
'            intSode1H = 3 * intMaisu
'        Else
'            intSode1H = 5 * intMaisu
'        End If

        intSode1H = intFncSode2Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        
        '20170105 K.Asayama Change
'        If strHinban Like "*DN-####*" Then
        If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
        '20170105 K.Asayama Change END

            'dblGakuYoko1 = dblDW - 386.5
            dblGakuYoko1 = dblDW - 371.5
        Else
            'dblGakuYoko1 = dblDW - 446.5
            dblGakuYoko1 = dblDW - 431.5
        End If
        '20160825 K.Asayama Change END
        
        Select Case dblDH
            
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                dblGakuYokoLVL30 = dblDW - 241.5
                intGakuYokoLVL30 = 2 * intMaisu
                dblGakutate3 = 150
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                    intGakutateH3 = 4 * intMaisu
                Else
                    intGakutateH3 = 8 * intMaisu
                End If
                    
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 174
                
                'strShingumizu = "SS-45"
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-44"
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 1 * intMaisu
                dblHashira = dblDH - 114
                
                'strShingumizu = "SS-44"
                
        End Select
        
        dblHashira2 = dblHashira
        
        '20180205 K.Asayama Change
        If dblDH > 2589 Then
            dblCupShitaji = 60
            intCupShitajiH = 6 * intMaisu
                
        '20170105 K.Asayama Change
'        If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" Or strHinban Like "*DH-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" _
'                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" Or strHinban Like "*DH-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" _
                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Then
        '20170105 K.Asayama Change END
        
            '20160524 K.Asayama Change 2530~2589カップ下地無し
            'dblCupShitaji = 35
            'intCupShitajiH = 2 * intMaisu
            
            If dblDH < 2530 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If
            '20160524 K.Asayama Change END
        End If
        
'   *CG4/EG4(旧VG4)**********************************************
'   '20151211 K.Asayama ADD
    '20180205 K.Asayama 柱LVL45追加対応
    
    ElseIf strHinban Like "F?V??*-####V*-*" Then
    
        dblShinAtsu = 30.2
        dblSan = dblDW - 166.5
        'dblGakuYoko1 = dblDW - 446.5
        'intHashiraH = 3 * intMaisu
        intHashiraH = 1 * intMaisu
        intHashiraH2 = 1 * intMaisu
        dblSode1 = 60
        'intSode1H = 6 * intMaisu
        
        '20160825 K.Asayama Change
        '20161121 K.Asayama Change
'        If strHinban Like "*DN-####*-*" Then
'            intSode1H = 3 * intMaisu
'        Else
'            intSode1H = 5 * intMaisu
'        End If

        intSode1H = intFncSode2Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END

        '20170105 K.Asayama Change
'        If strHinban Like "*DN-####*" Then
        If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
        '20170105 K.Asayama Change END

            'dblGakuYoko1 = dblDW - 386.5
            dblGakuYoko1 = dblDW - 371.5
        Else
            'dblGakuYoko1 = dblDW - 446.5
            dblGakuYoko1 = dblDW - 431.5
        End If
        '20160825 K.Asayama Change END
        
        Select Case dblDH
            
            '20180205 K.Asayama ADD
            Case 2589.5 To 2689
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                dblGakuYokoLVL30 = dblDW - 241.5
                intGakuYokoLVL30 = 2 * intMaisu
                dblGakutate3 = 150
                If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
                    intGakutateH3 = 4 * intMaisu
                Else
                    intGakutateH3 = 8 * intMaisu
                End If
                
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 174
                
                'strShingumizu = "PCS-15"
                
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                
                'strShingumizu = "PCS-14"
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 1 * intMaisu
                dblHashira = dblDH - 114
                
                'strShingumizu = "PCS-14"
                
        End Select
        
        dblHashira2 = dblHashira
        
        '20180205 K.Asayama Change
        If dblDH > 2589 Then
            dblCupShitaji = 60
            intCupShitajiH = 6 * intMaisu
            
        '20170105 K.Asayama Change
'        If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" Or strHinban Like "*DH-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" _
'                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
        ElseIf strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" Or strHinban Like "*DH-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" _
                Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Then
        '20170105 K.Asayama Change END
        
            '20160524 K.Asayama Change 2530~2589カップ下地無し
            'dblCupShitaji = 35
            'intCupShitajiH = 2 * intMaisu
            
            If dblDH < 2530 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If
            '20160524 K.Asayama Change END

        End If

'   *AF1*********************************************************
'   '20151211 K.Asayama ADD

    ElseIf strHinban Like "F?B??*-####A*-*" Then
        
        dblShinAtsu = 30.2
        dblSan = dblDW - 200
        intSanH = 4 * intMaisu
        dblGakuYoko1 = (dblDW / 2) - 270
        dblHashira = dblDH - 114
        intHashiraH = 5 * intMaisu
        
         '20160825 K.Asayama Change
        dblSode1 = 60
'        intSode1H = 2 * intMaisu
'        dblSode2 = 60
'        intSode2H = 3 * intMaisu
        
        '20161121 K.Asayama Change
'        If IsHikido(strHinban) Then
'            intSode1H = 4
'        Else
'            intSode1H = 5
'        End If
'
        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        '20160825 K.Asayama Change END
        
        dblDaboShitaji = 150
        intDaboShitajiH = 4 * intMaisu
        
        Select Case dblDH
        
            Case 1801 To 2529
                intGakuYokoH1 = 4 * intMaisu
                
            Case Is <= 1800
                intGakuYokoH1 = 2 * intMaisu
                
        End Select
        
        If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DU-####*" _
            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
            
            dblCupShitaji = 35
            intCupShitajiH = 2 * intMaisu
            'strShingumizu = "GCS-15"
            
        Else
            'strShingumizu = "GAS-10"
        End If
        
        If IsHidden_Hinge(strHinban) Then
            dblGakutate1 = 210
            intGakutateH1 = 2
        End If
        
'   *AF2*********************************************************
'   '20151211 K.Asayama ADD

    ElseIf strHinban Like "F?B??*-####B*-*" Then
        
        dblShinAtsu = 30.2
        dblSan = dblDW - 240
        intSanH = 4 * intMaisu
        dblGakuYoko1 = (dblDW / 2) - 290
        dblHashira = dblDH - 114
        intHashiraH = 5 * intMaisu
        
         '20160825 K.Asayama Change
        dblSode1 = 60
'        intSode1H = 2 * intMaisu
'        dblSode2 = 60
'        intSode2H = 3 * intMaisu
        
        '20161121 K.Asayama Change
'        If IsHikido(strHinban) Then
'            intSode1H = 4
'        Else
'            intSode1H = 5
'        End If
        
        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        '20160825 K.Asayama Change END
        
        dblDaboShitaji = 150
        intDaboShitajiH = 4 * intMaisu
        
        Select Case dblDH
        
            Case 1801 To 2529
                intGakuYokoH1 = 4 * intMaisu
                
            Case Is <= 1800
                intGakuYokoH1 = 2 * intMaisu
                
        End Select
        
        If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DU-####*" _
            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
            
            dblCupShitaji = 35
            intCupShitajiH = 2 * intMaisu
            'strShingumizu = "GCS-15"
            
        Else
            'strShingumizu = "GAS-10"
        End If
        
        If IsHidden_Hinge(strHinban) Then
            dblGakutate1 = 210
            intGakutateH1 = 2
        End If
        
'   *AF3*********************************************************
'   '20151211 K.Asayama ADD

    ElseIf strHinban Like "F?B??*-####O*-*" Then
        
        dblShinAtsu = 30.2
        dblSan = dblDW - 290
        intSanH = 4 * intMaisu
        dblGakuYoko1 = (dblDW / 2) - 315
        dblHashira = dblDH - 114
        intHashiraH = 5 * intMaisu
        
         '20160825 K.Asayama Change
        dblSode1 = 60
'        intSode1H = 2 * intMaisu
'        dblSode2 = 60
'        intSode2H = 3 * intMaisu
        
        '20161121 K.Asayama Change
'        If IsHikido(strHinban) Then
'            intSode1H = 4
'        Else
'            intSode1H = 5
'        End If
        
        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)
        '20161121 K.Asayama Change END
        '20160825 K.Asayama Change END
        
        dblDaboShitaji = 150
        intDaboShitajiH = 4 * intMaisu
        
        Select Case dblDH
        
            Case 1801 To 2529
                intGakuYokoH1 = 4 * intMaisu
                
            Case Is <= 1800
                intGakuYokoH1 = 2 * intMaisu
                
        End Select
        
        If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DU-####*" _
            Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" Then
            
            dblCupShitaji = 35
            intCupShitajiH = 2 * intMaisu
            'strShingumizu = "GCS-15"
            
        Else
            'strShingumizu = "GAS-10"
        End If
        
        If IsHidden_Hinge(strHinban) Then
            dblGakutate1 = 210
            intGakutateH1 = 2
        End If
        
    '20170517 K.Asayama ADD Terrace
'   *YF1*************************************************
    ElseIf strHinban Like "Y?C??*-####F*-*" Or strHinban Like "特 Y?C??*-####F*-*" Then
        
        dblShinAtsu = 30.2
        dblSan = dblDW + 2
        intHashiraH = 2 * intMaisu
        dblSode1 = 60
        
        
        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)

        If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
        
            dblGakuYoko1 = dblDW - 188
        Else
            dblGakuYoko1 = dblDW - 248
        End If
        
        Select Case dblDH
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 174

            Case 1801 To 2529
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 2 * intMaisu
                dblHashira = dblDH - 114
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                intGakuYokoH1 = 1 * intMaisu
                dblHashira = dblDH - 114

        End Select

        If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
        Or strHinban Like "*DH-####*" Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" _
        Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" _
        Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Then
        
            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If
        End If
        
        If IsHidden_Hinge(strHinban) Then
            dblGakutate1 = 210
            intGakutateH1 = 2
        End If
    
'   *YG6*************************************************
    ElseIf strHinban Like "Y?C??*-####T*-*" Or strHinban Like "特 Y?C??*-####T*-*" Then
        
        dblShinAtsu = 30.2
        dblSan = dblDW + 2
        intHashiraH = 2 * intMaisu
        dblSode1 = 60
        intGakuYokoH1 = 3 * intMaisu
        
        intSode1H = intFncSode1Honsu_Group1(strHinban, intMaisu)

        If strHinban Like "*DN-####*" Or strHinban Like "*VN-####*" Then
        
            dblGakuYoko1 = dblDW - 188
        Else
            dblGakuYoko1 = dblDW - 248
        End If
        
        Select Case dblDH
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                dblHashira = dblDH - 174

            Case 1801 To 2529
                intSanH = 4 * intMaisu
                dblHashira = dblDH - 114
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                dblHashira = dblDH - 114

        End Select

        If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
        Or strHinban Like "*DH-####*" Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" _
        Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" _
        Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Then
        
            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If
        End If
        
        dblGakutate1 = dblFncGakuTate1_YG6(strHinban, dblDW)
        If dblGakutate1 > 0 Then intGakutateH1 = 2 * intMaisu
        
        
        If IsHidden_Hinge(strHinban) Then
            dblGakutate2 = 210
            intGakutateH2 = 2
        End If
        
'   *YF5/YG5*********************************************
    ElseIf strHinban Like "Y?B??*-####W*-*" Or strHinban Like "特 Y?B??*-####W*-*" Then
        
        dblShinAtsu = 30.2
        dblSan = 280
        intHashiraH = 2 * intMaisu
        intHashiraH2 = 1 * intMaisu
        dblSode1 = 52.5
        dblSode2 = 52.5
        dblDaboShitaji = 150
        intDaboShitajiH = 4 * intMaisu
                    
        If IsHirakido(strHinban) Then
            intSode1H = 2 * intMaisu
            dblTegake = 52.5
        Else
            intSode1H = 3 * intMaisu
            dblTegake = 50
        End If
        
        If (IsEndWakunashi_Jou(strHinban) And Not (strHinban Like "*DN-####*-*" Or strHinban Like "*VN-####*-*")) _
            Or (strHinban Like "*DH-####*-*" Or strHinban Like "*DF-####*-*" Or strHinban Like "*DJ-####*-*" Or strHinban Like "*DQ-####*-*" Or strHinban Like "*VF-####*-*" Or strHinban Like "*VQ-####*-*") Then
            
            intSode2H = 4 * intMaisu
        Else
            intSode2H = 3 * intMaisu
        End If
                    
        
        Select Case dblDH
            Case 2530 To 2589
                intSanH = 6 * intMaisu
                dblHashira = dblDH - 174
                dblHashira2 = dblDH - 174
            Case 1801 To 2529
                intSanH = 4 * intMaisu
                dblHashira = dblDH - 114
                dblHashira2 = dblDH - 114
                
            Case Is <= 1800
                intSanH = 4 * intMaisu
                dblHashira = dblDH - 114
                dblHashira2 = dblDH - 114
        End Select
        

        If strHinban Like "*DC-####*" Or strHinban Like "*DT-####*" Or strHinban Like "*DP-####*" Or strHinban Like "*DU-####*" Or strHinban Like "*DE-####*" _
        Or strHinban Like "*DH-####*" Or strHinban Like "*KC-####*" Or strHinban Like "*KT-####*" Or strHinban Like "*KU-####*" _
        Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" _
        Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Then
        
            If dblDH <= 2529 Then
                dblCupShitaji = 35
                intCupShitajiH = 2 * intMaisu
            End If
        End If
        
        If IsHidden_Hinge(strHinban) Then
            dblGakutate1 = 210
            intGakutateH1 = 2
        End If
        
    '20170517 K.Asayama ADD END
    End If
'   *************************************************************
'   明り窓更新
'   （明り窓の時は袖に２本プラス）
'   *************************************************************
    'txt_明り窓 = "有" Or txt_明り窓 = "A" Or txt_明り窓 = "B"
    If (strAkarimado = "有" Or strAkarimado = "A" Or strAkarimado = "B") And (Nz(intSode1H, 0) > 0) Then
        intSode1H = intSode1H + (2 * intMaisu)
    End If
    
'   *************************************************************
'   芯組詳細図
'   20160308 K.Asayama ADD
'   *************************************************************
    
    '20180205 K.Asayama Change
    '2700から番号体系が変わっているので別モジュール化
    If dblDH > 2589 Then
        strShingumizu = fncstrShingumiShousai2700(strHinban, dblDH, varHandle)
    
    Else
        strShingumizu = fncstrShingumiShousai(strHinban, dblDH)
    
    End If
    
'   *************************************************************
'   OUTPUTへ出力（0の場合はNullを送る）
'
'   *************************************************************

'   20160308 K.Asayama Change dblFIVEorZERO関数追加
'   20160825 K.Asayama Change
'    引数をType型に変更 (以下コメントアウトなし)

    KidoriSunpo.out_dblShinAtsu = IIf(dblShinAtsu = 0, Null, dblShinAtsu)
    KidoriSunpo.out_dblsan = IIf(dblSan = 0, Null, dblFIVEorZERO(dblSan))
    KidoriSunpo.out_dblgakuyoko1 = IIf(dblGakuYoko1 = 0, Null, dblFIVEorZERO(dblGakuYoko1))
    KidoriSunpo.out_dblgakuYoko2 = IIf(dblGakuYoko2 = 0, Null, dblFIVEorZERO(dblGakuYoko2))
    KidoriSunpo.out_dblhashira = IIf(dblHashira = 0, Null, dblFIVEorZERO(dblHashira))
    '20170517 K.Asayama ADD
    KidoriSunpo.out_dblhashira2 = IIf(dblHashira2 = 0, Null, dblFIVEorZERO(dblHashira2))
    '20170517 K.Asayama ADD END
    KidoriSunpo.out_dblgakutate1 = IIf(dblGakutate1 = 0, Null, dblFIVEorZERO(dblGakutate1))
    KidoriSunpo.out_dblgakutate2 = IIf(dblGakutate2 = 0, Null, dblFIVEorZERO(dblGakutate2))
    KidoriSunpo.out_dblgakutate3 = IIf(dblGakutate3 = 0, Null, dblFIVEorZERO(dblGakutate3))
    KidoriSunpo.out_dbltegake = IIf(dblTegake = 0, Null, dblTegake)
    KidoriSunpo.out_dbltegakeShurui = IIf(dblTegakeShurui = 0, Null, dblTegakeShurui)
    KidoriSunpo.out_dblsode1 = IIf(dblSode1 = 0, Null, dblSode1)
    KidoriSunpo.out_dblsode2 = IIf(dblSode2 = 0, Null, dblSode2)
    KidoriSunpo.out_dbldaboshitaji = IIf(dblDaboShitaji = 0, Null, dblDaboShitaji)
    KidoriSunpo.out_dblCupShitaji = IIf(dblCupShitaji = 0, Null, dblCupShitaji)
    '20180205 K.Asayama ADD
    KidoriSunpo.out_dblGakuyokoLVL30 = IIf(dblGakuYokoLVL30 = 0, Null, dblGakuYokoLVL30)
    KidoriSunpo.out_dblhashira2_N = IIf(dblhashira2_N = 0, Null, dblhashira2_N)
    
    KidoriSunpo.out_intsan = IIf(intSanH = 0, Null, intSanH)
    KidoriSunpo.out_intgakuyoko1 = IIf(intGakuYokoH1 = 0, Null, intGakuYokoH1)
    KidoriSunpo.out_intgakuyoko2 = IIf(intGakuYokoH2 = 0, Null, intGakuYokoH2)
    KidoriSunpo.out_inthashira = IIf(intHashiraH = 0, Null, intHashiraH)
    '20170517 K.Asayama ADD
    KidoriSunpo.out_inthashira2 = IIf(intHashiraH2 = 0, Null, intHashiraH2)
    '20170517 K.Asayama ADD END
    KidoriSunpo.out_intgakutate1 = IIf(intGakutateH1 = 0, Null, intGakutateH1)
    KidoriSunpo.out_intgakutate2 = IIf(intGakutateH2 = 0, Null, intGakutateH2)
    KidoriSunpo.out_intgakutate3 = IIf(intGakutateH3 = 0, Null, intGakutateH3)
    KidoriSunpo.out_inttegake = IIf(intTegakeH = 0, Null, intTegakeH)
    KidoriSunpo.out_intsode1 = IIf(intSode1H = 0, Null, intSode1H)
    KidoriSunpo.out_intsode2 = IIf(intSode2H = 0, Null, intSode2H)
    KidoriSunpo.out_intdaboshitaji = IIf(intDaboShitajiH = 0, Null, intDaboShitajiH)
    KidoriSunpo.out_intcupshitaji = IIf(intCupShitajiH = 0, Null, intCupShitajiH)
    
    '20180205 K.Asayama ADD
    KidoriSunpo.out_intgakuyokoLVL30 = IIf(intGakuYokoLVL30 = 0, Null, intGakuYokoLVL30)
    
    KidoriSunpo.out_dblShinAtsu_N = IIf(dblShinAtsu_N = 0, Null, dblShinAtsu_N)
    KidoriSunpo.out_dblsan_N = IIf(dblSan_N = 0, Null, dblFIVEorZERO(dblSan_N))
    KidoriSunpo.out_dblgakuyoko1_N = IIf(dblGakuYoko1_N = 0, Null, dblFIVEorZERO(dblGakuYoko1_N))
    KidoriSunpo.out_dblhashira_N = IIf(dblHashira_N = 0, Null, dblFIVEorZERO(dblHashira_N))
    '20180205 K.Asayama ADD
    KidoriSunpo.out_dblsanH2_N = IIf(dblsanH2_N = 0, Null, dblsanH2_N)
    
    KidoriSunpo.out_intsanh_N = IIf(intSanH_N = 0, Null, intSanH_N)
    KidoriSunpo.out_intgakuyokoH1_N = IIf(intGakuYokoH1_N = 0, Null, intGakuYokoH1_N)
    KidoriSunpo.out_inthashiraH_N = IIf(intHashiraH_N = 0, Null, intHashiraH_N)
    '20180205 K.Asayama ADD
    KidoriSunpo.out_intsanh2_N = IIf(intsanh2_N = 0, Null, intsanh2_N)
    KidoriSunpo.out_inthashiraH2_N = IIf(inthashiraH2_N = 0, Null, inthashiraH2_N)
        
    '20160308 K.Asayama ADD
    KidoriSunpo.out_dblhashiraSt_N = IIf(dblHashiraShita_N = 0, Null, dblFIVEorZERO(dblHashiraShita_N))
    KidoriSunpo.out_dblYokosan_N = IIf(dblYokoSan_N = 0, Null, dblFIVEorZERO(dblYokoSan_N))
    
    KidoriSunpo.out_inthashiraStH_N = IIf(intHashiraShitaH_N = 0, Null, intHashiraShitaH_N)
    KidoriSunpo.out_intYokosanh_N = IIf(intYokoSanH_N = 0, Null, intYokoSanH_N)
    '20160308 K.Asayama ADD
    
    KidoriSunpo.out_strShingumizu = IIf(strShingumizu = "", Null, strShingumizu)
    '20160825 K.Asayama Change END

End Function

Public Function bolFncSan_Koteichi(ByVal in_dblDW As Double, ByVal in_dblDH As Double, ByVal in_strHinban As String, ByRef out_dblsan As Double, ByRef out_dblGakuYoko As Double, ByRef out_strShingumizu As String) As Boolean
'   *************************************************************
'   上下桟固定値（ガラス、ミラー扉の際の固定値）
'   'ADD by Asayama 20150917

'   戻り値:Boolean
'       →True              照合OK　数値戻し
'       →True              照合NG　数値なし
'
'    Input項目
'       in_dblDW            DW
'       in_dblDH            DH
'       in_strHinban        品番
'
'    Output項目
'      寸法
'       out_dblsan          上下桟
'       out_dblgakuyoko     額横
'       out_strShingumizu   芯組詳細図
'
' 20160308 K.Asayama 芯組詳細図はここでは送らない
'
'****************************************************************

    bolFncSan_Koteichi = True
    
    '20160308 K.Asayama Change
    
'    Select Case in_dblDW
'        Case 426 To 717.9
'            '20151211 K.Asayama Change 親扉が抜けていたので追加
'            If IsHirakido(in_strhinban) Or IsOyatobira(in_strhinban) Then
'                If in_dblDW >= 2530 Then
'                    out_strShingumizu = "KS-23"
'                Else
'                    out_strShingumizu = "KS-22"
'                End If
'            Else
'                If in_dblDW >= 2530 Then
'                    out_strShingumizu = "KS-25"
'                Else
'                    out_strShingumizu = "KS-24"
'                End If
'            End If
'
'        Case 718 To 1100
'
'            If in_dblDW >= 2530 Then
'                out_strShingumizu = "KS-19"
'            Else
'                out_strShingumizu = "KS-18"
'            End If
'
'        Case Else
'            out_strShingumizu = ""
'
'    End Select

    out_strShingumizu = ""
    
    '20160308 K.Asayama Change End
    
    Select Case in_dblDW
        Case 426 To 570.9
            out_dblsan = 346: out_dblGakuYoko = 3
            
        Case 571 To 618.9
            out_dblsan = 390: out_dblGakuYoko = 25
        
        Case 619 To 669.9
            out_dblsan = 422: out_dblGakuYoko = 41
        
        Case 670 To 717.9
            out_dblsan = 454: out_dblGakuYoko = 57
        
        Case 718 To 750.9
            out_dblsan = 488: out_dblGakuYoko = 74
                    
        Case 751 To 780.9
            out_dblsan = 502: out_dblGakuYoko = 81
        
        Case 781 To 819.9
            out_dblsan = 526: out_dblGakuYoko = 93
            
        Case 820 To 862.9
            out_dblsan = 552: out_dblGakuYoko = 106
                    
        Case 863 To 900.9
            out_dblsan = 576: out_dblGakuYoko = 118
                    
        Case 901 To 944.9
            out_dblsan = 610: out_dblGakuYoko = 135
                    
        Case 945 To 985.9
            out_dblsan = 638: out_dblGakuYoko = 149
        
        Case 986 To 1022.9
            out_dblsan = 662: out_dblGakuYoko = 161
                    
        Case 1023 To 1061.9
            out_dblsan = 688: out_dblGakuYoko = 174
            
        Case 1062 To 1100
            out_dblsan = 710: out_dblGakuYoko = 185
            
        Case Else
            out_dblsan = 0: out_dblGakuYoko = 0
            bolFncSan_Koteichi = False
    End Select

End Function

Public Function bolFncLVL30_Koteichi(ByVal in_dblDW As Double, ByVal in_dblDH As Double, ByVal in_strHinban As String, ByRef out_dblGakuyokoLVL30 As Double, ByRef out_dblCupshitajiLVL30 As Double) As Boolean
'   *************************************************************
'   額横、カップ下地 LVL30固定値（ガラス、ミラー扉の際の固定値）
'   'ADD by Asayama 20180201

'   戻り値:Boolean
'       →True              照合OK　数値戻し
'       →True              照合NG　数値なし
'
'    Input項目
'       in_dblDW            DW
'       in_dblDH            DH
'       in_strHinban        品番
'
'    Output項目
'      寸法
'       out_dblGakuyokoLVL30       上下桟
'       out_dblCupshitajiLVL30     カップ下地

'****************************************************************

    bolFncLVL30_Koteichi = True
    
    Select Case in_dblDW
        Case 426 To 570.9
            out_dblGakuyokoLVL30 = 98: out_dblCupshitajiLVL30 = 98
            
        Case 571 To 618.9
            out_dblGakuyokoLVL30 = 120: out_dblCupshitajiLVL30 = 60
        
        Case 619 To 669.9
            out_dblGakuyokoLVL30 = 136: out_dblCupshitajiLVL30 = 60
        
        Case 670 To 717.9
            out_dblGakuyokoLVL30 = 152: out_dblCupshitajiLVL30 = 60
        
        Case 718 To 750.9
            out_dblGakuyokoLVL30 = 169: out_dblCupshitajiLVL30 = 60
                    
        Case 751 To 780.9
            out_dblGakuyokoLVL30 = 176: out_dblCupshitajiLVL30 = 60
        
        Case 781 To 819.9
            out_dblGakuyokoLVL30 = 188: out_dblCupshitajiLVL30 = 60
            
        Case 820 To 862.9
            out_dblGakuyokoLVL30 = 201: out_dblCupshitajiLVL30 = 60
                    
        Case 863 To 900.9
            out_dblGakuyokoLVL30 = 213: out_dblCupshitajiLVL30 = 60
                    
        Case 901 To 944.9
            out_dblGakuyokoLVL30 = 230: out_dblCupshitajiLVL30 = 60
                    
        Case 945 To 985.9
            out_dblGakuyokoLVL30 = 244: out_dblCupshitajiLVL30 = 60
        
        Case 986 To 1022.9
            out_dblGakuyokoLVL30 = 256: out_dblCupshitajiLVL30 = 60
                    
        Case 1023 To 1061.9
            out_dblGakuyokoLVL30 = 269: out_dblCupshitajiLVL30 = 60
            
        Case 1062 To 1100
            out_dblGakuyokoLVL30 = 280: out_dblCupshitajiLVL30 = 60
            
        Case Else
            out_dblGakuyokoLVL30 = 0: out_dblCupshitajiLVL30 = 0
            bolFncLVL30_Koteichi = False
    End Select

End Function

Public Function dblfncTekake_Shurui(in_strHinban As String, in_strHandle As String, in_strSpec As String) As Double
'   *************************************************************
'   手掛け種類
'   'ADD by Asayama 20150917
'   'Change by Asayama 20160825 引数に品番を追加
'   戻り値:Double
'       →手掛け種類コードを戻す
'
'    Input項目
'       in_strHinban        建具品番
'       in_strHandle        ハンドル種類
'       in_strSpec          個別Spec
'   *************************************************************
'表示条件について(20170419記載）
'   1.引戸はすべて500（引き手レスを除く）
'   2.ヴェルチカは200
'   3.モンスターは0
'   4.蔵前製ハンドル、OLVARIは140
'   5.その他は100（特注はカワジュン製のみなので100でよい）
'   *************************************************************

'20170419 K.Asayama Change 引き手ではなく引戸は全部500
'    If fncbol_Handle_引手_長(in_strHandle, in_strSpec) Or fncbol_Handle_引手_短(in_strHandle, in_strSpec) Then
'
'        dblfncTekake_Shurui = 500
    '20180205 K.Asayama ADD
    If IsKotobira(in_strHinban) Then
        dblfncTekake_Shurui = 0
        Exit Function
    End If
    
    If IsHikido(in_strHinban) Then
        If fncstrHandle_Name(in_strHandle, in_strSpec) Like "-*" Then
        
            If IsVertica(in_strHinban) Then
                dblfncTekake_Shurui = 200
            Else
                dblfncTekake_Shurui = 0
            End If
        Else
            dblfncTekake_Shurui = 500
        End If
        
'20170419 K.Asayama Change END

    '20151211 K.Asayama Change 1601仕様追加
'    ElseIf fncstrHandle_Name(in_strHandle, in_strSpec) = "L" Or fncstrHandle_Name(in_strHandle, in_strSpec) = "M" _
'        Or fncstrHandle_Name(in_strHandle, in_strSpec) = "AL" Or fncstrHandle_Name(in_strHandle, in_strSpec) = "AM" _
'        Or fncstrHandle_Name(in_strHandle, in_strSpec) = "BY" Or fncstrHandle_Name(in_strHandle, in_strSpec) = "BZ" Then
    '20170419 K.Asayama キーノート追加
    '20170704 K.Asayama 1708仕様追加(DP,DQ)
    '20180205 K.Asayama 1801仕様追加(OLIVARI)
    '20180306 K.Asayama グランアートハンドルが漏れていたので追加
    ElseIf fncstrHandle_Name(in_strHandle, in_strSpec) = "L" Or fncstrHandle_Name(in_strHandle, in_strSpec) = "M" _
        Or fncstrHandle_Name(in_strHandle, in_strSpec) = "AL" Or fncstrHandle_Name(in_strHandle, in_strSpec) = "AM" _
        Or fncstrHandle_Name(in_strHandle, in_strSpec) = "CL" Or fncstrHandle_Name(in_strHandle, in_strSpec) = "CM" _
        Or fncstrHandle_Name(in_strHandle, in_strSpec) = "DL" Or fncstrHandle_Name(in_strHandle, in_strSpec) = "DM" _
        Or fncstrHandle_Name(in_strHandle, in_strSpec) = "BA" Or fncstrHandle_Name(in_strHandle, in_strSpec) = "BC" _
        Or fncstrHandle_Name(in_strHandle, in_strSpec) = "BD" Or fncstrHandle_Name(in_strHandle, in_strSpec) = "BE" _
        Or fncstrHandle_Name(in_strHandle, in_strSpec) = "BF" Or fncstrHandle_Name(in_strHandle, in_strSpec) = "BH" _
        Or fncstrHandle_Name(in_strHandle, in_strSpec) = "BI" Or fncstrHandle_Name(in_strHandle, in_strSpec) = "BJ" _
        Or fncstrHandle_Name(in_strHandle, in_strSpec) = "BL" Or fncstrHandle_Name(in_strHandle, in_strSpec) = "BM" _
        Or fncstrHandle_Name(in_strHandle, in_strSpec) = "BN" Or fncstrHandle_Name(in_strHandle, in_strSpec) = "BO" _
        Or fncstrHandle_Name(in_strHandle, in_strSpec) = "BP" Or fncstrHandle_Name(in_strHandle, in_strSpec) = "BQ" _
        Or fncstrHandle_Name(in_strHandle, in_strSpec) = "DP" Or fncstrHandle_Name(in_strHandle, in_strSpec) = "DQ" _
        Or fncstrHandle_Name(in_strHandle, in_strSpec) = "FC" Or fncstrHandle_Name(in_strHandle, in_strSpec) = "FD" _
        Or fncstrHandle_Name(in_strHandle, in_strSpec) = "FE" Or fncstrHandle_Name(in_strHandle, in_strSpec) = "FF" _
        Or fncstrHandle_Name(in_strHandle, in_strSpec) = "FG" Or fncstrHandle_Name(in_strHandle, in_strSpec) = "FH" _
        Or fncstrHandle_Name(in_strHandle, in_strSpec) = "BY" Or fncstrHandle_Name(in_strHandle, in_strSpec) = "BZ" _
        Or fncstrHandle_Name(in_strHandle, in_strSpec) = "BR" Then
    '20151211 K.Asayama Change End
    
        dblfncTekake_Shurui = 140

    Else
        dblfncTekake_Shurui = 100
        
    End If
    
End Function

Public Function fncstrTateguSearchKey(in_strHinban As String) As String
'   *************************************************************
'   T_建具関数ﾏｽﾀ検索用キー品番取得
'   'ADD by Asayama 20150917
'   戻り値:String
'       →検索用品番
'
'    Input項目
'       in_strHinban        建具品番
'
'   20160602 K.Asayama ADD モンスターで検索するときは色を追加
'   *************************************************************
    Dim strHinban As String
    
    On Error GoTo Err_fncstrTateguSearchKey

    '特注外し
    If in_strHinban Like "特 *" Then
        strHinban = Mid(in_strHinban, 3)
    Else
        strHinban = in_strHinban
    End If
    
    strHinban = left(strHinban, 1) & "?" & Mid(strHinban, 3, InStr(1, strHinban, "-") - 2) & "##" & Mid(strHinban, InStr(1, strHinban, "-") + 3, 3)
    
    '20160602 K.Asayama ADD
    'モンスターは色有(例:O?CDC-##24P*(PH))
    If IsMonster(in_strHinban) Then
        strHinban = strHinban & "*" & right(in_strHinban, 4)
    End If
    '20160602 K.Asayama ADD END
    
    fncstrTateguSearchKey = strHinban
    
    Exit Function
    
Err_fncstrTateguSearchKey:
    fncstrTateguSearchKey = ""
    
End Function

Public Function fncstrGetKihonzu(in_strKensakuHinban As String, in_spec As Variant) As String
'   *************************************************************
'   T_建具関数ﾏｽﾀ基本図番号取得

'   戻り値:String
'       →検索用品番
'
'    Input項目
'       in_strHinban        建具品番
'   *************************************************************
    Dim objREMOTEDB As New cls_BRAND_MASTER
    Dim strSQL As String
    
    fncstrGetKihonzu = ""
    
    On Error GoTo Err_fncstrGetKihonzu
    
    With objREMOTEDB
        If in_spec Like "*1007" Or in_spec Like "*1011" Or in_spec Like "*1111" Or in_spec Like "*1103" Or in_spec Like "*1105" Or in_spec Like "*1010" Or in_spec Like "*1009" Or right(in_spec, 4) >= "1304" Then
           strSQL = "select 基本図 from T_建具関数ﾏｽﾀ_1007仕様 where 建具品番 = '" & fncstrTateguSearchKey(in_strKensakuHinban) & "'"
        Else
           strSQL = "select 基本図 from T_建具関数ﾏｽﾀ where 建具品番 = '" & fncstrTateguSearchKey(in_strKensakuHinban) & "'"
        End If
        
        If .ExecSelect(strSQL) Then
            If Not .GetRS.EOF Then
                fncstrGetKihonzu = .GetRS![基本図]
            End If
        End If
    
    End With
    
    GoTo Exit_fncstrGetKihonzu
    
Err_fncstrGetKihonzu:

Exit_fncstrGetKihonzu:
    Set objREMOTEDB = Nothing

End Function

Public Function fncstrShingumiShousai2700(in_strHinban As String, dblDH As Double, in_varHandleName As Variant) As String
'   *************************************************************
'   芯組み詳細図関数(2700品番用）
'   'ADD by Asayama 20180201
'   戻り値:String
'       →詳細図番号        照合NGの場合は空欄
'
'    Input項目
'       in_strHinban        建具品番
'       dblDH               DH（作成時点では使用しない）
'       in_strHandleName    施錠名
'   *************************************************************

    Dim strShingumi As String
    Dim strHinban As String
    Dim strHandle As String
    
    
    fncstrShingumiShousai2700 = ""
    strShingumi = ""
    
    strHinban = Replace(in_strHinban, "特 ", "")
    strHandle = Nz(in_varHandleName, "")
    
'   *************************************************************
'   品番別データの挿入
'   （クローゼットと建具で品番が被っているのでクローゼットを先に処理）
'   *************************************************************

'   *MC1/ME1/MZ1*************************************************
'   *MS1*********************************************************
    If strHinban Like "*F?CME-####F*-*" Then
        strShingumi = "HPC-9"

'   *MC1/ME1/MZ1(ミラー)*****************************************
'   *MS1(ミラー)*************************************************
    ElseIf strHinban Like "*F?CME-####M*-*" Then
    
    '両側ミラー
        If strHinban Like "*-####MM*-*" Then

            strShingumi = "HPC-10"
        Else
            
            strShingumi = "HPC-9/10"
            
        End If
        
'   *CF1/EF1/ZF1*************************************************
'   *BF1*********************************************************
'   *RF1*********************************************************
'   *TF1*********************************************************
    ElseIf strHinban Like "F?C??*-####F*-*" Or strHinban Like "B?C??*-####F*-*" Or strHinban Like "R?C??*-####F*-*" Or strHinban Like "T?C??*-####F*-*" Then
    
        If IsHirakido(strHinban) Or IsOyatobira(strHinban) Then
            strShingumi = "HPA-1"
        Else
            If IsEndWakunashi(strHinban) And strHinban Like "*U-####*" And Not strHandle Like "*N" Then
                strShingumi = "HPC-2"
                
            ElseIf strHinban Like "*DH-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Then
                strShingumi = "HPC-2"
            Else
                strShingumi = "HPC-1"
            End If
        End If

   
'   *CG2/EG2/ZG2*************************************************
'   *RG2*********************************************************
'   *BG2*********************************************************
'   *TG2*********************************************************
    ElseIf strHinban Like "F?C??*-####C*-*" Or strHinban Like "R?C??*-####C*-*" Or strHinban Like "B?C??*-####C*-*" Or strHinban Like "T?C??*-####C*-*" Then
    
        If IsHirakido(strHinban) Or IsOyatobira(strHinban) Then
            strShingumi = "HGA-2"
            
        ElseIf IsKotobira(strHinban) Then
            strShingumi = "HGA-3"
        Else
            If IsEndWakunashi(strHinban) And strHinban Like "*U-####*" And Not strHandle Like "*N" Then
                strShingumi = "HGC-4"
                
            ElseIf strHinban Like "*DH-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Then
                strShingumi = "HGC-4"
            Else
                strShingumi = "HGC-3"
            End If
        End If
  
'   *CG1/EG1/ZG1*************************************************
'   *BG1*********************************************************
'   *TG1*********************************************************
'   *RG1*********************************************************
    ElseIf strHinban Like "F?C??*-####S*-*" Or strHinban Like "B?C??*-####S*-*" Or strHinban Like "T?C??*-####S*-*" Or strHinban Like "R?C??*-####S*-*" Then
      
        If IsHirakido(strHinban) Or IsOyatobira(strHinban) Then
              strShingumi = "HGA-1"
              
        ElseIf IsKotobira(strHinban) Then
            strShingumi = "HGA-3"
        Else
            If IsEndWakunashi(strHinban) And strHinban Like "*U-####*" And Not strHandle Like "*N" Then
                strShingumi = "HGC-2"
                
            ElseIf strHinban Like "*DH-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Then
                strShingumi = "HGC-2"
            Else
                strShingumi = "HGC-1"
            End If
        End If
        
 '   *CG3/EG3/ZG3*************************************************
    ElseIf strHinban Like "F?B??*-####G*-*" Then
    
        If IsHirakido(strHinban) Or IsOyatobira(strHinban) Then
            strShingumi = "HGA-4"
            
        ElseIf IsKotobira(strHinban) Then
            strShingumi = "HGA-5"
        Else
            If IsEndWakunashi(strHinban) And strHinban Like "*U-####*" And Not strHandle Like "*N" Then
                strShingumi = "HGC-6"
                
            ElseIf strHinban Like "*DH-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Then
                strShingumi = "HGC-6"
            Else
                strShingumi = "HGC-5"
            End If
        End If
     
   '   *新CG7/EG7/ZG7(1608以降)*************************************
    ElseIf strHinban Like "F?C??*-####M*-*" Then
    
        If IsHirakido(strHinban) Or IsOyatobira(strHinban) Then
            strShingumi = "HGA-6"
            
        Else
            If IsEndWakunashi(strHinban) And strHinban Like "*U-####*" And Not strHandle Like "*N" Then
                strShingumi = "HGC-8"
                
            ElseIf strHinban Like "*DH-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Then
                strShingumi = "HGC-8"
            Else
                strShingumi = "HGC-7"
            End If
        End If
     
    '   *KF1*********************************************************
    '   *KF7*********************************************************
    ElseIf strHinban Like "S?C??*-####Z*-*" Or strHinban Like "S?C??*-####F*-*" Then
    
        If IsHirakido(strHinban) Then
            strShingumi = "HPA-4"
        Else
            If IsEndWakunashi(strHinban) And strHinban Like "*U-####*" And Not strHandle Like "*N" Then
                strShingumi = "HPC-6"
                
            ElseIf strHinban Like "*DH-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Then
                strShingumi = "HPC-6"
            Else
                strShingumi = "HPC-5"
            End If
        End If
       
     '   *CF6/EF6/ZF6*************************************************
    ElseIf strHinban Like "F?D??*-####F*-*" Then
    
        If IsHirakido(strHinban) Then
            strShingumi = "HPA-3"

        Else
            If IsEndWakunashi(strHinban) And strHinban Like "*U-####*" And Not strHandle Like "*N" Then
                strShingumi = "HPC-4"
                
            ElseIf strHinban Like "*DH-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Then
                strShingumi = "HPC-4"
            Else
                strShingumi = "HPC-3"
            End If
        End If
        
     '   *CF4/EF4/CG4/EG4*********************************************
    ElseIf strHinban Like "F?V??*-####P*-*" Or strHinban Like "F?V??*-####V*-*" Then
    
        If IsEndWakunashi(strHinban) And strHinban Like "*U-####*" And Not strHandle Like "*N" Then
                strShingumi = "HPC-8"
                
            ElseIf strHinban Like "*DH-####*" Or strHinban Like "*DF-####*" Or strHinban Like "*DJ-####*" Or strHinban Like "*DQ-####*" Or strHinban Like "*VF-####*" Or strHinban Like "*VQ-####*" Then
                strShingumi = "HPC-8"
            Else
                strShingumi = "HPC-7"
            End If
    
    '   *GF1*********************************************************
    ElseIf strHinban Like "G?C??*-####F*-*" Then
        strShingumi = "HPA-5"
        
    '   *GG1*********************************************************
    '   *GG2*********************************************************
    ElseIf strHinban Like "*G?C??*-####S*-*" Or strHinban Like "*G?C??*-####C*-*" Then
        strShingumi = "HGA-9"
        
    End If
    
    fncstrShingumiShousai2700 = strShingumi
   
End Function

Public Function fncstrShingumiShousai(in_strHinban As String, dblDH As Double) As String

'   *************************************************************
'   芯組み詳細図関数
'   'ADD by Asayama 20160218
'   戻り値:String
'       →詳細図番号        照合NGの場合は空欄
'
'    Input項目
'       in_strhinban        建具品番
'       dblDH               DH
'   *************************************************************

    Dim strShingumi As String
    
    fncstrShingumiShousai = ""

'   *************************************************************
'   品番別データの挿入
'   （クローゼットと建具で品番が被っているのでクローゼットを先に処理）
'   *************************************************************
    
'   *MC1/ME1/MZ1*************************************************
'   *MS1*********************************************************
    If in_strHinban Like "*F?CME-####F*-*" Or in_strHinban Like "*T?CME-####F*-*" Then
        
        Select Case dblDH
            Case 2530 To 2589
                strShingumi = strShingumi & "PCS-11"
            Case Is <= 2529
                strShingumi = strShingumi & "PCS-10"
        End Select
        
'   *MC1/ME1/MZ1(ミラー)*****************************************
'   *MS1(ミラー)*************************************************
    ElseIf in_strHinban Like "*F?CME-####M*-*" Or in_strHinban Like "*T?CME-####M*-*" Then

        '両側ミラー
        If in_strHinban Like "*-####MM*-*" Then

            Select Case dblDH
                Case 2530 To 2589
                    strShingumi = strShingumi & "PCS-13"
                Case Is <= 2529
                    strShingumi = strShingumi & "PCS-12"
            End Select
        Else
            
            Select Case dblDH
                Case 2530 To 2589
                    strShingumi = strShingumi & "PCS-11/13"
                Case Is <= 2529
                    strShingumi = strShingumi & "PCS-10/12"
            End Select
        End If
   
'   *MP3*********************************************************
'20161108 K.Asayama Change 品番間違い修正
'    ElseIf in_strHinban Like "*F?CSA-####F*-*" Then
    ElseIf in_strHinban Like "*P?CSA-####F*-*" Then
'20161108 K.Asayama Change End
        Select Case dblDH
            Case 2530 To 2589
                strShingumi = strShingumi & "PAS-11"
            Case Is <= 2529
                strShingumi = strShingumi & "PAS-10"
        End Select
        
'   *CF1/EF1/ZF1*************************************************
'   *BF1*********************************************************
'   *RF1*********************************************************
'   *XF1*********************************************************
'   *TF1*********************************************************
    ElseIf in_strHinban Like "*F?C??*-####F*-*" Or in_strHinban Like "*B?C??*-####F*-*" Or in_strHinban Like "*R?C??*-####F*-*" Or in_strHinban Like "*X?C??*-####F*-*" Or in_strHinban Like "*T?C??*-####F*-*" Then
    
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "PAS"
        Else
            strShingumi = "PCS"
        End If
        
        Select Case dblDH
            Case 2530 To 2589
                strShingumi = strShingumi & "-2"
            Case Is <= 2529
                strShingumi = strShingumi & "-1"
        End Select
        
'   *CG2/EG2/ZG2*************************************************
'   *RG2*********************************************************
'   *BG2*********************************************************
'   *AG1*********************************************************
'   *AG2*********************************************************
'   *XG2*********************************************************
'   *TG2*********************************************************
    ElseIf in_strHinban Like "*F?C??*-####C*-*" Or in_strHinban Like "*R?C??*-####C*-*" Or in_strHinban Like "*B?C??*-####C*-*" Or in_strHinban Like "*F?C??*-####A*-*" Or in_strHinban Like "*F?C??*-####B*-*" Or in_strHinban Like "*X?C??*-####C*-*" Or in_strHinban Like "*T?C??*-####C*-*" Then
    
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "GAS"
        Else
            strShingumi = "GCS"
        End If
        
        Select Case dblDH
            Case 2530 To 2589
                If IsKotobira(in_strHinban) Then
                    strShingumi = strShingumi & "-6"
                Else
                    strShingumi = strShingumi & "-4"
                End If
            Case Is <= 2529
                If IsKotobira(in_strHinban) Then
                    strShingumi = strShingumi & "-5"
                Else
                    strShingumi = strShingumi & "-3"
                End If
        End Select
        
'   *CF6/EF6/ZF6*************************************************
    ElseIf in_strHinban Like "*F?D??*-####F*-*" Then
    
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "PAS"
        Else
            strShingumi = "PCS"
        End If
        
        Select Case dblDH
            Case 2530 To 2589
                strShingumi = strShingumi & "-4"
            Case Is <= 2529
                strShingumi = strShingumi & "-3"
        End Select
        
'   *CG8/EG8/ZG8*************************************************
    ElseIf in_strHinban Like "*F?C??*-####D*-*" Then
    
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "GAS"
        Else
            strShingumi = "GCS"
        End If
        
        Select Case dblDH
            Case Is <= 2529
                If IsKotobira(in_strHinban) Then
                    strShingumi = strShingumi & "-5"
                Else
                    '20160819 K.Asayama 誤り訂正
                    'strShingumi = strShingumi & "-1"
                    strShingumi = strShingumi & "-3"
                    '20160819 K.Asayama Change End
                End If
        End Select
        
'   *CG1/EG1/ZG1*************************************************
'   *BG1*********************************************************
'   *XG1*********************************************************
'   *TG1*********************************************************
'   *RG1*********************************************************
    ElseIf in_strHinban Like "*F?C??*-####S*-*" Or in_strHinban Like "*B?C??*-####S*-*" Or in_strHinban Like "*X?C??*-####S*-*" Or in_strHinban Like "*T?C??*-####S*-*" Or in_strHinban Like "*R?C??*-####S*-*" Then
    
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "GAS"
        Else
            strShingumi = "GCS"
        End If
        
        Select Case dblDH
            Case 2530 To 2589
                If IsKotobira(in_strHinban) Then
                    strShingumi = strShingumi & "-6"
                Else
'20160315 K.Asayama Change
'                    If strShingumi = "GAS" Then
                        strShingumi = strShingumi & "-2"
'                    Else
'                        strShingumi = strShingumi & "-4"
'                    End If
                End If
            Case Is <= 2529
                If IsKotobira(in_strHinban) Then
                    strShingumi = strShingumi & "-5"
                Else
'                    If strShingumi = "GAS" Then
                        strShingumi = strShingumi & "-1"
'                    Else
'                        strShingumi = strShingumi & "-3"
'                    End If
                End If
        End Select
        
'   *KF1*********************************************************
'   *KF7*********************************************************
    ElseIf in_strHinban Like "*S?C??*-####Z*-*" Or in_strHinban Like "*S?C??*-####F*-*" Then
    
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "PAS"
        Else
            strShingumi = "PCS"
        End If
        
        Select Case dblDH
            Case 2530 To 2589
                strShingumi = strShingumi & "-9"
            Case Is <= 2529
                strShingumi = strShingumi & "-8"
        End Select
      
'   *SG1*********************************************************
    ElseIf in_strHinban Like "*F?S??*-####S*-*" Then
    
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "GAS"
        Else
            strShingumi = "GCS"
        End If
        
        Select Case dblDH
            Case 2530 To 2589
                If IsKotobira(in_strHinban) Then
                    strShingumi = strShingumi & "-20"
                Else
                    If strShingumi = "GAS" Then
                        strShingumi = strShingumi & "-16"
                    Else
                        strShingumi = strShingumi & "-20"
                    End If
                End If
            Case Is <= 2529
                If IsKotobira(in_strHinban) Then
                    strShingumi = strShingumi & "-19"
                Else
                    If strShingumi = "GAS" Then
                        strShingumi = strShingumi & "-15"
                    Else
                        strShingumi = strShingumi & "-19"
                    End If
                End If
        End Select
        
'   *SG2*********************************************************
    ElseIf in_strHinban Like "*F?S??*-####C*-*" Then
               
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "GAS"
        Else
            strShingumi = "GCS"
        End If
        
        Select Case dblDH
            Case 2530 To 2589
                If IsKotobira(in_strHinban) Then
                    strShingumi = strShingumi & "-20"
                Else
                    If strShingumi = "GAS" Then
                        strShingumi = strShingumi & "-18"
                    Else
                        strShingumi = strShingumi & "-22"
                    End If
                End If
            Case Is <= 2529
                If IsKotobira(in_strHinban) Then
                    strShingumi = strShingumi & "-19"
                Else
                    If strShingumi = "GAS" Then
                        strShingumi = strShingumi & "-17"
                    Else
                        strShingumi = strShingumi & "-21"
                    End If
                End If
        End Select

'   *OF1*********************************************************
    '20170207 K.Asayama Change
'    ElseIf in_strHinban Like "*O?C??*-####P*-*" Then
    ElseIf in_strHinban Like "*O?C??*-####P*-*" Or in_strHinban Like "*O?C??*-####N*-*" Then
    '20170207 K.Asayama Change END
    
        Select Case dblDH
            Case 2530 To 2589
                strShingumi = strShingumi & "PCS-18"
            Case Is <= 2529
                strShingumi = strShingumi & "PCS-5"
        End Select
        
'   *SF1*********************************************************
    ElseIf in_strHinban Like "*F?S??*-####F*-*" Then
        
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "PAS"
        Else
            strShingumi = "PCS"
        End If
        
        Select Case dblDH
            Case 2530 To 2589
                If strShingumi = "PAS" Then
                    strShingumi = strShingumi & "-14"
                Else
                    strShingumi = strShingumi & "-17"
                End If
            Case Is <= 2529
                If strShingumi = "PAS" Then
                    strShingumi = strShingumi & "-13"
                Else
                    strShingumi = strShingumi & "-16"
                End If
        End Select
        
'   *PF6*********************************************************
    ElseIf in_strHinban Like "*P?D??*-####F*-*" Then
    
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "PAS"
        Else
            strShingumi = "PCS"
        End If
        
        Select Case dblDH
            Case 2530 To 2589
                strShingumi = strShingumi & "-7"
            Case Is <= 2529
                strShingumi = strShingumi & "-7"
        End Select
                
'   *PG2*********************************************************
    ElseIf in_strHinban Like "*P?C??*-####C*-*" Then
        
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "GAS"
        Else
            strShingumi = "GCS"
        End If
        
        Select Case dblDH
            Case Is <= 2529
                If IsKotobira(in_strHinban) Then
                    strShingumi = strShingumi & "-13"
                Else
                    If strShingumi = "GAS" Then
                        strShingumi = strShingumi & "-12"
                    Else
                        strShingumi = strShingumi & "-8"
                    End If
                End If
        End Select
        
'   *GG1*********************************************************
'   *GG2*********************************************************
    ElseIf in_strHinban Like "*G?C??*-####S*-*" Or in_strHinban Like "*G?C??*-####C*-*" Then

        strShingumi = strShingumi & "GAS-20"
        
'   *PG1*********************************************************
    ElseIf in_strHinban Like "*P?C??*-####S*-*" Then
    
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "GAS"
        Else
            strShingumi = "GCS"
        End If
        
        Select Case dblDH
            Case Is <= 2529
                If IsKotobira(in_strHinban) Then
                    strShingumi = strShingumi & "-13"
                Else
                    If strShingumi = "GAS" Then
                        strShingumi = strShingumi & "-11"
                    Else
                        strShingumi = strShingumi & "-7"
                    End If
                End If
        End Select

'   *GF1*********************************************************
    ElseIf in_strHinban Like "*G?C??*-####F*-*" Then

        strShingumi = strShingumi & "PAS-12"

    
'   *PF1*********************************************************
    ElseIf in_strHinban Like "*P?C??*-####F*-*" Then
        
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "PAS"
        Else
            strShingumi = "PCS"
        End If
        
        Select Case dblDH
            Case Is <= 2529
                If strShingumi = "PAS" Then
                    strShingumi = strShingumi & "-5"
                Else
                    strShingumi = strShingumi & "-6"
                End If
        End Select
        
'   *FA2*********************************************************
    ElseIf in_strHinban Like "*A?C??*-####SL*-*" Then
    
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "GAS"
        Else
            strShingumi = "GCS"
        End If
        
        Select Case dblDH
            Case 2530 To 2589
                    strShingumi = strShingumi & "-19"
                    
            Case Is <= 2529
                    strShingumi = strShingumi & "-18"
                    
        End Select
                
  '   *新CG7/EG7/ZG7(1608以降)*************************************
  '20160923 K.Asayama ADD
    ElseIf in_strHinban Like "*F?C??*-####M*-*" Then
    
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "GAS"
        Else
            strShingumi = "GCS"
        End If
        
        Select Case dblDH
            Case 2530 To 2589
                If strShingumi = "GAS" Then
                    strShingumi = strShingumi & "-28"
                    '隠し丁番
                    If IsHidden_Hinge(in_strHinban) Then
                        strShingumi = strShingumi & "B"
                    Else
                        strShingumi = strShingumi & "A"
                    End If
                Else
                    strShingumi = strShingumi & "-25"
                End If
            Case Is <= 2529
                    If strShingumi = "GAS" Then
                        strShingumi = strShingumi & "-27"
                        '隠し丁番
                        If IsHidden_Hinge(in_strHinban) Then
                            strShingumi = strShingumi & "B"
                        Else
                            strShingumi = strShingumi & "A"
                        End If
                    Else
                        strShingumi = strShingumi & "-24"
                    End If
        End Select
    '20160923 K.Asayama ADD END
    
 '   *CG3/EG3/ZG3*************************************************
 '   *CG7/EG7/ZG7*************************************************
    ElseIf in_strHinban Like "*F?B??*-####G*-*" Or in_strHinban Like "*F?B??*-####M*-*" Then
    
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "GAS"
        Else
            strShingumi = "GCS"
        End If
        
        Select Case dblDH
            Case 2530 To 2589
                If IsKotobira(in_strHinban) Then
                    strShingumi = strShingumi & "-10"
                Else
                    If strShingumi = "GAS" Then
                        strShingumi = strShingumi & "-8"
                        '隠し丁番
                        If IsHidden_Hinge(in_strHinban) Then
                            strShingumi = strShingumi & "B"
                        Else
                            strShingumi = strShingumi & "A"
                        End If
                    Else
                        strShingumi = strShingumi & "-6"
                    End If
                End If
            Case Is <= 2529
                If IsKotobira(in_strHinban) Then
                    strShingumi = strShingumi & "-9"
                Else
                    If strShingumi = "GAS" Then
                        strShingumi = strShingumi & "-7"
                        '隠し丁番
                        If IsHidden_Hinge(in_strHinban) Then
                            strShingumi = strShingumi & "B"
                        Else
                            strShingumi = strShingumi & "A"
                        End If
                    Else
                        strShingumi = strShingumi & "-5"
                    End If
                End If
        End Select

'   *FA1*********************************************************
    ElseIf in_strHinban Like "*A?C??*-####SC*-*" Then
        
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "GAS"
        Else
            strShingumi = "GCS"
        End If
        
        Select Case dblDH
            Case 2530 To 2589
                strShingumi = strShingumi & "-17"
 
            Case Is <= 2529
                strShingumi = strShingumi & "-16"

        End Select
        
'   *AG3*********************************************************
    ElseIf in_strHinban Like "*F?C??*-####O*-*" Then
    
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "GAS"
        Else
            strShingumi = "GCS"
        End If
        
        Select Case dblDH
            Case Is <= 2529
                strShingumi = strShingumi & "-14"
        End Select
        
'   *XG3*********************************************************
    ElseIf in_strHinban Like "*X?B??*-####G*-*" Then

        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "GAS"
        Else
            strShingumi = "GCS"
        End If
        
        Select Case dblDH
            Case 2530 To 2589
                If strShingumi = "GAS" Then
                    strShingumi = strShingumi & "-8"
                Else
                    strShingumi = strShingumi & "-6"
                End If
            Case Is <= 2529
                If strShingumi = "GAS" Then
                    strShingumi = strShingumi & "-7"
                Else
                    strShingumi = strShingumi & "-5"
                End If
        End Select
        
'   *VF1*********************************************************
'   *VG4*********************************************************
    ElseIf in_strHinban Like "*F?V??*-####P*-*" Or in_strHinban Like "*F?V??*-####V*-*" Then
    
        strShingumi = "PCS"
        
        Select Case dblDH
            Case 2530 To 2589
                strShingumi = strShingumi & "-15"
            Case Is <= 2529
                strShingumi = strShingumi & "-14"
        End Select

'   *AF1*********************************************************
'   *AF2*********************************************************
'   *AF3*********************************************************
    ElseIf in_strHinban Like "*F?B??*-####A*-*" Or in_strHinban Like "*F?B??*-####B*-*" Or in_strHinban Like "*F?B??*-####O*-*" Then
    
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "GAS"
        Else
            strShingumi = "GCS"
        End If

        strShingumi = strShingumi & "-15"
        
'20170517 K.Asayama ADD
'   *YF1*********************************************************
    ElseIf in_strHinban Like "Y?C??*-####F*-*" Or in_strHinban Like "特 Y?C??*-####F*-*" Then
        
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "PAS"
        Else
            strShingumi = "PCS"
        End If
        
        Select Case dblDH
            Case 2530 To 2589
                If strShingumi = "PAS" Then
                    strShingumi = strShingumi & "-16"
                Else
                    strShingumi = strShingumi & "-20"
                End If
            Case Is <= 2529
                If strShingumi = "PAS" Then
                    strShingumi = strShingumi & "-15"
                Else
                    strShingumi = strShingumi & "-19"
                End If
        End Select
        
'   *YG6*********************************************************
    ElseIf in_strHinban Like "Y?C??*-####T*-*" Or in_strHinban Like "特 Y?C??*-####T*-*" Then
        
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "GAS"
        Else
            strShingumi = "GCS"
        End If
        
        Select Case dblDH
            Case 2530 To 2589
                If strShingumi = "GAS" Then
                    If IsHidden_Hinge(in_strHinban) Then
                        strShingumi = strShingumi & "-32"
                    Else
                        strShingumi = strShingumi & "-31"
                    End If
                Else
                    If (IsEndWakunashi_Jou(in_strHinban) And Not (in_strHinban Like "*DN-####*-*" Or in_strHinban Like "*VN-####*-*")) _
                     Or (in_strHinban Like "*DH-####*-*" Or in_strHinban Like "*DF-####*-*" Or in_strHinban Like "*DJ-####*-*" Or in_strHinban Like "*DQ-####*-*" Or in_strHinban Like "*VF-####*-*" Or in_strHinban Like "*VQ-####*-*") Then
                        strShingumi = strShingumi & "-29"
                    Else
                        strShingumi = strShingumi & "-28"
                    End If
  
                End If
            Case Is <= 2529
                If strShingumi = "GAS" Then
                    If IsHidden_Hinge(in_strHinban) Then
                        strShingumi = strShingumi & "-30"
                    Else
                        strShingumi = strShingumi & "-29"
                    End If
                Else
                    If (IsEndWakunashi_Jou(in_strHinban) And Not (in_strHinban Like "*DN-####*-*" Or in_strHinban Like "*VN-####*-*")) _
                     Or (in_strHinban Like "*DH-####*-*" Or in_strHinban Like "*DF-####*-*" Or in_strHinban Like "*DJ-####*-*" Or in_strHinban Like "*DQ-####*-*" Or in_strHinban Like "*VF-####*-*" Or in_strHinban Like "*VQ-####*-*") Then
                        strShingumi = strShingumi & "-27"
                    Else
                        strShingumi = strShingumi & "-26"
                    End If
                End If
        End Select

'   *YF5/YG5******************************************************
    ElseIf in_strHinban Like "Y?B??*-####W*-*" Or in_strHinban Like "特 Y?B??*-####W*-*" Then
        
        If IsHirakido(in_strHinban) Or IsOyatobira(in_strHinban) Or IsKotobira(in_strHinban) Then
            strShingumi = "GAS"
        Else
            strShingumi = "GCS"
        End If
        
        Select Case dblDH
            Case 2530 To 2589
                If strShingumi = "GAS" Then
                    strShingumi = strShingumi & "-34"
                Else
                    strShingumi = strShingumi & "-31"
                End If
            Case Is <= 2529
                If strShingumi = "GAS" Then
                    strShingumi = strShingumi & "-33"
                Else
                    strShingumi = strShingumi & "-30"
                End If
        End Select
'20170517 K.Asayama ADD END
    End If
    
    '20161121 K.Asayama ADD
    '以下品番は末尾に「A」をつける******************************
    'アウトセットエンド枠なしで錠付（DNは除く）
    'DH
    '3枚扉
    '
    '20170517 K.Asayama Change 1701以降の新図面番号は例外
    'If strShingumi <> "" Then
    If strShingumi <> "" And bolfncShingumiShousai_SaibanReigai(in_strHinban) = False Then
    '20170517 K.Asayama Change End
    
        '20170105 K.Asayama Change
        'If (IsEndWakunashi_Jou(in_strHinban) And Not in_strHinban Like "*DN-####*-*") Or (in_strHinban Like "*DH-####*-*" Or in_strHinban Like "*DF-####*-*" Or in_strHinban Like "*DJ-####*-*" Or in_strHinban Like "*DQ-####*-*") Then
        '   strShingumi = strShingumi & "A"
        If IsEndWakunashi_Jou(in_strHinban) Or (in_strHinban Like "*DH-####*-*" Or in_strHinban Like "*DF-####*-*" Or in_strHinban Like "*DJ-####*-*" Or in_strHinban Like "*DQ-####*-*" Or in_strHinban Like "*VF-####*-*" Or in_strHinban Like "*VQ-####*-*") Then
        
            If in_strHinban Like "*DN-####*-*" Or in_strHinban Like "*VN-####*-*" Then
                '何もしない
            Else
        
                strShingumi = strShingumi & "A"
            End If
        '20170105 K.Asayama Change END
        End If
    End If
    
    fncstrShingumiShousai = strShingumi
    
End Function

Private Function bolfncShingumiShousai_SaibanReigai(ByVal in_strHinban As String) As Boolean
'   *************************************************************
'   芯組詳細図末尾付加番号例外品番確認
'   'ADD by Asayama 20170517
'   戻り値:boolean
'       →                  True 例外対象 False 例外の対象外
'
'    Input項目
'       in_strhinban        建具品番

'   201701以降新図面番号(Terrace以降)は対象外になる
'   *************************************************************
    On Error GoTo Err_bolfncShingumiShousai_SaibanReigai
    
    bolfncShingumiShousai_SaibanReigai = False
    
    If IsTerrace(in_strHinban) Then
        bolfncShingumiShousai_SaibanReigai = True
    End If

    Exit Function
    
Err_bolfncShingumiShousai_SaibanReigai:
    bolfncShingumiShousai_SaibanReigai = False
    
End Function

Public Function bolFncSan_Koteichi_Nakaita(ByVal in_dblDW As Double, ByVal in_dblDH As Double, ByVal in_strHinban As String, ByRef out_dblsan As Double, ByRef out_dblGakuYoko As Double) As Boolean
'   *************************************************************
'   上下桟固定値（ガラス、ミラー扉の際の固定値--中板）
'   'ADD by Asayama 20160923

'   戻り値:Boolean
'       →True              照合OK　数値戻し
'       →True              照合NG　数値なし
'
'    Input項目
'       in_dblDW            DW
'       in_dblDH            DH（作成時は未使用）
'       in_strHinban        品番
'
'    Output項目
'      寸法
'       out_dblsan          上下桟
'       out_dblgakuyoko     額横
'
'
'****************************************************************

    bolFncSan_Koteichi_Nakaita = True

    Select Case in_dblDW
        Case 426 To 570.9
            out_dblsan = in_dblDW - 281: out_dblGakuYoko = in_dblDW - 341
            
        Case 571 To 618.9
            out_dblsan = in_dblDW - 325: out_dblGakuYoko = in_dblDW - 385
        
        Case 619 To 669.9
            out_dblsan = in_dblDW - 357: out_dblGakuYoko = in_dblDW - 417
        
        Case 670 To 717.9
            out_dblsan = in_dblDW - 389: out_dblGakuYoko = in_dblDW - 449
        
        Case 718 To 750.9
            out_dblsan = in_dblDW - 423: out_dblGakuYoko = in_dblDW - 483
                    
        Case 751 To 780.9
            '20161011 K.Asayama Change 変更
            'out_dblsan = in_dblDW - 439: out_dblGakuYoko = in_dblDW - 499
            out_dblsan = in_dblDW - 437: out_dblGakuYoko = in_dblDW - 497
            
        Case 781 To 819.9
            out_dblsan = in_dblDW - 461: out_dblGakuYoko = in_dblDW - 521
            
        Case 820 To 862.9
            out_dblsan = in_dblDW - 487: out_dblGakuYoko = in_dblDW - 547
                    
        Case 863 To 900.9
            out_dblsan = in_dblDW - 511: out_dblGakuYoko = in_dblDW - 571
                    
        Case 901 To 944.9
            out_dblsan = in_dblDW - 545: out_dblGakuYoko = in_dblDW - 605
                    
        Case 945 To 985.9
            out_dblsan = in_dblDW - 573: out_dblGakuYoko = in_dblDW - 633
        
        Case 986 To 1022.9
            out_dblsan = in_dblDW - 597: out_dblGakuYoko = in_dblDW - 657
                    
        Case 1023 To 1061.9
            out_dblsan = in_dblDW - 623: out_dblGakuYoko = in_dblDW - 683
            
        Case 1062 To 1100
            out_dblsan = in_dblDW - 645: out_dblGakuYoko = in_dblDW - 705
            
        Case Else
            out_dblsan = 0: out_dblGakuYoko = 0
            bolFncSan_Koteichi_Nakaita = False
    End Select

End Function

Private Function IsEndWakunashi_Jou(ByVal in_varHinban As Variant) As Boolean
'   *************************************************************
'   エンド枠無し 錠付確認用関数
'   'ADD by Asayama 20161121

'   戻り値:Boolean
'       →True              エンド枠無し品番且つ錠付
'       →False             エンド枠無し品番、錠付以外
'
'    Input項目
'       in_strHinban        建具品番
'
'
'****************************************************************

    Dim strHinban As String

    On Error GoTo Err_IsEndWakunashi_Jou
        
    IsEndWakunashi_Jou = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "特 ", "")

    If IsEndWakunashi(strHinban) And (strHinban Like "*-####*-K*" Or strHinban Like "*-####*-M*") Then
    
        IsEndWakunashi_Jou = True
        
    End If

    Exit Function

Err_IsEndWakunashi_Jou:
    IsEndWakunashi_Jou = False

End Function

Private Function intFncSode1Honsu_Group1(in_varHinban As Variant, in_Maisu As Integer) As Integer
'   *************************************************************
'   袖1本数集計グループ1
'   'ADD by Asayama 20161121
    
'   対象グループ

'
'   戻り値:Integer
'       →                  袖1本数
'
'    Input項目
'       in_strHinban        建具品番
'       in_Maisu            建具枚数
'
'****************************************************************

    Dim strHinban As String
    Dim intHonsu As Integer
    
    On Error GoTo Err_intFncSode1Honsu_Group1
        
    intFncSode1Honsu_Group1 = 0
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "特 ", "")
    
    'Monster
    If IsMonster(strHinban) Then
        intHonsu = 6 * in_Maisu
        
    'DN例外
    '20170105 K.Asayama Change
'    ElseIf strHinban Like "*DN-####*-*" Then
    ElseIf strHinban Like "*DN-####*-*" Or strHinban Like "*VN-####*-*" Then
    '20170105 K.Asayama Change END
        
        '20170517 K.Asayama Change Terrace追加
        'intHonsu = 2 * in_Maisu
        If IsTerrace(strHinban) Then
            intHonsu = 5 * in_Maisu
        Else
            intHonsu = 2 * in_Maisu
        End If
        
    'アウトセットエンド枠なし錠付例外(DU,KU)
    ElseIf IsEndWakunashi_Jou(strHinban) Then
        intHonsu = 4 * in_Maisu
    
    'DHと3枚扉例外
    '20170105 K.Asayama Change
'    ElseIf strHinban Like "*DH-####*-*" Or strHinban Like "*DF-####*-*" Or strHinban Like "*DJ-####*-*" Or strHinban Like "*DQ-####*-*" Then
    ElseIf strHinban Like "*DH-####*-*" Or strHinban Like "*DF-####*-*" Or strHinban Like "*DJ-####*-*" Or strHinban Like "*DQ-####*-*" Or strHinban Like "*VF-####*-*" Or strHinban Like "*VQ-####*-*" Then
    '20170105 K.Asayama Change END
    
        intHonsu = 4 * in_Maisu
    
    Else
        intHonsu = 5 * in_Maisu
    End If
            
    intFncSode1Honsu_Group1 = intHonsu
    
    Exit Function
    
Err_intFncSode1Honsu_Group1:
    intFncSode1Honsu_Group1 = 0
    
End Function

Private Function intFncSode2Honsu_Group1(in_varHinban As Variant, in_Maisu As Integer) As Integer
'   *************************************************************
'   袖2本数集計グループ1
'   'ADD by Asayama 20161121
    

'
'   戻り値:Integer
'       →                  袖2本数（ヴェルチカのみ袖1本数に挿入すること）
'
'    Input項目
'       in_strHinban        建具品番
'       in_Maisu            建具枚数
'
'****************************************************************

    Dim strHinban As String
    Dim intHonsu As Integer
    
    On Error GoTo Err_intFncSode2Honsu_Group1
        
    intFncSode2Honsu_Group1 = 0
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "特 ", "")
    
    'DN例外
    '20170105 K.Asayama Change
'    If strHinban Like "*DN-####*-*" Then
    If strHinban Like "*DN-####*-*" Or strHinban Like "*VN-####*-*" Then
    '20170105 K.Asayama Change END

        intHonsu = 3 * in_Maisu
        
    'アウトセットエンド枠なし錠付例外(DU,KU)
    ElseIf IsEndWakunashi_Jou(strHinban) Then
        intHonsu = 5 * in_Maisu
    
    'DHと3枚扉例外
    '20170105 K.Asayama Change
'    ElseIf strHinban Like "*DH-####*-*" Or strHinban Like "*DF-####*-*" Or strHinban Like "*DJ-####*-*" Or strHinban Like "*DQ-####*-*" Then
    ElseIf strHinban Like "*DH-####*-*" Or strHinban Like "*DF-####*-*" Or strHinban Like "*DJ-####*-*" Or strHinban Like "*DQ-####*-*" Or strHinban Like "*VF-####*-*" Or strHinban Like "*VQ-####*-*" Then
    '20170105 K.Asayama Change END
    
        intHonsu = 5 * in_Maisu
    
    Else
        intHonsu = 6 * in_Maisu
    End If
            
    intFncSode2Honsu_Group1 = intHonsu
    
    Exit Function
    
Err_intFncSode2Honsu_Group1:
    intFncSode2Honsu_Group1 = 0
    
End Function

Private Function intFncSode2Honsu_Group2(in_varHinban As Variant, in_Maisu As Integer) As Integer
'   *************************************************************
'   袖2本数集計グループ2
'   'ADD by Asayama 20161121
    
'   対象グループ
'       サイドスルー(SG1/PG1型)

'
'   戻り値:Integer
'       →                  袖2本数
'
'    Input項目
'       in_strHinban        建具品番
'       in_Maisu            建具枚数
'
'****************************************************************

    Dim strHinban As String
    Dim intHonsu As Integer
    
    On Error GoTo Err_intFncSode2Honsu_Group2
        
    intFncSode2Honsu_Group2 = 0
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "特 ", "")
    
    'DN例外
    '20170105 K.Asayama Change
'    If strHinban Like "*DN-####*-*" Then
    If strHinban Like "*DN-####*-*" Or strHinban Like "*VN-####*-*" Then
    '20170105 K.Asayama Change END

        intHonsu = 2 * in_Maisu
        
    'アウトセットエンド枠なし錠付例外(DU,KU)
    ElseIf IsEndWakunashi_Jou(strHinban) Then
        intHonsu = 2 * in_Maisu
    
    ElseIf IsHikido(strHinban) Then
        intHonsu = 3 * in_Maisu
        
    Else
        intHonsu = 0
        
    End If
            
    intFncSode2Honsu_Group2 = intHonsu
    
    Exit Function
    
Err_intFncSode2Honsu_Group2:
    intFncSode2Honsu_Group2 = 0
    
End Function

Public Function dblFncGakuYokoNaisun(in_strHinban As String, in_dblDW As Double) As Double
'   *************************************************************
'   額横内々寸法
'   'ADD by Asayama 20170517
    
'   対象品目
'       YG6型

'
'   戻り値:Double
'       →                  額横内々寸法
'
'    Input項目
'       in_strHinban        建具品番
'       in_dblDW            DW
'
'****************************************************************
    dblFncGakuYokoNaisun = 0
    
    On Error GoTo Err_dblFncGakuYokoNaisun
    
    If IsTerrace(in_strHinban) And in_strHinban Like "*-####T*-*" Then
        Select Case in_dblDW
            Case Is > 900 'ありえない
                dblFncGakuYokoNaisun = 0
            Case Is > 740
                dblFncGakuYokoNaisun = 450
            Case 670 To 740
                dblFncGakuYokoNaisun = 379
            Case Is < 670  'ありえない
                dblFncGakuYokoNaisun = 0
        End Select
    End If
    
    Exit Function

Err_dblFncGakuYokoNaisun:

End Function

Public Function dblFncGakuTate1_YG6(in_strHinban As String, in_dblDW As Double) As Double
'   *************************************************************
'   額縦1寸法
'   'ADD by Asayama 20170517
    
'   対象品目
'       YG6型

'
'   戻り値:Double
'       →                  額縦内々寸法
'
'    Input項目
'       in_strHinban        建具品番
'       in_dblDW            DW
'
'****************************************************************
    dblFncGakuTate1_YG6 = 0
    
    On Error GoTo Err_dblFncGakuTate1_YG6
    
    Select Case in_dblDW
        Case Is > 900 'ありえない
            dblFncGakuTate1_YG6 = 0
        Case Is > 740
            dblFncGakuTate1_YG6 = 400
        Case 670 To 740
            dblFncGakuTate1_YG6 = 337
        Case Is < 670  'ありえない
            dblFncGakuTate1_YG6 = 0
    End Select

    
    Exit Function

Err_dblFncGakuTate1_YG6:

End Function

Public Function IsAWKansuSearch_Needed(in_varHinban As Variant) As Boolean
'   *************************************************************
'   T_AW関数検索対象識別
'   'ADD by Asayama 20161102

'   戻り値:Boolean
'       →True              T_AW関数検索対象
'       →False             T_AW関数検索対象外
'
'    Input項目
'       in_varHinban        建具品番

'   *************************************************************
    
    Dim strHinban As String
    
    On Error GoTo Err_IsAWKansuSearch_Needed
    
    IsAWKansuSearch_Needed = False
    
    '品番Nullの場合False
    If IsNull(in_varHinban) Then Exit Function
    
    '頭の[特 ]を外す
    strHinban = Replace(in_varHinban, "特 ", "")
    
    '子扉は対象外
    If IsKotobira(strHinban) Then Exit Function
    
    '-----------------------------------------------------------
    
    '3型旧7型は対象
    If strHinban Like "??B*-####MF*-*" Or strHinban Like "??B*-####G*-*" Then
        IsAWKansuSearch_Needed = True
        
        Exit Function
    End If
    
    '7型（フラッシュ）は対象
    If IsG7_Flush(strHinban) Then
        IsAWKansuSearch_Needed = True
        
        Exit Function
    End If
    
    Exit Function

Err_IsAWKansuSearch_Needed:
    IsAWKansuSearch_Needed = False
End Function