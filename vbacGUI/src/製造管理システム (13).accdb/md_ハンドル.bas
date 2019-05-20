Option Compare Database
Option Explicit

Public Function fncbol_Handle_引手_長(施錠 As String, 個別Spec As String) As Boolean
'--------------------------------------------------------------------------------
'           長サイズの引手錠か確認する
'
'引数       施錠            : 施錠コード(2 or 3桁)
'           個別Spec        : 仕様コード(7桁)
'戻り値     Boolean         : True  →長サイズ引手  False   →長サイズ引手ではない
'--------------------------------------------------------------------------------
    Dim str会社 As String
    Dim strSPEC As String
    
    str会社 = left(個別Spec, 3)
    strSPEC = right(個別Spec, 4)
    
    Select Case strSPEC
        Case Is >= "1410"

            If 施錠 Like "ZA?" Or 施錠 Like "ZC?" Or 施錠 Like "ZV?" Or 施錠 Like "ZE?" Then
            
                fncbol_Handle_引手_長 = True
            Else
                fncbol_Handle_引手_長 = False
                
            End If
        
        Case Is >= "1404"
            If 施錠 Like "A?" Or 施錠 Like "C?" Or 施錠 Like "V?" Then
               
                fncbol_Handle_引手_長 = True
                
            Else
                fncbol_Handle_引手_長 = False
                
            End If
        Case Is >= "1011"
            If 施錠 Like "W?" Then
               
                fncbol_Handle_引手_長 = True
                
            Else
                fncbol_Handle_引手_長 = False
                
            End If
        Case "0911" To "1010"
            If 施錠 Like "Z?" Or 施錠 Like "Y?" Then
               
                fncbol_Handle_引手_長 = True
                
            Else
                fncbol_Handle_引手_長 = False
                
            End If
            
        Case Else
            fncbol_Handle_引手_長 = False
            
    End Select

End Function

Public Function fncbol_Handle_引手_短(施錠 As String, 個別Spec As String) As Boolean
'--------------------------------------------------------------------------------
'           短サイズの引手錠か確認する
'
'引数       施錠            : 施錠コード(2 or 3桁)
'           個別Spec        : 仕様コード(7桁)
'戻り値     Boolean         : True  →短サイズ引手  False   →短サイズ引手ではない
'--------------------------------------------------------------------------------

    Dim str会社 As String
    Dim strSPEC As String
    
    str会社 = left(個別Spec, 3)
    strSPEC = right(個別Spec, 4)
    
    Select Case strSPEC
        Case Is >= "1410"

            If 施錠 Like "ZG?" Or 施錠 Like "ZH?" Or 施錠 Like "ZB?" Or 施錠 Like "ZD?" Then
            
                fncbol_Handle_引手_短 = True
            Else
                fncbol_Handle_引手_短 = False
                
            End If
            
        Case Is >= "1307"
            If 施錠 Like "G?" Or 施錠 Like "H?" Or 施錠 Like "B?" Then
               
                fncbol_Handle_引手_短 = True
                
            Else
                fncbol_Handle_引手_短 = False
                
            End If
            
        Case Else
            fncbol_Handle_引手_短 = False
    End Select
    
End Function

Public Function fncbol_Handle_EStyle(施錠 As String, 個別Spec As String) As Boolean
'--------------------------------------------------------------------------------
'E-Styleのデフォルトのハンドル(I,J,K)か確認する
'
'引数       施錠            : 施錠コード(2 or 3桁)
'           個別Spec        : 仕様コード(7桁)
'戻り値     Boolean         : True  →E-Styleハンドル  False   →E-Styleハンドルではない
'--------------------------------------------------------------------------------

    Dim str会社 As String
    Dim strSPEC As String
    
    str会社 = left(個別Spec, 3)
    strSPEC = right(個別Spec, 4)
    

    Select Case strSPEC

        Case Is >= "1601"
            If 施錠 Like "CI?" Or 施錠 Like "CJ?" Or 施錠 Like "CK?" Or 施錠 Like "CG?" Or 施錠 Like "CH?" Or 施錠 Like "CB?" _
                Or 施錠 Like "DI?" Or 施錠 Like "DJ?" Or 施錠 Like "DK?" Or 施錠 Like "DG?" Or 施錠 Like "DH?" Or 施錠 Like "DB?" Then
                fncbol_Handle_EStyle = True
            Else
                fncbol_Handle_EStyle = False
            End If

        Case Is >= "1410"
            If 施錠 Like "AI?" Or 施錠 Like "AJ?" Or 施錠 Like "AK?" Or 施錠 Like "AA?" Or 施錠 Like "AH?" Then
                fncbol_Handle_EStyle = True
            Else
                fncbol_Handle_EStyle = False
            End If
        Case Is >= "0911"
            If 施錠 Like "I?" Or 施錠 Like "J?" Or 施錠 Like "K?" Then
                fncbol_Handle_EStyle = True
            Else
                fncbol_Handle_EStyle = False
            End If
        Case Else
            fncbol_Handle_EStyle = False
    End Select

                               
End Function

Public Function fncbol_Handle_WanNyan(施錠 As String, 個別Spec As String) As Boolean
'--------------------------------------------------------------------------------

'           わんにゃんハンドル(O,N,E,F)か確認する
'
'引数       施錠            : 施錠コード(2 or 3桁)
'           個別Spec        : 仕様コード(7桁)
'戻り値     Boolean         : True  →わんにゃんハンドル  False   →わんにゃんハンドルではない
'--------------------------------------------------------------------------------
    Dim str会社 As String
    Dim strSPEC As String
    
    str会社 = left(個別Spec, 3)
    strSPEC = right(個別Spec, 4)
    

    Select Case strSPEC
        Case Is >= "1410"
            If 施錠 Like "AE?" Or 施錠 Like "AF?" Or 施錠 Like "AN?" Or 施錠 Like "AO?" Then
                fncbol_Handle_WanNyan = True
            Else
                fncbol_Handle_WanNyan = False
            End If
        Case Is >= "1404"
            If 施錠 Like "E?" Or 施錠 Like "F?" Or 施錠 Like "N?" Or 施錠 Like "O?" Then
                fncbol_Handle_WanNyan = True
            Else
                fncbol_Handle_WanNyan = False
            End If
        Case Else
            fncbol_Handle_WanNyan = False
    End Select

                               
End Function

Public Function fncstrHandle_Name(施錠 As String, 個別Spec As String) As String
'--------------------------------------------------------------------------------

'           ハンドルコードを抽出する
'
'引数       施錠            : 施錠コード(2 or 3桁)
'           個別Spec        : 仕様コード(7桁)
'戻り値     String(1~2桁)   : ハンドルコード
'--------------------------------------------------------------------------------

    Dim str会社 As String
    Dim strSPEC As String
    
    str会社 = left(個別Spec, 3)
    strSPEC = right(個別Spec, 4)
    
    If 施錠 Like "特*" Then
        fncstrHandle_Name = "特"
        
    ElseIf str会社 = "TSC" Then
        fncstrHandle_Name = left(施錠, 2)
    
    Else
        Select Case strSPEC
            Case Is >= "1410"
                If Len(施錠) > 2 Then 'ｸﾛｾﾞｯﾄは2桁
                    fncstrHandle_Name = left(施錠, 2)
                Else
                    fncstrHandle_Name = left(施錠, 1)
                End If
            Case Else
                fncstrHandle_Name = left(施錠, 1)
        End Select
     End If
End Function

Public Function fncbol錠(施錠 As String, 個別Spec As String) As Boolean
'--------------------------------------------------------------------------------

'           錠あり無しを判定する
'
'引数       施錠            : 施錠コード(2 or 3桁)
'           個別Spec        : 仕様コード(7桁)
'戻り値     Boolean         : True  →錠あり  False   →錠なし
'--------------------------------------------------------------------------------
    
    If 施錠 Like "*C" Or 施錠 Like "*M" Or 施錠 Like "*K" Then
        fncbol錠 = True
    Else
        fncbol錠 = False
    End If
    
End Function

Public Function fncbol_Handle_Vertica(施錠 As String, 個別Spec As String) As Boolean
'--------------------------------------------------------------------------------
'           ヴェルチカ（取手レス）(-N,-B)か確認する
'
'引数       施錠            : 施錠コード(3桁)
'           個別Spec        : 仕様コード(7桁)
'戻り値     Boolean         : True  →ヴェルチカ  False   →ヴェルチカではない
'--------------------------------------------------------------------------------
    Dim str会社 As String
    Dim strSPEC As String
    
    str会社 = left(個別Spec, 3)
    strSPEC = right(個別Spec, 4)
    

    Select Case strSPEC
        Case Is >= "1507"
            If 施錠 Like "-N?" Or 施錠 Like "-B?" Or 施錠 Like "-Q?" Or 施錠 Like "-K?" Then
                fncbol_Handle_Vertica = True
            Else
                fncbol_Handle_Vertica = False
            End If
        Case Else
            fncbol_Handle_Vertica = False
    End Select

                               
End Function

Public Function fncstrHandleKigoFileName(ByVal in_Hinban As String, ByVal in_Handle As String, ByVal in_spec As String) As String
'   *************************************************************
'   ハンドル記号から記号図のファイル名を取得

'       ※新ライン用データが出来るまでの暫定利用
'
'   戻り値:String
'       →画像ファイル名
'
'    Input項目
'       in_Hinban           建具品番
'       in_Handle           施錠
'       in_Spec             個別SPec

'2.1.0
'   →DP,DQハンドル追加（蔵前）
'2.3.0
'   →1801仕様追加
'2.7.0
'   →1808仕様追加
'2.13.0
'   →1901仕様追加(HE-HI)
'   *************************************************************
    Dim Reg As Object
    'ディクショナリ
    Dim dictHandle As Object
    '配列用（ディクショナリのパターン挿入）
    Dim varPattern As Variant
    
    Dim i As Integer
    Dim strHandleHikiteName As String
    
    fncstrHandleKigoFileName = ""
    
    If in_Hinban = "" Or in_Handle = "" Or in_spec = "" Then Exit Function
    
    '施錠コードが3桁（特注除く）以外は対応しない（過去のコードは対応しない）
    If Len(in_Handle) <> 3 Then Exit Function
    If in_Handle Like "*特*" Then Exit Function
    
    '正規表現
    Set Reg = CreateObject("VBScript.RegExp")
    'ディクショナリ
    Set dictHandle = CreateObject("Scripting.Dictionary")
    
    On Error GoTo Err_fncstrHandleKigoFileName
    
    'ディクショナリに正規表現のパターンを挿入
    '（キー→パターン、アイテム→ファイル名）
    '速度向上のために当たりやすいパターンから並べるべし
    
    With dictHandle
        If IsHirakido(in_Hinban) Or IsOyatobira(in_Hinban) Then
    
            
            .Add "^(C[B-KST]|D[B-FI-KST])", "SHIBUTANI"
            .Add "^(H[A-IPQR])", "KAWAJ_LJ"
            .Add "^(C[N-R]|D[NO])", "KAWAJUN"
            .Add "^(C[LM]|B[YZ]|B[ACDEFHIJLMNOPQRS]|D[PQ])", "KURAMAE"
            .Add "^A[EFON]", "NAGASAWA"
            .Add "^F[C-H]", "OLIVARI"

        ElseIf IsHikido(in_Hinban) Or IsCloset_Hikichigai(in_Hinban) Then
    
            .Add "^Z[ACEV]", "OKUDAIRATK"
            .Add "^Z[BDHG]", "OKUDAIRA"
            .Add "^-(?!-).(?!N)", "HIKITELESS"
           
        End If
        
        'キーにマッチしたらファイル名を取得
        If .Count > 0 Then '←引戸でも開き戸でもない場合はカウントがゼロ
            varPattern = .keys
            
            For i = 0 To UBound(varPattern)
            
                Reg.Pattern = varPattern(i)
                
                If Reg.Test(in_Handle) Then
                    strHandleHikiteName = .Item(varPattern(i))
                    'Debug.Print "Match! " & in_Hinban & " key=" & strHandleHikiteName
                    Exit For
                End If
                
            Next
            
        End If
    
    End With
    
    If strHandleHikiteName = "" Then
        'Debug.Print "Nomatch!"
    Else
        '錠付きの場合はファイル名にさらに錠の名前を付加
        If fncbol錠(in_Handle, in_spec) Then
        
            strHandleHikiteName = strHandleHikiteName & "_Lock"
            
            'ワンニャンの場合はシリンダー錠、間仕切り錠の区別あり
            If fncbol_Handle_WanNyan(in_Handle, in_spec) Then
                Select Case right(in_Handle, 1)
                    Case "C"
                        strHandleHikiteName = strHandleHikiteName & "_Cylinder"
                    Case "M", "K"
                        strHandleHikiteName = strHandleHikiteName & "_Majikiri"
                End Select
            
            'エンド枠無しはアウトセット引戸錠
            ElseIf IsEndWakunashi(in_Hinban) Then
            
                strHandleHikiteName = strHandleHikiteName & "_Outset"
                
            End If
        End If
    End If
    
    fncstrHandleKigoFileName = strHandleHikiteName
    
    GoTo Exit_fncstrHandleKigoFileName

Err_fncstrHandleKigoFileName:
    Debug.Print Err.Description
    fncstrHandleKigoFileName = ""
    
Exit_fncstrHandleKigoFileName:
    Set Reg = Nothing
    Set dictHandle = Nothing
        
End Function

Public Function fncstrHikiteColorName(ByVal in_Handle As String, ByVal in_spec As String) As String
'   *************************************************************
'   引き手記号から引き手色を取得

'
'   戻り値:String
'       →色名
'
'    Input項目
'       in_Handle           施錠
'       in_Spec             個別SPec

'   *************************************************************
    
    fncstrHikiteColorName = ""
    
    If in_Handle = "" Or in_spec = "" Then Exit Function
    
    '施錠コードが3桁（特注除く）以外は対応しない（過去のコードは対応しない）
    If Len(in_Handle) <> 3 Then Exit Function
    If in_Handle Like "*特*" Then Exit Function
    
    If in_Handle Like "ZV*" Or in_Handle Like "ZH*" Then
        fncstrHikiteColorName = "サテンニッケル"
    ElseIf in_Handle Like "ZC*" Or in_Handle Like "ZG*" Then
        fncstrHikiteColorName = "クローム"
    ElseIf in_Handle Like "ZA*" Or in_Handle Like "ZB*" Then
        fncstrHikiteColorName = "ホワイト"
    ElseIf in_Handle Like "ZE*" Or in_Handle Like "ZD*" Then
        fncstrHikiteColorName = "ブラック"
    End If
    
End Function

Public Function IsHikiteKako(ByVal in_varHinban As Variant, ByVal in_varTobiraichi As Variant, ByVal in_varTsurimoto As Variant, ByVal in_varSpec As Variant) As Boolean
'   *************************************************************
'   引き手加工があるか確認
'
'   戻り値:Boolean
'       True                引手加工あり
'       False               加工なし（なしの条件以外はすべてありで返す）
'
'    Input項目
'       in_varHinban        建具品番
'       in_varTobiraichi    位置（1.左、2.右、3.中）
'       in_Spec             個別SPec(作成時未使用）

'2.14.0 ADD
'   *************************************************************
    
    Dim strHinban As String
    Dim intTobiraichi As Integer
    Dim strSPEC As String
    Dim strTsurimoto As String
    
    IsHikiteKako = True
    
    '品番なしの場合はTrueで返す
    If IsNull(in_varHinban) Then
        Exit Function
    End If
    
    '位置なしの場合はTrueで返す

    If IsNull(in_varTobiraichi) Then
        Exit Function
    End If
    
    If IsNull(in_varTsurimoto) Then
        strTsurimoto = "Z"
    Else
        strTsurimoto = in_varTsurimoto
    End If
    
    If IsNull(in_varSpec) Then
        strSPEC = ""
    Else
        strSPEC = in_varSpec
    End If
    
    strHinban = Replace(in_varHinban, "特 ", "")
    intTobiraichi = in_varTobiraichi
    
    If IsSynchro(strHinban) Then
        If IsHikichigai(strHinban) Then
            If intTobiraichi = 3 Then
                IsHikiteKako = False
            End If
        Else
            If strTsurimoto = "L" Then 'L吊元
                If intTobiraichi <> 1 Then
                    IsHikiteKako = False
                End If
            ElseIf strTsurimoto = "R" Then 'R吊元
                If intTobiraichi <> 2 Then
                    IsHikiteKako = False
                End If
            End If
        End If
    
    End If

End Function