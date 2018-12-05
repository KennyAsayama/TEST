Option Compare Database
Option Explicit

Public Function IsCasing(in_strWakuHinban As Variant) As Boolean
'   *************************************************************
'   ケーシング確認
'
'   戻り値:Boolean
'       →True              ケーシング
'       →False             ケーシング以外
'
'    Input項目
'       in_strHinban        枠品番

'   *************************************************************
    IsCasing = False
    
    If in_strWakuHinban Like "*X*KH*-####*" Or in_strWakuHinban Like "*Y*KH*-####*" Then
        IsCasing = True
    End If
    
End Function

Public Function IsCloset(in_strSetHinban As Variant) As Boolean
'   *************************************************************
'   クローゼット確認
'
'   戻り値:Boolean
'       →True              クローゼット
'       →False             クローゼット以外
'
'    Input項目
'       in_strSetHinban     セット品番

'   *************************************************************
    IsCloset = False
    
    If in_strSetHinban Like "M??-?-?####*-*" Or in_strSetHinban Like "特 M??-?-?####*-*" Then
        IsCloset = True
    End If
    
End Function

Public Function IsCloset_Isehara(in_strHinban As Variant) As Boolean
'   *************************************************************
'   伊勢原生産クローゼット確認
'
'   戻り値:Boolean
'       →True              伊勢原生産クローゼット
'       →False             伊勢原生産クローゼット以外
'
'    Input項目
'       in_strHinban        建具品番

'   *************************************************************
    IsCloset_Isehara = False
    
    If in_strHinban Like "*CME-####*-*" Or in_strHinban Like "*CSA-####*-*" Then
        IsCloset_Isehara = True
    End If
    
End Function

Public Function IsStealth(in_strHinban As Variant) As Boolean
'   *************************************************************
'   ステルス確認
'
'   戻り値:Boolean
'       →True              ステルス
'       →False             ステルス以外
'
'    Input項目
'       in_strHinban        下地品番

'   *************************************************************
    IsStealth = False
    
    If Not in_strHinban Like "*KG*-####*" Then
        IsStealth = True
    End If
    
End Function
Public Function IsStealth_Seizo(in_strHinban As Variant) As Boolean
'   *************************************************************
'   ステルス（製造）確認
'
'   戻り値:Boolean
'       →True              ステルス（製造）
'       →False             ステルス（製造）以外
'
'    Input項目
'       in_strHinban        下地品番

'   *************************************************************
    '20150820現在未使用
    
    IsStealth_Seizo = False
    
'    If in_strHinban Like "*PW*-####*" Then 'エスパスライドウォールはインセット
'        IsStealth_Seizo = False
'        Exit Function
'    End If
'
'    If (in_strHinban Like "*SG*-####*" Or in_strHinban Like "*NG*-####*" Or in_strHinban Like "*AG*-####*" Or in_strHinban Like "*BG*-####*") _
'        And Not in_strHinban Like "*ML-####*" And Not in_strHinban Like "*MK-####*" And Not in_strHinban Like "*MT-####*" And Not in_strHinban Like "*DU-####*" And Not in_strHinban Like "*DN-####*" And Not in_strHinban Like "*CTSG*MK-####*" And Not in_strHinban Like "*CTSG*ML-####*" And Not in_strHinban Like "*CTSG*MT-####*" And Not in_strHinban Like "*KU-####*" And Not in_strHinban Like "*KN-####*" And Not in_strHinban Like "*DV-####*" Then
'        IsStealth_Seizo = True
'    End If
    
    
End Function
Public Function intFncSeizokubun(in_strShurui As String, in_varHinban As Variant) As Integer
'   *************************************************************
'   製造区分取得
'
'   戻り値:Integer
'       →                  製造区分
'
'    Input項目
'       in_strShurui        種類
'       in_varHinban        品番

'2.7.0
'   →フルガラスは区分0（製造しない）
'   *************************************************************
    
    intFncSeizokubun = 0
    
    Select Case in_strShurui
    
        Case "建具", "子扉"
            
            If IsFullGlass(in_varHinban) Then
                
                intFncSeizokubun = 0
            
            ElseIf IsKamachi(in_varHinban) Then
            
                intFncSeizokubun = 3
                
            ElseIf IsFkamachi(in_varHinban) Then
            
                intFncSeizokubun = 2
                
            Else
            
                intFncSeizokubun = 1
                
            End If
            
        Case "ｸﾛｾﾞｯﾄ"
        
            If IsCloset_Isehara(in_varHinban) Then  'ｸﾛｾﾞｯﾄ(伊勢原生産)
                intFncSeizokubun = 1
            End If
        Case "枠"
        
            If IsCasing(in_varHinban) Then
                intFncSeizokubun = 5
            Else
                intFncSeizokubun = 4
            End If
        Case "下地"
        
            If IsStealth_Seizo(in_varHinban) Then
                intFncSeizokubun = 7
            Else
                intFncSeizokubun = 6
            End If
            
    End Select
    
End Function

Public Function intFncSeihinkubun(in_strShurui As String, in_varHinban As Variant) As Integer
'   *************************************************************
'   製造区分取得
'   製品区分名からコードを取得
'
'   戻り値:Integer
'       →                  製造区分
'
'    Input項目
'       in_strShurui        種類
'       in_varHinban        品番

'   *************************************************************

'*****************************************
'1.下にコードから区分名を取得する逆引きあり
'　(更新の際は同期を取ること)
'2.コードの追加変更削除の場合チェックリストの関数も修正
'　(関数名:intFncSeizoKubunToChecklistCode)
'*****************************************

    Dim intChecklistikubun As Integer
    
    intFncSeihinkubun = 0
    
    Select Case in_strShurui
    
        Case "建具", "子扉"

            intFncSeihinkubun = 1
            
        Case "ｸﾛｾﾞｯﾄ"

            intFncSeihinkubun = 5
            
        Case "枠"
        
            If IsCasing(in_varHinban) Then
                intFncSeihinkubun = 4
            Else
                intFncSeihinkubun = 2
            End If

            
        Case "下地"
        
            intFncSeihinkubun = 3
            
        Case "造作材"
        
            intFncSeihinkubun = 6
            
        Case "玄関収納"
        
            intFncSeihinkubun = 7
            
        Case "金物"
        
            intFncSeihinkubun = 8
            
        Case "配送費"
        
            intFncSeihinkubun = 9
            
        Case "床材"

            intFncSeihinkubun = 10

        Case "階段"

            intFncSeihinkubun = 11

        Case "ﾌｧﾆﾁｭｱ"

            intFncSeihinkubun = 12
           
    End Select
    
        
End Function

Public Function strFncSeihinkubunMei(in_intSeihinkubun As Integer) As String
'   *************************************************************
'   製造区分名取得
'   製品コードから区分名を取得
'
'   戻り値:Integer
'       →                  製造区分
'
'    Input項目
'       in_strShurui        種類
'       in_varHinban        品番

'   *************************************************************

'*****************************************
'上に区分名からコードを取得する逆引きあり
'(更新の際は同期を取ること)
'*****************************************
    strFncSeihinkubunMei = ""
    
    Select Case in_intSeihinkubun

        Case 5
            strFncSeihinkubunMei = "ｸﾛｾﾞｯﾄ"

        Case 1
            strFncSeihinkubunMei = "建具"

        Case 2
            strFncSeihinkubunMei = "枠"

        Case 4
            strFncSeihinkubunMei = "三方枠"

        Case 3
            strFncSeihinkubunMei = "下地"

        Case 6
            strFncSeihinkubunMei = "造作材"

        Case 7
            strFncSeihinkubunMei = "玄関収納"

        Case 8
            strFncSeihinkubunMei = "金物"
        
        Case 9
            strFncSeihinkubunMei = "配送費"
            
        Case 10
            strFncSeihinkubunMei = "床材"

        Case 11
            strFncSeihinkubunMei = "階段"

        Case 12
            strFncSeihinkubunMei = "ﾌｧﾆﾁｭｱ"

          
    End Select
    
End Function

Public Function IsFkamachi(in_strHinban As Variant) As Boolean
'   *************************************************************
'   Flush框確認
'
'   戻り値:Boolean
'       →True              F框
'       →False             F框以外
'
'    Input項目
'       in_strHinban        建具品番

'   1.10.11 20160302 K.Asayama ADD
'           →エスパスライドウォール追加
'   1.10.12 20160322 K.Asayama Change
'           →AF1～AF3（カロ）追加
'   1.10.19 K.Asayama Change
'           →1608以降のミラーはFlush（スルーガラス）
'   1.11.0
'           →テラスドア追加
'   1.11.3
'           →モンスター品番変更対応（関数）
'   2.3.0
'           →1801仕様追加　G9型
'   2.5.2
'           →1801仕様追加　格子扉
'   *************************************************************
    
    IsFkamachi = False
    
    If IsNull(in_strHinban) Then Exit Function
       
    '1.10.19
    'If in_strHinban Like "*-####G*-*" Or in_strHinban Like "*-####MF*-*" Or in_strHinban Like "*O*-####P*-*" Then
    If in_strHinban Like "*-####G*-*" Or in_strHinban Like "F?B*-####MF*-*" Or in_strHinban Like "特 F?B*-####MF*-*" Or IsMonster(in_strHinban) Then
        IsFkamachi = True
       
    'Caro
    ElseIf in_strHinban Like "F?B??*-####A*-*" Or in_strHinban Like "F?B??*-####B*-*" Or in_strHinban Like "F?B??*-####O*-*" Then
         IsFkamachi = True
    
    'Terrace(YG6型,YG5型)
    ElseIf in_strHinban Like "Y?B??*-####W*-*" Then
         IsFkamachi = True
         
    'G9型
    ElseIf IsG9(in_strHinban) Then
         IsFkamachi = True
         
    '格子型
    ElseIf IsKousi(in_strHinban) Then
         IsFkamachi = True
    End If
    
    '1.10.11 ADD エスパスライドウォール
    If in_strHinban Like "*PSW*-####FV*-*" Then
        IsFkamachi = True
    End If
    
End Function

Public Function IsKamachi(in_strHinban As Variant) As Boolean
'   *************************************************************
'   框確認
'
'   戻り値:Boolean
'       →True              框
'       →False             框以外
'
'    Input項目
'       in_strHinban        建具品番
'
'   1.10.9 201602** K.Asayama ADD
'           →框条件作成
'   1.10.11 20160302 K.Asayama ADD
'           →エスパリアラート除外
'   *************************************************************
    
    IsKamachi = False
    
    '1.10.9 ADD
    On Error GoTo Err_IsKamachi
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "??R*-####*-*" Or in_strHinban Like "特 ??R*-####*-*" Then
        '1.10.11 Change
            'IsKamachi = True
            If Not in_strHinban Like "HER*-####*-*" And Not in_strHinban Like "特 HER*-####*-*" Then
            
                IsKamachi = True
            End If
        '1.10.11 Change END
    End If
    
    Exit Function
    
Err_IsKamachi:
    IsKamachi = False
    '1.10.9 ADD END
End Function

Public Function IsThruGlass(in_strHinban As Variant) As Boolean
'   *************************************************************
'   スルーガラス確認
'   サブフォームの条件付書式からの呼び出しで消去した際不要な呼び出しが発生するのでエラーロジックを追加
'
'   戻り値:Boolean
'       →True              スルー
'       →False             スルー以外
'
'    Input項目
'       in_strHinban        建具品番
'
'   1.10.12 20160322 K.Asayama Change
'           →AF1～AF3を除外（F框へ)
'   1.10.19 K.Asayama Change
'           →1608より7型はFlush（ガラス）扱い
'   1.11.0
'           →テラスドア(YG6型)
'   2.5.2
'           →YG6型はスルーガラスから外す
'   *************************************************************
    On Error GoTo Err_IsThruGlass
    
    IsThruGlass = False
    
    If IsNull(in_strHinban) Then Exit Function
     
    If in_strHinban Like "*-####S*-*" Or in_strHinban Like "*-####C*-*" Or in_strHinban Like "*-####D*-*" _
        Or in_strHinban Like "F?C??*-####A*-*" Or in_strHinban Like "F?C??*-####B*-*" Or in_strHinban Like "F?C??*-####O*-*" _
        Or in_strHinban Like "*ME-####M*-*" Or in_strHinban Like "*SA-####M*-*" Or IsVertica(in_strHinban) Or in_strHinban Like "F?C??*-####MF*-*" Then
        
        IsThruGlass = True
    'YG6型
'    ElseIf in_strHinban Like "Y*-####T*-*" Then
'        IsThruGlass = True
        
    Else
        IsThruGlass = False
    End If
    
    Exit Function
    
Err_IsThruGlass:
    IsThruGlass = False
End Function

Public Function IsOyatobira(in_strHinban As Variant) As Boolean
'   *************************************************************
'   親扉確認
'
'   戻り値:Boolean
'       →True              親扉
'       →False             親扉以外
'
'    Input項目
'       in_strHinban        建具品番

'   1.10.6 K.Asayama 1610仕様（隠し丁番）追加
'   *************************************************************

    IsOyatobira = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "*DO-####*" Or in_strHinban Like "*DOS-####*" _
       Or in_strHinban Like "*CO-####*" Or in_strHinban Like "*COS-####*" _
        Or in_strHinban Like "*KO-####*" Or in_strHinban Like "*KOS-####*" _
                                                                            Then
        IsOyatobira = True
    Else
        IsOyatobira = False
    End If
    
End Function

Public Function IsKotobira(in_strHinban As Variant) As Boolean
'   *************************************************************
'   子扉確認
'
'   戻り値:Boolean
'       →True              子扉
'       →False             子扉以外
'
'    Input項目
'       in_strHinban        建具品番

'   1.10.6 K.Asayama 1610仕様（隠し丁番）追加
'   *************************************************************

    IsKotobira = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "*DK-####*" Or in_strHinban Like "*DKS-####*" _
            Or in_strHinban Like "*KK-####*" Or in_strHinban Like "*KKS-####*" _
                                                                                Then
        IsKotobira = True
    Else
        IsKotobira = False
    End If
    
End Function

Public Function IsSxL(in_strHinban As Variant, out_strKamiyahinban As Variant) As Boolean
'   *************************************************************
'   エスバイエル確認
'
'   戻り値:Boolean
'       →True              エスバイエル
'       →False             エスバイエル以外
'
'    Input項目
'       in_strHinban        建具品番
'    Output項目
'       out_strSxLhinban    神谷品番(Falseの場合はNull)

'   1.10.6 K.Asayama SxLコピー初回のみ実行に変更したため本処理に追加
'   *************************************************************
    
    Dim objLOCALdb As New cls_LOCALDB
    Dim strHinban As String
    Dim bolMentori As Boolean
    
    IsSxL = False
    
    On Error GoTo Err_IsSxL
    
    If IsNull(in_strHinban) Then GoTo Exit_IsSxL

    '1.10.6 K.Asayama ADD 20161211********
    If Not fncbolSxL_Replace() Then
        MsgBox "SxL品番マスタのコピーに失敗しました" & vbCrLf & "ネットワークに問題がある場合は回復後再度実行してください"
        Err.Raise 9999, , "Quit"
    End If
    '*************************************
    
    
    '下地で面取り記号がある場合は外す
    If in_strHinban Like "*①?②?③?④*" Then
        strHinban = left(in_strHinban, Len(in_strHinban) - 10)
        bolMentori = True
    Else
        strHinban = in_strHinban
        bolMentori = False
    End If
    '1.10.3 K.Asayama 20151119 SxL品番読替表ローカルテーブル名変更
    If objLOCALdb.ExecSelect("select ブランド品番 from WK_SxL品番読替表 where S×L品番 = '" & Trim(strHinban) & "'") Then
        If Not objLOCALdb.GetRS.EOF Then
            out_strKamiyahinban = objLOCALdb.GetRS![ブランド品番]
            If bolMentori Then
                out_strKamiyahinban = out_strKamiyahinban & right(in_strHinban, 10)
            End If
            IsSxL = True
        End If
    End If
    
    GoTo Exit_IsSxL
    
Err_IsSxL:
    IsSxL = False
    
Exit_IsSxL:
'クラスのインスタンスを破棄
    Set objLOCALdb = Nothing
End Function

Public Function valfncHinmei(in_objRemoteDB As cls_BRAND_MASTER, in_RS As ADODB.Recordset, in_strHinban As Variant, in_intSeihinkubun As Integer, in_strSpec As Variant) As Variant
'   *************************************************************
'   品名抽出
'   20151116 1.10.2 個別SpecをVariantに変更（Nullの可能性があるため）
'   戻り値:Variant → 品名（見つからない場合はNULL）
'
'    Input項目
'       in_objREMOTEDB      データベースサーバ
'       in_strHinban        建具品番（特注は外しておく）
'       in_intSeihinkubun   品番区分
'       in_strSpec          個別Spec
'   *************************************************************
    Dim strSQL As String
    
    strSQL = ""
    valfncHinmei = Null
    
    On Error GoTo Err_valfncHinmei
    
    If IsNull(in_strHinban) Then GoTo Exit_valfncHinmei
    
    '1.10.2 廃止***********************************
    'If in_strSpec = "" Then GoTo Exit_valfncHinmei
    '**********************************************
    
    Select Case in_intSeihinkubun
        Case 1, 5 '建具,ｸﾛｾﾞｯﾄ
            strSQL = "select top 1 品名 from T_建具品番ﾏｽﾀ where "
                If IsKotobira(in_strHinban) Then
                    strSQL = strSQL & " 子扉品番 = '" & in_strHinban & "'"
                Else
                    strSQL = strSQL & " 建具品番 = '" & in_strHinban & "'"
                End If
        Case 2, 4 '枠,三方枠
            strSQL = "select top 1 品名 from T_枠品番ﾏｽﾀ where 枠品番 = '" & in_strHinban & "'"
            
        Case 3 '下地枠
            strSQL = "select top 1 品名 from T_下地材品番ﾏｽﾀ where 下地材品番 = '" & in_strHinban & "'"
          
        Case 6 '造作材
            strSQL = "select top 1 品名 from T_造作材品番ﾏｽﾀ where 造作材品番 = '" & in_strHinban & "'"
            
        Case 7 '玄関収納
            strSQL = "select top 1 品名 from T_玄関収納ﾏｽﾀ where 品番 = '" & in_strHinban & "'"
            
        Case 8 '金物
            strSQL = "select top 1 品名 from T_金物品番ﾏｽﾀ where 金物品番 = '" & in_strHinban & "'"
        
    End Select
    
    If strSQL = "" Then
        GoTo Exit_valfncHinmei
    Else
        '1.10.2 ****************************************************************************************************************
        'strSQL = strSQL & " and 仕様 = '" & left(in_strSpec, 3) & "' and '" & right(in_strSpec, 4) & "' between 開始 and 終了 "
        If Not IsNull(in_strSpec) And in_strSpec <> "" Then
            strSQL = strSQL & " and 仕様 = '" & left(in_strSpec, 3) & "' and '" & right(in_strSpec, 4) & "' between 開始 and 終了 "
        End If
        '***********************************************************************************************************************
    End If
    
    
    If in_objRemoteDB.ExecSelect_ExternalRS(in_RS, strSQL) Then
        If Not in_RS.EOF Then
            valfncHinmei = in_RS![品名]
        End If
    End If
    
    GoTo Exit_valfncHinmei
    
Err_valfncHinmei:
    'MsgBox Err.Description
Exit_valfncHinmei:

End Function

Public Function bolFncTokuHinban(in_varHinban As Variant, in_varTokuHinban As Variant, ByRef out_varTokuhinban As Variant) As Boolean
'   *************************************************************
'   特注品番確認
'   品番が特注品番か確認し特注品番の場合は通常品番を返す
'   SxL品番に該当する場合神谷品番を返す
'
'   戻り値:Boolean
'       →True              特注
'       →False             特注以外
'
'    Input項目
'       in_varHinban        受注品番
'       in_varTokuHinban    特注受注品番
'       out_varTokuhinban   受注品番（特注の場合--頭の「特 」を外したもの）
'   *************************************************************
    
    Dim varHinban As Variant
    
    out_varTokuhinban = Null
    
    If Not IsNull(in_varHinban) And in_varHinban <> "" Then
        Select Case in_varHinban
            Case "TATEGU", "TATEGUM", "TATEGUW", "TATEGUS", "BUZAI", "BUZAIK", "BUZAIG"
                varHinban = Mid(in_varTokuHinban, 3)
                bolFncTokuHinban = True
            Case Else
                varHinban = in_varHinban
        End Select
        
        'SxL品番チェック
        If Not IsSxL(varHinban, out_varTokuhinban) Then
            out_varTokuhinban = varHinban
        End If
        
    End If
End Function

Public Function intFncChecklistCode(in_Kubun As String) As Integer
'   *************************************************************
'   チェックリストの区分取得
'   コードはローカルルール
'
'   戻り値:Integer
'       →                  チェックリスト用コード
'                           扉;子扉;枠;造;下地;金;玄;床;階;フ
'    Input項目
'       in_Kubun            T_チェックリストの区分
'                          該当する区分が無い場合は0を返す
'   *************************************************************

    Select Case in_Kubun
        Case "扉"
            intFncChecklistCode = 1
        Case "子扉"
            intFncChecklistCode = 1
        Case "枠"
            intFncChecklistCode = 2
        Case "下地"
            intFncChecklistCode = 3
        Case "造"
            intFncChecklistCode = 6
        Case "金"
            intFncChecklistCode = 8
        Case "玄"
            intFncChecklistCode = 7
        Case "床"
            intFncChecklistCode = 10
        Case "階"
            intFncChecklistCode = 11
        Case "フ"
            intFncChecklistCode = 12
        Case Else
            intFncChecklistCode = 0
    End Select

End Function

Public Function intFncSeizoKubunToChecklistCode(in_Kubun As Integer) As Integer
'   *************************************************************
'   製造区分に対応するチェックリストのコード取得
'
'   戻り値:Integer
'       →                  チェックリスト用コード
'                           扉;子扉;枠;造;下地;金;玄;床;階;フ
'    Input項目
'       in_Kubun           製造区分
'                          該当する区分が無い場合は0を返す
'   *************************************************************

    '製造区分に対応するチェックリストのコードを返す

    Select Case in_Kubun
        Case 1, 2, 3, 5 'Flush,F框,框,ｸﾛｾﾞｯﾄ
            intFncSeizoKubunToChecklistCode = 1
        Case 4, 5 '枠と三方枠
            intFncSeizoKubunToChecklistCode = 2
        Case 6, 7 '下地とステルス
            intFncSeizoKubunToChecklistCode = 3
        Case Else 'その他は無視
            intFncSeizoKubunToChecklistCode = 0
    End Select

End Function

Public Function intFncSeihinKubunToChecklistCode(in_intSeihinkubun As Integer) As Integer
'   *************************************************************
'   製品区分に対応するチェックリストのコード取得
'
'   戻り値:Integer
'       →                  チェックリスト用コード
'                           扉;子扉;枠;造;下地;金;玄;床;階;フ
'    Input項目
'       in_Kubun           製品区分
'                          該当する区分が無い場合は0を返す
'   *************************************************************
    
    intFncSeihinKubunToChecklistCode = 0
    
    Select Case in_intSeihinkubun
    
        Case 5  '建具、ｸﾛｾﾞｯﾄ

            intFncSeihinKubunToChecklistCode = 1
            
        Case Else
        
            intFncSeihinKubunToChecklistCode = in_intSeihinkubun
           
    End Select
    
End Function

Public Function isCaro(in_varHinban As Variant) As Boolean
'   *************************************************************
'   Caro確認
'
'   戻り値:Boolean
'       →True              Caro
'       →False             Caro以外
'
'    Input項目
'       in_strHinban        建具品番

'   1.10.6 K.Asayama 1610仕様（AF1～AF3）追加
'   1.10.19 K.Asayama Change
'           →ワイルドカード誤り訂正(_→?)
'   *************************************************************

    isCaro = False
    
    If in_varHinban Like "F?C*-####A*-*" Or in_varHinban Like "F?C*-####B*-*" Or in_varHinban Like "F?C*-####O*-*" _
        Or in_varHinban Like "特 F?C*-####A*-*" Or in_varHinban Like "特 F?C*-####B*-*" Or in_varHinban Like "特 F?C*-####O*-*" _
            Or in_varHinban Like "F?B*-####A*-*" Or in_varHinban Like "F?B*-####B*-*" Or in_varHinban Like "F?B*-####O*-*" _
                Or in_varHinban Like "特 F?B*-####A*-*" Or in_varHinban Like "特 F?B*-####B*-*" Or in_varHinban Like "特 F?B*-####O*-*" _
                                                                                                                                        Then
        
        isCaro = True
        
    End If
    
End Function

Public Function strfncDaibunrui_Kamui(in_strShurui As String, in_varHinban As Variant) As String
'   *************************************************************
'   種類からカムイの大分類を取得
'
'   戻り値:String
'       →                  カムイの大分類
'                           該当する区分が無い場合は"00"を返す

'    Input項目
'       in_strShurui        種類
'       in_varHinban        品番
'   *************************************************************
'
'三方枠のみ品番が必要
    
    strfncDaibunrui_Kamui = "00"
    
    Select Case in_strShurui
    
        Case "建具", "子扉"

            strfncDaibunrui_Kamui = "11"
            
        Case "ｸﾛｾﾞｯﾄ"

            strfncDaibunrui_Kamui = "21"
            
        Case "枠"
        
            If IsCasing(in_varHinban) Then
                strfncDaibunrui_Kamui = "03"
            Else
                strfncDaibunrui_Kamui = "02"
            End If

            
        Case "下地"
        
            strfncDaibunrui_Kamui = "01"
            
        Case "造作材"
        
            strfncDaibunrui_Kamui = "41"
            
        Case "玄関収納"
        
            strfncDaibunrui_Kamui = "61"
            
        Case "金物"
        
            strfncDaibunrui_Kamui = "51"
            
        Case "配送費"
        
            
        Case "床材"


        Case "階段"


        Case "ﾌｧﾆﾁｭｱ"
    
    End Select
    
End Function

Public Function strfncSyobunrui_Kamui(in_strDaibunrui_Kamui As String, in_varHinban As Variant) As String
'   *************************************************************
'   カムイの大分類と品番からをカムイ小分類を取得
'
'   戻り値:String
'       →                              カムイの小分類
'
'    Input項目
'       in_strDaibunrui_Kamui           カムイの大分類
'       in_varHinban                    品番

'1.11.0
'       →分類変更に対応(一部関数化）
'1.11.3
'       →分類変更に対応
'   *************************************************************

    Dim strHinbanKigou As String
    
    Select Case in_strDaibunrui_Kamui
    
        Case "01" '下地
            strHinbanKigou = left(in_varHinban, 1)
            
            Select Case strHinbanKigou
                Case "S", "N", "A", "B"
                    strfncSyobunrui_Kamui = strHinbanKigou
                    
                Case Else
                    strfncSyobunrui_Kamui = "W"
                    
            End Select
        Case "02" '枠
            strfncSyobunrui_Kamui = "W"
            
        Case "03" '三方枠
            strfncSyobunrui_Kamui = "C"
            
        Case "11" '出入口
            strHinbanKigou = left(in_varHinban, 1)
            
            '関数化
'            If in_varHinban Like "F_V*-####*" Then 'Vertica
'                strfncSyobunrui_Kamui = "V"
'
'            ElseIf in_varHinban Like "F_C*-####*" Then 'Caro
'                strfncSyobunrui_Kamui = "A"

            If IsVertica(in_varHinban) Then  'Vertica
                strfncSyobunrui_Kamui = "V"

            ElseIf isCaro(in_varHinban) Then 'Caro
                strfncSyobunrui_Kamui = "A"
                
            Else
                Select Case strHinbanKigou
                    Case "F" '標準品はCUBEのコードを送る（分割されたら分ける必要あり）
                        strfncSyobunrui_Kamui = "C"
                    Case "S" 'F/S
                        strfncSyobunrui_Kamui = "K"
                    Case "A" 'Air
                        strfncSyobunrui_Kamui = "F"
                    Case Else
                        strfncSyobunrui_Kamui = strHinbanKigou
                End Select
            End If
            
        Case "21" 'クロゼット
            strfncSyobunrui_Kamui = "M"
        
        Case "31" 'ウォークスルー
            If in_varHinban Like "*-####G*" Then        'ガラス
                strfncSyobunrui_Kamui = "G"
            ElseIf in_varHinban Like "*-####L*" Then    'ルーバー
                strfncSyobunrui_Kamui = "L"
            Else
                strfncSyobunrui_Kamui = "C"             'コンビ
            End If
            
        Case "41" '造作材
            strfncSyobunrui_Kamui = "99999" '表示しない
            
        Case "51" '金物
            strfncSyobunrui_Kamui = 1
            
        Case "61" '玄関収納
            strfncSyobunrui_Kamui = 1
            
    End Select


End Function

Public Function IsGikan(in_strHinban As Variant) As Boolean
'   *************************************************************
'   技官製造確認
'   サブフォームの条件付書式からの呼び出しで消去した際不要な呼び出しが発生するのでエラーロジックを追加
'   'ADD by Asayama 20150903
'   戻り値:Boolean
'       →True              技官製造
'       →False             技官製造以外
'
'    Input項目
'       in_strHinban        建具品番

'   *************************************************************
    On Error GoTo Err_IsGikan
    
    IsGikan = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    'スルーガラス
    If IsThruGlass(in_strHinban) Then
        IsGikan = True
    
    '引き手レス（Vertica）
    ElseIf IsVertica(in_strHinban) Then
        IsGikan = True
    
    'Air
    ElseIf IsAir(in_strHinban) Then
        IsGikan = True
     
    Else
        IsGikan = False
        
    End If
    
    Exit Function
    
Err_IsGikan:
    IsGikan = False
End Function

Public Function IsVertica(in_strHinban As Variant) As Boolean
'   *************************************************************
'   引き手レス引戸確認
'   サブフォームの条件付書式からの呼び出しで消去した際不要な呼び出しが発生するのでエラーロジックを追加
'   'ADD by Asayama 20150903
'   戻り値:Boolean
'       →True              引き手レス
'       →False             引き手レス以外
'
'    Input項目
'       in_strHinban        建具品番

'   *************************************************************
    On Error GoTo Err_IsVertica
    
    IsVertica = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "??V*-####*-*" Or in_strHinban Like "特 ??V*-####*-*" Then
        IsVertica = True
    Else
        IsVertica = False
    End If
    
    Exit Function
    
Err_IsVertica:
    IsVertica = False
    
End Function

Public Function IsAir(in_strHinban As Variant) As Boolean
'   *************************************************************
'   FullHeight Air確認
'   サブフォームの条件付書式からの呼び出しで消去した際不要な呼び出しが発生するのでエラーロジックを追加
'   'ADD by Asayama 20150903
'   戻り値:Boolean
'       →True              Air
'       →False             Air以外
'
'    Input項目
'       in_strHinban        建具品番

'   *************************************************************
    On Error GoTo Err_IsAir
    
    IsAir = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "A*-####SC*-*" Or in_strHinban Like "A*-####SL*-*" Or in_strHinban Like "特 A*-####SC*-*" Or in_strHinban Like "特 A*-####SL*-*" Then
        IsAir = True
    Else
        IsAir = False
    End If
    
    Exit Function
    
Err_IsAir:
    IsAir = False
    
End Function

Public Function IsPainted(in_strHinban As Variant) As Boolean
'   *************************************************************
'   塗装扉確認
'   サブフォームの条件付書式からの呼び出しで消去した際不要な呼び出しが発生するのでエラーロジックを追加
'   'ADD by Asayama 201510**
'   '1.10.4 Change by Asayama 20151207
'       →全面改訂（リアラートに無塗装ができるので色コードベースに変更）
'
'   戻り値:Boolean
'       →True              塗装扉
'       →False             塗装扉以外
'
'    Input項目
'       in_strHinban        建具品番

'   1.10.11 K.Asayama ADD
'           →エスパのリアラートは塗装
'   1.12.3
'           →リアラート新色追加
'   2.1.1
'           →1801新色先行追加
'   *************************************************************
    On Error GoTo Err_IsPainted
    
    IsPainted = False
    
    If IsNull(in_strHinban) Then Exit Function
    
'    If in_strHinban Like "R*-####*-*" Or in_strHinban Like "特 R*-####*-*" Or in_strHinban Like "B*-####*-*" Or in_strHinban Like "特 B*-####*-*" Then
'        IsPainted = True
'    Else
'        IsPainted = False
'    End If
    
    If in_strHinban Like "*-####*-*(NW)*" Or in_strHinban Like "*-####*-*(NO)*" Or in_strHinban Like "*-####*-*(NC)*" Or in_strHinban Like "*-####*-*(NK)*" Or in_strHinban Like "*-####*-*(NA)*" Or in_strHinban Like "*-####*-*(NB)*" Then
        IsPainted = True
    Else
        IsPainted = False
    End If
    
    '1.12.3 ADD
    If in_strHinban Like "*-####*-*(NH)*" Then
        IsPainted = True
    End If
    
    If in_strHinban Like "*-####*-*(NT)*" Or in_strHinban Like "*-####*-*(NY)*" Then
        IsPainted = True
    End If
    
    '1.10.11 ADD
    If in_strHinban Like "*HER*-####*-*" Then
        IsPainted = True
    End If
    
    Exit Function
    
Err_IsPainted:
    IsPainted = False
    
End Function

Public Function IsMonster(in_strHinban As Variant) As Boolean
'   *************************************************************
'   モンスター扉確認
'   サブフォームの条件付書式からの呼び出しで消去した際不要な呼び出しが発生するのでエラーロジックを追加
'   'ADD by Asayama 201510**
'   戻り値:Boolean
'       →True              モンスター扉確認
'       →False             モンスター扉確認以外
'
'    Input項目
'       in_strHinban        建具品番

'   *************************************************************
    On Error GoTo Err_IsMonster
    
    IsMonster = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "O*-####*-*" Or in_strHinban Like "特 O*-####*-*" Then
        IsMonster = True
    Else
        IsMonster = False
    End If
    
    Exit Function
    
Err_IsMonster:
    IsMonster = False
    
End Function

Public Function IsStealth_Seizo_TEMP(in_strHinban As Variant) As Boolean
'   *************************************************************
'   ステルス（製造）確認（上記 IsStealth_Seizo使用開始時には差し替え）
'
'   戻り値:Boolean
'       →True              ステルス（製造）
'       →False             ステルス（製造）以外
'
'    Input項目
'       in_strHinban        下地品番

'1.10.9 K.Asayama
'       →特注開閉様式DVはインセット下地
'1.10.13 K.Asayama
'       →エスパ下地品番はインセット下地
'1.11.4 K.Asayama
'       →1701新品番追加(VM)
'2.9.0
'       →1808新品番追加(GU)
'   *************************************************************
    '
    IsStealth_Seizo_TEMP = False
    
    '1.10.13
    If in_strHinban Like "*PW*-####*" Then
        IsStealth_Seizo_TEMP = False
        Exit Function
    End If
    
    If (in_strHinban Like "*SG*-####*" Or in_strHinban Like "*NG*-####*" Or in_strHinban Like "*AG*-####*" Or in_strHinban Like "*BG*-####*") _
        And Not in_strHinban Like "*ML-####*" And Not in_strHinban Like "*MK-####*" And Not in_strHinban Like "*MT-####*" And Not in_strHinban Like "*DU-####*" And Not in_strHinban Like "*DN-####*" And Not in_strHinban Like "*VN-####*" And Not in_strHinban Like "*CTSG*MK-####*" And Not in_strHinban Like "*CTSG*ML-####*" And Not in_strHinban Like "*CTSG*MT-####*" And Not in_strHinban Like "*KU-####*" And Not in_strHinban Like "*KN-####*" And Not in_strHinban Like "*DV-####*" And Not in_strHinban Like "*GU-####*" Then
        IsStealth_Seizo_TEMP = True
    End If
    
End Function

Public Function fncbolSxL_Replace() As Boolean
'   *************************************************************
'   SxL品番読替表置換え処理
'   1.10.3 K.Asayama ADD 20151119 SxL品番表リモートからコピー
'   1.10.6 K.Asayama ADD 20151211 コピー済みの場合(bolSxLCopy=True）は処理しない
'
'   リモートデータベースからローカルにSxL品番読替表をコピーする
'
'   戻り値:Boolean
'       →True              置換成功
'       →False             置換失敗
'
'   *************************************************************

    fncbolSxL_Replace = False
    
    If bolSxLCopy Then
        fncbolSxL_Replace = True
        Exit Function
    End If
    
    Dim objREMOTEdb As New cls_BRAND_MASTER
    Dim objLOCALdb As New cls_LOCALDB
    
    On Error GoTo Err_fncbolSxL_Replace
    
    Dim strSQL_Insert As String
    Dim strSQL As String
    strSQL_Insert = "Insert into WK_SxL品番読替表(S×L品番,ブランド品番,DH,DW,CH) values ("
    
    '工場用コピー（T_Calendar_工場)
    If objLOCALdb.ExecSQL("delete from WK_SxL品番読替表") Then
        strSQL = "select distinct [S×L品番],ブランド品番,DW,DH,CH from SxL品番読替表 "
        If objREMOTEdb.ExecSelect(strSQL) Then
            Do While Not objREMOTEdb.GetRS.EOF
                If Not objLOCALdb.ExecSQL(strSQL_Insert & "'" & objREMOTEdb.GetRS![S×L品番] & "','" & objREMOTEdb.GetRS![ブランド品番] & "'," & objREMOTEdb.GetRS![DW] & "," & objREMOTEdb.GetRS![DH] & "," & objREMOTEdb.GetRS![CH] & ")") Then
                    Err.Raise 9999, , "SxL品番読替表 ローカルコピーエラー"
                End If
                objREMOTEdb.GetRS.MoveNext
            Loop
        End If
    End If
    
    '1.10.6 K.Asayama ADD 20151211 コピー完了の場合共通フラグをTrueにする
    bolSxLCopy = True
    
    fncbolSxL_Replace = True
    
    GoTo Exit_fncbolSxL_Replace
    
Err_fncbolSxL_Replace:
    MsgBox Err.Description
    
Exit_fncbolSxL_Replace:

    Set objREMOTEdb = Nothing
    Set objLOCALdb = Nothing
    
End Function

Public Function IsREALART(in_strHinban As Variant) As Boolean
'   *************************************************************
'   REALART確認
'   サブフォームの条件付書式からの呼び出しで消去した際不要な呼び出しが発生するのでエラーロジックを追加
'   '1.10.4 ADD by Asayama 20151207
'   戻り値:Boolean
'       →True              REALART
'       →False             REALART以外
'
'    Input項目
'       in_strHinban        建具品番

'2.3.0（コメントのみ挿入）
'   →1801仕様でリアラートとタモにシリーズ分割されたがシステムの内容上は一緒の方が
'     都合がよいのでこの関数はタモを含むようにする
'   *************************************************************
    On Error GoTo Err_IsREALART
    
    IsREALART = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "R*-####*-*" Or in_strHinban Like "特 R*-####*-*" Then
        IsREALART = True
    Else
        IsREALART = False
    End If
    
    Exit Function
    
Err_IsREALART:
    IsREALART = False
    
End Function

Public Function IsPALIO(in_strHinban As Variant) As Boolean
'   *************************************************************
'   PALIO確認
'   サブフォームの条件付書式からの呼び出しで消去した際不要な呼び出しが発生するのでエラーロジックを追加
'   '1.10.4 ADD by Asayama 20151207
'   戻り値:Boolean
'       →True              PALIO
'       →False             PALIO以外
'
'    Input項目
'       in_strHinban        建具品番

'   *************************************************************
    On Error GoTo Err_IsPALIO
    
    IsPALIO = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "B*-####*-*" Or in_strHinban Like "特 B*-####*-*" Then
        IsPALIO = True
    Else
        IsPALIO = False
    End If
    
    Exit Function
    
Err_IsPALIO:
    IsPALIO = False
    
End Function

Public Function fncvalDoorColor(inHinban As String) As Variant
'   *************************************************************
'   色確認
'   品番から色を返す。返せない場合は空欄を返す（Nullではない）
'   '1.10.7 ADD by Asayama 20160108
'   戻り値:Variant
'       →色コード（色コードが無い場合は空欄、エラーの場合はNull）
'
'    Input項目
'       inHinban            建具品番

'   *************************************************************
    Dim i As Integer
    Dim byteHirakiIchi As Byte
    Dim byteTojiIchi As Byte
    
    On Error GoTo Err_fncvalDoorColor
    
    fncvalDoorColor = Null
    
    byteTojiIchi = 0
    byteHirakiIchi = 0
    
    For i = Len(inHinban) To 1 Step -1
        If Mid(inHinban, i, 1) = ")" Then
            byteTojiIchi = i
        ElseIf Mid(inHinban, i, 1) = "(" Then
            byteHirakiIchi = i + 1
        End If
        
        If byteHirakiIchi <> 0 And byteTojiIchi <> 0 Then Exit For
        
    Next
        
    If byteHirakiIchi <> 0 And byteTojiIchi <> 0 And byteTojiIchi > byteHirakiIchi Then
        fncvalDoorColor = Mid(inHinban, byteHirakiIchi, byteTojiIchi - byteHirakiIchi)
    End If
    
    GoTo Exit_fncvalDoorColor
    
Err_fncvalDoorColor:
    fncvalDoorColor = Null
    MsgBox Err.Description, , "品番から色コードが取得できません"
    
Exit_fncvalDoorColor:
    
End Function

Public Function fncIntHalfGlassMirror_Maisu(in_strHinban As Variant, in_Maisu As Integer) As Integer
'   *************************************************************
'   複数枚で片側のみガラス・ミラーの品番確認し、ガラス枚数を返す
'   サブフォームの条件付書式からの呼び出しで消去した際不要な呼び出しが発生するのでエラーロジックを追加
'1.10.10 ADD by Asayama
'   戻り値:Integer
'       →ガラス扉枚数
'
'    Input項目
'       in_strHinban        建具品番
'        in_Maisu 建具枚数
'   *************************************************************
    On Error GoTo Err_fncIntHalfGlassMirror_Maisu
    
    fncIntHalfGlassMirror_Maisu = in_Maisu
    
    If IsNull(in_strHinban) Then Exit Function
    
    '2で割り切れない場合そのまま返す
    If in_Maisu Mod 2 <> 0 Then Exit Function
    
    If in_strHinban Like "*ME-####MR*-*" Or in_strHinban Like "*ME-####ML*-*" Then
        
        fncIntHalfGlassMirror_Maisu = in_Maisu / 2
    End If
    
    Exit Function
    
Err_fncIntHalfGlassMirror_Maisu:
    fncIntHalfGlassMirror_Maisu = in_Maisu
End Function

Public Function IsGranArt(in_strHinban As Variant) As Boolean
'   *************************************************************
'   グランアート確認
'   サブフォームの条件付書式からの呼び出しで消去した際不要な呼び出しが発生するのでエラーロジックを追加
'   '1.10.16 ADD
'
'   戻り値:Boolean
'       →True              グランアート
'       →False             グランアート以外
'
'    Input項目
'       in_strHinban        建具品番

'   *************************************************************
    On Error GoTo Err_IsGranArt
    
    IsGranArt = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "G*-####*-*" Or in_strHinban Like "特 G*-####*-*" Then
        IsGranArt = True
    Else
        IsGranArt = False
    End If
    
    Exit Function
    
Err_IsGranArt:
    IsGranArt = False
    
End Function
Public Function IsInset(in_strWakuHinban As Variant) As Boolean
'   *************************************************************
'   インセット枠確認
'   '1.10.16 ADD
'
'   戻り値:Boolean
'       →True              インセット枠
'       →False             インセット枠以外
'
'    Input項目
'       in_strHinban        枠品番

'1.11.1 Change K70品番がFalseになってしまう件対応
'   *************************************************************
    On Error GoTo Err_IsInset
    
    IsInset = False

    If in_strWakuHinban Like "K##*-####*" Or in_strWakuHinban Like "特 K##*-####*" Then
        IsInset = True
    End If
    
    Exit Function

Err_IsInset:
    IsInset = False
End Function
Public Function IsHirakido(in_strHinban As Variant) As Boolean
'   *************************************************************
'   開き戸確認（親子含む）
'   '1.10.16 ADD
'
'   戻り値:Boolean
'       →True              開き戸
'       →False             開き戸以外
'
'    Input項目
'       in_strHinban        建具（枠、下地）品番
'   1.10.19 K.Asayama Change
'           →隠し丁番親子追加
'   *************************************************************
    
    On Error GoTo Err_IsHirakido
    
    If in_strHinban Like "*CA-####*" Or in_strHinban Like "*CAS-####*" _
        Or in_strHinban Like "*DA-####*" Or in_strHinban Like "*DAS-####*" _
        Or in_strHinban Like "*PA-####*" Or in_strHinban Like "*PAS-####*" _
        Or in_strHinban Like "*KA-####*" Or in_strHinban Like "*KAS-####*" _
        Or in_strHinban Like "*DO-####*" Or in_strHinban Like "*DOS-####*" _
        Or in_strHinban Like "*DK-####*" Or in_strHinban Like "*DKS-####*" _
        Or in_strHinban Like "*KO-####*" Or in_strHinban Like "*KOS-####*" _
        Or in_strHinban Like "*KK-####*" Or in_strHinban Like "*KKS-####*" _
    Then
        
        IsHirakido = True
        
    Else
    
        IsHirakido = False
        
    End If
    
    Exit Function
    
Err_IsHirakido:
    IsHirakido = False
End Function

Public Function IsWallThru(in_strHinban As Variant) As Boolean
'   *************************************************************
'   ウォールスルー確認
'   1.11.0 ADD
'
'   戻り値:Boolean
'       →True              WallThrough
'       →False             WallThrough以外
'
'    Input項目
'       in_strHinban        下地品番

'   *************************************************************
    '
    IsWallThru = False

    If in_strHinban Like "*WS*-####*" Then
        IsWallThru = True
        Exit Function
    End If

    
End Function

Public Function IsTerrace(in_varHinban As Variant) As Boolean
'   *************************************************************
'   テラスドア確認
'
'   戻り値:Boolean
'       →True              Terrace
'       →False             Terrace以外
'
'    Input項目
'       in_strHinban        建具品番

'   1.11.0 ADD
'   *************************************************************

    IsTerrace = False
    
    On Error GoTo Err_IsTerrace
        
    If IsNull(in_varHinban) Then
        Exit Function
    End If
    
    If in_varHinban Like "Y*-####*-*" Then
        
        IsTerrace = True
        
    End If
    
    Exit Function
    
Err_IsTerrace:
    IsTerrace = False
    
End Function

Public Function IsMirror(in_varHinban As Variant) As Boolean
'   *************************************************************
'   ミラー扉確認
'
'   戻り値:Boolean
'       →True              ミラー
'       →False             ミラー以外
'
'    Input項目
'       in_strHinban        建具品番

'   1.11.0 ADD
'   *************************************************************

    IsMirror = False
    
    On Error GoTo Err_IsMirror
        
    If IsNull(in_varHinban) Then
        Exit Function
    End If
    
    If in_varHinban Like "*-####M*-*" Then
        
        IsMirror = True
        
    End If
    
    Exit Function
    
Err_IsMirror:
    IsMirror = False
End Function

Public Function IsCloset_Hiraki(in_varHinban As Variant) As Boolean
'   *************************************************************
'   クロゼット品番確認（開き戸）

'   ※両／片開き（下地枠兼用） スライド収納は対象としない

'   戻り値:Boolean
'       →True              クロゼット開き
'       →False             クロゼット開き以外
'
'    Input項目
'       in_varHinban        建具品番／下地品番

'   1.12.0 ADD
'   *************************************************************
    On Error GoTo Err_IsCloset_Hiraki
    
    IsCloset_Hiraki = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*MA-####*" Or in_varHinban Like "*MB-####*" Or in_varHinban Like "*MAS-####*" Or in_varHinban Like "*MBS-####*" Then
        IsCloset_Hiraki = True
    End If
    
    Exit Function

Err_IsCloset_Hiraki:
    IsCloset_Hiraki = False
End Function

Public Function IsCloset_Oredo(in_varHinban As Variant) As Boolean
'   *************************************************************
'   クロゼット品番確認（折れ戸）


'   戻り値:Boolean
'       →True              クロゼット折れ戸
'       →False             クロゼット折れ戸以外
'
'    Input項目
'       in_varHinban        建具品番／下地品番

'   1.12.0 ADD
'   *************************************************************
    On Error GoTo Err_IsCloset_Oredo
    
    IsCloset_Oredo = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*ML-####*" Or in_varHinban Like "*MK-####*" Or in_varHinban Like "*MT-####*" Then
        IsCloset_Oredo = True
    End If
    
    Exit Function

Err_IsCloset_Oredo:
    IsCloset_Oredo = False
End Function

Public Function IsSoftMotion(ByVal in_varHinban As Variant) As Boolean
'   *************************************************************
'   ソフトモーション確認

'   戻り値:Boolean
'       →True              ソフトモーション無し
'       →False             ソフトモーション以外
'
'    Input項目
'       in_varHinban        建具品番／下地品番

'   1.12.0 ADD
'   *************************************************************
    IsSoftMotion = False
    
    If in_varHinban Like "*CA-####*" Or in_varHinban Like "*CO-####*" Or in_varHinban Like "*CAS-####*" Or in_varHinban Like "*COS-####*" Then
    
        IsSoftMotion = True
    
    End If
    

End Function

Public Function IsCloset_Slide(in_varHinban As Variant) As Boolean
'   *************************************************************
'   スライド収納確認

'   戻り値:Boolean
'       →True              スライド収納
'       →False             スライド収納以外
'
'    Input項目
'       in_strHinban        建具品番,又は下地品番


'   1.12.0 ADD
'   *************************************************************
    
    Dim strHinban As String
    
    On Error GoTo Err_IsCloset_Slide
    
    IsCloset_Slide = False
    
    
    If IsNull(in_varHinban) Then Exit Function

    If in_varHinban Like "*SA-####*" Then
        IsCloset_Slide = True
    End If
    
    Exit Function

Err_IsCloset_Slide:
    IsCloset_Slide = False
    
End Function

Public Function IsYukazukeRail(in_varHinban As Variant) As Boolean
'   *************************************************************
'   床付けレール品番確認

'   ※上吊り連動は含まない

'   戻り値:Boolean
'       →True              床付けレール
'       →False             床付けレール以外
'
'    Input項目
'       in_varHinban        品番

'   1.12.0 ADD
'   *************************************************************
    On Error GoTo Err_IsYukazukeRail
    
    IsYukazukeRail = False
    
    If in_varHinban Like "*DM-####*" Or in_varHinban Like "*DL-####*" Or in_varHinban Like "*DN-####*" Then
        IsYukazukeRail = True
    'Vレール
    ElseIf in_varHinban Like "*VM-####*" Or in_varHinban Like "*VL-####*" Or in_varHinban Like "*VN-####*" Then
        IsYukazukeRail = True
    End If
    
    Exit Function

Err_IsYukazukeRail:
    IsYukazukeRail = False
End Function

Public Function IsLUCENTE(in_varHinban As Variant) As Boolean
'   *************************************************************
'   ルチェンテ確認

'   戻り値:Boolean
'       →True              ルチェンテ
'       →False             ルチェンテ以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.1.0 ADD
'   *************************************************************

    Dim strHinban As String
    
    IsLUCENTE = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "特 ", "")
    
    If strHinban Like "P*-####*-*" Then
        If strHinban Like "*(XW)" Or strHinban Like "*(XB)" Then
            IsLUCENTE = True
        End If
    End If
    
End Function

Public Function IsSINA(in_varHinban As Variant) As Boolean
'   *************************************************************
'   シナ確認
'   'ADD by Asayama 20150903
'   戻り値:Boolean
'       →True              シナ品番
'       →False             シナ品番以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.1.0 ADD
'   *************************************************************

    Dim strHinban As String
    
    IsSINA = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "特 ", "")
    
    If strHinban Like "T*-####*-*" Then
        If IsSINAColor(strHinban) Then
            IsSINA = True
        End If
    End If
    
End Function

Public Function IsSINAColor(in_varHinban As Variant) As Boolean
'   *************************************************************
'   シナ色確認
'
'   戻り値:Boolean
'       →True              色がシナ色
'       →False             シナ色以外
'
'    Input項目
'       in_strHinban        建具品番

'   2.1.0 ADD
'   *************************************************************

    IsSINAColor = False
    
    If IsNull(in_varHinban) Then Exit Function

    If in_varHinban Like "*-*-*(ZZ)" Or in_varHinban Like "*-*-*(AA)" Or in_varHinban Like "*-*-*(BB)" Or in_varHinban Like "*-*-*(CC)" Or in_varHinban Like "*-*-*(DD)" Then
        IsSINAColor = True
    Else
        IsSINAColor = False
    End If
    
End Function

Public Function IsFs(in_varHinban As Variant) As Boolean
'   *************************************************************
'   F/S確認

'   戻り値:Boolean
'       →True              F/S
'       →False             F/S以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.1.0 ADD
'   *************************************************************

    Dim strHinban As String
    
    IsFs = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "特 ", "")
    
    If strHinban Like "S*-####*-*" Then
        IsFs = True
    End If
    
End Function

Public Function IsCloset_Hikichigai(in_varHinban As Variant) As Boolean
'   *************************************************************
'   物入れ引き違い確認

'   戻り値:Boolean
'       →True              物入れ引き違い
'       →False             物入れ引き違い以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.1.0 ADD
'   *************************************************************
    
    IsCloset_Hikichigai = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*ME-####*-*" Then
        IsCloset_Hikichigai = True
    End If
    
End Function

Public Function IsSideThru(in_varHinban As Variant) As Boolean
'   *************************************************************
'   サイドスルー確認

'   戻り値:Boolean
'       →True              サイドスルー
'       →False             サイドスルー以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.1.0 ADD

'2.3.0
'   →1801仕様追加
'   *************************************************************
    
    IsSideThru = False
    
    If IsNull(in_varHinban) Then Exit Function

    If in_varHinban Like "*-####ST*-*" Or in_varHinban Like "*-####SS*-*" Or in_varHinban Like "*-####SG*-*" Or in_varHinban Like "*-####SH*-*" Then
        IsSideThru = True
    End If
    
End Function

Public Function IsCenterThru(in_varHinban As Variant) As Boolean
'   *************************************************************
'   センタースルー確認

'   戻り値:Boolean
'       →True              センタースルー
'       →False             センタースルー以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.1.0 ADD
'   *************************************************************
    
    IsCenterThru = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*-####C*-*" Then
        IsCenterThru = True
    End If
    
End Function

Public Function IsWideThru(in_varHinban As Variant) As Boolean
'   *************************************************************
'   幅広スルーガラス確認

'   戻り値:Boolean
'       →True              幅広センタースルー
'       →False             幅広センタースルー以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.1.0 ADD
'   *************************************************************
    
    IsWideThru = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*-####D*-*" Then
        IsWideThru = True
    End If
    
End Function

Public Function IsG7_Flush(in_varHinban As Variant) As Boolean
'   *************************************************************
'   G7型(1608仕様以降)確認

'   戻り値:Boolean
'       →True              G7型
'       →False             G7型以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.1.0 ADD
'   *************************************************************

    Dim strHinban As String
    
    IsG7_Flush = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "特 ", "")
    
    If strHinban Like "??C*-####MF*" Then
        IsG7_Flush = True
    End If

End Function

Public Function IsHikido(ByVal in_varHinban As Variant) As Boolean
'   *************************************************************
'   引戸確認

'   戻り値:Boolean
'       →True              引戸
'       →False             引戸以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.1.0 ADD

'2.3.0
'   →1801仕様追加
'2.7.0
'   →1808仕様追加
'   *************************************************************

    IsHikido = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*DC-####*" Or in_varHinban Like "*DT-####*" Or _
        in_varHinban Like "*KC-####*" Or in_varHinban Like "*KT-####*" Or _
        in_varHinban Like "*DM-####*" Or in_varHinban Like "*DL-####*" Or _
        in_varHinban Like "*DP-####*" Or in_varHinban Like "*DH-####*" Or _
        in_varHinban Like "*DE-####*" Or in_varHinban Like "*DJ-####*" Or _
        in_varHinban Like "*DF-####*" Or in_varHinban Like "*DQ-####*" Or _
        in_varHinban Like "*DU-####*" Or in_varHinban Like "*DN-####*" Or in_varHinban Like "*KU-####*" Or _
        in_varHinban Like "*DI-####*" Or in_varHinban Like "*DG-####*" Or _
        in_varHinban Like "*DD-####*" Or in_varHinban Like "*DV-####*" Or _
        in_varHinban Like "*VM-####*" Or in_varHinban Like "*VL-####*" Or in_varHinban Like "*VN-####*" Or _
        in_varHinban Like "*VF-####*" Or in_varHinban Like "*VQ-####*" Or _
        in_varHinban Like "*VI-####*" Or in_varHinban Like "*VG-####*" Or _
        in_varHinban Like "*JC-####*" Or in_varHinban Like "*GU-####*" Or _
        in_varHinban Like "*DY-####*" _
    Then

        IsHikido = True
        
    End If
    

End Function

Public Function IsKabetsukeGuide(in_varHinban As Variant) As Boolean
'   *************************************************************
'   壁付ガイド引戸確認

'   戻り値:Boolean
'       →True              壁付ガイド引戸
'       →False             壁付ガイド引戸以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.1.0 ADD
'   2.5.0
'       →バグ修正 KTとKUの頭の[*]が抜けていた
'   *************************************************************
    
    IsKabetsukeGuide = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*KC-####*-*" Or in_varHinban Like "*KT-####*-*" Or in_varHinban Like "*KU-####*-*" Then
        IsKabetsukeGuide = True
    End If
    
End Function

Public Function IsEndWakunashi(in_varHinban As Variant) As Boolean
'   *************************************************************
'   エンド枠無し確認

'   戻り値:Boolean
'       →True              エンド枠無し
'       →False             エンド枠無し以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.1.0 ADD
'   *************************************************************
    
    IsEndWakunashi = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*DU-####*" Or in_varHinban Like "*DN-####*" Or in_varHinban Like "*KU-####*" Or in_varHinban Like "*VN-####*" Then
        IsEndWakunashi = True
    End If
    
End Function

Public Function IsCaro_Panel(in_varHinban As Variant) As Boolean
'   *************************************************************
'   カロ（AF-1～3型　パネル戸）確認

'   戻り値:Boolean
'       →True              Caro（パネル）
'       →False             Caro（パネル）以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.1.0 ADD
'   *************************************************************
    
    Dim strHinban As String

    IsCaro_Panel = False
    
    If IsNull(in_varHinban) Then Exit Function

    strHinban = Replace(in_varHinban, "特 ", "")
    
    If strHinban Like "??B*-####A*-*" Or strHinban Like "??B*-####B*-*" Or strHinban Like "??B*-####O*-*" Then
        IsCaro_Panel = True
    End If
    
End Function

Public Function IsTerraceGlass(in_varHinban As Variant) As Boolean
'   *************************************************************
'   テラスガラスドア確認

'   戻り値:Boolean
'       →True              テラスガラスドア型
'       →False             テラスガラスドア型以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.1.0 ADD
'   *************************************************************

    Dim strHinban As String

    IsTerraceGlass = False
    
    If IsNull(in_varHinban) Then Exit Function

    strHinban = Replace(in_varHinban, "特 ", "")
    
    If strHinban Like "Y*-####?A*" Or strHinban Like "Y*-####?C*" Or strHinban Like "Y*-####?D*" Or strHinban Like "Y*-####?P*" Or strHinban Like "Y*-####?V*" Then
        IsTerraceGlass = True
    End If

End Function

Public Function IsHidden_Hinge(in_varHinban As Variant) As Boolean
'   *************************************************************
'   「隠し丁番」確認

'   戻り値:Boolean
'       →True              隠し丁番
'       →False             隠し丁番でない
'
'    Input項目
'       in_varHinban        建具品番

'   2.1.0 ADD
'   *************************************************************

    IsHidden_Hinge = False
      
     If IsNull(in_varHinban) Then Exit Function
     
    If in_varHinban Like "*KA-####*" Or in_varHinban Like "*KAS-####*" Or in_varHinban Like "*KO-####*" Or in_varHinban Like "*KOS-####*" Or in_varHinban Like "*KK-####*" Or in_varHinban Like "*KKS-####*" Then
        IsHidden_Hinge = True
    End If
    
End Function

Public Function IsYG6(in_varHinban As Variant) As Boolean
'   *************************************************************
'   テラス面縁ガラスドア確認

'   戻り値:Boolean
'       →True              テラス面縁ガラスドア
'       →False             テラス面縁ガラスドア以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.5.2 ADD
'   *************************************************************

    Dim strHinban As String

    IsYG6 = False
    
    If IsNull(in_varHinban) Then Exit Function

    strHinban = Replace(in_varHinban, "特 ", "")
    
    If strHinban Like "Y*-####T*" Then
        IsYG6 = True
    End If

End Function

Public Function IsPALIOBlack(in_varHinban As Variant) As Boolean
'   *************************************************************
'   パリオブラック（ビアンコ）確認
'   とりあえず中止になったのでFalseのみを返す

'   戻り値:Boolean
'       →True              パリオブラック
'       →False             パリオブラック以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.1.0 ADD
'   *************************************************************
'    If in_varHinban Like "*-*-*(NN)" Then
'        IsPALIOBlack = True
'    Else
        IsPALIOBlack = False
'    End If
    
End Function

Public Function IsTateguInset(in_varHinban As Variant) As Boolean
'   *************************************************************
'   建具品番の枠仕様がインセットか確認


'   戻り値:boolen
'       →True 枠仕様がインセット
'
'    Input項目
'       in_varHinban        品番

'   2.1.0 ADD
'   *************************************************************
    Dim strHinban As String
    
    IsTateguInset = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "特 ", "")
    
    If strHinban Like "?Z*-####*" Or strHinban Like "?Y*-####*" Or strHinban Like "?T*-####*" Then
        IsTateguInset = True
    End If
    
End Function

Public Function IsGuidePiece_ShitaanaKakou(in_varHinban As Variant, in_varTobiraIchi As Variant, in_varSpec As Variant, Optional in_SekkeiBikou As Variant) As Boolean
'   *************************************************************
'   ガイドピース下穴加工確認

'   戻り値:Boolean
'       →True              下穴加工あり
'       →False             下穴加工なし
'
'    Input項目
'       in_varHinban        建具品番
'       in_varTobiraIchi    扉位置（右、中、左）、それ以外は無条件にFalseを返す
'       in_varSpec          個別Spec 20160923時点では使用しない
'       in_SekkeiBikou      建具設計備考
       
'   2.1.0 ADD
'   *************************************************************
    
    Dim strTsurimoto As String
    
    On Error GoTo Err_IsGuidePiece_ShitaanaKakou
    
    If in_varTobiraIchi <> "右" And in_varTobiraIchi <> "左" And in_varTobiraIchi <> "中" Then
        Err.Raise 9999, , "ErrEnd"
    End If
    
    If IsNull(in_varHinban) Then
        Err.Raise 9999, , "ErrEnd"
    End If
    
    'エスパスライドウォールは除外
    If IsSlideWall_Espacio(in_varHinban) Then
        IsGuidePiece_ShitaanaKakou = False
        Exit Function
    End If
    
    '設計備考に以下コメントがある場合は除外
    If Not IsMissing(in_SekkeiBikou) Then
        If in_SekkeiBikou Like "*図面あり*戸首･戸車*" Then
            IsGuidePiece_ShitaanaKakou = False
            Exit Function
        End If
    End If
    
    If in_varHinban Like "*DF-####*-*" Or in_varHinban Like "*VF-####*-*" Then

    
        If in_varTobiraIchi = "中" Then
            IsGuidePiece_ShitaanaKakou = True
        End If
    ElseIf in_varHinban Like "*DH-####*-*" Then
        strTsurimoto = Mid(in_varHinban, InStr(1, in_varHinban, "(") - 1, 1)
        If strTsurimoto = "L" And in_varTobiraIchi = "右" Then
            IsGuidePiece_ShitaanaKakou = True
        ElseIf strTsurimoto = "R" And in_varTobiraIchi = "左" Then
            IsGuidePiece_ShitaanaKakou = True
        End If
    ElseIf in_varHinban Like "*DJ-####*-*" Then
        If in_varTobiraIchi = "中" Then
            IsGuidePiece_ShitaanaKakou = True
        Else
            strTsurimoto = Mid(in_varHinban, InStr(1, in_varHinban, "(") - 1, 1)
            If strTsurimoto = "L" And in_varTobiraIchi = "右" Then
                IsGuidePiece_ShitaanaKakou = True
            ElseIf strTsurimoto = "R" And in_varTobiraIchi = "左" Then
                IsGuidePiece_ShitaanaKakou = True
            End If
        End If
    End If
    
    Exit Function
    
Err_IsGuidePiece_ShitaanaKakou:
    IsGuidePiece_ShitaanaKakou = False
    
End Function

Public Function IsSlideWall_Espacio(in_varHinban As Variant) As Boolean
'   *************************************************************
'   スライドウォール（エスパ）確認

'   戻り値:Boolean
'       →True              スライドウォール
'       →False             スライドウォール以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.1.0 ADD
'   *************************************************************

    On Error GoTo Err_IsSlideWall_Espacio
    
    Dim strHinban As String
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "特 ", "")
    
    If strHinban Like "PSW*-####FV-*" Or strHinban Like "ESW*-####FV-*" Then
        IsSlideWall_Espacio = True
    Else
        IsSlideWall_Espacio = False
    End If
    
    Exit Function

Err_IsSlideWall_Espacio:
    IsSlideWall_Espacio = False
    
End Function

Public Function dblfncGuidePiece_ShitaanaSunpo(in_varHinban As Variant, in_varSpec As Variant) As Double
'   *************************************************************
'   ガイドピース下穴加工寸法

'   戻り値:Double
'       →下穴加工寸法（該当しない場合は0を返す）
'
'    Input項目
'       in_varHinban        建具品番
'       in_varSpec          個別Spec 20160923時点では使用しない

'   2.1.0 ADD
'   *************************************************************

    On Error GoTo Err_dblfncGuidePiece_ShitaanaSunpo
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*DF-####*-*" Or in_varHinban Like "*VF-####*-*" Then
    
        dblfncGuidePiece_ShitaanaSunpo = 60
    ElseIf in_varHinban Like "*DH-####*-*" Then
        dblfncGuidePiece_ShitaanaSunpo = 52.5
    ElseIf in_varHinban Like "*DJ-####*-*" Then
        dblfncGuidePiece_ShitaanaSunpo = 52.5
    Else
        dblfncGuidePiece_ShitaanaSunpo = 0
    End If
    
    Exit Function
    
Err_dblfncGuidePiece_ShitaanaSunpo:
    dblfncGuidePiece_ShitaanaSunpo = 0
End Function

Public Function intfncGuidePiece_ShitaanaSu(in_varHinban As Variant, in_varSpec As Variant) As Integer
'   *************************************************************
'   ガイドピース下穴加工数

'   戻り値:Integer
'       →下穴加工数（該当しない場合は0を返す）
'
'    Input項目
'       in_varHinban        建具品番
'       in_varSpec          個別Spec 20160923時点では使用しない
        
'   2.1.0 ADD
'   *************************************************************
    
    On Error GoTo Err_intfncGuidePiece_ShitaanaSu
    
    If IsNull(in_varHinban) Then Exit Function
    
    If in_varHinban Like "*DF-####*-*" Or in_varHinban Like "*VF-####*-*" Then
    
        intfncGuidePiece_ShitaanaSu = 2
    ElseIf in_varHinban Like "*DH-####*-*" Then
        intfncGuidePiece_ShitaanaSu = 1
    ElseIf in_varHinban Like "*DJ-####*-*" Then
        intfncGuidePiece_ShitaanaSu = 1
    Else
        intfncGuidePiece_ShitaanaSu = 0
    End If
    
    Exit Function
    
Err_intfncGuidePiece_ShitaanaSu:
    intfncGuidePiece_ShitaanaSu = 0

End Function

Public Function IsMirrorUsed(in_varHinban As Variant, Optional in_Tobiraichi As Variant = Null) As Boolean
'   *************************************************************
'   鏡使用扉確認

'   ※鏡を使用する品番を確認

'   戻り値:Boolean
'       →True              鏡使用
'       →False             鏡未使用
'
'    Input項目
'       in_varHinban        建具品番
'       in_Tobiraichi       扉位置(L or R or C or LC or RC) --オプション（引数がない場合は品番のみで判断）--201708時点では使用しない

'   *************************************************************
    Dim strHinban As String
    
    On Error GoTo Err_IsMirrorUsed
    
    IsMirrorUsed = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "特 ", "")
            
    '扉位置に関係なくTrueの品番
    If strHinban Like "*-####MF*" Or strHinban Like "*-####MM*" Then
        
        IsMirrorUsed = True
        
    '扉位置によって扉がある場合
    ElseIf strHinban Like "*-####ML*" Or strHinban Like "*-####MR*" Then
        
        '扉位置の指示がない場合はTrue
        If Nz(in_Tobiraichi, "") = "" Then
            IsMirrorUsed = True
        Else
            Select Case in_Tobiraichi
                Case "L"
                    If strHinban Like "*-####ML*" Then IsMirrorUsed = True
                        
                Case "R"
                    If strHinban Like "*-####MR*" Then IsMirrorUsed = True
            End Select
        End If
    End If
    
    'SxL建具
    '1701仕様時点ではなし

    Exit Function

Err_IsMirrorUsed:
    IsMirrorUsed = False
End Function

Public Function valfncHinmei_Local(in_strHinban As Variant, in_intSeihinkubun As Integer, in_strSpec As Variant) As Variant
'   *************************************************************
'   品名抽出（品名を返すのみ版）

'   戻り値:Variant → 品名（見つからない場合はNULL）
'
'    Input項目
'       in_strHinban        建具品番
'       in_intSeihinkubun   品番区分
'       in_strSpec          個別Spec
'   *************************************************************
    Dim objREMOTEdb As cls_BRAND_MASTER
    
    Dim strSQL As String
    Dim strHinban As String
    
    strSQL = ""
    valfncHinmei_Local = Null
    
    On Error GoTo Err_valfncHinmei_Local
    
    If IsNull(in_strHinban) Then GoTo Exit_valfncHinmei_Local
    
    strHinban = Replace(in_strHinban, "特 ", "")
    
    Select Case in_intSeihinkubun
        Case 1, 5 '建具,ｸﾛｾﾞｯﾄ
            strSQL = "select top 1 品名 from T_建具品番ﾏｽﾀ where "
                If IsKotobira(strHinban) Then
                    strSQL = strSQL & " 子扉品番 = '" & strHinban & "'"
                Else
                    strSQL = strSQL & " 建具品番 = '" & strHinban & "'"
                End If
        Case 2, 4 '枠,三方枠
            strSQL = "select top 1 品名 from T_枠品番ﾏｽﾀ where 枠品番 = '" & strHinban & "'"
            
        Case 3 '下地枠
            strSQL = "select top 1 品名 from T_下地材品番ﾏｽﾀ where 下地材品番 = '" & strHinban & "'"
          
        Case 6 '造作材
            strSQL = "select top 1 品名 from T_造作材品番ﾏｽﾀ where 造作材品番 = '" & strHinban & "'"
            
        Case 7 '玄関収納
            strSQL = "select top 1 品名 from T_玄関収納ﾏｽﾀ where 品番 = '" & strHinban & "'"
            
        Case 8 '金物
            strSQL = "select top 1 品名 from T_金物品番ﾏｽﾀ where 金物品番 = '" & strHinban & "'"
        
    End Select
    
    If strSQL = "" Then
        GoTo Exit_valfncHinmei_Local
    Else

        If Not IsNull(in_strSpec) And in_strSpec <> "" Then
            strSQL = strSQL & " and 仕様 = '" & left(in_strSpec, 3) & "' and '" & right(in_strSpec, 4) & "' between 開始 and 終了 "
        End If

    End If
    
    With objREMOTEdb
        If .ExecSelect(strSQL) Then
            If Not .GetRS.EOF Then
                valfncHinmei_Local = .GetRS![品名]
            End If
        End If
    End With
    
    GoTo Exit_valfncHinmei_Local
    
Err_valfncHinmei_Local:
    'MsgBox Err.Description
Exit_valfncHinmei_Local:
    Set objREMOTEdb = Nothing
End Function

Public Function IsTerraceKamachi(in_varHinban As Variant) As Boolean
'   *************************************************************
'   テラス框(YF5/YG5)ドア確認

'   戻り値:Boolean
'       →True              テラスドア型
'       →False             テラスドア型以外
'
'    Input項目
'       in_strHinban        建具品番

'   2.1.0 ADD
'   *************************************************************

    On Error GoTo Err_IsTerraceKamachi
    
    Dim strHinban As String
    
    IsTerraceKamachi = False
    
    If IsNull(in_varHinban) Then Exit Function
    
    strHinban = Replace(in_varHinban, "特 ", "")
    
    If strHinban Like "Y?B*-####*" Then
        IsTerraceKamachi = True
    Else
        IsTerraceKamachi = False
    End If
    
    Exit Function

Err_IsTerraceKamachi:
    IsTerraceKamachi = False
End Function

Public Function IsG9(in_varHinban As Variant) As Boolean
'   *************************************************************
'   Ｇ９型（細框戸）確認
'
'   戻り値:Boolean
'       →True              G9型
'       →False             G9型以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.3.0 ADD
'   *************************************************************

    Dim strHinban As String
    
    On Error GoTo Err_IsG9
    
    IsG9 = False
    
    If IsNull(in_varHinban) Then Exit Function

    strHinban = Replace(in_varHinban, "特 ", "")
    
    If strHinban Like "??B*-####E*-*" Then
        IsG9 = True
    End If
    
    Exit Function

Err_IsG9:
    IsG9 = False
    
End Function

Public Function IsTamo(in_varHinban As Variant) As Boolean
'   *************************************************************
'   タモシリーズ（JF1,JG1,JG2)確認
'
'   戻り値:Boolean
'       →True              タモシリーズ
'       →False             タモシリーズ以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.3.0 ADD
'   *************************************************************
    Dim strHinban As String
    
    On Error GoTo Err_IsTamo
    
    IsTamo = False
    
    If IsNull(in_varHinban) Then Exit Function

    strHinban = Replace(in_varHinban, "特 ", "")
    
    If strHinban Like "R*-####*-*(NT)*" Or strHinban Like "R*-####*-*(ZT)*" Then
        IsTamo = True
    End If
    
    Exit Function

Err_IsTamo:
    IsTamo = False
    
End Function

Public Function IsRendouTategu(in_varHinban As Variant) As Boolean
'   *************************************************************
'   連動建具確認
'
'   戻り値:Boolean
'       →True              連動建具（ガイドピース用下穴がある）
'       →False             連動建具以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.5.0 ADD
'   *************************************************************
    Dim strHinban As String
    
    On Error GoTo Err_IsRendouTategu
    
    IsRendouTategu = False
    
    If IsNull(in_varHinban) Then Exit Function

    strHinban = Replace(in_varHinban, "特 ", "")
    
    If strHinban Like "*VF-####*-*" Or strHinban Like "*DF-####*-*" Or strHinban Like "*DH-####*-*" Or strHinban Like "*DJ-####*-*" Then
        IsRendouTategu = True
    End If
    
    Exit Function

Err_IsRendouTategu:
    IsRendouTategu = False

End Function

Public Function IsHiRendouTategu(in_varHinban As Variant) As Boolean
'   *************************************************************
'   非連動建具確認
'
'   戻り値:Boolean
'       →True              非連動建具
'       →False             非連動建具以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.5.0 ADD
'   *************************************************************
    Dim strHinban As String
    
    On Error GoTo Err_IsHiRendouTategu
    
    IsHiRendouTategu = False
    
    If IsNull(in_varHinban) Then Exit Function

    strHinban = Replace(in_varHinban, "特 ", "")
    
    If strHinban Like "*DQ-####*-*" Or strHinban Like "*VQ-####*-*" Then
        IsHiRendouTategu = True
    End If
    
    Exit Function

Err_IsHiRendouTategu:
    IsHiRendouTategu = False

End Function

Public Function IsKousi(in_varHinban As Variant) As Boolean
'   *************************************************************
'   格子扉確認
'
'   戻り値:Boolean
'       →True              格子扉
'       →False             格子扉以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.5.2 ADD
'   *************************************************************
    Dim strHinban As String
    
    On Error GoTo Err_IsKousi
    
    IsKousi = False
    
    If IsNull(in_varHinban) Then Exit Function

    strHinban = Replace(in_varHinban, "特 ", "")
    
    If strHinban Like "Z?B*-####*-*" Then
        IsKousi = True
    End If
    
    Exit Function

Err_IsKousi:
    IsKousi = False
    
End Function

Public Function IsReversible(in_varHinban As Variant, varTateguSekkeiBikou As Variant, varSpec As Variant) As Boolean
'   *************************************************************
'   リバーシブル確認
'
'   戻り値:Boolean
'       →True                  リバーシブル
'       →False                 リバーシブル以外
'
'    Input項目
'       in_varHinban            建具品番
'       varTateguSekkeiBikou    建具設計備考
'       varSpec                 個別Spec

'   2.5.3 ADD
'   2.7.0
'       →ME扉1808にてリバーシブル廃止
'   *************************************************************

    Dim strTateguSekkeiBikou As String
    
    On Error GoTo Err_IsReversible
    
    IsReversible = False

    If IsNull(in_varHinban) Or IsNull(varSpec) Then
        Exit Function
    End If
    
    strTateguSekkeiBikou = Nz(varTateguSekkeiBikou, "")


    '   *************************************************************
    '   建具設計備考に「リバーシブル」が含まれている場合
    '   リバーシブル扱い
    '   *************************************************************
    
    If strTateguSekkeiBikou Like "*リバーシブル*" Or strTateguSekkeiBikou Like "*ﾘﾊﾞｰｼﾌﾞﾙ*" Then
        IsReversible = True
        
    '   *************************************************************
    '   F/Sシリーズ
    '   ZZ色(KF1)はリバーシブルでない
    '   *************************************************************
    
    ElseIf IsFs(CStr(in_varHinban)) Then
        If in_varHinban Like "*(ZZ)" Then 'KF1型
            Exit Function
        Else
            IsReversible = True 'KF7型はリバーシブル
        End If
        
    '   *************************************************************
    '   物入れ引き違い
    '   PH色はリバーシブルでない
    '   ミラーオプションはPHでもリバーシブル扱い
    '   *************************************************************
    
    ElseIf IsCloset_Hikichigai(CStr(in_varHinban)) Then '物入引き違い

        
        'ミラーはリバーシブル（中板があるため）
        If in_varHinban Like "*-####M*" Then
            
            IsReversible = True
                
        '1808仕様以降はリバーシブルでない
        ElseIf right(varSpec, 4) >= "1808" Then
            
            IsReversible = False
            
        '1801以前でも白はリバーシブルでない
        ElseIf in_varHinban Like "*-####*-*(PH)" Or (in_varHinban Like "*-####*-*(SH)" And right(varSpec, 4) >= "1701") Then

            IsReversible = False
        
        '1801以前の残り全てはリバーシブル
        Else
            IsReversible = True

        End If
        
    End If
    
    Exit Function
    
Err_IsReversible:
    IsReversible = False
    
End Function

Public Function IsFullGlass(in_varHinban As Variant) As Boolean
'   *************************************************************
'   ＶＧ１型（フルガラス）確認
'
'   戻り値:Boolean
'       →True              VG1型
'       →False             VG1型以外
'
'    Input項目
'       in_varHinban        建具品番

'   2.7.0 ADD
'   *************************************************************

    Dim strHinban As String
    
    On Error GoTo Err_IsFullglass
    
    IsFullGlass = False
    
    If IsNull(in_varHinban) Then Exit Function

    strHinban = Replace(in_varHinban, "特 ", "")
    
    If strHinban Like "X*-####X*-*" Then
        IsFullGlass = True
    End If
    
    Exit Function

Err_IsFullglass:
    IsFullGlass = False
    
End Function