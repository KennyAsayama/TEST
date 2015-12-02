Attribute VB_Name = "md_製品関数"
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
'       in_strHinban        建具品番

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
    
'    If (in_strHinban Like "*SG*-####*" Or in_strHinban Like "*NG*-####*" Or in_strHinban Like "*AG*-####*" Or in_strHinban Like "*BG*-####*") _
'        And Not in_strHinban Like "*ML-####*" And Not in_strHinban Like "*MK-####*"  And Not in_strHinban Like "*MT-####*" And Not in_strHinban Like "*DU-####*" And Not in_strHinban Like "*DN-####*" And Not in_strHinban Like "*CTSG*MK-####*" And Not in_strHinban Like "*CTSG*ML-####*"  And Not in_strHinban Like "*CTSG*MT-####*"  And Not in_strHinban Like "*KU-####*"  And Not in_strHinban Like "*KN-####*" Then
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

'   *************************************************************
    
    intFncSeizokubun = 0
    
    Select Case in_strShurui
    
        Case "建具", "子扉"
            
            If IsKamachi(in_varHinban) Then
            
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

'   *************************************************************
    
    IsFkamachi = False
    
    If IsNull(in_strHinban) Then Exit Function
       
    If in_strHinban Like "*-####G*-*" Or in_strHinban Like "*-####MF*-*" Or in_strHinban Like "*O*-####P*-*" Then
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

'   *************************************************************

    IsKamachi = False
    
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

'   *************************************************************
    On Error GoTo Err_IsThruGlass
    
    IsThruGlass = False
    
    If IsNull(in_strHinban) Then Exit Function
     
    If in_strHinban Like "*-####S*-*" Or in_strHinban Like "*-####C*-*" Or in_strHinban Like "*-####D*-*" _
        Or in_strHinban Like "*-####A*-*" Or in_strHinban Like "*-####B*-*" Or in_strHinban Like "*-####O*-*" _
        Or in_strHinban Like "*ME-####M*-*" Or in_strHinban Like "*SA-####M*-*" Or IsVertica(in_strHinban) Then
        
        IsThruGlass = True
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
'   *************************************************************

    IsOyatobira = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "*DO-####*" Or in_strHinban Like "*DOS-####*" _
       Or in_strHinban Like "*CO-####*" Or in_strHinban Like "*COS-####*" _
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
'   *************************************************************

    IsKotobira = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "*DK-####*" Or in_strHinban Like "*DKS-####*" Then
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
'       out_strSxLhinban    神谷品番(Falseの場合はNull)
'   *************************************************************
    
    Dim objLOCALDB As New cls_LOCALDB
    Dim strHinban As String
    Dim bolMentori As Boolean
    
    IsSxL = False
    
    On Error GoTo Err_IsSxL
    
    If IsNull(in_strHinban) Then GoTo Exit_IsSxL
    
    '下地で面取り記号がある場合は外す
    If in_strHinban Like "*①?②?③?④*" Then
        strHinban = left(in_strHinban, Len(in_strHinban) - 10)
        bolMentori = True
    Else
        strHinban = in_strHinban
        bolMentori = False
    End If
    '1.10.3 K.Asayama 20151119 SxL品番読替表ローカルテーブル名変更
    If objLOCALDB.ExecSelect("select ブランド品番 from WK_SxL品番読替表 where S×L品番 = '" & Trim(strHinban) & "'") Then
        If Not objLOCALDB.GetRS.EOF Then
            out_strKamiyahinban = objLOCALDB.GetRS![ブランド品番]
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
    Set objLOCALDB = Nothing
End Function

Public Function valfncHinmei(in_objRemoteDB As cls_BRAND_MASTER, in_Rs As ADODB.Recordset, in_strHinban As Variant, in_intSeihinkubun As Integer, in_strSpec As Variant) As Variant
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
    
    
    If in_objRemoteDB.ExecSelect_ExternalRS(in_Rs, strSQL) Then
        If Not in_Rs.EOF Then
            valfncHinmei = in_Rs![品名]
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

'   *************************************************************

    isCaro = False
    
    If in_varHinban Like "F_C*-####A*-*" Or in_varHinban Like "F_C*-####B*-*" Or in_varHinban Like "F_C*-####O*-*" _
        Or in_varHinban Like "特 F_C*-####A*-*" Or in_varHinban Like "特 F_C*-####B*-*" Or in_varHinban Like "特 F_C*-####O*-*" Then
    
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
            
            If in_varHinban Like "F_V*-####*" Then 'Vertica
                strfncSyobunrui_Kamui = "V"
                
            ElseIf in_varHinban Like "F_C*-####*" Then 'Caro
                strfncSyobunrui_Kamui = "A"
            
            Else
                Select Case strHinbanKigou
                    Case "F" '標準品はCUBEのコードを送る
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
'   戻り値:Boolean
'       →True              塗装扉
'       →False             塗装扉以外
'
'    Input項目
'       in_strHinban        建具品番

'   *************************************************************
    On Error GoTo Err_IsPainted
    
    IsPainted = False
    
    If IsNull(in_strHinban) Then Exit Function
    
    If in_strHinban Like "R*-####*-*" Or in_strHinban Like "特 R*-####*-*" Or in_strHinban Like "B*-####*-*" Or in_strHinban Like "特 B*-####*-*" Then
        IsPainted = True
    Else
        IsPainted = False
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

'   *************************************************************
    '
    IsStealth_Seizo_TEMP = False
    
    If (in_strHinban Like "*SG*-####*" Or in_strHinban Like "*NG*-####*" Or in_strHinban Like "*AG*-####*" Or in_strHinban Like "*BG*-####*") _
        And Not in_strHinban Like "*ML-####*" And Not in_strHinban Like "*MK-####*" And Not in_strHinban Like "*MT-####*" And Not in_strHinban Like "*DU-####*" And Not in_strHinban Like "*DN-####*" And Not in_strHinban Like "*CTSG*MK-####*" And Not in_strHinban Like "*CTSG*ML-####*" And Not in_strHinban Like "*CTSG*MT-####*" And Not in_strHinban Like "*KU-####*" And Not in_strHinban Like "*KN-####*" Then
        IsStealth_Seizo_TEMP = True
    End If
    
End Function

Public Function fncbolSxL_Replace() As Boolean
'   *************************************************************
'   SxL品番読替表置換え処理
'   1.10.3 K.Asayama ADD 20151119 SxL品番表リモートからコピー
'
'   リモートデータベースからローカルにSxL品番読替表をコピーする
'
'   戻り値:Boolean
'       →True              置換成功
'       →False             置換失敗
'
'   *************************************************************

    fncbolSxL_Replace = False
    
    Dim objREMOTEDB As New cls_BRAND_MASTER
    Dim objLOCALDB As New cls_LOCALDB
    
    On Error GoTo Err_fncbolSxL_Replace
    
    Dim strSQL_Insert As String
    Dim strSQL As String
    strSQL_Insert = "Insert into WK_SxL品番読替表(S×L品番,ブランド品番,DH,DW,CH) values ("
    
    '工場用コピー（T_Calendar_工場)
    If objLOCALDB.ExecSQL("delete from WK_SxL品番読替表") Then
        strSQL = "select distinct [S×L品番],ブランド品番,DW,DH,CH from SxL品番読替表 "
        If objREMOTEDB.ExecSelect(strSQL) Then
            Do While Not objREMOTEDB.GetRS.EOF
                If Not objLOCALDB.ExecSQL(strSQL_Insert & "'" & objREMOTEDB.GetRS![S×L品番] & "','" & objREMOTEDB.GetRS![ブランド品番] & "'," & objREMOTEDB.GetRS![DW] & "," & objREMOTEDB.GetRS![DH] & "," & objREMOTEDB.GetRS![CH] & ")") Then
                    Err.Raise 9999, , "SxL品番読替表 ローカルコピーエラー"
                End If
                objREMOTEDB.GetRS.MoveNext
            Loop
        End If
    End If
    
    fncbolSxL_Replace = True
    
    GoTo Exit_fncbolSxL_Replace
    
Err_fncbolSxL_Replace:
    MsgBox Err.Description
    
Exit_fncbolSxL_Replace:

    Set objREMOTEDB = Nothing
    Set objLOCALDB = Nothing
    
End Function
