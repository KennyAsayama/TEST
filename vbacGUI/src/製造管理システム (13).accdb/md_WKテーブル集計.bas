Option Compare Database
Option Explicit

Public Function SetOrderData(ByVal inDate As Date, ByVal inDateKbn As Byte, inSeizoKbn As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'製造データをWKファイルに転送する
'
'   :引数
'       inDate          :納品日
'       inDateKbn       :1:納品日ベース集計、2:製造日ベース集計
'       inSeizoKbn      :建具、枠、下地
'
'   :戻り値
'       True            :成功
'       False           :失敗
'1.10.7 K.Asayama ADD 20160108
'       →「F_邸別_数量」工程表ボタンを使用可能にする引数を追加
'       →「WK_札データ」に出荷方法、色（塗装のみ）を追加
'       → 製造日ベースの時は未確定も集計
'       → inDateが[9999/12/31]の時は日付Nullのデータを出力（製造日ベース）
'       → inDateが[9999/12/30]の時は日付は関係なく未確定のデータを出力（製造日ベース）
'1.10.8 K.Asayama ADD 20160114
'       →ヴェルチカ分割
'1.10.10 K.Asayama Change 20160212
'       →物入れ引き違い片側ミラーオプション追加
'1.10.14 K.Asayama Change 20160418
'       →バグ修正計算出荷日がNullで戻った場合の対応
'1.10.16 K.Asayama Change
'       →下地、ステルス分割
'2.5.0
'       →出荷日計算をリードタイムに変更
'2.5.2
'       →F框の塗装集計対応
'2.8.0
'       →アラジンオフィス納期情報データ取り込み
'2.9.0
'       →リモートとの接続時間短縮化のためワークテーブル作成
'       →リードタイムから出荷日計算をサーバサイド化（パフォーマンス改善）
'       →DoEvents追加
'2.13.0
'       →Verticaシンクロ枚数対応
'3.0.0
'       →Specを転送
'--------------------------------------------------------------------------------------------------------------------
    Dim objREMOTEdb As New cls_BRAND_MASTER
    Dim objLOCALdb As New cls_LOCALDB
    Dim objLOCALDB_2 As New cls_LOCALDB
    
    Dim strSQL As String
    Dim strSQLWK As String
    
    Dim bolTRAN As Boolean
    Dim strKeiyakuNo As String
    Dim varCalcShukkaBi As Variant
    Dim intMinusDays As Integer
    Dim dblWindowTop As Double, dblWindowLeft As Double, dblWindowHight As Double, dblWindowWidth As Double
    Dim bolFormOpen As Boolean
    Dim strKubun As String
    Dim strLT As String
    Dim i As Integer
    
    bolFormOpen = False
    
    On Error GoTo Err_SetOrderData
    
'    Me.Painting = False
    Application.Echo False
    
    SetOrderData = False
    bolTRAN = False
    strKeiyakuNo = ""
    
    Select Case inSeizoKbn
        Case "建具"
            strKubun = "1,2,3"
            strLT = "建具LT"
        Case "枠"
            strKubun = "4,5"
            strLT = "枠LT"
        '1.10.16 Change
'        Case "下地"
'            strKubun = "6,7"
        Case "下地"
            strKubun = "6"
            strLT = "下地材LT"
        Case "ステルス"
            strKubun = "7"
            strLT = "下地材LT"
        Case Else
            Err.Raise 9999, , "製造区分転送エラー"
    End Select
    
    strSQL = ""
    strSQL = strSQL & "select * "
    strSQL = strSQL & ",case when 出荷日 is null then dbo.fnc出荷日取得_LTのみ(確定日," & strLT & ") "
    strSQL = strSQL & " else null end 計算出荷日 "
    strSQL = strSQL & "from ( "
    strSQL = strSQL & "select s.契約番号,s.棟番号,s.部屋番号 "
    strSQL = strSQL & ",s.契約番号 + '-' + s.棟番号 + '-' + s.部屋番号 AS 契約No "
    strSQL = strSQL & ",s.確定日 "
    strSQL = strSQL & ",dbo.fncSeizosyukkaDate(s.契約番号,s.棟番号,s.部屋番号,s.項,"
    strSQL = strSQL & "case s.製造区分 "
    strSQL = strSQL & "when 1 then 1 "
    strSQL = strSQL & "when 2 then 1 "
    strSQL = strSQL & "when 3 then 1 "
    strSQL = strSQL & "when 4 then 2 "
    strSQL = strSQL & "when 5 then 5 "
    strSQL = strSQL & "when 6 then 3 "
    strSQL = strSQL & "when 7 then 3 "
    strSQL = strSQL & "else 999 "
    strSQL = strSQL & "end) as 出荷日 "
    strSQL = strSQL & ",dbo.fncNohinAddress(s.契約番号,s.棟番号,s.部屋番号,s.項,"
    strSQL = strSQL & "case s.製造区分 "
    strSQL = strSQL & "when 1 then 1 "
    strSQL = strSQL & "when 2 then 1 "
    strSQL = strSQL & "when 3 then 1 "
    strSQL = strSQL & "when 4 then 2 "
    strSQL = strSQL & "when 5 then 5 "
    strSQL = strSQL & "when 6 then 3 "
    strSQL = strSQL & "when 7 then 3 "
    strSQL = strSQL & "else 999 "
    strSQL = strSQL & "end) as 納品住所 "
    strSQL = strSQL & ",s.製造日 "
    strSQL = strSQL & ",s.項 "
    strSQL = strSQL & ",s.製造区分 "
    strSQL = strSQL & ",s.特注 "
    strSQL = strSQL & ",m.物件名 "
    strSQL = strSQL & ",m.施工店 "
    strSQL = strSQL & ",case s.製造区分 when 1  then s.数量 else 0 end AS [Flush数] "
    strSQL = strSQL & ",case s.製造区分 when 2  then s.数量 else 0 end AS [F框数] "
    strSQL = strSQL & ",case s.製造区分 when 3  then s.数量 else 0 end AS [框数] "
    strSQL = strSQL & ",case s.製造区分 when 4  then s.数量 else 0 end AS [枠数] "
    strSQL = strSQL & ",case s.製造区分 when 5  then s.数量 else 0 end AS [三方枠数] "
    strSQL = strSQL & ",case s.製造区分 when 6  then s.数量 else 0 end AS [下地枠数] "
    strSQL = strSQL & ",case s.製造区分 when 7  then s.数量 else 0 end AS [ステルス枠数] "
    strSQL = strSQL & ",s.登録時品番 "
    strSQL = strSQL & ",s.数量 "
    strSQL = strSQL & ",s.確定 "
    strSQL = strSQL & ",y.コメント as 備考 "
    '1.10.7 ADD
    strSQL = strSQL & ",dbo.fncNohinHaiso(s.契約番号,s.棟番号,s.部屋番号,s.項,"
    strSQL = strSQL & "case s.製造区分 "
    strSQL = strSQL & "when 1 then 1 "
    strSQL = strSQL & "when 2 then 1 "
    strSQL = strSQL & "when 3 then 1 "
    strSQL = strSQL & "when 4 then 2 "
    strSQL = strSQL & "when 5 then 5 "
    strSQL = strSQL & "when 6 then 3 "
    strSQL = strSQL & "when 7 then 3 "
    strSQL = strSQL & "else 999 "
    strSQL = strSQL & "end) as 出荷方法 "
    '1.10.7 ADD End
    strSQL = strSQL & ",建具LT,枠LT,下地材LT,WTLT,金物LT,玄関収納LT,造作材LT "
    
    strSQL = strSQL & ",m.Spec,s.登録時個別Spec 個別Spec "
    
    If inSeizoKbn = "建具" Then
        strSQL = strSQL & ",ガラス入荷日,ルーバー入荷日,その他入荷日,出荷金物入荷日 "
    Else
        strSQL = strSQL & ",Null as ガラス入荷日,Null as ルーバー入荷日,Null as その他入荷日,Null as 出荷金物入荷日 "
    End If
    
    strSQL = strSQL & "from T_製造指示 s "
    strSQL = strSQL & "inner join T_受注ﾏｽﾀ m "
    strSQL = strSQL & "on s.契約番号 = m.契約番号 and s.棟番号 = m.棟番号 and s.部屋番号 = m.部屋番号 "
    strSQL = strSQL & "inner join T_受注ﾏｽﾀ_2 m2 "
    strSQL = strSQL & "on s.契約番号 = m2.契約番号 and s.棟番号 = m2.棟番号 and s.部屋番号 = m2.部屋番号 "
    strSQL = strSQL & "left join T_製造予備 y "
    strSQL = strSQL & "on s.契約番号 = y.契約番号 and s.棟番号 = y.棟番号 and s.部屋番号 = y.部屋番号 and s.製造区分 = y.製造区分 "
    
    If inSeizoKbn = "建具" Then
        strSQL = strSQL & "left join (select 契約番号,棟番号,部屋番号,項 "
        strSQL = strSQL & ",max(case when 部材種別CD like '%ｶﾞﾗｽ%' or 部材種別CD like '%ﾐﾗｰ%'  then 入荷日 end) ガラス入荷日 "
        strSQL = strSQL & ",max(case when 部材種別CD like '%ﾙｰﾊﾞｰﾕﾆｯﾄ%' then 入荷日 end) ルーバー入荷日 "
        strSQL = strSQL & ",max(case when 部材種別CD not like '%ｶﾞﾗｽ%' and 部材種別CD not like '%ﾐﾗｰ%' and 部材種別CD not like '%ﾙｰﾊﾞｰﾕﾆｯﾄ%' and (同梱品 is null or 同梱品 <> '○') then 入荷日 end) その他入荷日 "
        strSQL = strSQL & ",max(case when 同梱品 = '○' then 入荷日 end) 出荷金物入荷日 "
        strSQL = strSQL & "from T_AO資材納期情報 AO "
        strSQL = strSQL & "where 品番区分 = 1 and 製造区分 = 1 "
        strSQL = strSQL & "group by 契約番号,棟番号,部屋番号,項 "
        strSQL = strSQL & ") AO "
        strSQL = strSQL & "on s.契約番号 = AO.契約番号 and s.棟番号 = AO.棟番号 and s.部屋番号 = AO.部屋番号 and s.項 = AO.項 "
    End If
    
    If inDateKbn = 1 Then
        strSQL = strSQL & "where s.確定日 = '" & Format(inDate, "yyyy/mm/dd") & "' "
    Else
        '1.10.7 ADD
        If inDate = #12/31/9999# Then
            strSQL = strSQL & "where s.製造日 is Null "
            
        ElseIf inDate = #12/30/9999# Then
            strSQL = strSQL & " where 確定 < 2 "
        Else
        '1.10.7 ADD End
            strSQL = strSQL & "where s.製造日 = '" & Format(inDate, "yyyy/mm/dd") & "' "
            '1.10.7 DEL
            'strSQL = strSQL & " and 確定 > 0 "
            '1.10.7 DEL END
        
        '1.10.7 ADD
        End If
        '1.10.7 ADD End
    End If
    strSQL = strSQL & " and s.製造区分 in ( " & strKubun & ") "
    
    'ウォールスルーは製造日を入れていないので対象外
    If inSeizoKbn = "下地" Then
        strSQL = strSQL & " and s.登録時品番 not like 'WS%' "
    End If
    
    strSQL = strSQL & " ) WKTABLE "
    
    If Not objLOCALdb.ExecSQL("delete from WK_札データ") Then
        Err.Raise 9999, , "製造指示データワーク（ローカル）初期化エラー"
    End If
    
    '最初はないのでエラーは無視
    objLOCALdb.ExecSQL ("drop table TMP_製造指示データ")
    
    strSQLWK = ""
    strSQLWK = strSQLWK & "CREATE TABLE TMP_製造指示データ( "
    strSQLWK = strSQLWK & " 契約番号            TEXT(10) "
    strSQLWK = strSQLWK & ",棟番号              TEXT(10) "
    strSQLWK = strSQLWK & ",部屋番号            TEXT(10) "
    strSQLWK = strSQLWK & ",契約No              TEXT(30) "
    strSQLWK = strSQLWK & ",確定日              DATE "
    strSQLWK = strSQLWK & ",出荷日              DATE "
    strSQLWK = strSQLWK & ",納品住所            TEXT(255) "
    strSQLWK = strSQLWK & ",製造日              DATE "
    strSQLWK = strSQLWK & ",項                  INT "
    strSQLWK = strSQLWK & ",製造区分            INT "
    strSQLWK = strSQLWK & ",特注                INT "
    strSQLWK = strSQLWK & ",物件名              TEXT(255) "
    strSQLWK = strSQLWK & ",施工店              TEXT(255) "
    strSQLWK = strSQLWK & ",Flush数             INT "
    strSQLWK = strSQLWK & ",F框数               INT "
    strSQLWK = strSQLWK & ",框数                INT "
    strSQLWK = strSQLWK & ",枠数                INT "
    strSQLWK = strSQLWK & ",三方枠数            INT "
    strSQLWK = strSQLWK & ",下地枠数            INT "
    strSQLWK = strSQLWK & ",ステルス枠数        INT "
    strSQLWK = strSQLWK & ",登録時品番          TEXT(50) "
    strSQLWK = strSQLWK & ",数量                INT "
    strSQLWK = strSQLWK & ",確定                INT "
    strSQLWK = strSQLWK & ",備考                TEXT(255) "
    strSQLWK = strSQLWK & ",出荷方法            TEXT(50) "
    strSQLWK = strSQLWK & ",建具LT              INT "
    strSQLWK = strSQLWK & ",枠LT                INT "
    strSQLWK = strSQLWK & ",下地材LT            INT "
    strSQLWK = strSQLWK & ",WTLT                INT "
    strSQLWK = strSQLWK & ",金物LT              INT "
    strSQLWK = strSQLWK & ",玄関収納LT          INT "
    strSQLWK = strSQLWK & ",造作材LT            INT "
    strSQLWK = strSQLWK & ",ガラス入荷日        DATE "
    strSQLWK = strSQLWK & ",ルーバー入荷日      DATE "
    strSQLWK = strSQLWK & ",その他入荷日        DATE "
    strSQLWK = strSQLWK & ",出荷金物入荷日      DATE "
    strSQLWK = strSQLWK & ",計算出荷日          DATE "
    strSQLWK = strSQLWK & ",Spec                TEXT(20) "
    strSQLWK = strSQLWK & ",個別Spec            TEXT(20) "
    strSQLWK = strSQLWK & ") "
        
    If Not objLOCALdb.ExecSQL(strSQLWK) Then
        Err.Raise 9999, , "製造指示データワーク（ローカル）作成エラー"
    End If
    
    
    With objREMOTEdb
        If .ExecSelect(strSQL) Then
            If Not bolfncTableCopyToLocal(.GetRS, "TMP_製造指示データ", False) Then
                Err.Raise 9999, , "TMP_製造指示データローカルコピーエラー。管理者に連絡してください"
            End If
        Else
            Err.Raise 9999, , ""
        End If
    End With
    
    strSQL = ""
    strSQL = strSQL & "select * from TMP_製造指示データ "
    
    i = 0
    
    With objLOCALDB_2
        If .ExecSelect(strSQL) Then
            If objLOCALdb.ExecSelect_Writable("select * from WK_札データ") Then
            
                objLOCALdb.BeginTrans
                bolTRAN = True
                
                Do While Not .GetRS.EOF
                        objLOCALdb.GetRS.AddNew

                        objLOCALdb.GetRS![契約番号] = .GetRS![契約番号]
                        objLOCALdb.GetRS![棟番号] = .GetRS![棟番号]
                        objLOCALdb.GetRS![部屋番号] = .GetRS![部屋番号]
                        objLOCALdb.GetRS![物件名] = .GetRS![物件名]
                        objLOCALdb.GetRS![施工店] = .GetRS![施工店]
                        objLOCALdb.GetRS![契約No] = .GetRS![契約No]
                        objLOCALdb.GetRS![項] = .GetRS![項]
                        objLOCALdb.GetRS![製造区分] = .GetRS![製造区分]
                        objLOCALdb.GetRS![確定日] = .GetRS![確定日]
'                        If IsNull(.GetRS![出荷日]) Then
'                            objLOCALDB.GetRS![出荷日登録] = False
'                            'If fncbolSyukkaBiFromAddress(.GetRS![納品住所], .GetRS![確定日], varCalcShukkaBi, intMinusDays) Then
'                            If fncbolSyukkaBiFromLeadTime(.GetRS(strLT), .GetRS![確定日], varCalcShukkaBi, intMinusDays) Then
'                                '1.10.14
'                                If Not IsNull(varCalcShukkaBi) Then
'                                    objLOCALDB.GetRS![出荷日] = CDate(varCalcShukkaBi)
'                                End If
'                            Else
'                                objLOCALDB.GetRS![出荷日] = .GetRS![出荷日]
'                            End If
'                        Else
'                            objLOCALDB.GetRS![出荷日登録] = True
'                            objLOCALDB.GetRS![出荷日] = .GetRS![出荷日]
'                        End If
                        
                        If IsNull(.GetRS![出荷日]) Then
                            objLOCALdb.GetRS![出荷日登録] = False
                            If Not IsNull(.GetRS![計算出荷日]) Then
                                objLOCALdb.GetRS![出荷日] = CDate(.GetRS![計算出荷日])
                            Else
                                objLOCALdb.GetRS![出荷日] = Null
                            End If
                        Else
                            objLOCALdb.GetRS![出荷日登録] = True
                           objLOCALdb.GetRS![出荷日] = .GetRS![出荷日]
                        End If

                        objLOCALdb.GetRS![製造日] = .GetRS![製造日]
                        objLOCALdb.GetRS![納品住所] = .GetRS![納品住所]
                        
                        'If IsNull(.GetRS![確定]) Or .GetRS![確定] = 0 Then
                        '    objLocalDB.GetRS![確定] = 0
                        'Else
                        '    objLocalDB.GetRS![確定] = -1
                        'End If
                        
                        objLOCALdb.GetRS![確定] = .GetRS![確定]
                        
                        '1.10.7 ADD
                        objLOCALdb.GetRS![出荷方法] = .GetRS![出荷方法]
                        '1.10.7 ADD End
                        
                        objLOCALdb.GetRS![Flush数] = .GetRS![Flush数] + .GetRS![F框数]
                        objLOCALdb.GetRS![F框数] = .GetRS![F框数]
                        objLOCALdb.GetRS![框数] = .GetRS![框数]
                        objLOCALdb.GetRS![枠数] = .GetRS![枠数]
                        objLOCALdb.GetRS![三方枠数] = .GetRS![三方枠数]
                                                
                        If IsStealth_Seizo_TEMP(Nz(.GetRS![登録時品番], "nz")) Then
                            objLOCALdb.GetRS![下地枠数] = 0
                            '1.10.16 change
                            'objLOCALDB.GetRS![ステルス枠数] = .GetRS![下地枠数]
                            If .GetRS![製造区分] = 7 Then
                                objLOCALdb.GetRS![ステルス枠数] = .GetRS![ステルス枠数]
                            Else
                                objLOCALdb.GetRS![ステルス枠数] = .GetRS![下地枠数]
                            End If
                        Else
                            objLOCALdb.GetRS![ステルス枠数] = 0
                            objLOCALdb.GetRS![下地枠数] = .GetRS![下地枠数]
                        End If
                        
                        If .GetRS![製造区分] >= 1 And .GetRS![製造区分] <= 3 Then
                            If IsThruGlass(.GetRS![登録時品番]) Then
                                If IsVertica(.GetRS![登録時品番]) Then
                                    objLOCALdb.GetRS![スルーガラス数] = IsVertica_Maisu(.GetRS![登録時品番], .GetRS![Flush数])
                                Else
                                    '1.10.10 K.Asayama Change
                                    'objLOCALDB.GetRS![スルーガラス数] = .GetRS![Flush数]
                                    objLOCALdb.GetRS![スルーガラス数] = fncIntHalfGlassMirror_Maisu(.GetRS![登録時品番], .GetRS![Flush数])
                                    '1.10.10 K.Asayama Change End
                                End If
                            Else
                                objLOCALdb.GetRS![スルーガラス数] = 0
                            End If
                            
                            If IsAir(.GetRS![登録時品番]) Then
                                objLOCALdb.GetRS![ルーバー扉数] = .GetRS![Flush数]
                            Else
                                objLOCALdb.GetRS![ルーバー扉数] = 0
                            End If
                            
                            If IsPainted(.GetRS![登録時品番]) Then
                                If .GetRS![F框数] > 0 Then
                                    objLOCALdb.GetRS![塗装扉数] = .GetRS![F框数]
                                Else
                                    objLOCALdb.GetRS![塗装扉数] = .GetRS![Flush数]
                                End If
                                '1.10.7 ADD
                                objLOCALdb.GetRS![色] = fncvalDoorColor(.GetRS![登録時品番])
                                '1.10.7 ADD End
                            Else
                                objLOCALdb.GetRS![塗装扉数] = 0
                            End If
                            
                            If IsMonster(.GetRS![登録時品番]) Then
                                objLOCALdb.GetRS![モンスター数] = .GetRS![F框数]
                            Else
                                objLOCALdb.GetRS![モンスター数] = 0
                            End If
                            '1.10.8 ADD
                            If IsVertica(.GetRS![登録時品番]) Then
                                'objLOCALdb.GetRS![ヴェルチカ数] = .GetRS![Flush数]
                                objLOCALdb.GetRS![ヴェルチカ数] = IsVertica_Maisu(.GetRS![登録時品番], .GetRS![Flush数])
                            Else
                                objLOCALdb.GetRS![ヴェルチカ数] = 0
                            End If
                            '1.10.8 ADD End
                        Else
                            objLOCALdb.GetRS![スルーガラス数] = 0
                            objLOCALdb.GetRS![ルーバー扉数] = 0
                            objLOCALdb.GetRS![塗装扉数] = 0
                            objLOCALdb.GetRS![モンスター数] = 0
                            '1.10.8 ADD
                            objLOCALdb.GetRS![ヴェルチカ数] = 0
                            '1.10.8 ADD End
                        End If
                        
                        objLOCALdb.GetRS![備考] = .GetRS![備考]
                    
                    objLOCALdb.GetRS![ガラス入荷日] = .GetRS![ガラス入荷日]
                    objLOCALdb.GetRS![ルーバー入荷日] = .GetRS![ルーバー入荷日]
                    objLOCALdb.GetRS![その他入荷日] = .GetRS![その他入荷日]
                    objLOCALdb.GetRS![出荷金物入荷日] = .GetRS![出荷金物入荷日]
                    
                    objLOCALdb.GetRS![Spec] = .GetRS![Spec]
                    objLOCALdb.GetRS![個別Spec] = .GetRS![個別Spec]
                    
                    objLOCALdb.GetRS.Update
                                        
                    i = i + 1
                    
                    If i Mod 100 = 0 Then
                        DoEvents
                    End If
                    
                    .GetRS.MoveNext
                Loop
                
                If bolTRAN Then objLOCALdb.Commit
                bolTRAN = False
            Else
                Err.Raise 9999, , "チェックリストワーク（ローカル）オープンエラー"
            
            End If
        Else
            Err.Raise 9999, , "チェックリスト抽出エラー"
        End If
    End With
    
    '1.10.7 ADD 備考データ呼び出し
    If Not SetBikouData() Then
        Err.Raise 9999, , "備考情報呼び出し異常"
    End If
    
    DoCmd.SetWarnings False
    
    
    If Form_IsLoaded("F_邸別_数量") Then
        bolFormOpen = True
    End If
    
    
    If Not bolFormOpen Then
        DoCmd.OpenForm "F_邸別_数量", acNormal, , , , , inDateKbn
    Else
        '1.10.7 Change
        'If Not Form_F_邸別_数量.bolfncData_Update(inSeizoKbn) Then
        If Not Form_F_邸別_数量.bolfncData_Update(inSeizoKbn, inDateKbn) Then
        '1.10.7 Change End
            DoCmd.Close acForm, "F_邸別_数量", acSaveNo
        End If
    End If
    
    If bolFormOpen = True Then
        DoCmd.SelectObject acForm, "F_邸別_数量"
    End If
    
    DoCmd.SetWarnings True
    
    SetOrderData = True
    GoTo Exit_SetOrderData
    
Err_SetOrderData:
    If bolTRAN Then objLOCALdb.Rollback
    bolTRAN = False
    MsgBox Err.Description

Exit_SetOrderData:
    Set objREMOTEdb = Nothing
    Set objLOCALdb = Nothing
    Set objLOCALDB_2 = Nothing
    
    Application.Echo True
    'Me.Painting = True
    
End Function

Public Function SetOrderCount(ByVal inDateKbn As Byte, ByRef Captionctl() As cls_Labelset, ByRef Graphctl() As cls_Labelset, ByRef Graphctl_Kakutei() As cls_Labelset, ByRef Graphctl_Temp() As cls_Labelset, ByRef Itemctl() As cls_Labelset, ByVal in_HinbanKubun As Integer, ByVal in_KojoCD As Integer)
'--------------------------------------------------------------------------------------------------------------------
'数量集計処理
'
'   :引数
'       inDateKbn       :1:納品日ベース集計、2:製造日ベース集計
'       Captionctl      :日付表示ラベル（コントロール配列）
'       Graphctl        :数量表示ラベル（コントロール配列）
'       Graphctl_Kakutei:確定数量表示ラベル（コントロール配列）
'       Graphctl_Temp   :仮確定数量表示ラベル（コントロール配列）
'       Itemctl         :製品表示ラベル（2次元コントロール配列[日付,製品]）
'       in_HinbanKubun  :1,建具、2,枠、3,下地
'       in_KojoCD       :工場CD

'
'   :戻り値
'       True            :成功
'       False           :失敗
'---------------------------
'   変更
'       1.10.1 K.Asayama 下地数、ステルス数をラベル表示（各々ガラス数、モンスター数を流用）
'       1.10.7 K.Asayama 引数に仮確定（Graphctl_Temp）追加、確定数集計追加 暫定で表示はしない。確定数のラベルにカッコとじで数量表示
'       1.10.8 K.Asayama
'                       →框用ラベルをヴェルチカ用ラベルに変更
'                       →グラフラベルの数量（Caption）のControlTipText対応
'       1.10.10 K.Asayama Change 20160212
'                       →物入れ引き違い片側ミラーオプション追加
'       2.0.0
'                       →工場CDを使用しない

'       2.13.0
'                       →Verticaシンクロ枚数対応
'--------------------------------------------------------------------------------------------------------------------
    Dim objREMOTEdb As New cls_BRAND_MASTER
    Dim strSQL_C As String
    Dim strSQL As String
    Dim strKubun As String
    Dim i As Integer
    Dim bolToku As Boolean
    
    Dim intFlushM As Integer
    Dim intFkamachiM As Integer
    Dim intKamachiM As Integer
    Dim intThruM As Integer
    Dim intPaintM As Integer
    Dim intAirM As Integer
    Dim intMonsterM As Integer
    Dim intKakuteiM As Integer
    Dim intShitajiM As Integer
    Dim intStealthM As Integer
    '1.10.7 ADD
    Dim intKakuteiTempM As Integer
    '1.10.7 ADD End
    '1.10.8 ADD
    Dim intVerticaM As Integer
    '1.10.8 ADD End
    
    On Error GoTo Err_SetOrderCount
    
    '下地と共用するラベルを一旦建具で初期化
    For i = 0 To UBound(Itemctl)
        Itemctl(i, 1).CaptionSet ("ガラス")
        Itemctl(i, 4).CaptionSet ("Monster")
        
        If inDateKbn = 1 Then
            Itemctl(i, 1).SetWidth (197)
            Itemctl(i, 4).SetWidth (197)
        Else
            Itemctl(i, 1).SetWidth (107)
            Itemctl(i, 4).SetWidth (107)
        End If
    Next
    
    Select Case in_HinbanKubun
        Case 1 'Flush
            strKubun = "1,2,3"
        Case 2 'Waku
            strKubun = "4,5"
        '1.10.16 Change
'        Case 3 'Shitaji
'            strKubun = "6,7"
        Case 3 'Shitaji
            strKubun = "6"
        Case 4 'Stealth
            strKubun = "7"
        Case Else
            strKubun = CStr(in_HinbanKubun)
    End Select
    
    strSQL_C = "select s.登録時品番, s.製造区分, s.確定, s.数量 as 枚数 from T_製造指示 s "
    strSQL_C = strSQL_C & "inner join T_受注明細 m "
    strSQL_C = strSQL_C & "on m.契約番号 = s.契約番号 and m.棟番号 = s.棟番号 and m.部屋番号 = s.部屋番号 and m.項 = s.項 "
    
    
    For i = 0 To UBound(Captionctl)
        strSQL = strSQL_C
        If inDateKbn = 1 Then
            strSQL = strSQL & " where s.確定日 = '" & Captionctl(i).GetTag & "'"
        Else
            strSQL = strSQL & " where s.製造日 = '" & Captionctl(i).GetTag & "'"
            '1.10.7 K.Asayama Change
            'strSQL = strSQL & " and 確定 > 0 "
            '1.10.7 K.Asayama Change End
        End If
        strSQL = strSQL & " and s.製造区分 in ( " & strKubun & ")"
'        Select Case in_HinbanKubun
'            Case 1
'                strSQL = strSQL & " and 製造区分 = 1"
'            Case 6
'                strSQL = strSQL & " and 製造区分 between 6 and 7"
'            Case Else
'                strSQL = strSQL & " and 製造区分= " & in_HinbanKubun
'        End Select
'
        'ウォールスルーは製造日を入れていないので対象外
        If in_HinbanKubun = 3 Then
            strSQL = strSQL & " and s.登録時品番 not like 'WS%' "
        End If
    
'        strSQL = strSQL & " and s.工場CD = " & in_KojoCD
        strSQL = strSQL & " "
        
        With objREMOTEdb
           If .ExecSelect(strSQL) Then
                intFlushM = 0
                intFkamachiM = 0
                intKamachiM = 0
                intThruM = 0
                intPaintM = 0
                intAirM = 0
                intMonsterM = 0
                intKakuteiM = 0
                intShitajiM = 0
                intStealthM = 0
                '1.10.7 ADD
                intKakuteiTempM = 0
                '1.10.7 ADD End
                '1.10.8 ADD
                intVerticaM = 0
                '1.10.8 ADD End
    
                Do Until .GetRS.EOF
                    
                    intFlushM = intFlushM + .GetRS("枚数")
                    
                    '1.10.7 ADD 製造日ベースの時は未確定は集計しない
                    If (inDateKbn = 1) Or (inDateKbn = 2 And .GetRS("確定") <> 0) Then
                    '1.10.7 ADD End
                        Select Case .GetRS("製造区分")
                            Case 1, 2, 3
                                If .GetRS("製造区分") = 2 Then intFkamachiM = intFkamachiM + .GetRS("枚数")
                                If .GetRS("製造区分") = 3 Then intKamachiM = intKamachiM + .GetRS("枚数")
                                '1.10.10 K.Asayama Change
                                'If IsThruGlass(.GetRS("登録時品番")) Then intThruM = intThruM + .GetRS("枚数")
                                If IsThruGlass(.GetRS("登録時品番")) Then
                                    If IsVertica(.GetRS("登録時品番")) Then
                                        intThruM = intThruM + IsVertica_Maisu(.GetRS("登録時品番"), .GetRS("枚数"))
                                    Else
                                        intThruM = intThruM + fncIntHalfGlassMirror_Maisu(.GetRS("登録時品番"), .GetRS("枚数"))
                                    End If
                                End If
                                '1.10.10 K.Asayama Change End
                                If IsPainted(.GetRS("登録時品番")) Then intPaintM = intPaintM + .GetRS("枚数")
                                If IsAir(.GetRS("登録時品番")) Then intAirM = intAirM + .GetRS("枚数")
                                If IsMonster(.GetRS("登録時品番")) Then intMonsterM = intMonsterM + .GetRS("枚数")
                                '1.10.8 K.Asayama ADD
                                'If IsVertica(.GetRS("登録時品番")) Then intVerticaM = intVerticaM + .GetRS("枚数")
                                If IsVertica(.GetRS("登録時品番")) Then intVerticaM = intVerticaM + IsVertica_Maisu(.GetRS("登録時品番"), .GetRS("枚数"))
                                '1.10.8 K.Asayama ADD End
                            Case 6
                                If IsStealth_Seizo_TEMP(.GetRS("登録時品番")) Then
                                    intStealthM = intStealthM + .GetRS("枚数")
                                Else
                                    intShitajiM = intShitajiM + .GetRS("枚数")
                                End If
                                'intShitajiM = intShitajiM + .GetRS("枚数")
                            Case 7
                                intStealthM = intStealthM + .GetRS("枚数")
                                
                        End Select
                    '1.10.7 ADD
                    End If
                    '1.10.7 ADD End
                    
                    '1.10.7 Change
                    'If .GetRS("確定") <> 0 Then intKakuteiM = intKakuteiM + .GetRS("枚数")
                    If .GetRS("確定") = 2 Then
                        intKakuteiM = intKakuteiM + .GetRS("枚数")
                    ElseIf .GetRS("確定") = 1 Then
                        intKakuteiTempM = intKakuteiTempM + .GetRS("枚数")
                    End If
                    '1.10.7 Change End
                    
                    .GetRS.MoveNext
                Loop
                
                Graphctl(i).SetTag (CStr(intFlushM))
                Graphctl(i).CaptionSet Graphctl(i).GetTag
                '1.10.8 ADD
                Graphctl(i).SetControlTipText Graphctl(i).GetTag
                '1.10.8 ADD End
                
                If intKakuteiM > 0 Then
                    Graphctl_Kakutei(i).SetTag (CStr(intKakuteiM))
                    Graphctl_Kakutei(i).myVisible (True)
                Else
                    Graphctl_Kakutei(i).SetTag "0"
                    Graphctl_Kakutei(i).myVisible (False)
                End If
                               
                
                Graphctl_Kakutei(i).CaptionSet Graphctl_Kakutei(i).GetTag
                '1.10.8 ADD
                Graphctl_Kakutei(i).SetControlTipText Graphctl_Kakutei(i).GetTag
                '1.10.8 ADD End
                
                '1.10.7 ADD
                If intKakuteiTempM + intKakuteiM > 0 Then
                    Graphctl_Temp(i).SetTag (CStr(intKakuteiM + intKakuteiTempM))
                    Graphctl_Temp(i).myVisible (True)
                Else
                    Graphctl_Temp(i).SetTag "0"
                    Graphctl_Temp(i).myVisible (False)
                End If
                
                Graphctl_Temp(i).CaptionSet Graphctl_Temp(i).GetTag
                '1.10.7 ADD End
                '1.10.8 ADD
                Graphctl_Temp(i).SetControlTipText Graphctl_Temp(i).GetTag
                '1.10.8 ADD End
                
                If intFkamachiM > 0 Then Itemctl(i, 0).myVisible (True): Itemctl(i, 0).SetControlTipText (intFkamachiM) Else Itemctl(i, 0).myVisible (False)
                If intThruM > 0 Then Itemctl(i, 1).myVisible (True): Itemctl(i, 1).SetControlTipText (intThruM) Else Itemctl(i, 1).myVisible (False)
                If intPaintM > 0 Then Itemctl(i, 2).myVisible (True): Itemctl(i, 2).SetControlTipText (intPaintM) Else Itemctl(i, 2).myVisible (False)
                If intAirM > 0 Then Itemctl(i, 3).myVisible (True): Itemctl(i, 3).SetControlTipText (intAirM) Else Itemctl(i, 3).myVisible (False)
                If intMonsterM > 0 Then Itemctl(i, 4).myVisible (True): Itemctl(i, 4).SetControlTipText (intMonsterM) Else Itemctl(i, 4).myVisible (False)
                '1.10.8 Change
                'If intKamachiM > 0 Then Itemctl(i, 5).myVisible (True): Itemctl(i, 5).SetControlTipText (intKamachiM) Else Itemctl(i, 5).myVisible (False)
                If intVerticaM > 0 Then Itemctl(i, 5).myVisible (True): Itemctl(i, 5).SetControlTipText (intVerticaM) Else Itemctl(i, 5).myVisible (False)
                '1.10.8 Change End
                
                If intShitajiM > 0 Then Itemctl(i, 1).myVisible (True): Itemctl(i, 1).SetControlTipText ("下地数"): Itemctl(i, 1).CaptionSet (CStr(intShitajiM)): Itemctl(i, 1).SetWidth (240)
                If intStealthM > 0 Then Itemctl(i, 4).myVisible (True): Itemctl(i, 4).SetControlTipText ("ステルス数"): Itemctl(i, 4).CaptionSet (CStr(intStealthM)): Itemctl(i, 4).SetWidth (240)
                
           End If
        End With
    Next
    
    GoTo Exit_SetOrderCount
    
Err_SetOrderCount:
    MsgBox Err.Description

Exit_SetOrderCount:
    Set objREMOTEdb = Nothing


End Function

Public Function fncbolSetComboKubun(inKubun As String, inCombobox As Access.ComboBox) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'コンボボックスセット（共通）
'
'   :引数
'       inKubun         :コンボボックス区分名
'       inCombobox      :コンボボックスオブジェクト
'
'   :戻り値
'       True            :成功
'       False           :失敗

'1.10.16 K.Asayama ADD
'   →ステルス区分追加
'--------------------------------------------------------------------------------------------------------------------
    On Error GoTo Err_fncbolSetComboKubun
    
    inCombobox.RowSourceType = "Value List"
    
    If inKubun = "製造区分" Then
        inCombobox.AddItem "建具,1", 0
        inCombobox.AddItem "枠,2", 1
        inCombobox.AddItem "下地,3", 2
        '1.10.16 ADD
        inCombobox.AddItem "ステルス,4", 3
        inCombobox.value = inCombobox.ItemData(0)
    End If
    
    
    fncbolSetComboKubun = True
    
    GoTo Exit_fncbolSetComboKubun
    
Err_fncbolSetComboKubun:
    fncbolSetComboKubun = False
    MsgBox Err.Description
    
Exit_fncbolSetComboKubun:
    
End Function

Public Function SetBikouData() As Boolean
'--------------------------------------------------------------------------------------------------------------------
'WK_札データ_備考ファイルを作成する
'
'
'   :戻り値
'       True            :成功
'       False           :失敗
'1.10.7 K.Asayama ADD 20160108
'       →作成したWK_札データから備考ファイルを作成する
'1.10.8 K.Asayama Change 20160114
'       →バグ修正 Firstだとうまくデータが出ないのでMaxに変更
'--------------------------------------------------------------------------------------------------------------------
    Dim objLOCALdb As New cls_LOCALDB
    
    Dim strSQL As String
    Dim strErrMsg As String
    
    SetBikouData = False
     
    On Error GoTo Err_SetBikouData
    
    strSQL = ""
    
    strSQL = strSQL & "select 契約番号,棟番号,部屋番号"
'1.10.8 Change
'    strSQL = strSQL & ",First(IIf([製造区分] = 1,[備考],Null)) as Flush備考 "
'    strSQL = strSQL & ",First(IIf([製造区分] = 2,[備考],Null)) as F框備考 "
'    strSQL = strSQL & ",First(IIf([製造区分] = 3,[備考],Null)) as 框備考 "
'    strSQL = strSQL & ",First(IIf([製造区分] = 4,[備考],Null)) as 枠備考 "
'    strSQL = strSQL & ",First(IIf([製造区分] = 5,[備考],Null)) as 三方枠備考 "
'    strSQL = strSQL & ",First(IIf([製造区分] = 6,[備考],Null)) as 下地備考 "
'    strSQL = strSQL & ",First(IIf([製造区分] = 7,[備考],Null)) as ステルス枠備考 "
    strSQL = strSQL & ",Max(IIf([製造区分] = 1,[備考],Null)) as Flush備考 "
    strSQL = strSQL & ",Max(IIf([製造区分] = 2,[備考],Null)) as F框備考 "
    strSQL = strSQL & ",Max(IIf([製造区分] = 3,[備考],Null)) as 框備考 "
    strSQL = strSQL & ",Max(IIf([製造区分] = 4,[備考],Null)) as 枠備考 "
    strSQL = strSQL & ",Max(IIf([製造区分] = 5,[備考],Null)) as 三方枠備考 "
    strSQL = strSQL & ",Max(IIf([製造区分] = 6,[備考],Null)) as 下地備考 "
    strSQL = strSQL & ",Max(IIf([製造区分] = 7,[備考],Null)) as ステルス枠備考 "
'1.10.8 Change End
    strSQL = strSQL & "from WK_札データ "
    strSQL = strSQL & "where 備考 is not null "
    strSQL = strSQL & "group by 契約番号,棟番号,部屋番号 "
    
    If Not objLOCALdb.ExecSQL("delete from WK_札データ_備考") Then
        Err.Raise 9999, , "備考データワーク（ローカル）初期化エラー"
    End If
    
    With objLOCALdb
        If .ExecSelect(strSQL) Then
            
            Do While Not .GetRS.EOF
                strSQL = "insert into WK_札データ_備考 ("
                strSQL = strSQL & "契約番号,棟番号,部屋番号 "
                strSQL = strSQL & ",Flush備考,F框備考,枠備考,三方枠備考,下地備考,ステルス枠備考"
                strSQL = strSQL & ") values ( "
                strSQL = strSQL & "'" & .GetRS![契約番号] & "','" & .GetRS![棟番号] & "','" & .GetRS![部屋番号] & "'"
                strSQL = strSQL & "," & varNullChk(.GetRS![Flush備考], 1) & " "
                strSQL = strSQL & "," & varNullChk(.GetRS![F框備考], 1) & " "
                strSQL = strSQL & "," & varNullChk(.GetRS![枠備考], 1) & " "
                strSQL = strSQL & "," & varNullChk(.GetRS![三方枠備考], 1) & " "
                strSQL = strSQL & "," & varNullChk(.GetRS![下地備考], 1) & " "
                strSQL = strSQL & "," & varNullChk(.GetRS![ステルス枠備考], 1) & " "
                strSQL = strSQL & ")"
                
                'Debug.Print strSQL
                
                If Not .ExecSQL(strSQL, strErrMsg) Then
                    Err.Raise 9999, , strErrMsg
                End If
                
                .GetRS.MoveNext
            Loop
        Else
            Err.Raise 9999, , "札データ（ローカル）オープンエラー(Input)"
        End If
    End With
    
    SetBikouData = True
    
    GoTo Exit_SetBikouData
    
Err_SetBikouData:
    SetBikouData = False
    MsgBox Err.Description
    
Exit_SetBikouData:
     Set objLOCALdb = Nothing
     
End Function

Public Function varNullChk(in_Data As Variant, in_DBType As Integer) As Variant
'--------------------------------------------------------------------------------------------------------------------
'引数がNullの場合は文字列[Null]を返す。それ以外はそのまま返す(DBインサート用)
'
'   :引数
'       in_Data     Variant(型不定 exデータベースのカラム）
'       in_DBType   1:Local(Jet) 2:SQLServer
'
'   :戻り値
'       Variant　   引数がNullの場合は文字列[Null]、それ以外はそのまま(日付、文字列は加工する）
'
'1.10.7 K.Asayama ADD 20160108
'
'1.10.16 Change
'       →  文字列と日付の精査順入れ替え
'           文字列でも日付変換可能なものは日付で変換する（SQLServerからダイレクトに列を受け取った場合型を文字列と認識してしまうため）
'           空欄の文字列はNullとする
'2.0.0
'       →  データ内に「'（アポストロフィ）」があった場合「''」複数に置換え

'2.1.0
'       →　アポストロフィは全角に置き換える
'       →  String値の日付が誤って判断される場合があるので修正
'--------------------------------------------------------------------------------------------------------------------
    Dim datDate As Date
    
    If IsNull(in_Data) Then
    
        varNullChk = "Null"
    
'    ElseIf VarType(in_Data) = vbDate Or (VarType(in_Data) = vbString And IsDate(in_Data)) Then
    ElseIf VarType(in_Data) = vbDate Or (VarType(in_Data) = vbString And (CStr(in_Data) Like "#*/#*/#*" Or CStr(in_Data) Like "#*-#*-#*")) Then
    
        If VarType(in_Data) = vbDate Then
            datDate = in_Data
            
            Select Case in_DBType
                Case 1
                    varNullChk = "#" & Format(datDate, "yyyy/mm/dd") & "#"
                Case Else
                    varNullChk = "'" & Format(datDate, "yyyy/mm/dd") & "'"
            End Select
            
        ElseIf IsDate(in_Data) Then
            datDate = CDate(in_Data)
            
            Select Case in_DBType
                Case 1
                    varNullChk = "#" & Format(datDate, "yyyy/mm/dd") & "#"
                Case Else
                    varNullChk = "'" & Format(datDate, "yyyy/mm/dd") & "'"
            End Select
            
        Else
            in_Data = Replace(in_Data, "'", "’")
            varNullChk = "'" & in_Data & "'"
        End If
        
        
        
    ElseIf VarType(in_Data) = vbString Then
        '1.10.16
        'varNullChk = "'" & in_Data & "'"
        If in_Data = "" Then
            varNullChk = "Null"
        Else
            '↓2.0.0 ADD
            in_Data = Replace(in_Data, "'", "’")
            varNullChk = "'" & in_Data & "'"
        End If
    Else
        varNullChk = in_Data
    End If

End Function

Public Function bolfncTableCopyToLocal(in_RS As ADODB.Recordset, out_LocalTableName As String, Optional in_ADDMode As Boolean = False) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'リモートDBのレコードセットをローカルのテーブルにコピーする
'   （リモートとローカルのカラム名は同じである前提）
'
'   :引数
'       in_RS                   リモートデータベースのレコードセット
'       out_LocalTableName      ローカルデータベースのテーブル名
'       in_ADDMode              True:追加 False:Replace（最初にレコードをDELETEする）
'
'   :戻り値
'       Boolean　               True:成功   False:失敗
'
'1.11.2 ADD
'--------------------------------------------------------------------------------------------------------------------

    Dim objLOCALdb As New cls_LOCALDB
    Dim i As Integer
    Dim strErrMsg As String
    Dim varAutoNumber As Variant
    
    Dim daoDB As DAO.Database
    Dim DAORs As DAO.Recordset
    
    Set daoDB = CurrentDb
    Set DAORs = daoDB.OpenRecordset(out_LocalTableName)
    
    On Error GoTo Err_bolfncTableCopyToLocal
    
    bolfncTableCopyToLocal = False
    
    'オートナンバー型チェック
    'オートナンバーは移送しない
    varAutoNumber = Null
    
    With DAORs
        For i = 0 To .Fields.Count - 1
            If (.Fields(i).Type = dbLong) And (.Fields(i).Attributes And dbAutoIncrField) Then
                varAutoNumber = .Fields(i).Name
                Exit For
            End If
        Next
    End With
    
    DAORs.Close
    daoDB.Close
    
    With objLOCALdb
    
        If Not in_ADDMode Then
            If Not .ExecSQL("delete * from " & out_LocalTableName & " ", strErrMsg) Then
                Err.Raise 9999, , strErrMsg
            Else
                'オートナンバー初期化
                If Not IsNull(varAutoNumber) Then
                    DoCmd.RunSQL "ALTER TABLE " & out_LocalTableName & " ALTER COLUMN " & varAutoNumber & " COUNTER(1, 1)"
                End If
                
            End If
        End If
        
        If .ExecSelect_Writable("select * from " & out_LocalTableName & " ") Then
        
            in_RS.MoveFirst
            
            Do While Not in_RS.EOF
                .GetRS.AddNew
                
                For i = 0 To .GetRS.Fields.Count - 1
                    If .GetRS.Fields(i).Name <> Nz(varAutoNumber, "") Then
                        .GetRS(.GetRS.Fields(i).Name) = in_RS(.GetRS.Fields(i).Name)
                    End If
                Next
                .GetRS.Update
               in_RS.MoveNext
            Loop
        End If
    End With
    
    bolfncTableCopyToLocal = True
    
    GoTo Exit_bolfncTableCopyToLocal
    
Err_bolfncTableCopyToLocal:
    MsgBox Err.Description
    
Exit_bolfncTableCopyToLocal:
    Set objLOCALdb = Nothing
    Set DAORs = Nothing
    Set daoDB = Nothing
End Function

Public Function bolfncMiseizoToExcel() As Boolean
'--------------------------------------------------------------------------------------------------------------------
'未製造データExcelへエクスポート
'1.12.2 ADD

'   :引数

'   :戻り値
'       True            :成功
'       False           :失敗

'2.2.0
'   →ウォールスルー未製造追加
'2.8.0
'   →クロゼット未出荷追加
'2.13.0
'   →サーバパスを共通変数に変更
'--------------------------------------------------------------------------------------------------------------------

    Dim objApp As New cls_Excel
    Dim objREMOTEdb As New cls_BRAND_MASTER
    
    Dim xlsBookName As String
    Dim i As Integer
    Dim intSheetDel As Integer
    Dim strSQL As String
    Dim strSQLJ As String
    Dim strKBName(4) As String
    Dim strMidashiVal As String
    
    On Error GoTo Err_bolfncMiseizoToExcel
    intSheetDel = 0
    
    Screen.MousePointer = 11
    
    With objApp.getExcel

        .Workbooks.Add
        
        strKBName(0) = "建具"
        strKBName(1) = "下地"
        strKBName(2) = "枠"
        strKBName(3) = "ウォールスルー未出荷"
        strKBName(4) = "クロゼット未出荷"
        
        strSQL = ""
        strSQL = strfncTextFileToString(conServerPath & "\SQL\subMISEIZO.sql")
        'strSQL = strfncTextFileToString("\\db\prog\製造管理システム\SQL\subMISEIZO.sql")
        If strSQL <> "" Then
            strSQL = Replace(strSQL, vbCrLf, " ")
        Else
            Err.Raise 9999, , "未製造出力異常終了"
        End If
        
        strMidashiVal = "建具未生産残 " & Format(Now, "yyyy-MM-dd")
        
        If Not objREMOTEdb.ExecSelect(strSQL) Then
            Err.Raise 9999, , "台帳集計データ異常終了"
        End If
        
        objApp.WorkSheetADD strKBName(i)
                
        If Not bolfncexp_EXCELOBJECT(objREMOTEdb.GetRS, objApp.getExcel, True, strMidashiVal) Then
            Err.Raise 9999, , "Excelエクスポート異常終了"
        End If
        
        strSQL = ""
        'strSQL = strfncTextFileToString("\\db\prog\製造管理システム\SQL\subMISEIZOWaku.sql")
        strSQL = strfncTextFileToString(conServerPath & "\SQL\subMISEIZOWaku.sql")
        If strSQL <> "" Then
            strSQL = Replace(strSQL, vbCrLf, " ")
        Else
            Err.Raise 9999, , "未製造出力異常終了"
        End If
        
        For i = 1 To 2
            strMidashiVal = "枠未生産残 (" & strKBName(i) & ") "
            
            strSQLJ = Replace(strSQL, "@WakuKubun", "'" & strKBName(i) & "'")

            strMidashiVal = strMidashiVal & " " & Format(Now, "yyyy-MM-dd")
            
            objApp.WorkSheetADD strKBName(i)
            If Not objREMOTEdb.ExecSelect(strSQLJ) Then
                Err.Raise 9999, , "台帳集計データ異常終了"
            End If
            
            If Not bolfncexp_EXCELOBJECT(objREMOTEdb.GetRS, objApp.getExcel, True, strMidashiVal) Then
                Err.Raise 9999, , "Excelエクスポート異常終了"
            End If

        Next
        
        i = 3
        strSQL = ""
        'strSQL = strfncTextFileToString("\\db\prog\製造管理システム\SQL\subMISHUKKA_Wallthru.sql")
        strSQL = strfncTextFileToString(conServerPath & "\SQL\subMISHUKKA_Wallthru.sql")
        If strSQL <> "" Then
            strSQL = Replace(strSQL, vbCrLf, " ")
        Else
            Err.Raise 9999, , "未製造出力異常終了"
        End If
        
        strMidashiVal = "ウォールスルー未出荷残 " & Format(Now, "yyyy-MM-dd")
        
        If Not objREMOTEdb.ExecSelect(strSQL) Then
            Err.Raise 9999, , "台帳集計データ異常終了"
        End If
        
        objApp.WorkSheetADD strKBName(i)
                
        If Not bolfncexp_EXCELOBJECT(objREMOTEdb.GetRS, objApp.getExcel, True, strMidashiVal) Then
            Err.Raise 9999, , "Excelエクスポート異常終了"
        End If
        
        i = 4
        strSQL = ""
        strSQL = strfncTextFileToString("\\db\prog\製造管理システム\SQL\subMISHUKKA_Oredo.sql")
        strSQL = strfncTextFileToString(conServerPath & "\SQL\subMISHUKKA_Oredo.sql")
        If strSQL <> "" Then
            strSQL = Replace(strSQL, vbCrLf, " ")
        Else
            Err.Raise 9999, , "未製造出力異常終了"
        End If
        
        strMidashiVal = "クロゼット未出荷残 " & Format(Now, "yyyy-MM-dd")
        
        If Not objREMOTEdb.ExecSelect(strSQL) Then
            Err.Raise 9999, , "台帳集計データ異常終了"
        End If
        
        objApp.WorkSheetADD strKBName(i)
                
        If Not bolfncexp_EXCELOBJECT(objREMOTEdb.GetRS, objApp.getExcel, True, strMidashiVal) Then
            Err.Raise 9999, , "Excelエクスポート異常終了"
        End If
        
        '不要なワークシートの削除
        For i = 1 To .Worksheets.Count
            If .Worksheets(i - intSheetDel).Name Like "Sheet*" Then
                .Worksheets(i - intSheetDel).Delete
                intSheetDel = intSheetDel + 1
            End If
        Next
        
        .Worksheets(1).Activate
        
        objApp.ContinueOpen = True
        
    End With
    
    bolfncMiseizoToExcel = True
    
    GoTo Exit_bolfncMiseizoToExcel

Err_bolfncMiseizoToExcel:
    Screen.MousePointer = 0
    MsgBox Err.Description
    bolfncMiseizoToExcel = False
    
    
Exit_bolfncMiseizoToExcel:
    Screen.MousePointer = 0
    Set objApp = Nothing
    Set objREMOTEdb = Nothing
    
    
End Function