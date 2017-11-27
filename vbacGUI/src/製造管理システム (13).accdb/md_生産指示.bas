Option Compare Database
Option Explicit
'2.1.0 ADD

Public Function bolfnc製造指示データ抽出(Optional SeizoDate As Date = #6/19/2017#) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'指定の製造日の指示データ、部材展開をワークファイルに作成する
'
'   :引数
'       製造日
'
'   :戻り値
'       True            :成功
'       False           :失敗

'--------------------------------------------------------------------------------------------------------------------

    Dim objREMOTEDB As New cls_BRAND_MASTER
    Dim objLocalDB As New cls_LOCALDB
    Dim objTateguHinban As New cls_建具品番
    
    Dim conSQL As String
    Dim strSQL As String
    Dim strErrMsg As String
    Dim intCnt As Integer
    Dim intMaisu As Integer
    Dim varBrandHinban As Variant
    Dim ChuuiFlg As Boolean
    
    bolfnc製造指示データ抽出 = False
    
    On Error GoTo Err_bolfnc製造指示データ抽出
    
    conSQL = conSQL & "insert into WK_製造依頼書建具 "
    conSQL = conSQL & "( "
    conSQL = conSQL & "通番 "
    conSQL = conSQL & ",契約番号 "
    conSQL = conSQL & ",棟番号 "
    conSQL = conSQL & ",部屋番号 "
    conSQL = conSQL & ",契約No "
    conSQL = conSQL & ",物件名 "
    conSQL = conSQL & ",施工店 "
    conSQL = conSQL & ",項 "
    conSQL = conSQL & ",設置場所 "
    conSQL = conSQL & ",品番1 "
    conSQL = conSQL & ",子扉 "
    conSQL = conSQL & ",商品名 "
    conSQL = conSQL & ",色 "
    conSQL = conSQL & ",色コード "
    conSQL = conSQL & ",吊元 "
    conSQL = conSQL & ",開閉様式 "
    conSQL = conSQL & ",施錠 "
    conSQL = conSQL & ",数量 "
    conSQL = conSQL & ",枚数 "
    conSQL = conSQL & ",品番区分 "
    conSQL = conSQL & ",製造区分 "
    conSQL = conSQL & ",工場CD "
    conSQL = conSQL & ",追加 "
    conSQL = conSQL & ",階 "
    conSQL = conSQL & ",旧ヒンジ "
    conSQL = conSQL & ",欠品FLG "
    conSQL = conSQL & ",種類 "
    conSQL = conSQL & ",工数 "
    conSQL = conSQL & ",旧仕様 "
    conSQL = conSQL & ",出荷方法 "
    conSQL = conSQL & ",設計備考 "
    conSQL = conSQL & ",08カタログ "
    conSQL = conSQL & ",標準ハンドル "
    conSQL = conSQL & ",DW "
    conSQL = conSQL & ",DH "
    conSQL = conSQL & ",CH "
    conSQL = conSQL & ",明り窓 "
    conSQL = conSQL & ",個別Spec "
    conSQL = conSQL & ",Spec "
    conSQL = conSQL & ",Style "
    conSQL = conSQL & ",受注明細Style "
    conSQL = conSQL & ",新面材割付 "
    conSQL = conSQL & ",注意FLG "
    conSQL = conSQL & ",物流施工日 "
    conSQL = conSQL & ",建具出荷方法 "
    conSQL = conSQL & ",クレーム用備考 "
    conSQL = conSQL & ",建具確定日 "
    conSQL = conSQL & ",製造日 "
    conSQL = conSQL & ") values ("
    conSQL = conSQL & "@通番@ "
    conSQL = conSQL & ",@契約番号@ "
    conSQL = conSQL & ",@棟番号@ "
    conSQL = conSQL & ",@部屋番号@ "
    conSQL = conSQL & ",@契約No@ "
    conSQL = conSQL & ",@物件名@ "
    conSQL = conSQL & ",@施工店@ "
    conSQL = conSQL & ",@項@ "
    conSQL = conSQL & ",@設置場所@ "
    conSQL = conSQL & ",@品番1@ "
    conSQL = conSQL & ",@子扉@ "
    conSQL = conSQL & ",@商品名@ "
    conSQL = conSQL & ",@色@ "
    conSQL = conSQL & ",@色コード@ "
    conSQL = conSQL & ",@吊元@ "
    conSQL = conSQL & ",@開閉様式@ "
    conSQL = conSQL & ",@施錠@ "
    conSQL = conSQL & ",@数量@ "
    conSQL = conSQL & ",@枚数@ "
    conSQL = conSQL & ",@品番区分@ "
    conSQL = conSQL & ",@製造区分@ "
    conSQL = conSQL & ",@工場CD@ "
    conSQL = conSQL & ",@追加@ "
    conSQL = conSQL & ",@階@ "
    conSQL = conSQL & ",@旧ヒンジ@ "
    conSQL = conSQL & ",@欠品FLG@ "
    conSQL = conSQL & ",@種類@ "
    conSQL = conSQL & ",@工数@ "
    conSQL = conSQL & ",@旧仕様@ "
    conSQL = conSQL & ",@出荷方法@ "
    conSQL = conSQL & ",@設計備考@ "
    conSQL = conSQL & ",@08カタログ@ "
    conSQL = conSQL & ",@標準ハンドル@ "
    conSQL = conSQL & ",@DW@ "
    conSQL = conSQL & ",@DH@ "
    conSQL = conSQL & ",@CH@ "
    conSQL = conSQL & ",@明り窓@ "
    conSQL = conSQL & ",@個別Spec@ "
    conSQL = conSQL & ",@Spec@ "
    conSQL = conSQL & ",@Style@ "
    conSQL = conSQL & ",@受注明細Style@ "
    conSQL = conSQL & ",@新面材割付@ "
    conSQL = conSQL & ",@注意FLG@ "
    conSQL = conSQL & ",@物流施工日@ "
    conSQL = conSQL & ",@建具出荷方法@ "
    conSQL = conSQL & ",@クレーム用備考@ "
    conSQL = conSQL & ",@建具確定日@ "
    conSQL = conSQL & ",@製造日@ "
    conSQL = conSQL & ") "
    
    strSQL = strSQL & "select "
    strSQL = strSQL & "case "
    strSQL = strSQL & "when replace(dbo.fncgethinban(品番1,特注建具品番),'特 ','') like 'S%(ZZ)' then 1 "
    strSQL = strSQL & "when  replace(dbo.fncgethinban(品番1,特注建具品番),'特 ','') like '%SA-[0-9][0-9][0-9][0-9]%' then 3 "
    strSQL = strSQL & "else 2 "
    strSQL = strSQL & "end ロット順1 "
    strSQL = strSQL & ", "
    strSQL = strSQL & "case "
    strSQL = strSQL & "when right(b.個別Spec,4) >= '1006' and b.色 = 'PW' then 0 "
    strSQL = strSQL & "else d.製造順 "
    strSQL = strSQL & "end ロット順2 "
    strSQL = strSQL & ",case "
    strSQL = strSQL & "when b.Style like 'S%' then 0 "
    strSQL = strSQL & "else 1 "
    strSQL = strSQL & "end ロット順3 "
    strSQL = strSQL & ",case "
    strSQL = strSQL & "when replace(dbo.fncgethinban(品番1,特注建具品番),'特 ','') like '%-[0-9][0-9][0-9][0-9]C%' then 1 "
    strSQL = strSQL & "when replace(dbo.fncgethinban(品番1,特注建具品番),'特 ','') like '%-[0-9][0-9][0-9][0-9]S[TSG]%' then 2 "
    strSQL = strSQL & "when replace(dbo.fncgethinban(品番1,特注建具品番),'特 ','') like '%-[0-9][0-9][0-9][0-9]G%' then 3 "
    strSQL = strSQL & "when replace(dbo.fncgethinban(品番1,特注建具品番),'特 ','') like '%-[0-9][0-9][0-9][0-9]MF%' then 3 "
    strSQL = strSQL & "when replace(dbo.fncgethinban(品番1,特注建具品番),'特 ','') like '%-[0-9][0-9][0-9][0-9]D%' then 3 "
    strSQL = strSQL & "else 4 "
    strSQL = strSQL & "end ロット順4 "
    strSQL = strSQL & ",case when replace(dbo.fncgethinban(品番1,特注建具品番),'特 ','') like '%DA-[0-9][0-9][0-9][0-9]%' then 1 "
    strSQL = strSQL & "when replace(dbo.fncgethinban(品番1,特注建具品番),'特 ','') like '%DAS-[0-9][0-9][0-9][0-9]%' then 1 "
    strSQL = strSQL & "when replace(dbo.fncgethinban(品番1,特注建具品番),'特 ','') like '%DO-[0-9][0-9][0-9][0-9]%' then 1 "
    strSQL = strSQL & "when replace(dbo.fncgethinban(品番1,特注建具品番),'特 ','') like '%DOS-[0-9][0-9][0-9][0-9]%' then 1 "
    strSQL = strSQL & "when replace(dbo.fncgethinban(品番1,特注建具品番),'特 ','') like '%DK-[0-9][0-9][0-9][0-9]%' then 1 "
    strSQL = strSQL & "when replace(dbo.fncgethinban(品番1,特注建具品番),'特 ','') like '%DKS-[0-9][0-9][0-9][0-9]%' then 1 "
    strSQL = strSQL & "else 2 end ロット順5 "
    strSQL = strSQL & ",d.色名称 "
    strSQL = strSQL & ",a.契約番号,a.棟番号,a.部屋番号 "
    strSQL = strSQL & ",a.契約番号 + '-' + a.棟番号 + '-' + a.部屋番号 契約No "
    strSQL = strSQL & ",物件名 "
    strSQL = strSQL & ",施工店 "
    strSQL = strSQL & ",発注元 "
    strSQL = strSQL & ",[2次問屋] "
    strSQL = strSQL & ",c.項 "
    strSQL = strSQL & ",設置場所 "
    strSQL = strSQL & ",dbo.fncgethinban(品番1,特注建具品番) 品番1 "
    strSQL = strSQL & ",dbo.fncgethinban(子扉品番,特注子扉品番) 子扉品番 "
    strSQL = strSQL & ",商品名 "
    strSQL = strSQL & ",色 "
    strSQL = strSQL & ",吊元 "
    strSQL = strSQL & ",施錠 "
    strSQL = strSQL & ",b.数量 "
    strSQL = strSQL & ",枚数 "
    strSQL = strSQL & ",c.品番区分 "
    strSQL = strSQL & ",c.製造区分 "
    strSQL = strSQL & ",b.工場CD "
    strSQL = strSQL & ",b.追加 "
    strSQL = strSQL & ",階 "
    strSQL = strSQL & ",旧ヒンジ "
    strSQL = strSQL & ",欠品FLG "
    strSQL = strSQL & ",Null 種類 "
    strSQL = strSQL & ",1 工数 "
    strSQL = strSQL & ",旧仕様 "
    strSQL = strSQL & ",null 出荷方法 "
    strSQL = strSQL & ",建具設計備考 設計備考 "
    strSQL = strSQL & ",[08カタログ] "
    strSQL = strSQL & ",標準ハンドル "
    strSQL = strSQL & ",DW "
    strSQL = strSQL & ",DW子扉 "
    strSQL = strSQL & ",DH "
    strSQL = strSQL & ",明り窓 "
    strSQL = strSQL & ",case  "
    strSQL = strSQL & " when replace(dbo.fncgethinban(品番1,特注建具品番),'特 ','') like 'K_SD%' then FL上枠下地外H "
    strSQL = strSQL & " else FL仕上りH "
    strSQL = strSQL & " end CH "
    strSQL = strSQL & ",個別Spec "
    strSQL = strSQL & ",a.Spec "
    strSQL = strSQL & ",style 受注明細style "
    strSQL = strSQL & ",新面材割付 "
    strSQL = strSQL & ",物流施工日  "
    strSQL = strSQL & ",dbo.fncNohinHaisoCode(b.契約番号,b.棟番号,b.部屋番号,b.項,1) 建具出荷方法 "
    strSQL = strSQL & ",クレーム用備考 "
    strSQL = strSQL & ",建具確定日 "
    strSQL = strSQL & ",c.製造日 "
    strSQL = strSQL & "from T_受注ﾏｽﾀ a "
    strSQL = strSQL & "inner join T_受注ﾏｽﾀ_2 a2 "
    strSQL = strSQL & "on a.契約番号 = a2.契約番号 and a.棟番号 = a2.棟番号 and a.部屋番号 = a2.部屋番号 "
    strSQL = strSQL & "inner join T_受注明細 b "
    strSQL = strSQL & "on a.契約番号 = b.契約番号 and a.棟番号 = b.棟番号 and a.部屋番号 = b.部屋番号 "
    strSQL = strSQL & "inner join T_製造指示 c "
    strSQL = strSQL & "on c.契約番号 = b.契約番号 and c.棟番号 = b.棟番号 and c.部屋番号 = b.部屋番号 and c.項 = b.項 "
    strSQL = strSQL & "left join T_色記号ﾏｽﾀ d "
    strSQL = strSQL & "on b.色 = d.色記号 "
    strSQL = strSQL & "where  c.製造日 between " & varNullChk(SeizoDate, 2) & " and " & varNullChk(SeizoDate, 2) & " "
    strSQL = strSQL & "and c.確定 = 2 "
    strSQL = strSQL & "and 製造区分 in (1,2,3) "
    strSQL = strSQL & "order by ロット順1 "
    strSQL = strSQL & ",ロット順2"
    strSQL = strSQL & ",ロット順3"
    strSQL = strSQL & ",ロット順4"
    strSQL = strSQL & ",ロット順5"
    strSQL = strSQL & ",品番1"
    strSQL = strSQL & ",吊元"
    strSQL = strSQL & ",b.契約番号,b.棟番号,b.部屋番号,b.項 "
    
    If Not objLocalDB.ExecSQL("delete * from WK_製造依頼書建具 ", strErrMsg) Then
        Err.Raise 9999, , strErrMsg
    End If
    
    With objREMOTEDB

        If .ExecSelect(strSQL) Then
            If Not .GetRS.EOF Then
            
                intCnt = 1
                
                Do Until .GetRS.EOF
                    
                    intMaisu = 0
                    varBrandHinban = ""
                    
                    If Not IsNull(.GetRS![品番1]) Then
                        
                        If Not IsSxL(Nz(.GetRS![品番1], ""), varBrandHinban) Then
                            varBrandHinban = .GetRS![品番1]
                        End If
                        
                        If Not IsNull(.GetRS![子扉品番]) Then
                            intMaisu = .GetRS![枚数] / 2
                        Else
                            intMaisu = .GetRS![枚数]
                        End If
                        
                        If .GetRS![発注元] Like "*水戸建親*" Or .GetRS![2次問屋] Like "*チャネル*" Then
                            ChuuiFlg = True
                        Else
                            ChuuiFlg = False
                        End If
                        
                        strSQL = conSQL
                        
                        If .GetRS![工場CD] = 10 Then
                            strSQL = Replace(strSQL, "@通番@", intCnt)
                            intCnt = intCnt + 1
                        Else
                            strSQL = Replace(strSQL, "@通番@", "Null")
                        End If
                        
                        strSQL = Replace(strSQL, "@契約番号@", varNullChk(.GetRS![契約番号], 1))
                        strSQL = Replace(strSQL, "@棟番号@", varNullChk(.GetRS![棟番号], 1))
                        strSQL = Replace(strSQL, "@部屋番号@", varNullChk(.GetRS![部屋番号], 1))
                        strSQL = Replace(strSQL, "@契約No@", varNullChk(.GetRS![契約No], 1))
                        strSQL = Replace(strSQL, "@物件名@", varNullChk(.GetRS![物件名], 1))
                        strSQL = Replace(strSQL, "@施工店@", varNullChk(.GetRS![施工店], 1))
                        strSQL = Replace(strSQL, "@項@", varNullChk(.GetRS![項], 1))
                        strSQL = Replace(strSQL, "@設置場所@", varNullChk(.GetRS![設置場所], 1))
                        strSQL = Replace(strSQL, "@品番1@", varNullChk(.GetRS![品番1], 1))
                        strSQL = Replace(strSQL, "@商品名@", varNullChk(.GetRS![商品名], 1))
                        strSQL = Replace(strSQL, "@色@", varNullChk(.GetRS![色名称], 1))
                        strSQL = Replace(strSQL, "@色コード@", varNullChk(.GetRS![色], 1))
                        strSQL = Replace(strSQL, "@吊元@", varNullChk(.GetRS![吊元], 1))
                        strSQL = Replace(strSQL, "@開閉様式@", varNullChk(objTateguHinban.開閉様式(varBrandHinban), 1))
                        strSQL = Replace(strSQL, "@施錠@", varNullChk(.GetRS![施錠], 1))
                        strSQL = Replace(strSQL, "@数量@", varNullChk(.GetRS![数量], 1))
                        strSQL = Replace(strSQL, "@枚数@", varNullChk(intMaisu, 1))
                        strSQL = Replace(strSQL, "@品番区分@", varNullChk(.GetRS![品番区分], 1))
                        strSQL = Replace(strSQL, "@製造区分@", varNullChk(.GetRS![製造区分], 1))
                        strSQL = Replace(strSQL, "@工場CD@", varNullChk(.GetRS![工場CD], 1))
                        strSQL = Replace(strSQL, "@追加@", varNullChk(.GetRS![追加], 1))
                        strSQL = Replace(strSQL, "@階@", varNullChk(.GetRS![階], 1))
                        strSQL = Replace(strSQL, "@旧ヒンジ@", varNullChk(.GetRS![旧ヒンジ], 1))
                        strSQL = Replace(strSQL, "@欠品FLG@", varNullChk(.GetRS![欠品FLG], 1))
                        strSQL = Replace(strSQL, "@種類@", varNullChk(.GetRS![種類], 1))
                        strSQL = Replace(strSQL, "@工数@", varNullChk(.GetRS![工数], 1))
                        strSQL = Replace(strSQL, "@旧仕様@", varNullChk(.GetRS![旧仕様], 1))
                        strSQL = Replace(strSQL, "@出荷方法@", varNullChk(.GetRS![出荷方法], 1))
                        strSQL = Replace(strSQL, "@設計備考@", varNullChk(.GetRS![設計備考], 1))
                        strSQL = Replace(strSQL, "@08カタログ@", varNullChk(.GetRS![08カタログ], 1))
                        strSQL = Replace(strSQL, "@標準ハンドル@", varNullChk(.GetRS![標準ハンドル], 1))
                        strSQL = Replace(strSQL, "@DW@", varNullChk(.GetRS![DW], 1))
                        strSQL = Replace(strSQL, "@DH@", varNullChk(.GetRS![DH], 1))
                        strSQL = Replace(strSQL, "@CH@", varNullChk(.GetRS![CH], 1))
                        strSQL = Replace(strSQL, "@明り窓@", varNullChk(.GetRS![明り窓], 1))
                        strSQL = Replace(strSQL, "@個別Spec@", varNullChk(.GetRS![個別Spec], 1))
                        strSQL = Replace(strSQL, "@Spec@", varNullChk(.GetRS![Spec], 1))
                        strSQL = Replace(strSQL, "@受注明細Style@", varNullChk(.GetRS![受注明細Style], 1))
                        strSQL = Replace(strSQL, "@style@", varNullChk(objTateguHinban.Style(varBrandHinban), 1))
                        strSQL = Replace(strSQL, "@新面材割付@", varNullChk(.GetRS![新面材割付], 1))
                        strSQL = Replace(strSQL, "@注意FLG@ ", varNullChk(ChuuiFlg, 1))
                        strSQL = Replace(strSQL, "@物流施工日@", varNullChk(.GetRS![物流施工日], 1))
                        strSQL = Replace(strSQL, "@建具出荷方法@", varNullChk(.GetRS![建具出荷方法], 1))
                        strSQL = Replace(strSQL, "@クレーム用備考@", varNullChk(.GetRS![クレーム用備考], 1))
                        strSQL = Replace(strSQL, "@建具確定日@", varNullChk(.GetRS![建具確定日], 1))
                        strSQL = Replace(strSQL, "@製造日@", varNullChk(.GetRS![製造日], 1))
                        strSQL = Replace(strSQL, "@子扉@", False)

                        If Not objLocalDB.ExecSQL(strSQL, strErrMsg) Then
                            Err.Raise 9999, , strErrMsg
                        End If
                    
                    End If
                    
                    If Not IsNull(.GetRS![子扉品番]) Then
                        
                        If Not IsSxL(.GetRS![子扉品番], varBrandHinban) Then
                            varBrandHinban = .GetRS![子扉品番]
                        End If
                        
                        If intMaisu = 0 Then
                            intMaisu = .GetRS![枚数]
                        End If
                        
                        strSQL = conSQL
                        
                        If .GetRS![工場CD] = 10 Then
                            strSQL = Replace(strSQL, "@通番@", intCnt)
                            intCnt = intCnt + 1
                        Else
                            strSQL = Replace(strSQL, "@通番@", "Null")
                        End If
                        
                        strSQL = Replace(strSQL, "@契約番号@", varNullChk(.GetRS![契約番号], 1))
                        strSQL = Replace(strSQL, "@棟番号@", varNullChk(.GetRS![棟番号], 1))
                        strSQL = Replace(strSQL, "@部屋番号@", varNullChk(.GetRS![部屋番号], 1))
                        strSQL = Replace(strSQL, "@契約No@", varNullChk(.GetRS![契約No], 1))
                        strSQL = Replace(strSQL, "@物件名@", varNullChk(.GetRS![物件名], 1))
                        strSQL = Replace(strSQL, "@施工店@", varNullChk(.GetRS![施工店], 1))
                        strSQL = Replace(strSQL, "@項@", varNullChk(.GetRS![項], 1))
                        strSQL = Replace(strSQL, "@設置場所@", varNullChk(.GetRS![設置場所], 1))
                        strSQL = Replace(strSQL, "@品番1@", varNullChk(.GetRS![子扉品番], 1))
                        strSQL = Replace(strSQL, "@商品名@", varNullChk(.GetRS![商品名], 1))
                        strSQL = Replace(strSQL, "@色@", varNullChk(.GetRS![色名称], 1))
                        strSQL = Replace(strSQL, "@色コード@", varNullChk(.GetRS![色], 1))
                        strSQL = Replace(strSQL, "@吊元@", varNullChk(.GetRS![吊元], 1))
                        strSQL = Replace(strSQL, "@開閉様式@", varNullChk(objTateguHinban.開閉様式(varBrandHinban), 1))
                        strSQL = Replace(strSQL, "@施錠@", varNullChk(.GetRS![施錠], 1))
                        strSQL = Replace(strSQL, "@数量@", varNullChk(.GetRS![数量], 1))
                        strSQL = Replace(strSQL, "@枚数@", varNullChk(intMaisu, 1))
                        strSQL = Replace(strSQL, "@品番区分@", varNullChk(.GetRS![品番区分], 1))
                        strSQL = Replace(strSQL, "@製造区分@", varNullChk(.GetRS![製造区分], 1))
                        strSQL = Replace(strSQL, "@工場CD@", varNullChk(.GetRS![工場CD], 1))
                        strSQL = Replace(strSQL, "@追加@", varNullChk(.GetRS![追加], 1))
                        strSQL = Replace(strSQL, "@階@", varNullChk(.GetRS![階], 1))
                        strSQL = Replace(strSQL, "@旧ヒンジ@", varNullChk(.GetRS![旧ヒンジ], 1))
                        strSQL = Replace(strSQL, "@欠品FLG@", varNullChk(.GetRS![欠品FLG], 1))
                        strSQL = Replace(strSQL, "@種類@", varNullChk(.GetRS![種類], 1))
                        strSQL = Replace(strSQL, "@工数@", varNullChk(.GetRS![工数], 1))
                        strSQL = Replace(strSQL, "@旧仕様@", varNullChk(.GetRS![旧仕様], 1))
                        strSQL = Replace(strSQL, "@出荷方法@", varNullChk(.GetRS![出荷方法], 1))
                        strSQL = Replace(strSQL, "@設計備考@", varNullChk(.GetRS![設計備考], 1))
                        strSQL = Replace(strSQL, "@08カタログ@", varNullChk(.GetRS![08カタログ], 1))
                        strSQL = Replace(strSQL, "@標準ハンドル@", varNullChk(.GetRS![標準ハンドル], 1))
                        strSQL = Replace(strSQL, "@DW@", varNullChk(.GetRS![DW子扉], 1))
                        strSQL = Replace(strSQL, "@DH@", varNullChk(.GetRS![DH], 1))
                        strSQL = Replace(strSQL, "@CH@", varNullChk(.GetRS![CH], 1))
                        strSQL = Replace(strSQL, "@明り窓@", varNullChk(.GetRS![明り窓], 1))
                        strSQL = Replace(strSQL, "@個別Spec@", varNullChk(.GetRS![個別Spec], 1))
                        strSQL = Replace(strSQL, "@Spec@", varNullChk(.GetRS![Spec], 1))
                        strSQL = Replace(strSQL, "@受注明細Style@", varNullChk(.GetRS![受注明細Style], 1))
                        strSQL = Replace(strSQL, "@Style@", varNullChk(objTateguHinban.Style(varBrandHinban), 1))
                        strSQL = Replace(strSQL, "@新面材割付@", varNullChk(.GetRS![新面材割付], 1))
                        strSQL = Replace(strSQL, "@注意FLG@ ", varNullChk(ChuuiFlg, 1))
                        strSQL = Replace(strSQL, "@物流施工日@", varNullChk(.GetRS![物流施工日], 1))
                        strSQL = Replace(strSQL, "@建具出荷方法@", varNullChk(.GetRS![建具出荷方法], 1))
                        strSQL = Replace(strSQL, "@クレーム用備考@", varNullChk(.GetRS![クレーム用備考], 1))
                        strSQL = Replace(strSQL, "@建具確定日@", varNullChk(.GetRS![建具確定日], 1))
                        strSQL = Replace(strSQL, "@製造日@", varNullChk(.GetRS![製造日], 1))
                        strSQL = Replace(strSQL, "@子扉@", True)
                        
                        If Not objLocalDB.ExecSQL(strSQL, strErrMsg) Then
                            Err.Raise 9999, , strErrMsg
                        End If
                    End If
                    
                    .GetRS.MoveNext
                Loop
            End If
        Else
            Err.Raise 9999, , "製造指示データがありません"
        End If
    End With
    
    If DCount("*", "WK_製造依頼書建具") = 0 Then
        Err.Raise 9999, , "製造データがありません"
    End If
    
    If Not 部材展開WK作成(SeizoDate) Then
        Err.Raise 9999, , "製造データ作成異常終了（部材展開ワーク作成）"
    End If
    
    If Not 部材展開WK符号更新() Then
        Err.Raise 9999, , "製造データ作成異常終了（部材展開符号追加）"
    End If
    
    bolfnc製造指示データ抽出 = True
    
    GoTo Exit_bolfnc製造指示データ抽出
    
Err_bolfnc製造指示データ抽出:
    MsgBox Err.Description
    Debug.Print strSQL
    
    'Resume
Exit_bolfnc製造指示データ抽出:
    Set objREMOTEDB = Nothing
    Set objLocalDB = Nothing
    Set objTateguHinban = Nothing

    
End Function

Private Function 部材展開WK作成(ByVal inSeizoDate As Date) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'指定の製造日の部材展開をワークファイルに作成する
'
'
'   :戻り値
'       True            :成功
'       False           :失敗

'--------------------------------------------------------------------------------------------------------------------
    Dim objREMOTEDB As New cls_BRAND_MASTER
    
    Dim strSQL As String
    
    On Error GoTo Err_部材展開WK作成
    
    strSQL = ""
    strSQL = strSQL & "select b.契約番号 "
    strSQL = strSQL & ",b.棟番号 "
    strSQL = strSQL & ",b.部屋番号 "
    strSQL = strSQL & ",b.項 "
    strSQL = strSQL & ",邸名 "
    strSQL = strSQL & ",品番 "
    strSQL = strSQL & ",特注品番 "
    strSQL = strSQL & ",品名 "
    strSQL = strSQL & ",商品CD "
    strSQL = strSQL & ",部材種別CD "
    strSQL = strSQL & ",部材名 "
    strSQL = strSQL & ",同梱 "
    strSQL = strSQL & ",同梱品 "
    strSQL = strSQL & ",取付 "
    strSQL = strSQL & ",b.数量 "
    strSQL = strSQL & ",部材数 "
    strSQL = strSQL & ",部材数合計 "
    strSQL = strSQL & ",枚数 "
    strSQL = strSQL & ",b.品番区分 "
    strSQL = strSQL & ",b.製造区分 "
    strSQL = strSQL & ",追加 "
    strSQL = strSQL & ",キャンセル "
    strSQL = strSQL & ",色 "
    strSQL = strSQL & ",単位 "
    strSQL = strSQL & ",クレーム "
    strSQL = strSQL & ",メーカーCD "
    strSQL = strSQL & ",金物色 "
    strSQL = strSQL & ",日時 "
    strSQL = strSQL & ",[PC名] "
    strSQL = strSQL & ",[No] "
    strSQL = strSQL & ",Null as 符号 "
    strSQL = strSQL & ",case when dbo.IsKotobira(b.品番) = 1 then 1 else 0 end 子扉 "
    strSQL = strSQL & "from T_製造指示 a "
    strSQL = strSQL & "inner join BRAND_BOM.dbo.T_部材展開 b "
    strSQL = strSQL & "on a.契約番号 = b.契約番号 and a.棟番号 = b.棟番号 and a.部屋番号 = b.部屋番号 and a.項 = b.項 "
    strSQL = strSQL & "where a.製造日 between " & varNullChk(inSeizoDate, 2) & " and " & varNullChk(inSeizoDate, 2) & " "
    strSQL = strSQL & "and a.確定 = 2 "
    strSQL = strSQL & "and a.製造区分 between 1 and 3 "
    strSQL = strSQL & "and b.製造区分 in (1,2,3) "
    
    With objREMOTEDB
    
        If .ExecSelect(strSQL) Then
            If Not bolfncTableCopyToLocal(.GetRS, "WK_部材展開") Then
                Err.Raise 9999, , "部材展開コピー異常終了 "
            End If
        Else
            Err.Raise 9999, , "SQL実行エラー SQL = " & strSQL
        End If
    End With
    
    部材展開WK作成 = True
    
        
    GoTo Exit_部材展開WK作成

Err_部材展開WK作成:
    MsgBox Err.Description

Exit_部材展開WK作成:
    Set objREMOTEDB = Nothing
    
End Function

Private Function 部材展開WK符号更新() As Boolean
'--------------------------------------------------------------------------------------------------------------------
'部材展開をワークファイルに資材のDBより符号コードを更新する
'
'
'   :戻り値
'       True            :成功
'       False           :失敗

'--------------------------------------------------------------------------------------------------------------------
    Dim objLocalDB As New cls_LOCALDB
    Dim objSKAMIYADB As New cls_SKAMIYADB
    
    Dim strSQL As String
    
    On Error GoTo Err_部材展開WK符号更新
    
    strSQL = ""
    strSQL = strSQL & "select 商品CD,符号 "
    strSQL = strSQL & "from WK_部材展開 "
    strSQL = strSQL & "where 商品CD is not null "
    
    
    With objLocalDB
        If .ExecSelect_Writable(strSQL) Then
            
            If Not .GetRS.EOF Then
                Do Until .GetRS.EOF
                    strSQL = ""
                    strSQL = strSQL & "select 符号 from BRAND.T_KANAMONO_FUGO "
                    strSQL = strSQL & "where 部材CD = 'B' and 商品CD = '" & .GetRS![商品CD] & "' "
                    
                    If objSKAMIYADB.ExecSelect(strSQL) Then
                        If Not objSKAMIYADB.GetRS.EOF Then
                            .GetRS![符号] = objSKAMIYADB.GetRS![符号]
                            .GetRS.Update
                        End If
                    Else
                        Err.Raise 9999, , "部材展開WK符号更新エラー SQL = " & strSQL
                    End If
                    
                    objSKAMIYADB.RecordSetClose
                
                    .GetRS.MoveNext
                Loop
                
            End If
        Else
            Err.Raise 9999, , "部材展開WK符号更新 SQL実行エラー SQL = " & strSQL
        End If
    End With
    
    部材展開WK符号更新 = True
        
        
    GoTo Exit_部材展開WK符号更新

Err_部材展開WK符号更新:
    MsgBox Err.Description

Exit_部材展開WK符号更新:
    Set objSKAMIYADB = Nothing
    Set objLocalDB = Nothing
    
End Function

Public Function bolfnc製造指示_フルハイト() As Boolean
'--------------------------------------------------------------------------------------------------------------------
'当日の製造指示データからフルハイトライン用の指示書を出力する
'
'
'   :戻り値
'       True            :成功
'       False           :失敗

'--------------------------------------------------------------------------------------------------------------------
    Dim objREMOTEDB As cls_BRAND_MASTER
    Dim objLocalDB As New cls_LOCALDB
    Dim objTateguHinban As New cls_建具品番
    Dim objFullHeight As New Cls_FullHeight
    
    Dim rsADO As New ADODB.Recordset
    
    Dim strSQL As String
    Dim conSQL As String
    Dim conSQLW As String
    
    Dim strErrMsg As String
    
    Dim inRecordCount As Long
    Dim inReadCount As Long
    Dim intCnt As Integer
    
    Dim strBrandHinban As String
    Dim strOEMHinban As String
    
    Dim strTateguTsurimoto() As String
    Dim intTateguShurui As Integer
    Dim intOyako As Integer
    
    On Error GoTo Err_bolfnc製造指示_フルハイト
    
    conSQL = conSQL & "insert into WK_製造依頼書建具_FullHeight "
    conSQL = conSQL & "( "
    conSQL = conSQL & "通番 "
    conSQL = conSQL & ",契約番号 "
    conSQL = conSQL & ",棟番号 "
    conSQL = conSQL & ",部屋番号 "
    conSQL = conSQL & ",契約No "
    conSQL = conSQL & ",物件名 "
    conSQL = conSQL & ",施工店 "
    conSQL = conSQL & ",項 "
    conSQL = conSQL & ",扉位置 "
    conSQL = conSQL & ",設置場所 "
    conSQL = conSQL & ",品番1 "
    conSQL = conSQL & ",開閉様式 "
    conSQL = conSQL & ",商品名 "
    conSQL = conSQL & ",色 "
    conSQL = conSQL & ",色コード "
    conSQL = conSQL & ",吊元 "
    conSQL = conSQL & ",施錠 "
    conSQL = conSQL & ",数量 "
    conSQL = conSQL & ",枚数 "
    conSQL = conSQL & ",受注枚数 "
    conSQL = conSQL & ",品番区分 "
    conSQL = conSQL & ",製造区分 "
    conSQL = conSQL & ",工場CD "
    conSQL = conSQL & ",追加 "
    conSQL = conSQL & ",階 "
    conSQL = conSQL & ",旧ヒンジ "
    conSQL = conSQL & ",欠品FLG "
    conSQL = conSQL & ",種類 "
    conSQL = conSQL & ",工数 "
    conSQL = conSQL & ",旧仕様 "
    conSQL = conSQL & ",出荷方法 "
    conSQL = conSQL & ",設計備考 "
    conSQL = conSQL & ",08カタログ "
    conSQL = conSQL & ",標準ハンドル "
    conSQL = conSQL & ",DW "
    conSQL = conSQL & ",DH "
    conSQL = conSQL & ",CH "
    conSQL = conSQL & ",明り窓 "
    conSQL = conSQL & ",個別Spec "
    conSQL = conSQL & ",Spec "
    conSQL = conSQL & ",Style "
    conSQL = conSQL & ",新面材割付 "
    conSQL = conSQL & ",注意FLG "
    conSQL = conSQL & ",物流施工日 "
    conSQL = conSQL & ",建具出荷方法 "
    conSQL = conSQL & ",クレーム用備考 "
    conSQL = conSQL & ",建具確定日 "
    conSQL = conSQL & ",製造日 "
    
    conSQL = conSQL & ") values ("
    conSQL = conSQL & "@通番@ "
    conSQL = conSQL & ",@契約番号@ "
    conSQL = conSQL & ",@棟番号@ "
    conSQL = conSQL & ",@部屋番号@ "
    conSQL = conSQL & ",@契約No@ "
    conSQL = conSQL & ",@物件名@ "
    conSQL = conSQL & ",@施工店@ "
    conSQL = conSQL & ",@項@ "
    conSQL = conSQL & ",@扉位置@ "
    conSQL = conSQL & ",@設置場所@ "
    conSQL = conSQL & ",@品番1@ "
    conSQL = conSQL & ",@開閉様式@ "
    conSQL = conSQL & ",@商品名@ "
    conSQL = conSQL & ",@色@ "
    conSQL = conSQL & ",@色コード@ "
    conSQL = conSQL & ",@吊元@ "
    conSQL = conSQL & ",@施錠@ "
    conSQL = conSQL & ",@数量@ "
    conSQL = conSQL & ",@枚数@ "
    conSQL = conSQL & ",@受注枚数@ "
    conSQL = conSQL & ",@品番区分@ "
    conSQL = conSQL & ",@製造区分@ "
    conSQL = conSQL & ",@工場CD@ "
    conSQL = conSQL & ",@追加@ "
    conSQL = conSQL & ",@階@ "
    conSQL = conSQL & ",@旧ヒンジ@ "
    conSQL = conSQL & ",@欠品FLG@ "
    conSQL = conSQL & ",@種類@ "
    conSQL = conSQL & ",@工数@ "
    conSQL = conSQL & ",@旧仕様@ "
    conSQL = conSQL & ",@出荷方法@ "
    conSQL = conSQL & ",@設計備考@ "
    conSQL = conSQL & ",@08カタログ@ "
    conSQL = conSQL & ",@標準ハンドル@ "
    conSQL = conSQL & ",@DW@ "
    conSQL = conSQL & ",@DH@ "
    conSQL = conSQL & ",@CH@ "
    conSQL = conSQL & ",@明り窓@ "
    conSQL = conSQL & ",@個別Spec@ "
    conSQL = conSQL & ",@Spec@ "
    conSQL = conSQL & ",@Style@ "
    conSQL = conSQL & ",@新面材割付@ "
    conSQL = conSQL & ",@注意FLG@ "
    conSQL = conSQL & ",@物流施工日@ "
    conSQL = conSQL & ",@建具出荷方法@ "
    conSQL = conSQL & ",@クレーム用備考@ "
    conSQL = conSQL & ",@建具確定日@ "
    conSQL = conSQL & ",@製造日@ "
    conSQL = conSQL & ") "
    
    strSQL = "select * from WK_製造依頼書建具 "
        
    conSQLW = conSQLW & "where 工場CD = 10 "
    conSQLW = conSQLW & "order by 通番 "
        
    
    With objLocalDB
        If Not .ExecSQL("delete * from WK_製造依頼書建具_Fullheight") Then Err.Raise 9999, , "建具製造リスト_ローカル削除エラー"
        
        .CursorLocation = adUseClient
        
        strSQL = strSQL & conSQLW
        If Not .ExecSelect(strSQL) Then Err.Raise 9999, , "WK_製造依頼書建具 抽出エラー"
        
        inRecordCount = .GetRS.RecordCount
        intCnt = 1
        
        If Not .GetRS.EOF Then
            Do Until .GetRS.EOF
            
                If .GetRS![子扉] Then
                    intOyako = 2
                Else
                    intOyako = 1
                End If
                
                Set rsADO = objFullHeight.Rs建具種類(.GetRS![契約番号], .GetRS![棟番号], .GetRS![部屋番号], .GetRS![項], intOyako)
                
                If Not rsADO.EOF Then
                
                    Do Until rsADO.EOF
                        
                        intTateguShurui = rsADO![建具種類]
                        
                        strSQL = conSQL

                        strSQL = Replace(strSQL, "@通番@", intCnt)
                        intCnt = intCnt + 1
                        
                        strSQL = Replace(strSQL, "@契約番号@", varNullChk(.GetRS![契約番号], 1))
                        strSQL = Replace(strSQL, "@棟番号@", varNullChk(.GetRS![棟番号], 1))
                        strSQL = Replace(strSQL, "@部屋番号@", varNullChk(.GetRS![部屋番号], 1))
                        strSQL = Replace(strSQL, "@契約No@", varNullChk(.GetRS![契約No], 1))
                        strSQL = Replace(strSQL, "@物件名@", varNullChk(.GetRS![物件名], 1))
                        strSQL = Replace(strSQL, "@施工店@", varNullChk(.GetRS![施工店], 1))
                        strSQL = Replace(strSQL, "@項@", varNullChk(.GetRS![項], 1))
                        strSQL = Replace(strSQL, "@扉位置@", varNullChk(intTateguShurui, 1))
                        strSQL = Replace(strSQL, "@設置場所@", varNullChk(.GetRS![設置場所], 1))
                        strSQL = Replace(strSQL, "@品番1@", varNullChk(.GetRS![品番1], 1))
                        strSQL = Replace(strSQL, "@商品名@", varNullChk(.GetRS![商品名], 1))
                        strSQL = Replace(strSQL, "@色@", varNullChk(.GetRS![色], 1))
                        strSQL = Replace(strSQL, "@色コード@", varNullChk(.GetRS![色コード], 1))
                        strSQL = Replace(strSQL, "@吊元@", varNullChk(.GetRS![吊元], 1))
                        strSQL = Replace(strSQL, "@施錠@", varNullChk(.GetRS![施錠], 1))
                        strSQL = Replace(strSQL, "@数量@", varNullChk(.GetRS![数量], 1))
                        strSQL = Replace(strSQL, "@枚数@", varNullChk(.GetRS![数量], 1))
                        strSQL = Replace(strSQL, "@受注枚数@", varNullChk(.GetRS![枚数], 1))
                        strSQL = Replace(strSQL, "@品番区分@", varNullChk(.GetRS![品番区分], 1))
                        strSQL = Replace(strSQL, "@製造区分@", varNullChk(.GetRS![製造区分], 1))
                        strSQL = Replace(strSQL, "@工場CD@", varNullChk(.GetRS![工場CD], 1))
                        strSQL = Replace(strSQL, "@追加@", varNullChk(.GetRS![追加], 1))
                        strSQL = Replace(strSQL, "@階@", varNullChk(.GetRS![階], 1))
                        strSQL = Replace(strSQL, "@旧ヒンジ@", varNullChk(.GetRS![旧ヒンジ], 1))
                        strSQL = Replace(strSQL, "@欠品FLG@", varNullChk(.GetRS![欠品FLG], 1))
                        strSQL = Replace(strSQL, "@種類@", varNullChk(.GetRS![種類], 1))
                        strSQL = Replace(strSQL, "@工数@", varNullChk(.GetRS![工数], 1))
                        strSQL = Replace(strSQL, "@旧仕様@", varNullChk(.GetRS![旧仕様], 1))
                        strSQL = Replace(strSQL, "@出荷方法@", varNullChk(.GetRS![出荷方法], 1))
                        strSQL = Replace(strSQL, "@設計備考@", varNullChk(.GetRS![設計備考], 1))
                        strSQL = Replace(strSQL, "@08カタログ@", varNullChk(.GetRS![08カタログ], 1))
                        strSQL = Replace(strSQL, "@標準ハンドル@", varNullChk(.GetRS![標準ハンドル], 1))
                        strSQL = Replace(strSQL, "@DW@", varNullChk(.GetRS![DW], 1))
                        strSQL = Replace(strSQL, "@DH@", varNullChk(.GetRS![DH], 1))
                        strSQL = Replace(strSQL, "@CH@", varNullChk(.GetRS![CH], 1))
                        strSQL = Replace(strSQL, "@明り窓@", varNullChk(.GetRS![明り窓], 1))
                        strSQL = Replace(strSQL, "@個別Spec@", varNullChk(.GetRS![個別Spec], 1))
                        strSQL = Replace(strSQL, "@Spec@", varNullChk(.GetRS![Spec], 1))
                        strSQL = Replace(strSQL, "@Style@", varNullChk(.GetRS![Style], 1))
                        strSQL = Replace(strSQL, "@Style@", varNullChk(objTateguHinban.Style(.GetRS![品番1]), 1))
                        strSQL = Replace(strSQL, "@開閉様式@", varNullChk(.GetRS![開閉様式], 1))
                        strSQL = Replace(strSQL, "@新面材割付@", varNullChk(.GetRS![新面材割付], 1))
                        strSQL = Replace(strSQL, "@注意FLG@", varNullChk(.GetRS![注意FLG], 1))
                        strSQL = Replace(strSQL, "@物流施工日@", varNullChk(.GetRS![物流施工日], 1))
                        strSQL = Replace(strSQL, "@建具出荷方法@", varNullChk(.GetRS![建具出荷方法], 1))
                        strSQL = Replace(strSQL, "@クレーム用備考@", varNullChk(.GetRS![クレーム用備考], 1))
                        strSQL = Replace(strSQL, "@建具確定日@", varNullChk(.GetRS![建具確定日], 1))
                        strSQL = Replace(strSQL, "@製造日@", varNullChk(.GetRS![製造日], 1))
                        
                        If Not objLocalDB.ExecSQL(strSQL, strErrMsg) Then
                            Err.Raise 9999, , strErrMsg
                        End If
                
                        inReadCount = inReadCount + 1

                        
                        SysCmd acSysCmdSetStatus, "実行中.... " & inReadCount & "/" & inRecordCount
                        If inReadCount Mod 10 = 0 Then
                            DoEvents
                        End If
                    
                        rsADO.MoveNext
                        
                    Loop
                    
                    If rsADO.State = adStateOpen Then
                        rsADO.Close
                    End If
                        
                Else
                    Err.Raise 9999, , "フルハイトライン用の加工データ（設計）がありません。契約番号 = " & .GetRS![契約No] & " ,項No." & .GetRS![項]
                End If
            
                .GetRS.MoveNext
            Loop
        Else
            Err.Raise 9999, , "製造指示データがありません"
        End If
        
    End With
    
    GoTo Exit_bolfnc製造指示_フルハイト
    
Err_bolfnc製造指示_フルハイト:
    MsgBox Err.Description
    
Exit_bolfnc製造指示_フルハイト:
    
    Set objREMOTEDB = Nothing
    Set objLocalDB = Nothing
    Set objTateguHinban = Nothing
    Set objFullHeight = Nothing
    Set rsADO = Nothing
    
    SysCmd acSysCmdSetStatus, " "
    
End Function

Public Function bolfnc製造指示フルハイト帳票データ() As Boolean

Dim objLocalDB As New cls_LOCALDB
    Dim objTateguKansu As New cls_建具製造関数
    Dim objCheckLabel As New cls_建具識別ラベル
    
    Dim strSQL As String
    Dim strErrMsg As String
    
    Dim inRecordCount As Long
    Dim inReadCount As Long
    
    Dim strBrandHinban As String
    Dim strOEMHinban As String
    
    Dim intHanWari As Integer
    Dim intTateguStyle As Integer
    Dim intTateguKeijo As Integer
    Dim intOyako As Integer
    Dim strTsurimoto As String
    Dim strLineKbn As String
    
    Dim strQRLabel As String
    
    On Error GoTo Err_bolfnc製造指示フルハイト帳票データ
    
    SysCmd acSysCmdSetStatus, " "
    
    strSQL = "select * from WK_製造依頼書建具_FullHeight "
    strSQL = strSQL & "order by 通番 "
    
    With objLocalDB
        If Not .ExecSQL("delete * from WK_建具製造リスト_FullHeight") Then Err.Raise 9999, , "建具製造リスト_ローカル削除エラー"
        
        .CursorLocation = adUseClient
        
        If Not .ExecSelect(strSQL) Then Err.Raise 9999, , "Inputファイル抽出エラー"
        
        inRecordCount = .GetRS.RecordCount
        
        If Not .GetRS.EOF Then
            SysCmd acSysCmdClearStatus
            Do Until .GetRS.EOF
                
                If objTateguKansu.Bind(.GetRS![契約No], .GetRS![品番1], .GetRS![数量], .GetRS![受注枚数], .GetRS![DW], Nz(.GetRS![DH], 0), .GetRS![CH], .GetRS![施錠], Nz(.GetRS![明り窓], "-"), .GetRS![個別Spec]) Then
                    
                    If Nz(objTateguKansu.W1, 0) > 0 And Nz(objTateguKansu.W2, 0) Then
                        intHanWari = 1
                        
                        If Replace(objTateguKansu.ブランド品番, "特 ", "") Like "??B*" Then
                        
                            intTateguKeijo = 2
                        Else
                            intTateguKeijo = 1
                        End If
                        
                    Else
                        intHanWari = 0
                        intTateguKeijo = 0
                    End If
                    
                    If IsHirakido(objTateguKansu.ブランド品番) Or IsOyatobira(objTateguKansu.ブランド品番) Or IsKotobira(objTateguKansu.ブランド品番) Then
                        intTateguStyle = 0
                    ElseIf IsHikido(objTateguKansu.ブランド品番) Then
                        intTateguStyle = 1
                    ElseIf IsCloset_Hikichigai(objTateguKansu.ブランド品番) Then
                        intTateguStyle = 1
                    ElseIf IsCloset_Slide(objTateguKansu.ブランド品番) Then
                        intTateguStyle = 3
                    Else
                        Err.Raise 9999, , "開閉様式エラー 品番 = " & objTateguKansu.ブランド品番
                    End If
                    
                    If IsKotobira(objTateguKansu.ブランド品番) Then
                        intOyako = 2
                        
                        If Nz(.GetRS![吊元], "") = "R" Then
                            strTsurimoto = "L"
                        ElseIf Nz(.GetRS![吊元], "") = "L" Then
                            strTsurimoto = "R"
                        Else
                           Err.Raise 9999, , "子扉吊元エラー 契約No = " & .GetRS![契約No] & " 項 = " & .GetRS![項]
                        End If
                    Else
                        intOyako = 1
                        strTsurimoto = Nz(.GetRS![吊元], "Z")
                    End If
                    
                    strQRLabel = ""
                    strQRLabel = strQRLabel & RPAD(StrConv(.GetRS![契約No], vbNarrow), " ", 20)
                    strQRLabel = strQRLabel & LPAD(.GetRS![項], "0", 3)
                    strQRLabel = strQRLabel & intOyako
                    If Nz(.GetRS![扉位置], 0) > 3 Then
                        strQRLabel = strQRLabel & "03"
                    Else
                        strQRLabel = strQRLabel & LPAD(.GetRS![扉位置], "0", 2)
                    End If
                    strQRLabel = strQRLabel & LPAD(CStr(CInt(Nz(objTateguKansu.扉厚, 0) * 10)), "0", 3)
                    strQRLabel = strQRLabel & CStr(intHanWari)
                    
                    If intHanWari = 1 Then
                        strQRLabel = strQRLabel & "00000"
                    Else
                        strQRLabel = strQRLabel & LPAD(CStr(CInt(Nz(objTateguKansu.テノナW, 0) * 10)), "0", 5)
                    End If
                    
                    If intHanWari = 0 Then
                        strQRLabel = strQRLabel & "0000"
                    Else
                        strQRLabel = strQRLabel & LPAD(CStr(CInt(Nz(objTateguKansu.W1, 0) * 10)), "0", 4)
                    End If
                    
                    If intHanWari = 0 Then
                        strQRLabel = strQRLabel & "00000"
                    Else
                        strQRLabel = strQRLabel & LPAD(CStr(CInt(Nz(objTateguKansu.W2, 0) * 10)), "0", 5)
                    End If
                    
                    strQRLabel = strQRLabel & LPAD(CStr(CInt(Nz(objTateguKansu.テノナH, 0) * 10)), "0", 5)
                    
                    strQRLabel = strQRLabel & LPAD(CStr(CInt(Nz(objTateguKansu.芯厚, 0) * 10)), "0", 3)
                    
                    strQRLabel = strQRLabel & LPAD(CStr(CInt(Nz(objTateguKansu.表面材厚み, 0) * 10)), "0", 2)
                    
                    strQRLabel = strQRLabel & LPAD(.GetRS![通番], "0", 4)
                    
                    strOEMHinban = objTateguKansu.OEM品番

                    If IsNull(.GetRS![製造区分]) Then
                        strLineKbn = "S"
                        
                    ElseIf .GetRS![製造区分] = 2 Or .GetRS![製造区分] = 3 Then
                        strLineKbn = "K"
                    
                    ElseIf IsGikan(objTateguKansu.ブランド品番) Then
                        '物入れ引き違い（ガラス）--両面の場合は特注ライン、片面の場合はフラッシュライン
                        If IsCloset_Hikichigai(objTateguKansu.ブランド品番) Then
                        
                            If fncIntHalfGlassMirror_Maisu(objTateguKansu.ブランド品番, 2) = 2 Then
                                strLineKbn = "T"
                            Else
                                strLineKbn = "F"
                            End If
                        Else
                            strLineKbn = "T"
                        End If
                    Else
                        strLineKbn = "F"
                    End If
                        
                    strSQL = ""
                    strSQL = strSQL & "insert into WK_建具製造リスト_FullHeight( "
                    strSQL = strSQL & " 通番,契約No "
                    strSQL = strSQL & ",契約番号,棟番号,部屋番号 "
                    strSQL = strSQL & ", 物件名, 設置場所, 施工店, 項, ブランド品番"
                    strSQL = strSQL & ", OEM品番, 施錠, 色コード, 色, 数量"
                    strSQL = strSQL & ", 枚数, 受注枚数, 吊元, 扉位置, 開閉様式, 明り窓, DW, DH, CH"
                    strSQL = strSQL & ", 半割り, 建具スタイル, 建具形状, 親子"
                    strSQL = strSQL & ", 個別Spec, 基本図, 芯組詳細図, ガラス厚"
                    strSQL = strSQL & ", ガラス種類, ガラス枚数, ガラス種類2"
                    strSQL = strSQL & ", ガラス枚数2, テノナW, テノナW2"
                    strSQL = strSQL & ", テノナW中板, テノナW中板2, W1, W2"
                    strSQL = strSQL & ", ガラスW, テノナH, テノナH2, ガラスH"
                    strSQL = strSQL & ", テノナH中板, テノナH中板2, H1, H2, H3"
                    strSQL = strSQL & ", ガラス厚1, ガラスW1, ガラスH1, 芯厚"
                    strSQL = strSQL & ", 扉厚, 中板扉厚, 表面材厚み, 表面材厚み2"
                    strSQL = strSQL & ", CU_AS_AW表面材W, CU_AS_AW表面材W2, CU_AS_AW表面材H"
                    strSQL = strSQL & ", CU_AS_AW表面材H2, CU_AS_AW表面材枚数, CU_AS_AW表面材枚数2"
                    strSQL = strSQL & ", AS_AW中板W, AS_AW中板W1, AS_AW中板W2, AS_AW中板H, AS_AW中板H1"
                    strSQL = strSQL & ", AS_AW中板H2, AS_AW中板枚数, AS_AW中板枚数1, AS_AW中板枚数2"
                    strSQL = strSQL & ", 上パネルW, 上パネルH, 上パネル枚数, 上HFLG, 中パネルH"
                    strSQL = strSQL & ", 中パネル枚数, 中HFLG, 下パネルH, 下パネル枚数, 下HFLG"
                    strSQL = strSQL & ", ミラー表H, ミラー表枚数, ミラー裏H, ミラー裏枚数"
                    strSQL = strSQL & ", メタルW, メタル枚数, 上框種類, 上框長さ, 上框本数"
                    strSQL = strSQL & ", 下框種類, 下框長さ, 下框本数, アクセントライン "
                    strSQL = strSQL & ", ダボピッチ1, ダボピッチ2, ダボピッチ3, ダボピッチ4, ダボピッチ5 "
                    strSQL = strSQL & ", ダボピッチ6, ダボピッチ7, ダボピッチ8, ダボピッチ9, ダボピッチ10 "
                    strSQL = strSQL & ", 丁番位置上, 丁番位置中上, 丁番位置中下, 丁番位置下 "
                    strSQL = strSQL & ", 溝 "
                    strSQL = strSQL & ", ハンドル引手センター, ハンドル引手BS "
                    strSQL = strSQL & ", 鎌錠センター, 鎌錠BS "
                    strSQL = strSQL & ", 引込取手センター "
                    strSQL = strSQL & ", 子扉ラッチ受けセンター, 子扉鎌錠受けセンター "
                    strSQL = strSQL & ", 縁張外, 縁張内 "
                    strSQL = strSQL & ", カセット加工図, 戸開き図 "
                    strSQL = strSQL & ", 欠品FLG, QRコード, QRコード_2, Spec,設計備考 "
                    strSQL = strSQL & ", 新面材割付 "
                    strSQL = strSQL & ", 注意FLG "
                    strSQL = strSQL & ", 物流施工日  "
                    strSQL = strSQL & ", 建具出荷方法 "
                    strSQL = strSQL & ", クレーム用備考 "
                    strSQL = strSQL & ", 建具確定日 "
                    strSQL = strSQL & ", 08カタログ "
                    strSQL = strSQL & ", 製造区分 "
                    strSQL = strSQL & ", ライン区分 "
                    strSQL = strSQL & ", 製造日 "
                    strSQL = strSQL & " ) "
                    strSQL = strSQL & "values "
                    strSQL = strSQL & "( "
                    strSQL = strSQL & varNullChk(.GetRS![通番], 1) & "," & varNullChk(.GetRS![契約No], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![契約番号], 1) & "," & varNullChk(.GetRS![棟番号], 1) & "," & varNullChk(.GetRS![部屋番号], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![物件名], 1) & "," & varNullChk(.GetRS![設置場所], 1) & "," & varNullChk(.GetRS![施工店], 1) & "," & varNullChk(.GetRS![項], 1) & "," & varNullChk(objTateguKansu.ブランド品番, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.OEM品番, 1) & "," & varNullChk(.GetRS![施錠], 1) & "," & varNullChk(.GetRS![色コード], 1) & "," & varNullChk(.GetRS![色], 1) & "," & varNullChk(.GetRS![数量], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![枚数], 1) & "," & varNullChk(.GetRS![受注枚数], 1) & "," & varNullChk(strTsurimoto, 1) & "," & varNullChk(.GetRS![扉位置], 1) & "," & varNullChk(.GetRS![開閉様式], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![明り窓], 1) & "," & varNullChk(.GetRS![DW], 1) & "," & varNullChk(.GetRS![DH], 1) & "," & varNullChk(.GetRS![CH], 1)
                    strSQL = strSQL & "," & varNullChk(intHanWari, 1) & "," & varNullChk(intTateguStyle, 1) & "," & varNullChk(intTateguKeijo, 1) & "," & varNullChk(intOyako, 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![個別Spec], 1) & "," & varNullChk(objTateguKansu.基本図, 1) & "," & varNullChk(objTateguKansu.芯組詳細図, 1) & "," & varNullChk(objTateguKansu.ガラス厚, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.ガラス種類枚数("種類"), 1) & "," & varNullChk(objTateguKansu.ガラス種類枚数("枚数"), 1) & "," & varNullChk(objTateguKansu.ガラス種類枚数2("種類"), 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.ガラス種類枚数2("枚数"), 1) & "," & varNullChk(objTateguKansu.テノナW, 1) & "," & varNullChk(objTateguKansu.テノナW2, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.テノナW中板, 1) & "," & varNullChk(objTateguKansu.テノナW中板2, 1) & "," & varNullChk(objTateguKansu.W1, 1) & "," & varNullChk(objTateguKansu.W2, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.ガラスW, 1) & "," & varNullChk(objTateguKansu.テノナH, 1) & "," & varNullChk(objTateguKansu.テノナH2, 1) & "," & varNullChk(objTateguKansu.ガラスH, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.テノナH中板, 1) & "," & varNullChk(objTateguKansu.テノナH中板2, 1) & "," & varNullChk(objTateguKansu.H1, 1) & "," & varNullChk(objTateguKansu.H2, 1) & "," & varNullChk(objTateguKansu.H3, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.ガラス厚1, 1) & "," & varNullChk(objTateguKansu.ガラスW1, 1) & "," & varNullChk(objTateguKansu.ガラスH1, 1) & "," & varNullChk(objTateguKansu.芯厚, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.扉厚, 1) & "," & varNullChk(objTateguKansu.中板扉厚, 1) & "," & varNullChk(objTateguKansu.表面材厚み, 1) & "," & varNullChk(objTateguKansu.表面材厚み2, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.CU_AS_AW表面材W, 1) & "," & varNullChk(objTateguKansu.CU_AS_AW表面材W2, 1) & "," & varNullChk(objTateguKansu.CU_AS_AW表面材H, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.CU_AS_AW表面材H2, 1) & "," & varNullChk(objTateguKansu.CU_AS_AW表面材枚数, 1) & "," & varNullChk(objTateguKansu.CU_AS_AW表面材枚数2, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.AS_AW中板W, 1) & "," & varNullChk(objTateguKansu.AS_AW中板W1, 1) & "," & varNullChk(objTateguKansu.AS_AW中板W2, 1) & "," & varNullChk(objTateguKansu.AS_AW中板H, 1) & "," & varNullChk(objTateguKansu.AS_AW中板H1, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.AS_AW中板H2, 1) & "," & varNullChk(objTateguKansu.AS_AW中板枚数, 1) & "," & varNullChk(objTateguKansu.AS_AW中板枚数1, 1) & "," & varNullChk(objTateguKansu.AS_AW中板枚数2, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.上パネルW, 1) & "," & varNullChk(objTateguKansu.上パネルH, 1) & "," & varNullChk(objTateguKansu.上パネル枚数, 1) & "," & varNullChk(objTateguKansu.上HFLG, 1) & "," & varNullChk(objTateguKansu.中パネルH, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.中パネル枚数, 1) & "," & varNullChk(objTateguKansu.中HFLG, 1) & "," & varNullChk(objTateguKansu.下パネルH, 1) & "," & varNullChk(objTateguKansu.下パネル枚数, 1) & "," & varNullChk(objTateguKansu.下HFLG, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.ミラー表H, 1) & "," & varNullChk(objTateguKansu.ミラー表枚数, 1) & "," & varNullChk(objTateguKansu.ミラー裏H, 1) & "," & varNullChk(objTateguKansu.ミラー裏枚数, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.メタルW, 1) & "," & varNullChk(objTateguKansu.メタル枚数, 1) & "," & varNullChk(objTateguKansu.上框種類, 1) & "," & varNullChk(objTateguKansu.上框長さ, 1) & "," & varNullChk(objTateguKansu.上框本数, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.下框種類, 1) & "," & varNullChk(objTateguKansu.下框長さ, 1) & "," & varNullChk(objTateguKansu.下框本数, 1) & "," & varNullChk(objTateguKansu.アクセントライン, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.ダボピッチ1, 1) & "," & varNullChk(objTateguKansu.ダボピッチ2, 1) & "," & varNullChk(objTateguKansu.ダボピッチ3, 1) & "," & varNullChk(objTateguKansu.ダボピッチ4, 1) & "," & varNullChk(objTateguKansu.ダボピッチ5, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.ダボピッチ6, 1) & "," & varNullChk(objTateguKansu.ダボピッチ7, 1) & "," & varNullChk(objTateguKansu.ダボピッチ8, 1) & "," & varNullChk(objTateguKansu.ダボピッチ9, 1) & "," & varNullChk(objTateguKansu.ダボピッチ10, 1)
                    strSQL = strSQL & "," & varNullChk(objTateguKansu.丁番上位置, 1) & "," & varNullChk(objTateguKansu.丁番中上位置, 1) & "," & varNullChk(objTateguKansu.丁番中下位置, 1) & "," & varNullChk(objTateguKansu.丁番下位置, 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.溝(objTateguKansu.ブランド品番, .GetRS![施錠], .GetRS![個別Spec], .GetRS![扉位置]), 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.ハンドル引手センター(objTateguKansu.ブランド品番, .GetRS![施錠], .GetRS![個別Spec]), 1) & "," & varNullChk(objCheckLabel.ハンドル引手BS(objTateguKansu.ブランド品番, .GetRS![施錠], .GetRS![個別Spec]), 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.鎌錠センター(objTateguKansu.ブランド品番, .GetRS![施錠], .GetRS![個別Spec]), 1) & "," & varNullChk(objCheckLabel.鎌錠BS(objTateguKansu.ブランド品番, .GetRS![施錠], .GetRS![個別Spec]), 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.引込取手センター(objTateguKansu.ブランド品番, .GetRS![施錠], .GetRS![個別Spec]), 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.子扉ラッチ受けセンター(objTateguKansu.ブランド品番, .GetRS![施錠], .GetRS![個別Spec]), 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.子扉鎌錠受けセンター(objTateguKansu.ブランド品番, .GetRS![施錠], .GetRS![個別Spec]), 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.縁張_外(.GetRS![契約番号], objTateguKansu.ブランド品番, .GetRS![色コード]), 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.縁張_内(.GetRS![契約番号], objTateguKansu.ブランド品番, .GetRS![色コード]), 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.カセット加工図パス(objTateguKansu.ブランド品番, .GetRS![施錠], .GetRS![個別Spec]), 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.戸開き図パス(objTateguKansu.ブランド品番, .GetRS![開閉様式], .GetRS![吊元], .GetRS![扉位置]), 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![欠品FLG], 1) & "," & varNullChk(strQRLabel, 1) & "," & varNullChk(strQRLabel, 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![Spec], 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.建具設計備考(objTateguKansu.ブランド品番, .GetRS![個別Spec], .GetRS![設計備考]), 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![新面材割付], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![注意FLG], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![物流施工日], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![建具出荷方法], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![クレーム用備考], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![建具確定日], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![08カタログ], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![製造区分], 1)
                    strSQL = strSQL & "," & varNullChk(strLineKbn, 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![製造日], 1)
                    strSQL = strSQL & ") "
                    
                    'Debug.Print strSQL
                    If Not .ExecSQL(strSQL, strErrMsg) Then
                        Err.Raise 9999, , strErrMsg
                    End If
                    
                Else
                    If objTateguKansu.ブランド品番 <> "" Then
                        strBrandHinban = objTateguKansu.ブランド品番
                    Else
                        strBrandHinban = .GetRS![品番1]
                    End If
                    
                    Err.Raise 9998, , "建具関数引当エラー 品番=" & .GetRS![品番1]

                    strSQL = ""
                    strSQL = strSQL & "insert into WK_建具製造リスト_FullHeight( "
                    strSQL = strSQL & "  通番,契約No "
                    strSQL = strSQL & ",契約番号,棟番号,部屋番号 "
                    strSQL = strSQL & ", 物件名, 項, ブランド品番, OEM品番 "
                    strSQL = strSQL & ", 施錠, 色コード, 色, 数量"
                    strSQL = strSQL & ", 枚数, 受注枚数, 吊元, 扉位置, 明り窓,開閉様式"
                    strSQL = strSQL & ", DW, DH, CH"
                    strSQL = strSQL & ", 半割り, 建具スタイル, 建具形状, 親子"
                    strSQL = strSQL & ", 個別Spec "
                    strSQL = strSQL & ", 欠品FLG, QRコード, QRコード_2, Spec,設計備考 "
                    strSQL = strSQL & ", 新面材割付 "
                    strSQL = strSQL & ", 注意FLG "
                    strSQL = strSQL & ", 物流施工日  "
                    strSQL = strSQL & ", 建具出荷方法 "
                    strSQL = strSQL & ", クレーム用備考 "
                    strSQL = strSQL & ", 建具確定日 "
                    strSQL = strSQL & ", 製造区分 "
                    strSQL = strSQL & ", 製造日 "
                    strSQL = strSQL & " ) "
                    strSQL = strSQL & "values "
                    strSQL = strSQL & "( "
                    strSQL = strSQL & varNullChk(.GetRS![通番], 1) & "," & varNullChk(.GetRS![契約No], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![契約番号], 1) & "," & varNullChk(.GetRS![棟番号], 1) & "," & varNullChk(.GetRS![部屋番号], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![物件名], 1) & "," & varNullChk(.GetRS![項], 1) & "," & varNullChk(strBrandHinban, 1) & "," & varNullChk(strOEMHinban, 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![施錠], 1) & "," & varNullChk(.GetRS![色コード], 1) & "," & varNullChk(.GetRS![色], 1) & "," & varNullChk(.GetRS![数量], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![枚数], 1) & "," & varNullChk(.GetRS![受注枚数], 1) & "," & varNullChk(strTsurimoto, 1) & "," & varNullChk(.GetRS![扉位置], 1) & "," & varNullChk(.GetRS![明り窓], 1) & "," & varNullChk(.GetRS![開閉様式], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![DW], 1) & "," & varNullChk(.GetRS![DH], 1) & "," & varNullChk(.GetRS![CH], 1)
                    strSQL = strSQL & "," & varNullChk(intHanWari, 1) & "," & varNullChk(intTateguStyle, 1) & "," & varNullChk(intTateguKeijo, 1) & "," & varNullChk(intOyako, 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![個別Spec], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![欠品FLG], 1) & "," & varNullChk(.GetRS![QRコード], 1) & "," & varNullChk(.GetRS![QRコード_2], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![Spec], 1)
                    strSQL = strSQL & "," & varNullChk(objCheckLabel.建具設計備考(objTateguKansu.ブランド品番, .GetRS![個別Spec], .GetRS![設計備考]), 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![新面材割付], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![注意FLG], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![物流施工日], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![建具出荷方法], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![クレーム用備考], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![建具確定日], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![製造区分], 1)
                    strSQL = strSQL & "," & varNullChk(.GetRS![製造日], 1)
                    strSQL = strSQL & ") "
                    
                    'Debug.Print strSQL
                    If Not .ExecSQL(strSQL, strErrMsg) Then
                        Err.Raise 9999, , strErrMsg
                    End If
                    
                End If
                
                    inReadCount = inReadCount + 1

                    
                    SysCmd acSysCmdSetStatus, "実行中.... " & inReadCount & "/" & inRecordCount
                    If inReadCount Mod 10 = 0 Then
                        DoEvents
                    End If
                    .GetRS.MoveNext

            Loop
        End If
    
    End With
    
    If bolfnc製造指示フルハイト帳票データ_金物更新 Then
        If bolfnc製造指示フルハイト帳票データ_他ライン指示数取得 Then
            bolfnc製造指示フルハイト帳票データ = True
        End If
    End If
    

    GoTo Exit_bolfnc製造指示フルハイト帳票データ
    
Err_bolfnc製造指示フルハイト帳票データ:
    If Err.Number = 9998 Then
        Debug.Print Err.Description
        Resume Next
    Else
        MsgBox Err.Description
        
    End If
    'Resume
Exit_bolfnc製造指示フルハイト帳票データ:
    SysCmd acSysCmdSetStatus, " "
    Set objLocalDB = Nothing
    Set objTateguKansu = Nothing
End Function

Public Function bolfnc製造指示フルハイト帳票データ_金物更新() As Boolean
    Dim objLocalDB As New cls_LOCALDB
    Dim objREMOTEDB As New cls_BRAND_MASTER
    
    Dim strSQL As String
    Dim conSQL As String
    
    Dim strBuzaimei As String
    Dim dblBuzaisuGoukei As Double
    
    Dim i As Integer
    
    bolfnc製造指示フルハイト帳票データ_金物更新 = False
    
    On Error GoTo Err_bolfnc製造指示フルハイト帳票データ_金物更新
    
    strBuzaimei = ""
    dblBuzaisuGoukei = 0
    
    strSQL = "update WK_建具製造リスト_FullHeight "
    strSQL = strSQL & "set 取付金物1 = null "
    strSQL = strSQL & ", 取付金物1数量 = null "
    strSQL = strSQL & ", 取付金物2 = null "
    strSQL = strSQL & ", 取付金物2数量 = null "
    strSQL = strSQL & ", 取付金物3 = null "
    strSQL = strSQL & ", 取付金物3数量 = null "
    strSQL = strSQL & ", 取付金物4 = null "
    strSQL = strSQL & ", 取付金物4数量 = null "
    strSQL = strSQL & ", 取付金物5 = null "
    strSQL = strSQL & ", 取付金物5数量 = null "
    strSQL = strSQL & ", 取付金物6 = null "
    strSQL = strSQL & ", 取付金物6数量 = null "
    strSQL = strSQL & ", 取付金物7 = null "
    strSQL = strSQL & ", 取付金物7数量 = null "
    strSQL = strSQL & ", 取付金物8 = null "
    strSQL = strSQL & ", 取付金物8数量 = null "
    strSQL = strSQL & ", 取付金物9 = null "
    strSQL = strSQL & ", 取付金物9数量 = null "
    strSQL = strSQL & ", 取付金物10 = null "
    strSQL = strSQL & ", 取付金物10数量 = null "
    
    If Not objLocalDB.ExecSQL(strSQL) Then
        Err.Raise 9999, , "初期化エラー"
    End If
    
    conSQL = ""
    conSQL = conSQL & "select a.*,b.枚数 受注枚数 from BRAND_BOM.dbo.T_部材展開 a "
    conSQL = conSQL & "inner join T_受注明細 b "
    conSQL = conSQL & "on a.契約番号 = b.契約番号 and a.棟番号 = b.棟番号 and a.部屋番号 = b.部屋番号 and a.項 = b.項 "
    conSQL = conSQL & "where a.契約番号 = '@契約番号@' and a.棟番号 = '@棟番号@' and a.部屋番号 = '@部屋番号@' and a.項 = @項@ "
    conSQL = conSQL & "and a.品番区分 = 1 "
    conSQL = conSQL & "and a.取付 = '○' "
    
    strSQL = "select * from WK_建具製造リスト_FullHeight "
    
    With objLocalDB
        If .ExecSelect_Writable(strSQL) Then
            If Not .GetRS.EOF Then
                Do Until .GetRS.EOF
                    strSQL = conSQL
                    strSQL = Replace(strSQL, "@契約番号@", .GetRS![契約番号])
                    strSQL = Replace(strSQL, "@棟番号@", .GetRS![棟番号])
                    strSQL = Replace(strSQL, "@部屋番号@", .GetRS![部屋番号])
                    strSQL = Replace(strSQL, "@項@", .GetRS![項])

                    If objREMOTEDB.ExecSelect(strSQL) Then
                        If Not objREMOTEDB.GetRS.EOF Then
                            i = 1
                            Do Until objREMOTEDB.GetRS.EOF
                                If .GetRS![扉位置] = 0 Then
                                    strBuzaimei = objREMOTEDB.GetRS![部材名]
                                    dblBuzaisuGoukei = objREMOTEDB.GetRS![部材数合計] / objREMOTEDB.GetRS![受注枚数]
                                Else
                                    If (.GetRS![開閉様式] = "DF" And .GetRS![ブランド品番] Like "*-####*HY-*") Or (.GetRS![開閉様式] = "VF" And .GetRS![ブランド品番] Like "*-####*HF-*") Then
                                        If objREMOTEDB.GetRS![部材種別CD] Like "*ｴﾝﾄﾞｸｯｼｮﾝ*" Then
                                            If .GetRS![扉位置] <> 3 Then
                                                strBuzaimei = objREMOTEDB.GetRS![部材名]
                                                dblBuzaisuGoukei = objREMOTEDB.GetRS![部材数合計] / (objREMOTEDB.GetRS![受注枚数] - 1)
                                            End If
                                        ElseIf objREMOTEDB.GetRS![部材種別CD] Like "*戸車*" Then
                                            If .GetRS![扉位置] = 3 Then
                                                strBuzaimei = objREMOTEDB.GetRS![部材名]
                                                dblBuzaisuGoukei = objREMOTEDB.GetRS![部材数合計] / (objREMOTEDB.GetRS![受注枚数] - 2)
                                            End If
                                        ElseIf objREMOTEDB.GetRS![部材種別CD] Like "*ｸﾘｱﾊﾞﾝﾎﾞﾝ*" And IsTateguInset(.GetRS![ブランド品番]) Then
                                            If .GetRS![扉位置] <> 3 Then
                                                strBuzaimei = objREMOTEDB.GetRS![部材名]
                                                dblBuzaisuGoukei = objREMOTEDB.GetRS![部材数合計] / (objREMOTEDB.GetRS![受注枚数] - 1)
                                            End If
                                        Else
                                            strBuzaimei = objREMOTEDB.GetRS![部材名]
                                            dblBuzaisuGoukei = objREMOTEDB.GetRS![部材数合計] / objREMOTEDB.GetRS![受注枚数]
                                        End If
                                    Else
                                        strBuzaimei = objREMOTEDB.GetRS![部材名]
                                        dblBuzaisuGoukei = objREMOTEDB.GetRS![部材数合計] / objREMOTEDB.GetRS![受注枚数]
                                    End If
                                End If
                                
                                If strBuzaimei <> "" Then
                                    .GetRS.Update "取付金物" & i, strBuzaimei
                                    .GetRS.Update "取付金物" & i & "数量", dblBuzaisuGoukei
                                    i = i + 1
                                End If
                                
                                strBuzaimei = ""
                                dblBuzaisuGoukei = 0
                                
                                If i > 10 Then Exit Do
                                
                                objREMOTEDB.GetRS.MoveNext
                            Loop
                            objREMOTEDB.RecordSetClose
                        End If
                    Else
                        Err.Raise 9999, , "部材展開抽出実行エラー " & strSQL
                    End If
                    .GetRS.MoveNext
                Loop
            End If

        Else
            Err.Raise 9999, , "Input実行エラー " & strSQL
        End If
    End With
    
    bolfnc製造指示フルハイト帳票データ_金物更新 = True
    
    GoTo Exit_bolfnc製造指示フルハイト帳票データ_金物更新
   
Err_bolfnc製造指示フルハイト帳票データ_金物更新:
    MsgBox Err.Description
    'Resume
Exit_bolfnc製造指示フルハイト帳票データ_金物更新:
    Set objREMOTEDB = Nothing
    Set objLocalDB = Nothing
    
End Function

Private Function bolfnc製造指示フルハイト帳票データ_他ライン指示数取得() As Boolean
    Dim objLocalDB As New cls_LOCALDB
    Dim objREMOTEDB As New cls_BRAND_MASTER
    
    Dim strSQL As String
    Dim strSQLR As String
    Dim conSQL As String
    Dim consqlR As String
    
    On Error GoTo Err_bolfnc製造指示フルハイト帳票データ_他ライン指示数取得
    
    conSQL = ""
    conSQL = conSQL & "update WK_建具製造リスト_FullHeight "
    conSQL = conSQL & "set ラインALL数量 = @建具枚数ALL@ "
    conSQL = conSQL & ",ライン2数量 = @建具枚数2@ "
    conSQL = conSQL & ",折戸枚数 = @折戸枚数@ "
    conSQL = conSQL & "where 契約No = '@契約No@' "
    
    consqlR = ""
    consqlR = consqlR & "select sum(枚数) as 折戸枚数 from T_受注明細 "
    consqlR = consqlR & "where 契約番号 = '@契約番号@' "
    consqlR = consqlR & "and 棟番号 = '@棟番号@' "
    consqlR = consqlR & "and 部屋番号 = '@部屋番号@' "
    consqlR = consqlR & "and 種類 = 'ｸﾛｾﾞｯﾄ' "
    consqlR = consqlR & "and dbo.IsCloset_Isehara(dbo.fncgetHinban(品番1,特注建具品番)) = 0 "

    strSQL = ""
    strSQL = strSQL & "select ALLDATA.契約番号,ALLDATA.棟番号,ALLDATA.部屋番号,ALLDATA.契約No,建具枚数ALL,建具枚数2 from "
    strSQL = strSQL & "(select 契約番号,棟番号,部屋番号,契約No,sum(枚数) as 建具枚数ALL from WK_製造依頼書建具 group by 契約番号,棟番号,部屋番号,契約No) as ALLDATA "
    strSQL = strSQL & "left join "
    strSQL = strSQL & "(select 契約No,sum(枚数) as 建具枚数2 from WK_製造依頼書建具_FullHeight group by 契約No) as FillHeightLine "
    strSQL = strSQL & "on ALLDATA.契約No = FillHeightLine.契約No "
    
    With objLocalDB
        If .ExecSelect(strSQL) Then
            Do Until .GetRS.EOF
                                
                strSQLR = consqlR
                strSQLR = Replace(strSQLR, "@契約番号@", .GetRS![契約番号])
                strSQLR = Replace(strSQLR, "@棟番号@", .GetRS![棟番号])
                strSQLR = Replace(strSQLR, "@部屋番号@", .GetRS![部屋番号])
                
                strSQL = conSQL
                strSQL = Replace(strSQL, "@建具枚数ALL@", Nz(.GetRS![建具枚数ALL], 0))
                strSQL = Replace(strSQL, "@建具枚数2@", Nz(.GetRS![建具枚数2], 0))
                strSQL = Replace(strSQL, "@契約No@", .GetRS![契約No])
                
                With objREMOTEDB
                    If .ExecSelect(strSQLR) Then
                        If Not .GetRS.EOF Then
                            strSQL = Replace(strSQL, "@折戸枚数@", Nz(.GetRS![折戸枚数], 0))
                        Else
                            strSQL = Replace(strSQL, "@折戸枚数@", 0)
                        End If
                    Else
                        Err.Raise 9999, , "折戸枚数検索エラー SQL=" & strSQLR
                    End If
                End With
                'Debug.Print strSQL
                
                If Not .ExecSQL(strSQL) Then
                    Err.Raise 9999, , "SQL実行エラー SQL=" & strSQL
                End If
                
                .GetRS.MoveNext
            Loop
        Else
            Err.Raise 9999, , "総数集計実行エラー "
        End If
    
    End With
    
    bolfnc製造指示フルハイト帳票データ_他ライン指示数取得 = True
    
    GoTo Exit_bolfnc製造指示フルハイト帳票データ_他ライン指示数取得

Err_bolfnc製造指示フルハイト帳票データ_他ライン指示数取得:
    MsgBox Err.Description
Exit_bolfnc製造指示フルハイト帳票データ_他ライン指示数取得:
    Set objLocalDB = Nothing
    Set objREMOTEDB = Nothing
    
End Function