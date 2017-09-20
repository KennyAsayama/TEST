Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'2.0.0 ADD

Private objREMOTEDB As cls_BRAND_MASTER
Private rsADO As ADODB.Recordset

Private intSeiZoKanoSu(2) As Integer
Private conSQL_FullHeight As String

Private Sub Class_Initialize()
'--------------------------------------------------------------------------------------------------------------------
'コンストラクタ
'--------------------------------------------------------------------------------------------------------------------

    Set objREMOTEDB = New cls_BRAND_MASTER
    
    conSQL_FullHeight = conSQL_FullHeight & "from BRAND_NEWLINE.dbo.T_FullHeightLineデータ作成対象マスタ a "
    conSQL_FullHeight = conSQL_FullHeight & "inner join "
    conSQL_FullHeight = conSQL_FullHeight & "( "
    conSQL_FullHeight = conSQL_FullHeight & "select distinct x.*,y.Style, left(y.個別Spec,3) 仕様, right(y.個別Spec,4) 年月,s.確定 製造確定 "
    conSQL_FullHeight = conSQL_FullHeight & ",物件名,施工店,s.製造区分 "
    conSQL_FullHeight = conSQL_FullHeight & ",dbo.fncGetHinban(品番1,特注建具品番) 製造品番 "
    conSQL_FullHeight = conSQL_FullHeight & ",s.工場CD,s.確定日,s.製造日 ,s.数量 製造数 "
    conSQL_FullHeight = conSQL_FullHeight & ",[dbo].[fncGetKaiheiYoshiki](x.建具品番,1) 開閉様式 "
    conSQL_FullHeight = conSQL_FullHeight & "from BRAND_NEWLINE.dbo.T_新ﾗｲﾝ製造ﾃﾞｰﾀ x "
    conSQL_FullHeight = conSQL_FullHeight & "inner join T_受注明細 y "
    conSQL_FullHeight = conSQL_FullHeight & "on x.契約番号 = y.契約番号 and x.棟番号 = y.棟番号 and x.部屋番号 = y.部屋番号 and x.項 = y.項 "
    conSQL_FullHeight = conSQL_FullHeight & "inner join T_受注ﾏｽﾀ jm "
    conSQL_FullHeight = conSQL_FullHeight & "on jm.契約番号 = y.契約番号 and jm.棟番号 = y.棟番号 and jm.部屋番号 = y.部屋番号 "
    conSQL_FullHeight = conSQL_FullHeight & "inner join T_製造指示  s "
    conSQL_FullHeight = conSQL_FullHeight & "on x.契約番号 = s.契約番号 and x.棟番号 = s.棟番号 and x.部屋番号 = s.部屋番号 and x.項 = s.項 "
    conSQL_FullHeight = conSQL_FullHeight & "where s.製造区分 in (1,2,3) "
    conSQL_FullHeight = conSQL_FullHeight & "and s.製造日 = '@@SeizoDate@@' "
    conSQL_FullHeight = conSQL_FullHeight & "and x.製造ライン = '00' "
    conSQL_FullHeight = conSQL_FullHeight & ") b "
    conSQL_FullHeight = conSQL_FullHeight & "on a.シリーズ名 = b.Style "
    conSQL_FullHeight = conSQL_FullHeight & "and a.開閉様式 = b.開閉様式 "
    conSQL_FullHeight = conSQL_FullHeight & "and a.仕様 = b.仕様 and b.年月 between 開始 and 終了 "
    conSQL_FullHeight = conSQL_FullHeight & "and a.製造対象 = 1 "
    
End Sub

Private Sub Class_Terminate()
'--------------------------------------------------------------------------------------------------------------------
'ディストラクタ
'--------------------------------------------------------------------------------------------------------------------
        
    On Error Resume Next
    rsADO.Close
    Set rsADO = Nothing
    Set objREMOTEDB = Nothing
    
End Sub

Public Property Get 製造可能数(ByVal 確定区分 As Byte) As Variant
'--------------------------------------------------------------------------------------------------------------------
'フルハイトラインで製造可能な枚数(確定区分別）
'   ↓の製造可能数集計を先に実行しておく必要あり

'   確定区分はT_製造指示の確定区分
'       0:未確定
'       1:仮確定
'       2:確定

'--------------------------------------------------------------------------------------------------------------------
    On Error GoTo Err_製造可能数
    
    If intSeiZoKanoSu(確定区分) = 0 Then
        製造可能数 = Null
    Else
        製造可能数 = intSeiZoKanoSu(確定区分)
    End If
    
    Exit Property

Err_製造可能数:
    MsgBox Err.Description
    製造可能数 = Null
    
End Property

Public Function 製造可能数集計(ByVal 製造日 As Date) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'フルハイトラインで製造可能な枚数を取得
'--------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strDate As String
    
    strDate = Format(製造日, "yyyy-mm-dd")
    
    Erase intSeiZoKanoSu
    
    On Error GoTo Err_製造可能数集計
    
    strSQL = ""
    strSQL = strSQL & "select 製造確定,count(*) 製造可能枚数 "
    strSQL = strSQL & Replace(conSQL_FullHeight, "@@SeizoDate@@", strDate)
    strSQL = strSQL & "group by 製造確定 "
    

    With objREMOTEDB
        If .ExecSelect(strSQL) Then
            Do Until .GetRS.EOF
                intSeiZoKanoSu(.GetRS![製造確定]) = .GetRS![製造可能枚数]
                .GetRS.MoveNext
            Loop
        Else
            Err.Raise 9999, , "製造可能数集計 SQL実行エラー "
        End If
        
    End With
    
    製造可能数集計 = True
    
    Exit Function
    
Err_製造可能数集計:
    MsgBox Err.Description
    製造可能数集計 = False
    
End Function

Public Property Get 製造可能データ(ByVal 製造日 As Date) As ADODB.Recordset
'--------------------------------------------------------------------------------------------------------------------
'フルハイトラインで製造可能なデータを取得
'--------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strDate As String
    
    On Error GoTo Err_製造可能データ
    
    objREMOTEDB.RecordSetClose
    strDate = Format(製造日, "yyyy-mm-dd")
    
    strSQL = ""
    strSQL = strSQL & "select 契約番号,棟番号,部屋番号,物件名,施工店 "
    strSQL = strSQL & ",契約番号 + '-' + 棟番号 + '-' + 部屋番号 契約No "
    strSQL = strSQL & ",dbo.fncNohinAddress(契約番号,棟番号,部屋番号,項,1) 納品住所 "
    strSQL = strSQL & ",項,製造区分,製造品番,工場CD,確定日,製造日 "
    strSQL = strSQL & ",dbo.fncSeizoSyukkaDate(契約番号,棟番号,部屋番号,項,1) 出荷日 "
    strSQL = strSQL & ",case 製造区分 when 1 then 製造数 else 0 end Flush数 "
    strSQL = strSQL & ",case 製造区分 when 2 then 製造数 else 0 end F框数 "
    strSQL = strSQL & ",case 製造区分 when 3 then 製造数 else 0 end 框数 "
    strSQL = strSQL & ",style,b.開閉様式 "
    strSQL = strSQL & ",dbo.fncGetColor(製造品番) 面材色 "
    strSQL = strSQL & ",case 製造確定 when 1 then '仮' when 2 then '確' else '未' end 確定 "
    strSQL = strSQL & Replace(conSQL_FullHeight, "@@SeizoDate@@", strDate)
    
    With objREMOTEDB
        If .ExecSelect(strSQL) Then
            Set 製造可能データ = .GetRS
        Else
            Debug.Print strSQL
            Err.Raise 9999, , "製造可能数 SQL実行エラー "
        End If
    End With
    
    GoTo Exit_製造可能データ
    
Err_製造可能データ:
    MsgBox Err.Description
   
Exit_製造可能データ:

End Property

Public Function フルハイトライン製造数(ByVal 製造日 As Date) As Integer
'--------------------------------------------------------------------------------------------------------------------
'フルハイトライン指示数を取得(仮確定以上)
'--------------------------------------------------------------------------------------------------------------------

    Dim strSQL As String
    Dim strDate As String
    
    On Error GoTo Err_フルハイトライン製造数
    
    objREMOTEDB.RecordSetClose
    strDate = Format(製造日, "yyyy-mm-dd")
    
    strSQL = ""
    strSQL = strSQL & "select sum(数量) 枚数 "
    strSQL = strSQL & "from T_製造指示 "
    strSQL = strSQL & "where 製造日 = '" & strDate & "' "
    strSQL = strSQL & "and 製造区分 in (1,2,3) "
    strSQL = strSQL & "and 確定 > 0 "
    strSQL = strSQL & "and 工場CD = 10 "
    
    With objREMOTEDB
        If .ExecSelect(strSQL) Then
            フルハイトライン製造数 = Nz(.GetRS![枚数], 0)
        Else
            Err.Raise 9999, , "フルハイトライン製造数 SQL実行エラー "
        End If
    End With
    
    GoTo Exit_フルハイトライン製造数
    
Err_フルハイトライン製造数:
    MsgBox Err.Description
   
Exit_フルハイトライン製造数:
    objREMOTEDB.RecordSetClose
    
End Function

Public Function fnc契約番号別製造可能数集計(ByVal in区分 As Byte, ByVal 日付 As Variant) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'フルハイトラインで製造可能な枚数を取得（製造日又は確定日、邸別）
'
'      in区分 : 1→確定日 2→製造日
'--------------------------------------------------------------------------------------------------------------------
    Dim objLOCALDB As New cls_LOCALDB
    Dim strSQL As String
    Dim strDate As String
    
    'Nullの場合はワークファイルを消してそのまま終了
    If IsNull(日付) Then
        objLOCALDB.ExecSQL ("delete * from WK_フルハイトライン製造数邸別")
        fnc契約番号別製造可能数集計 = True
        Set objLOCALDB = Nothing
        Exit Function
    End If
    
    strDate = Format(日付, "yyyy-mm-dd")
    fnc契約番号別製造可能数集計 = False
    
    On Error GoTo Err_fnc契約番号別製造可能数集計
    
    strSQL = ""
    strSQL = strSQL & "select 契約番号,棟番号,部屋番号,count(*) 製造可能枚数,sum(IIf(工場CD=10,1,0)) 指示済枚数 "
    
    If in区分 = 1 Then '確定日
        strSQL = strSQL & Replace(conSQL_FullHeight, "s.製造日 = '@@SeizoDate@@'", "s.確定日 = '@@SeizoDate@@'")
    Else
        strSQL = strSQL & conSQL_FullHeight
    End If
    
    strSQL = Replace(strSQL, "@@SeizoDate@@", strDate)
    
'    strSQL = strSQL & "and 契約番号 = '" & 契約番号 & "' and 棟番号 = '" & 棟番号 & "' and 部屋番号 = '" & 部屋番号 & "' "
    strSQL = strSQL & "group by 契約番号,棟番号,部屋番号 "
    
'    Debug.Print strSQL
    
    If objLOCALDB.ExecSQL("delete * from WK_フルハイトライン製造数邸別") Then
    
    
        With objREMOTEDB
            If .ExecSelect(strSQL) Then
                If Not .GetRS.EOF Then
                    Do Until .GetRS.EOF
                        strSQL = ""
                        strSQL = strSQL & "insert into WK_フルハイトライン製造数邸別( "
                        strSQL = strSQL & "契約番号,棟番号,部屋番号,フルハイトライン製造可能数,フルハイトライン指示済数 "
                        strSQL = strSQL & ") values ( "
                        strSQL = strSQL & varNullChk(.GetRS![契約番号], 1) & " "
                        strSQL = strSQL & "," & varNullChk(.GetRS![棟番号], 1) & " "
                        strSQL = strSQL & "," & varNullChk(.GetRS![部屋番号], 1) & " "
                        strSQL = strSQL & "," & varNullChk(.GetRS![製造可能枚数], 1) & " "
                        strSQL = strSQL & "," & varNullChk(.GetRS![指示済枚数], 1) & " "
                        strSQL = strSQL & ") "
                        If Not objLOCALDB.ExecSQL(strSQL) Then
                           Err.Raise 9999, , "製造可能数集計 ローカルファイル追加エラー "
                        End If
                        
                        .GetRS.MoveNext
                    Loop
                End If
            Else
                Err.Raise 9999, , "製造可能数集計 SQL実行エラー "
            End If
            
        End With
    
    Else
        Err.Raise 9999, , "ワークファイル削除エラー(WK_フルハイトライン邸別) "
    End If
    
    fnc契約番号別製造可能数集計 = True
    
    GoTo Exit_fnc契約番号別製造可能数集計
    
Err_fnc契約番号別製造可能数集計:
    Debug.Print Err.Description
    fnc契約番号別製造可能数集計 = False
Exit_fnc契約番号別製造可能数集計:
    Set objLOCALDB = Nothing
End Function

Public Function fnc契約番号別製造可能数集計_期間(ByVal in区分 As Byte, ByVal from日付 As Variant, ByVal to日付 As Variant) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'フルハイトラインで製造可能な枚数を取得（製造日又は確定日、From~Toの期間内で取得、邸別）
'
'      in区分 : 1→確定日 2→製造日
'--------------------------------------------------------------------------------------------------------------------
    Dim objLOCALDB As New cls_LOCALDB
    Dim strSQL As String
    Dim strDateF As String
    Dim strDateT As String
    
    'Nullの場合はワークファイルを消してそのまま終了
    If IsNull(from日付) Or IsNull(to日付) Then
        objLOCALDB.ExecSQL ("delete * from WK_フルハイトライン製造数邸別")
        fnc契約番号別製造可能数集計_期間 = True
        Set objLOCALDB = Nothing
        Exit Function
    End If
    
    strDateF = Format(from日付, "yyyy-mm-dd")
    strDateT = Format(to日付, "yyyy-mm-dd")
    
    fnc契約番号別製造可能数集計_期間 = False
    
    On Error GoTo Err_fnc契約番号別製造可能数集計_期間
    
    strSQL = ""
    strSQL = strSQL & "select 契約番号,棟番号,部屋番号,count(*) 製造可能枚数,sum(IIf(工場CD=10,1,0)) 指示済枚数 "
    
    If in区分 = 1 Then '確定日
        strSQL = strSQL & Replace(conSQL_FullHeight, "s.製造日 = '@@SeizoDate@@'", "s.確定日 between '" & strDateF & "' and '" & strDateT & "'")
    Else
        strSQL = strSQL & Replace(conSQL_FullHeight, "s.製造日 = '@@SeizoDate@@'", "s.製造日 between '" & strDateF & "' and '" & strDateT & "'")
    End If
    
   strSQL = strSQL & "group by 契約番号,棟番号,部屋番号 "
    
'    Debug.Print strSQL
    
    If objLOCALDB.ExecSQL("delete * from WK_フルハイトライン製造数邸別") Then
    
    
        With objREMOTEDB
            If .ExecSelect(strSQL) Then
                If Not .GetRS.EOF Then
                    Do Until .GetRS.EOF
                        strSQL = ""
                        strSQL = strSQL & "insert into WK_フルハイトライン製造数邸別( "
                        strSQL = strSQL & "契約番号,棟番号,部屋番号,フルハイトライン製造可能数,フルハイトライン指示済数 "
                        strSQL = strSQL & ") values ( "
                        strSQL = strSQL & varNullChk(.GetRS![契約番号], 1) & " "
                        strSQL = strSQL & "," & varNullChk(.GetRS![棟番号], 1) & " "
                        strSQL = strSQL & "," & varNullChk(.GetRS![部屋番号], 1) & " "
                        strSQL = strSQL & "," & varNullChk(.GetRS![製造可能枚数], 1) & " "
                        strSQL = strSQL & "," & varNullChk(.GetRS![指示済枚数], 1) & " "
                        strSQL = strSQL & ") "
                        If Not objLOCALDB.ExecSQL(strSQL) Then
                           Err.Raise 9999, , "製造可能数集計 ローカルファイル追加エラー "
                        End If
                        
                        .GetRS.MoveNext
                    Loop
                End If
            Else
                Err.Raise 9999, , "製造可能数集計 SQL実行エラー "
            End If
            
        End With
    
    Else
        Err.Raise 9999, , "ワークファイル削除エラー(WK_フルハイトライン邸別) "
    End If
    
    fnc契約番号別製造可能数集計_期間 = True
    
    GoTo Exit_fnc契約番号別製造可能数集計_期間
    
Err_fnc契約番号別製造可能数集計_期間:
    Debug.Print Err.Description
    fnc契約番号別製造可能数集計_期間 = False
Exit_fnc契約番号別製造可能数集計_期間:
    Set objLOCALDB = Nothing
End Function