Option Compare Database
Option Explicit
                
Public Function bolfncUwawakuShitaji_GroupADD() As Boolean
'   *************************************************************
'   上枠加工有無確認
'   20170301 K.Asayama ADD
'
'   上枠下地カット表に集計グループと厚さを追加する

'   1.12.1
'       →バグ修正(SQL文の like条件に *[アスタリスク]を使用していたため %[パーセント]に変更
'       →図面ありの時に上枠下地がある品番か確認する
'   1.12.2
'        →Err.Raise時に引数の数が間違っているところを修正
'   *************************************************************
    
    Dim cnADO As ADODB.Connection
    Dim rsADO As ADODB.Recordset
    Dim strSQL As String
    Dim intGroupName As Integer
    Dim dblcutLength() As Double
    Dim dblGroupSu As Double
    Dim i As Integer
    Dim bolStealth As Boolean
    Dim varHinban As Variant
    Dim bolToku As Boolean
    
    Set cnADO = CurrentProject.Connection
    Set rsADO = New ADODB.Recordset
    
    bolfncUwawakuShitaji_GroupADD = False
    
    On Error GoTo Err_fncbolUwawakuShitaji_GroupADD
    
'    DoCmd.RunSQL "delete from WK_ｽﾃﾙｽ上枠下地集計表"
'    DoCmd.RunSQL "delete from WK_ｲﾝｾｯﾄ上枠下地集計表"
    DoCmd.RunSQL "delete from WK_上枠下地グループ集計"
    
    strSQL = ""
    strSQL = strSQL & "select S.*,(数量 - 在庫引当数) AS 加工数 from TMP_製造指示書下地材 S "
    strSQL = strSQL & "inner join TMP_当日製造 T "
    strSQL = strSQL & "on S.契約番号 = T.契約番号 and S.棟番号 = T.棟番号 and S.部屋番号 = T.部屋番号 and S.項 = T.項 "
    strSQL = strSQL & "where (数量 - 在庫引当数) > 0 and 下地材欠品FLG = False "
    strSQL = strSQL & "and (下地枠設計備考 like '%図面あり%' or 上枠下地W <> 0) "
    
    rsADO.Open strSQL, cnADO, adOpenStatic, adLockPessimistic
    
    Do Until rsADO.EOF
        strSQL = ""
        dblGroupSu = 0
        bolStealth = False
        
        '品番取得
        bolToku = bolFncTokuHinban(rsADO![下地材品番], rsADO![特注下地材品番], varHinban)
        
        'ステルス確認
        If IsStealth_Seizo_TEMP(varHinban) Then bolStealth = True
        
        If Not Nz(rsADO![下地枠設計備考], "") Like "*図面あり*" Then
            intGroupName = intFncUwawakuShitajiLengthGroup(rsADO![上枠下地W])
            
            If bolfncUwawakuShitajiLength(rsADO![上枠下地W], dblcutLength()) Then
                For i = 1 To 3
                    If dblcutLength(i) > 0 Then
                        strSQL = ""
                        strSQL = strSQL & "insert into WK_上枠下地グループ集計(  "
                        strSQL = strSQL & "契約番号,棟番号,部屋番号,邸名,項,下地材品番,開き戸,クローゼット,ステルス"
                        strSQL = strSQL & ",商品名 "
                        strSQL = strSQL & ",上枠下地W,分割後上枠下地W,上枠下地W長さグループ,上総巾,上固定値巾,上変動値巾,厚さ,数量,カット後数量,備考,入力順 "
                        strSQL = strSQL & ") values ( "
                        strSQL = strSQL & varNullChk(rsADO![契約番号], 1)
                        strSQL = strSQL & "," & varNullChk(rsADO![棟番号], 1) & " "
                        strSQL = strSQL & "," & varNullChk(rsADO![部屋番号], 1) & " "
                        strSQL = strSQL & "," & varNullChk(rsADO![物件名], 1) & " "
                        strSQL = strSQL & "," & varNullChk(rsADO![項], 1) & " "
                        If bolToku Then
                            strSQL = strSQL & "," & varNullChk(rsADO![特注下地材品番], 1) & " "
                        Else
                        
                            strSQL = strSQL & "," & varNullChk(rsADO![下地材品番], 1) & " "
                        End If
                        If IsHirakido(varHinban) _
                            Or IsOyatobira(varHinban) _
                                Or IsCloset_Hiraki(varHinban) Then
                            strSQL = strSQL & "," & True & " "
                        Else
                            strSQL = strSQL & "," & False & " "
                        End If
                        
                        If IsCloset_Hiraki(varHinban) _
                            Or IsCloset_Oredo(varHinban) Then
                                strSQL = strSQL & "," & True & " "
                        Else
                                strSQL = strSQL & "," & False & " "
                        End If
                        strSQL = strSQL & "," & bolStealth & " "
                        strSQL = strSQL & "," & varNullChk(rsADO![商品名], 1) & " "
                        strSQL = strSQL & "," & varNullChk(rsADO![上枠下地W], 1) & " "
                        strSQL = strSQL & "," & varNullChk(dblcutLength(i), 1) & " "
                        strSQL = strSQL & "," & IIf(intFncUwawakuShitajiLengthGroup(dblcutLength(i)) > 0, CStr(intFncUwawakuShitajiLengthGroup(dblcutLength(i))), Null)
                        strSQL = strSQL & "," & varNullChk(rsADO![上枠下地総巾], 1) & " "
                        strSQL = strSQL & "," & varNullChk(rsADO![上枠下地固定値側], 1) & " "
                        strSQL = strSQL & "," & varNullChk(rsADO![上枠下地変動値側], 1) & " "
                        strSQL = strSQL & "," & varNullChk(fncstrUwawakuShitajiT(varHinban, rsADO![下がり壁], rsADO![ボード厚]), 1) & " "

                        If i = 1 Then
                            strSQL = strSQL & "," & varNullChk(rsADO![加工数], 1) & " "
                        Else
                            strSQL = strSQL & ",0 "
                        End If
                        
                        If dblcutLength(i) > 300 And dblcutLength(i) <= 900 Then
                            strSQL = strSQL & "," & varNullChk(rsADO![加工数] * 0.5, 1) & " "
                        Else
                            strSQL = strSQL & "," & varNullChk(rsADO![加工数], 1) & " "
                        End If
                        strSQL = strSQL & ",Null "
                        strSQL = strSQL & "," & varNullChk(rsADO![入力順], 1) & " "
                        strSQL = strSQL & ") "
                        
                        cnADO.Execute strSQL
                        
                    Else
                        Exit For
                    End If
                Next
            Else
                Err.Raise 9999, , "以降の帳票の出力を中止します" '1.12.2
            End If
        Else '図面あり
            If fncstrUwawakuShitajiT(varHinban, rsADO![下がり壁], rsADO![ボード厚]) <> "" Then
                strSQL = ""
                strSQL = strSQL & "insert into WK_上枠下地グループ集計(  "
                strSQL = strSQL & "契約番号,棟番号,部屋番号,邸名,項,下地材品番,開き戸,クローゼット,ステルス"
                strSQL = strSQL & ",商品名 "
                strSQL = strSQL & ",上枠下地W,分割後上枠下地W,上枠下地W長さグループ,上総巾,上固定値巾,上変動値巾,厚さ,数量,カット後数量,備考,入力順 "
                strSQL = strSQL & ") values ( "
                strSQL = strSQL & varNullChk(rsADO![契約番号], 1)
                strSQL = strSQL & "," & varNullChk(rsADO![棟番号], 1) & " "
                strSQL = strSQL & "," & varNullChk(rsADO![部屋番号], 1) & " "
                strSQL = strSQL & "," & varNullChk(rsADO![物件名], 1) & " "
                strSQL = strSQL & "," & varNullChk(rsADO![項], 1) & " "
                If bolToku Then
                    strSQL = strSQL & "," & varNullChk(rsADO![特注下地材品番], 1) & " "
                Else
                
                    strSQL = strSQL & "," & varNullChk(rsADO![下地材品番], 1) & " "
                End If
                If IsHirakido(varHinban) _
                    Or IsOyatobira(varHinban) _
                        Or IsCloset_Hiraki(varHinban) Then
                    strSQL = strSQL & "," & True & " "
                Else
                    strSQL = strSQL & "," & False & " "
                End If
                
                If IsCloset_Hiraki(varHinban) _
                    Or IsCloset_Oredo(varHinban) Then
                        strSQL = strSQL & "," & True & " "
                Else
                        strSQL = strSQL & "," & False & " "
                End If
                strSQL = strSQL & "," & bolStealth & " "
                strSQL = strSQL & "," & varNullChk(rsADO![商品名], 1) & " "
                strSQL = strSQL & "," & "0,0,Null,0,0,0,Null "
                strSQL = strSQL & "," & varNullChk(rsADO![加工数], 1) & " "
                strSQL = strSQL & "," & varNullChk(rsADO![加工数], 1) & " "
                strSQL = strSQL & ",'" & rsADO![物件名] & "　" & rsADO![項] & "　" & rsADO![下地枠設計備考] & "' "
                strSQL = strSQL & "," & varNullChk(rsADO![入力順], 1) & " "
                strSQL = strSQL & ") "
                
                cnADO.Execute strSQL
            End If
            
        End If
        
        rsADO.MoveNext
    Loop
    
    bolfncUwawakuShitaji_GroupADD = True
    
    GoTo Exit_fncbolUwawakuShitaji_GroupADD
    
Err_fncbolUwawakuShitaji_GroupADD:
    MsgBox Err.Description
    Debug.Print strSQL
    
Exit_fncbolUwawakuShitaji_GroupADD:
    If rsADO.State = adStateOpen Then rsADO.Close
    Set rsADO = Nothing
    
    If cnADO.State = adStateOpen Then cnADO.Close
    Set cnADO = Nothing
    
End Function

Public Function bolfncUwawakuShitajiLength(ByVal in_varLength As Variant, ByRef out_dblLength() As Double) As Boolean
'   *************************************************************
'   上枠下地長さ抽出
'   'ADD by K.Asayama 20170301
'   戻り値:Boolean
'       →True              上枠下地長さ抽出
'       →False             上枠下地長さ抽出不可
'
'    Input項目
'       in_varLength        上枠下地W
'       out_dblLength()     分割されたW(Falseの際は全て0) --Output

'   *************************************************************

    Dim dblLength As Double
    Dim dblWork As Double
    
    On Error GoTo Err_bolfncUwawakuShitajiLength
    
    '変数初期化
    bolfncUwawakuShitajiLength = False
    
    ReDim out_dblLength(1 To 3)
    dblWork = 0
    
    '長さが数字以外の場合はFalse
    If IsNumeric(in_varLength) Then
        '少数点第２位以下切り捨て
        dblLength = RoundDown(CDbl(in_varLength), 1)
        dblLength = dblFIVEorZERO(dblLength)
    Else
        Exit Function
    End If
    
    '上枠下地取得
    '2420mm以下はそのまま1本
    'それ以上は条件により分割する
    
    Select Case dblLength
        
        Case Is < 2420.5
            out_dblLength(1) = dblLength
            
        Case 2420.5 To 2720
            out_dblLength(1) = dblLength - 300
            out_dblLength(2) = 300
        
        Case 2720.5 To 4840
            out_dblLength(1) = dblLength - 2420
            out_dblLength(2) = 2420
        
        Case Is > 4840
            dblWork = dblFIVEorZERO(Roundx(dblLength / 3, 1))
            out_dblLength(1) = dblWork
            out_dblLength(2) = dblWork
            out_dblLength(3) = dblLength - (dblWork * 2)
            
    End Select
    
    bolfncUwawakuShitajiLength = True
    Exit Function

Err_bolfncUwawakuShitajiLength:
    MsgBox Err.Description, vbCritical, "上枠下地W分割エラー"
    
End Function

Public Function intFncUwawakuShitajiLengthGroup(in_dblLength As Double) As Integer
'   *************************************************************
'   上枠下地長さ集計グループ抽出
'   'ADD by K.Asayama 20170301
'   戻り値:Integer
'       →長さ
'
'    Input項目
'       in_dblLength        上枠下地W

'   *************************************************************
    intFncUwawakuShitajiLengthGroup = 0
    
    Select Case in_dblLength
        
        Case Is = 300
            intFncUwawakuShitajiLengthGroup = 300
            
        Case Is > 1820
            intFncUwawakuShitajiLengthGroup = 2430
                       
        Case Else
            intFncUwawakuShitajiLengthGroup = 1820
            
    End Select
    
End Function

Public Function fncstrUwawakuShitajiT(ByVal in_varHinban As Variant, ByVal in_varSagari As Variant, ByVal in_BoardT As Variant) As String
'   *************************************************************
'   上枠下地厚み抽出
'   'ADD by K.Asayama 20170301
'   戻り値:String
'       →厚み（数値だがA+Bの表記もあるので文字列型式で出力）
'
'    Input項目
'       in_varHinban        下地材品番
'       in_varSagari        下がり壁

'1.12.2
'   →ウォールスルー除外追加
'3.0.0
'   →ボード厚考慮 引数追加（BRD1908）
'   *************************************************************
    Dim varHinban As String
    Dim strSagari As String
    Dim strBoardT As String
    
    fncstrUwawakuShitajiT = ""
    
    On Error GoTo Err_fncstrUwawakuShitajiT
    
    If IsWallThru(in_varHinban) Then Exit Function
    
    If Not IsNull(in_varHinban) Then
        varHinban = Replace(in_varHinban, "特 ", "")
    Else
        Exit Function
    End If
    
    If in_varSagari = "有" Then
        strSagari = "有"
    Else
        strSagari = "無"
    End If
    
    Select Case Nz(in_BoardT, 0)
        Case 9.5
            strBoardT = "21"
        
        Case 12.5
            strBoardT = "18"
            
        Case 15
            strBoardT = "15"
            
        Case Else
            strBoardT = "30"
    End Select
    
    Select Case IsStealth_Seizo_TEMP(varHinban)
        'ステルス
        Case True

           Select Case strSagari
                '下がり壁
                Case "有"

                        If IsHirakido(varHinban) Or IsOyatobira(varHinban) Then
                            If varHinban Like "*G114*-####*" Then
                                fncstrUwawakuShitajiT = "30+9"
                            ElseIf IsSoftMotion(varHinban) Then
                                fncstrUwawakuShitajiT = "18+9"
                            Else
                                fncstrUwawakuShitajiT = "12+9"
                            End If
                        Else
                            If IsCloset_Hiraki(varHinban) Then
                                If varHinban Like "*G114*-####*" Then
                                    fncstrUwawakuShitajiT = "30"
                                Else
                                    fncstrUwawakuShitajiT = "12"
                                End If
                            Else
                                fncstrUwawakuShitajiT = strBoardT
                            End If
                        End If
                    
                '天井収まり
                Case "無"
                    
                    If IsSoftMotion(varHinban) Then
                        fncstrUwawakuShitajiT = "18"
                        
                    ElseIf IsOyatobira(varHinban) Then
                        fncstrUwawakuShitajiT = "12"
                        
                    ElseIf IsCloset_Hiraki(varHinban) Then
                        
                        fncstrUwawakuShitajiT = "12"
                        
                    ElseIf Not IsHirakido(varHinban) And Not IsCloset_Slide(varHinban) And Not IsYukazukeRail(varHinban) Then
                        
                        fncstrUwawakuShitajiT = strBoardT
                        
                    End If

                    
            End Select
            
        'インセット下地
        Case False
            
            Select Case strSagari
            
                '下がり壁
                Case "有"
                    If IsHirakido(varHinban) Or IsOyatobira(varHinban) Then
                        If varHinban Like "*G114*-####*" Then
                            fncstrUwawakuShitajiT = "30"
                        ElseIf IsSoftMotion(varHinban) Then
                            fncstrUwawakuShitajiT = "18"
                        Else
                            fncstrUwawakuShitajiT = "12"
                        End If
                    Else
                        fncstrUwawakuShitajiT = strBoardT
                    End If
                    
                '天井収まり
                Case "無"
                    If IsSoftMotion(varHinban) Then
                        fncstrUwawakuShitajiT = "18"
                    ElseIf IsOyatobira(varHinban) Then
                        fncstrUwawakuShitajiT = "12"
                    ElseIf Not IsHirakido(varHinban) Then
                        fncstrUwawakuShitajiT = strBoardT
                    End If
                    
                        
            End Select
    End Select
    
    Exit Function

Err_fncstrUwawakuShitajiT:
    
End Function

Public Function strfncFullHeightHinge(in_varHinban As Variant, in_varSpec As Variant) As String
'   *************************************************************
'   フルハイトヒンジメーカー確認

'   引数:
'       →建具品番
'         個別Spec

'
'   戻り値:String
'       →メーカー名
'
'  2.7.0 ADD

'   *************************************************************

    Dim datNOW As Date
    
    strfncFullHeightHinge = ""
    
    '2018/9/6以降から適用とする**************
    datNOW = Date
    If datNOW < #9/6/2018# Then Exit Function
    '****************************************
    
    If IsNull(in_varHinban) Then Exit Function
    If IsNull(in_varSpec) Then Exit Function
    
    If (IsHirakido(CStr(in_varHinban)) Or IsOyatobira(CStr(in_varHinban)) Or IsKotobira(CStr(in_varHinban))) And Not IsHidden_Hinge(in_varHinban) Then
    
        If left(in_varSpec, 3) = "BRD" Then
        
            If right(in_varSpec, 4) >= "1808" Then
                strfncFullHeightHinge = "NISHIMURA"
            Else
                strfncFullHeightHinge = "YOGO"
            End If

        Else
        
            strfncFullHeightHinge = "YOGO"
            
        End If
    End If
    
End Function

Public Function bolFncCloset_IseharaToso(in_varHinban As Variant) As Boolean
'   *************************************************************
'   伊勢原工場塗装クローゼット確認
'
'   戻り値:Boolean
'       →True              伊勢原工場塗装
'       →False             伊勢原工場塗装以外
'
'    Input項目
'       in_varHinban        建具（折戸）品番

'   2.14.0 ADD
'   *************************************************************
    
    Dim strHinban As String
    
    bolFncCloset_IseharaToso = False
    
    If IsNull(in_varHinban) Then Exit Function
        
    If IsCloset_Oredo(in_varHinban) Or IsCloset_Hiraki(in_varHinban) Then
        If in_varHinban Like "*-####*(NI)*" Then
            bolFncCloset_IseharaToso = True
        End If
    End If
    
End Function

Public Function strFncFuchibariColor(in_varHinban As Variant, in_strColor As String, in_varSpec As Variant) As String

'   *************************************************************
'   縁貼り色確認

'   戻り値:Boolean
'       →縁貼り色
'
'    Input項目
'       in_varHinban        建具品番
'       in_strColor         色
'       in_varSpec          個別Spec

'   2.14.0 ADD

'   3.0.0
'   →ジュリアの縁貼り色変更(BA→MO)
'   *************************************************************
    
    Dim strHinban As String
    
    strFncFuchibariColor = ""
    
    On Error GoTo Err_strFncFuchibariColor
    
    If in_strColor = "" Then Exit Function
    
    If IsNull(in_varHinban) Then
        strFncFuchibariColor = in_strColor
        Exit Function
    End If
    
    strHinban = Replace(in_varHinban, "特 ", "")
    
    If IsCarloGiulia(strHinban) Then
        If in_strColor = "SB" Then
            strFncFuchibariColor = "EW"
        ElseIf in_strColor = "SH" Then
            If Is40mm(in_varSpec) Then
                strFncFuchibariColor = "MO"
            Else
                strFncFuchibariColor = "BA"
            End If
        End If
    Else
        strFncFuchibariColor = in_strColor
    End If
    
    Exit Function

Err_strFncFuchibariColor:
    MsgBox Err.Description
    strFncFuchibariColor = ""
    
End Function

Public Function varFncKanamonoSeizoBi(ByVal strKeiyakuNo As String, ByVal strTouNo As String, ByVal strHeyaNo As String, ByVal varDate As Variant, ByVal bytHinbanKubun As Byte) As Variant
'   *************************************************************
'   varFncKanamonoSeizoBi
'       出荷金物がある場合、製造指示データから製造日（最大値）を返す
'
'   戻り値:Variant
'       →Date型            製造日
'       →Null              製造日なし、又はエラー
'
'    Input項目
'       strKeiyakuNo        契約番号
'       strTouNo            棟番号
'       strHeyaNo           部屋番号
'       varDate             製造日
'       bytHinbanKubun      品番区分 (1→建具）

'3.0.0 ADD
'   *************************************************************
    Dim strSQL As String
    
    varFncKanamonoSeizoBi = Null
    
    On Error GoTo Err_varFncKanamonoSeizoBi
    
    strSQL = ""
    strSQL = strSQL & "契約番号 = '" & strKeiyakuNo & "' and  棟番号 = '" & strTouNo & "' and 部屋番号 = '" & strHeyaNo & "' "
    strSQL = strSQL & "and  出荷金物あり = True AND 製造日 < #" & Format(varDate, "yyyy/MM/dd") & "# "
    
    varFncKanamonoSeizoBi = DMax("製造日", "TEMP_部材展開_製造日", strSQL)
    
    Exit Function
    
Err_varFncKanamonoSeizoBi:
    'Debug.Print Err.Description
    varFncKanamonoSeizoBi = Null
End Function

Public Function varFncKanamonoDaisha(ByVal strKeiyakuNo As String, ByVal strTouNo As String, ByVal strHeyaNo As String, ByVal bytHinbanKubun As Byte) As Variant
'   *************************************************************
'   varFncKanamonoDaisha
'       出荷金物が台車データに登録されている場合、製造指示データから台車コードを返す

'
'   戻り値:Variant
'       →String型          台車コード
'       →Null              製造日なし、又はエラー
'
'    Input項目
'       strKeiyakuNo        契約番号
'       strTouNo            棟番号
'       strHeyaNo           部屋番号
'       bytHinbanKubun      品番区分 (1→建具）

'3.0.0 ADD
'   *************************************************************
    Dim strSQL As String
    Dim strKeiyakuBango As String
    Dim bytKubun As Byte
    
    varFncKanamonoDaisha = Null
    
    On Error GoTo Err_varFncKanamonoDaisha
    
    strKeiyakuBango = strKeiyakuNo & "-" & strTouNo & "-" & strHeyaNo
    
    Select Case bytHinbanKubun
        Case 1
            bytKubun = 10
    End Select
    
    If Not IsNull(bytKubun) Then
            strSQL = ""
            strSQL = strSQL & "契約番号 = '" & strKeiyakuNo & "' and  棟番号 = '" & strTouNo & "' and 部屋番号 = '" & strHeyaNo & "' "
    
            varFncKanamonoDaisha = DMax("台車コード", "TEMP_部材展開_製造日", strSQL)
    End If
    
    Exit Function
    
Err_varFncKanamonoDaisha:
    'Debug.Print Err.Description
    varFncKanamonoDaisha = Null
End Function