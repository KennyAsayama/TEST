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
                        strSQL = strSQL & "," & varNullChk(fncstrUwawakuShitajiT(varHinban, rsADO![下がり壁]), 1) & " "
                        
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
                Err.Raise "以降の帳票の出力を中止します"
            End If
        Else '図面あり
            If fncstrUwawakuShitajiT(varHinban, rsADO![下がり壁]) <> "" Then
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

Public Function fncstrUwawakuShitajiT(ByVal in_varHinban As Variant, ByVal in_varSagari As Variant) As String
'   *************************************************************
'   上枠下地厚み抽出
'   'ADD by K.Asayama 20170301
'   戻り値:String
'       →厚み（数値だがA+Bの表記もあるので文字列型式で出力）
'
'    Input項目
'       in_varHinban        下地材品番
'       in_varSagari        下がり壁

'   *************************************************************
    Dim varHinban As String
    Dim strSagari As String
    
    fncstrUwawakuShitajiT = ""
    
    On Error GoTo Err_fncstrUwawakuShitajiT
    
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
                                fncstrUwawakuShitajiT = "30"
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
                        
                        fncstrUwawakuShitajiT = "30"
                        
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
                        fncstrUwawakuShitajiT = "30"
                    End If
                    
                '天井収まり
                Case "無"
                    If IsSoftMotion(varHinban) Then
                        fncstrUwawakuShitajiT = "18"
                    ElseIf IsOyatobira(varHinban) Then
                        fncstrUwawakuShitajiT = "12"
                    ElseIf Not IsHirakido(varHinban) Then
                        fncstrUwawakuShitajiT = "30"
                    End If
                    
                        
            End Select
    End Select
    
    Exit Function

Err_fncstrUwawakuShitajiT:
    
End Function