Attribute VB_Name = "md_日付関数"
Option Compare Database
Option Explicit

Public Function intfncSeizoNissu(in_varHinban As Variant) As Integer
'   *************************************************************
'   建具製造所要日数確認
'   カタログに記載されている最短製造可能日数を返す
'
'   戻り値:Integer
'                       →  所要日数
'                           品番不正の場合は0を返す
'                           クローゼットは0を返す (伊勢原生産以外)
'
'    Input項目
'       in_strHinban        建具品番

'   *************************************************************

    If Not in_varHinban Like "*-####*-*" Then
        intfncSeizoNissu = 0
        Exit Function
    End If
    
    'Caro(Flushより先に記載する)
    If isCaro(in_varHinban) Then
    
        intfncSeizoNissu = 20
    'ｸﾛｾﾞｯﾄ(Flushより先に記載する)
    ElseIf in_varHinban Like "F*CME-####*-*" Then
    
        intfncSeizoNissu = 20
    'ｸﾛｾﾞｯﾄ(SINAより先に記載する)
    ElseIf in_varHinban Like "T*CME-####*-*" Then
    
        intfncSeizoNissu = 20
    'ｸﾛｾﾞｯﾄ
    ElseIf in_varHinban Like "P*CSA-####*-*" Then
    
        intfncSeizoNissu = 20
    'Flush
    ElseIf in_varHinban Like "F*-####*-*" Then
    
        intfncSeizoNissu = 13
    'F/S
    ElseIf in_varHinban Like "S*-####*-*" Then
    
        intfncSeizoNissu = 13
    'LUCENTE
    ElseIf in_varHinban Like "P*-####*-*" Then
    
        intfncSeizoNissu = 20
    'SINA
    ElseIf in_varHinban Like "T*-####*-*" Then
    
        intfncSeizoNissu = 20
    'Air
    ElseIf in_varHinban Like "A*-####*-*" Then
    
        intfncSeizoNissu = 20
    'MONSTER
    ElseIf in_varHinban Like "O*-####*-*" Then
    
        intfncSeizoNissu = 20
    'PALIO
    ElseIf in_varHinban Like "B*-####*-*" Then
    
        intfncSeizoNissu = 23
    'REALART
    ElseIf in_varHinban Like "R*-####*-*" Then
    
        intfncSeizoNissu = 23
        
    Else
    
        intfncSeizoNissu = 23
    
    End If
    
End Function

Public Function fncbolCalender_Replace() As Boolean
'   *************************************************************
'   ローカルカレンダー置換え処理
'   リモートデータベースからローカルにカレンダーデータをコピーする
'
'   戻り値:Boolean
'       →True              置換成功
'       →False             置換失敗
'
'   *************************************************************

    fncbolCalender_Replace = False
    
    Dim objREMOTEDB As New cls_BRAND_MASTER
    Dim objLOCALDB As New cls_LOCALDB
    
    Dim strSQL_Insert As String
    Dim strSQL As String
    strSQL_Insert = "Insert into WK_Calendar_工場(休日) values (#"
    
    '工場用コピー（T_Calendar_工場)
    If objLOCALDB.ExecSQL("delete from WK_Calendar_工場") Then
        strSQL = "select 休日 from T_Calendar_工場 "
        'strSQL = strSQL & "where convert(datetime,休日) > '" & "2015/01/01" & "'"
        If objREMOTEDB.ExecSelect(strSQL) Then
            Do While Not objREMOTEDB.GetRS.EOF
                If Not objLOCALDB.ExecSQL(strSQL_Insert & objREMOTEDB.GetRS![休日] & "#)") Then
                    Err.Raise 9999, , "休日カレンダー（工場）ローカルコピーエラー"
                End If
                objREMOTEDB.GetRS.MoveNext
            Loop
        End If
    End If
    
    strSQL_Insert = "Insert into WK_Calendar_業務(休日) values (#"
    
    '業務用コピー（T_Calendar)
    If objLOCALDB.ExecSQL("delete from WK_Calendar_業務") Then
        strSQL = "select 休日 from T_Calendar "
        'strSQL = strSQL & "where convert(datetime,休日) > '" & "2015/01/01" & "'"
        If objREMOTEDB.ExecSelect(strSQL) Then
            Do While Not objREMOTEDB.GetRS.EOF
                If Not objLOCALDB.ExecSQL(strSQL_Insert & objREMOTEDB.GetRS![休日] & "#)") Then
                    Err.Raise 9999, , "休日カレンダー（業務）ローカルコピーエラー"
                End If
                objREMOTEDB.GetRS.MoveNext
            Loop
            fncbolCalender_Replace = True
        End If
    End If
    
    GoTo Exit_fncbolCalender_Replace
    
Err_fncbolCalender_Replace:
    MsgBox Err.Description
    
Exit_fncbolCalender_Replace:
    Set objREMOTEDB = Nothing
    Set objLOCALDB = Nothing
End Function

Public Function bolfncCalc_DayOn(in_datNouhinDate As Variant, in_varHinban As Variant, in_intDays As Integer, out_datDay As Variant, out_datNextDay As Variant) As Boolean
'   *************************************************************
'   製造部門日付加算処理
'   工場カレンダーを参照しN日後の日付を返す（N営業日後）
'
'   戻り値:Boolean
'       →True              日付取得成功
'       →False             日付取得成功失敗
'
'    Input項目
'       in_datNouhinDate    Input用日付
'       in_varHinban        品番
'       in_intDays          加算日付
'    Output項目
'       out_datDay          Input用日付にin_intDaysを加算後の日付
'       out_datNextDay      out_datDayの1営業日後の日付(F框と技官製造扉以外はNull）
'   *************************************************************

    Dim objLOCALDB As New cls_LOCALDB
    
    Dim strSQL As String
    
    Dim datDayBefore As Date

    Dim datNextDay As Date
    
    Dim i As Integer, j As Integer
    
    bolfncCalc_DayOn = False
    i = in_intDays
    j = 0
    out_datDay = Null
    out_datNextDay = Null
    
    If Not IsDate(in_datNouhinDate) Then GoTo Err_bolfncCalc_DayOn
    
    datDayBefore = DateDiff("d", -1, in_datNouhinDate)
    
    strSQL = ""
    strSQL = strSQL & "select 休日 from WK_Calendar_工場 "
    strSQL = strSQL & "where 休日 > #" & in_datNouhinDate & "# "
    strSQL = strSQL & "order by 休日 "
    
    If objLOCALDB.ExecSelect(strSQL) Then
        Do While Not objLOCALDB.GetRS.EOF
            If datDayBefore = objLOCALDB.GetRS![休日] Then
                objLOCALDB.GetRS.MoveNext
            Else
                i = i - 1
            End If
            
            If i = 0 Then Exit Do
            
            datDayBefore = DateDiff("d", -1, datDayBefore)
            
        Loop
        
        If i <> 0 Then Err.Raise 9999, , "製造日取得エラー"
        
        out_datDay = datDayBefore
        
        '技官製造日
        If IsFkamachi(in_varHinban) Or IsGikan(in_varHinban) Then
                
            If Not bolfncNextDate(datDayBefore, out_datNextDay) Then
                Err.Raise 9999, , "技官（框）製造日取得エラー"
            End If
        
'            strSQL = ""
'            strSQL = strSQL & "select 休日 from WK_Calendar_工場 "
'            strSQL = strSQL & "where 休日 > #" & datDayBefore & "# "
'            strSQL = strSQL & "order by 休日 "
'
'            datNextDay = DateDiff("d", -1, datDayBefore)
'
'            If objLocalDB.ExecSelect(strSQL) Then
'                i = 1
'                Do While Not objLocalDB.GetRS.EOF
'
'                     If datNextDay = objLocalDB.GetRS![休日] Then
'                         objLocalDB.GetRS.MoveNext
'                     Else
'                         i = i - 1
'                     End If
'
'                     If i = 0 Then Exit Do
'
'                     datNextDay = DateDiff("d", -1, datNextDay)
'
'                Loop
'
'                If i <> 0 Then Err.Raise 9999, , "技官（框）製造日取得エラー"
'
'                out_datNextDay = datNextDay
'
'            Else
'                Err.Raise 9999, , "休日カレンダー取得エラー"
'            End If
'
        End If
    Else
        Err.Raise 9999, , "休日カレンダー取得エラー"
    End If
    
    
    bolfncCalc_DayOn = True
    GoTo Exit_bolfncCalc_DayOn
    
Err_bolfncCalc_DayOn:
    out_datDay = Null
    out_datNextDay = Null
    bolfncCalc_DayOn = False
    
Exit_bolfncCalc_DayOn:
    Set objLOCALDB = Nothing
    
End Function

Public Function bolfncCalc_DayOff(in_datNouhinDate As Variant, in_intDays As Integer, out_datDay As Variant, out_datNextDay As Variant) As Boolean
'   *************************************************************
'   製造部門日付減算処理
'   工場カレンダーを参照しN日前の日付を返す（N営業日後）
'
'   戻り値:Boolean
'       →True              日付取得成功
'       →False             日付取得成功失敗
'
'    Input項目
'       in_datNouhinDate    Input用日付
'       in_intDays          加算日付
'    Output項目
'       out_datDay          Input用日付にin_intDaysを加算後の日付
'       out_datNextDay      out_datDayの1営業日後の日付
'   *************************************************************

    Dim objLOCALDB As New cls_LOCALDB
    
    Dim strSQL As String
    
    Dim datDayBefore As Date

    Dim datNextDay As Date
    
    Dim i As Integer, j As Integer
    
    bolfncCalc_DayOff = False
    i = in_intDays
    j = 0
    out_datDay = Null
    out_datNextDay = Null
    
    If Not IsDate(in_datNouhinDate) Then GoTo Err_bolfncCalc_DayOff
    
    datDayBefore = DateDiff("d", 1, in_datNouhinDate)
    
    strSQL = ""
    strSQL = strSQL & "select 休日 from WK_Calendar_工場 "
    strSQL = strSQL & "where 休日 < #" & in_datNouhinDate & "# "
    strSQL = strSQL & "order by 休日 desc "
    
    If objLOCALDB.ExecSelect(strSQL) Then
        Do While Not objLOCALDB.GetRS.EOF
            If datDayBefore = objLOCALDB.GetRS![休日] Then
                objLOCALDB.GetRS.MoveNext
            Else
                i = i - 1
            End If
            
            If i = 0 Then Exit Do
            
            datDayBefore = DateDiff("d", 1, datDayBefore)
            
        Loop
        
        If i <> 0 Then Err.Raise 9999, , "製造日取得エラー"
        
        out_datDay = datDayBefore
        
        '技官製造日
        If Not bolfncNextDate(datDayBefore, out_datNextDay) Then
            Err.Raise 9999, , "技官（框）製造日取得エラー"
        End If
        
'            strSQL = ""
'            strSQL = strSQL & "select 休日 from WK_Calendar_工場 "
'            strSQL = strSQL & "where 休日 > #" & datDayBefore & "# "
'            strSQL = strSQL & "order by 休日 "
'
'            datNextDay = DateDiff("d", -1, datDayBefore)
'
'            If objLocalDB.ExecSelect(strSQL) Then
'                i = 1
'                Do While Not objLocalDB.GetRS.EOF
'
'                     If datNextDay = objLocalDB.GetRS![休日] Then
'                         objLocalDB.GetRS.MoveNext
'                     Else
'                         i = i - 1
'                     End If
'
'                     If i = 0 Then Exit Do
'
'                     datNextDay = DateDiff("d", -1, datNextDay)
'
'                Loop
'
'                If i <> 0 Then Err.Raise 9999, , "技官（框）製造日取得エラー"
'
'                out_datNextDay = datNextDay
'
'            Else
'                Err.Raise 9999, , "休日カレンダー取得エラー"
'            End If

    Else
        Err.Raise 9999, , "休日カレンダー取得エラー"
    End If
    
    
    bolfncCalc_DayOff = True
    GoTo Exit_bolfncCalc_DayOff
    
Err_bolfncCalc_DayOff:
    out_datDay = Null
    out_datNextDay = Null
    bolfncCalc_DayOff = False
    
Exit_bolfncCalc_DayOff:
    Set objLOCALDB = Nothing
    
End Function

Public Function bolfncNextDate(in_datStartDate As Variant, ByRef out_datNextDay As Variant) As Boolean
'   *************************************************************
'   製造部門日付加算処理（翌日）
'   input日付の翌営業日を取得
'
'   戻り値:Boolean
'       →True              日付取得成功
'       →False             日付取得成功失敗
'
'    Input項目
'       in_datStartDate     Input用日付
'    Output項目
'       out_datNextDay      Input用日付の1営業日後の日付
'   *************************************************************
    Dim objLOCALDB As New cls_LOCALDB
    
    Dim strSQL As String
    Dim datNextDay As Date
    Dim i As Integer
    bolfncNextDate = False
    
    strSQL = ""
    strSQL = strSQL & "select 休日 from WK_Calendar_工場 "
    strSQL = strSQL & "where 休日 > #" & in_datStartDate & "# "
    strSQL = strSQL & "order by 休日 "
    
    datNextDay = DateDiff("d", -1, in_datStartDate)
    
    If objLOCALDB.ExecSelect(strSQL) Then
        i = 1
        Do While Not objLOCALDB.GetRS.EOF
        
             If datNextDay = objLOCALDB.GetRS![休日] Then
                 objLOCALDB.GetRS.MoveNext
             Else
                 i = i - 1
             End If
             
             If i = 0 Then Exit Do
             
             datNextDay = DateDiff("d", -1, datNextDay)
        
        Loop
        
        If i <> 0 Then Err.Raise 9999, , "技官（框）製造日取得エラー"
        
        out_datNextDay = datNextDay
        
    Else
        Err.Raise 9999, , "休日カレンダー取得エラー（技官製造日）"
    End If
            
    bolfncNextDate = True
    GoTo Exit_bolfncNextDate
    
Err_bolfncNextDate:
    out_datNextDay = Null
    bolfncNextDate = False
    
Exit_bolfncNextDate:
    Set objLOCALDB = Nothing
    
End Function

Public Function fncbolSyukkaBiFromAddress(in_varAddress As Variant, in_varNouhinBi As Variant, ByRef out_SyukkaBi As Variant, ByRef out_MinusDay As Integer) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'住所から出荷日取得
'   →納品先住所から配送日数を引き出し、出荷日を作成する
'
'-------------------------------------------------------
'20151021 K.Asayama フォームモジュールから移動
'-------------------------------------------------------
'
'   :引数
'       in_varAddress       :納付先住所
'       in_varNouhinBi      :納品日
'       out_SyukkaBi        :出荷日（出力）　取得できない場合はNull
'       out_MinusDay        :納品日-出荷日（営業日数）

'
'   :戻り値
'       True            :取得成功
'       False           :取得失敗
'--------------------------------------------------------------------------------------------------------------------
    Dim objLOCALDB As New cls_LOCALDB
    Dim intMinusDays As Integer
    Dim datTMPSyukkaBi As Date
    Dim datTMPKeisan As Date
    Dim i As Integer
    Dim strSQL As String
    
    fncbolSyukkaBiFromAddress = False
    strSQL = ""
    
    On Error GoTo Err_fncbolSyukkaBiFromAddress
    
    If IsNull(in_varAddress) Then
        Exit Function
    End If
    
    '以下に該当する都道府県の場合は2日
    If in_varAddress Like "北海道*" Or _
        in_varAddress Like "青森県*" Or in_varAddress Like "岩手県*" Or in_varAddress Like "秋田県*" Or _
        in_varAddress Like "宮城県*" Or in_varAddress Like "福島県*" Or in_varAddress Like "山形県*" Or _
        in_varAddress Like "三重県*" Or in_varAddress Like "兵庫県*" Or in_varAddress Like "和歌山県*" Or _
        in_varAddress Like "島根県*" Or in_varAddress Like "鳥取県*" Or in_varAddress Like "山口県*" Or _
        in_varAddress Like "広島県*" Or in_varAddress Like "岡山県*" Or in_varAddress Like "香川県*" Or _
        in_varAddress Like "愛媛県*" Or in_varAddress Like "徳島県*" Or in_varAddress Like "高知県*" Or _
        in_varAddress Like "福岡県*" Or in_varAddress Like "大分県*" Or in_varAddress Like "佐賀県*" Or _
        in_varAddress Like "長崎県*" Or in_varAddress Like "宮崎県*" Or in_varAddress Like "熊本県*" Or _
        in_varAddress Like "鹿児島県*" Or _
        in_varAddress Like "沖縄県*" Then
       
            intMinusDays = 2
    Else
    
            intMinusDays = 1
    End If
    
    '画面表示用
    out_MinusDay = intMinusDays
    
    '------------------------------------------------------------
    '出荷日と納品日の間に日、祝が含まれている場合はその日数を加算
    '（土曜は配送日に含まれる）
    datTMPKeisan = in_varNouhinBi
        
    i = intMinusDays
    
    While i <> 0
        '祝日、日曜だった場合は1日加算
        If ktHolidayName(datTMPKeisan) <> "" Or Weekday(datTMPKeisan, vbSunday) = 1 Then '祝日又は日曜
            intMinusDays = intMinusDays + 1
        Else
            i = i - 1
            
        End If
        
        '日付から1引く
        datTMPKeisan = DateDiff("d", 1, datTMPKeisan)
    Wend
    '------------------------------------------------------------
    
    '出荷日取得
    datTMPSyukkaBi = DateDiff("d", intMinusDays, in_varNouhinBi)
    
    '出荷日が土日祝でないかチェック（営業の土曜日でも出荷はしない）
    Do
        If ktHolidayName(datTMPSyukkaBi) = "" Then '祝日でない
            If Weekday(datTMPSyukkaBi, vbSunday) = 1 Or Weekday(datTMPSyukkaBi, vbSunday) = 7 Then '日か土
                
            Else    '平日
                Exit Do
            End If
        End If
        
        datTMPSyukkaBi = DateDiff("d", 1, datTMPSyukkaBi)
        
    Loop
    
    '会社が休日の場合は前営業日を返す
    strSQL = ""
    strSQL = strSQL & "select 休日 from WK_Calendar_業務 "
    strSQL = strSQL & "where 休日 =< #" & datTMPSyukkaBi & "# "
    strSQL = strSQL & "order by 休日 desc "
    
    If objLOCALDB.ExecSelect(strSQL) Then
        Do While Not objLOCALDB.GetRS.EOF
            If datTMPSyukkaBi <> objLOCALDB.GetRS![休日] Then
                Exit Do
            End If
            
            datTMPSyukkaBi = DateDiff("d", 1, datTMPSyukkaBi)
            objLOCALDB.GetRS.MoveNext
            
        Loop
    End If
    
    out_SyukkaBi = datTMPSyukkaBi
    fncbolSyukkaBiFromAddress = True
    
    GoTo Exit_fncbolSyukkaBiFromAddress
    
Err_fncbolSyukkaBiFromAddress:

Exit_fncbolSyukkaBiFromAddress:
    Set objLOCALDB = Nothing
End Function

Public Function IsHoliday(ByVal in_Date As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'   製造部門休日確認処理
'   製造部門が休日かどうか確認
'

'   Ver 1.01.* K.Asayama ADD 201510**
'
'   戻り値:Boolean
'       →True              休日
'       →False             稼働日
'
'    Input項目
'       in_Date     日付（文字列型式）
'--------------------------------------------------------------------------------------------------------------------

    Dim objLOCALDB As New cls_LOCALDB
    
    Dim strSQL As String
    
    On Error GoTo Err_IsHoliday
    
    If Not IsDate(in_Date) Then GoTo Err_IsHoliday
    
    strSQL = ""
    strSQL = strSQL & "select 休日 from WK_Calendar_工場 "
    strSQL = strSQL & "where 休日 = #" & in_Date & "# "
    
    
    If objLOCALDB.ExecSelect(strSQL) Then
        If Not objLOCALDB.GetRS.EOF Then
            IsHoliday = True
        End If
    End If
        
    GoTo Exit_IsHoliday

Err_IsHoliday:
    IsHoliday = False
    
Exit_IsHoliday:
    Set objLOCALDB = Nothing
End Function
