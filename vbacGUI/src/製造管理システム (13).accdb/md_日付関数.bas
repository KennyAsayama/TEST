Option Compare Database
Option Explicit

Public Function intfncSeizoNissu(in_varHinban As Variant) As Integer
'   *************************************************************
'   寶嬶惢憿強梫擔悢妋擣
'   僇僞儘僌偵婰嵹偝傟偰偄傞嵟抁惢憿壜擻擔悢傪曉偡
'
'   栠傝抣:Integer
'                       仺  強梫擔悢
'                           昳斣晄惓偺応崌偼0傪曉偡
'                           僋儘乕僛僢僩偼0傪曉偡 (埳惃尨惗嶻埲奜)
'
'    Input崁栚
'       in_strHinban        寶嬶昳斣
'
'   1.10.7
'           仺 惢昳娭悢偵抲姺偊
'   *************************************************************

    If Not in_varHinban Like "*-####*-*" Then
        intfncSeizoNissu = 0
        Exit Function
    End If
    
    'Caro(Flush傛傝愭偵婰嵹偡傞)
    If isCaro(in_varHinban) Then
    
        intfncSeizoNissu = 20
    '港巨(Flush傛傝愭偵婰嵹偡傞)
    ElseIf in_varHinban Like "F*CME-####*-*" Then
    
        intfncSeizoNissu = 20
    '港巨(SINA傛傝愭偵婰嵹偡傞)
    ElseIf in_varHinban Like "T*CME-####*-*" Then
    
        intfncSeizoNissu = 20
    '港巨
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
    ElseIf IsAir(in_varHinban) Then
    
        intfncSeizoNissu = 20
    'MONSTER
    ElseIf IsMonster(in_varHinban) Then
    
        intfncSeizoNissu = 20
    'PALIO
    ElseIf IsPALIO(in_varHinban) Then
    
        intfncSeizoNissu = 23
    'REALART
    ElseIf IsREALART(in_varHinban) Then
    
        intfncSeizoNissu = 23
        
    Else
    
        intfncSeizoNissu = 23
    
    End If
    
End Function

Public Function fncbolCalender_Replace() As Boolean
'   *************************************************************
'   儘乕僇儖僇儗儞僟乕抲姺偊張棟
'   儕儌乕僩僨乕僞儀乕僗偐傜儘乕僇儖偵僇儗儞僟乕僨乕僞傪僐僺乕偡傞
'
'   栠傝抣:Boolean
'       仺True              抲姺惉岟
'       仺False             抲姺幐攕
'
'   1.10.6 K.Asayama ADD 20151211 僐僺乕嵪傒偺応崌(bolCalendarCopy=True乯偼張棟偟側偄
'   *************************************************************

    fncbolCalender_Replace = False
    
    If bolCalendarCopy Then
        fncbolCalender_Replace = True
        Exit Function
    End If
    
    Dim objREMOTEDB As New cls_BRAND_MASTER
    Dim objLOCALDB As New cls_LOCALDB
    
    Dim strSQL_Insert As String
    Dim strSQL As String
    
    '1.10.5 ADD By Asayama 僄儔乕捛壛 20151209
    On Error GoTo Err_fncbolCalender_Replace
    
    strSQL_Insert = "Insert into WK_Calendar_岺応(媥擔) values (#"
    
    '岺応梡僐僺乕乮T_Calendar_岺応)
    If objLOCALDB.ExecSQL("delete from WK_Calendar_岺応") Then
        strSQL = "select 媥擔 from T_Calendar_岺応 "
        'strSQL = strSQL & "where convert(datetime,媥擔) > '" & "2015/01/01" & "'"
        If objREMOTEDB.ExecSelect(strSQL) Then
            Do While Not objREMOTEDB.GetRS.EOF
                If Not objLOCALDB.ExecSQL(strSQL_Insert & objREMOTEDB.GetRS![媥擔] & "#)") Then
                    Err.Raise 9999, , "媥擔僇儗儞僟乕乮岺応乯儘乕僇儖僐僺乕僄儔乕"
                End If
                objREMOTEDB.GetRS.MoveNext
            Loop
        End If
    End If
    
    strSQL_Insert = "Insert into WK_Calendar_嬈柋(媥擔) values (#"
    
    '嬈柋梡僐僺乕乮T_Calendar)
    If objLOCALDB.ExecSQL("delete from WK_Calendar_嬈柋") Then
        strSQL = "select 媥擔 from T_Calendar "
        'strSQL = strSQL & "where convert(datetime,媥擔) > '" & "2015/01/01" & "'"
        If objREMOTEDB.ExecSelect(strSQL) Then
            Do While Not objREMOTEDB.GetRS.EOF
                If Not objLOCALDB.ExecSQL(strSQL_Insert & objREMOTEDB.GetRS![媥擔] & "#)") Then
                    Err.Raise 9999, , "媥擔僇儗儞僟乕乮嬈柋乯儘乕僇儖僐僺乕僄儔乕"
                End If
                objREMOTEDB.GetRS.MoveNext
            Loop
            fncbolCalender_Replace = True
        End If
    End If
    
    '1.10.6 K.Asayama ADD 20151211 僐僺乕姰椆偺応崌嫟捠僼儔僌傪True偵偡傞
    bolCalendarCopy = True
    
    GoTo Exit_fncbolCalender_Replace
    
Err_fncbolCalender_Replace:
    MsgBox Err.Description
    
Exit_fncbolCalender_Replace:
    Set objREMOTEDB = Nothing
    Set objLOCALDB = Nothing
End Function

Public Function bolfncCalc_DayOn(in_datNouhinDate As Variant, in_varHinban As Variant, in_intDays As Integer, out_datDay As Variant, out_datNextDay As Variant) As Boolean
'   *************************************************************
'   惢憿晹栧擔晅壛嶼張棟
'   岺応僇儗儞僟乕傪嶲徠偟N擔屻偺擔晅傪曉偡乮N塩嬈擔屻乯
'
'   栠傝抣:Boolean
'       仺True              擔晅庢摼惉岟
'       仺False             擔晅庢摼惉岟幐攕
'
'    Input崁栚
'       in_datNouhinDate    Input梡擔晅
'       in_varHinban        昳斣
'       in_intDays          壛嶼擔晅
'    Output崁栚
'       out_datDay          Input梡擔晅偵in_intDays傪壛嶼屻偺擔晅
'       out_datNextDay      out_datDay偺1塩嬈擔屻偺擔晅(F瀥偲媄姱惢憿斷埲奜偼Null乯
'   *************************************************************

    Dim objLOCALDB As New cls_LOCALDB
    
    Dim strSQL As String
    
    Dim datDayBefore As Date

    Dim datNextDay As Date
    
    Dim i As Integer, j As Integer
    
    bolfncCalc_DayOn = False
    
    '1.10.5 ADD By Asayama 僄儔乕捛壛 20151209
    On Error GoTo Err_bolfncCalc_DayOn
    
    i = in_intDays
    j = 0
    out_datDay = Null
    out_datNextDay = Null
    
    If Not IsDate(in_datNouhinDate) Then GoTo Err_bolfncCalc_DayOn
    
    datDayBefore = DateDiff("d", -1, in_datNouhinDate)
 
    strSQL = ""
    strSQL = strSQL & "select 媥擔 from WK_Calendar_岺応 "
    strSQL = strSQL & "where 媥擔 > #" & in_datNouhinDate & "# "
    strSQL = strSQL & "order by 媥擔 "
    
    If objLOCALDB.ExecSelect(strSQL) Then
        Do While Not objLOCALDB.GetRS.EOF
            If datDayBefore = objLOCALDB.GetRS![媥擔] Then
                objLOCALDB.GetRS.MoveNext
            Else
                i = i - 1
            End If
            
            If i = 0 Then Exit Do
            
            datDayBefore = DateDiff("d", -1, datDayBefore)
            
        Loop
        
        If i <> 0 Then Err.Raise 9999, , "惢憿擔庢摼僄儔乕"
        
        out_datDay = datDayBefore
        
        '媄姱惢憿擔
        If IsFkamachi(in_varHinban) Or IsGikan(in_varHinban) Then
                
            If Not bolfncNextDate(datDayBefore, out_datNextDay) Then
                Err.Raise 9999, , "媄姱乮瀥乯惢憿擔庢摼僄儔乕"
            End If
        
'            strSQL = ""
'            strSQL = strSQL & "select 媥擔 from WK_Calendar_岺応 "
'            strSQL = strSQL & "where 媥擔 > #" & datDayBefore & "# "
'            strSQL = strSQL & "order by 媥擔 "
'
'            datNextDay = DateDiff("d", -1, datDayBefore)
'
'            If objLocalDB.ExecSelect(strSQL) Then
'                i = 1
'                Do While Not objLocalDB.GetRS.EOF
'
'                     If datNextDay = objLocalDB.GetRS![媥擔] Then
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
'                If i <> 0 Then Err.Raise 9999, , "媄姱乮瀥乯惢憿擔庢摼僄儔乕"
'
'                out_datNextDay = datNextDay
'
'            Else
'                Err.Raise 9999, , "媥擔僇儗儞僟乕庢摼僄儔乕"
'            End If
'
        End If
    Else
        Err.Raise 9999, , "媥擔僇儗儞僟乕庢摼僄儔乕"
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
'   惢憿晹栧擔晅尭嶼張棟
'   岺応僇儗儞僟乕傪嶲徠偟N擔慜偺擔晅傪曉偡乮N塩嬈擔屻乯
'
'   栠傝抣:Boolean
'       仺True              擔晅庢摼惉岟
'       仺False             擔晅庢摼惉岟幐攕
'
'    Input崁栚
'       in_datNouhinDate    Input梡擔晅
'       in_intDays          壛嶼擔晅
'    Output崁栚
'       out_datDay          Input梡擔晅偵in_intDays傪壛嶼屻偺擔晅
'       out_datNextDay      out_datDay偺1塩嬈擔屻偺擔晅

'   *************************************************************

    Dim objLOCALDB As New cls_LOCALDB
    
    Dim strSQL As String
    
    Dim datDayBefore As Date

    Dim datNextDay As Date
    
    Dim i As Integer, j As Integer
    
    bolfncCalc_DayOff = False
    
    '1.10.5 ADD By Asayama 僄儔乕捛壛 20151209
    On Error GoTo Err_bolfncCalc_DayOff
    
    i = in_intDays
    j = 0
    out_datDay = Null
    out_datNextDay = Null
    
    If Not IsDate(in_datNouhinDate) Then GoTo Err_bolfncCalc_DayOff
    
    datDayBefore = DateDiff("d", 1, in_datNouhinDate)

    strSQL = ""
    strSQL = strSQL & "select 媥擔 from WK_Calendar_岺応 "
    strSQL = strSQL & "where 媥擔 < #" & in_datNouhinDate & "# "
    strSQL = strSQL & "order by 媥擔 desc "
    
    If objLOCALDB.ExecSelect(strSQL) Then
        Do While Not objLOCALDB.GetRS.EOF
            If datDayBefore = objLOCALDB.GetRS![媥擔] Then
                objLOCALDB.GetRS.MoveNext
            Else
                i = i - 1
            End If
            
            If i = 0 Then Exit Do
            
            datDayBefore = DateDiff("d", 1, datDayBefore)
            
        Loop
        
        If i <> 0 Then Err.Raise 9999, , "惢憿擔庢摼僄儔乕"
        
        out_datDay = datDayBefore
        
        '媄姱惢憿擔
        If Not bolfncNextDate(datDayBefore, out_datNextDay) Then
            Err.Raise 9999, , "媄姱乮瀥乯惢憿擔庢摼僄儔乕"
        End If
        
'            strSQL = ""
'            strSQL = strSQL & "select 媥擔 from WK_Calendar_岺応 "
'            strSQL = strSQL & "where 媥擔 > #" & datDayBefore & "# "
'            strSQL = strSQL & "order by 媥擔 "
'
'            datNextDay = DateDiff("d", -1, datDayBefore)
'
'            If objLocalDB.ExecSelect(strSQL) Then
'                i = 1
'                Do While Not objLocalDB.GetRS.EOF
'
'                     If datNextDay = objLocalDB.GetRS![媥擔] Then
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
'                If i <> 0 Then Err.Raise 9999, , "媄姱乮瀥乯惢憿擔庢摼僄儔乕"
'
'                out_datNextDay = datNextDay
'
'            Else
'                Err.Raise 9999, , "媥擔僇儗儞僟乕庢摼僄儔乕"
'            End If

    Else
        Err.Raise 9999, , "媥擔僇儗儞僟乕庢摼僄儔乕"
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
'   惢憿晹栧擔晅壛嶼張棟乮梻擔乯
'   input擔晅偺梻塩嬈擔傪庢摼
'
'   栠傝抣:Boolean
'       仺True              擔晅庢摼惉岟
'       仺False             擔晅庢摼惉岟幐攕
'
'    Input崁栚
'       in_datStartDate     Input梡擔晅
'    Output崁栚
'       out_datNextDay      Input梡擔晅偺1塩嬈擔屻偺擔晅

'   *************************************************************
    Dim objLOCALDB As New cls_LOCALDB
    
    Dim strSQL As String
    Dim datNextDay As Date
    Dim i As Integer
    
    bolfncNextDate = False
    
    '1.10.5 ADD By Asayama 僄儔乕捛壛 20151209
    On Error GoTo Err_bolfncNextDate
    
    strSQL = ""
    strSQL = strSQL & "select 媥擔 from WK_Calendar_岺応 "
    strSQL = strSQL & "where 媥擔 > #" & in_datStartDate & "# "
    strSQL = strSQL & "order by 媥擔 "
    
    datNextDay = DateDiff("d", -1, in_datStartDate)
    
    If objLOCALDB.ExecSelect(strSQL) Then
        i = 1
        Do While Not objLOCALDB.GetRS.EOF
        
             If datNextDay = objLOCALDB.GetRS![媥擔] Then
                 objLOCALDB.GetRS.MoveNext
             Else
                 i = i - 1
             End If
             
             If i = 0 Then Exit Do
             
             datNextDay = DateDiff("d", -1, datNextDay)
        
        Loop
        
        If i <> 0 Then Err.Raise 9999, , "媄姱乮瀥乯惢憿擔庢摼僄儔乕"
        
        out_datNextDay = datNextDay
        
    Else
        Err.Raise 9999, , "媥擔僇儗儞僟乕庢摼僄儔乕乮媄姱惢憿擔乯"
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
'廧強偐傜弌壸擔庢摼
'   仺擺昳愭廧強偐傜攝憲擔悢傪堷偒弌偟丄弌壸擔傪嶌惉偡傞
'
'-------------------------------------------------------
'20151021 K.Asayama 僼僅乕儉儌僕儏乕儖偐傜堏摦
'-------------------------------------------------------
'
'   :堷悢
'       in_varAddress       :擺晅愭廧強
'       in_varNouhinBi      :擺昳擔
'       out_SyukkaBi        :弌壸擔乮弌椡乯丂庢摼偱偒側偄応崌偼Null
'       out_MinusDay        :擺昳擔-弌壸擔乮塩嬈擔悢乯

'
'   :栠傝抣
'       True            :庢摼惉岟
'       False           :庢摼幐攕

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
   
    '埲壓偵奩摉偡傞搒摴晎導偺応崌偼2擔
    If in_varAddress Like "杒奀摴*" Or _
        in_varAddress Like "惵怷導*" Or in_varAddress Like "娾庤導*" Or in_varAddress Like "廐揷導*" Or _
        in_varAddress Like "媨忛導*" Or in_varAddress Like "暉搰導*" Or in_varAddress Like "嶳宍導*" Or _
        in_varAddress Like "嶰廳導*" Or in_varAddress Like "暫屔導*" Or in_varAddress Like "榓壧嶳導*" Or _
        in_varAddress Like "搰崻導*" Or in_varAddress Like "捁庢導*" Or in_varAddress Like "嶳岥導*" Or _
        in_varAddress Like "峀搰導*" Or in_varAddress Like "壀嶳導*" Or in_varAddress Like "崄愳導*" Or _
        in_varAddress Like "垽昋導*" Or in_varAddress Like "摽搰導*" Or in_varAddress Like "崅抦導*" Or _
        in_varAddress Like "暉壀導*" Or in_varAddress Like "戝暘導*" Or in_varAddress Like "嵅夑導*" Or _
        in_varAddress Like "挿嶈導*" Or in_varAddress Like "媨嶈導*" Or in_varAddress Like "孎杮導*" Or _
        in_varAddress Like "幁帣搰導*" Or _
        in_varAddress Like "壂撽導*" Then
       
            intMinusDays = 2
    Else
    
            intMinusDays = 1
    End If
    
    '夋柺昞帵梡
    out_MinusDay = intMinusDays
    
    '------------------------------------------------------------
    '弌壸擔偲擺昳擔偺娫偵擔丄廽偑娷傑傟偰偄傞応崌偼偦偺擔悢傪壛嶼
    '乮搚梛偼攝憲擔偵娷傑傟傞乯
    datTMPKeisan = in_varNouhinBi
        
    i = intMinusDays
    
    While i <> 0
        '廽擔丄擔梛偩偭偨応崌偼1擔壛嶼
        If ktHolidayName(datTMPKeisan) <> "" Or Weekday(datTMPKeisan, vbSunday) = 1 Then '廽擔枖偼擔梛
            intMinusDays = intMinusDays + 1
        Else
            i = i - 1
            
        End If
        
        '擔晅偐傜1堷偔
        datTMPKeisan = DateDiff("d", 1, datTMPKeisan)
    Wend
    '------------------------------------------------------------
    
    '弌壸擔庢摼
    datTMPSyukkaBi = DateDiff("d", intMinusDays, in_varNouhinBi)
    
    '弌壸擔偑搚擔廽偱側偄偐僠僃僢僋乮塩嬈偺搚梛擔偱傕弌壸偼偟側偄乯
    Do
        If ktHolidayName(datTMPSyukkaBi) = "" Then '廽擔偱側偄
            If Weekday(datTMPSyukkaBi, vbSunday) = 1 Or Weekday(datTMPSyukkaBi, vbSunday) = 7 Then '擔偐搚
                
            Else    '暯擔
                Exit Do
            End If
        End If
        
        datTMPSyukkaBi = DateDiff("d", 1, datTMPSyukkaBi)
        
    Loop
    
    '夛幮偑媥擔偺応崌偼慜塩嬈擔傪曉偡
    strSQL = ""
    strSQL = strSQL & "select 媥擔 from WK_Calendar_嬈柋 "
    strSQL = strSQL & "where 媥擔 =< #" & datTMPSyukkaBi & "# "
    strSQL = strSQL & "order by 媥擔 desc "
    
    If objLOCALDB.ExecSelect(strSQL) Then
        Do While Not objLOCALDB.GetRS.EOF
            If datTMPSyukkaBi <> objLOCALDB.GetRS![媥擔] Then
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
'   惢憿晹栧媥擔妋擣張棟
'   惢憿晹栧偑媥擔偐偳偆偐妋擣
'

'   Ver 1.01.* K.Asayama ADD 201510**
'
'   栠傝抣:Boolean
'       仺True              媥擔
'       仺False             壱摥擔
'
'    Input崁栚
'       in_Date     擔晅乮暥帤楍宆幃乯

'--------------------------------------------------------------------------------------------------------------------

    Dim objLOCALDB As New cls_LOCALDB
    
    Dim strSQL As String
    
    On Error GoTo Err_IsHoliday
    
    If Not IsDate(in_Date) Then GoTo Err_IsHoliday
    
    strSQL = ""
    strSQL = strSQL & "select 媥擔 from WK_Calendar_岺応 "
    strSQL = strSQL & "where 媥擔 = #" & in_Date & "# "
    
    
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

Public Function intfncSeizoNissu_FromSyukkaBi(in_varHinban As Variant, in_intDefaultDays As Integer) As Integer
'   *************************************************************
'   寶嬶惢憿強梫擔悢妋擣乮弌壸擔傛傝媡嶼乯
'   弌壸擔傛傝惢憿壜擻擔傪寁嶼偡傞
'
'   1.10.7 ADD
'
'   栠傝抣:Integer
'                       仺  強梫擔悢
'                           昳斣晄惓偺応崌偼嵟戝擔悢乮揾憰斷乯傪曉偡
'                           僋儘乕僛僢僩偼0傪曉偡 (埳惃尨惗嶻埲奜)
'
'    Input崁栚
'       in_strHinban        寶嬶昳斣
'       in_intDefaultDays   昗弨昳(CUBE摍強梫擔悢乯
'   *************************************************************

    If Not in_varHinban Like "*-####*-*" Then
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 13
        Exit Function
    End If
    
    'Caro(Flush傛傝愭偵婰嵹偡傞)
    If isCaro(in_varHinban) Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
    '港巨(Flush傛傝愭偵婰嵹偡傞)
    ElseIf in_varHinban Like "F*CME-####*-*" Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
    '港巨(SINA傛傝愭偵婰嵹偡傞)
    ElseIf in_varHinban Like "T*CME-####*-*" Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
    '港巨
    ElseIf in_varHinban Like "P*CSA-####*-*" Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
    'Flush
    ElseIf in_varHinban Like "F*-####*-*" Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays
    'F/S
    ElseIf in_varHinban Like "S*-####*-*" Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays
    'LUCENTE
    ElseIf in_varHinban Like "P*-####*-*" Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
    'SINA
    ElseIf in_varHinban Like "T*-####*-*" Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
    'Air
    ElseIf IsAir(in_varHinban) Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
    'MONSTER
    ElseIf IsMonster(in_varHinban) Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 7
    'PALIO
    ElseIf IsPALIO(in_varHinban) Then
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 9
    'REALART
    ElseIf IsREALART(in_varHinban) Then
        If IsPainted(in_varHinban) Then
            intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 9
        Else
            intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays
        End If
        
    Else
    
        intfncSeizoNissu_FromSyukkaBi = in_intDefaultDays + 9
    
    End If
    
End Function