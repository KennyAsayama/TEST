Option Compare Database
Option Explicit

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/  --- VB / VBA 版 ( Update: 2018/12/8 ) ---
'_/
'_/  CopyRight(C) K.Tsunoda(AddinBox) 2001 All Rights Reserved.
'_/  ( AddinBox  http://addinbox.sakura.ne.jp/index.htm )
'_/  (  旧サイト  http://www.h3.dion.ne.jp/~sakatsu/index.htm )
'_/
'_/    この祝日マクロは『kt関数アドイン』で使用しているものです。
'_/    このロジックは、レスポンスを第一義として、可能な限り少ない
'_/    【条件判定の実行】で結果を出せるように設計してあります。
'_/
'_/    この関数では以下の祝日変更までサポートしています。
'_/    (a) 2019年施行の「天皇誕生日の変更」 12/23⇒2/23 (補：2019年には[天皇誕生日]はありません)
'_/    (b) 2019年の徳仁親王の即位日(5/1) および
'_/       祝日に挟まれて「国民の休日」となる 4/30(平成天皇の退位日) ＆ 5/2 の２休日
'_/    (c) 2019年の「即位の礼 正殿の儀 (10/22) 」
'_/    (d) 2020年施行の「体育の日の改名」⇒スポーツの日
'_/    (e) 五輪特措法による2020年の「祝日移動」
'_/       海の日：7/20(3rd Mon)⇒7/23, スポーツの日:10/12(2nd Mon)⇒7/24, 山の日：8/11⇒8/10
'_/
'_/  (*1)このマクロを引用するに当たっては、必ずこのコメントも
'_/      一緒に引用する事とします。
'_/  (*2)他サイト上で本マクロを直接引用する事は、ご遠慮願います。
'_/      【 http://addinbox.sakura.ne.jp/holiday_logic.htm 】
'_/      へのリンクによる紹介で対応して下さい。
'_/  (*3)[ktHolidayName]という関数名そのものは、各自の環境に
'_/      おける命名規則に沿って変更しても構いません。
'_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Public Function ktHolidayName(ByVal 日付 As Date) As String
Dim dtm日付 As Date
Dim str祝日名 As String
Const cst振替休日施行日 As Date = "1973/4/12"

'時刻/時刻誤差の削除(Now関数などへの対応)
    dtm日付 = DateSerial(Year(日付), Month(日付), Day(日付))

    'シリアル値は[±0.5秒]の誤差範囲で認識されます。2002/6/21はシリアル値で
    '[37428.0]ですが､これに[-0.5秒]の誤差が入れば[37427.9999942130]となり､
    'Int関数で整数部分を取り出せば[37427]で前日日付になってしまいます。
    '※ 但し､引数に指定する値が必ず【手入力した日付】等で､時刻や時刻誤差を
    '  考慮しなくても良いならば､このステップは不要です。引数[日付]をそのまま
    '  使用しても問題ありません(ほとんどの利用形態ではこちらでしょうが‥‥)。


    str祝日名 = prv祝日(dtm日付)

    ' ----- 振替休日の判定 (振替休日施行日:1973/4/12) -----
    ' [ 対象日≠祝日/休日 ＆ 対象日＝月曜日 ]のみ、前日(＝日曜日)を祝日判定する。
    ' 前日(＝日曜日)が祝日の場合は”振替休日”となる。
    ' 尚、５月６日の扱いを
    '     「火曜 or 水曜(みどりの日(5/4) or 憲法記念日(5/3)の振替休日)」⇒５月ブロック内で判定済
    '     「月曜(こどもの日(5/5)の振替休日」⇒ここの判定処理で判定
    ' とする事により、ここでの判定対象は『対象日が月曜日』のみ となります。
    If (str祝日名 = "") Then
        If (Weekday(dtm日付) = vbMonday) Then
            If (dtm日付 >= cst振替休日施行日) Then
                str祝日名 = prv祝日(dtm日付 - 1)
                If (str祝日名 <> "") Then
                    ktHolidayName = "振替休日"
                Else
                    ktHolidayName = ""
                End If
            Else
                ktHolidayName = ""
            End If
        Else
            ktHolidayName = ""
        End If
    Else
        ktHolidayName = str祝日名
    End If
End Function

'========================================================================
Private Function prv祝日(ByVal 日付 As Date) As String
Dim int年 As Integer
Dim int月 As Integer
Dim int日 As Integer
Dim int秋分日 As Integer
Dim str第N曜日 As String
' 時刻データ(小数部)は取り除いてあるので、下記の日付との比較はＯＫ
Const cst祝日法施行 As Date = "1948/7/20"
Const cst昭和天皇の大喪の礼 As Date = "1989/2/24"
Const cst明仁親王の結婚の儀 As Date = "1959/4/10"
Const cst徳仁親王の結婚の儀 As Date = "1993/6/9"
Const cst即位礼正殿の儀 As Date = "1990/11/12"    '平成天皇

Const cst平成天皇の退位 As Date = "2019/4/30"    '祝日ではなく「国民の休日」です
Const cst徳仁親王の即位 As Date = "2019/5/1"
Const cst2019GW国民の休日 As Date = "2019/5/2"    '祝日ではなく「国民の休日」です
Const cst即位礼正殿の儀_徳仁親王 As Date = "2019/10/22"

    int年 = Year(日付)
    int月 = Month(日付)
    int日 = Day(日付)

    prv祝日 = ""
    If (日付 < cst祝日法施行) Then
        Exit Function    ' 祝日法施行以前
    End If

    Select Case int月
    '-- １月 --
    Case 1
        If (int日 = 1) Then
            prv祝日 = "元日"
        Else
            If (int年 >= 2000) Then
                str第N曜日 = (((int日 - 1) \ 7) + 1) & Weekday(日付)
                If (str第N曜日 = "22") Then  '2nd Monday(2)
                    prv祝日 = "成人の日"
                End If
            Else
                If (int日 = 15) Then
                    prv祝日 = "成人の日"
                End If
            End If
        End If

    '-- ２月 --
    Case 2
        If (int日 = 11) Then
            If (int年 >= 1967) Then
                prv祝日 = "建国記念の日"
            End If
        ElseIf (int日 = 23) Then
            If (int年 >= 2020) Then
                prv祝日 = "天皇誕生日"
            End If
        ElseIf (日付 = cst昭和天皇の大喪の礼) Then
            prv祝日 = "昭和天皇の大喪の礼"
        End If

    '-- ３月 --
    Case 3
        If (int日 = prv春分日(int年)) Then  ' 1948～2150以外は[99]
            prv祝日 = "春分の日"            ' が返るので､必ず≠になる
        End If

    '-- ４月 --
    Case 4
        If (int日 = 29) Then
            If (int年 >= 2007) Then
                prv祝日 = "昭和の日"
            ElseIf (int年 >= 1989) Then
                prv祝日 = "みどりの日"
            Else
                prv祝日 = "天皇誕生日"    ' 昭和天皇
            End If
        ElseIf (日付 = cst平成天皇の退位) Then    ' 2019/4/30
            prv祝日 = "国民の休日"    '祝日に挟まれた国民の休日です
        ElseIf (日付 = cst明仁親王の結婚の儀) Then
            prv祝日 = "皇太子明仁親王の結婚の儀"
        End If

    '-- ５月 --
    Case 5
        If (int日 = 3) Then
            prv祝日 = "憲法記念日"
        ElseIf (int日 = 4) Then
            If (int年 >= 2007) Then
                prv祝日 = "みどりの日"
            ElseIf (int年 >= 1986) Then
                ' 5/4が日曜日は『只の日曜』､月曜日は『憲法記念日の振替休日』(～2006年)
                If (Weekday(日付) > vbMonday) Then   ' 火曜 以降(火～土)
                    prv祝日 = "国民の休日"
                End If
            End If
        ElseIf (int日 = 5) Then
            prv祝日 = "こどもの日"
        ElseIf (int日 = 6) Then
            If (int年 >= 2007) Then
                Select Case Weekday(日付)
                    Case vbTuesday, vbWednesday
                        prv祝日 = "振替休日"    ' [5/3,5/4が日曜]ケースのみ、ここで判定
                End Select
            End If
        Else
            If (int年 = 2019) Then
                If (日付 = cst徳仁親王の即位) Then    ' 2019/5/1
                    prv祝日 = "即位の日"    ' 徳仁親王
                ElseIf (日付 = cst2019GW国民の休日) Then    ' 2019/5/2
                    prv祝日 = "国民の休日"    '祝日に挟まれた国民の休日です
                End If
            End If
        End If

    '-- ６月 --
    Case 6
        If (日付 = cst徳仁親王の結婚の儀) Then
            prv祝日 = "皇太子徳仁親王の結婚の儀"
        End If

    '-- ７月 --
    Case 7
        str第N曜日 = (((int日 - 1) \ 7) + 1) & Weekday(日付)
        Select Case int年
          Case Is >= 2021
            If (str第N曜日 = "32") Then  '3rd Monday(2)
                prv祝日 = "海の日"
            End If
          Case 2020
            '2020年はオリンピック特措法により
            '「海の日」が 7/23 / 「スポーツの日」が 7/24 に移動
            Select Case int日
              Case 23
                prv祝日 = "海の日"
              Case 24
                prv祝日 = "スポーツの日"
              Case Else
            End Select
          Case Is >= 2003
            If (str第N曜日 = "32") Then  '3rd Monday(2)
                prv祝日 = "海の日"
            End If
          Case Is >= 1996
            If (int日 = 20) Then
                prv祝日 = "海の日"
            End If
          Case Else
        End Select

    '-- ８月 --
    Case 8
        Select Case int年
          Case Is >= 2021
            If (int日 = 11) Then
                prv祝日 = "山の日"
            End If
          Case 2020
            '2020年はオリンピック特措法により「山の日」が 8/10 に移動
            If (int日 = 10) Then
                prv祝日 = "山の日"
            End If
          Case Is >= 2016
            If (int日 = 11) Then
                prv祝日 = "山の日"
            End If
          Case Else
        End Select

    '-- ９月 --
    Case 9
        '第３月曜日(15～21)と秋分日(22～24)が重なる事はない
        int秋分日 = prv秋分日(int年)
        If (int日 = int秋分日) Then  ' 1948～2150以外は[99]
            prv祝日 = "秋分の日"      ' が返るので､必ず≠になる
        Else
            If (int年 >= 2003) Then
                str第N曜日 = (((int日 - 1) \ 7) + 1) & Weekday(日付)
                If (str第N曜日 = "32") Then  '3rd Monday(2)
                    prv祝日 = "敬老の日"
                ElseIf (Weekday(日付) = vbTuesday) Then
                    If (int日 = (int秋分日 - 1)) Then
                        prv祝日 = "国民の休日"  '火曜日＆[秋分日の前日]
                    End If
                End If
            ElseIf (int年 >= 1966) Then
                If (int日 = 15) Then
                    prv祝日 = "敬老の日"
                End If
            End If
        End If

    '-- １０月 --
    Case 10
        str第N曜日 = (((int日 - 1) \ 7) + 1) & Weekday(日付)
        Select Case int年
          Case Is >= 2021
            If (str第N曜日 = "22") Then  '2nd Monday(2)
                prv祝日 = "スポーツの日"  '2020年より改名
            End If
          Case 2020
            '2020年はオリンピック特措法により「スポーツの日」が 7/24 に移動
          Case Is >= 2000
            If (str第N曜日 = "22") Then  '2nd Monday(2)
                prv祝日 = "体育の日"
            ElseIf (日付 = cst即位礼正殿の儀_徳仁親王) Then
                prv祝日 = "即位礼正殿の儀"    ' 徳仁親王(2019/10/22)
            End If
          Case Is >= 1966
            If (int日 = 10) Then
                prv祝日 = "体育の日"
            End If
          Case Else
        End Select

    '-- １１月 --
    Case 11
        If (int日 = 3) Then
            prv祝日 = "文化の日"
        ElseIf (int日 = 23) Then
            prv祝日 = "勤労感謝の日"
        ElseIf (日付 = cst即位礼正殿の儀) Then
            prv祝日 = "即位礼正殿の儀"    ' 平成天皇
        End If

    '-- １２月 --
    Case 12
        If (int日 = 23) Then
            If ((int年 >= 1989) And (int年 <= 2018)) Then
                prv祝日 = "天皇誕生日"    ' 平成天皇
            End If
        End If
    End Select
End Function

'======================================================================
'  春分/秋分日の略算式は
'    『海上保安庁水路部 暦計算研究会編 新こよみ便利帳』
'  で紹介されている式です。
Private Function prv春分日(ByVal 年 As Integer) As Integer
    If (年 <= 1947) Then
        prv春分日 = 99        '祝日法施行前
    ElseIf (年 <= 1979) Then
        '(年 - 1983)がマイナスになるので『Fix関数』にする
        prv春分日 = Fix(20.8357 + (0.242194 * (年 - 1980)) - Fix((年 - 1983) / 4))
    ElseIf (年 <= 2099) Then
        prv春分日 = Fix(20.8431 + (0.242194 * (年 - 1980)) - Fix((年 - 1980) / 4))
    ElseIf (年 <= 2150) Then
        prv春分日 = Fix(21.851 + (0.242194 * (年 - 1980)) - Fix((年 - 1980) / 4))
    Else
        prv春分日 = 99        '2151年以降は略算式が無いので不明
    End If
End Function

'========================================================================
Private Function prv秋分日(ByVal 年 As Integer) As Integer
    If (年 <= 1947) Then
        prv秋分日 = 99        '祝日法施行前
    ElseIf (年 <= 1979) Then
        '(年 - 1983)がマイナスになるので『Fix関数』にする
        prv秋分日 = Fix(23.2588 + (0.242194 * (年 - 1980)) - Fix((年 - 1983) / 4))
    ElseIf (年 <= 2099) Then
        prv秋分日 = Fix(23.2488 + (0.242194 * (年 - 1980)) - Fix((年 - 1980) / 4))
    ElseIf (年 <= 2150) Then
        prv秋分日 = Fix(24.2488 + (0.242194 * (年 - 1980)) - Fix((年 - 1980) / 4))
    Else
        prv秋分日 = 99        '2151年以降は略算式が無いので不明
    End If
End Function

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/  CopyRight(C) K.Tsunoda(AddinBox) 2001 All Rights Reserved.
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/