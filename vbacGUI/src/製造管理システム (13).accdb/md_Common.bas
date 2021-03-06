Option Compare Database
Option Explicit
'--------------------------------------------------------------------------------------------------------------------
'共通変数

'2.6.0
'   →bolShizaiUpdatable　追加
'3.0.0
'   →bolsekkei 追加
'--------------------------------------------------------------------------------------------------------------------
'本番データベース名
Public Const strDBName As String = "DB02"

'パスワード桁数
Public Const constintPassWordLength As Integer = 5

'ユーザーID,権限等
Public strUserID As String
Public strUserName As String
Public bolUpdatable As Boolean
Public bolAdministrator As Boolean

Public bolShizaiUpdatable As Boolean

Public bolSekkei As Boolean

'1.10.6 K.Asayama 20151211 追加
'SxLローカルコピー,カレンダーコピー
Public bolSxLCopy As Boolean
Public bolCalendarCopy As Boolean

'1.12.3 ADD
'サーバパス
Public Const conServerPath As String = "\\db\prog\製造管理システム"
Public Const conUserPath As String = "\\db\prog\製造管理システム"

Public Sub UserINIT()
'--------------------------------------------------------------------------------------------------------------------
'ユーザー関連関数初期化

''1.10.6 K.Asayama bolSxLCopy,bolCalendarCopy 初期化追加 20151211 追加
'--------------------------------------------------------------------------------------------------------------------
    strUserID = ""
    strUserName = ""
    bolUpdatable = False
    bolAdministrator = False
    
    bolSxLCopy = False
    bolCalendarCopy = False
    
End Sub

Public Function Connection_DB() As String
'--------------------------------------------------------------------------------------------------------------------
'現在使用しているDB名にセットされているODBCの接続先DB名を返す
'戻り値 : SQLSERVERのDB名 (String形式）

'注意：32bit版のACCESS用なので64bit版に移行した場合レジストリのディレクトリをあわせる必要あり
'--------------------------------------------------------------------------------------------------------------------

    Dim ConnectDB As String
    
    Dim objWshShell As Object
    Dim OSBit As Byte
    Dim strConnDB As String
    
    ConnectDB = strDBName & "_BRAND_MASTER" 'strDBNameはPUBLIC変数
    
    OSBit = OS_Architecture()
    strConnDB = ""
    
    Set objWshShell = CreateObject("WScript.Shell")
    
    If OSBit = 64 Then
        strConnDB = objWshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\ODBC\ODBC.INI\" & ConnectDB & "\Server")
    Else
        strConnDB = objWshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI\" & ConnectDB & "\Server")
    End If
    
    Connection_DB = strConnDB
    
    Set objWshShell = Nothing
    
End Function

Public Function OS_Architecture() As Byte
'--------------------------------------------------------------------------------------------------------------------
'OSが32bit版か64bit版かをチェックして数字で返す
'戻り値 : 32 または 64 (byte形式）
'--------------------------------------------------------------------------------------------------------------------

    Dim objWshShell As Object
    Dim strMode As String

    Set objWshShell = CreateObject("WScript.Shell")
    
    strMode = objWshShell.Environment("Process").Item("PROCESSOR_ARCHITECTURE")

    If UCase(strMode) = "X86" Then
         OS_Architecture = 32
    Else
         OS_Architecture = 64
    End If
    
    Set objWshShell = Nothing
    
End Function

Public Function to_Capital(intKeyASCII As Integer) As Integer
'--------------------------------------------------------------------------------------------------------------------
'小文字→大文字変換
'小文字の場合は大文字に変換して返す。それ以外はそのまま返す。
'--------------------------------------------------------------------------------------------------------------------
        Select Case intKeyASCII
            'Case 48 To 57 '0〜9
            'Case 65 To 90 'A〜Z
            Case 97 To 122 'a〜z
                to_Capital = intKeyASCII - 32 '大文字に強制変換
            Case Else
                to_Capital = intKeyASCII
        End Select
End Function

Public Function RPAD(strValue As String, strCharactor As String, intKeta As Integer) As String
'--------------------------------------------------------------------------------------------------------------------
'文字埋め
'stringの右側に指定された文字を文字数分になるよう埋める
'--------------------------------------------------------------------------------------------------------------------
        RPAD = strValue & String(intKeta - Len(strValue), strCharactor)
        
End Function

Public Function IsNumber(intKeyASCII As Integer, Optional bolBackSpace As Variant) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'   数字確認
'
'   戻り値:Boolean
'       →True              数字
'       →False             数字以外
'
'    Input項目
'       intKeyASCII         キーボード入力値（ASCII値)
'       bolBackSpace        Trueの場合はBackSpaceキー(Keyascii=8)を数字として認識する

'--------------------------------------------------------------------------------------------------------------------

        IsNumber = False
        
        Select Case intKeyASCII
            Case 48 To 57 '0〜9
                IsNumber = True
            Case Else
                If Not IsMissing(bolBackSpace) Then
                    If bolBackSpace And intKeyASCII = 8 Then
                       IsNumber = True
                    End If
                End If
        End Select
End Function

Public Function LPAD(strValue As String, strCharactor As String, intKeta As Integer) As String
'--------------------------------------------------------------------------------------------------------------------
'stringの左側に指定された文字を文字数分になるよう埋める
'--------------------------------------------------------------------------------------------------------------------
    
        LPAD = String(intKeta - Len(strValue), strCharactor) & strValue
        
End Function

Public Function IsKeiyakuNo(in_Text As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'入力された文字列が契約番号か確認する(*-*-*型式）
'   戻り値:Boolean
'       →True              契約番号
'       →False             契約番号以外
'
'    Input項目
'       in_Text             入力値
'--------------------------------------------------------------------------------------------------------------------

    If in_Text Like "SR*" Then
        If in_Text Like "SR####-###-####" Or in_Text Like "SR#####-###-####" Then
            IsKeiyakuNo = True '住友
        End If
    Else
        If in_Text Like "??####-###-####" Then
            IsKeiyakuNo = True 'ブランド
        End If
    End If
End Function

Public Function bytfncCheckDigit_JAN(in_strCode As String) As Byte
'--------------------------------------------------------------------------------------------------------------------
'チェックディジット計算（JANコード専用）
'計算方法は[モジュラス10/ウェイト3]
'   戻り値:Byte
'                           →チェックディジット
'                             （エラーの時は99を返す）
'
'    Input項目
'       in_strCode            JANコード（12桁）
'--------------------------------------------------------------------------------------------------------------------
    Dim bytCode(11) As Byte
    Dim i As Byte
    Dim strDigit As String
    
    bytfncCheckDigit_JAN = 99
    
    On Error GoTo Err_bytfncCheckDigit_JAN
    
    
    If Not IsNumeric(in_strCode) Then Exit Function
    If Len(in_strCode) <> 12 Then Exit Function
    
    For i = 0 To 11
        bytCode(i) = Mid(in_strCode, i + 1, 1)
    Next
    
    strDigit = right(CStr(((bytCode(11) + bytCode(9) + bytCode(7) + bytCode(5) + bytCode(3) + bytCode(1)) * 3) + (bytCode(10) + bytCode(8) + bytCode(6) + bytCode(4) + bytCode(2) + bytCode(0))), 1)
    
    If strDigit = "0" Then
        bytfncCheckDigit_JAN = 0
    Else
        bytfncCheckDigit_JAN = 10 - CByte(strDigit)
    End If
    
    Exit Function
    
Err_bytfncCheckDigit_JAN:
    MsgBox Err.Description
    bytfncCheckDigit_JAN = 99

End Function

Public Function strfncGetVersion() As String
'--------------------------------------------------------------------------------------------------------------------
'バージョン取得処理
'   →バージョンを取得し文字列で戻す
'   →取得できない場合は空文字で返す

'3.0.0
'   →Rev追加
'--------------------------------------------------------------------------------------------------------------------
    Dim objLOCALdb As New cls_LOCALDB
    Dim strRev As String
    
    On Error GoTo Err_strfncGetVersion
    
    If objLOCALdb.ExecSelect("select Version,Rev from T_Version管理 order by 更新日時 desc ") Then
        
        If Not objLOCALdb.GetRS.EOF Then
            strfncGetVersion = objLOCALdb.GetRS![Version]
            strRev = CStr(objLOCALdb.GetRS![Rev])
        Else
            Err.Raise 9999, , "バージョン取得エラー。メニューを開くことができません"
        End If
        
    Else
    
        Err.Raise 9999, , "バージョン取得エラー。メニューを開くことができません"
    
    End If
    
    strfncGetVersion = strfncGetVersion & "." & strRev
    
    GoTo Exit_strfncGetVersion

Err_strfncGetVersion:
    strfncGetVersion = ""
    MsgBox Err.Description
    
Exit_strfncGetVersion:
'クラスのインスタンスを破棄
    Set objLOCALdb = Nothing
    
End Function

Public Function strfncGetVersionOnly() As String
'--------------------------------------------------------------------------------------------------------------------
'バージョン(レビジョン無し)取得処理
'   →バージョンを取得し文字列で戻す
'   →取得できない場合は空文字で返す

'3.0.0 ADD
'--------------------------------------------------------------------------------------------------------------------
    Dim objLOCALdb As New cls_LOCALDB
    
    On Error GoTo Err_strfncGetVersionOnly
    
    If objLOCALdb.ExecSelect("select Version from T_Version管理 order by 更新日時 desc ") Then
        
        If Not objLOCALdb.GetRS.EOF Then
            strfncGetVersionOnly = objLOCALdb.GetRS![Version]
        Else
            Err.Raise 9999, , "バージョン取得エラー。メニューを開くことができません"
        End If
        
    Else
    
        Err.Raise 9999, , "バージョン取得エラー。メニューを開くことができません"
    
    End If
    
    GoTo Exit_strfncGetVersionOnly

Err_strfncGetVersionOnly:
    strfncGetVersionOnly = ""
    MsgBox Err.Description
    
Exit_strfncGetVersionOnly:
'クラスのインスタンスを破棄
    Set objLOCALdb = Nothing
    
End Function

Public Function bolfncOpen_LogOnMenu(in_strMenuName As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'ログオン済み確認処理
'   →ログオン画面を表示する
'   →UserIDが空欄の場合はキャンセルしたことになる
'   :引数
'       in_strMenuName  :メニュー名

'   :戻り値
'       True            :ログオン済み
'       False           :未ログオン
'--------------------------------------------------------------------------------------------------------------------
    bolfncOpen_LogOnMenu = False
    
    On Error GoTo Err_bolfncOpen_LogOnMenu
    
    DoCmd.OpenForm "F_ログオン", acNormal, , , , acDialog, in_strMenuName
    
    If strUserID <> "" Then
        bolfncOpen_LogOnMenu = True
    End If
    
    GoTo Exit_bolfncOpen_LogOnMenu
    
Err_bolfncOpen_LogOnMenu:
    MsgBox Err.Description
    bolfncOpen_LogOnMenu = False
Exit_bolfncOpen_LogOnMenu:

End Function

Public Function bolfncEnableSystem() As Boolean
'--------------------------------------------------------------------------------------------------------------------
'日時処理終了確認
'   →夜間バッチ処理が終了しているか確認する
'
'   :戻り値
'       True            :システム使用可能
'       False           :日時処理未終了
'--------------------------------------------------------------------------------------------------------------------
    Dim objREMOTEdb As New cls_BRAND_MASTER
    
    bolfncEnableSystem = False
    
    On Error GoTo Err_bolfncEnableSystem
    
    If objREMOTEdb.ExecSelect("select 値 from T_Control where [key] = 4") Then
        
        If Not objREMOTEdb.GetRS.EOF Then
            If objREMOTEdb.GetRS![値] <> Format(Date, "yyyymmdd") Then
                Err.Raise 9999, , "AM0:00〜日時処理終了までシステムは使用できません"
            End If
        Else
            Err.Raise 9999, , "コントロールマスタにキー[4]（処理日）が存在しません"
        End If
    Else
        Err.Raise 9999, , "コントロールマスタの取得エラー"
    
    End If

    bolfncEnableSystem = True
    
    GoTo Exit_bolfncEnableSystem
    
Err_bolfncEnableSystem:
    MsgBox Err.Description
    bolfncEnableSystem = False
    
Exit_bolfncEnableSystem:

    Set objREMOTEdb = Nothing
End Function

Public Function bolfncTextCompare(in_BeforeUpdate As Variant, in_AfterUpdate As Variant, Optional in_CompareMode As Variant) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'テキスト比較処理
'   →2つのテキストを比較し同じであればTrueを返す
'
'   :引数
'       in_BeforeUpdate     :変更前
'       in_AfterUpdate      :変更前
'       in_CompareMode      :比較モード
'                            0:バイナリモード（全角/半角、大文字/小文字、ひらがな/カタカナを区別する）（Default)
'                            1:テキストモード（全角/半角、大文字/小文字、ひらがな/カタカナを区別しない）
'                            2:Accessの設定にしたがう
'
'   :戻り値
'       True            :変更あり
'       False           :変更なし
'--------------------------------------------------------------------------------------------------------------------
    Dim intComparemode As Byte
    Dim strBefore As String, strAfter As String
    
    On Error GoTo Err_bolfncTextCompare
    
    bolfncTextCompare = False
    
    If IsMissing(in_CompareMode) Then
        intComparemode = vbBinaryCompare
    Else
        intComparemode = in_CompareMode
    End If
    
    strBefore = Nz(in_BeforeUpdate, "")
    strAfter = Nz(in_AfterUpdate, "")
    
    If StrComp(strBefore, strAfter, intComparemode) Then
        '変更あり
        bolfncTextCompare = True
    End If
        
    Exit Function
    
Err_bolfncTextCompare:
    MsgBox Err.Description, , "テキスト比較エラー"
    
End Function

Public Sub WindowSize_Restore()
'--------------------------------------------------------------------------------------------------------------------
'アプリケーションウィンドウサイズを標準に戻す
'   Ver 1.01.1 K.Asayama ADD 20150910
'--------------------------------------------------------------------------------------------------------------------

    Dim lngRt As Long
    lngRt = ShowWindow(Application.hWndAccessApp, 1)
    
End Sub

Public Sub WindowSize_Minimize()
'--------------------------------------------------------------------------------------------------------------------
'アプリケーションウィンドウサイズを最小化する
'   Ver 1.01.1 K.Asayama ADD 20150910
'--------------------------------------------------------------------------------------------------------------------
    Dim lngRt As Long
    lngRt = ShowWindow(Application.hWndAccessApp, 2)
    
End Sub

Public Function fncMenuCall(ByVal strMenuName As String)
'--------------------------------------------------------------------------------------------------------------------
'メニューを開く
'   Ver 1.01.1 K.Asayama ADD 20150910
'   Ver 1.01.* K.Asayama Change 201510**
'--------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo Err_fncMenuCall
''Window最小化
    WindowSize_Minimize

'メニューコール
   'DoCmd.OpenForm strMenuName, acNormal, , , , acDialog
   DoCmd.OpenForm strMenuName, acNormal, , , , acWindowNormal
   
'Windowを元に戻す
    'WindowSize_Restore
    
    Exit Function
    
Err_fncMenuCall:
    Select Case Err.Number
        Case 2501 'キャンセル終了の場合
        
        Case Else
            MsgBox Err.Number
    End Select
    'Windowを元に戻す
    WindowSize_Restore
End Function

Public Function Form_IsLoaded(ByVal in_FormName As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'フォームが開いているか確認する
'   Ver 1.01.* K.Asayama ADD 201510**
'--------------------------------------------------------------------------------------------------------------------
    If CurrentProject.AllForms(in_FormName).IsLoaded Then
    
        Form_IsLoaded = True
    
    Else
    
        Form_IsLoaded = False
            
    End If

End Function

Public Function MainMenu_ReVisible()
'--------------------------------------------------------------------------------------------------------------------
'メインメニューを再表示する
'   Ver 1.01.* K.Asayama ADD 201510**
'--------------------------------------------------------------------------------------------------------------------
    If Form_IsLoaded("F_工程管理メニュー") Then
    
        Form_F_工程管理メニュー.Visible = True
    
    End If

End Function

Public Function TwipsToCm(ByVal value As Double) As Double
'--------------------------------------------------------------------------------------------------------------------
'   1 cm = 567 twips
'   1インチ = 1440 twips = 2.54cm = 72 ポイント
'
'   twips から cm に変換
'
'   :引数
'       value               :twipsでの長さ
'
'   :戻り値
'       Cm                  :センチメートル
'--------------------------------------------------------------------------------------------------------------------

    TwipsToCm = value / 567

End Function

Public Function CmToTwips(ByVal value As Double) As Double
'--------------------------------------------------------------------------------------------------------------------
'
'   cm から twips に変換
'
'   :引数
'       value               :cmでの長さ
'
'   :戻り値
'       twips               :Twips
'--------------------------------------------------------------------------------------------------------------------
    CmToTwips = value * 567
    
End Function

Public Sub subAllbutton_Enabled(ByVal in_FormName As String, ByVal in_bolTF As Boolean)
'--------------------------------------------------------------------------------------------------------------------
'
'   フォームのボタンの使用可能、不能一括変更
'
'   :引数
'       in_FormName         :フォーム名
'       in_bolTF            :使用可能（True）/不能（False）
'
'--------------------------------------------------------------------------------------------------------------------
    Dim ctl As Access.Control
    Dim i As Byte
    i = 0
    
    On Error GoTo Err_subAllbutton_Enabled
    
    'このフォーム内のすべてのコントロールを検索
    For Each ctl In Forms(in_FormName).Controls
        With ctl
            If .ControlType = acCommandButton Then
                   
                ctl.Enabled = in_bolTF

            End If
        End With
    Next ctl
        
    GoTo Exit_subAllbutton_Enabled
    
Err_subAllbutton_Enabled:

Exit_subAllbutton_Enabled:
    Set ctl = Nothing
End Sub

Public Function bolfncinputDate(ByVal in_MidashiText As String, ByRef out_Date As Variant) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'
'   汎用日付入力フォーム表示
'
'   :引数
'       in_MidashiText      :見出し名（8文字くらいが適当）
'       out_Date            :入力日付
'
'   :戻り値
'                           :日付入力済み（True）/キャンセル（False）
'--------------------------------------------------------------------------------------------------------------------
Dim objLOCALdb As New cls_LOCALDB
Dim strErrMsg As String

On Error GoTo Err_bolfncinputDate

out_Date = Null

If Not objLOCALdb.ExecSQL("delete from WK_対象日付", strErrMsg) Then
    Err.Raise 9999, , strErrMsg
End If

DoCmd.OpenForm "F_汎用日付入力", acNormal, , , , acDialog, in_MidashiText

If Not objLOCALdb.ExecSelect("select date1 from WK_対象日付") Then
    Err.Raise 9999, , "日付読み込みエラー"
Else
    If Not objLOCALdb.GetRS.EOF Then
        out_Date = objLOCALdb.GetRS!Date1
    End If
End If

If IsNull(out_Date) Then
    Err.Raise 9998, , "日付の入力がキャンセルされました"
End If

bolfncinputDate = True

GoTo Exit_bolfncinputDate

Err_bolfncinputDate:
    bolfncinputDate = False
    MsgBox Err.Description
    
Exit_bolfncinputDate:
    Set objLOCALdb = Nothing
    
End Function

Public Sub subAllbutton_noPrintable(ByVal in_FormName As String)
'--------------------------------------------------------------------------------------------------------------------
'
'   フォームのボタンの印刷不可
'
'   :引数
'       in_FormName         :フォーム名
'
'1.10.9 K.Asayama ADD
'--------------------------------------------------------------------------------------------------------------------
    Dim ctl As Access.Control
    Dim i As Byte
    i = 0
    
    On Error GoTo Err_subAllbutton_noPrintable
    
    'このフォーム内のすべてのコントロールを検索
    For Each ctl In Forms(in_FormName).Controls
        With ctl
            If .ControlType = acCommandButton Then
                   
                ctl.DisplayWhen = 2

            End If
        End With
    Next ctl
        
    GoTo Exit_subAllbutton_noPrintable
    
Err_subAllbutton_noPrintable:

Exit_subAllbutton_noPrintable:
    Set ctl = Nothing
End Sub

Public Sub subScreenShot_AllArea()
'--------------------------------------------------------------------------------------------------------------------
'
'   全スクリーンのスクリーンショット取得
'
'
'1.10.14 K.Asayama ADD
'--------------------------------------------------------------------------------------------------------------------
    keybd_event vbKeySnapshot, 0&, &H1, 0&
    keybd_event vbKeySnapshot, 0&, &H1 Or &H2, 0&
    
End Sub

Public Sub subScreenShot_ActiveArea()
'--------------------------------------------------------------------------------------------------------------------
'
'   アクティブスクリーンのスクリーンショット取得
'
'
'1.10.14 K.Asayama ADD
'--------------------------------------------------------------------------------------------------------------------
    keybd_event &HA4, 0&, &H1, 0&
    keybd_event vbKeySnapshot, 0&, &H1, 0&
    keybd_event vbKeySnapshot, 0&, &H1 Or &H2, 0&
    keybd_event &HA4, 0&, &H1 Or &H2, 0&
End Sub

Public Function bolfncinputDate_FromTo(ByVal in_MidashiText As String, ByVal in_DateDetail As String, ByRef out_DateFrom As Variant, ByRef out_DateTo As Variant) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'
'   汎用日付入力フォーム(FromTo)表示
'
'   :引数
'       in_MidashiText      :見出し名（8文字くらいが適当）
'       in_DateDetail       :日付詳細（製造日、納品日等画面表示用）
'       out_DateFrom        :入力日付(From)
'       out_DateTo          :入力日付(To)
'
'   :戻り値
'                           :日付入力済み（True）/キャンセル（False）
'1.10.15 ADD
'--------------------------------------------------------------------------------------------------------------------
Dim objLOCALdb As New cls_LOCALDB
Dim strErrMsg As String

On Error GoTo Err_bolfncinputDate_FromTo

out_DateFrom = Null
out_DateTo = Null

If Not objLOCALdb.ExecSQL("delete from WK_対象日付", strErrMsg) Then
    Err.Raise 9999, , strErrMsg
End If

DoCmd.OpenForm "F_汎用日付入力_FromTo", acNormal, , , , acDialog, in_MidashiText & vbTab & in_DateDetail

If Not objLOCALdb.ExecSelect("select date1,date2 from WK_対象日付") Then
    Err.Raise 9999, , "日付読み込みエラー"
Else
    If Not objLOCALdb.GetRS.EOF Then
        out_DateFrom = objLOCALdb.GetRS!Date1
        out_DateTo = objLOCALdb.GetRS!Date2
    End If
End If

If IsNull(out_DateFrom) Or IsNull(out_DateTo) Then
    Err.Raise 9998, , "日付の入力がキャンセルされました"
End If

bolfncinputDate_FromTo = True

GoTo Exit_bolfncinputDate_FromTo

Err_bolfncinputDate_FromTo:
    bolfncinputDate_FromTo = False
    MsgBox Err.Description
    
Exit_bolfncinputDate_FromTo:
    Set objLOCALdb = Nothing
    
End Function

Public Function bolfncReport(in_ReportName As String, in_Preview As Boolean, Optional in_Message As Boolean) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'
'   レポート出力
'
'   :引数
'       in_ReportName       :レポート名
'       in_Preview          :True→プレビュー False→プリンタ出力
'       in_Message(Option)  :True→データ0件の際メッセージを出力する   False→しない
'
'   :戻り値
'       True            :成功
'       False           :失敗
'
'   1.11.0 ADD
'--------------------------------------------------------------------------------------------------------------------
    Dim bytPrintmode As Byte

    On Error GoTo Err_bolfncReport

    If in_Preview Then
        bytPrintmode = 2
    Else
        bytPrintmode = 0
    End If
    
    DoCmd.OpenReport in_ReportName, bytPrintmode
    
    bolfncReport = True
    
    Exit Function

Err_bolfncReport:
    
    If Err.Number = 2501 Then
        If in_Message Then
            MsgBox in_ReportName & vbCrLf & "データがありません"
        End If
        Resume Next
    Else
        MsgBox Err.Description
    End If
    
    bolfncReport = False
    

End Function

Public Function strfncTextFileToString(strFileFullpath As String) As String
'--------------------------------------------------------------------------------------------------------------------
'TextからStringへフルコピー
'   →ファイル（フルパス）を読み込んでそのままString変数にインポート
'
'   :引数
'       strFileFullpath     ファイル名（フルパス）

'1.11.1 ADD
'--------------------------------------------------------------------------------------------------------------------
    Dim strTxt As String
    
    strfncTextFileToString = ""
    strTxt = ""
    
    On Error GoTo Err_strfncTextFileToString
    
    If Dir(strFileFullpath) <> "" Then
        With CreateObject("Scripting.FileSystemObject")
            With .GetFile(strFileFullpath).OpenAsTextStream
                strTxt = .ReadAll
                .Close
            End With
        End With
        
    Else
        Err.Raise 9999, , "変換用ファイルが存在しません。管理者に連絡してください"
    End If
    
    strfncTextFileToString = strTxt
    
    Exit Function
    
Err_strfncTextFileToString:
    Close
    MsgBox Err.Description
    
End Function

Public Sub subAllOption_Enabled(ByVal in_FormName As String, ByVal in_bolTF As Boolean)
'--------------------------------------------------------------------------------------------------------------------
'
'   フォームのOptionコントロールの使用可能、不能一括変更
'
'   :引数
'       in_FormName         :フォーム名
'       in_bolTF            :使用可能（True）/不能（False）
'
'1.11.1 ADD
'1.12.0
'   →  コンボボックスとチェックボックス追加
'--------------------------------------------------------------------------------------------------------------------
    Dim ctl As Access.Control
    Dim i As Byte
    i = 0
    
    On Error GoTo Err_subAllOption_Enabled_Enabled
    
    'このフォーム内のすべてのコントロールを検索
    For Each ctl In Forms(in_FormName).Controls
        With ctl
            If .ControlType = acOptionGroup Or .ControlType = acComboBox Or .ControlType = acCheckBox Then
                   
                ctl.Enabled = in_bolTF

            End If
        End With
    Next ctl
        
    GoTo Exit_subAllOption_Enabled_Enabled
    
Err_subAllOption_Enabled_Enabled:

Exit_subAllOption_Enabled_Enabled:
    Set ctl = Nothing
End Sub

Public Sub subAllTextBox_Enabled(ByVal in_FormName As String, ByVal in_bolTF As Boolean)
'--------------------------------------------------------------------------------------------------------------------
'
'   フォームのTextBoxの使用可能、不能一括変更
'
'   :引数
'       in_FormName         :フォーム名
'       in_bolTF            :使用可能（True）/不能（False）
'
'2.0.0 ADD
'--------------------------------------------------------------------------------------------------------------------
    Dim ctl As Access.Control
    Dim i As Byte
    i = 0
    
    On Error GoTo Err_subAllOption_Enabled_Enabled
    
    'このフォーム内のすべてのコントロールを検索
    For Each ctl In Forms(in_FormName).Controls
        With ctl
            If .ControlType = acTextBox Then
                   
                ctl.Enabled = in_bolTF

            End If
        End With
    Next ctl
        
    GoTo Exit_subAllOption_Enabled_Enabled
    
Err_subAllOption_Enabled_Enabled:

Exit_subAllOption_Enabled_Enabled:
    Set ctl = Nothing
End Sub
Public Function Report_IsLoaded(ByVal in_ReportName As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'レポートが開いているか確認する
'   Ver 1.11.2 ADD
'--------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo Err_Report_IsLoaded
    
    If CurrentProject.AllReports(in_ReportName).IsLoaded Then
    
        Report_IsLoaded = True
    
    Else
    
        Report_IsLoaded = False
            
    End If
    
    Exit Function
    
Err_Report_IsLoaded:
'    If Err.Number = 2467 Then
'        Resume Next
'    End If
    Report_IsLoaded = False
    
End Function

Public Function bolfncinFlieGet(ByVal in_KeyName As String, ByRef out_iniData As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
'iniファイルからを指定のキーを抽出

'   :引数
'       in_KeyName             :iniファイルキー名
'       out_iniData            :変数名
'
'   :戻り値
'       True            :成功
'       False           :失敗

'   Ver 1.11.2 ADD
'   1.11.3  Change テスト環境識別追加（ローカル(C:\kamiya_Brandにiniファイルがある場合はそちらを優先する
'   1.12.3
'           →サーバパス共通化
'2.13.0
'   →ローカル（テスト用）を使用する場合はメッセージを表示する
'--------------------------------------------------------------------------------------------------------------------
       
    Const strIniPath As String = conServerPath & "\製造管理システム.ini"
    
    Const strTestPath As String = "C:\Kamiya_Brand\製造管理システム.ini"
    
    Dim strBuf As String
    Dim varText As Variant
    Dim varPath As Variant
    
    Dim strInputPath As String
    
    Dim i As Integer
    
    bolfncinFlieGet = False
    
    On Error GoTo Err_bolfncinFlieGet
    
    varPath = Null
    
    'iniファイルがローカルにある場合はそちらを優先
    If Dir(strTestPath) <> "" Then
        MsgBox "テスト用のIniファイルを使用しています", vbInformation
        strInputPath = strTestPath
    Else
        strInputPath = strIniPath
    End If
    
    'iniファイルをバッファに読み込み
    With CreateObject("Scripting.FileSystemObject")
        With .GetFile(strInputPath).OpenAsTextStream
            strBuf = .ReadAll
            .Close
        End With
    End With
    
    varText = Split(strBuf, vbCrLf)
    
    If VarType(varText) > vbArray Then
        For i = LBound(varText) To UBound(varText)
            If varText(i) Like in_KeyName & vbTab & "*" Then
                varPath = Split(varText(i), vbTab)
                Exit For
            End If
        Next
    Else
        'Debug.Print varText
    End If
        
    If VarType(varPath) > vbArray Then
       out_iniData = varPath(1)
       bolfncinFlieGet = True
    End If
    
    GoTo Exit_bolfncinFlieGet
    
Err_bolfncinFlieGet:
    MsgBox Err.Description
    Close
Exit_bolfncinFlieGet:

End Function

Public Sub OpenExplorer(in_Path As String)
'--------------------------------------------------------------------------------------------------------------------
'指定したファイル位置をエクスプローラで開く

'   :引数
'       in_Path             :フルパス名

'1.12.0 ADD
'--------------------------------------------------------------------------------------------------------------------
    Call Shell("explorer.exe /select," & in_Path, vbNormalFocus)
End Sub

Public Function NullIFNothing(ByVal InputData As Variant) As Variant
'***********************************************************
'0または空欄（Empty)の時Nullに置き換える
'
'   値が0又は空欄の時Nullに置き換える
'       主にDBへの値登録やテキストボックス等のオブジェクトへ
'       値を戻すときに使用
'
'   戻り値（Variant型)
'
'2.1.0 ADD
'***********************************************************
    Select Case VarType(InputData)
    
        Case Is <= 1 '空欄 又は Null
            
            NullIFNothing = Null
            
        Case Is <= 6 '数値系
            
            If InputData = 0 Then
                NullIFNothing = Null
            Else
                NullIFNothing = InputData
            End If
        
        Case 8  'String型
            
            If InputData = "" Then
                NullIFNothing = Null
            ElseIf IsNumeric(InputData) Then
                If CDbl(InputData) = 0 Then
                    NullIFNothing = Null
                Else
                    NullIFNothing = InputData
                End If
            Else
                NullIFNothing = InputData
            End If
            
        Case Else
            
            NullIFNothing = InputData
    
    End Select
    
End Function

Public Function RPADB(strValue As String, strCharactor As String, intKeta As Integer) As String
'***********************************************************
'RPAD（引数の桁数へ文字を埋める）のバイト数版
'
'   引数
'       :strValue               元になる値
'       :strCharactor           埋める文字（0の場合右側を0で埋める）
'       :intKeta                埋めた後の桁数
'
'   戻り値（String型)
'
'2.1.0 ADD
'***********************************************************
    'stringの右側に指定された文字を文字数分になるよう埋める（バイト数）
    Dim strSJIS As String
    
    'UnicodeをShift-JISに変換
    strSJIS = StrConv(strValue, vbFromUnicode)
    RPADB = strValue & String(intKeta - LenB(strSJIS), strCharactor)
        
End Function

Public Function fncFileSelector(ByVal inFolder As String, ByVal inShurui As Integer) As String
'***********************************************************
'ダイアログを表示してファイルを選択する

'
'   引数
'       :inFolder               初期パス
'       :inShurui               ファイル種類 0:xlsx

'   戻り値（String型)
'
'2.8.0 ADD
'***********************************************************

    Dim strFile As String
    Dim intRet As Integer
    
    On Error GoTo Err_fncFileSelector
    
    fncFileSelector = ""
    With Application.FileDialog(msoFileDialogOpen)
    
        'ダイアログのタイトルを設定
        '.Title = "ダイアログ"
        
        'ファイルの種類を設定
        .Filters.Clear
        
        Select Case inShurui
            Case 0
                .Filters.Add "Microsoft Office Excelファイル", "*.xlsx"
        End Select
        
        .FilterIndex = 1
        
        '複数ファイル選択を許可しない
        .AllowMultiSelect = False
        '初期パスを設定
        
        .InitialFileName = inFolder
        'ダイアログを表示
        intRet = .Show
        
        If intRet <> 0 Then
          'ファイルが選択されたとき
          'そのフルパスを返り値に設定
          strFile = Trim(.SelectedItems.Item(1))
        Else
          'ファイルが選択されなければブランク
          strFile = ""
        End If
        
    End With

    If strFile <> "" Then
        fncFileSelector = strFile
    End If
    
    Exit Function
    
Err_fncFileSelector:
    MsgBox Err.Description
    
    
End Function

Public Function isTableExist(ByVal strTableName As String) As Boolean
'   *************************************************************
'   isTableExist
'   テーブル存在確認

'
'   戻り値:Boolean
'       →True              テーブル有
'       →False             テーブル無し
'
'    Input項目
'       strTableName        テーブル名

'3.0.0 ADD
'   *************************************************************

    Dim daoDB As DAO.Database
    Dim daoTableDef As DAO.TableDef
    Set daoDB = CurrentDb
    
    isTableExist = False
    
    On Error GoTo Err_isTableExist
    
    For Each daoTableDef In CurrentDb.TableDefs
        If daoTableDef.Name = strTableName Then
            isTableExist = True
            Exit For
        End If
    Next
    
    GoTo Exit_isTableExist
    
Err_isTableExist:

Exit_isTableExist:
    Set daoTableDef = Nothing
    Set daoDB = Nothing
End Function

Public Function UrlEncodeUtf8(ByVal strSource As String) As String
'   *************************************************************
'   UrlEncodeUtf8
'   文字列をUTF8エンコードして戻す

'
'   戻り値:String
'       →エンコード後の文字列
'
'    Input項目
'       strSource        エンコード前の文字列

'3.0.0 ADD
'   *************************************************************
    Dim objSC As Object
    Set objSC = CreateObject("ScriptControl")
    objSC.Language = "Jscript"
    UrlEncodeUtf8 = objSC.CodeObject.encodeURIComponent(strSource)
    Set objSC = Nothing
End Function

Public Function bolint40mmOrder() As Integer
'--------------------------------------------------------------------------------------------------------------------
'40mmソート順
'   →ソート順で40mmを優先するか36mmを優先するか確認
'
'   :戻り値
'       0       :40mm,36mmの順
'       1       :36mm,40mmの順
'--------------------------------------------------------------------------------------------------------------------
    Dim objREMOTEdb As New cls_BRAND_MASTER
    
    bolint40mmOrder = 1
    
    On Error GoTo Err_bolint40mmOrder
    
    If objREMOTEdb.ExecSelect("select 値 from T_Control where [key] = 14") Then
        
        If Not objREMOTEdb.GetRS.EOF Then
            If IsNull(objREMOTEdb.GetRS![値]) Then
                Err.Raise 9999, , "40mmソート順が取得できないので36mm→40mmの順でソートします"
                If Not IsNumeric(objREMOTEdb.GetRS![値]) Then
                    Err.Raise 9999, , "コントロールマスタのソート順の値が異常です"
                End If
            End If
        Else
            Err.Raise 9999, , "コントロールマスタにキー[14]（40mmソート順）が存在しません"
        End If
    Else
        Err.Raise 9999, , "コントロールマスタの取得エラー"
    
    End If

    bolint40mmOrder = objREMOTEdb.GetRS![値]
    
    GoTo Exit_bolint40mmOrder
    
Err_bolint40mmOrder:
    MsgBox Err.Description
    
Exit_bolint40mmOrder:

    Set objREMOTEdb = Nothing
End Function

Public Function fncUserGroup_Belongs(in_strGroup As String) As Boolean
'   *************************************************************
'   ユーザーグループ所属確認
'       ログインユーザーのADアカウントが
'       引数のグループに登録されていればTrueを返す

'       ※ドメインに参加していない場合、PCがネットワークに接続されていない場合
'       　実行時エラーになる

'   戻り値:Boolean
'       →True              グループに所属している
'       →False             グループに所属していない
'
'   引数
'       in_strGroup        ActiveDirectoryユーザーグループ名

'  3.0.0 K.Asayama ADD
'   *************************************************************

    Dim objADSysInfo As Object
    Dim objUser As Object
    Dim objGroup As Object
    Dim varGroup As Variant
    Dim strUser As String
    
    
    Set objADSysInfo = CreateObject("ADSystemInfo")
    strUser = objADSysInfo.UserName
    Set objUser = GetObject("LDAP://" & strUser)
    
    On Error GoTo Err_fncUserGroup_Belongs
        
    strUser = objADSysInfo.UserName

    'memberOf属性は複数ある場合のみオブジェクト配列になるため
    '以下で配列か確認して処理を分岐

    If IsArray(objUser.memberOf) Then
    
        For Each varGroup In objUser.memberOf
            Set objGroup = GetObject("LDAP://" & varGroup)
            If objGroup.cn = in_strGroup Then
                fncUserGroup_Belongs = True
                Exit For
            End If
        Next
        
        Set objGroup = Nothing

    Else
        '↓グループに１つも属してないとGetObjectでエラーになるためトラップ
        On Error Resume Next
        
        Set objGroup = GetObject("LDAP://" & objUser.memberOf)
        
        If Err.Number = 0 Then
            
            If objGroup.cn = in_strGroup Then
                fncUserGroup_Belongs = True
            End If
             
            Set objGroup = Nothing
        End If

    End If
    
    GoTo Exit_fncUserGroup_Belongs
    
Err_fncUserGroup_Belongs:
    'エラーの場合はFalseを返す
    fncUserGroup_Belongs = False

Exit_fncUserGroup_Belongs:
    Set objADSysInfo = Nothing
    Set objUser = Nothing
    
End Function