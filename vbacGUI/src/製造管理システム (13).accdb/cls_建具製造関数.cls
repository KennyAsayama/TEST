Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Ver2.10

'*************************************
'接続先データベース定数              *
' "ACCESS"又は"SQLSERVER"を指定      *
'Private Const ConnectionDB As String = "ACCESS"
Private Const ConnectionDB As String = "SQLSERVER"
'*************************************

Private cnADO As ADODB.Connection
Private rsADO As ADODB.Recordset
Private rsADOAW As ADODB.Recordset
Private rsADOSF2 As ADODB.Recordset
Private rsADOSF4 As ADODB.Recordset
Private rsADOSTREET As ADODB.Recordset
Private rsADOKihonzu As ADODB.Recordset
Private rsADODaboPitch As ADODB.Recordset

Private blnKansuMatch As Boolean
Private blnZumenKansuMatch As Boolean
Private blnAWKansuMatch As Boolean
Private blnSF4KansuMatch As Boolean
Private blnSF2KansuMatch As Boolean
Private blnSTREETKansuMatch As Boolean
Private blnDaboPitchKansuMatch As Boolean

Private varDaboPitch(1 To 10) As Variant

Private strGetTateguKansuErrMsg As String

Private Type 引数格納
'*************************************
'Bind引数構造体                      *
' Bindで受け取った引数を登録する     *
'　　　　　　　　　　　　　　　　　　*
'------------------------------------*
' --品番について                  　 *
'  「in受注品番」                    *
'      引数で受け取った品番          *
'                                    *
'  「inブランド品番」                *
'　　　SxL品番はブランド品番に置換え *
'      その他はin受注品番を移送      *
'                                    *
'  「in検索品番」                    *
'　　　特注の場合は頭の「特 」を除去 *
'      その他はinブランド品番を移送  *
'*************************************
    in受注品番 As String
    inブランド品番 As String
    in検索品番 As String
    inセット数 As Integer
    in枚数 As Integer
    inDW As Double
    inDH As Double
    inCH As Double
    in施錠 As String
    in明り窓 As String
    in個別Spec As Variant
End Type

Private inArgument As 引数格納

Private Type 建具関数
'*************************************
'建具関数マスタ構造体                *
'マスタのカラムが増えたらここにも登録*
'*************************************
    基本図 As String
    芯組詳細図 As String
    組立図 As String
    建具品番 As String 'Not Null
    DW As String
    DH As String
    芯厚 As Double
    上下桟関数 As String
    上下桟本数 As Double
    額横関数1 As String
    額横本数1 As Double
    額横関数2 As String
    額横本数2 As Double
    柱関数 As String
    柱本数 As Double
    額縦関数 As String
    額縦本数 As Double
    額縦関数1 As String
    額縦本数1 As Double
    額縦関数2 As String
    額縦本数2 As Double
    手掛け長さ As Double
    手掛け個数 As Double
    袖長さ As Double
    袖本数 As Double
    ダボ下地長さ As Double
    ダボ下地本数 As Double
    中板芯厚 As Double
    中板上下桟関数 As String
    中板上下桟本数 As Double
    中板額横関数 As String
    中板額横本数 As Double
    中板柱関数 As String
    中板柱本数 As Double
    扉厚 As Double
    中板扉厚 As Double
    テノーナW関数 As String
    テノーナH関数 As String
    テノーナW関数2 As String
    テノーナH関数2 As String
    中板テノーナW関数 As String
    中板テノーナH関数 As String
    中板テノーナW関数2 As String
    中板テノーナH関数2 As String
    表面材厚み As Double
    表面材厚み2 As Double
    CU_AS_AW表面材W関数 As String
    CU_AS_AW表面材H関数 As String
    CU_AS_AW枚数 As Double
    CU_AS_AW表面材W関数2 As String
    CU_AS_AW表面材H関数2 As String
    CU_AS_AW枚数2 As Double
    AS_AW中板W関数 As String
    AS_AW中板H関数 As String
    AS_AW中板枚数 As String
    AS_AW中板W関数1 As String
    AS_AW中板H関数1 As String
    AS_AW中板枚数1 As Double
    AS_AW中板W関数2 As String
    AS_AW中板H関数2 As String
    AS_AW中板枚数2 As Double
    ST表面材上W関数 As String
    ST表面材上H関数 As String
    ST表面材上枚数 As String
    ST表面材中W関数 As String
    ST表面材中H関数 As String
    ST表面材中枚数 As String
    ST表面材下W関数 As String
    ST表面材下H関数 As String
    ST表面材下枚数 As String
    上框種類 As String
    上框長さ関数 As String
    上框長さ個数 As String
    下框種類 As String
    下框長さ関数 As String
    下框長さ個数 As Double
    '2.1.0 ADD**
    中框種類 As String
    上框高さ As Double
    下框高さ As Double
    '************************
    パネルソーW1関数 As String
    パネルソーW2関数 As String
    パネルソーH1関数 As String
    パネルソーH2関数 As String
    パネルソーH3関数 As String
    ガラス厚み As Double
    ガラスW関数 As String
    ガラスH関数 As String
    ガラス厚み1 As Double
    ガラスW関数1 As String
    ガラスH関数1 As String
    ダボセンター関数 As String
    ミラー表W As String
    ミラー表H As String
    ミラー表枚数 As Double
    ミラー裏W As String
    ミラー裏H As String
    ミラー裏枚数 As Double
End Type

Private TateguKansu As 建具関数

Private Type 建具図面
'*************************************
'建具図面マスタ構造体                *
'*************************************
    加工図 As String
    組立図 As String
End Type

Private TateguZumen As 建具図面

Private Type AW関数
'*************************************
'AW関数構造体 　　　　               *
'関数のカラムが増えたらここにも登録  *
'*************************************
    通番 As Integer 'Not Null
    DW_Max As Double 'Not Null
    DW_Min As Double 'Not Null
    芯組詳細図DADO As String
    芯組詳細図DCDEDPDU As String
    芯組詳細図DHDFDJ As String
    上下桟関数AW As Double
    額横関数1AW As Double
    額横本数1AW As Double
    額横関数2AW As Double
    額横関数2AW_DHDFDJ As Double
    額横本数2AW As Double
    額横本数2AW_DE As Double
    額横本数2AW_DH As Double
    額横本数2AW_DFDJ As Double
    額縦関数AW As String
    額縦本数AW As Double
    額縦本数AW_DE As Double
    額縦本数AW_DH As Double
    額縦本数AW_DFDJ As Double
    テノーナW関数AW As Double
    中板テノーナW関数AW As String
    CU_AS_AW表面材W関数AW As Double
    AS_AW中板W関数AW As String
    上框長さ関数AW As String
    下框長さ関数AW As String
    パネルソーW1関数AW As Double
    パネルソーW2関数AW As Double
    ガラスW関数AW As String
    鏡W関数AW As String
    開始 As String
    終了 As String
End Type

Private AWKansu As AW関数

Private Type SF4SG6関数
'*************************************
'SF4SG6関数構造体                    *
'*************************************
    区分 As Double
    CH_Max As Double
    CH_Min As Double
    CH枠外H As String
    上HFLG As Boolean
    上H関数 As String
    上枚数 As Double
    中HFLG As Boolean
    中H As Double
    中枚数 As Double
    中HFLG1 As Boolean
    中H1 As Double
    中枚数1 As Double
    下HFLG As Boolean
    下H関数 As String
    下枚数 As Double
    中板柱関数 As String
    中板テノーナH関数 As String
    パネルソーH1関数 As String
    パネルソーH2関数 As String
    パネルソーH3関数 As String
    上框長さ個数 As String
    ガラスH関数1 As String
    ダボピッチ1 As String
    ダボピッチ2 As String
    ダボピッチ3 As String
    ダボピッチ4 As String
    ダボピッチ5 As String
    ダボピッチ6 As String
    ダボピッチ7 As String
    ダボピッチ8 As String
    ダボピッチ9 As String
    ダボピッチ10 As String
End Type

Private SF4Kansu As SF4SG6関数

Private Type SF2SF3関数
'*************************************
'SF2SF3関数構造体                    *
'*************************************
    区分 As Double
    CH_Max As Double
    CH_Min As Double
    CH枠外H As String
    上HFLG As Boolean
    上H関数 As String
    上H関数_引戸 As String
    上枚数 As Double
    中HFLG As Boolean
    中H As Double
    中枚数 As Double
    中HFLG1 As Boolean
    中H1 As Double
    中枚数1 As Double
    下HFLG As Boolean
    下H関数 As String
    下H関数_引戸 As String
    下枚数 As Double
End Type

Private SF2Kansu As SF2SF3関数

Private Type STREET割付関数
'*************************************
'STREET割付関数                      *
'*************************************
    区分 As Double
    CH枠外H As String
    上HFLG As Boolean
    上H関数 As String
    上枚数 As Double
    中HFLG As Boolean
    中H As Double
    中枚数 As Double
    下HFLG As Boolean
    下H関数 As String
    下枚数 As Double
End Type

Private STREETKansu As STREET割付関数

Private Sub Class_Initialize()
'*************************************
'コンストラクタ                      *
'*************************************
    
    'DB接続
    setADOConnection
    
    FLG_Init
    
    strGetTateguKansuErrMsg = ""
    
End Sub

Private Sub Class_Terminate()
'*************************************
'デストラクタ　                      *
'*************************************
    On Error Resume Next
    
    RS_AllClose
    
    Set rsADO = Nothing
    Set rsADOAW = Nothing
    Set rsADOSF2 = Nothing
    Set rsADOSF4 = Nothing
    Set rsADOSTREET = Nothing
    Set rsADOKihonzu = Nothing
    Set rsADODaboPitch = Nothing
    
    cnADO.Close
    Set cnADO = Nothing

End Sub

Public Function Bind(ByVal 契約No As String, ByVal 品番 As String, ByVal 数量 As Integer, ByVal 枚数 As Integer, ByVal DW As Variant, ByVal DH As Variant, CH As Variant, ByVal 施錠 As Variant, ByVal 明り窓 As Variant, ByVal 個別Spec As Variant) As Boolean
'***********************************************************
'バインド処理（必須）
' プロパティを取得する場合は必ずこの関数を実行してから行う
'
'   戻り値:
'           True →引当成功
'           False→    失敗
'
'-----------------------------------------------------------
'   注意！
'     子扉の場合はDWに[子扉DW]を渡すこと
'***********************************************************

    Dim typArgument As 引数格納 '初期化用。データは移送しないこと！
    
    Dim bolEnd As Boolean
    
    On Error GoTo Err_Bind
    
    FLG_Init
    
    inArgument = typArgument
    
    With inArgument
        
        .in受注品番 = 品番
        .inセット数 = 数量
        .in枚数 = 枚数
        'DW,DH,CHはNullの場合あり
        .inDW = CDbl(Nz(DW, 0))
        .inDH = CDbl(Nz(DH, 0))
        .inCH = CDbl(Nz(CH, 0))
        .in施錠 = Nz(施錠, "")
        .in明り窓 = Nz(明り窓, "-")
        
        .in個別Spec = 個別Spec
        
        .inブランド品番 = TateguHinban()
        
        If .inブランド品番 = "" Then
            Err.Raise 9999, , "SxL品番読替エラー 契約No =" & 契約No & " 品番= " & .in受注品番
        Else
            .in検索品番 = Trim(Replace(.inブランド品番, "特 ", ""))
        End If
        
        If IsNull(個別Spec) Then
            Err.Raise 9999, , "SpecNullエラー 契約No =" & 契約No & " 品番= " & .in受注品番
        End If
        
    End With
    
    If 契約No Like "SR*" Then
        'クレスト処理は別に定義
    ElseIf GetTateguKansu() Then
        bolEnd = True
    End If
    
    Bind = bolEnd
    
    Exit Function
    
Err_Bind:
    If Err.Number <> 9999 Then
        MsgBox Err.Description, vbCritical
    Else
        Debug.Print Err.Description
    End If
    
    inArgument = typArgument
    Bind = False
End Function

Private Property Get TateguHinban() As String
'***********************************************************
'建具品番取得
' SxL品番の時はここでブランド品番に置き換える
'
'
'***********************************************************
    Dim rsSxL As New ADODB.Recordset
    Dim strSQL As String
    Dim varTateguHinban As Variant
    
    varTateguHinban = Null
    
    With inArgument
        If IsSxL(.in受注品番, varTateguHinban) Then
            TateguHinban = Nz(varTateguHinban, "")
        Else
            TateguHinban = .in受注品番
        End If
    End With

End Property

Private Function GetTateguKansu() As Boolean
'***********************************************************
'関数テーブル引当
'
'
'***********************************************************

    '初期化用。データは移送しないこと！******
    Dim typTateguKansuInit As 建具関数
    Dim typAWKansuInit As AW関数
    Dim typSF4SG6KansuInit As SF4SG6関数
    Dim typSF2SF3KansuInit As SF2SF3関数
    Dim typSTREETKansuInit As STREET割付関数
    Dim typTateguZumenKansuInit As 建具図面
    '****************************************
    
    Dim strSQL As String
    Dim strSQLZM As String
    Dim strSQLAW As String
    Dim strSQLSF4 As String
    Dim strSQLSF2 As String
    Dim strSQLSTREET As String
    Dim strSQLDabo As String
    
    Dim strKansuTable As String
    
    Dim bytStreetKbn As Byte
    
    Dim i As Integer
    
    On Error GoTo Err_GetTateguKansu
    
    TateguKansu = typTateguKansuInit
    AWKansu = typAWKansuInit
    SF4Kansu = typSF4SG6KansuInit
    SF2Kansu = typSF2SF3KansuInit
    STREETKansu = typSTREETKansuInit
    TateguZumen = typTateguZumenKansuInit
    Erase varDaboPitch

    strGetTateguKansuErrMsg = ""
    
    'オープン済みのレコードセットがある場合、一旦閉じる
    RS_AllClose
    
    With inArgument
    
        If .in個別Spec Like "*1007" Or .in個別Spec Like "*1011" Or .in個別Spec Like "*1111" Or .in個別Spec Like "*1103" Or .in個別Spec Like "*1105" Or .in個別Spec Like "*1010" Or .in個別Spec Like "*1009" Or right(.in個別Spec, 4) >= "1304" Then
            strKansuTable = "T_建具関数ﾏｽﾀ_1007仕様"
        Else
            strKansuTable = "T_建具関数ﾏｽﾀ"
        End If
    
        strSQL = ""
        strSQL = strSQL & "select * "
        strSQL = strSQL & "FROM " & strKansuTable & " "
        strSQL = strSQL & "where '" & .in検索品番 & " ' Like '%' + Replace(Replace(Replace(建具品番, '?', '_'), '#', '[0-9]'), '*', '%') + '%';"
        
        
    End With
    
    Set rsADO = New ADODB.Recordset
    rsADO.Open strSQL, cnADO, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rsADO.EOF
        If inArgument.in検索品番 Like "*" & rsADO![建具品番] & "*" Then
            
            With TateguKansu
                .基本図 = Nz(rsADO![基本図], "")
                .芯組詳細図 = Nz(rsADO![芯組詳細図], "")
                .組立図 = Nz(rsADO![組立図], "")
                .建具品番 = inArgument.inブランド品番
                .DW = Nz(rsADO![DW], "0")
                .DH = Nz(rsADO![DH], "0")
                .芯厚 = Nz(rsADO![芯厚], 0)
                .上下桟関数 = Nz(rsADO![上下桟関数], "0")
                .上下桟本数 = Nz(rsADO![上下桟本数], 0)
                .額横関数1 = Nz(rsADO![額横関数1], "0")
                .額横本数1 = Nz(rsADO![額横本数1], 0)
                .額横関数2 = Nz(rsADO![額横関数2], "0")
                .額横本数2 = Nz(rsADO![額横本数2], 0)
                .柱関数 = Nz(rsADO![柱関数], "0")
                .柱本数 = Nz(rsADO![柱本数], 0)
                .額縦関数 = Nz(rsADO![額縦関数], "0")
                .額縦本数 = Nz(rsADO![額縦本数], 0)
                .額縦関数1 = Nz(rsADO![額縦関数1], "0")
                .額縦本数1 = Nz(rsADO![額縦本数1], 0)
                .額縦関数2 = Nz(rsADO![額縦関数2], "0")
                .額縦本数2 = Nz(rsADO![額縦本数2], 0)
                .手掛け長さ = Nz(rsADO![手掛け長さ], 0)
                .手掛け個数 = Nz(rsADO![手掛け個数], 0)
                .袖長さ = Nz(rsADO![袖長さ], 0)
                .袖本数 = Nz(rsADO![袖本数], 0)
                .ダボ下地長さ = Nz(rsADO![ダボ下地長さ], 0)
                .ダボ下地本数 = Nz(rsADO![ダボ下地本数], 0)
                .中板芯厚 = Nz(rsADO![中板芯厚], 0)
                .中板上下桟関数 = Nz(rsADO![中板上下桟関数], "0")
                .中板上下桟本数 = Nz(rsADO![中板上下桟本数], 0)
                .中板額横関数 = Nz(rsADO![中板額横関数], "0")
                .中板額横本数 = Nz(rsADO![中板額横本数], 0)
                .中板柱関数 = Nz(rsADO![中板柱関数], "0")
                .中板柱本数 = Nz(rsADO![中板柱本数], 0)
                .扉厚 = Nz(rsADO![扉厚], 0)
                .中板扉厚 = Nz(rsADO![中板扉厚], 0)
                .テノーナW関数 = Nz(rsADO![テノーナW関数], "0")
                .テノーナH関数 = Nz(rsADO![テノーナH関数], "0")
                .テノーナW関数2 = Nz(rsADO![テノーナW関数2], "0")
                .テノーナH関数2 = Nz(rsADO![テノーナH関数2], "0")
                .中板テノーナW関数 = Nz(rsADO![中板テノーナW関数], "0")
                .中板テノーナH関数 = Nz(rsADO![中板テノーナH関数], "0")
                .中板テノーナW関数2 = Nz(rsADO![中板テノーナW関数2], "0")
                .中板テノーナH関数2 = Nz(rsADO![中板テノーナH関数2], "0")
                .表面材厚み = Nz(rsADO![表面材厚み], 0)
                .表面材厚み2 = Nz(rsADO![表面材厚み2], 0)
                .CU_AS_AW表面材W関数 = Nz(rsADO![CU_AS_AW表面材W関数], "0")
                .CU_AS_AW表面材H関数 = Nz(rsADO![CU_AS_AW表面材H関数], "0")
                .CU_AS_AW枚数 = Nz(rsADO![CU_AS_AW枚数], 0)
                .CU_AS_AW表面材W関数2 = Nz(rsADO![CU_AS_AW表面材W関数2], "0")
                .CU_AS_AW表面材H関数2 = Nz(rsADO![CU_AS_AW表面材H関数2], "0")
                .CU_AS_AW枚数2 = Nz(rsADO![CU_AS_AW枚数2], 0)
                .AS_AW中板W関数 = Nz(rsADO![AS_AW中板W関数], "0")
                .AS_AW中板H関数 = Nz(rsADO![AS_AW中板H関数], "0")
                .AS_AW中板枚数 = Nz(rsADO![AS_AW中板枚数], "0")
                .AS_AW中板W関数1 = Nz(rsADO![AS_AW中板W関数1], "0")
                .AS_AW中板H関数1 = Nz(rsADO![AS_AW中板H関数1], "0")
                .AS_AW中板枚数1 = Nz(rsADO![AS_AW中板枚数1], 0)
                .AS_AW中板W関数2 = Nz(rsADO![AS_AW中板W関数2], "0")
                .AS_AW中板H関数2 = Nz(rsADO![AS_AW中板H関数2], "0")
                .AS_AW中板枚数2 = Nz(rsADO![AS_AW中板枚数2], 0)
                .ST表面材上W関数 = Nz(rsADO![ST表面材上W関数], "0")
                .ST表面材上H関数 = Nz(rsADO![ST表面材上H関数], "0")
                .ST表面材上枚数 = Nz(rsADO![ST表面材上枚数], "0")
                .ST表面材中W関数 = Nz(rsADO![ST表面材中W関数], "0")
                .ST表面材中H関数 = Nz(rsADO![ST表面材中H関数], "0")
                .ST表面材中枚数 = Nz(rsADO![ST表面材中枚数], "0")
                .ST表面材下W関数 = Nz(rsADO![ST表面材下W関数], "0")
                .ST表面材下H関数 = Nz(rsADO![ST表面材下H関数], "0")
                .ST表面材下枚数 = Nz(rsADO![ST表面材下枚数], "0")
                .上框種類 = Nz(rsADO![上框種類], "")
                .上框長さ関数 = Nz(rsADO![上框長さ関数], "0")
                .上框長さ個数 = Nz(rsADO![上框長さ個数], "0")
                .下框種類 = Nz(rsADO![下框種類], "")
                .下框長さ関数 = Nz(rsADO![下框長さ関数], "0")
                .下框長さ個数 = Nz(rsADO![下框長さ個数], 0)
                .パネルソーW1関数 = Nz(rsADO![パネルソーW1関数], "0")
                .パネルソーW2関数 = Nz(rsADO![パネルソーW2関数], "0")
                .パネルソーH1関数 = Nz(rsADO![パネルソーH1関数], "0")
                .パネルソーH2関数 = Nz(rsADO![パネルソーH2関数], "0")
                .パネルソーH3関数 = Nz(rsADO![パネルソーH3関数], "0")
                .ガラス厚み = Nz(rsADO![ガラス厚み], 0)
                .ガラスW関数 = Nz(rsADO![ガラスW関数], "0")
                .ガラスH関数 = Nz(rsADO![ガラスH関数], "0")
                .ガラス厚み1 = Nz(rsADO![ガラス厚み1], 0)
                .ガラスW関数1 = Nz(rsADO![ガラスW関数1], "0")
                .ガラスH関数1 = Nz(rsADO![ガラスH関数1], "0")
                .ダボセンター関数 = Nz(rsADO![ダボセンター関数], "0")
                .ミラー表W = Nz(rsADO![ミラー表W], "0")
                .ミラー表H = Nz(rsADO![ミラー表H], "0")
                .ミラー表枚数 = Nz(rsADO![ミラー表枚数], 0)
                .ミラー裏W = Nz(rsADO![ミラー裏W], "0")
                .ミラー裏H = Nz(rsADO![ミラー裏H], "0")
                .ミラー裏枚数 = Nz(rsADO![ミラー裏枚数], 0)
                '2.1.0 ADD**************************
                .中框種類 = Nz(rsADO![中框種類], "")
                .上框高さ = Nz(rsADO![上框高さ], 0)
                .下框高さ = Nz(rsADO![下框高さ], 0)
                '************************************************
            End With
            
            blnKansuMatch = True
            Exit Do
        End If
        
        rsADO.MoveNext
    Loop
    
    If Not blnKansuMatch Then
        Err.Raise 9999, , "品番に該当する建具関数がありません。" & vbCrLf & "品番 = " & inArgument.in受注品番
    End If
    
    With inArgument
        If Not .in施錠 Like "*特*" Then
            strSQLZM = ""
            strSQLZM = strSQLZM & "select 加工図,組立図 from T_建具図面ﾏｽﾀ "
            strSQLZM = strSQLZM & "where 基本図 = '" & TateguKansu.基本図 & "' "
            strSQLZM = strSQLZM & "and (ﾊﾝﾄﾞﾙ = '" & fncstrHandle_Name(.in施錠, CStr(.in個別Spec)) & "' "
            strSQLZM = strSQLZM & "or ﾊﾝﾄﾞﾙ = '" & .in施錠 & "') "
            strSQLZM = strSQLZM & "and 仕様 = '" & left(.in個別Spec, 3) & "' "
            strSQLZM = strSQLZM & "and '" & right(.in個別Spec, 4) & "' between 開始 and 終了 "
            
            Set rsADOKihonzu = New ADODB.Recordset
            rsADOKihonzu.Open strSQLZM, cnADO, adOpenForwardOnly, adLockReadOnly
            
            If Not rsADOKihonzu.EOF Then
                TateguZumen.加工図 = Nz(rsADOKihonzu![加工図], "")
                TateguZumen.組立図 = Nz(rsADOKihonzu![組立図], "")
                
                blnZumenKansuMatch = True
            Else
                '☆☆☆ここ重要！☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆
                '                                                                                                        ☆
                '☆建具図面マスタに引当ができなかった場合は建具関数を初期化する                                          ☆
                '  （建具図面マスタがない場合建具関数もないものとする。データは移送しない）                              ☆
                '                                                                                                        ☆
                '   --理由：元のシステムがそうなっていたから。なぜかは聞かないでください・・・ 　　                      ☆
                '                                                                                                        ☆
                '   --補足                                                                                               ☆
                '         20170608  設計課担当（奈良課長 敬称登録時）に確認                                              ☆
                '                   『影響範囲が判らないのでそのままにしておいてください』                               ☆
                '☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆
                
                Err.Raise 9999, , "品番に該当する建具図面がありません。" & vbCrLf & "品番 = " & inArgument.in受注品番
                
            End If
            
        End If
        
    End With
    
    
    
    If IsAWKansuSearch_Needed(inArgument.in検索品番) Then
    
        strSQLAW = ""
        strSQLAW = strSQLAW & "SELECT * FROM T_AW関数 "
        strSQLAW = strSQLAW & "WHERE " & inArgument.inDW & " Between DW_Min AND DW_Max "
        strSQLAW = strSQLAW & "AND '" & right(Nz(inArgument.in個別Spec, "0000"), 4) & "' Between 開始 AND 終了;"
        
       
        Set rsADOAW = New ADODB.Recordset
        rsADOAW.Open strSQLAW, cnADO, adOpenForwardOnly, adLockReadOnly
        
        If Not rsADO.EOF Then
            With AWKansu
                .通番 = rsADOAW![通番]
                .DW_Max = rsADOAW![DW_Max]
                .DW_Min = rsADOAW![DW_Min]
                .芯組詳細図DADO = Nz(rsADOAW![芯組詳細図DADO], "")
                .芯組詳細図DCDEDPDU = Nz(rsADOAW![芯組詳細図DCDEDPDU], "")
                .芯組詳細図DHDFDJ = Nz(rsADOAW![芯組詳細図DHDFDJ], "")
                .上下桟関数AW = Nz(rsADOAW![上下桟関数AW], 0)
                .額横関数1AW = Nz(rsADOAW![額横関数1AW], 0)
                .額横本数1AW = Nz(rsADOAW![額横本数1AW], 0)
                .額横関数2AW = Nz(rsADOAW![額横関数2AW], 0)
                .額横関数2AW_DHDFDJ = Nz(rsADOAW![額横関数2AW_DHDFDJ], 0)
                .額横本数2AW = Nz(rsADOAW![額横本数2AW], 0)
                .額横本数2AW_DE = Nz(rsADOAW![額横本数2AW_DE], 0)
                .額横本数2AW_DH = Nz(rsADOAW![額横本数2AW_DH], 0)
                .額横本数2AW_DFDJ = Nz(rsADOAW![額横本数2AW_DFDJ], 0)
                .額縦関数AW = Nz(rsADOAW![額縦関数AW], "0")
                .額縦本数AW = Nz(rsADOAW![額縦本数AW], 0)
                .額縦本数AW_DE = Nz(rsADOAW![額縦本数AW_DE], 0)
                .額縦本数AW_DH = Nz(rsADOAW![額縦本数AW_DH], 0)
                .額縦本数AW_DFDJ = Nz(rsADOAW![額縦本数AW_DFDJ], 0)
                .テノーナW関数AW = Nz(rsADOAW![ﾃﾉｰﾅW関数AW], 0)
                .中板テノーナW関数AW = Nz(rsADOAW![中板ﾃﾉｰﾅW関数AW], "0")
                .CU_AS_AW表面材W関数AW = Nz(rsADOAW![CU_AS_AW表面材W関数AW], 0)
                .AS_AW中板W関数AW = Nz(rsADOAW![AS_AW中板W関数AW], "0")
                .上框長さ関数AW = Nz(rsADOAW![上框長さ関数AW], "0")
                .下框長さ関数AW = Nz(rsADOAW![下框長さ関数AW], "0")
                .パネルソーW1関数AW = Nz(rsADOAW![ﾊﾟﾈﾙｿｰW1関数AW], 0)
                .パネルソーW2関数AW = Nz(rsADOAW![ﾊﾟﾈﾙｿｰW2関数AW], 0)
                .ガラスW関数AW = Nz(rsADOAW![ガラスW関数AW], "0")
                .鏡W関数AW = Nz(rsADOAW![鏡W関数AW], "0")
                .開始 = Nz(rsADOAW![開始], "0")
                .終了 = Nz(rsADOAW![終了], "0")
            End With
            
            blnAWKansuMatch = True
        End If
        
        If Not blnAWKansuMatch Then
            Err.Raise 9999, , "注文のDWに該当するAW関数がありません。" & vbCrLf & "DW = " & inArgument.inDW
        End If
    
    End If
    
    strSQLSF4 = ""
    
    With inArgument
        If .in検索品番 Like "F?WD?-????G*" Or .in検索品番 Like "F?WP?-????G*" Or .in検索品番 Like "F?WD??-????G*" Or .in検索品番 Like "F?WP??-????G*" Then
            
            strSQLSF4 = strSQLSF4 & "SELECT * FROM T_SF4SG6割付関数 "
            strSQLSF4 = strSQLSF4 & "WHERE " & .inCH & " Between CH_Min AND CH_Max "
            strSQLSF4 = strSQLSF4 & "and [CH/枠外H] = 'SG6' "
            
        ElseIf .in検索品番 Like "F?WD?-????P*" Or .in検索品番 Like "F?WP?-????P*" Or .in検索品番 Like "F?WD??-????P*" Or .in検索品番 Like "F?WP??-????P*" Then

            strSQLSF4 = strSQLSF4 & "SELECT * FROM T_SF4SG6割付関数 "
            strSQLSF4 = strSQLSF4 & "WHERE " & .inCH & " Between CH_Min AND CH_Max "
            strSQLSF4 = strSQLSF4 & "and [CH/枠外H] = 'SF4' "
        
        End If
    End With
    
    If strSQLSF4 <> "" Then
        Set rsADOSF4 = New ADODB.Recordset
        rsADOSF4.Open strSQLSF4, cnADO, adOpenForwardOnly, adLockReadOnly
        
        If Not rsADOSF4.EOF Then
            With SF4Kansu
                .区分 = Nz(rsADOSF4![区分], 0)
                .CH_Max = Nz(rsADOSF4![CH_Max], 0)
                .CH_Min = Nz(rsADOSF4![CH_Min], 0)
                .CH枠外H = Nz(rsADOSF4![CH/枠外H], "0")
                .上HFLG = Nz(rsADOSF4![上HFLG], 0)
                .上H関数 = Nz(rsADOSF4![上H関数], "0")
                .上枚数 = Nz(rsADOSF4![上枚数], 0)
                .中HFLG = Nz(rsADOSF4![中HFLG], 0)
                .中H = Nz(rsADOSF4![中H], 0)
                .中枚数 = Nz(rsADOSF4![中枚数], 0)
                .中HFLG1 = Nz(rsADOSF4![中HFLG1], 0)
                .中H1 = Nz(rsADOSF4![中H1], 0)
                .中枚数1 = Nz(rsADOSF4![中枚数1], 0)
                .下HFLG = Nz(rsADOSF4![下HFLG], 0)
                .下H関数 = Nz(rsADOSF4![下H関数], "0")
                .下枚数 = Nz(rsADOSF4![下枚数], 0)
                .中板柱関数 = Nz(rsADOSF4![中板柱関数], "0")
                .中板テノーナH関数 = Nz(rsADOSF4![中板ﾃﾉｰﾅH関数], "0")
                .パネルソーH1関数 = Nz(rsADOSF4![ﾊﾟﾈﾙｿｰH1関数], "0")
                .パネルソーH2関数 = Nz(rsADOSF4![ﾊﾟﾈﾙｿｰH2関数], "0")
                .パネルソーH3関数 = Nz(rsADOSF4![ﾊﾟﾈﾙｿｰH3関数], "0")
                .上框長さ個数 = Nz(rsADOSF4![上框長さ個数], "0")
                .ガラスH関数1 = Nz(rsADOSF4![ガラスH関数1], "0")
                .ダボピッチ1 = Nz(rsADOSF4![ダボピッチ1], "0")
                .ダボピッチ2 = Nz(rsADOSF4![ダボピッチ2], "0")
                .ダボピッチ3 = Nz(rsADOSF4![ダボピッチ3], "0")
                .ダボピッチ4 = Nz(rsADOSF4![ダボピッチ4], "0")
                .ダボピッチ5 = Nz(rsADOSF4![ダボピッチ5], "0")
                .ダボピッチ6 = Nz(rsADOSF4![ダボピッチ6], "0")
                .ダボピッチ7 = Nz(rsADOSF4![ダボピッチ7], "0")
                .ダボピッチ8 = Nz(rsADOSF4![ダボピッチ8], "0")
                .ダボピッチ9 = Nz(rsADOSF4![ダボピッチ9], "0")
                .ダボピッチ10 = Nz(rsADOSF4![ダボピッチ10], "0")
            End With
            
            blnSF4KansuMatch = True
                
        End If
        
        If Not blnSF4KansuMatch Then
            Err.Raise 9999, , "注文の品番、CHに該当するSF4SG6割付関数がありません。" & vbCrLf & "品番 = " & inArgument.in受注品番 & ",CH = " & inArgument.inCH
        End If
    End If
    
    strSQLSF2 = ""
    
    With inArgument
        If .in検索品番 Like "F?PD*" Or .in検索品番 Like "F?PP*" Or .in検索品番 Like "F?PC*" Or .in検索品番 Like "F?OD*" Or .in検索品番 Like "F?OP*" Or .in検索品番 Like "F?OC*" Then
            
            strSQLSF2 = strSQLSF2 & "SELECT * FROM T_SF2SF3割付関数 "
            strSQLSF2 = strSQLSF2 & "WHERE " & .inCH & " Between CH_Min AND CH_Max "
        
        End If
    End With
    
    If strSQLSF2 <> "" Then
        Set rsADOSF2 = New ADODB.Recordset
        rsADOSF2.Open strSQLSF2, cnADO, adOpenForwardOnly, adLockReadOnly
        
        If Not rsADOSF2.EOF Then
            With SF2Kansu
                .区分 = Nz(rsADOSF2![区分], 0)
                .CH_Max = Nz(rsADOSF2![CH_Max], 0)
                .CH_Min = Nz(rsADOSF2![CH_Min], 0)
                .CH枠外H = Nz(rsADOSF2![CH/枠外H], "0")
                .上HFLG = Nz(rsADOSF2![上HFLG], 0)
                .上H関数 = Nz(rsADOSF2![上H関数], "0")
                .上H関数_引戸 = Nz(rsADOSF2![上H関数_引戸], "0")
                .上枚数 = Nz(rsADOSF2![上枚数], 0)
                .中HFLG = Nz(rsADOSF2![中HFLG], 0)
                .中H = Nz(rsADOSF2![中H], 0)
                .中枚数 = Nz(rsADOSF2![中枚数], 0)
                .中HFLG1 = Nz(rsADOSF2![中HFLG1], 0)
                .中H1 = Nz(rsADOSF2![中H1], 0)
                .中枚数1 = Nz(rsADOSF2![中枚数1], 0)
                .下HFLG = Nz(rsADOSF2![下HFLG], 0)
                .下H関数 = Nz(rsADOSF2![下H関数], "0")
                .下H関数_引戸 = Nz(rsADOSF2![下H関数_引戸], "0")
                .下枚数 = Nz(rsADOSF2![下枚数], 0)
            End With
            
            blnSF2KansuMatch = True
                
        End If
        
        If Not blnSF2KansuMatch Then
            Err.Raise 9999, , "注文の品番、CHに該当するSF2SF3割付関数がありません。" & vbCrLf & "品番 = " & inArgument.in受注品番 & ",CH = " & inArgument.inCH
        End If
    End If
        
    strSQLSTREET = ""
    
    With inArgument
        If .in検索品番 Like "F?SD*" Or .in検索品番 Like "F?SP*" Or .in検索品番 Like "F?SC*" Or .in検索品番 Like "K?SD*" Or .in検索品番 Like "F?SSC-*" Then
            
            Select Case .inCH
                Case Is > 2600
                    bytStreetKbn = 0
                Case Is >= 2451
                    bytStreetKbn = 1
                Case Is >= 2401
                    bytStreetKbn = 2
                Case Is >= 2255
                    bytStreetKbn = 3
                Case Is >= 2101
                    bytStreetKbn = 4
                Case Is >= 1951
                    bytStreetKbn = 5
                Case Is >= 1901
                    bytStreetKbn = 6
                Case Is >= 1801
                    bytStreetKbn = 7
                Case Is >= 1551
                    bytStreetKbn = 8
                Case Is >= 1421
                    bytStreetKbn = 9
                Case Is >= 1381
                    bytStreetKbn = 10
                Case Is >= 1331
                    bytStreetKbn = 11
                Case Is >= 1082
                    bytStreetKbn = 12
                Case Else
                    bytStreetKbn = 0
                
            End Select
            
            strSQLSTREET = strSQLSTREET & "SELECT * FROM T_STREET割付関数 "
            strSQLSTREET = strSQLSTREET & "WHERE 区分 = " & bytStreetKbn & " "
        
        End If
    End With
        
    If strSQLSTREET <> "" Then
        Set rsADOSTREET = New ADODB.Recordset
        rsADOSTREET.Open strSQLSTREET, cnADO, adOpenForwardOnly, adLockReadOnly
        
        If Not rsADOSTREET.EOF Then
            With STREETKansu
                .区分 = Nz(rsADOSTREET![区分], 0)
                .CH枠外H = Nz(rsADOSTREET![CH/枠外H], "0")
                .上HFLG = Nz(rsADOSTREET![上HFLG], 0)
                .上H関数 = Nz(rsADOSTREET![上H関数], "0")
                .上枚数 = Nz(rsADOSTREET![上枚数], 0)
                .中HFLG = Nz(rsADOSTREET![中HFLG], 0)
                .中H = Nz(rsADOSTREET![中H], 0)
                .中枚数 = Nz(rsADOSTREET![中枚数], 0)
                .下HFLG = Nz(rsADOSTREET![下HFLG], 0)
                .下H関数 = Nz(rsADOSTREET![下H関数], "0")
                .下枚数 = Nz(rsADOSTREET![下枚数], 0)
            End With
            
            blnSTREETKansuMatch = True
                
        End If
        
        If Not blnSTREETKansuMatch Then
            Err.Raise 9999, , "注文の品番、CHに該当するSTREET割付関数がありません。" & vbCrLf & "品番 = " & inArgument.in受注品番 & ",CH = " & inArgument.inCH
        End If
    
    End If
    
    
    'ダボピッチ関数
    If Not blnSF4KansuMatch Then
        strSQLDabo = ""
        strSQLDabo = strSQLDabo & "select * from T_ﾀﾞﾎﾞﾋﾟｯﾁ関数 "
        strSQLDabo = strSQLDabo & "where '" & inArgument.in検索品番 & " ' Like '%' + Replace(Replace(Replace(品番, '?', '_'), '#', '[0-9]'), '*', '%') + '%';"
        
        Set rsADODaboPitch = New ADODB.Recordset
        rsADODaboPitch.Open strSQLDabo, cnADO, adOpenForwardOnly, adLockReadOnly
        
        If Not rsADODaboPitch.EOF Then
            
            For i = LBound(varDaboPitch) To UBound(varDaboPitch)
                varDaboPitch(i) = rsADODaboPitch("ﾀﾞﾎﾞﾋﾟｯﾁ" & i)
            Next
            
            blnDaboPitchKansuMatch = True
            
        End If
        
    End If
        
    GetTateguKansu = True
    
    GoTo Exit_GetTateguKansu

Err_GetTateguKansu:
    
    strGetTateguKansuErrMsg = Err.Description
    
    'エラーの場合は初期化
    TateguKansu = typTateguKansuInit
    AWKansu = typAWKansuInit
    SF4Kansu = typSF4SG6KansuInit
    SF2Kansu = typSF2SF3KansuInit
    STREETKansu = typSTREETKansuInit
    TateguZumen = typTateguZumenKansuInit
    
    GetTateguKansu = False
    
Exit_GetTateguKansu:

    RS_AllClose
    
End Function

Public Property Get 基本図() As String
    
    基本図 = TateguKansu.基本図
    
End Property

Public Property Get 扉厚() As Variant
    
    扉厚 = TateguKansu.扉厚
    
End Property

Public Property Get 中板扉厚() As Variant
    
    中板扉厚 = TateguKansu.中板扉厚
    
End Property

Public Property Get 芯厚() As Variant
    
'芯厚を取得するには以下のモジュールがスコープ範囲にあることが必要です
'①Module_帳票関数

    Dim typKidori As KidoriData
    
    With inArgument
        bolFncKidoriData .in個別Spec, .in検索品番, .in枚数, .inDW, .inDH, .in明り窓, .in施錠, typKidori
        
        If Nz(typKidori.out_dblShinAtsu, 0) > 0 Then
            芯厚 = typKidori.out_dblShinAtsu
        Else
            芯厚 = 0
        End If
    End With
    
End Property

Public Property Get 表面材厚み() As Variant
    
    表面材厚み = TateguKansu.表面材厚み
    
End Property

Public Property Get 表面材厚み2() As Variant

    表面材厚み2 = TateguKansu.表面材厚み2

End Property

Public Property Get 芯組詳細図() As String
   
'2.13.0
'   →H2700が漏れていたため追加

    With inArgument
        If .inDH > 2589 Then
            芯組詳細図 = fncstrShingumiShousai2700(.in検索品番, .inDH, .in施錠)
        Else
            芯組詳細図 = fncstrShingumiShousai(.in検索品番, .inDH)
        End If
    End With
    
End Property

Public Property Get ガラス厚() As Variant
    
    ガラス厚 = TateguKansu.ガラス厚み
    
End Property

Public Property Get ガラス種類枚数(ByVal 種類枚数 As String) As Variant
    
'2.3.0 1801仕様追加

    Dim strShurui As String
    Dim intGlass As Integer
    
    With inArgument
        If .in検索品番 Like "F?WD?-####G*" Or .in検索品番 Like "F?WP?-####G*" Or .in検索品番 Like "F?WD??-####G*" Or .in検索品番 Like "F?WP??-####G*" Then
            strShurui = "ﾀﾍﾟ"
            intGlass = .in枚数 * 2
            
        ElseIf .in検索品番 Like "F?WD?-####P*" Or .in検索品番 Like "F?WP?-####P*" Or .in検索品番 Like "F?WD??-####P*" Or .in検索品番 Like "F?WP??-####P*" Then
            strShurui = "ﾀﾍﾟ"
            intGlass = .in枚数 * 2
            
        ElseIf .in検索品番 Like "*-####MF*" Then
            strShurui = "鏡"
            intGlass = .in枚数
            
        ElseIf .in検索品番 Like "F?GD*-####GF*" Or .in検索品番 Like "F?GP*-####GF*" Or .in検索品番 Like "F?GC*-####GF*" Or .in検索品番 Like "F?OD*-####G*" Or .in検索品番 Like "F?OP*-####G*" Or .in検索品番 Like "F?OC*-####G*" Then
            strShurui = "ﾐﾗｰ"
            intGlass = .in枚数 * 2
            
        ElseIf .in検索品番 Like "F?GD*-####G*" Or .in検索品番 Like "F?GP*-####G*" Or .in検索品番 Like "F?GC*-####G*" Then
            If .in検索品番 Like "*-####?W*" Then strShurui = "SW塗装"
            If .in検索品番 Like "*-####?J*" Then strShurui = "JB塗装"
            intGlass = .in枚数 * 2
            
        ElseIf .in検索品番 Like "*-####C*" Or .in検索品番 Like "*-####S*" Or .in検索品番 Like "*-####D*" Or .in検索品番 Like "*-####G*" Or .in検索品番 Like "*-####V*" Or .in検索品番 Like "*-####Q*" Or .in検索品番 Like "*-####E*" Then
            If .in検索品番 Like "*-####?T*" Then strShurui = "透明"
            If .in検索品番 Like "*-####?R*" Then strShurui = "ｸﾞﾘｰﾝ"
            If .in検索品番 Like "*-####?G*" Then strShurui = "ｸﾞﾚｰ"
            If .in検索品番 Like "*-####?B*" Then strShurui = "ﾌﾞﾛﾝｽﾞ"
            If .in検索品番 Like "*-####?S*" Then strShurui = "ﾀﾍﾟ"
            If .in検索品番 Like "*-####?H*" Then strShurui = "ﾊｰﾌﾐﾗｰ"
            
            If .in検索品番 Like "*-####?I*" Then strShurui = "強化透明"
            If .in検索品番 Like "*-####?J*" Then strShurui = "強化ﾀﾍﾟ"
            If .in検索品番 Like "*-####?K*" Then strShurui = "強化ｸﾞﾚｰ"
            If .in検索品番 Like "*-####?L*" Then strShurui = "強化ﾊｰﾌ"
            
            If .in検索品番 Like "F?ED?-*" Or .in検索品番 Like "F?FD?-*" Or .in検索品番 Like "F?EP?-*" Or .in検索品番 Like "F?FP?-*" Or .in検索品番 Like "F?ED??-*" Or .in検索品番 Like "F?FD??-*" Or .in検索品番 Like "F?EP??-*" Or .in検索品番 Like "F?FP??-*" Then
                intGlass = .in枚数 * 2
            '20171220 K.Asayama Change Airが含まれないよう変更
            ElseIf Not IsAir(.in検索品番) Then
                intGlass = .in枚数
            End If
            
        ElseIf IsTerraceGlass(.in検索品番) Then
            If .in検索品番 Like "*-####?C*" Then strShurui = "ﾁｪｯｶｰｶﾞﾗｽ"
            If .in検索品番 Like "*-####?A*" Then strShurui = "ｱﾙﾏｼﾞﾛｶﾞﾗｽ"
            If .in検索品番 Like "*-####?D*" Then strShurui = "ﾄﾞｯﾄｶﾞﾗｽ"
            If .in検索品番 Like "*-####?P*" Then strShurui = "ﾋﾟﾝｽﾄﾗｲﾌﾟ"
            If .in検索品番 Like "*-####?V*" Then strShurui = "ﾊﾞｰﾁｶﾙｶﾞﾗｽ"
            intGlass = .in枚数
        Else
            strShurui = ""
            intGlass = 0
        End If
    End With
    
    If 種類枚数 = "種類" Then
        ガラス種類枚数 = strShurui

    ElseIf 種類枚数 = "枚数" Then
        ガラス種類枚数 = intGlass
        
    End If
    
End Property

Public Property Get ガラス種類枚数2(ByVal 種類枚数 As String) As Variant

    Dim strShurui As String
    Dim intGlass As Integer
    
    With inArgument
        If .in検索品番 Like "F?WD?-####G*" Or .in検索品番 Like "F?WP?-####G*" Or .in検索品番 Like "F?WD??-####G*" Or .in検索品番 Like "F?WP??-####G*" Then

            If .in検索品番 Like "*-####?T*" Then strShurui = "透明"
            If .in検索品番 Like "*-####?R*" Then strShurui = "ｸﾞﾘｰﾝ"
            If .in検索品番 Like "*-####?G*" Then strShurui = "ｸﾞﾚｰ"
            If .in検索品番 Like "*-####?B*" Then strShurui = "ﾌﾞﾛﾝｽﾞ"
            If .in検索品番 Like "*-####?S*" Then strShurui = "ﾀﾍﾟ"

            intGlass = .in枚数
            
        ElseIf TateguKansu.ガラス厚み1 > 0 Then
            strShurui = ""
            intGlass = .in枚数 * 2
            
        Else
            strShurui = ""
            intGlass = 0
        End If
    End With
    
    If 種類枚数 = "種類" Then
        ガラス種類枚数2 = strShurui

    ElseIf 種類枚数 = "枚数" Then
        ガラス種類枚数2 = intGlass
        
    End If
    
End Property

Public Property Get テノナW() As Variant
    
    With inArgument
        If blnAWKansuMatch Then
            テノナW = dblFIVEorZERO(Eval(Replace(AWKansu.テノーナW関数AW, "DW", .inDW)))
        Else
            テノナW = dblFIVEorZERO(Eval(Replace(TateguKansu.テノーナW関数, "DW", .inDW)))
        End If
    End With
    
End Property

Public Property Get テノナW2() As Variant
    
    With inArgument
    
        テノナW2 = dblFIVEorZERO(Eval(Replace(TateguKansu.テノーナW関数2, "DW", .inDW)))
        
    End With
    
End Property

Public Property Get テノナW中板() As Variant
    
    With inArgument
        If blnAWKansuMatch Then
            'G7型のみ表示
            If .in検索品番 Like "*-####MF*" Then
                テノナW中板 = dblFIVEorZERO(Eval(Replace(AWKansu.中板テノーナW関数AW, "DW", .inDW)))
            Else
                テノナW中板 = 0
            End If
        Else
            テノナW中板 = dblFIVEorZERO(Eval(Replace(TateguKansu.中板テノーナW関数, "DW", .inDW)))
        End If
    End With
    
End Property

Public Property Get テノナW中板2() As Variant
    
    With inArgument
        If blnAWKansuMatch Then
            'G7型(スルー)のみ表示
            If IsG7_Flush(.in検索品番) Then
                テノナW中板2 = dblFIVEorZERO(Eval(Replace(AWKansu.中板テノーナW関数AW, "DW", .inDW)))
            Else
                テノナW中板2 = 0
            End If
        Else
            テノナW中板2 = dblFIVEorZERO(Eval(Replace(TateguKansu.中板テノーナW関数2, "DW", .inDW)))
        End If

    End With
    
End Property

Public Property Get W1() As Variant
    
    With inArgument
        If blnAWKansuMatch Then
            W1 = dblFIVEorZERO(Eval(Replace(AWKansu.パネルソーW1関数AW, "DW", .inDW)))
        Else
            W1 = dblFIVEorZERO(Eval(Replace(TateguKansu.パネルソーW1関数, "DW", .inDW)))
        End If
    End With
    
End Property

Public Property Get W2() As Variant
    
    With inArgument
        If blnAWKansuMatch Then
            W2 = dblFIVEorZERO(Eval(Replace(AWKansu.パネルソーW2関数AW, "DW", .inDW)))
        Else
            W2 = dblFIVEorZERO(Eval(Replace(TateguKansu.パネルソーW2関数, "DW", .inDW)))
        End If
    End With
    
End Property

Public Property Get ガラスW() As Variant
    
    With inArgument
        If blnAWKansuMatch Then
            If .in検索品番 Like "*-####MF*" Then
                ガラスW = dblFIVEorZERO(Eval(Replace(AWKansu.鏡W関数AW, "DW", .inDW)))
            Else
                ガラスW = dblFIVEorZERO(Eval(Replace(AWKansu.ガラスW関数AW, "DW", .inDW)))
            End If
        ElseIf IsTerraceGlass(.in検索品番) Then
            
            If .in検索品番 Like "Y?B*-####*" Then     ' YG-5型の場合
                ガラスW = .inDW - 247
            ElseIf .in検索品番 Like "Y?C*-####*" Then ' YG-6型の場合
                Select Case .inDW
                    Case 670 To 740
                        ガラスW = 376
                    Case 740.5 To 900
                        ガラスW = 447
                    Case Else
                        ガラスW = 0
                End Select
            Else
                ガラスW = 0
            End If
        Else
            ガラスW = dblFIVEorZERO(Eval(Replace(TateguKansu.ガラスW関数, "DW", .inDW)))
        End If
    End With
    
End Property

Public Property Get テノナH() As Variant
    
    With inArgument

        テノナH = dblFIVEorZERO(Eval(Replace(TateguKansu.テノーナH関数, "DH", .inDH)))

    End With
    
End Property

Public Property Get テノナH2() As Variant
    
    With inArgument

        テノナH2 = dblFIVEorZERO(Eval(Replace(TateguKansu.テノーナH関数2, "DH", .inDH)))

    End With
    
End Property

Public Property Get ガラスH() As Variant
    
    With inArgument
        If IsTerraceGlass(.in検索品番) Then
            If .in検索品番 Like "Y?B*-####*" Then     ' YG-5型の場合
                Select Case .inCH
                    Case 1800 To 2100
                        ガラスH = .inDH - 1197
                    Case 2100.5 To 2400
                        ガラスH = .inDH - 1297
                    Case Else
                        ガラスH = 0
                End Select
            
            ElseIf .in検索品番 Like "Y?C*-####*" Then ' YG-6型の場合
                Select Case .inDW
                    Case 670 To 740
                        ガラスH = 334
                    Case 740.5 To 900
                        ガラスH = 397
                    Case Else
                        ガラスH = 0
                End Select
            
            Else
                ガラスH = 0
                
            End If
        Else
            ガラスH = dblFIVEorZERO(Eval(Replace(TateguKansu.ガラスH関数, "DH", .inDH)))
        
        End If
        
    End With
    
End Property

Public Property Get テノナH中板() As Variant
'2.1.0 ADD テラス計算式追加

    With inArgument
        If blnSF4KansuMatch Then
            If (.in検索品番 Like "F?WD?-????P*" Or .in検索品番 Like "F?WP?-????P*" Or .in検索品番 Like "F?WD??-????P*" Or .in検索品番 Like "F?WP??-????P*") And .inCH < 1881 Then
                テノナH中板 = dblFIVEorZERO(.inDH - 476)
            Else
                テノナH中板 = dblFIVEorZERO(Eval(Replace(SF4Kansu.中板テノーナH関数, "DH", .inDH)))
            End If
            
        ElseIf IsTerraceKamachi(.in検索品番) Then
            If .inCH <= 2100 And .inCH >= 1800 Then
                テノナH中板 = 723
                
            ElseIf .inCH <= 2400 And .inCH > 2100 Then
                テノナH中板 = 823
                
            End If
        Else
            テノナH中板 = dblFIVEorZERO(Eval(Replace(TateguKansu.中板テノーナH関数, "DH", .inDH)))
        End If
    End With
    
End Property

Public Property Get テノナH中板2() As Variant
'2.1.0 ADD テラス計算式追加

    With inArgument
        If IsTerraceKamachi(.in検索品番) And IsTerraceGlass(.in検索品番) = False Then
            If .inCH <= 2100 And .inCH >= 1800 Then
                テノナH中板2 = dblFIVEorZERO(Eval(.inDH - 1197))
                
            ElseIf .inCH <= 2400 And .inCH > 2100 Then
                テノナH中板2 = dblFIVEorZERO(Eval(.inDH - 1297))
                
            End If
        Else
            テノナH中板2 = dblFIVEorZERO(Eval(Replace(TateguKansu.中板テノーナH関数2, "DH", .inDH)))
        End If
    End With
    
End Property

Public Property Get CU_AS_AW表面材W() As Variant
    
    With inArgument
    
        If blnAWKansuMatch Then
            CU_AS_AW表面材W = dblFIVEorZERO(Eval(Replace(AWKansu.CU_AS_AW表面材W関数AW, "DW", .inDW)))
        Else
            CU_AS_AW表面材W = dblFIVEorZERO(Eval(Replace(TateguKansu.CU_AS_AW表面材W関数, "DW", .inDW)))
        End If
        
    End With
    
End Property

Public Property Get CU_AS_AW表面材W2() As Variant
    
    With inArgument
        
        If Not blnAWKansuMatch Then
            CU_AS_AW表面材W2 = dblFIVEorZERO(Eval(Replace(TateguKansu.CU_AS_AW表面材W関数2, "DW", .inDW)))
        Else
            CU_AS_AW表面材W2 = 0
        End If
        
    End With
    
End Property

Public Property Get AS_AW中板W() As Variant
    
    With inArgument
    
        'AS_AW中板W = dblFIVEorZERO(Eval(Replace(TateguKansu.AS_AW中板W関数, "DW", .inDW)))
        If IsTerrace(.in検索品番) Then
        
            'YF1/YG6型
            If .in検索品番 Like "Y?C*" Then
            
                If .inDW <= 594 And .inDW >= 551 Then
                    AS_AW中板W = dblFIVEorZERO(.inDW / 2 + 611)
                ElseIf .inDW <= 696 And .inDW >= 594.5 Then
                    AS_AW中板W = dblFIVEorZERO(.inDW / 2 + 560)
                ElseIf .inDW <= 798 And .inDW >= 696.5 Then
                    AS_AW中板W = dblFIVEorZERO(.inDW / 2 + 509)
                ElseIf .inDW <= 900 And .inDW >= 798.5 Then
                    AS_AW中板W = dblFIVEorZERO(.inDW / 2 + 458)
                End If
            'YF5/YG5型
            ElseIf .in検索品番 Like "Y?B*" Then
            
                If .inDW <= 641 And .inDW >= 551 Then
                    AS_AW中板W = dblFIVEorZERO(.inDW / 2 + 587.5)
                ElseIf .inDW <= 743 And .inDW >= 641.5 Then
                    AS_AW中板W = dblFIVEorZERO(.inDW / 2 + 536.5)
                ElseIf .inDW <= 845 And .inDW >= 743.5 Then
                    AS_AW中板W = dblFIVEorZERO(.inDW / 2 + 485.5)
                ElseIf .inDW <= 900 And .inDW >= 845.5 Then
                    AS_AW中板W = dblFIVEorZERO(.inDW / 2 + 434.5)
                End If
            Else
                AS_AW中板W = 0
            End If
        Else
            AS_AW中板W = 0
        End If
        
    End With
    
End Property

Public Property Get AS_AW中板W1() As Variant
    
    With inArgument
    
        If blnAWKansuMatch Then
            If IsG7_Flush(.in検索品番) Then
                AS_AW中板W1 = dblFIVEorZERO(Eval(Replace(AWKansu.AS_AW中板W関数AW, "DW", .inDW)))
            Else
                AS_AW中板W1 = 0
            End If
        Else
            AS_AW中板W1 = dblFIVEorZERO(Eval(Replace(TateguKansu.AS_AW中板W関数1, "DW", .inDW)))
        End If

    End With
    
End Property

Public Property Get AS_AW中板W2() As Variant
    
    With inArgument
        
        If blnAWKansuMatch Then
            If IsG7_Flush(.in検索品番) Then
                AS_AW中板W2 = dblFIVEorZERO(Eval(Replace(AWKansu.AS_AW中板W関数AW, "DW", .inDW)))
            Else
                AS_AW中板W2 = 0
            End If
        Else
            AS_AW中板W2 = dblFIVEorZERO(Eval(Replace(TateguKansu.AS_AW中板W関数2, "DW", .inDW)))
        End If

    End With
    
End Property

Public Property Get AS_AWミラー表W() As Variant
    
    With inArgument
    
        AS_AWミラー表W = dblFIVEorZERO(Eval(Replace(TateguKansu.ミラー表W, "DW", .inDW)))

    End With
    
End Property

Public Property Get AS_AWミラー裏W() As Variant
    
    With inArgument
    
        AS_AWミラー裏W = dblFIVEorZERO(Eval(Replace(TateguKansu.ミラー裏W, "DW", .inDW)))

    End With
    
End Property

Public Property Get CU_AS_AW表面材H() As Variant
    
    With inArgument
    
        If (TateguKansu.CU_AS_AW表面材H関数 <> "0") And (.in検索品番 Like "F?DD*" Or .in検索品番 Like "F?DP*" Or .in検索品番 Like "F?DC*" Or .in検索品番 Like "F?DV*") And right(.in個別Spec, 4) > "1110" Then
            CU_AS_AW表面材H = dblFIVEorZERO(Eval(Replace(TateguKansu.CU_AS_AW表面材H関数, "DH", .inDH)) - 3)
        'F6型でカワジュン製カセットの場合は以下数値に差し替え
        ElseIf (TateguKansu.CU_AS_AW表面材H関数 <> "0") And .in検索品番 Like "??D*-####F?*-*" And .in個別Spec >= "BRD1507" And (.in施錠 Like "CN*" Or .in施錠 Like "CO*" Or .in施錠 Like "DN*" Or .in施錠 Like "DO*" Or .in施錠 Like "CP*" Or .in施錠 Like "CQ*" Or .in施錠 Like "CR*") Then
            CU_AS_AW表面材H = dblFIVEorZERO(.inDH - 1047)
        Else
            CU_AS_AW表面材H = dblFIVEorZERO(Eval(Replace(TateguKansu.CU_AS_AW表面材H関数, "DH", .inDH)))
        End If

    End With
    
End Property

Public Property Get 上パネルW() As Variant
    
    With inArgument
    
        上パネルW = dblFIVEorZERO(Eval(Replace(TateguKansu.ST表面材上W関数, "DW", .inDW)))

    End With
    
End Property

Public Property Get アクセントライン() As String
    
    With inArgument
    
        If .in検索品番 Like "??D*-####FT*-*" Then
            アクセントライン = "ティネオ"
        ElseIf .in検索品番 Like "??D*-####FR*-*" Then
            アクセントライン = "ローズウッド"
        ElseIf .in検索品番 Like "??D*-####FH*-*" Then
            アクセントライン = "ヘアライン"
        Else
            アクセントライン = ""
        End If

    End With
    
End Property

Public Property Get CU_AS_AW表面材H2() As Variant
    
    With inArgument
    
        If (TateguKansu.CU_AS_AW表面材H関数2 <> "0") And (.in検索品番 Like "F?DD*" Or .in検索品番 Like "F?DP*" Or .in検索品番 Like "F?DC*" Or .in検索品番 Like "F?DV*") And (right(.in個別Spec, 4)) > "1110" Then
            CU_AS_AW表面材H2 = dblFIVEorZERO(Eval(Replace(TateguKansu.CU_AS_AW表面材H関数2, "DH", .inDH)) - 3)
        Else
            CU_AS_AW表面材H2 = dblFIVEorZERO(Eval(Replace(TateguKansu.CU_AS_AW表面材H関数2, "DH", .inDH)))
        End If
        
        'F6型でカワジュン製カセットの場合は以下数値に差し替え
        If .in検索品番 Like "??D*-####F?*-*" And .in個別Spec >= "BRD1507" And (.in施錠 Like "CN*" Or .in施錠 Like "CO*" Or .in施錠 Like "DN*" Or .in施錠 Like "DO*" Or .in施錠 Like "CP*" Or .in施錠 Like "CQ*" Or .in施錠 Like "CR*") Then
            CU_AS_AW表面材H2 = 1013
        End If
        
    End With
    
End Property

Public Property Get CU_AS_AW表面材枚数() As Variant
    
    With inArgument

        CU_AS_AW表面材枚数 = TateguKansu.CU_AS_AW枚数 * .inセット数
        
    End With
    
End Property

Public Property Get CU_AS_AW表面材枚数2() As Variant
    
    With inArgument

        CU_AS_AW表面材枚数2 = TateguKansu.CU_AS_AW枚数2 * .inセット数
        
    End With
    
End Property

Public Property Get AS_AW中板H() As Variant
    
    With inArgument
    
        AS_AW中板H = dblFIVEorZERO(Eval(Replace(TateguKansu.AS_AW中板H関数, "DH", .inDH)))

    End With
    
End Property

Public Property Get AS_AW中板枚数() As Variant
    
    With inArgument

        AS_AW中板枚数 = TateguKansu.AS_AW中板枚数 * .inセット数
        
    End With
    
End Property

Public Property Get AS_AW中板H1() As Variant
'2.1.0 ADD テラス計算式追加

    With inArgument
        
        If IsTerraceKamachi(.in検索品番) Then
            If .inCH <= 2100 And .inCH >= 1800 Then
                テノナH中板 = 723
                
            ElseIf .inCH <= 2400 And .inCH > 2100 Then
                テノナH中板 = 823
                
            End If
        Else
            AS_AW中板H1 = dblFIVEorZERO(Eval(Replace(TateguKansu.AS_AW中板H関数1, "DH", .inDH)))
        End If
    End With
    
End Property

Public Property Get AS_AW中板枚数1() As Variant
    
    With inArgument

        AS_AW中板枚数1 = TateguKansu.AS_AW中板枚数1 * .inセット数
        
    End With
    
End Property

Public Property Get AS_AW中板H2() As Variant
'2.1.0 ADD テラス計算式追加
    
    With inArgument
        If IsTerraceKamachi(.in検索品番) And IsTerraceGlass(.in検索品番) = False Then
            If .inCH <= 2100 And .inCH >= 1800 Then
                テノナH中板2 = dblFIVEorZERO(Eval(.inDH - 1197))
                
            ElseIf .inCH <= 2400 And .inCH > 2100 Then
                テノナH中板2 = dblFIVEorZERO(Eval(.inDH - 1297))
                
            End If
        Else
            AS_AW中板H2 = dblFIVEorZERO(Eval(Replace(TateguKansu.AS_AW中板H関数2, "DH", .inDH)))
        End If
    End With
    
End Property

Public Property Get AS_AW中板枚数2() As Variant
    
    With inArgument

        AS_AW中板枚数2 = TateguKansu.AS_AW中板枚数2 * .inセット数
        
    End With
    
End Property

Public Property Get ミラー表H() As Variant
    
    With inArgument
    
        ミラー表H = dblFIVEorZERO(Eval(Replace(TateguKansu.ミラー表H, "DH", .inDH)))

    End With
    
End Property

Public Property Get ミラー表枚数() As Variant
    
    With inArgument

        ミラー表枚数 = TateguKansu.ミラー表枚数 * .inセット数
        
    End With
    
End Property


Public Property Get ミラー裏H() As Variant
    
    With inArgument
    
        ミラー裏H = dblFIVEorZERO(Eval(Replace(TateguKansu.ミラー裏H, "DH", .inDH)))

    End With
    
End Property

Public Property Get ミラー裏枚数() As Variant
    
    With inArgument

        ミラー裏枚数 = TateguKansu.ミラー裏枚数 * .inセット数
        
    End With
    
End Property

Public Property Get メタルW() As Variant
    
    With inArgument

        If .in検索品番 Like "F?D*" Or .in検索品番 Like "P?D*" Then
            メタルW = dblFIVEorZERO(Eval(Replace(TateguKansu.CU_AS_AW表面材W関数, "DW", .inDW)))
        Else
            メタルW = 0
        End If
        
    End With
    
End Property

Public Property Get メタル枚数() As Variant
    
    With inArgument

        If .in検索品番 Like "F?D*" Or .in検索品番 Like "P?D*" Then
            メタル枚数 = TateguKansu.CU_AS_AW枚数 * .inセット数
        Else
            メタル枚数 = 0
        End If
        
    End With
    
End Property

Public Property Get 上框種類() As String
    
    With inArgument

        上框種類 = TateguKansu.上框種類
        
    End With
    
End Property

Public Property Get 上框本数() As Variant
    
    With inArgument
    
        If blnSF4KansuMatch Then
            If .in検索品番 Like "F?WD?-????P*" Or .in検索品番 Like "F?WP?-????P*" Or .in検索品番 Like "F?WD??-????P*" Or .in検索品番 Like "F?WP??-????P*" Then
                If .inCH < 1881 Then
                    上框本数 = 2 * .in枚数
                Else
                    上框本数 = 4 * .in枚数
                End If
            Else
                上框本数 = TateguKansu.上框長さ個数 * .inセット数
            End If
        Else
            上框本数 = TateguKansu.上框長さ個数 * .inセット数
        End If

    End With
    
End Property

Public Property Get 下框種類() As String
    
    With inArgument

        下框種類 = TateguKansu.下框種類
        
    End With
    
End Property

Public Property Get 下框本数() As Variant
    
    With inArgument
    
        下框本数 = TateguKansu.下框長さ個数 * .inセット数
        
    End With
    
End Property

Public Property Get 中框種類() As String
'2.1.0 ADD

    中框種類 = TateguKansu.中框種類

End Property

Public Property Get 上框高さ() As Variant
'2.1.0 ADD

    上框高さ = dblFIVEorZERO(TateguKansu.上框高さ)

End Property

Public Property Get 下框高さ() As Variant
'2.1.0 ADD

    下框高さ = dblFIVEorZERO(TateguKansu.下框高さ)

End Property

Public Property Get H1() As Variant
    
    With inArgument
        If blnSF4KansuMatch Then
            H1 = dblFIVEorZERO(Eval(Replace(SF4Kansu.パネルソーH1関数, "DH", .inDH)))
        Else
            H1 = dblFIVEorZERO(Eval(Replace(TateguKansu.パネルソーH1関数, "DH", .inDH)))
        End If
    End With
    
End Property

Public Property Get H2() As Variant
    
    With inArgument
        If blnSF4KansuMatch Then
            H2 = dblFIVEorZERO(Eval(Replace(SF4Kansu.パネルソーH2関数, "DH", .inDH)))
        Else
            H2 = dblFIVEorZERO(Eval(Replace(TateguKansu.パネルソーH2関数, "DH", .inDH)))
        End If
    End With
    
End Property

Public Property Get H3() As Variant
    
    With inArgument
        If blnSF4KansuMatch Then
            H3 = dblFIVEorZERO(Eval(Replace(SF4Kansu.パネルソーH3関数, "DH", .inDH)))
        Else
            H3 = dblFIVEorZERO(Eval(Replace(TateguKansu.パネルソーH3関数, "DH", .inDH)))
        End If
    End With
    
End Property

Public Property Get ガラス厚1() As Variant
    
    ガラス厚1 = TateguKansu.ガラス厚み1
    
End Property

Public Property Get ガラスW1() As Variant
    
    With inArgument

        ガラスW1 = dblFIVEorZERO(Eval(Replace(TateguKansu.ガラスW関数1, "DW", .inDW)))

    End With
    
End Property

Public Property Get ガラスH1() As Variant
    
    With inArgument
        If .in検索品番 Like "F?WD?-????G*" Or .in検索品番 Like "F?WP?-????G*" Or .in検索品番 Like "F?WD??-????G*" Or .in検索品番 Like "F?WP??-????G*" Then
            If .inCH > 2400 Then
                ガラスH1 = 558
            Else
                ガラスH1 = 448
            End If
        Else
            ガラスH1 = dblFIVEorZERO(Eval(Replace(TateguKansu.ガラスH関数1, "DH", .inDH)))
        End If
    End With
    
End Property

Public Property Get 上框長さ() As Variant
    
    With inArgument
        If blnAWKansuMatch Then
            If Not IsG7_Flush(.in検索品番) Then
                上框長さ = dblFIVEorZERO(Eval(Replace(AWKansu.上框長さ関数AW, "DW", .inDW)))
            Else
                上框長さ = 0
            End If
        Else
            上框長さ = dblFIVEorZERO(Eval(Replace(TateguKansu.上框長さ関数, "DW", .inDW)))
        End If
    End With
    
End Property

Public Property Get 下框長さ() As Variant
    
    With inArgument
        If blnAWKansuMatch Then
            If Not IsG7_Flush(.in検索品番) Then
                下框長さ = dblFIVEorZERO(Eval(Replace(AWKansu.下框長さ関数AW, "DW", .inDW)))
            Else
                下框長さ = 0
            End If
        ElseIf .in検索品番 Like "F?WD?-????G*" Or .in検索品番 Like "F?WP?-????G*" Or .in検索品番 Like "F?WD??-????G*" Or .in検索品番 Like "F?WP??-????G*" Then
            下框長さ = dblFIVEorZERO(Eval(Replace(TateguKansu.上框長さ関数, "DW", .inDW)))
        Else
            下框長さ = dblFIVEorZERO(Eval(Replace(TateguKansu.下框長さ関数, "DW", .inDW)))
        End If
    End With
    
End Property

Public Property Get 上パネルH() As Variant
    
    With inArgument
        
        If blnSTREETKansuMatch Then
            上パネルH = dblFIVEorZERO(Eval(Replace(STREETKansu.上H関数, "DH", .inDH)))
        ElseIf blnSF2KansuMatch Then
            If IsHirakido(.in検索品番) Then
                上パネルH = dblFIVEorZERO(Eval(Replace(SF2Kansu.上H関数, "DH", .inDH)))
            Else
                上パネルH = dblFIVEorZERO(Eval(Replace(SF2Kansu.上H関数_引戸, "DH", .inDH)))
            End If
        ElseIf blnSF4KansuMatch Then
            上パネルH = dblFIVEorZERO(Eval(Replace(SF4Kansu.上H関数, "DH", .inDH)))
        Else
            上パネルH = 0
        End If
        
    End With
    
End Property

Public Property Get 上パネル枚数() As Variant
    
    With inArgument
        
        If blnSTREETKansuMatch Then
            上パネル枚数 = STREETKansu.上枚数 * .in枚数
        ElseIf blnSF2KansuMatch Then
            上パネル枚数 = SF2Kansu.上枚数 * .in枚数
        ElseIf blnSF4KansuMatch Then
            上パネル枚数 = SF4Kansu.上枚数 * .in枚数
        Else
            上パネル枚数 = 0
        End If
        
    End With
    
End Property

Public Property Get 上HFLG() As Boolean
    
    With inArgument
        
        If blnSTREETKansuMatch Then
            上HFLG = STREETKansu.上HFLG
        ElseIf blnSF2KansuMatch Then
            上HFLG = SF2Kansu.上HFLG
        ElseIf blnSF4KansuMatch Then
            上HFLG = SF4Kansu.上HFLG
        Else
            上HFLG = False
        End If
        
    End With
    
End Property

Public Property Get 中パネルH() As Variant
    
    With inArgument
        
        If blnSTREETKansuMatch Then
            中パネルH = STREETKansu.中H
        ElseIf blnSF2KansuMatch Then
            中パネルH = SF2Kansu.中H
        ElseIf blnSF4KansuMatch Then
            中パネルH = SF4Kansu.中H
        Else
            中パネルH = 0
        End If
        
    End With
    
End Property


Public Property Get 中パネル2H() As Variant
    
    With inArgument
        
        If blnSF2KansuMatch Then
            中パネル2H = SF2Kansu.中H1
        ElseIf blnSF4KansuMatch Then
            中パネル2H = SF4Kansu.中H1
        Else
            中パネル2H = 0
        End If
        
    End With
    
End Property

Public Property Get 中パネル枚数() As Variant
    
    With inArgument
        
        If blnSTREETKansuMatch Then
            中パネル枚数 = STREETKansu.中枚数 * .in枚数
        ElseIf blnSF2KansuMatch Then
            中パネル枚数 = SF2Kansu.中枚数 * .in枚数
        ElseIf blnSF4KansuMatch Then
            中パネル枚数 = SF4Kansu.中枚数 * .in枚数
        Else
            中パネル枚数 = 0
        End If
        
    End With
    
End Property

Public Property Get 中パネル2枚数() As Variant
    
    With inArgument
        
        If blnSF2KansuMatch Then
            中パネル2枚数 = SF2Kansu.中枚数1 * .in枚数
        ElseIf blnSF4KansuMatch Then
            中パネル2枚数 = SF4Kansu.中枚数1 * .in枚数
        Else
            中パネル2枚数 = 0
        End If
        
    End With
    
End Property

Public Property Get 中HFLG() As Boolean
    
    With inArgument
        
        If blnSTREETKansuMatch Then
            中HFLG = STREETKansu.中HFLG
        ElseIf blnSF2KansuMatch Then
            中HFLG = SF2Kansu.中HFLG
        ElseIf blnSF4KansuMatch Then
            中HFLG = SF4Kansu.中HFLG
        Else
            中HFLG = False
        End If
        
    End With
    
End Property

Public Property Get 下パネルH() As Variant
    
    With inArgument
        
        If blnSTREETKansuMatch Then
            下パネルH = dblFIVEorZERO(Eval(Replace(STREETKansu.下H関数, "DH", .inDH)))
        ElseIf blnSF2KansuMatch Then
            If IsHirakido(.in検索品番) Then
                下パネルH = dblFIVEorZERO(Eval(Replace(SF2Kansu.下H関数, "DH", .inDH)))
            Else
                下パネルH = dblFIVEorZERO(Eval(Replace(SF2Kansu.下H関数_引戸, "DH", .inDH)))
            End If
        ElseIf blnSF4KansuMatch Then
            下パネルH = SF4Kansu.下H関数
        Else
            下パネルH = 0
        End If
        
    End With
    
End Property

Public Property Get 下パネル枚数() As Variant
    
    With inArgument
        
        If blnSTREETKansuMatch Then
            下パネル枚数 = STREETKansu.下枚数 * .in枚数
        ElseIf blnSF2KansuMatch Then
            下パネル枚数 = SF2Kansu.下枚数 * .in枚数
        ElseIf blnSF4KansuMatch Then
            下パネル枚数 = SF4Kansu.下枚数 * .in枚数
        Else
            下パネル枚数 = 0
        End If
        
    End With
    
End Property

Public Property Get 下HFLG() As Boolean
    
    With inArgument
        
        If blnSTREETKansuMatch Then
            下HFLG = STREETKansu.下HFLG
        ElseIf blnSF2KansuMatch Then
            下HFLG = SF2Kansu.下HFLG
        ElseIf blnSF4KansuMatch Then
            下HFLG = SF4Kansu.下HFLG
        Else
            下HFLG = False
        End If
        
    End With
    
End Property

Public Property Get Terrace_中框ホゾ加工センター() As Variant
    
    
    Terrace_中框ホゾ加工センター = 0
  
    With inArgument

        Select Case .inCH
        
            Case 2100.5 To 2400
                
                Terrace_中框ホゾ加工センター = 1095
            
            Case 1800 To 2100
            
                Terrace_中框ホゾ加工センター = 995
            
        End Select

    End With
    
End Property

Public Property Get ダボピッチ1() As Variant

    With inArgument
        
        If blnSF4KansuMatch Then
            ダボピッチ1 = SF4Kansu.ダボピッチ1
        ElseIf blnDaboPitchKansuMatch Then
            If IsCaro_Panel(.in検索品番) Then
                ダボピッチ1 = dblFIVEorZERO(.inDH - Nz(varDaboPitch(1), 0))
            Else
                ダボピッチ1 = dblFIVEorZERO(Nz(varDaboPitch(1), 0))
            End If
        Else
            ダボピッチ1 = 0
        End If
        
    End With
    
End Property

Public Property Get ダボピッチ2() As Variant

    With inArgument
        
        If blnSF4KansuMatch Then
            ダボピッチ2 = SF4Kansu.ダボピッチ2
        ElseIf blnDaboPitchKansuMatch Then
            If IsCaro_Panel(.in検索品番) Then
                ダボピッチ2 = dblFIVEorZERO(.inDH - Nz(varDaboPitch(2), 0))
            Else
                ダボピッチ2 = dblFIVEorZERO(Nz(varDaboPitch(2), 0))
            End If
        Else
            ダボピッチ2 = 0
        End If
        
    End With
    
End Property

Public Property Get ダボピッチ3() As Variant

    With inArgument
        
        If blnSF4KansuMatch Then
            ダボピッチ3 = SF4Kansu.ダボピッチ3
        ElseIf blnDaboPitchKansuMatch Then
            ダボピッチ3 = dblFIVEorZERO(Eval(Replace(Nz(varDaboPitch(3), 0), "DH", .inDH)))
        Else
            ダボピッチ3 = 0
        End If
        
    End With
    
End Property

Public Property Get ダボピッチ4() As Variant

    With inArgument
        
        If blnSF4KansuMatch Then
            ダボピッチ4 = SF4Kansu.ダボピッチ4
        ElseIf blnDaboPitchKansuMatch Then
            ダボピッチ4 = dblFIVEorZERO(Eval(Replace(Nz(varDaboPitch(4), 0), "DH", .inDH)))
        Else
            ダボピッチ4 = 0
        End If
        
    End With
    
End Property

Public Property Get ダボピッチ5() As Variant

    With inArgument
        
        If blnSF4KansuMatch Then
            ダボピッチ5 = SF4Kansu.ダボピッチ5
        ElseIf blnDaboPitchKansuMatch Then
            ダボピッチ5 = dblFIVEorZERO(Eval(Replace(Nz(varDaboPitch(5), 0), "DH", .inDH)))
        Else
            ダボピッチ5 = 0
        End If
        
    End With
    
End Property

Public Property Get ダボピッチ6() As Variant

    With inArgument
        
        If blnSF4KansuMatch Then
            ダボピッチ6 = SF4Kansu.ダボピッチ6
        ElseIf blnDaboPitchKansuMatch Then
            ダボピッチ6 = dblFIVEorZERO(Eval(Replace(Nz(varDaboPitch(6), 0), "DH", .inDH)))
        Else
            ダボピッチ6 = 0
        End If
        
    End With
    
End Property

Public Property Get ダボピッチ7() As Variant

    With inArgument
        
        If blnSF4KansuMatch Then
            ダボピッチ7 = SF4Kansu.ダボピッチ7
        ElseIf blnDaboPitchKansuMatch Then
            ダボピッチ7 = dblFIVEorZERO(Eval(Replace(Nz(varDaboPitch(7), 0), "DH", .inDH)))
        Else
            ダボピッチ7 = 0
        End If
        
    End With
    
End Property

Public Property Get ダボピッチ8() As Variant

    With inArgument
        
        If blnSF4KansuMatch Then
            ダボピッチ8 = SF4Kansu.ダボピッチ8
        ElseIf blnDaboPitchKansuMatch Then
            ダボピッチ8 = dblFIVEorZERO(Eval(Replace(Nz(varDaboPitch(8), 0), "DH", .inDH)))
        Else
            ダボピッチ8 = 0
        End If
        
    End With
    
End Property

Public Property Get ダボピッチ9() As Variant

    With inArgument
        
        If blnSF4KansuMatch Then
            ダボピッチ9 = SF4Kansu.ダボピッチ9
        ElseIf blnDaboPitchKansuMatch Then
            ダボピッチ9 = dblFIVEorZERO(Eval(Replace(Nz(varDaboPitch(9), 0), "DH", .inDH)))
        Else
            ダボピッチ9 = 0
        End If
        
    End With
    
End Property

Public Property Get ダボピッチ10() As Variant

    With inArgument
        
        If blnSF4KansuMatch Then
            ダボピッチ10 = SF4Kansu.ダボピッチ10
        ElseIf blnDaboPitchKansuMatch Then
            ダボピッチ10 = dblFIVEorZERO(Eval(Replace(Nz(varDaboPitch(10), 0), "DH", .inDH)))
        Else
            ダボピッチ10 = 0
        End If
        
    End With
    
End Property

Public Property Get 金物加工図() As Variant
    
    金物加工図 = Null
    
    With inArgument
        '特注
        If .in施錠 Like "*特*" Then
    
            If .in明り窓 <> "-" Then
                金物加工図 = "FKT-17"
            End If
            
            Exit Property
        End If
        
        '特注錠以外
        If (fncbol_Handle_EStyle(.in施錠, CStr(.in個別Spec)) And .in施錠 Like "*C") And (.in検索品番 Like "*DK-*" Or .in検索品番 Like "*DKS-*") Then
                If TateguKansu.基本図 Like "H*" Then
                    金物加工図 = "HKT-14"
                Else
                    If right(.in個別Spec, 4) < "1011" Then
                        金物加工図 = "FKT-21"
                    Else
                        金物加工図 = "FKT-55"
                    End If
                End If
        ElseIf (fncbol_Handle_EStyle(.in施錠, CStr(.in個別Spec))) And (.in検索品番 Like "*DK-*" Or .in検索品番 Like "*DKS-*") Then
            If TateguKansu.基本図 Like "H*" Then
                金物加工図 = "HKT-13"
            Else
                If right(.in個別Spec, 4) < "1011" Then
                    金物加工図 = "FKT-20"
                Else
                    金物加工図 = "FKT-55"
                End If
            End If
        ElseIf right(.in個別Spec, 4) < "1011" And ((fncbol_Handle_EStyle(.in施錠, CStr(.in個別Spec)) And .in施錠 Like "*C")) Then
            If TateguKansu.基本図 Like "H*" Then
                金物加工図 = "HKT-12"
            Else
                金物加工図 = "FKT-19"
            End If
        ElseIf right(.in個別Spec, 4) < "1011" And (fncbol_Handle_EStyle(.in施錠, CStr(.in個別Spec))) Then
            If TateguKansu.基本図 Like "H*" Then
                金物加工図 = "HKT-11"
            Else
                金物加工図 = "FKT-18"
            End If
        ElseIf .in施錠 = "ZC" And .in検索品番 Like "*DP-*" Then
            If .in検索品番 Like "*F?OD*" Or .in検索品番 Like "*F?GD*" Then
                金物加工図 = "FKT-52"
            Else
                金物加工図 = "FKT-32"
            End If
        ElseIf .in施錠 = "ZC" Then
            If .in検索品番 Like "*F?OD*" Or .in検索品番 Like "*F?GD*" Then
                金物加工図 = "FKT-50"
            Else
                金物加工図 = "FKT-27"
            End If
        Else
            If blnZumenKansuMatch Then 金物加工図 = TateguZumen.加工図
        End If
        
        If .in明り窓 <> "-" Then
            If Not IsNull(金物加工図) Then
                金物加工図 = 金物加工図 & "_17"
            Else
                金物加工図 = "FKT-17"
            End If
        End If
        
    End With
                    
End Property

Public Property Get 建具組立図() As Variant
    
    建具組立図 = Null
    
    With inArgument
        '特注錠
        If .in施錠 Like "*特*" Then
            If .in検索品番 Like "*F?FD*-??21G*" Or .in検索品番 Like "*F?FP*-??21G*" Then
            
                建具組立図 = "KK-21"
            Else
            
                 建具組立図 = TateguKansu.組立図
            End If
            
            Exit Property
            
        End If
        
        '特注錠以外
        If .in検索品番 Like "*F?FD*-??21G*" Or .in検索品番 Like "*F?FP*-??21G*" Then
        
            建具組立図 = "KK-21"
            
        ElseIf (.in検索品番 Like "*F?DD*" Or .in検索品番 Like "*F?DP*" Or .in検索品番 Like "*F?DC*") And right(.in個別Spec, 4) > 1110 Then
        
            建具組立図 = "KK-35"
            
        Else
            If blnZumenKansuMatch Then 建具組立図 = TateguZumen.組立図
            
        End If
        
     End With
     
End Property

Public Property Get ブランド品番() As Variant
    
    ブランド品番 = inArgument.inブランド品番
    
End Property

Public Property Get OEM品番() As Variant

    With inArgument
    
        If .inブランド品番 <> .in受注品番 Then
            OEM品番 = inArgument.in受注品番
        Else
            OEM品番 = ""
        End If
        
    End With
    
End Property

Public Property Get 建具形状() As Variant
'0:Flush
'1:スルーガラス
'2:F框
'3:框

    With inArgument
        
        If Nz(Me.W1, 0) > 0 And Nz(Me.W2, 0) > 0 Then
            If .in検索品番 Like "??B*" Then
                建具形状 = 2
            Else
                建具形状 = 1
            End If
        Else
            建具形状 = 0
        End If
        
    End With
    
End Property


Public Property Get 丁番上位置() As Variant


    丁番上位置 = Null

    With inArgument
        If IsHidden_Hinge(.in検索品番) Then
            丁番上位置 = 200
            
        ElseIf IsCloset_Hikichigai(.in検索品番) Then
            丁番上位置 = Null
        End If
        
    End With
    
End Property

Public Property Get 丁番下位置() As Variant


    丁番下位置 = Null

    With inArgument
        If IsHidden_Hinge(.in検索品番) Then
            丁番下位置 = 200
            
        ElseIf IsCloset_Hikichigai(.in検索品番) Then
            丁番下位置 = Null
        End If
        
    End With
    
End Property

Public Property Get 丁番中上位置() As Variant


    丁番中上位置 = Null

    With inArgument
        If IsCloset_Hikichigai(.in検索品番) Then
            丁番中上位置 = Null
        End If
        
    End With
    
End Property

Public Property Get 丁番中下位置() As Variant


    丁番中下位置 = Null

    With inArgument
        If IsCloset_Hikichigai(.in検索品番) Then
            丁番中下位置 = Null
        End If
        
    End With
    
End Property

Public Property Get 框材溝巾(ByVal 框材種類 As String) As String
'***********************************************************
'框材溝巾
'
'2.2.0 ADD
'***********************************************************
    
    Select Case 框材種類
        Case "GRU12"
            框材溝巾 = "8.4"
        
        Case "HRU12", "HRU15", "HRN15", "HRS22"
            框材溝巾 = "5.4"
            
        Case "HRG12"
            框材溝巾 = "6.4"
            
        Case "HRU15"
            框材溝巾 = ""
            
     End Select
            

End Property

Private Sub setADOConnection()
'***********************************************************
'ADOコネクション処理
'  共通変数が"SQLServer"の場合はSQLServerへ
'  それ以外はCurrentDBへコネクションを貼る
'
'***********************************************************
    
    On Error GoTo Err_setADOConnection
    
    If ConnectionDB = "SQLSERVER" Then
    
        'SQLServer ODBCの場合
        Set cnADO = New ADODB.Connection

        cnADO.Open "DSN=DB02_BRAND_MASTER;DATABASE=BRAND_MASTER;Trusted_connection=yes;"
        
        Exit Sub
    Else
        'CurrentDBの場合
        Set cnADO = CurrentProject.Connection
    End If
    
    Exit Sub
    
Err_setADOConnection:
    MsgBox Err.Description
End Sub
Private Sub RS_AllClose()
'***********************************************************
'レコードセットをすべて閉じる
'  品番が該当しない場合は開いていないのでエラーは無視する
'***********************************************************
    On Error Resume Next
    
    rsADO.Close
    rsADOAW.Close
    rsADOSF4.Close
    rsADOSF2.Close
    rsADOSTREET.Close
    rsADOKihonzu.Close
    rsADODaboPitch.Close
    
End Sub

Private Sub FLG_Init()
'***********************************************************
'検索FLG初期化
'る
'***********************************************************
    
    blnKansuMatch = False
    blnAWKansuMatch = False
    blnSF4KansuMatch = False
    blnZumenKansuMatch = False
    blnSF2KansuMatch = False
    blnDaboPitchKansuMatch = False
    
End Sub