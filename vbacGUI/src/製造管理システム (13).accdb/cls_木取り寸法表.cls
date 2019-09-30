Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'***********************************************************
'木取り寸法表関数
' BRD1908以降
'

'3.0.0 ADD
'***********************************************************

'*************************************
'接続先データベース定数              *
' "ACCESS"又は"SQLSERVER"を指定      *
Private Const ConnectionDB As String = "ACCESS"
'Private Const ConnectionDB As String = "SQLSERVER"
'*************************************

Private cnADO As ADODB.Connection
Private rsADO As ADODB.Recordset

Private blnKansuMatch As Boolean
Private blnAWKansuMatch As Boolean

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
    inハンドル引手種類 As String
    in建具品番_シリーズ As String
    in建具品番_建具スタイル As String
    in建具品番_開閉様式 As String
    in建具品番_パネルガラス As String
    in建具品番_ガラスレザー As String
    in建具品番_色 As String
End Type

Private inArgument As 引数格納

Private Type 木取り寸法関数
'*************************************
'木取り寸法関数マスタ構造体          *
'マスタのカラムが増えたらここにも登録*
'*************************************
    芯組詳細図番号 As String
    芯厚 As Double
    上下桟_長さ As String
    上下桟_本数 As Double
    額横1_長さ  As String
    額横1_本数  As Double
    額横2_長さ  As String
    額横2_本数  As Double
    額横_LVL30_長さ As String
    額横_LVL30_本数 As Double
    柱_長さ As String
    柱_本数 As Double
    柱_LVL45_長さ   As String
    柱_LVL45_本数   As Double
    額縦1_長さ  As String
    額縦1_本数  As Double
    額縦_LVL30_長さ As String
    額縦_LVL30_本数 As Double
    額縦2_長さ  As String
    額縦2_本数  As Double
    手掛け_種類 As String
    手掛け_長さ As String
    手掛け_本数 As Double
    袖1_長さ    As String
    袖1_本数    As Double
    袖2_長さ    As String
    袖2_本数    As Double
    ダボ下地_長さ   As String
    ダボ下地_本数   As Double
    カップ下地_長さ As String
    カップ下地_本数 As Double
    カップ下地_LVL30_長さ   As String
    カップ下地_LVL30_本数   As Double
    中板芯厚    As Double
    中板上下桟_長さ As String
    中板上下桟_本数 As Double
    中板額横_長さ   As String
    中板額横_本数   As Double
    中板柱_長さ As String
    中板柱_本数 As Double
    中板柱_LVL45_長さ   As String
    中板柱_LVL45_本数   As Double
End Type

Private KidoriKansu As 木取り寸法関数

Private Type AW関数
'*************************************
'AW関数構造体 　　　　               *
'  　　　　　　　　　　　　　　　　　*
'*************************************
    上下桟_長さ As String
    額横1_長さ  As String
    額横2_長さ  As String
    額横_LVL30_長さ  As String
    カップ下地_LVL30_長さ As String
    中板上下桟_LVL30_長さ As String
    中板上下桟_長さ As String
    中板額横_長さ   As String
End Type

Private AWKansu As AW関数

Private Type テープ誤差
'*************************************
'ローカンとの寸法誤差を登録          *
'　　　　　　　　　　　　　　　　　  *
'*************************************
    上下桟_長さ差異 As Double
    額横_長さ差異  As Double
    額横_LVL30_長さ差異  As Double
    中板上下桟_長さ差異 As Double
    中板額横_長さ差異  As Double
End Type

Private TapeGosa As テープ誤差

Public Function Bind(ByVal 品番 As String, ByVal 数量 As Integer, ByVal 枚数 As Integer, ByVal DW As Variant, ByVal DH As Variant, ByVal CH As Variant, ByVal 施錠 As Variant, ByVal 明り窓 As Variant, ByVal 個別Spec As Variant) As Boolean
'***********************************************************
'バインド処理（必須）
' プロパティを取得する場合は必ずこの関数を実行してから行う
'
'   戻り値:
'           True →引当成功
'           False→    失敗
'
'***********************************************************

    Dim typArgument As 引数格納 '初期化用。データは移送しないこと！
    
    Dim objTateguHinban As New cls_建具品番
    Dim blnEnd As Boolean
    
    
    On Error GoTo Err_Bind
    
    FLG_Init
    blnEnd = False
    
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
            Err.Raise 9999, , "SxL品番読替エラー  品番= " & .in受注品番
        Else
            .in検索品番 = Trim(Replace(.inブランド品番, "特 ", ""))
        End If
        
        If IsNull(個別Spec) Then
            Err.Raise 9999, , "SpecNullエラー  品番= " & .in受注品番
        End If
        
        If Not Is40mm(個別Spec) Then
            Err.Raise 9999, , "40mm建具のSpecではありません 品番= " & .in受注品番 & " 仕様= " & 個別Spec
        End If
        
        If objTateguHinban.IsTateguHinban(.inブランド品番) Then
            .in建具品番_シリーズ = objTateguHinban.シリーズ(.inブランド品番)
            .in建具品番_建具スタイル = objTateguHinban.建具スタイル(.inブランド品番)
            .in建具品番_開閉様式 = objTateguHinban.開閉様式(.inブランド品番)
            .in建具品番_パネルガラス = objTateguHinban.パネルガラス(.inブランド品番)
            .in建具品番_ガラスレザー = objTateguHinban.ガラスレザー種類(.inブランド品番)
            .in建具品番_色 = objTateguHinban.色(.inブランド品番)
        Else
            Err.Raise 9999, , "品番エラー  品番= " & .in受注品番
        End If
        
        .inハンドル引手種類 = fncstrHandleKigoFileName(.in検索品番, .in施錠, .in個別Spec)
        
    End With
    
    'テープ毎の修正誤差を登録
    Call bolfncGetTapeGosa
    
    If GetTateguKansu() Then
        blnEnd = True
    End If
    
    Bind = blnEnd
    
    GoTo Exit_Bind
    
Err_Bind:
    If Err.Number <> 9999 Then
        MsgBox Err.Description, vbCritical
    Else
        Debug.Print Err.Description
    End If
    
    inArgument = typArgument
    Bind = False
    
Exit_Bind:

    Set objTateguHinban = Nothing

End Function

Private Function GetTateguKansu() As Boolean
'***********************************************************
'関数テーブル引当
'
'***********************************************************

    '初期化用。データは移送しないこと！******
    Dim typKidoriKansuInit As 木取り寸法関数
    Dim typAWKansuInit As AW関数
    '****************************************
    
    Dim strSQL As String
    Dim strZumenNo As String
    
    Dim i As Integer
        
    On Error GoTo Err_GetTateguKansu
    
    KidoriKansu = typKidoriKansuInit
    AWKansu = typAWKansuInit
    
    strGetTateguKansuErrMsg = ""
    
    blnKansuMatch = False
    blnAWKansuMatch = False
    
    'オープン済みのレコードセットがある場合、一旦閉じる
    RS_AllClose
        
    strZumenNo = ZumenBango()
    
    strSQL = ""
    strSQL = strSQL & "select * "
    strSQL = strSQL & "FROM WK_木取り寸法表関数 "
    strSQL = strSQL & "where 芯組詳細図番号 = '" & strZumenNo & "' "
    
    Set rsADO = New ADODB.Recordset
    rsADO.Open strSQL, cnADO, adOpenForwardOnly, adLockReadOnly
    
    If Not rsADO.EOF Then
        With KidoriKansu
            .芯組詳細図番号 = strZumenNo
            .芯厚 = Nz(rsADO![芯厚], 0)
            .上下桟_長さ = Nz(rsADO![上下桟_長さ], "0")
            .上下桟_本数 = Nz(rsADO![上下桟_本数], 0)
            .額横1_長さ = Nz(rsADO![額横1_長さ], "0")
            .額横1_本数 = Nz(rsADO![額横1_本数], 0)
            .額横2_長さ = Nz(rsADO![額横2_長さ], "0")
            .額横2_本数 = Nz(rsADO![額横2_本数], 0)
            .額横_LVL30_長さ = Nz(rsADO![額横_LVL30_長さ], "0")
            .額横_LVL30_本数 = Nz(rsADO![額横_LVL30_本数], 0)
            .柱_長さ = Nz(rsADO![柱_長さ], "0")
            .柱_本数 = Nz(rsADO![柱_本数], 0)
            .柱_LVL45_長さ = Nz(rsADO![柱_LVL45_長さ], "0")
            .柱_LVL45_本数 = Nz(rsADO![柱_LVL45_本数], 0)
            .額縦1_長さ = Nz(rsADO![額縦1_長さ], "0")
            .額縦1_本数 = Nz(rsADO![額縦1_本数], 0)
            .額縦_LVL30_長さ = Nz(rsADO![額縦_LVL30_長さ], "0")
            .額縦_LVL30_本数 = Nz(rsADO![額縦_LVL30_本数], 0)
            .額縦2_長さ = Nz(rsADO![額縦2_長さ], "0")
            .額縦2_本数 = Nz(rsADO![額縦2_本数], 0)
            .手掛け_種類 = Nz(rsADO![手掛け_種類], "")
            .手掛け_長さ = Nz(rsADO![手掛け_長さ], "0")
            .手掛け_本数 = Nz(rsADO![手掛け_本数], 0)
            .袖1_長さ = Nz(rsADO![袖1_長さ], "0")
            .袖1_本数 = Nz(rsADO![袖1_本数], 0)
            .袖2_長さ = Nz(rsADO![袖2_長さ], "0")
            .袖2_本数 = Nz(rsADO![袖2_本数], 0)
            .ダボ下地_長さ = Nz(rsADO![ダボ下地_長さ], "0")
            .ダボ下地_本数 = Nz(rsADO![ダボ下地_本数], 0)
            .カップ下地_長さ = Nz(rsADO![カップ下地_長さ], "0")
            .カップ下地_本数 = Nz(rsADO![カップ下地_本数], 0)
            .カップ下地_LVL30_長さ = Nz(rsADO![カップ下地_LVL30_長さ], "0")
            .カップ下地_LVL30_本数 = Nz(rsADO![カップ下地_LVL30_本数], 0)
            .中板芯厚 = Nz(rsADO![中板芯厚], 0)
            .中板上下桟_長さ = Nz(rsADO![中板上下桟_長さ], "0")
            .中板上下桟_本数 = Nz(rsADO![中板上下桟_本数], 0)
            .中板額横_長さ = Nz(rsADO![中板額横_長さ], "0")
            .中板額横_本数 = Nz(rsADO![中板額横_本数], 0)
            .中板柱_長さ = Nz(rsADO![中板柱_長さ], "0")
            .中板柱_本数 = Nz(rsADO![中板柱_本数], 0)
            .中板柱_LVL45_長さ = Nz(rsADO![中板柱_LVL45_長さ], "0")
            .中板柱_LVL45_本数 = Nz(rsADO![中板柱_LVL45_本数], 0)
            
        End With
        
        With AWKansu
            '3型/7型
            If IsAWKansuSearch_Needed(inArgument.in検索品番) Then
                
                '初期化
                .上下桟_長さ = "0"
                .額横1_長さ = "0"
                .額横2_長さ = "0"
                .額横_LVL30_長さ = "0"
                .カップ下地_LVL30_長さ = "0"
                .中板上下桟_LVL30_長さ = "0"
                .中板上下桟_長さ = "0"
                .中板額横_長さ = "0"
                
                Select Case inArgument.inDW
                
                    Case 426 To 570.5
                        .上下桟_長さ = "346"
                        .額横1_長さ = "8"
                        'DNのみ
                        If inArgument.in建具品番_開閉様式 Like "[DV]N" Then
                            .額横2_長さ = "68"
                        End If
                        '2700以上のみ
                        If inArgument.inDH > 2539 Then
                            .額横_LVL30_長さ = "98"
                        End If
                        '7型のみ
                        If inArgument.in検索品番 Like "*-####MF*" Then
                            .中板上下桟_長さ = "DW-281"
                            .中板額横_長さ = "DW-341"
                        End If
                        
                    Case 571 To 618.5
                        .上下桟_長さ = "390"
                        .額横1_長さ = "30"
                        'DNのみ
                        If inArgument.in建具品番_開閉様式 Like "[DV]N" Then
                            .額横2_長さ = "90"
                        End If
                        '2700以上のみ
                        If inArgument.inDH > 2539 Then
                            .額横_LVL30_長さ = "120"
                        End If
                        '7型のみ
                        If inArgument.in検索品番 Like "*-####MF*" Then
                            .中板上下桟_長さ = "DW-325"
                            .中板額横_長さ = "DW-385"
                        End If
                        
                    Case 619 To 669.5
                        .上下桟_長さ = "422"
                        .額横1_長さ = "46"
                        'DNのみ
                        If inArgument.in建具品番_開閉様式 Like "[DV]N" Then
                            .額横2_長さ = "106"
                        End If
                        '2700以上のみ
                        If inArgument.inDH > 2539 Then
                            .額横_LVL30_長さ = "136"
                            .カップ下地_LVL30_長さ = "60"
                        End If
                        '7型のみ
                        If inArgument.in検索品番 Like "*-####MF*" Then
                            .中板上下桟_長さ = "DW-357"
                            .中板額横_長さ = "DW-417"
                        End If
                        
                    Case 670 To 717.5
                        .上下桟_長さ = "454"
                        .額横1_長さ = "62"
                        'DNのみ
                        If inArgument.in建具品番_開閉様式 Like "[DV]N" Then
                            .額横2_長さ = "122"
                        End If
                        '2700以上のみ
                        If inArgument.inDH > 2539 Then
                            .額横_LVL30_長さ = "152"
                            .カップ下地_LVL30_長さ = "60"
                        End If
                        '7型のみ
                        If inArgument.in検索品番 Like "*-####MF*" Then
                            .中板上下桟_長さ = "DW-389"
                            .中板額横_長さ = "DW-449"
                        End If
                        
                    Case 718 To 750.5
                        .上下桟_長さ = "488"
                        .額横1_長さ = "79"
                        'DNのみ
                        If inArgument.in建具品番_開閉様式 Like "[DV]N" Then
                            .額横2_長さ = "139"
                        End If
                        '2700以上のみ
                        If inArgument.inDH > 2539 Then
                            .額横_LVL30_長さ = "169"
                            .カップ下地_LVL30_長さ = "60"
                        End If
                        '7型のみ
                        If inArgument.in検索品番 Like "*-####MF*" Then
                            .中板上下桟_長さ = "DW-423"
                            .中板額横_長さ = "DW-483"
                        End If
                        
                    Case 751 To 780.5
                        .上下桟_長さ = "502"
                        .額横1_長さ = "86"
                        'DNのみ
                        If inArgument.in建具品番_開閉様式 Like "[DV]N" Then
                            .額横2_長さ = "146"
                        End If
                        '2700以上のみ
                        If inArgument.inDH > 2539 Then
                            .額横_LVL30_長さ = "176"
                            .カップ下地_LVL30_長さ = "60"
                        End If
                        '7型のみ
                        If inArgument.in検索品番 Like "*-####MF*" Then
                            .中板上下桟_長さ = "DW-437"
                            .中板額横_長さ = "DW-497"
                        End If
                        
                    Case 781 To 819.5
                        .上下桟_長さ = "526"
                        .額横1_長さ = "98"
                        'DNのみ
                        If inArgument.in建具品番_開閉様式 Like "[DV]N" Then
                            .額横2_長さ = "158"
                        End If
                        '2700以上のみ
                        If inArgument.inDH > 2539 Then
                            .額横_LVL30_長さ = "188"
                            .カップ下地_LVL30_長さ = "60"
                        End If
                        '7型のみ
                        If inArgument.in検索品番 Like "*-####MF*" Then
                            .中板上下桟_長さ = "DW-461"
                            .中板額横_長さ = "DW-521"
                        End If
                        
                    Case 820 To 862.5
                        .上下桟_長さ = "552"
                        .額横1_長さ = "111"
                        'DNのみ
                        If inArgument.in建具品番_開閉様式 Like "[DV]N" Then
                            .額横2_長さ = "171"
                        End If
                        '2700以上のみ
                        If inArgument.inDH > 2539 Then
                            .額横_LVL30_長さ = "201"
                            .カップ下地_LVL30_長さ = "60"
                        End If
                        '7型のみ
                        If inArgument.in検索品番 Like "*-####MF*" Then
                            .中板上下桟_長さ = "DW-487"
                            .中板額横_長さ = "DW-547"
                        End If
                        
                    Case 863 To 900.5
                        .上下桟_長さ = "576"
                        .額横1_長さ = "123"
                        'DNのみ
                        If inArgument.in建具品番_開閉様式 Like "[DV]N" Then
                            .額横2_長さ = "183"
                        End If
                        '2700以上のみ
                        If inArgument.inDH > 2539 Then
                            .額横_LVL30_長さ = "213"
                            .カップ下地_LVL30_長さ = "60"
                        End If
                        '7型のみ
                        If inArgument.in検索品番 Like "*-####MF*" Then
                            .中板上下桟_長さ = "DW-511"
                            .中板額横_長さ = "DW-571"
                        End If
                        
                    Case 901 To 944.5
                        .上下桟_長さ = "610"
                        .額横1_長さ = "140"
                        'DNのみ
                        If inArgument.in建具品番_開閉様式 Like "[DV]N" Then
                            .額横2_長さ = "200"
                        End If
                        '2700以上のみ
                        If inArgument.inDH > 2539 Then
                            .額横_LVL30_長さ = "230"
                            .カップ下地_LVL30_長さ = "60"
                        End If
                        '7型のみ
                        If inArgument.in検索品番 Like "*-####MF*" Then
                            .中板上下桟_長さ = "DW-545"
                            .中板額横_長さ = "DW-605"
                        End If
                        
                    Case 945 To 985.5
                        .上下桟_長さ = "638"
                        .額横1_長さ = "154"
                        'DNのみ
                        If inArgument.in建具品番_開閉様式 Like "[DV]N" Then
                            .額横2_長さ = "214"
                        End If
                        '2700以上のみ
                        If inArgument.inDH > 2539 Then
                            .額横_LVL30_長さ = "244"
                            .カップ下地_LVL30_長さ = "60"
                        End If
                        '7型のみ
                        If inArgument.in検索品番 Like "*-####MF*" Then
                            .中板上下桟_長さ = "DW-573"
                            .中板額横_長さ = "DW-633"
                        End If
                        
                    Case 986 To 1022.5
                        .上下桟_長さ = "662"
                        .額横1_長さ = "166"
                        'DNのみ
                        If inArgument.in建具品番_開閉様式 Like "[DV]N" Then
                            .額横2_長さ = "226"
                        End If
                        '2700以上のみ
                        If inArgument.inDH > 2539 Then
                            .額横_LVL30_長さ = "256"
                            .カップ下地_LVL30_長さ = "60"
                        End If
                        '7型のみ
                        If inArgument.in検索品番 Like "*-####MF*" Then
                            .中板上下桟_長さ = "DW-597"
                            .中板額横_長さ = "DW-657"
                        End If
                        
                    Case 1023 To 1061.5
                        .上下桟_長さ = "688"
                        .額横1_長さ = "179"
                        'DNのみ
                        If inArgument.in建具品番_開閉様式 Like "[DV]N" Then
                            .額横2_長さ = "239"
                        End If
                        '2700以上のみ
                        If inArgument.inDH > 2539 Then
                            .額横_LVL30_長さ = "269"
                            .カップ下地_LVL30_長さ = "60"
                        End If
                        '7型のみ
                        If inArgument.in検索品番 Like "*-####MF*" Then
                            .中板上下桟_長さ = "DW-623"
                            .中板額横_長さ = "DW-683"
                        End If
                        
                    Case 1062 To 1100
                        .上下桟_長さ = "710"
                        .額横1_長さ = "190"
                        'DNのみ
                        If inArgument.in建具品番_開閉様式 Like "[DV]N" Then
                            .額横2_長さ = "250"
                        End If
                        '2700以上のみ
                        If inArgument.inDH > 2539 Then
                            .額横_LVL30_長さ = "280"
                            .カップ下地_LVL30_長さ = "60"
                        End If
                        '7型のみ
                        If inArgument.in検索品番 Like "*-####MF*" Then
                            .中板上下桟_長さ = "DW-645"
                            .中板額横_長さ = "DW-705"
                        End If

                End Select
                
                blnAWKansuMatch = True
            End If
            
        End With
        
        blnKansuMatch = True
        
    End If
        
    GetTateguKansu = True
    
    Exit Function
    
Err_GetTateguKansu:
    blnKansuMatch = False
    blnAWKansuMatch = False
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

Private Function ZumenBango() As String
'***********************************************************
'図面番号取得
' 品番とDHと錠から図面番号を取得
'
'
'***********************************************************

    With inArgument
        ZumenBango = fncstrShingumiShousai_40mm(.in検索品番, .inDH, .in施錠)
    End With
    
End Function

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
    
    cnADO.Close
    Set cnADO = Nothing

End Sub

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
    
End Sub

Private Sub FLG_Init()
'***********************************************************
'検索FLG初期化
'
'***********************************************************
    
    blnKansuMatch = False
    blnAWKansuMatch = False
    
End Sub

Public Property Get 芯組詳細図番号(Optional ByVal in_strTateguHinban As String = "", Optional ByVal in_strJou As String = "", Optional ByVal in_dblDH As Double = 0) As String
'***********************************************************
'芯組詳細図番号取得
'
'***********************************************************
    
    With KidoriKansu
        If blnKansuMatch Then
            芯組詳細図番号 = .芯組詳細図番号
        End If
    End With
    
End Property

Public Property Get 芯厚() As Double
'***********************************************************
'芯厚取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            芯厚 = .芯厚
        End If
    End With
    
End Property

Public Property Get 上下桟_長さ() As Double
'***********************************************************
'上下桟_長さ取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            If blnAWKansuMatch Then
                上下桟_長さ = dblFIVEorZERO(Eval(Replace(AWKansu.上下桟_長さ, "DW", inArgument.inDW)))
            Else
                上下桟_長さ = dblFIVEorZERO(Eval(Replace(.上下桟_長さ, "DW", inArgument.inDW)))
                '縁貼り誤差修正
                If 上下桟_長さ > 0 Then 上下桟_長さ = 上下桟_長さ + TapeGosa.上下桟_長さ差異
            End If
        End If
    End With
    
End Property

Public Property Get 上下桟_本数() As Double
'***********************************************************
'上下桟_本数取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            上下桟_本数 = .上下桟_本数 * inArgument.in枚数
        End If
    End With
    
End Property

Public Property Get 額横1_長さ() As Double
'***********************************************************
'額横1_長さ取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            '本数が0の場合は表示しない
            If 額横1_本数 = 0 Then Exit Property
            
            If blnAWKansuMatch Then
                額横1_長さ = dblFIVEorZERO(Eval(Replace(AWKansu.額横1_長さ, "DW", inArgument.inDW)))
            Else
                額横1_長さ = dblFIVEorZERO(Eval(Replace(.額横1_長さ, "DW", inArgument.inDW)))
                '縁貼り誤差修正
                If 額横1_長さ > 0 Then 額横1_長さ = 額横1_長さ + TapeGosa.額横_長さ差異
            End If
        End If
    End With
    
End Property

Public Property Get 額横1_本数() As Double
'***********************************************************
'額横1_本数取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            額横1_本数 = .額横1_本数 * inArgument.in枚数
            'DNでDH1800以下は1本抜く
            If 額横1_本数 > 0 And inArgument.in建具品番_開閉様式 Like "[DV]N" And inArgument.inDH <= 1800 Then
                額横1_本数 = (額横1_本数 - 1) * inArgument.in枚数
            End If
        End If
    End With
    
End Property

Public Property Get 額横2_長さ() As Double
'***********************************************************
'額横2_長さ取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            If blnAWKansuMatch Then
                額横2_長さ = dblFIVEorZERO(Eval(Replace(AWKansu.額横2_長さ, "DW", inArgument.inDW)))
            Else
                額横2_長さ = dblFIVEorZERO(Eval(Replace(.額横2_長さ, "DW", inArgument.inDW)))
                If 額横2_長さ > 0 Then 額横2_長さ = 額横2_長さ + TapeGosa.額横_長さ差異
            End If
        End If
    End With
    
End Property

Public Property Get 額横2_本数() As Double
'***********************************************************
'額横2_本数取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            額横2_本数 = .額横2_本数 * inArgument.in枚数
        End If
    End With
    
End Property

Public Property Get 額横_LVL30_長さ() As Double
'***********************************************************
'額横_LVL30_長さ取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            If blnAWKansuMatch Then
                額横_LVL30_長さ = dblFIVEorZERO(Eval(Replace(AWKansu.額横_LVL30_長さ, "DW", inArgument.inDW)))
            Else
                額横_LVL30_長さ = dblFIVEorZERO(Eval(Replace(.額横_LVL30_長さ, "DW", inArgument.inDW)))
                '縁貼り誤差修正
                If 額横_LVL30_長さ > 0 Then 額横_LVL30_長さ = 額横_LVL30_長さ + TapeGosa.額横_LVL30_長さ差異
            End If
        End If
    End With
    
End Property

Public Property Get 額横_LVL30_本数() As Double
'***********************************************************
'額横_LVL30_本数取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            If blnAWKansuMatch Then
                '2700以上のみ
                If inArgument.inDH > 2539 Then
                    If inArgument.inDW < 619 Then
                        額横_LVL30_本数 = 8 * inArgument.in枚数
                    Else
                        額横_LVL30_本数 = 4 * inArgument.in枚数
                    End If
                End If
            Else
                額横_LVL30_本数 = .額横_LVL30_本数 * inArgument.in枚数
            End If
        End If
    End With
    
End Property

Public Property Get 柱_長さ() As Double
'***********************************************************
'柱_長さ取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            柱_長さ = dblFIVEorZERO(Eval(Replace(.柱_長さ, "DH", inArgument.inDH)))
        End If
    End With
    
End Property

Public Property Get 柱_本数() As Double
'***********************************************************
'柱_本数取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            柱_本数 = .柱_本数 * inArgument.in枚数
        End If
    End With
    
End Property

Public Property Get 柱_LVL45_長さ() As Double
'***********************************************************
'柱_LVL45_長さ取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            柱_LVL45_長さ = dblFIVEorZERO(Eval(Replace(.柱_LVL45_長さ, "DH", inArgument.inDH)))
        End If
    End With
    
End Property

Public Property Get 柱_LVL45_本数() As Double
'***********************************************************
'柱_LVL45_本数取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            柱_LVL45_本数 = .柱_LVL45_本数 * inArgument.in枚数
        End If
    End With
    
End Property

Public Property Get 額縦1_長さ() As Double
'***********************************************************
'額縦1_長さ取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            額縦1_長さ = dblFIVEorZERO(Eval(Replace(.額縦1_長さ, "DH", inArgument.inDH)))
        End If
    End With
    
End Property

Public Property Get 額縦1_本数() As Double
'***********************************************************
'額縦1_本数取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            額縦1_本数 = .額縦1_本数 * inArgument.in枚数
        End If
    End With
    
End Property

Public Property Get 額縦_LVL30_長さ() As Double
'***********************************************************
'額縦_LVL30_長さ取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            額縦_LVL30_長さ = dblFIVEorZERO(Eval(Replace(.額縦_LVL30_長さ, "DH", inArgument.inDH)))
        End If
    End With
    
End Property

Public Property Get 額縦_LVL30_本数() As Double
'***********************************************************
'額縦_LVL30_本数取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            額縦_LVL30_本数 = .額縦_LVL30_本数 * inArgument.in枚数
        End If
    End With
    
End Property

Public Property Get 額縦2_長さ() As Double
'***********************************************************
'額縦2_長さ取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            'YF1/YF6型
            If inArgument.in建具品番_シリーズ = "Y" And inArgument.in建具品番_建具スタイル = "C" And 額縦2_本数 > 0 Then
                額縦2_長さ = dblFncGakuTate1_YG6(inArgument.in検索品番, inArgument.inDW)
            Else
                額縦2_長さ = dblFIVEorZERO(Eval(Replace(.額縦2_長さ, "DH", inArgument.inDH)))
            End If
        End If
    End With
    
End Property

Public Property Get 額縦2_本数() As Double
'***********************************************************
'額縦2_本数取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            額縦2_本数 = .額縦2_本数 * inArgument.in枚数
        End If
    End With
    
End Property

Public Property Get 手掛け_種類() As String
'***********************************************************
'手掛け_種類取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            If IsNumeric(.手掛け_種類) Then
                手掛け_種類 = .手掛け_種類
            Else
                If inArgument.inハンドル引手種類 Like "KAWAJ_LJ*" Then
                    'カワジュンLJ
                    手掛け_種類 = "100"
                    
                ElseIf inArgument.inハンドル引手種類 Like "KURAMAE*" Then
                    '蔵前
                    手掛け_種類 = "140"
                    
                ElseIf inArgument.inハンドル引手種類 Like "OKUDAIRATK*" _
                        Or inArgument.inハンドル引手種類 Like "PC422*" _
                        Or inArgument.inハンドル引手種類 Like "PC210*" _
                        Or inArgument.inハンドル引手種類 Like "BEST*" Then
                    '引戸
                    手掛け_種類 = "500"
                End If
            End If
        End If
    End With
    
End Property

Public Property Get 手掛け_長さ() As Double
'***********************************************************
'手掛け_長さ取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            手掛け_長さ = dblFIVEorZERO(Eval(.手掛け_長さ))
        End If
    End With
    
End Property

Public Property Get 手掛け_本数() As Double
'***********************************************************
'手掛け_本数取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            手掛け_本数 = .手掛け_本数 * inArgument.in枚数
        End If
    End With
    
End Property

Public Property Get 袖1_長さ() As Double
'***********************************************************
'袖1_長さ取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            袖1_長さ = dblFIVEorZERO(Eval(.袖1_長さ))
        End If
    End With
    
End Property

Public Property Get 袖1_本数() As Double
'***********************************************************
'袖1_本数取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            袖1_本数 = .袖1_本数 * inArgument.in枚数
        End If
    End With
    
End Property

Public Property Get 袖2_長さ() As Double
'***********************************************************
'袖2_長さ取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            袖2_長さ = dblFIVEorZERO(Eval(.袖2_長さ))
        End If
    End With
    
End Property

Public Property Get 袖2_本数() As Double
'***********************************************************
'袖2_本数取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            袖2_本数 = .袖2_本数 * inArgument.in枚数
        End If
    End With
    
End Property

Public Property Get ダボ下地_長さ() As Double
'***********************************************************
'ダボ下地_長さ取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            ダボ下地_長さ = dblFIVEorZERO(Eval(.ダボ下地_長さ))
        End If
    End With
    
End Property

Public Property Get ダボ下地_本数() As Double
'***********************************************************
'ダボ下地_本数取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            ダボ下地_本数 = .ダボ下地_本数 * inArgument.in枚数
        End If
    End With
    
End Property

Public Property Get カップ下地_長さ() As Double
'***********************************************************
'カップ下地_長さ取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            カップ下地_長さ = dblFIVEorZERO(Eval(.カップ下地_長さ))
        End If
    End With
        
End Property

Public Property Get カップ下地_本数() As Double
'***********************************************************
'カップ下地_本数取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            カップ下地_本数 = .カップ下地_本数 * inArgument.in枚数
        End If
    End With
    
End Property

Public Property Get カップ下地_LVL30_長さ() As Double
'***********************************************************
'カップ下地_LVL30_長さ取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            If blnAWKansuMatch Then
                カップ下地_LVL30_長さ = dblFIVEorZERO(Eval(AWKansu.カップ下地_LVL30_長さ))
            Else
                カップ下地_LVL30_長さ = dblFIVEorZERO(Eval(.カップ下地_LVL30_長さ))
            End If
        End If
    End With
    
End Property

Public Property Get カップ下地_LVL30_本数() As Double
'***********************************************************
'カップ下地_LVL30_本数取得
'
'***********************************************************

    With KidoriKansu
        If blnKansuMatch Then
            If blnAWKansuMatch Then
                '2700以上のみ
                If inArgument.inDH > 2539 Then
                    If inArgument.inDW < 619 Then
                        カップ下地_LVL30_本数 = 0
                    Else
                        カップ下地_LVL30_本数 = 8 * inArgument.in枚数
                    End If
                End If
            Else
                カップ下地_LVL30_本数 = .カップ下地_LVL30_本数 * inArgument.in枚数
            End If
        End If
    End With
    
End Property

Public Property Get 中板芯厚() As Double
'***********************************************************
'中板芯厚取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            中板芯厚 = .中板芯厚
        End If
    End With
    
End Property

Public Property Get 中板上下桟_長さ() As Double
'***********************************************************
'中板上下桟_長さ取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            If blnAWKansuMatch Then
                中板上下桟_長さ = dblFIVEorZERO(Eval(Replace(AWKansu.中板上下桟_長さ, "DW", inArgument.inDW)))
            Else
                中板上下桟_長さ = dblFIVEorZERO(Eval(Replace(.中板上下桟_長さ, "DW", inArgument.inDW)))
                '縁貼り誤差修正
                If 中板上下桟_長さ > 0 Then 中板上下桟_長さ = 中板上下桟_長さ + TapeGosa.中板上下桟_長さ差異
            End If
        End If
    End With
    
End Property

Public Property Get 中板上下桟_本数() As Double
'***********************************************************
'中板上下桟_本数取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            中板上下桟_本数 = .中板上下桟_本数 * inArgument.in枚数
        End If
    End With
    
End Property

Public Property Get 中板額横_長さ() As Double
'***********************************************************
'中板額横_長さ取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            If blnAWKansuMatch Then
                中板額横_長さ = dblFIVEorZERO(Eval(Replace(AWKansu.中板額横_長さ, "DW", inArgument.inDW)))
            Else
                中板額横_長さ = dblFIVEorZERO(Eval(Replace(.中板額横_長さ, "DW", inArgument.inDW)))
                '縁貼り誤差修正
                If 中板額横_長さ > 0 Then 中板額横_長さ = 中板額横_長さ + TapeGosa.中板額横_長さ差異
            End If
        End If
    End With
    
End Property

Public Property Get 中板額横_本数() As Double
'***********************************************************
'中板額横_本数取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            中板額横_本数 = .中板額横_本数 * inArgument.in枚数
        End If
    End With
    
End Property

Public Property Get 中板柱_長さ() As Double
'***********************************************************
'中板柱_長さ取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            中板柱_長さ = dblFIVEorZERO(Eval(Replace(.中板柱_長さ, "DH", inArgument.inDH)))
        End If
    End With
    
End Property

Public Property Get 中板柱_本数() As Double
'***********************************************************
'中板柱_本数取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            中板柱_本数 = .中板柱_本数 * inArgument.in枚数
        End If
    End With
    
End Property

Public Property Get 中板柱_LVL45_長さ() As Double
'***********************************************************
'中板柱_LVL45_長さ取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            中板柱_LVL45_長さ = dblFIVEorZERO(Eval(Replace(.中板柱_LVL45_長さ, "DH", inArgument.inDH)))
        End If
    End With
    
End Property

Public Property Get 中板柱_LVL45_本数() As Double
'***********************************************************
'中板柱_LVL45_本数取得
'
'***********************************************************
    With KidoriKansu
        If blnKansuMatch Then
            中板柱_LVL45_本数 = .中板柱_LVL45_本数 * inArgument.in枚数
        End If
    End With
    
End Property

Public Function bolFncKidoriData(ByVal varSpec As Variant, ByVal in_strHinban As String, intMaisu As Integer, ByVal dblDW As Double, ByVal dblDH As Double, ByVal strAkariMado As Variant, ByVal varHandle As Variant _
                            , ByRef KidoriSunpo As KidoriData) As Boolean
'***********************************************************
'木取り寸法表取得関数

'過去帳票システムとの互換のためのもの

'***********************************************************
    '以下定数は現状使用していないのでゼロを送る
    Const dblCH As Double = 0
    Const intSuryo As Integer = 0
    
    Call subKidoriData_Init(KidoriSunpo)
    
    If Me.Bind(in_strHinban, intSuryo, intMaisu, dblDW, dblDH, dblCH, varHandle, strAkariMado, varSpec) Then
    
        With KidoriSunpo
            .out_dblShinAtsu = NullIFNothing(Me.芯厚)
            .out_dblsan = NullIFNothing(Me.上下桟_長さ)
            .out_dblgakuyoko1 = NullIFNothing(Me.額横1_長さ)
            .out_dblgakuYoko2 = NullIFNothing(Me.額横2_長さ)
            .out_dblGakuyokoLVL30 = NullIFNothing(Me.額横_LVL30_長さ)
            .out_dblhashira = NullIFNothing(Me.柱_長さ)
            .out_dblhashira2 = NullIFNothing(Me.柱_LVL45_長さ)
            .out_dblgakutate1 = NullIFNothing(Me.額縦1_長さ)
            .out_dblgakutate2 = NullIFNothing(Me.額縦2_長さ)
            .out_dblgakutate3 = NullIFNothing(Me.額縦_LVL30_長さ)
            .out_dbltegakeShurui = NullIFNothing(Me.手掛け_種類)
            .out_dbltegake = NullIFNothing(Me.手掛け_長さ)
            .out_dblsode1 = NullIFNothing(Me.袖1_長さ)
            .out_dblsode2 = NullIFNothing(Me.袖2_長さ)
            .out_dbldaboshitaji = NullIFNothing(Me.ダボ下地_長さ)
            .out_dblCupShitaji = NullIFNothing(Me.カップ下地_長さ)
            .out_intsan = NullIFNothing(Me.上下桟_本数)
            .out_intgakuyoko1 = NullIFNothing(Me.額横1_本数)
            .out_intgakuyoko2 = NullIFNothing(Me.額横2_本数)
            .out_intgakuyokoLVL30 = NullIFNothing(Me.額横_LVL30_本数)
            .out_inthashira = NullIFNothing(Me.柱_本数)
            .out_inthashira2 = NullIFNothing(Me.柱_LVL45_本数)
            .out_intgakutate1 = NullIFNothing(Me.額縦1_本数)
            .out_intgakutate2 = NullIFNothing(Me.額縦2_本数)
            .out_intgakutate3 = NullIFNothing(Me.額縦_LVL30_本数)
            .out_inttegake = NullIFNothing(Me.手掛け_本数)
            .out_intsode1 = NullIFNothing(Me.袖1_本数)
            .out_intsode2 = NullIFNothing(Me.袖2_本数)
            .out_intdaboshitaji = NullIFNothing(Me.ダボ下地_本数)
            .out_intcupshitaji = NullIFNothing(Me.カップ下地_本数)
            .out_dblShinAtsu_N = NullIFNothing(Me.中板芯厚)
            If IsMonster(inArgument.in検索品番) Then
                .out_dblsan_N = NullIFNothing(Me.中板上下桟_長さ)
                .out_intsanh_N = NullIFNothing(Me.中板上下桟_本数)
                .out_dblsanH2_N = NullIFNothing(0)
                .out_intsanh2_N = NullIFNothing(0)
            Else
                .out_dblsan_N = NullIFNothing(0)
                .out_intsanh_N = NullIFNothing(0)
                .out_dblsanH2_N = NullIFNothing(Me.中板上下桟_長さ)
                .out_intsanh2_N = NullIFNothing(Me.中板上下桟_本数)
            End If
            .out_dblgakuyoko1_N = NullIFNothing(Me.中板額横_長さ)
            .out_dblhashira_N = NullIFNothing(Me.中板柱_長さ)
            .out_dblhashira2_N = NullIFNothing(Me.中板柱_LVL45_長さ)
            .out_dblhashiraSt_N = NullIFNothing(0)
            .out_dblYokosan_N = NullIFNothing(0)
            
            .out_inthashiraH2_N = NullIFNothing(Me.中板柱_LVL45_本数)
            .out_intgakuyokoH1_N = NullIFNothing(Me.中板額横_本数)
            .out_inthashiraH_N = NullIFNothing(Me.中板柱_本数)
            .out_inthashiraStH_N = NullIFNothing(0)
            .out_intYokosanh_N = NullIFNothing(0)
            .out_strShingumizu = NullIFNothing(Me.芯組詳細図番号)
        
        End With
    Else

'        With KidoriSunpo
'            .out_dblShinAtsu = Null
'            .out_dblsan = Null
'            .out_dblgakuyoko1 = Null
'            .out_dblgakuYoko2 = Null
'            .out_dblGakuyokoLVL30 = Null
'            .out_dblhashira = Null
'            .out_dblhashira2 = Null
'            .out_dblgakutate1 = Null
'            .out_dblgakutate2 = Null
'            .out_dblgakutate3 = Null
'            .out_dbltegakeShurui = Null
'            .out_dbltegake = Null
'            .out_dblsode1 = Null
'            .out_dblsode2 = Null
'            .out_dbldaboshitaji = Null
'            .out_dblCupShitaji = Null
'            .out_intsan = Null
'            .out_intgakuyoko1 = Null
'            .out_intgakuyoko2 = Null
'            .out_intgakuyokoLVL30 = Null
'            .out_inthashira = Null
'            .out_inthashira2 = Null
'            .out_intgakutate1 = Null
'            .out_intgakutate2 = Null
'            .out_intgakutate3 = Null
'            .out_inttegake = Null
'            .out_intsode1 = Null
'            .out_intsode2 = Null
'            .out_intdaboshitaji = Null
'            .out_intcupshitaji = Null
'            .out_dblShinAtsu_N = Null
'            .out_dblsan_N = Null
'            .out_dblgakuyoko1_N = Null
'            .out_dblhashira_N = Null
'            .out_dblhashira2_N = Null
'            .out_dblhashiraSt_N = Null
'            .out_dblYokosan_N = Null
'            .out_dblsanH2_N = Null
'            .out_intsanh_N = Null
'            .out_intsanh2_N = Null
'            .out_inthashiraH2_N = Null
'            .out_intgakuyokoH1_N = Null
'            .out_inthashiraH_N = Null
'            .out_inthashiraStH_N = Null
'            .out_intYokosanh_N = Null
'            .out_strShingumizu = Null
'        End With
    End If
    
    bolFncKidoriData = True
    
End Function

Private Function bolfncGetTapeGosa() As Boolean
'***********************************************************
'縁貼りテープによる誤差データ登録

'***********************************************************
    
    With TapeGosa
        
        .上下桟_長さ差異 = 0
        .額横_長さ差異 = 0
        .額横_LVL30_長さ差異 = 0
        .中板上下桟_長さ差異 = 0
        .中板額横_長さ差異 = 0
        
        Select Case inArgument.in建具品番_シリーズ
        
            Case "R" 'リアラート/TAMO
                Select Case inArgument.in建具品番_パネルガラス
                
                    Case "F" 'パネル戸
                        .上下桟_長さ差異 = 2
                        .額横_長さ差異 = 2
                        .額横_LVL30_長さ差異 = 2
                        
                    Case "S", "C" 'サイド/センタースルー
                        .上下桟_長さ差異 = 1
                        .額横_長さ差異 = 0.5
                        .額横_LVL30_長さ差異 = 0.5
                        
                End Select
                
            Case "T", "B" 'SINA/PALIO
                Select Case inArgument.in建具品番_パネルガラス
                
                    Case "F" 'パネル戸
                        .上下桟_長さ差異 = 3
                        .額横_長さ差異 = 3
                        .額横_LVL30_長さ差異 = 3
                        
                    Case "S", "C" 'サイド/センタースルー
                        .上下桟_長さ差異 = 3
                        .額横_長さ差異 = 1.5
                        .額横_LVL30_長さ差異 = 1.5
                        
                End Select
        
            Case "S" 'F/S
                
                Select Case inArgument.in建具品番_パネルガラス
                
                    Case "F" 'KF7
                        .上下桟_長さ差異 = -4
                        .額横_長さ差異 = -4
                        .額横_LVL30_長さ差異 = -4

                        
                End Select
            
            Case "O" 'Monster
                'SH以外
                If inArgument.in建具品番_色 <> "SH" Then
                    .上下桟_長さ差異 = 1.5
                    .額横_長さ差異 = 0.5
                    .額横_LVL30_長さ差異 = 1.5
                    .中板上下桟_長さ差異 = 1.5
                    .中板額横_長さ差異 = 1.5
                End If
        End Select
        
        bolfncGetTapeGosa = True
        
    End With
        
End Function