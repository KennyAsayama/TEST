Version =20
VersionRequired =20
Checksum =1849243254
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =128
    TabularFamily =119
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6679
    DatasheetFontHeight =9
    ItemSuffix =10
    Left =9450
    Top =120
    Right =16080
    Bottom =4305
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x7c13a689729ae440
    End
    GUID = Begin
        0x7f2781420a8e4845b2e9dbd3d485d02f
    End
    NameMap = Begin
        0x0acc0e5500000000000000000000000000000000000000000c00000005000000 ,
        0x0000000000000000000000000000
    End
    Caption ="F_製造番号変更"
    DatasheetFontName ="ＭＳ Ｐゴシック"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    NoSaveCTIWhenDisabled =1
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =128
            FontSize =9
            FontName ="ＭＳ Ｐゴシック"
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            TextFontCharSet =128
            Width =1701
            Height =390
            FontSize =9
            FontWeight =400
            ForeColor =-2147483630
            FontName ="ＭＳ Ｐゴシック"
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =128
            BorderLineStyle =0
            Width =1701
            Height =225
            LabelX =-1701
            FontSize =9
            FontName ="ＭＳ Ｐゴシック"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =128
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =9
            FontName ="ＭＳ Ｐゴシック"
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =128
            BorderLineStyle =0
            Width =1701
            Height =225
            LabelX =-1701
            FontSize =9
            FontName ="ＭＳ Ｐゴシック"
        End
        Begin Section
            Height =4195
            BackColor =-2147483633
            Name ="詳細"
            GUID = Begin
                0x123bb1e5e1d9de4fb5f5b7d5141a0ab3
            End
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =50
                    Left =2154
                    Top =2494
                    Width =2325
                    Height =964
                    FontWeight =700
                    ForeColor =8388608
                    Name ="cmd_Change"
                    Caption ="変更"
                    OnClick ="[Event Procedure]"
                    GUID = Begin
                        0x3793e93e6873c841b7af0ebf97783e8a
                    End

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =50
                    Left =270
                    Top =963
                    Width =2578
                    Height =465
                    FontSize =14
                    TabIndex =1
                    Name ="txt_SeizoNum_Orijinal"
                    FontName ="メイリオ"
                    GUID = Begin
                        0x6e50dfe9a5b89f41aa1a6ae1a11b78eb
                    End

                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =50
                    Left =282
                    Top =630
                    Width =1605
                    Height =300
                    FontWeight =700
                    ForeColor =32768
                    Name ="ラベル461"
                    Caption ="変更前製造番号"
                    FontName ="メイリオ"
                    GUID = Begin
                        0x92f430a04e5257449d2a8fe87861ec9f
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    TextFontFamily =50
                    Left =3825
                    Top =975
                    Width =2793
                    Height =459
                    FontSize =14
                    TabIndex =2
                    GUID = Begin
                        0x1bf937082f96a94d8f7190919ba17405
                    End
                    Name ="list_SeizoNum_After"
                    RowSourceType ="Value List"
                    FontName ="メイリオ"

                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    TextFontFamily =50
                    Left =2664
                    Top =737
                    Width =1304
                    Height =953
                    FontSize =36
                    ForeColor =8388608
                    Name ="ラベル8"
                    Caption ="⇒"
                    FontName ="メイリオ"
                    GUID = Begin
                        0xd7ff1cb2e8181348b1f377355f8513b1
                    End
                End
                Begin Label
                    OverlapFlags =247
                    TextFontFamily =50
                    Left =3870
                    Top =690
                    Width =1530
                    Height =300
                    FontWeight =700
                    ForeColor =255
                    Name ="ラベル9"
                    Caption ="変更後製造番号"
                    FontName ="メイリオ"
                    GUID = Begin
                        0xaa49fc12ed1f0541941a73aa4f603de4
                    End
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Change_Click()
'--------------------------------------------------------------------------------------------------------------------
'登録済のデータを同じく登録済の他の製造番号に変更する
'1.変更対象製造番号
'   →F_製造指示の製造番号リストボックス(list_SeizoNum)で選択されている番号（Load時にtxt_SeizoNum_Orijinalに移送済み）
'
'2.変更後製造番号
'   →リストボックス（list_SeizoNum_After）で選択されている製造番号
'
'※変更後は親画面を再ロード(F_製造指示の[Master_Kensaku_FromOtherForm])してクローズ
'--------------------------------------------------------------------------------------------------------------------
    Dim objREMOTEdb As New cls_BRAND_MASTER
    Dim bolUpdate As Boolean
    
    Dim strSQL As String
    
    bolUpdate = False
    
    strSQL = ""
    strSQL = strSQL & "update T_製造指示 "
    strSQL = strSQL & "set 製造番号 = '" & list_SeizoNum_After.Column(0, list_SeizoNum_After.ListIndex) & "' "
    strSQL = strSQL & "where 製造番号 = '" & txt_SeizoNum_Orijinal & "'"
    If MsgBox("製造番号を変更します。よろしいですか?" & vbCrLf & "（製造日等データは変更されません。ご注意ください）", vbInformation + vbOKCancel) = vbOK Then

        With objREMOTEdb
            .BeginTrans

            If Not .ExecSQL(strSQL) Then
                MsgBox "T_製造実績 更新エラー"
                .Rollback
            Else
                .Commit
                bolUpdate = True
                Call Form_F_製造指示.Master_Kensaku_FromOtherForm
            End If
        End With
        
    End If
    
    Set objREMOTEdb = Nothing
    If bolUpdate Then DoCmd.Close acForm, Me.Form.Name
    'MsgBox list_SeizoNum_After.Column(0, list_SeizoNum_After.ListIndex)
    
End Sub

Private Sub Form_Load()
'--------------------------------------------------------------------------------------------------------------------
'親画面で選択されている製造番号を取り込む
'   →親画面が開いていない場合はクローズ
'   →移行先の製造番号が無い場合はクローズ
'--------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    
    '20150914 K.Asayama Change
    'Application.Echo False
    Me.Painting = False
    '20150914 K.Asayama Change End
    
    On Error GoTo Err_Form_Load
    
    If Not CurrentProject.AllForms("F_製造指示").IsLoaded Then
        Err.Raise 9999, , "このフォームモジュールは[F_製造指示が開いている場合のみ使用できます"
        
    End If
    
    
    txt_SeizoNum_Orijinal = Forms![F_製造指示].list_SeizoNum.Column(0, Forms![F_製造指示].list_SeizoNum.ListIndex)
    
    '配列初期化
    For i = list_SeizoNum_After.ListCount - 1 To 0 Step -1
            list_SeizoNum_After.RemoveItem (i)
    Next i
    
    For i = 1 To Forms![F_製造指示].list_SeizoNum.ListCount
        '移送元と同じ製造番号は送らない
        If (Forms![F_製造指示].list_SeizoNum.Column(0, i - 1) <> txt_SeizoNum_Orijinal) Then
            'NEW(未登録）データは対象外
            If Forms![F_製造指示].list_SeizoNum.Column(0, i - 1) <> "NEW" Then
                'すべて(すべての製造番号を選択）は対象外
                If Forms![F_製造指示].list_SeizoNum.Column(0, i - 1) <> "すべて" Then
                    'Debug.Print Forms![F_製造指示].list_SeizoNum.Column(0, i - 1)
                    list_SeizoNum_After.AddItem Forms![F_製造指示].list_SeizoNum.Column(0, i - 1)
                End If
            End If
        End If
    Next
    
    If list_SeizoNum_After.ListCount < 1 Then
        Err.Raise 9999, , "移行先できる製造番号がありません"
    Else
        list_SeizoNum_After = list_SeizoNum_After.ItemData(0)

    End If
    
    '20150914 K.Asayama Change
    'Application.Echo True
    Me.Painting = True
    '20150914 K.Asayama Change End
    
    Exit Sub
    
Err_Form_Load:
    MsgBox Err.Description
    '20150914 K.Asayama Change
    'Application.Echo True
    Me.Painting = True
    '20150914 K.Asayama Change End
    DoCmd.Close acForm, Me.Form.Name
    
End Sub