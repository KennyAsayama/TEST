Version =20
VersionRequired =20
Checksum =-805947968
Begin Form
    PopUp = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowUpdating =2
    ScrollBars =2
    ViewsAllowed =1
    TabularCharSet =128
    TabularFamily =49
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12271
    DatasheetFontHeight =9
    ItemSuffix =587
    Left =-14633
    Top =-18998
    Right =-1058
    Bottom =-11183
    DatasheetGridlinesColor =12632256
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0x644ad60416f0e440
    End
    GUID = Begin
        0x7f3a920e64961c4d9b603e5fbedd7c97
    End
    NameMap = Begin
        0x0acc0e55000000009ad480493adfb8418b1aca7110135754000000008bd50c46 ,
        0x34dce440000000000000000057004b005f005300740061007400750073000000 ,
        0x0000000097e7df5072e22c45aba1f2c914468d28070000009ad480493adfb841 ,
        0x8b1aca71101357544a5264960000000000009f300146ba199c40b365fff9e6a0 ,
        0x0069070000009ad480493adfb8418b1aca71101357545159047d6a75f7530000 ,
        0x00000000118fdc316cfcdf479af951259810a35d070000009ad480493adfb841 ,
        0x8b1aca7110135754df686a75f7530000000000005753a7caad1fbd40ab3b4754 ,
        0x432cd6ce070000009ad480493adfb8418b1aca7110135754e8904b5c6a75f753 ,
        0x00000000000022c6cade7c720e4fb6c579e0397bc2f0070000009ad480493adf ,
        0xb8418b1aca711013575405980000000000001dddd618e227ed46bc8c0d8c7e6d ,
        0x216a070000009ad480493adfb8418b1aca7110135754c1546a753a5306520000 ,
        0x00000000767f12a49bb049458a02772b167f2e68070000009ad480493adfb841 ,
        0x8b1aca71101357543a5306520d54f07900000000000090cb30bf675e5d4eb844 ,
        0xddf6ece73543070000009ad480493adfb8418b1aca7110135754c1546a750000 ,
        0x00000000ec631f11510d5043bf248eb76ed75851070000009ad480493adfb841 ,
        0x8b1aca711013575470650000000000004b4874c63bc78e4698888680d32797c5 ,
        0x070000009ad480493adfb8418b1aca71101357544d7a977bd753d84ee5650000 ,
        0x00000000700a5883d48ad143acc44416b4802b3b070000009ad480493adfb841 ,
        0x8b1aca71101357544d7a977b1a904e90e5650000000000003bf6e5b5e8b46848 ,
        0xa0379f57d43105f1070000009ad480493adfb8418b1aca7110135754937ae353 ,
        0xd753d84ee5650000000000000ce927a9e35460468d1f840087dc084307000000 ,
        0x9ad480493adfb8418b1aca7110135754937ae3531a904e90e565000000000000 ,
        0x9abf1db78c8d5241ade9ae0f411ead90070000009ad480493adfb8418b1aca71 ,
        0x101357542d8a088ad753d84ee565000000000000e5b80d04e473774cadd02edf ,
        0x811f9eec070000009ad480493adfb8418b1aca71101357542d8a088a1a904e90 ,
        0xe565000000000000873d268a292cb648808d55c154e69296070000009ad48049 ,
        0x3adfb8418b1aca7110135754c78c5067d753d84ee565000000000000df57eaf8 ,
        0x5ae094479bea24aa5406d39c070000009ad480493adfb8418b1aca7110135754 ,
        0xc78c50671a904e90e565000000000000757304ddd4f83d42adf70fcc86ff5582 ,
        0x070000009ad480493adfb8418b1aca7110135754fd882090d753d84ee5650000 ,
        0x00000000960166900483fb41bf858e73dd5a4df7070000009ad480493adfb841 ,
        0x8b1aca7110135754155f216e1f675096e565000000000000cb5f52c307d82c4d ,
        0x9d47249871585137070000009ad480493adfb8418b1aca7110135754c1546a75 ,
        0xd753e86ce565000000000000f89a0fb47736d244a4999e4fac4fce6907000000 ,
        0x9ad480493adfb8418b1aca7110135754bd30fc30c8300698000000000000096e ,
        0xead23e8a4d49b982f9ac7cee6376070000009ad480493adfb8418b1aca711013 ,
        0x57545159047d4e006f0000000000000000000000000000000000000000000000 ,
        0x0c000000050000000000000000000000000000000000
    End
    RecordSource ="SELECT WK_Status.削除, [契約番号] & \"-\" & [棟番号] & \"-\" & [部屋番号] AS 契約No, WK_Status."
        "契約番号, WK_Status.棟番号, WK_Status.部屋番号, WK_Status.項, WK_Status.品番区分, WK_Status.区分名称"
        ", WK_Status.品番, WK_Status.数, WK_Status.積算受付日 AS 製造日, WK_Status.積算通過日, WK_Status."
        "窓口受付日, WK_Status.窓口通過日, WK_Status.設計受付日, WK_Status.設計通過日, WK_Status.資材受付日, WK_St"
        "atus.資材通過日, WK_Status.製造受付日, WK_Status.引渡期限日, WK_Status.品番受注日, WK_Status.[ソート順] "
        "FROM WK_Status; "
    Caption ="F_製造先行受付確認"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="ＭＳ ゴシック"
    PrtMip = Begin
        0x8805000088050000880500008805000000000000b85500000d20000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    PrtDevMode = Begin
        0x0097365f2613971ff8c5bf0000000000904d6e0170ccdd55400000007ddc3e22 ,
        0x010403069c00e40553ef8001020008009a0b3408640001000f00b00402000100 ,
        0xb00403000000413400486e0134cddd5540cddd55904d6e0198c5bf1920cedd55 ,
        0xf8c5bf1902000000000000000000000000000000010000000000000001000000 ,
        0x0100000001000000ffffffff0000000000000000000000000000000044494e55 ,
        0x2200380194035002b9da64890000000000000000000000000000000000000000 ,
        0x00000000000000000b0000000000000004000400000001000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000100000000000000000000000000000000000000 ,
        0x00000000000000000000000000000000000000000000000038010000534d544a ,
        0x0000000010002801000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000050020000465058460000006a50020000 ,
        0x465058460000006ad114000000000000ff7f0000000000000000000001000000 ,
        0x0000000101000001030100000000000000000000000000000200000000000100 ,
        0x0034089a0b000000000000000200000000020000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x00000153616d654173506170657253697a650000030005000000020001000000 ,
        0x00000101020064000200392e3500000000000001030001000000010000000001 ,
        0x0000000000000000000064000000000000000004640000000000000000000000 ,
        0x0100042e002e002e002e002e002e002e002e002e002e002e002e002e002e002e ,
        0x002e002e002e002e002e002e002e002e002e002e002e002e002e000000000101 ,
        0x0000000100010900000000000000000000000100c80001010096000b00000000 ,
        0x0000000000000000000100000000102700001027000001000000000000000000
    End
    PrtDevNames = Begin
        0x08001f0022000100000000000000000000000000000000000000000000000000 ,
        0x00003139322e3136382e312e313400
    End
    OnKeyDown ="[Event Procedure]"
    OnTimer ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    PrtDevModeW = Begin
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x01040306dc00e40553ef8001020008009a0b3408640001000f00b00402000100 ,
        0xb004030000004100340000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000010000000000000001000000 ,
        0x0100000001000000ffffffff0000000000000000000000000000000044494e55 ,
        0x2200380194035002b9da64890000000000000000000000000000000000000000 ,
        0x00000000000000000b0000000000000004000400000001000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000100000000000000000000000000000000000000 ,
        0x00000000000000000000000000000000000000000000000038010000534d544a ,
        0x0000000010002801000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000050020000465058460000006a50020000 ,
        0x465058460000006ad114000000000000ff7f0000000000000000000001000000 ,
        0x0000000101000001030100000000000000000000000000000200000000000100 ,
        0x0034089a0b000000000000000200000000020000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x00000153616d654173506170657253697a650000030005000000020001000000 ,
        0x00000101020064000200392e3500000000000001030001000000010000000001 ,
        0x0000000000000000000064000000000000000004640000000000000000000000 ,
        0x0100042e002e002e002e002e002e002e002e002e002e002e002e002e002e002e ,
        0x002e002e002e002e002e002e002e002e002e002e002e002e002e000000000101 ,
        0x0000000100010900000000000000000000000100c80001010096000b00000000 ,
        0x0000000000000000000100000000102700001027000001000000000000000000
    End
    PrtDevNamesW = Begin
        0x04001b001e000100000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000031003900 ,
        0x32002e003100360038002e0031002e00310034000000
    End
    NoSaveCTIWhenDisabled =1
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =128
            FontSize =9
            FontName ="ＭＳ Ｐゴシック"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin CommandButton
            TextFontCharSet =128
            Height =390
            FontSize =9
            FontWeight =400
            ForeColor =-2147483630
            FontName ="ＭＳ Ｐゴシック"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =128
            BorderLineStyle =0
            FontSize =9
            FontName ="ＭＳ Ｐゴシック"
        End
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =128
            BorderLineStyle =0
            FontSize =9
            FontName ="ＭＳ Ｐゴシック"
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =128
            BorderLineStyle =0
            FontSize =9
            FontName ="ＭＳ Ｐゴシック"
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin ToggleButton
            TextFontCharSet =128
            FontSize =9
            FontWeight =400
            ForeColor =-2147483630
            FontName ="ＭＳ Ｐゴシック"
            BorderLineStyle =0
        End
        Begin Tab
            TextFontCharSet =128
            FontSize =9
            FontName ="ＭＳ Ｐゴシック"
            BorderLineStyle =0
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin FormHeader
            Height =2415
            BackColor =12632256
            Name ="フォームヘッダー"
            GUID = Begin
                0xa5b8cce1de22e14499a4ef82499b82cc
            End
            Begin
                Begin Label
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =50
                    Left =-45
                    Width =4440
                    Height =480
                    FontSize =18
                    TopMargin =11
                    BackColor =8388608
                    BorderColor =16777215
                    ForeColor =12632256
                    Name ="lbl_Midashi"
                    Caption ="製造先行受付確認"
                    FontName ="HGP創英角ｺﾞｼｯｸUB"
                    GUID = Begin
                        0x004f1220dc8eb84ea4ae0d72233d08f5
                    End
                    LayoutCachedLeft =-45
                    LayoutCachedWidth =4395
                    LayoutCachedHeight =480
                End
                Begin CommandButton
                    OverlapFlags =93
                    AccessKey =81
                    TextFontFamily =50
                    Left =9361
                    Width =1545
                    Height =600
                    FontWeight =700
                    ForeColor =8388608
                    Name ="cmd_Close"
                    Caption ="終了(&Q)"
                    OnClick ="[Event Procedure]"
                    FontName ="メイリオ"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="ﾌｫｰﾑを閉じる"
                    GUID = Begin
                        0x994b2d04c8430e44ba1821e1dda65cc3
                    End
                    UnicodeAccessKey =81

                    LayoutCachedLeft =9361
                    LayoutCachedWidth =10906
                    LayoutCachedHeight =600
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =50
                    Left =7320
                    Width =1545
                    Height =600
                    FontWeight =700
                    TabIndex =2
                    ForeColor =8388608
                    Name ="cmd_Enter"
                    Caption ="Excel(F8)"
                    OnClick ="[Event Procedure]"
                    FontName ="メイリオ"
                    GUID = Begin
                        0x69e81a9084944f4a8c2e83d77fd9292e
                    End

                    LayoutCachedLeft =7320
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =600
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =50
                    Left =315
                    Top =30
                    Width =3675
                    Height =375
                    FontSize =18
                    BackColor =8388608
                    BorderColor =16777215
                    ForeColor =16777215
                    Name ="ラベル467"
                    Caption ="製造先行受付確認"
                    FontName ="HGP創英角ｺﾞｼｯｸUB"
                    GUID = Begin
                        0x1ba5765e9310de43ba27394f49d8ff1c
                    End
                    LayoutCachedLeft =315
                    LayoutCachedTop =30
                    LayoutCachedWidth =3990
                    LayoutCachedHeight =405
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =50
                    Left =4575
                    Width =1545
                    Height =600
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="cmd_Reload"
                    Caption ="再読込み(F1)"
                    OnClick ="[Event Procedure]"
                    FontName ="メイリオ"
                    GUID = Begin
                        0xd57d8ff99c9ffc4483a9de0210811ab5
                    End

                    LayoutCachedLeft =4575
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =600
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =2
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =50
                    Left =11090
                    Width =1134
                    Height =259
                    FontWeight =700
                    ForeColor =13056
                    Name ="lbl_UserID"
                    Caption ="00000"
                    FontName ="メイリオ"
                    GUID = Begin
                        0xdb060ca261f9d6489387aa6dceff475b
                    End
                    LayoutCachedLeft =11090
                    LayoutCachedWidth =12224
                    LayoutCachedHeight =259
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =2
                    OverlapFlags =87
                    TextAlign =3
                    TextFontFamily =50
                    Left =11090
                    Top =270
                    Width =1134
                    Height =259
                    FontWeight =700
                    ForeColor =13056
                    Name ="lbl_UserName"
                    Caption ="XXXX"
                    FontName ="メイリオ"
                    GUID = Begin
                        0x446bd6d5b3a2a04794ecb989679e6a00
                    End
                    LayoutCachedLeft =11090
                    LayoutCachedTop =270
                    LayoutCachedWidth =12224
                    LayoutCachedHeight =529
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontFamily =50
                    Left =1215
                    Top =1350
                    Width =4650
                    Height =705
                    ColumnOrder =1
                    FontSize =12
                    TabIndex =3
                    Name ="txt_物件名"
                    FontName ="メイリオ"
                    GUID = Begin
                        0x68a2353f00312a40ad85abda49109b3f
                    End

                    LayoutCachedLeft =1215
                    LayoutCachedTop =1350
                    LayoutCachedWidth =5865
                    LayoutCachedHeight =2055
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =50
                    Left =75
                    Top =1350
                    Width =1110
                    Height =300
                    FontSize =12
                    ForeColor =8388608
                    Name ="ラベル133"
                    Caption ="現 場 名"
                    FontName ="メイリオ"
                    GUID = Begin
                        0x32268acc74130745864024daf06c8bad
                    End
                    LayoutCachedLeft =75
                    LayoutCachedTop =1350
                    LayoutCachedWidth =1185
                    LayoutCachedHeight =1650
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontFamily =50
                    Left =10891
                    Top =555
                    Width =1380
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    ForeColor =8388608
                    Name ="cmd_Screenshot"
                    Caption ="ScreenShot"
                    OnClick ="[Event Procedure]"
                    FontName ="メイリオ"
                    GUID = Begin
                        0x572cee29d409b4409e2bcf6b3a6fdfa2
                    End

                    LayoutCachedLeft =10891
                    LayoutCachedTop =555
                    LayoutCachedWidth =12271
                    LayoutCachedHeight =825
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =50
                    Left =1395
                    Top =2130
                    Width =2310
                    Height =285
                    FontSize =10
                    FontWeight =700
                    BackColor =10319446
                    BorderColor =13020235
                    ForeColor =16777215
                    Name ="lbl契約No_Sort"
                    Caption ="契約番号"
                    FontName ="メイリオ"
                    GUID = Begin
                        0x4a49c69c27d97444864cb5544fc50c45
                    End
                    GridlineStyleBottom =1
                    LayoutCachedLeft =1395
                    LayoutCachedTop =2130
                    LayoutCachedWidth =3705
                    LayoutCachedHeight =2415
                    ForeThemeColorIndex =1
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    TextFontFamily =50
                    Left =3705
                    Top =2130
                    Width =345
                    Height =285
                    FontSize =10
                    FontWeight =700
                    BackColor =10319446
                    BorderColor =13020235
                    ForeColor =16777215
                    Name ="lbl項"
                    Caption ="項"
                    FontName ="メイリオ"
                    GUID = Begin
                        0x27504c813ef145438986b07b0fd75f6a
                    End
                    GridlineStyleBottom =1
                    LayoutCachedLeft =3705
                    LayoutCachedTop =2130
                    LayoutCachedWidth =4050
                    LayoutCachedHeight =2415
                    ForeThemeColorIndex =1
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    TextFontFamily =50
                    Left =4047
                    Top =2130
                    Width =1005
                    Height =285
                    FontSize =10
                    FontWeight =700
                    BackColor =10319446
                    BorderColor =13020235
                    ForeColor =16777215
                    Name ="lbl区分名称_Sort"
                    Caption ="区分名称"
                    FontName ="メイリオ"
                    GUID = Begin
                        0x3f96ed4a970b0e4eb379f4cbf5318030
                    End
                    GridlineStyleBottom =1
                    LayoutCachedLeft =4047
                    LayoutCachedTop =2130
                    LayoutCachedWidth =5052
                    LayoutCachedHeight =2415
                    ForeThemeColorIndex =1
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    TextFontFamily =50
                    Left =5049
                    Top =2130
                    Width =3765
                    Height =285
                    FontSize =10
                    FontWeight =700
                    BackColor =10319446
                    BorderColor =13020235
                    ForeColor =16777215
                    Name ="lbl品番_Sort"
                    Caption ="品番"
                    FontName ="メイリオ"
                    GUID = Begin
                        0x6f71b72372d5af4997fbd5475c3eacd2
                    End
                    GridlineStyleBottom =1
                    LayoutCachedLeft =5049
                    LayoutCachedTop =2130
                    LayoutCachedWidth =8814
                    LayoutCachedHeight =2415
                    ForeThemeColorIndex =1
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    TextFontFamily =50
                    Left =8814
                    Top =2130
                    Width =645
                    Height =285
                    FontSize =10
                    FontWeight =700
                    BackColor =10319446
                    BorderColor =13020235
                    ForeColor =16777215
                    Name ="lbl数"
                    Caption ="数"
                    FontName ="メイリオ"
                    GUID = Begin
                        0x5dbfa484528b2a4b89f53581cff05b7d
                    End
                    GridlineStyleBottom =1
                    LayoutCachedLeft =8814
                    LayoutCachedTop =2130
                    LayoutCachedWidth =9459
                    LayoutCachedHeight =2415
                    ForeThemeColorIndex =1
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    TextFontFamily =50
                    Left =9461
                    Top =2130
                    Width =2805
                    Height =285
                    FontSize =10
                    FontWeight =700
                    BackColor =10319446
                    BorderColor =13020235
                    ForeColor =16777215
                    Name ="lbl製造受付日_Sort"
                    Caption ="製造受付日"
                    FontName ="メイリオ"
                    GUID = Begin
                        0x8791313085fc8a4daf664001e3ac986a
                    End
                    GridlineStyleBottom =1
                    LayoutCachedLeft =9461
                    LayoutCachedTop =2130
                    LayoutCachedWidth =12266
                    LayoutCachedHeight =2415
                    ForeThemeColorIndex =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =50
                    Left =75
                    Top =780
                    Width =1110
                    Height =300
                    FontSize =12
                    ForeColor =8388608
                    Name ="ラベル581"
                    Caption ="契約NO"
                    FontName ="メイリオ"
                    GUID = Begin
                        0x3008d4dc44771f4896bbdcfcda9bca8e
                    End
                    LayoutCachedLeft =75
                    LayoutCachedTop =780
                    LayoutCachedWidth =1185
                    LayoutCachedHeight =1080
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontFamily =50
                    Left =1245
                    Top =780
                    Width =2535
                    Height =330
                    ColumnOrder =0
                    FontSize =12
                    TabIndex =5
                    Name ="txt_契約No"
                    FontName ="メイリオ"
                    GUID = Begin
                        0xb26ef0255870ba439592e30e2c0c3dad
                    End

                    LayoutCachedLeft =1245
                    LayoutCachedTop =780
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =1110
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    TextFontFamily =50
                    Top =2130
                    Width =1395
                    Height =285
                    FontSize =10
                    FontWeight =700
                    BackColor =10319446
                    BorderColor =13020235
                    ForeColor =16777215
                    Name ="lbl製造日_Sort"
                    Caption ="製造日"
                    FontName ="メイリオ"
                    GUID = Begin
                        0x9bc93f7d75ed0040a925402417027610
                    End
                    GridlineStyleBottom =1
                    LayoutCachedTop =2130
                    LayoutCachedWidth =1395
                    LayoutCachedHeight =2415
                    ForeThemeColorIndex =1
                End
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =50
                    Left =11010
                    Top =1800
                    Width =570
                    Height =284
                    FontSize =11
                    BorderColor =16777215
                    ForeColor =5026082
                    Name ="ラベル60"
                    Caption ="昇順"
                    FontName ="メイリオ"
                    GUID = Begin
                        0xd334b1a10d0ee54ab455369e655bd96f
                    End
                    GridlineColor =10921638
                    LayoutCachedLeft =11010
                    LayoutCachedTop =1800
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =2084
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =50
                    Left =11640
                    Top =1800
                    Width =570
                    Height =284
                    FontSize =11
                    BorderColor =16777215
                    ForeColor =5676533
                    Name ="ラベル61"
                    Caption ="降順"
                    FontName ="メイリオ"
                    GUID = Begin
                        0x6a529108eee990439feed6ed73e4208c
                    End
                    GridlineColor =10921638
                    LayoutCachedLeft =11640
                    LayoutCachedTop =1800
                    LayoutCachedWidth =12210
                    LayoutCachedHeight =2084
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =50
                    Left =11490
                    Top =1800
                    Width =225
                    Height =284
                    FontSize =11
                    BorderColor =16777215
                    Name ="ラベル62"
                    Caption ="/"
                    FontName ="メイリオ"
                    GUID = Begin
                        0x3ed46dfecb052742ad1a6664230f8eb6
                    End
                    GridlineColor =10921638
                    LayoutCachedLeft =11490
                    LayoutCachedTop =1800
                    LayoutCachedWidth =11715
                    LayoutCachedHeight =2084
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =285
            Name ="詳細"
            GUID = Begin
                0xcb8bc16c660ada4996a9b0ada2eca08c
            End
            AlternateBackColor =16247774
            AlternateBackThemeColorIndex =4
            AlternateBackTint =20.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =50
                    IMEMode =1
                    Left =5694
                    Width =990
                    Height =285
                    ColumnOrder =14
                    FontSize =10
                    BackColor =62207
                    Name ="txt_契約番号"
                    ControlSource ="契約番号"
                    FontName ="メイリオ"
                    OnClick ="[Event Procedure]"
                    GUID = Begin
                        0x9054f45c0b177f4f8f1243c86952b8ef
                    End

                    LayoutCachedLeft =5694
                    LayoutCachedWidth =6684
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =50
                    IMEMode =1
                    Left =7674
                    Width =990
                    Height =285
                    ColumnOrder =12
                    FontSize =10
                    TabIndex =1
                    BackColor =62207
                    Name ="txt_部屋番号"
                    ControlSource ="部屋番号"
                    FontName ="メイリオ"
                    OnClick ="[Event Procedure]"
                    GUID = Begin
                        0x20ae9896e1dee9419b0483539312a525
                    End

                    LayoutCachedLeft =7674
                    LayoutCachedWidth =8664
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    TextFontFamily =50
                    IMEMode =1
                    Left =6690
                    Width =990
                    Height =285
                    ColumnOrder =13
                    FontSize =10
                    TabIndex =2
                    BackColor =62207
                    Name ="txt_棟番号"
                    ControlSource ="棟番号"
                    FontName ="メイリオ"
                    OnClick ="[Event Procedure]"
                    GUID = Begin
                        0x74bd478362a2714d80362a740acfd82a
                    End

                    LayoutCachedLeft =6690
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =50
                    BackStyle =0
                    Left =3705
                    Width =345
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="txt_項"
                    ControlSource ="項"
                    FontName ="メイリオ"
                    OnClick ="[Event Procedure]"
                    GUID = Begin
                        0x828ca53a37c3e4448a22ddd66e75eec2
                    End

                    LayoutCachedLeft =3705
                    LayoutCachedWidth =4050
                    LayoutCachedHeight =285
                    BackThemeColorIndex =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    TextFontFamily =50
                    IMEMode =1
                    BackStyle =0
                    Left =4047
                    Width =1005
                    Height =285
                    FontSize =10
                    TabIndex =4
                    Name ="txt_区分名称"
                    ControlSource ="区分名称"
                    FontName ="メイリオ"
                    OnClick ="[Event Procedure]"
                    GUID = Begin
                        0x71e2a7f19e851d49afee639402bd587d
                    End

                    LayoutCachedLeft =4047
                    LayoutCachedWidth =5052
                    LayoutCachedHeight =285
                    BackThemeColorIndex =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =255
                    TextAlign =1
                    TextFontFamily =50
                    IMEMode =1
                    BackStyle =0
                    Left =5049
                    Width =3765
                    Height =285
                    FontSize =10
                    TabIndex =5
                    Name ="txt_品番"
                    ControlSource ="品番"
                    FontName ="メイリオ"
                    OnClick ="[Event Procedure]"
                    GUID = Begin
                        0x08e820794217544f812cc55659f21df4
                    End

                    LayoutCachedLeft =5049
                    LayoutCachedWidth =8814
                    LayoutCachedHeight =285
                    BackThemeColorIndex =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =127
                    TextAlign =3
                    TextFontFamily =50
                    BackStyle =0
                    Left =8814
                    Width =645
                    Height =285
                    FontSize =10
                    TabIndex =6
                    Name ="txt_数"
                    ControlSource ="数"
                    FontName ="メイリオ"
                    OnClick ="[Event Procedure]"
                    GUID = Begin
                        0x05552d7c0c2b6648b62d410917ef63b0
                    End

                    LayoutCachedLeft =8814
                    LayoutCachedWidth =9459
                    LayoutCachedHeight =285
                    BackThemeColorIndex =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =50
                    IMEMode =2
                    BackStyle =0
                    Width =1395
                    Height =285
                    FontSize =10
                    TabIndex =7
                    Name ="txt_積算受付日"
                    ControlSource ="製造日"
                    FontName ="メイリオ"
                    GUID = Begin
                        0xaad570b376f8794cbff0259d247f3264
                    End
                    ShowDatePicker =0

                    LayoutCachedWidth =1395
                    LayoutCachedHeight =285
                    BackThemeColorIndex =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =119
                    TextAlign =2
                    TextFontFamily =50
                    IMEMode =2
                    BackStyle =0
                    Left =9461
                    Width =2805
                    Height =285
                    FontSize =10
                    TabIndex =8
                    Name ="txt_製造受付日"
                    ControlSource ="製造受付日"
                    FontName ="メイリオ"
                    OnClick ="[Event Procedure]"
                    GUID = Begin
                        0x9a541917e707c349ae18ecce0b901d0b
                    End

                    LayoutCachedLeft =9461
                    LayoutCachedWidth =12266
                    LayoutCachedHeight =285
                    BackThemeColorIndex =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    TextFontFamily =50
                    IMEMode =1
                    BackStyle =0
                    Left =1395
                    Width =2310
                    Height =285
                    FontSize =10
                    TabIndex =9
                    BackColor =62207
                    Name ="txt_契約"
                    ControlSource ="契約No"
                    FontName ="メイリオ"
                    GUID = Begin
                        0xb07898abcfcc0d488627582b4e49a376
                    End

                    LayoutCachedLeft =1395
                    LayoutCachedWidth =3705
                    LayoutCachedHeight =285
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =12632256
            Name ="フォームフッター"
            GUID = Begin
                0x2e1cc3c6a81bcb418e524111764d1aff
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'1.12.2 ADD
Private ctrlSortLbl() As New cls_SortLabelset

Private Sub cmd_Close_Click()
'--------------------------------------------------------------------------------------------------------------------
'画面を閉じる
'--------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    DoCmd.Close acForm, Me.Form.Name, acSaveNo
End Sub

Private Sub cmd_Kensaku_Click()
'--------------------------------------------------------------------------------------------------------------------
'検索開始
'--------------------------------------------------------------------------------------------------------------------
    If IsNull(txt_契約番号) Or IsNull(txt_棟番号) Or IsNull(txt_部屋番号) Then
        MsgBox "契約番号を正しく入力して下さい"
    Else
        Call Master_Kensaku
    End If
    
End Sub

Private Sub Form_Init()
'--------------------------------------------------------------------------------------------------------------------
'フォーム初期化
'--------------------------------------------------------------------------------------------------------------------
    
    Dim i As Integer
    
    On Error GoTo Err_Form_Init
    
    'ワークテーブル削除
    DoCmd.RunSQL "delete * from WK_Status"
    

    'ユーザー情報表示
    lbl_UserID.Caption = strUserID
    lbl_UserName.Caption = strUserName
    
    Me.Refresh

    Exit Sub
    
Err_Form_Init:
        MsgBox Err.Number & Err.Description
        DoCmd.Close acForm, Me.Form.Name
End Sub

Private Sub cmd_Enter_Click()
    
    Dim objExcel As New cls_Excel
    Dim objLOCALdb As New cls_LOCALDB
    Dim strSQL As String
    
    On Error GoTo Err_cmd_Enter_Click
    
    cmd_Enter.Enabled = False
    Screen.MousePointer = 11
    
    strSQL = ""
    strSQL = strSQL & "select [契約番号] & '-' & [棟番号] & '-' & [部屋番号] AS 契約No "
    strSQL = strSQL & ", WK_Status.項 "
    strSQL = strSQL & ", WK_Status.区分名称 "
    strSQL = strSQL & ", WK_Status.品番 "
    strSQL = strSQL & ", WK_Status.数 "
    strSQL = strSQL & ", WK_Status.積算受付日 AS 製造日 "
    strSQL = strSQL & ", WK_Status.製造受付日 "
    strSQL = strSQL & "FROM WK_Status "
    
    objExcel.getExcel.Workbooks.Add
    
    If objLOCALdb.ExecSelect(strSQL) Then
    
        If Not bolfncexp_EXCELOBJECT(objLOCALdb.GetRS, objExcel.getExcel, True, "製造先行受付商品一覧") Then
            Err.Raise 9999, , "Excel抽出エラー"
        End If
    
    Else
        Err.Raise 9999, , "製造先行テーブルコピーエラー"
    End If
    
    objExcel.ContinueOpen = True
    
    GoTo Exit_cmd_Enter_Click
    
Err_cmd_Enter_Click:
    MsgBox Err.Description
    
Exit_cmd_Enter_Click:
    Screen.MousePointer = 0
    cmd_Enter.Enabled = True
    
    Set objLOCALdb = Nothing
    Set objExcel = Nothing
    
End Sub

Private Sub cmd_Reload_Click()
'--------------------------------------------------------------------------------------------------------------------
'リロードボタンクリック時
'   →再読み込み
'--------------------------------------------------------------------------------------------------------------------
    Application.Echo False
    
    subStatus_Copy
    
    Me.Requery
    Me.Refresh
    
    Application.Echo True
    
End Sub


Private Sub cmd_Screenshot_Click()
'--------------------------------------------------------------------------------------------------------------------
'スクリーンショット取得、Excelに貼り付け
'Excelへの貼り付けは別プロシージャで実行しないとうまく起動しないのでタイマーで2秒後に実行する

'--------------------------------------------------------------------------------------------------------------------
'アクティブ画面のスクリーンショット取得
    
    '2秒後に貼り付けコマンドを実行
    Me.TimerInterval = 2000
    
    subScreenShot_ActiveArea
End Sub

Private Sub Form_Close()
'--------------------------------------------------------------------------------------------------------------------
'フォームクローズ時はアプリケーションウィンドウを元に戻す
'--------------------------------------------------------------------------------------------------------------------
     WindowSize_Restore
End Sub

Private Sub Form_Current()
'--------------------------------------------------------------------------------------------------------------------
'カレントレコード変更時
'   →受注ﾏｽﾀ邸名検索
'--------------------------------------------------------------------------------------------------------------------
    Master_Kensaku
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'--------------------------------------------------------------------------------------------------------------------
'フォーム上のボタン押下イベント受信
'   →ファンクションキーでのショートカット用
'--------------------------------------------------------------------------------------------------------------------
    
    If Fnc_KeyEvent(KeyCode) = 0 Then KeyCode = 0

End Sub

Private Sub Form_Load()
'--------------------------------------------------------------------------------------------------------------------
'フォーム読み込み処理
'   →データベース名はODBC設定(DB02_BRAND_MASTER)の接続先データベース名を取得
'--------------------------------------------------------------------------------------------------------------------

    Dim strDB As String
    Dim ctl As Access.Control
    
    'DB名取得（SQLServerであることが前提、型式は[サーバ名\データベース名])
    strDB = Connection_DB()
    
    '本番DB以外に接続する場合は見出しを赤表示
    If strDBName <> left(strDB, InStr(1, strDB, "\") - 1) Then 'PUBLIC変数のDB名と違う場合
        lbl_Midashi.BackColor = vbRed
    Else
        lbl_Midashi.BackColor = 8388608
    End If
    
    'タイマー初期化
    Me.TimerInterval = 0
    
    '画面初期化
    Form_Init
    
    'データコピー
    subStatus_Copy
    
    Me.OrderBy = "契約No,項"
    Me.OrderByOn = True
    
    'サブフォームの条件付書式設定
'    For Each ctl In frmSub.Form.Controls
'        With ctl
'            If .ControlType = acTextBox Then
'                If .Visible Then
'                '条件付き書式を設定する
'                        ctl.FormatConditions.Delete
''                        With ctl.FormatConditions.Add(acExpression, acEqual, "削除 = 0")
''                            .ForeColor = vbBlack
''                            .BackColor = vbWhite
''                            .FontBold = False
''                        End With
'                        With ctl.FormatConditions.Add(acExpression, acEqual, "削除<>0")
'                            .ForeColor = vbWhite
'                            .BackColor = vbRed
''                            .FontBold = True
'                        End With
'
'                End If
'            End If
'        End With
'    Next ctl
'
'    Set ctl = Nothing
    
    Me.Requery
    Me.Refresh
End Sub

Private Sub Form_Open(Cancel As Integer)
'--------------------------------------------------------------------------------------------------------------------
'フォーム開始時チェック
'
'--------------------------------------------------------------------------------------------------------------------
    Dim ctl As Access.Control
    Dim i As Byte
    
    'F_Status表示とワークファイルを共用しているため開いていたら終了
    If Form_IsLoaded("F_Status表示") Then
        MsgBox "Status画面を開いていると表示できません", vbCritical
        Cancel = True
        Exit Sub
    End If
    
    '日時処理終了チェック（未終了の場合は強制終了）
    If Not bolfncEnableSystem Then
        Cancel = True
        Exit Sub
    '共通変数が空欄の場合はログオン画面表示
    ElseIf strUserID = "" Then
        If Not bolfncOpen_LogOnMenu(Me.Form.Name) Then
            Cancel = True
        End If
    End If
    
    '見出しラベルにソート用コントロールバインド（ダブルクリックで並べ替え）
    For Each ctl In フォームヘッダー.Controls
        With ctl
            If .ControlType = acLabel And .Name Like "*_Sort" Then
                
                ReDim Preserve ctrlSortLbl(i)
                ctrlSortLbl(i).Bind ctl
                
                i = i + 1
            End If
        End With
    Next ctl
    
End Sub

Private Sub Master_Kensaku()
'--------------------------------------------------------------------------------------------------------------------
'データ検索処理
'   →T_受注マスタを検索して対象が存在するか確認
'   →存在していた場合はStatusを取得する
'
'--------------------------------------------------------------------------------------------------------------------
    Dim objREMOTEdb As New cls_BRAND_MASTER
    
    Dim strSQL As String
    
    On Error GoTo Err_Master_Kensaku
    
    DoCmd.Hourglass True

    Me.Painting = False
    
    txt_契約No = Null
    txt_物件名 = Null

    
    If Not IsNull(txt_契約番号) And Not IsNull(txt_棟番号) And Not IsNull(txt_部屋番号) Then
    
        strSQL = ""
        strSQL = strSQL & "select M.契約番号,M.棟番号,M.部屋番号,M.物件名 from T_受注ﾏｽﾀ AS M "
        strSQL = strSQL & "where M.契約番号 = '" & txt_契約番号
        strSQL = strSQL & "' and M.棟番号 = '" & txt_棟番号
        strSQL = strSQL & "' and M.部屋番号 = '" & txt_部屋番号
        strSQL = strSQL & "' "

        If objREMOTEdb.ExecSelect(strSQL) Then
            
            If Not objREMOTEdb.GetRS.EOF Then
                txt_物件名 = objREMOTEdb.GetRS![物件名]
                txt_契約No = txt_契約番号 & "-" & txt_棟番号 & "-" & txt_部屋番号
            Else
                Err.Raise 9999, , "邸が存在しません"
            End If

        Else
            MsgBox "エラーです"
        End If
    End If
    
    
    GoTo Exit_Master_Kensaku
    
Err_Master_Kensaku:
    MsgBox Err.Description
    Err.Clear
    On Error Resume Next
    Form_Init
     
Exit_Master_Kensaku:

    'クラスのインスタンスを破棄
    Set objREMOTEdb = Nothing
    
    Application.Echo True
    Me.Painting = True
    
    DoCmd.Hourglass False
End Sub

Private Sub subStatus_Copy()
'--------------------------------------------------------------------------------------------------------------------
'リモートテーブル（T_Status製造先行）をローカルテーブル（WK_Status)にコピー
'   →WK_Statusテーブルを代用するので積算受付日に製造日を入れる
'
'2.5.3
'   →エラー文言変更
'--------------------------------------------------------------------------------------------------------------------
    Dim objREMOTEdb As New cls_BRAND_MASTER
    
    Dim strSQL As String
    
    On Error GoTo Err_subStatus_Copy
    
    strSQL = ""
    strSQL = strSQL & "select cast(0 as BIT)  削除,a.契約番号 ,a.棟番号 ,a.部屋番号 ,a.項 "
    strSQL = strSQL & ",a.品番区分 "
    strSQL = strSQL & ",dbo.fncStatusKbnName(a.品番区分,種類,dbo.fncGetHinban(品番1,特注建具品番))  区分名称 "
    strSQL = strSQL & ",dbo.fncStatusShohinName(a.品番区分,dbo.fncGetHinban(品番1,特注建具品番),dbo.fncGetHinban(枠品番1,特注枠品番),dbo.fncGetHinban(下地材品番,特注下地材品番)) 品番 "
    strSQL = strSQL & ",case when 種類 = '出入口' and a.品番区分 = 1 then b.枚数 else b.数量 end 数 "
    strSQL = strSQL & ",case when c.確定 > 0 then cast(c.製造日 as datetime) else Null end as 積算受付日"
    strSQL = strSQL & ",Null as 積算通過日,Null as 窓口受付日,Null as 窓口通過日,Null as 設計受付日 "
    strSQL = strSQL & ",Null as 設計通過日,Null as 資材受付日,Null as 資材通過日 "
    strSQL = strSQL & ",製造受付日 "
    strSQL = strSQL & ",Null as 引渡期限日,Null as 品番受注日 "
    strSQL = strSQL & ",case dbo.fncStatusKbnName(a.品番区分,種類,dbo.fncGetHinban(品番1,特注建具品番)) when '建具' then 0 else a.品番区分 end ソート順 "
    strSQL = strSQL & "from T_Status製造先行 a "
    strSQL = strSQL & "inner join T_受注明細 b "
    strSQL = strSQL & "on a.契約番号 = b.契約番号 and a.棟番号 = b.棟番号 and a.部屋番号 = b.部屋番号 and a.項 = b.項 "
    strSQL = strSQL & "left join T_製造指示 c "
    strSQL = strSQL & "on a.契約番号 = c.契約番号 and a.棟番号 = c.棟番号 and a.部屋番号 = c.部屋番号 and a.項 = c.項 and a.製造指示品番区分 = c.品番区分 "
    
    
    With objREMOTEdb
        
        If .ExecSelect(strSQL) Then
            If Not .GetRS.EOF Then
                If Not bolfncTableCopyToLocal(.GetRS, "WK_Status", False) Then
                    Err.Raise 9999, , "T_status製造先行ローカルコピーエラー。管理者に連絡してください"
                End If
            Else
                Err.Raise 9998, , "Status製造先行データはありません"
            End If
        Else
            Err.Raise 9999, , "T_status製造先行検索エラー"
        End If
        
 
    End With
    
    GoTo Exit_subStatus_Copy
    
Err_subStatus_Copy:
    If Err.Number = 9998 Then
        MsgBox Err.Description, vbInformation
    Else
        MsgBox Err.Description
    End If
    
Exit_subStatus_Copy:

    Set objREMOTEdb = Nothing
    
End Sub

Private Sub Form_Timer()
'--------------------------------------------------------------------------------------------------------------------
'タイマー起動時
'   →クリップボードをExcelに貼り付ける。その後タイマーを終了する

'--------------------------------------------------------------------------------------------------------------------
    sub_ClipBord_Paste_to_Excel
    Me.TimerInterval = 0
    
End Sub

Private Sub frmSub_KeyDown(KeyCode As Integer, Shift As Integer)
    If Fnc_KeyEvent(KeyCode) = 0 Then KeyCode = 0
End Sub

Private Function Fnc_KeyEvent(KeyCode As Integer) As Integer

'--------------------------------------------------------------------------------------------------------------------
'フォーム上のボタン押下イベント受信
'   →ファンクションキーでのショートカット用
'--------------------------------------------------------------------------------------------------------------------
    
    Fnc_KeyEvent = KeyCode
    
    Select Case KeyCode
    
        Case vbKeyF1
            '再読込みボタン押下
            If cmd_Reload.Enabled Then
                cmd_Reload.SetFocus
                cmd_Reload_Click
            End If
            Fnc_KeyEvent = 0
        Case vbKeyF2
    
        Case vbKeyF3
    
        Case vbKeyF4

        Case vbKeyF5
          
        Case vbKeyF6
          
        Case vbKeyF7
          'F7キーには「スペルチェック」の機能が
          '割り当てられています
          Fnc_KeyEvent = 0
        Case vbKeyF8
            '更新ボタン押下
    '        If cmd_Enter.Enabled Then
    '            cmd_Enter.SetFocus
    '            cmd_Enter_Click
    '        End If
            
            Fnc_KeyEvent = 0
        Case vbKeyF9
          'F9キーには「画面再表示」の機能が
          '割り当てられています
          
            Fnc_KeyEvent = 0
            
        Case vbKeyF10
            
        Case vbKeyF11
          'F11キーには「データベースウィンドウ表示」の機能が
          '割り当てられています
          Fnc_KeyEvent = 0
        Case vbKeyF12
          'F11キーには「フォーム保存」の機能が
          '割り当てられています
          Fnc_KeyEvent = 0
    End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
'--------------------------------------------------------------------------------------------------------------------
'フォームを閉じる前処理
'   →ラベルコントロールを無効にする
'--------------------------------------------------------------------------------------------------------------------
 
    Dim i As Byte
    
    On Error GoTo Err_Form_Unload
    
    For i = LBound(ctrlSortLbl) To UBound(ctrlSortLbl)
        Set ctrlSortLbl(i) = Nothing
    Next
    
    GoTo Exit_Form_Unload
    
Err_Form_Unload:

Exit_Form_Unload:
End Sub