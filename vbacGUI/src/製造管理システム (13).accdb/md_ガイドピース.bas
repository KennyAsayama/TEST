Option Compare Database
Option Explicit
'2.1.0 ADD

Public Function IsGuidePieceJ(in_Hinban As String, in_個別Spec As Variant) As Boolean
'   *************************************************************
'   ガイドピース取付定規出力対象チェック
'
'   戻り値:Boolean
'       →True              定規あり
'       →False             定規無し
'
'    Input項目
'       in_Hinban           建具品番
'       in_個別Spec         個別Spec

'   *************************************************************

    Dim strSpec_BRAND As String
    Dim strSPEC As String
        
    IsGuidePieceJ = False
    
    If IsYukazukeRail(in_Hinban) Then
    
        If Not IsNull(in_個別Spec) Then
            strSpec_BRAND = left(in_個別Spec, 3)
            strSPEC = right(in_個別Spec, 4)
            
            If strSpec_BRAND = "BRD" And strSPEC >= "1507" Then
                IsGuidePieceJ = True
            End If
            
        End If
    End If

End Function

Public Function fncdblGuidePiece_SH(in_Hinban As String, dblDH As Double) As Double
'   *************************************************************
'   ガイドピースSH値
'       ガイドピース品番かどうかは本関数ではチェックしないので
'       事前にIsGuidePieceJ関数でチェックしておくこと

'   戻り値:Double
'       →SH値
'
'    Input項目
'       in_Hinban           建具品番
'       dblDH               DH

'   *************************************************************

    If in_Hinban Like "*VL-####*" Or in_Hinban Like "*VM-####*" Or in_Hinban Like "*VN-####*" Then
        fncdblGuidePiece_SH = dblDH - 47
    Else
        fncdblGuidePiece_SH = dblDH - 41
    End If
    
End Function

Public Function fncstrGuidePiece_MM(dblDH As Double) As String
'   *************************************************************
'   ガイドピースmm値
'       ガイドピース品番かどうかは本関数ではチェックしないので
'       事前にIsGuidePieceJ関数でチェックしておくこと

'   戻り値:String
'       →mm値
'
'    Input項目
'       dblDH               DH

'   *************************************************************
    If dblDH < 2411 Then
        fncstrGuidePiece_MM = "4mm"
    Else
        fncstrGuidePiece_MM = "7mm"
    End If
    
End Function

Public Function fncstrGuidePiece_Name(in_Hinban As String, in_個別Spec As Variant) As String
'   *************************************************************
'   ガイドピース名称（コード）取得
'       ガイドピース品番かどうかは本関数ではチェックしないので
'       事前にIsGuidePieceJ関数でチェックしておくこと

'       事前に個別SpecがNullでないことを確認すること（万が一Nullの場合は空欄を返す）

'   戻り値:String
'       →ガイドピース名称

'           1608より前は A

'           1608以降
'               →DMは A
'               →それ以外は B

'           1701以降のVレール
'               →VMはC

'           →それ以外はD

'    Input項目
'       in_Hinban           建具品番
'       in_個別Spec         個別Spec

'   *************************************************************

    If IsNull(in_個別Spec) Then
        fncstrGuidePiece_Name = ""
        
    ElseIf right(in_個別Spec, 4) < "1608" Then
        fncstrGuidePiece_Name = "A"
    Else
        
        If in_Hinban Like "*DM-####*" Then
            fncstrGuidePiece_Name = "A"
        ElseIf in_Hinban Like "*DL-####*" Or in_Hinban Like "*DN-####*" Then
            fncstrGuidePiece_Name = "B"
        ElseIf in_Hinban Like "*VM-####*" Then
            fncstrGuidePiece_Name = "C"
        ElseIf in_Hinban Like "*VL-####*" Or in_Hinban Like "*VN-####*" Then
            fncstrGuidePiece_Name = "D"
        End If
        
    End If

End Function