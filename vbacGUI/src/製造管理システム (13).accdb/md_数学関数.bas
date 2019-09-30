Option Compare Database
Option Explicit

Public Function dblFIVEorZERO(dblValue As Double) As Double
'   *************************************************************
'   少数点以下を0又は0.5に丸める
'
'   戻り値:Double
'       →丸め後の数字
'
'    Input項目
'       dblValue        入力値
'
'   例      10.4 → 10
'           10.5 → 10.5
'           10.6 → 10.5
'
'   1.10.18 ADD
'   *************************************************************

    If dblValue * 10 Mod 10 < 5 Then
        dblFIVEorZERO = RoundDown(CCur(dblValue), 0)
    Else
        dblFIVEorZERO = RoundDown(CCur(dblValue), 0) + 0.5
    End If
    
End Function

Public Function RoundDown(CurValue As Currency, Optional Intdp As Integer) As Double
'   *************************************************************
'   切捨て関数
'
'   戻り値:Double
'       →丸め後の数字
'
'    Input項目
'       CurValue        入力値
'       Intdp           丸める桁数
'
'   IntDPの例       -2  →  10の位を丸める
'                   -1  →  1の位を丸める
'                   0   →  少数点以下第1位を丸める
'                   1   →  少数点以下第2位を丸める
'
'   1.10.18 ADD
'   *************************************************************

    RoundDown = (Int(Abs(CurValue) * 10 ^ Intdp) / 10 ^ Intdp) * Sgn(CurValue)

End Function

Public Function Roundx(CurValue As Currency, Optional Intdp As Integer) As Double
'   *************************************************************
'   四捨五入関数
'
'   戻り値:Double
'       →丸め後の数字
'
'    Input項目
'       CurValue        入力値
'       Intdp           丸める桁数
'
'   IntDPの例       -2  →  10の位を丸める
'                   -1  →  1の位を丸める
'                   0   →  少数点以下第1位を丸める
'                   1   →  少数点以下第2位を丸める
'
'   1.12.0 ADD
'   *************************************************************

    Roundx = (Int((Abs(CurValue) * 10 ^ Intdp) + 0.5) / 10 ^ Intdp) * Sgn(CurValue)

End Function

Public Function RoundUp(CurValue As Currency, Optional Intdp As Integer) As Currency
'   *************************************************************
'   切上げ関数
'
'   戻り値:Currency
'       →丸め後の数字
'
'    Input項目
'       CurValue        入力値
'       Intdp           丸める桁数
'
'   IntDPの例       -2  →  10の位を丸める
'                   -1  →  1の位を丸める
'                   0   →  少数点以下第1位を丸める
'                   1   →  少数点以下第2位を丸める
'
'   3.0.0 ADD
'   *************************************************************

    Dim W As Currency
  
    W = 10 ^ Abs(Intdp)
    
    If Intdp > 0 Then
        RoundUp = Int(CurValue * W + 0.999) / W
    Else
        RoundUp = Int(CurValue / W + 0.999) * W
    End If
    
End Function