Option Compare Database
Option Explicit

Public Function dblFIVEorZERO(dblValue As Double) As Double
'   *************************************************************
'   �����_�ȉ���0����0.5�Ɋۂ߂�
'
'   �߂�l:Double
'       ���ۂߌ�̐���
'
'    Input����
'       dblValue        ���͒l
'
'   ��      10.4 �� 10
'           10.5 �� 10.5
'           10.6 �� 10.5
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
'   �؎̂Ċ֐�
'
'   �߂�l:Double
'       ���ۂߌ�̐���
'
'    Input����
'       CurValue        ���͒l
'       Intdp           �ۂ߂錅��
'
'   IntDP�̗�       -2  ��  10�̈ʂ��ۂ߂�
'                   -1  ��  1�̈ʂ��ۂ߂�
'                   0   ��  �����_�ȉ���1�ʂ��ۂ߂�
'                   1   ��  �����_�ȉ���2�ʂ��ۂ߂�
'
'   1.10.18 ADD
'   *************************************************************

    RoundDown = (Int(Abs(CurValue) * 10 ^ Intdp) / 10 ^ Intdp) * Sgn(CurValue)

End Function

Public Function Roundx(CurValue As Currency, Optional Intdp As Integer) As Double
'   *************************************************************
'   �l�̌ܓ��֐�
'
'   �߂�l:Double
'       ���ۂߌ�̐���
'
'    Input����
'       CurValue        ���͒l
'       Intdp           �ۂ߂錅��
'
'   IntDP�̗�       -2  ��  10�̈ʂ��ۂ߂�
'                   -1  ��  1�̈ʂ��ۂ߂�
'                   0   ��  �����_�ȉ���1�ʂ��ۂ߂�
'                   1   ��  �����_�ȉ���2�ʂ��ۂ߂�
'
'   1.12.0 ADD
'   *************************************************************

    Roundx = (Int((Abs(CurValue) * 10 ^ Intdp) + 0.5) / 10 ^ Intdp) * Sgn(CurValue)

End Function

Public Function RoundUp(CurValue As Currency, Optional Intdp As Integer) As Currency
'   *************************************************************
'   �؏グ�֐�
'
'   �߂�l:Currency
'       ���ۂߌ�̐���
'
'    Input����
'       CurValue        ���͒l
'       Intdp           �ۂ߂錅��
'
'   IntDP�̗�       -2  ��  10�̈ʂ��ۂ߂�
'                   -1  ��  1�̈ʂ��ۂ߂�
'                   0   ��  �����_�ȉ���1�ʂ��ۂ߂�
'                   1   ��  �����_�ȉ���2�ʂ��ۂ߂�
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