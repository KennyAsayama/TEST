Option Compare Database
Option Explicit

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/�@CopyRight(C) K.Tsunoda(AddinBox) 2001 All Rights Reserved.
'_/�@( http://www.h3.dion.ne.jp/~sakatsu/index.htm )
'_/
'_/�@�@���̏j���}�N���́wkt�֐��A�h�C���x�Ŏg�p���Ă�����̂ł��B
'_/�@�@���̃��W�b�N�́A���X�|���X����`�Ƃ��āA�\�Ȍ��菭�Ȃ�
'_/�@  �y��������̎��s�z�Ō��ʂ��o����悤�ɐ݌v���Ă���܂��B
'_/�@�@���̊֐��ł́A�Q�O�P�U�N�{�s�̉����j���@(�R�̓�)�܂ł�
'_/�@  �T�|�[�g���Ă��܂��B
'_/
'_/�@(*1)���̃}�N�������p����ɓ������ẮA�K�����̃R�����g��
'_/�@�@�@�ꏏ�Ɉ��p���鎖�Ƃ��܂��B
'_/�@(*2)���T�C�g��Ŗ{�}�N���𒼐ڈ��p���鎖�́A�������肢�܂��B
'_/�@�@�@�y http://www.h3.dion.ne.jp/~sakatsu/holiday_logic.htm �z
'_/�@�@�@�ւ̃����N�ɂ��Љ�őΉ����ĉ������B
'_/�@(*3)[ktHolidayName]�Ƃ����֐������̂��̂́A�e���̊���
'_/�@�@�@�����閽���K���ɉ����ĕύX���Ă��\���܂���B
'_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Public Function ktHolidayName(ByVal ���t As Date) As String
Dim dtm���t As Date
Dim str�j���� As String
Const cst�U�֋x���{�s�� As Date = "1973/4/12"

'����/�����덷�̍폜(Now�֐��Ȃǂւ̑Ή�)
    dtm���t = DateSerial(Year(���t), Month(���t), Day(���t))
    '�V���A���l��[�}0.5�b]�̌덷�͈͂ŔF������܂��B2002/6/21�̓V���A���l��
    '[37428.0]�ł���������[-0.5�b]�̌덷�������[37427.9999942130]�ƂȂ�
    'Int�֐��Ő������������o����[37427]�őO�����t�ɂȂ��Ă��܂��܂��B
    '�� �A��������Ɏw�肷��l���K���y����͂������t�z���Ť�����⎞���덷��
    '�@ �l�����Ȃ��Ă��ǂ��Ȃ�Τ���̃X�e�b�v�͕s�v�ł��B����[���t]�����̂܂�
    '�@ �g�p���Ă���肠��܂���(�قƂ�ǂ̗��p�`�Ԃł͂�����ł��傤���d�d)�B

    str�j���� = prv�j��(dtm���t)
    If (str�j���� = "") Then
        If (Weekday(dtm���t) = vbMonday) Then
            ' ���j�ȊO�͐U�֋x������s�v
            ' 5/6(��,��)�̔����[prv�j��]�ŏ�����
            ' 5/6(��)�͂����Ŕ��肷��
            If (dtm���t >= cst�U�֋x���{�s��) Then
                str�j���� = prv�j��(dtm���t - 1)
                If (str�j���� <> "") Then
                    ktHolidayName = "�U�֋x��"
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
        ktHolidayName = str�j����
    End If
End Function

'========================================================================
Private Function prv�j��(ByVal ���t As Date) As String
Dim int�N As Integer
Dim int�� As Integer
Dim int�� As Integer
Dim int�H���� As Integer
Dim str��N�j�� As String
' �����f�[�^(������)�͎�菜���Ă���̂ŁA���L�̓��t�Ƃ̔�r�͂n�j
Const cst�j���@�{�s As Date = "1948/7/20"
Const cst���a�V�c�̑�r�̗� As Date = "1989/2/24"
Const cst���m�e���̌����̋V As Date = "1959/4/10"
Const cst���m�e���̌����̋V As Date = "1993/6/9"
Const cst���ʗ琳�a�̋V As Date = "1990/11/12"

    int�N = Year(���t)
    int�� = Month(���t)
    int�� = Day(���t)

    prv�j�� = ""
    If (���t < cst�j���@�{�s) Then
        Exit Function    ' �j���@�{�s�ȑO
    End If

    Select Case int��
    Case 1
        If (int�� = 1) Then
            prv�j�� = "����"
        Else
            If (int�N >= 2000) Then
                str��N�j�� = (((int�� - 1) \ 7) + 1) & Weekday(���t)
                If (str��N�j�� = "22") Then  'Monday:2
                    prv�j�� = "���l�̓�"
                End If
            Else
                If (int�� = 15) Then
                    prv�j�� = "���l�̓�"
                End If
            End If
        End If
    Case 2
        If (int�� = 11) Then
            If (int�N >= 1967) Then
                prv�j�� = "�����L�O�̓�"
            End If
        ElseIf (���t = cst���a�V�c�̑�r�̗�) Then
            prv�j�� = "���a�V�c�̑�r�̗�"
        End If
    Case 3
        If (int�� = prv�t����(int�N)) Then  ' 1948�`2150�ȊO��[99]
            prv�j�� = "�t���̓�"            ' ���Ԃ�̂Ť�K�����ɂȂ�
        End If
    Case 4
        If (int�� = 29) Then
            If (int�N >= 2007) Then
                prv�j�� = "���a�̓�"
            ElseIf (int�N >= 1989) Then
                prv�j�� = "�݂ǂ�̓�"
            Else
                prv�j�� = "�V�c�a����"
            End If
        ElseIf (���t = cst���m�e���̌����̋V) Then
            prv�j�� = "�c���q���m�e���̌����̋V"
        End If
    Case 5
        If (int�� = 3) Then
            prv�j�� = "���@�L�O��"
        ElseIf (int�� = 4) Then
            If (int�N >= 2007) Then
                prv�j�� = "�݂ǂ�̓�"
            ElseIf (int�N >= 1986) Then
                ' 5/4�����j���́w���̓��j�x����j���́w���@�L�O���̐U�֋x���x(�`2006�N)
                If (Weekday(���t) > vbMonday) Then
                    prv�j�� = "�����̋x��"
                End If
            End If
        ElseIf (int�� = 5) Then
            prv�j�� = "���ǂ��̓�"
        ElseIf (int�� = 6) Then
            If (int�N >= 2007) Then
                Select Case Weekday(���t)
                    Case vbTuesday, vbWednesday
                        prv�j�� = "�U�֋x��"    ' [5/3,5/4�����j]�P�[�X�̂݁A�����Ŕ���
                End Select
            End If
        End If
    Case 6
        If (���t = cst���m�e���̌����̋V) Then
            prv�j�� = "�c���q���m�e���̌����̋V"
        End If
    Case 7
        If (int�N >= 2003) Then
            str��N�j�� = (((int�� - 1) \ 7) + 1) & Weekday(���t)
            If (str��N�j�� = "32") Then  'Monday:2
                prv�j�� = "�C�̓�"
            End If
        ElseIf (int�N >= 1996) Then
            If (int�� = 20) Then
                prv�j�� = "�C�̓�"
            End If
        End If
    Case 8
        If (int�� = 11) Then
            If (int�N >= 2016) Then
                prv�j�� = "�R�̓�"
            End If
        End If
    Case 9
        '��R���j��(15�`21)�ƏH����(22�`24)���d�Ȃ鎖�͂Ȃ�
        int�H���� = prv�H����(int�N)
        If (int�� = int�H����) Then  ' 1948�`2150�ȊO��[99]
            prv�j�� = "�H���̓�"      ' ���Ԃ�̂Ť�K�����ɂȂ�
        Else
            If (int�N >= 2003) Then
                str��N�j�� = (((int�� - 1) \ 7) + 1) & Weekday(���t)
                If (str��N�j�� = "32") Then  'Monday:2
                    prv�j�� = "�h�V�̓�"
                ElseIf (Weekday(���t) = vbTuesday) Then
                    If (int�� = (int�H���� - 1)) Then
                        prv�j�� = "�����̋x��"
                    End If
                End If
            ElseIf (int�N >= 1966) Then
                If (int�� = 15) Then
                    prv�j�� = "�h�V�̓�"
                End If
            End If
        End If
    Case 10
        If (int�N >= 2000) Then
            str��N�j�� = (((int�� - 1) \ 7) + 1) & Weekday(���t)
            If (str��N�j�� = "22") Then  'Monday:2
                prv�j�� = "�̈�̓�"
            End If
        ElseIf (int�N >= 1966) Then
            If (int�� = 10) Then
                prv�j�� = "�̈�̓�"
            End If
        End If
    Case 11
        If (int�� = 3) Then
            prv�j�� = "�����̓�"
        ElseIf (int�� = 23) Then
            prv�j�� = "�ΘJ���ӂ̓�"
        ElseIf (���t = cst���ʗ琳�a�̋V) Then
            prv�j�� = "���ʗ琳�a�̋V"
        End If
    Case 12
        If (int�� = 23) Then
            If (int�N >= 1989) Then
                prv�j�� = "�V�c�a����"
            End If
        End If
    End Select
End Function

'======================================================================
'�@�t��/�H�����̗��Z����
'�@�@�w�C��ۈ������H�� ��v�Z������� �V����ݕ֗����x
'�@�ŏЉ��Ă��鎮�ł��B
Private Function prv�t����(ByVal �N As Integer) As Integer
    If (�N <= 1947) Then
        prv�t���� = 99        '�j���@�{�s�O
    ElseIf (�N <= 1979) Then
        '(�N - 1983)���}�C�i�X�ɂȂ�̂ŁwFix�֐��x�ɂ���
        prv�t���� = Fix(20.8357 + (0.242194 * (�N - 1980)) - Fix((�N - 1983) / 4))
    ElseIf (�N <= 2099) Then
        prv�t���� = Fix(20.8431 + (0.242194 * (�N - 1980)) - Fix((�N - 1980) / 4))
    ElseIf (�N <= 2150) Then
        prv�t���� = Fix(21.851 + (0.242194 * (�N - 1980)) - Fix((�N - 1980) / 4))
    Else
        prv�t���� = 99        '2151�N�ȍ~�͗��Z���������̂ŕs��
    End If
End Function

'========================================================================
Private Function prv�H����(ByVal �N As Integer) As Integer
    If (�N <= 1947) Then
        prv�H���� = 99        '�j���@�{�s�O
    ElseIf (�N <= 1979) Then
        '(�N - 1983)���}�C�i�X�ɂȂ�̂ŁwFix�֐��x�ɂ���
        prv�H���� = Fix(23.2588 + (0.242194 * (�N - 1980)) - Fix((�N - 1983) / 4))
    ElseIf (�N <= 2099) Then
        prv�H���� = Fix(23.2488 + (0.242194 * (�N - 1980)) - Fix((�N - 1980) / 4))
    ElseIf (�N <= 2150) Then
        prv�H���� = Fix(24.2488 + (0.242194 * (�N - 1980)) - Fix((�N - 1980) / 4))
    Else
        prv�H���� = 99        '2151�N�ȍ~�͗��Z���������̂ŕs��
    End If
End Function

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/�@CopyRight(C) K.Tsunoda(AddinBox) 2001 All Rights Reserved.
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/