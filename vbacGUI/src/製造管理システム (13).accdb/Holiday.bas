Option Compare Database
Option Explicit

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/  --- VB / VBA �� ( Update: 2018/12/8 ) ---
'_/
'_/  CopyRight(C) K.Tsunoda(AddinBox) 2001 All Rights Reserved.
'_/  ( AddinBox  http://addinbox.sakura.ne.jp/index.htm )
'_/  (  ���T�C�g  http://www.h3.dion.ne.jp/~sakatsu/index.htm )
'_/
'_/    ���̏j���}�N���́wkt�֐��A�h�C���x�Ŏg�p���Ă�����̂ł��B
'_/    ���̃��W�b�N�́A���X�|���X����`�Ƃ��āA�\�Ȍ��菭�Ȃ�
'_/    �y��������̎��s�z�Ō��ʂ��o����悤�ɐ݌v���Ă���܂��B
'_/
'_/    ���̊֐��ł͈ȉ��̏j���ύX�܂ŃT�|�[�g���Ă��܂��B
'_/    (a) 2019�N�{�s�́u�V�c�a�����̕ύX�v 12/23��2/23 (��F2019�N�ɂ�[�V�c�a����]�͂���܂���)
'_/    (b) 2019�N�̓��m�e���̑��ʓ�(5/1) �����
'_/       �j���ɋ��܂�āu�����̋x���v�ƂȂ� 4/30(�����V�c�̑ވʓ�) �� 5/2 �̂Q�x��
'_/    (c) 2019�N�́u���ʂ̗� ���a�̋V (10/22) �v
'_/    (d) 2020�N�{�s�́u�̈�̓��̉����v�˃X�|�[�c�̓�
'_/    (e) �ܗ֓��[�@�ɂ��2020�N�́u�j���ړ��v
'_/       �C�̓��F7/20(3rd Mon)��7/23, �X�|�[�c�̓�:10/12(2nd Mon)��7/24, �R�̓��F8/11��8/10
'_/
'_/  (*1)���̃}�N�������p����ɓ������ẮA�K�����̃R�����g��
'_/      �ꏏ�Ɉ��p���鎖�Ƃ��܂��B
'_/  (*2)���T�C�g��Ŗ{�}�N���𒼐ڈ��p���鎖�́A�������肢�܂��B
'_/      �y http://addinbox.sakura.ne.jp/holiday_logic.htm �z
'_/      �ւ̃����N�ɂ��Љ�őΉ����ĉ������B
'_/  (*3)[ktHolidayName]�Ƃ����֐������̂��̂́A�e���̊���
'_/      �����閽���K���ɉ����ĕύX���Ă��\���܂���B
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
    '  �l�����Ȃ��Ă��ǂ��Ȃ�Τ���̃X�e�b�v�͕s�v�ł��B����[���t]�����̂܂�
    '  �g�p���Ă���肠��܂���(�قƂ�ǂ̗��p�`�Ԃł͂�����ł��傤���d�d)�B


    str�j���� = prv�j��(dtm���t)

    ' ----- �U�֋x���̔��� (�U�֋x���{�s��:1973/4/12) -----
    ' [ �Ώۓ����j��/�x�� �� �Ώۓ������j�� ]�̂݁A�O��(�����j��)���j�����肷��B
    ' �O��(�����j��)���j���̏ꍇ�́h�U�֋x���h�ƂȂ�B
    ' ���A�T���U���̈�����
    '     �u�Ηj or ���j(�݂ǂ�̓�(5/4) or ���@�L�O��(5/3)�̐U�֋x��)�v�˂T���u���b�N���Ŕ����
    '     �u���j(���ǂ��̓�(5/5)�̐U�֋x���v�˂����̔��菈���Ŕ���
    ' �Ƃ��鎖�ɂ��A�����ł̔���Ώۂ́w�Ώۓ������j���x�̂� �ƂȂ�܂��B
    If (str�j���� = "") Then
        If (Weekday(dtm���t) = vbMonday) Then
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
Const cst���ʗ琳�a�̋V As Date = "1990/11/12"    '�����V�c

Const cst�����V�c�̑ވ� As Date = "2019/4/30"    '�j���ł͂Ȃ��u�����̋x���v�ł�
Const cst���m�e���̑��� As Date = "2019/5/1"
Const cst2019GW�����̋x�� As Date = "2019/5/2"    '�j���ł͂Ȃ��u�����̋x���v�ł�
Const cst���ʗ琳�a�̋V_���m�e�� As Date = "2019/10/22"

    int�N = Year(���t)
    int�� = Month(���t)
    int�� = Day(���t)

    prv�j�� = ""
    If (���t < cst�j���@�{�s) Then
        Exit Function    ' �j���@�{�s�ȑO
    End If

    Select Case int��
    '-- �P�� --
    Case 1
        If (int�� = 1) Then
            prv�j�� = "����"
        Else
            If (int�N >= 2000) Then
                str��N�j�� = (((int�� - 1) \ 7) + 1) & Weekday(���t)
                If (str��N�j�� = "22") Then  '2nd Monday(2)
                    prv�j�� = "���l�̓�"
                End If
            Else
                If (int�� = 15) Then
                    prv�j�� = "���l�̓�"
                End If
            End If
        End If

    '-- �Q�� --
    Case 2
        If (int�� = 11) Then
            If (int�N >= 1967) Then
                prv�j�� = "�����L�O�̓�"
            End If
        ElseIf (int�� = 23) Then
            If (int�N >= 2020) Then
                prv�j�� = "�V�c�a����"
            End If
        ElseIf (���t = cst���a�V�c�̑�r�̗�) Then
            prv�j�� = "���a�V�c�̑�r�̗�"
        End If

    '-- �R�� --
    Case 3
        If (int�� = prv�t����(int�N)) Then  ' 1948�`2150�ȊO��[99]
            prv�j�� = "�t���̓�"            ' ���Ԃ�̂Ť�K�����ɂȂ�
        End If

    '-- �S�� --
    Case 4
        If (int�� = 29) Then
            If (int�N >= 2007) Then
                prv�j�� = "���a�̓�"
            ElseIf (int�N >= 1989) Then
                prv�j�� = "�݂ǂ�̓�"
            Else
                prv�j�� = "�V�c�a����"    ' ���a�V�c
            End If
        ElseIf (���t = cst�����V�c�̑ވ�) Then    ' 2019/4/30
            prv�j�� = "�����̋x��"    '�j���ɋ��܂ꂽ�����̋x���ł�
        ElseIf (���t = cst���m�e���̌����̋V) Then
            prv�j�� = "�c���q���m�e���̌����̋V"
        End If

    '-- �T�� --
    Case 5
        If (int�� = 3) Then
            prv�j�� = "���@�L�O��"
        ElseIf (int�� = 4) Then
            If (int�N >= 2007) Then
                prv�j�� = "�݂ǂ�̓�"
            ElseIf (int�N >= 1986) Then
                ' 5/4�����j���́w���̓��j�x����j���́w���@�L�O���̐U�֋x���x(�`2006�N)
                If (Weekday(���t) > vbMonday) Then   ' �Ηj �ȍ~(�΁`�y)
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
        Else
            If (int�N = 2019) Then
                If (���t = cst���m�e���̑���) Then    ' 2019/5/1
                    prv�j�� = "���ʂ̓�"    ' ���m�e��
                ElseIf (���t = cst2019GW�����̋x��) Then    ' 2019/5/2
                    prv�j�� = "�����̋x��"    '�j���ɋ��܂ꂽ�����̋x���ł�
                End If
            End If
        End If

    '-- �U�� --
    Case 6
        If (���t = cst���m�e���̌����̋V) Then
            prv�j�� = "�c���q���m�e���̌����̋V"
        End If

    '-- �V�� --
    Case 7
        str��N�j�� = (((int�� - 1) \ 7) + 1) & Weekday(���t)
        Select Case int�N
          Case Is >= 2021
            If (str��N�j�� = "32") Then  '3rd Monday(2)
                prv�j�� = "�C�̓�"
            End If
          Case 2020
            '2020�N�̓I�����s�b�N���[�@�ɂ��
            '�u�C�̓��v�� 7/23 / �u�X�|�[�c�̓��v�� 7/24 �Ɉړ�
            Select Case int��
              Case 23
                prv�j�� = "�C�̓�"
              Case 24
                prv�j�� = "�X�|�[�c�̓�"
              Case Else
            End Select
          Case Is >= 2003
            If (str��N�j�� = "32") Then  '3rd Monday(2)
                prv�j�� = "�C�̓�"
            End If
          Case Is >= 1996
            If (int�� = 20) Then
                prv�j�� = "�C�̓�"
            End If
          Case Else
        End Select

    '-- �W�� --
    Case 8
        Select Case int�N
          Case Is >= 2021
            If (int�� = 11) Then
                prv�j�� = "�R�̓�"
            End If
          Case 2020
            '2020�N�̓I�����s�b�N���[�@�ɂ��u�R�̓��v�� 8/10 �Ɉړ�
            If (int�� = 10) Then
                prv�j�� = "�R�̓�"
            End If
          Case Is >= 2016
            If (int�� = 11) Then
                prv�j�� = "�R�̓�"
            End If
          Case Else
        End Select

    '-- �X�� --
    Case 9
        '��R���j��(15�`21)�ƏH����(22�`24)���d�Ȃ鎖�͂Ȃ�
        int�H���� = prv�H����(int�N)
        If (int�� = int�H����) Then  ' 1948�`2150�ȊO��[99]
            prv�j�� = "�H���̓�"      ' ���Ԃ�̂Ť�K�����ɂȂ�
        Else
            If (int�N >= 2003) Then
                str��N�j�� = (((int�� - 1) \ 7) + 1) & Weekday(���t)
                If (str��N�j�� = "32") Then  '3rd Monday(2)
                    prv�j�� = "�h�V�̓�"
                ElseIf (Weekday(���t) = vbTuesday) Then
                    If (int�� = (int�H���� - 1)) Then
                        prv�j�� = "�����̋x��"  '�Ηj����[�H�����̑O��]
                    End If
                End If
            ElseIf (int�N >= 1966) Then
                If (int�� = 15) Then
                    prv�j�� = "�h�V�̓�"
                End If
            End If
        End If

    '-- �P�O�� --
    Case 10
        str��N�j�� = (((int�� - 1) \ 7) + 1) & Weekday(���t)
        Select Case int�N
          Case Is >= 2021
            If (str��N�j�� = "22") Then  '2nd Monday(2)
                prv�j�� = "�X�|�[�c�̓�"  '2020�N������
            End If
          Case 2020
            '2020�N�̓I�����s�b�N���[�@�ɂ��u�X�|�[�c�̓��v�� 7/24 �Ɉړ�
          Case Is >= 2000
            If (str��N�j�� = "22") Then  '2nd Monday(2)
                prv�j�� = "�̈�̓�"
            ElseIf (���t = cst���ʗ琳�a�̋V_���m�e��) Then
                prv�j�� = "���ʗ琳�a�̋V"    ' ���m�e��(2019/10/22)
            End If
          Case Is >= 1966
            If (int�� = 10) Then
                prv�j�� = "�̈�̓�"
            End If
          Case Else
        End Select

    '-- �P�P�� --
    Case 11
        If (int�� = 3) Then
            prv�j�� = "�����̓�"
        ElseIf (int�� = 23) Then
            prv�j�� = "�ΘJ���ӂ̓�"
        ElseIf (���t = cst���ʗ琳�a�̋V) Then
            prv�j�� = "���ʗ琳�a�̋V"    ' �����V�c
        End If

    '-- �P�Q�� --
    Case 12
        If (int�� = 23) Then
            If ((int�N >= 1989) And (int�N <= 2018)) Then
                prv�j�� = "�V�c�a����"    ' �����V�c
            End If
        End If
    End Select
End Function

'======================================================================
'  �t��/�H�����̗��Z����
'    �w�C��ۈ������H�� ��v�Z������� �V����ݕ֗����x
'  �ŏЉ��Ă��鎮�ł��B
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
'_/  CopyRight(C) K.Tsunoda(AddinBox) 2001 All Rights Reserved.
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/