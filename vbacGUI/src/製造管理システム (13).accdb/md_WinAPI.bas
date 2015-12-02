Attribute VB_Name = "md_WinAPI"
Option Compare Database
Option Explicit
'Ver1.01.1 20150910 K.Asayama ADD
'   *************************************************************
'   Window�T�C�Y�ύXAPI
'
'
'    Input����
'       hwnd        Window�n���h��(Application.hWndAccessApp)
'       nCmdShow    Window�T�C�Y
'                   1   ���m�[�}��
'                   2   ���ŏ���
'                   3   ���ő剻
'
'   *************************************************************
Declare Function ShowWindow Lib "user32" _
       (ByVal hwnd As Long, _
        ByVal nCmdShow As Long) As Long


'Ver1.01.* 201510** K.Asayama ADD
'   *************************************************************
'   Window�T�C�Y�E�B���h�E�̈ʒu�Ƒ傫�����擾�A�ύX
'
'
'    Input����
'       hwnd        Window�n���h��(Application.hWndAccessApp)
'       X           ���ʒu�i�s�N�Z���j
'       Y           �c�ʒu�i�s�N�Z���j
'       nWidth      ���i�s�N�Z���j
'       nHeight     �����i�s�N�Z���j
'       nHeight     �ĕ`��
'       bRepaint    True   ���E�B���h�E���ĕ`�悷��
'                   False  �����Ȃ�
'
'       ��64bit�p��ACCDB�ł̂ݎg�p�\
'   *************************************************************
#If VBA7 Then
'64�r�b�g�p
Public Declare PtrSafe Function MoveWindow _
           Lib "user32" _
        (ByVal hwnd As Long _
       , ByVal X As Long _
       , ByVal Y As Long _
       , ByVal nWidth As Long _
       , ByVal nHeight As Long _
       , ByVal bRepaint As Long) As Long

Public Declare PtrSafe Function GetWindowRect _
           Lib "user32" _
        (ByVal hwnd As Long _
       , ByRef lpRect As RECT) As Long
#Else
'32�r�b�g�p
Public Declare Function MoveWindow _
           Lib "user32" _
          (ByVal hwnd As Long _
       , ByVal X As Long _
       , ByVal Y As Long _
       , ByVal nWidth As Long _
       , ByVal nHeight As Long _
       , ByVal bRepaint As Long) As Long

Public Declare Function GetWindowRect _
           Lib "user32" _
        (ByVal hwnd As Long _
       , ByRef lpRect As RECT) As Long
#End If

               
'RECT�\����
Public Type RECT
       left   As Long
       top    As Long
       right  As Long
       bottom As Long
End Type

Public Const LOGPIXELSX = 88
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

