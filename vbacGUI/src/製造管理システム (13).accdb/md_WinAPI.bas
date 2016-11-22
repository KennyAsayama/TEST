Option Compare Database
Option Explicit
'Ver1.01.1 20150910 K.Asayama ADD
'   *************************************************************
'   Windowサイズ変更API
'
'
'    Input項目
'       hwnd        Windowハンドル(Application.hWndAccessApp)
'       nCmdShow    Windowサイズ
'                   1   →ノーマル
'                   2   →最小化
'                   3   →最大化
'
'   *************************************************************
Declare Function ShowWindow Lib "user32" _
       (ByVal hwnd As Long, _
        ByVal nCmdShow As Long) As Long


'Ver1.01.* 201510** K.Asayama ADD
'   *************************************************************
'   Windowサイズウィンドウの位置と大きさを取得、変更
'
'
'    Input項目
'       hwnd        Windowハンドル(Application.hWndAccessApp)
'       X           横位置（ピクセル）
'       Y           縦位置（ピクセル）
'       nWidth      幅（ピクセル）
'       nHeight     高さ（ピクセル）
'       nHeight     再描画
'       bRepaint    True   →ウィンドウを再描画する
'                   False  →しない
'
'       ※64bit用はACCDBでのみ使用可能
'   *************************************************************
#If VBA7 Then
'64ビット用
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
'32ビット用
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

               
'RECT構造体
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

'   *************************************************************
'   スクリーンショット用キー受信API
'
'
'1.10.14 K.Asayama ADD
'   *************************************************************
'64bit版
#If VBA7 Then
Public Declare PtrSafe Sub keybd_event Lib "user32" ( _
    ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, _
    ByVal dwExtraInfo As Long _
        )
#Else
'32bit版
Public Declare Sub keybd_event Lib "user32" ( _
    ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, _
    ByVal dwExtraInfo As Long _
        )
#End If

'1.11.1 ADD
'Sleep関数
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)