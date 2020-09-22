Attribute VB_Name = "CallBack"
Option Explicit

Public Const NOTIFYICON_VERSION = &H3

' Enumerations


Public Enum WindowsMessage
    WM_MOUSEMOVE = &H200       ' Mouse Move
    WM_LBUTTONDOWN = &H201     ' Left Button Down
    WM_LBUTTONUP = &H202       ' Left Button Up
    WM_LBUTTONDBLCLK = &H203   ' Left Button Double Click
    WM_RBUTTONDOWN = &H204     ' Right Button Down
    WM_RBUTTONUP = &H205       ' Right Button Up
    WM_RBUTTONDBLCLK = &H206   ' Right Button Double Click
    WM_USER = &H400
    WM_USER_SYSTRAY = WM_USER + &H5
End Enum

Public Enum BalloonMessage
    balloonShow = WM_USER + 2
    balloonHide = WM_USER + 3
    balloonTimeout = WM_USER + 4
    balloonUserClick = WM_USER + 5
End Enum

Public Enum IconState
    Hidden = &H1
    SHAREDICON = &H2
End Enum

Public Enum Action ' Actions for dealing with the system Tray
    Add = &H0       ' Add a system tray icon
    MODIFY = &H1    ' Modify a system tray icon
    Delete = &H2    ' Delete a system tray icon
    SetFocus = &H3
    SetVersion = &H4
    Version = &H5
End Enum
 
Public Enum Info_Flags
    NONE = &H0 ' No icon.
    Info = &H1 ' An information icon.
    warning = &H2 ' A warning icon.0
    Error = &H3 ' An error icon.
    GUID = &H5
    ICON_MASK = &HF ' Version 6.0. Reserved.
    NOSOOUND = &H10 ' Version 6.0. Do not play the associated sound. Applies only to balloon ToolTips.
End Enum

Public Enum Icon_Flags  ' Flags you can set on the system tray
    Message = &H1   ' System Messages
    Icon = &H2      ' Icon
    Tip = &H4       ' Tooltip
    State = &H8
    Info = &H10
End Enum

Public Enum ImageTypes ' Image Types
    IBITMAP = 0         ' Bitmap
    Icon                ' Icon
    CURSOR              ' Cursor
End Enum

Public Enum LoadTypes      ' Load Types
    FROMFILE = &H10         ' From file
    TRANSPARENT = &H20      '
    MAP3DCOLORS = &H1000    '
End Enum

' Types (Structs)
Public Type POINTAPI       ' Point Type
    X As Long               ' X Coordinate
    Y As Long               ' Y Coordinate
End Type

Public Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Public Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uId As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type

' API Calls
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Action, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal iType As Long, ByVal cx As Long, ByVal cy As Long, ByVal fOptions As Long) As Long
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal Length As Long)
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' Global Variables
Public AppObjPtr As Long
Public PrevProc As Long
Public M_TaskbarRestart As Long
Public nidProgramData As NOTIFYICONDATA
Public nidBalloon As NOTIFYICONDATA

Public Function CallBack(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error GoTo ErrorHandler
    
    If uMsg = WindowsMessage.WM_USER_SYSTRAY Or uMsg = M_TaskbarRestart Or uMsg = WM_MOUSEMOVE Then
        Dim objTemp As Window
        
        Call CopyMemory(objTemp, AppObjPtr, 4)
        Call objTemp.Receive(hWnd, uMsg, wParam, lParam)
        Call CopyMemory(objTemp, 0&, 4)
    End If
    
    CallBack = DefWindowProc(hWnd, uMsg, wParam, lParam)
Exit Function
ErrorHandler:
    Resume Next
End Function
