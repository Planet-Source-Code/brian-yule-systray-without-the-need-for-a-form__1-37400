VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BalloonObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub IconInfo(strTitle As String, IconType As Long, strMessage As String)

   With nidBalloon
      .cbSize = Len(nidBalloon)
      .hWnd = wnd.myhWnd
      .uId = App.hInstance
      .uFlags = Info_Flags.Info
      .dwInfoFlags = IconType
      .szInfoTitle = strTitle
      .szInfo = strMessage
   End With

   Call Shell_NotifyIcon(Action.MODIFY, nidBalloon)
End Sub
Public Sub IconError(strTitle As String, IconType As Long, strMessage As String)

   With nidBalloon
      .cbSize = Len(nidBalloon)
      .hWnd = wnd.myhWnd
      .uId = App.hInstance
      .uFlags = Info_Flags.Error
      .dwInfoFlags = IconType
      .szInfoTitle = strTitle
      .szInfo = strMessage
   End With

   Call Shell_NotifyIcon(Action.MODIFY, nidBalloon)
End Sub
Public Sub IconNone(strTitle As String, IconType As Long, strMessage As String)

   With nidBalloon
      .cbSize = Len(nidBalloon)
      .hWnd = wnd.myhWnd
      .uId = App.hInstance
      .uFlags = Info_Flags.NONE
      .dwInfoFlags = IconType
      .szInfoTitle = strTitle
      .szInfo = strMessage
   End With

   Call Shell_NotifyIcon(Action.MODIFY, nidBalloon)
End Sub
Public Sub Warning(strTitle As String, IconType As Long, strMessage As String)

   With nidBalloon
      .cbSize = Len(nidBalloon)
      .hWnd = wnd.myhWnd
      .uId = App.hInstance
      .uFlags = Info_Flags.Warning
      .dwInfoFlags = IconType
      .szInfoTitle = strTitle
      .szInfo = strMessage
   End With

   Call Shell_NotifyIcon(Action.MODIFY, nidBalloon)
End Sub
