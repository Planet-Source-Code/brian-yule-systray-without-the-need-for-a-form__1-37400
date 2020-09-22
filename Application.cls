VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'   Project:    System Tray
'
'   Disclaimer:
'   The use of this code is permitted.  I do not accept any consequences from the use of
'   this code or this application.
'
'   Creator:    Brian Yule
'   Email:      byule@braxtel.com
'   Created:    16/07/2002
'   Modified:   16/07/2002
'
'   Reason:
'   The system tray needs to associate a window with the tray icon to send the call back message.
'   We need to have a system tray icon that doesn't need a window handle so that we can show the
'   state of a process without a user interface.  Also we need to know the position of the task bar
'   so that we can place our form in the correct position.
'
'   Call Back Flow:
'   Register a window class
'   Create a window with this class
'   Get call back information for the window through the call back procedure
'   Send the call back information into the window class
'   Root the call back information required into the application class
'   Send events back to parent application
'
'   Application input:
'   Method 1: Create the icon in the system tray
'   Method 2: Update the icon in the system tray
'   Method 3: Update the tooltip of the icon in the system tray
'   Method 4: Delete the icon in the system tray
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit




' Return Events
Public Event MouseMove(ByVal X As Long, ByVal Y As Long)
Public Event DblClick(ByVal Button As Integer)
Public Event ButtonUp(ByVal Button As Integer)
Public Event ButtonDown(ByVal Button As Integer)
Public Event Error(ByVal ErrNum As Long, ByVal ErrDesc As String)

Public Event BalloonHide()
Public Event BalloonShow()
Public Event BalloonTimeout()
Public Event BalloonClick()

Private WithEvents wnd As Window
Attribute wnd.VB_VarHelpID = -1
Public Taskbar As New TaskBarObj

Public Sub ShowBalloon(strTitle As String, IconType As Long, strMessage As String)

   With nidBalloon
      .cbSize = Len(nidBalloon)
      .hWnd = wnd.myhWnd
      .uId = App.hInstance
      .uFlags = Icon_Flags.Info
      .dwInfoFlags = IconType
      .szInfoTitle = strTitle
      .szInfo = strMessage
   End With

   Call Shell_NotifyIcon(Action.MODIFY, nidBalloon)
End Sub

Private Function GetImage(ByVal fileName As String) As Long
    On Error GoTo ErrorHandler
    
    GetImage = LoadImage(App.hInstance, fileName, ImageTypes.Icon, 0, 0, LoadTypes.FROMFILE)
    
    If GetImage = 0 Then RaiseEvent Error(1, "Load Icon Failed")
Exit Function
ErrorHandler:
    RaiseEvent Error(Err.Number, Err.Description)
End Function

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Set wnd = New Window
    
    With nidProgramData
        Let .cbSize = Len(nidProgramData)
        Let .uId = App.hInstance
        Let .uFlags = Icon_Flags.Icon Or Icon_Flags.Tip Or Icon_Flags.Message
        Let .uCallbackMessage = WindowsMessage.WM_MOUSEMOVE
        Let .dwState = IconState.SHAREDICON
    End With
Exit Sub
ErrorHandler:
    RaiseEvent Error(Err.Number, Err.Description)
End Sub

Public Sub CreateIcon(ByVal fileName As String, Optional ByVal Tooltip As String)
    On Error GoTo ErrorHandler
    
    Call wnd.HookWindow   ' Hook all System Tray Messages
    
    With nidProgramData
        Let .hWnd = wnd.myhWnd
        Let .hIcon = GetImage(fileName)
        Let .szTip = Tooltip & vbNullChar
    End With
    
    Call Shell_NotifyIcon(Action.Add, nidProgramData) ' Add the system tray icon
    Call Shell_NotifyIcon(Action.SetVersion, nidProgramData)
    
    ' Register to receive the message that the taskbar has been created.
    ' We receive this message when Explorer has crashed and restarted.
    CallBack.M_TaskbarRestart = RegisterWindowMessage("TaskbarCreated")
Exit Sub
ErrorHandler:
    RaiseEvent Error(Err.Number, Err.Description)
End Sub

Private Sub ReCreateIcon()
    On Error GoTo ErrorHandler
    
    Call Shell_NotifyIcon(Action.Add, nidProgramData) ' Add the system tray icon
    Call Shell_NotifyIcon(Action.SetVersion, nidProgramData)
Exit Sub
ErrorHandler:
    RaiseEvent Error(Err.Number, Err.Description)
End Sub

Public Sub UpdateIcon(ByVal fileName As String)
    On Error GoTo ErrorHandler
    
    Let nidProgramData.hIcon = GetImage(fileName)
    
    Call Shell_NotifyIcon(Action.MODIFY, nidProgramData) ' Modify the file Name of the system tray icon
Exit Sub
ErrorHandler:
    RaiseEvent Error(Err.Number, Err.Description)
End Sub

Public Sub UpdateTooltip(ByVal Tooltip As String)
    On Error GoTo ErrorHandler
    
    Let nidProgramData.szTip = Tooltip & vbNullChar
    
    Call Shell_NotifyIcon(Action.MODIFY, nidProgramData) ' Modify the Tooltip of the system tray icon
Exit Sub
ErrorHandler:
    RaiseEvent Error(Err.Number, Err.Description)
End Sub



Public Sub DeleteBalloon(title As String, IconType As Long, strMessage As String)
   Call Shell_NotifyIcon(Action.Delete, nidBalloon)
End Sub

Public Sub DeleteIcon()
    On Error GoTo ErrorHandler
    
    Call wnd.UnHookWindow ' Stop receiving messages from the Window
    Call Shell_NotifyIcon(Action.Delete, nidProgramData) ' Delete System Tray icon
Exit Sub
ErrorHandler:
    RaiseEvent Error(Err.Number, Err.Description)
End Sub

Private Sub wnd_Error(ByVal ErrNum As Long, ByVal ErrDesc As String)
    On Error GoTo ErrorHandler
    
    RaiseEvent Error(ErrNum, ErrDesc)
Exit Sub
ErrorHandler:
    Resume Next
End Sub

Private Sub wnd_Incomming(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    On Error GoTo ErrorHandler
    
    ' Select which Window Event was received and send event to parent application
    If uMsg = WindowsMessage.WM_USER_SYSTRAY Then ' System Tray Message
        Select Case lParam
            Case WindowsMessage.WM_LBUTTONDBLCLK: RaiseEvent DblClick(1)
            Case WindowsMessage.WM_LBUTTONDOWN: RaiseEvent ButtonDown(1)
            Case WindowsMessage.WM_LBUTTONUP:  RaiseEvent ButtonUp(1)
            Case WindowsMessage.WM_RBUTTONDOWN: RaiseEvent ButtonDown(2)
            Case WindowsMessage.WM_RBUTTONUP: RaiseEvent ButtonUp(2)
            Case WindowsMessage.WM_RBUTTONDBLCLK: RaiseEvent DblClick(2)
            Case WindowsMessage.WM_MOUSEMOVE
                Dim pt As POINTAPI
                    
                Call GetCursorPos(pt) ' Get the cursor position of the mouse over the system tray icon
                
                RaiseEvent MouseMove(pt.X, pt.Y)
        End Select
    ' The Balloon uses the mousemove message to return balloon window events.
    ElseIf uMsg = WM_MOUSEMOVE Then
        Debug.Print uMsg, wParam, lParam
        Select Case lParam
            Case BalloonMessage.BalloonHide: RaiseEvent BalloonHide
            Case BalloonMessage.BalloonShow: RaiseEvent BalloonShow
            Case BalloonMessage.BalloonTimeout: RaiseEvent BalloonTimeout
            Case BalloonMessage.balloonUserClick: RaiseEvent BalloonClick
        End Select
    ' If explorer has crashed and restarted then recreate the icon in the system tray
    ElseIf uMsg = CallBack.M_TaskbarRestart Then ' TaskbarCreated Message
        Call ReCreateIcon
    End If
Exit Sub
ErrorHandler:
    Resume Next
End Sub
