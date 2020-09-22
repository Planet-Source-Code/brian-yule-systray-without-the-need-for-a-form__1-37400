VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Tray Icon Chooser"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1875
      ScaleWidth      =   5235
      TabIndex        =   4
      Top             =   720
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Balloon"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdRed 
      Caption         =   "Red"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdYellow 
      Caption         =   "Yellow"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdGreen 
      Caption         =   "Green"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents sysTray As SystemTray.Application
Private sysTray As Object

Private Sub cmdGreen_Click()
    sysTray.UpdateIcon (App.Path & "\Trffc10a.ico")
End Sub

Private Sub cmdRed_Click()
    sysTray.UpdateIcon (App.Path & "\Trffc10c.ico")
End Sub

Private Sub cmdYellow_Click()
    sysTray.UpdateIcon (App.Path & "\Trffc10b.ico")
End Sub

Private Sub Command1_Click()
    Picture1.Cls
    Picture1.Print "Left", sysTray.Taskbar.Left
    Picture1.Print "Width", sysTray.Taskbar.Width
    Picture1.Print "Height", sysTray.Taskbar.Height
    Picture1.Print "Top", sysTray.Taskbar.Top
    Picture1.Print "OnTop", sysTray.Taskbar.OnTop
    Picture1.Print "Autohide", sysTray.Taskbar.Autohide
    
    Me.ScaleMode = vbPixels
    Picture1.Print "ScaleWidth", Me.ScaleWidth
    Picture1.Print "ScaleHeight", Me.ScaleHeight
    
    Select Case sysTray.Taskbar.Alignment
        Case "Top"
            Let Me.Left = (sysTray.Taskbar.Width + 1) * Screen.TwipsPerPixelX - Me.Width
            Let Me.Top = (sysTray.Taskbar.Height - 1) * Screen.TwipsPerPixelY
        Case "Left"
            Let Me.Left = (sysTray.Taskbar.Width - 1) * Screen.TwipsPerPixelX
            Let Me.Top = (sysTray.Taskbar.Height + 1) * Screen.TwipsPerPixelY - Me.Height
        Case "Bottom"
            Let Me.Left = (sysTray.Taskbar.Width + 1) * Screen.TwipsPerPixelX - Me.Width
            Let Me.Top = (sysTray.Taskbar.Top - 1) * Screen.TwipsPerPixelY - Me.Height
        Case "Right"
            Let Me.Left = (sysTray.Taskbar.Left + 1) * Screen.TwipsPerPixelX - Me.Width
            Let Me.Top = (sysTray.Taskbar.Height - 1) * Screen.TwipsPerPixelY - Me.Height
    End Select
    
    Call sysTray.ShowBalloon("Hello", &H1, "www.braxtel.com")
End Sub

Private Sub Form_Load()
    'Set sysTray = New SystemTray.Application
    Set sysTray = CreateObject("SystemTray.Application")
    Call sysTray.CreateIcon(App.Path & "\Trffc10c.ico", "My Icon")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call sysTray.DeleteIcon
    Set sysTray = Nothing
End Sub

Private Sub sysTray_balloonClick()
'
    Debug.Print "sysTray_balloonClick"
End Sub

Private Sub sysTray_balloonHide()
'
    Debug.Print "sysTray_balloonHide"
End Sub

Private Sub sysTray_balloonShow()
'
    Debug.Print "sysTray_balloonShow"
End Sub

Private Sub sysTray_balloonTimeout()
'
    Debug.Print "sysTray_balloonTimeout"
End Sub

Private Sub sysTray_ButtonDown(ByVal Button As Integer)
'
    Debug.Print "sysTray_ButtonDown"
End Sub

Private Sub sysTray_ButtonUp(ByVal Button As Integer)
'
    Debug.Print "sysTray_ButtonUp"
End Sub

Private Sub sysTray_DblClick(ByVal Button As Integer)
'
    Debug.Print "sysTray_DblClick"
End Sub

Private Sub sysTray_MouseMove(ByVal x As Long, ByVal Y As Long)
'
    Debug.Print "sysTray_MouseMove"; x, Y
End Sub
