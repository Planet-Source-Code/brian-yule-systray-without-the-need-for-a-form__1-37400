VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskBarObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type APPBARDATA
    cbSize As Long
    hWnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long
End Type

Private Enum Taskbar
    Autohidden = &H1
    AlwaysOnTop = &H2
    GetState = &H4
    GetPos = &H5
End Enum

Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long

Public Property Get Left() As Long
    Dim ABD As APPBARDATA, Ret As Long
    ABD.cbSize = LenB(ABD)
    Call SHAppBarMessage(Taskbar.GetPos, ABD)
    Left = ABD.rc.Left
End Property

Public Property Get Top() As Long
    Dim ABD As APPBARDATA
    ABD.cbSize = LenB(ABD)
    Call SHAppBarMessage(Taskbar.GetPos, ABD)
    Top = ABD.rc.Top
End Property

Public Property Get Height() As Long
    Dim ABD As APPBARDATA
    ABD.cbSize = LenB(ABD)
    Call SHAppBarMessage(Taskbar.GetPos, ABD)
    Height = ABD.rc.Bottom - ABD.rc.Top
End Property

Public Property Get Width() As Long
    Dim ABD As APPBARDATA
    ABD.cbSize = LenB(ABD)
    Call SHAppBarMessage(Taskbar.GetPos, ABD)
    Width = ABD.rc.Right - ABD.rc.Left
End Property

Public Property Get Autohide() As Boolean
    Dim ABD As APPBARDATA, Ret As Long
    ABD.cbSize = LenB(ABD)
    Ret = SHAppBarMessage(Taskbar.GetState, ABD)
    If Ret And Taskbar.Autohidden Then Autohide = True
End Property

Public Property Get OnTop() As Boolean
    Dim ABD As APPBARDATA, Ret As Long
    ABD.cbSize = LenB(ABD)
    Ret = SHAppBarMessage(Taskbar.GetState, ABD)
    If Ret And Taskbar.AlwaysOnTop Then OnTop = True
End Property

Public Property Get Alignment() As String
    If Me.Left < 0 Then
        If Me.Top < 0 Then
            If Me.Width > Me.Height Then
                Alignment = "Top"
            ElseIf Me.Width < Me.Height Then
                Alignment = "Left"
            End If
        Else
            Alignment = "Bottom"
        End If
    Else
        Alignment = "Right"
    End If
End Property
