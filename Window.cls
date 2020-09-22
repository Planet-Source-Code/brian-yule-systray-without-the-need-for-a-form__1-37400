VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Constants
Private Const WS_OVERLAPPED = &H0&

' Types (Structs)
Private Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long       ' Application instance number
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    Y As Long
    X As Long
    style As Long
    lpszName As String
    lpszClass As String
    ExStyle As Long
End Type

Private Type WNDCLASS
    style As Long
    lpfnwndproc As Long     ' Pointer to Call back procedure
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long       ' Application instance number
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String ' Window Class Name
End Type

' API calls
Private Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

' Public Events
Public Event Incomming(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Event Error(ByVal ErrNum As Long, ByVal ErrDesc As String)

' Global Variables
Public myhWnd As Long
Private WndCls As WNDCLASS

Public Function HookWindow()
    On Error GoTo ErrorHandler
    
    ' Create Window
    Dim CS As CREATESTRUCT
    
    Let myhWnd = CreateWindowEx(0, WndCls.lpszClassName, "MyWindow", WS_OVERLAPPED, 0, 0, 0, 0, 0, 0, App.hInstance, CS)
Exit Function
ErrorHandler:
    RaiseEvent Error(Err.Number, Err.Description)
End Function

Public Function UnHookWindow()
    On Error GoTo ErrorHandler
    
    ' Remove Window
    Call DestroyWindow(myhWnd)
    
    Let myhWnd = 0 ' Reset Window handle
Exit Function
ErrorHandler:
    RaiseEvent Error(Err.Number, Err.Description)
End Function

Friend Sub Receive(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    On Error GoTo ErrorHandler
    
    RaiseEvent Incomming(hWnd, uMsg, wParam, lParam) ' Send call back information to parent class
Exit Sub
ErrorHandler:
    RaiseEvent Error(Err.Number, Err.Description)
End Sub

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Let CallBack.AppObjPtr = ObjPtr(Me) ' Pointer to this class
    
    ' Register Window Class
    Let WndCls.hInstance = App.hInstance
    Let WndCls.lpfnwndproc = CLng(Val(AddressOf CallBack.CallBack))
    Let WndCls.lpszClassName = "clsCallBackWindow"
    
    Call RegisterClass(WndCls)
Exit Sub
ErrorHandler:
    RaiseEvent Error(Err.Number, Err.Description)
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    ' Remove Window
    If myhWnd <> 0 Then Call DestroyWindow(myhWnd)
    
    ' Remove Window Class
    Call UnregisterClass(WndCls.lpszClassName, App.hInstance)
Exit Sub
ErrorHandler:
    RaiseEvent Error(Err.Number, Err.Description)
End Sub

