Attribute VB_Name = "SysTray"
Option Explicit

Public Const NIM_ADD As Long = &H0
Public Const NIM_MODIFY As Long = &H1
Public Const NIM_DELETE As Long = &H2

Public Const NIF_ICON As Long = &H2
Public Const NIF_TIP As Long = &H4
Public Const NIF_MESSAGE As Long = &H1

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP As Long = &H202
Public Const WM_LBUTTONDBLCLK As Long = &H203

Public Const WM_MBUTTONDOWN As Long = &H207
Public Const WM_MBUTTONUP As Long = &H208
Public Const WM_MBUTTONDBLCLK As Long = &H209

Public Const WM_RBUTTONDOWN As Long = &H204
Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_RBUTTONDBLCLK As Long = &H206

Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Public nID As NOTIFYICONDATA

Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Sub ShellTrayAdd(ByVal ToolTip As String, frm As Form)
    Dim r As Long
    
    With nID
        .cbSize = LenB(nID)
        .hwnd = frm.hwnd
        .uID = 125&
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = frm.Icon
        .szTip = ToolTip & Chr$(0)
     End With
     r = Shell_NotifyIcon(NIM_ADD, nID)
End Sub

Public Sub ShellTrayModify(ByVal ToolTip As String, frm As Form)
  Dim r As Long

    With nID
        .cbSize = LenB(nID)
        .hwnd = frm.hwnd
        .uID = 125&
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = frm.Icon
        .szTip = ToolTip & Chr$(0)
     End With
     r = Shell_NotifyIcon(NIM_MODIFY, nID)
End Sub

Public Sub ShellTrayRemove()
    Call Shell_NotifyIcon(NIM_DELETE, nID)
End Sub
