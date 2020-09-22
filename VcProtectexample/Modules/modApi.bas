Attribute VB_Name = "modApi"
Option Explicit

Private Type NOTIFYICONDATA
 cbSize As Long
 hwnd As Long
 uID As Long
 uFlags As Long
 uCallbackMessage As Long
 hIcon As Long
 szTip As String * 128
 dwState As Long
 dwStateMask As Long
 szInfo As String * 256
 uTimeout As Long
 szInfoTitle As String * 64
 dwInfoFlags As Long
End Type

Private m_IconData As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Sub AddTrayIcon(StrTIP As String, Frm As Form)
On Error Resume Next
 With m_IconData
.cbSize = Len(m_IconData)
.hwnd = Frm.hwnd
.uID = vbNull
.uFlags = &H2 Or &H10 Or &H1 Or &H4
.uCallbackMessage = &H200
.hIcon = Frm.Icon
.szTip = StrTIP & vbNullChar
.dwState = 0
.dwStateMask = 0
End With
Shell_NotifyIcon &H0, m_IconData
End Sub

Public Sub ModifyTrayIcon(StrTIP As String, PB As PictureBox)
On Error Resume Next
With m_IconData
.hIcon = PB.Picture
.szTip = StrTIP & vbNullChar
End With
Shell_NotifyIcon &H1, m_IconData
End Sub

Public Sub DeleteTrayIcon()
On Error Resume Next
Shell_NotifyIcon &H2, m_IconData
End Sub
