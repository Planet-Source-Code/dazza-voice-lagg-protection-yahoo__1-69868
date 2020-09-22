VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   750
   ClientLeft      =   79335
   ClientTop       =   14715
   ClientWidth     =   1560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   2520
      Width           =   375
      Begin VB.PictureBox Picture3 
         Height          =   375
         Left            =   840
         Picture         =   "Form1.frx":08E1
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Left            =   1200
         Picture         =   "Form1.frx":11AB
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         Height          =   375
         Left            =   1680
         Picture         =   "Form1.frx":1A75
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   360
         Top             =   240
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   360
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Project1.VcHook VcHook1 
         Left            =   1080
         Top             =   720
         _ExtentX        =   741
         _ExtentY        =   741
      End
      Begin Project1.YTCP YTCP1 
         Left            =   1560
         Top             =   720
         _ExtentX        =   741
         _ExtentY        =   741
      End
   End
   Begin VB.Menu mnumain 
      Caption         =   " "
      Visible         =   0   'False
      Begin VB.Menu EnableProtect 
         Caption         =   "Enable Protect"
         Enabled         =   0   'False
      End
      Begin VB.Menu DisableProtect 
         Caption         =   "Disable Protect"
      End
      Begin VB.Menu spcr2 
         Caption         =   "-"
      End
      Begin VB.Menu OpenClient 
         Caption         =   "Open Voice Client"
      End
      Begin VB.Menu SetClient 
         Caption         =   "Set Client To Protect"
      End
      Begin VB.Menu spcr1 
         Caption         =   "-"
      End
      Begin VB.Menu ExitApplication 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private YVcServer As String, YVcPort As String, YVcApp As String, YVcPath As String, LibPath As String

Private Sub DisableProtect_Click()
On Error Resume Next
Protect_off
DisableProtect.Enabled = False
EnableProtect.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub EnableProtect_Click()
On Error Resume Next
If YVcApp = "" Then GoTo 1
Protect_on
DisableProtect.Enabled = True
EnableProtect.Enabled = False
Timer1.Enabled = True
Exit Sub
1
DisableProtect_Click
SetClient_Click
End Sub

Private Sub ExitApplication_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then End
On Error Resume Next
LibPath = App.Path & "\vchook.dll"
YVcServer = GetSetting("YVCPRO", "OPTIONS", "HOST", "vc.yahoo.com")
YVcPort = GetSetting("YVCPRO", "OPTIONS", "PORT", "5001")
YVcApp = GetSetting("YVCPRO", "OPTIONS", "EXE")
YVcPath = GetSetting("YVCPRO", "OPTIONS", "PATH")
If YVcServer = "" Then YVcServer = "vc.yahoo.com"
If YVcPort = "" Then YVcPort = "5001"
If Dir(LibPath, vbNormal) = "" Then MsgBox "Unable to locate: " & LibPath: End
AddTrayIcon "Voice Protect - Ready", Me
EnableProtect_Click
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim Action As Integer, Result As String
If Me.ScaleMode = vbPixels Then
Action = X
Else
Action = X / Screen.TwipsPerPixelX
End If
Select Case X
Case 7725
Me.WindowState = vbNormal
Result = SetForegroundWindow(Me.hwnd)
GoTo 1
Case 7755
Me.WindowState = vbNormal
Result = SetForegroundWindow(Me.hwnd)
PopupMenu mnumain
GoTo 1
End Select
1
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
VcHook1.StopService
SaveSetting "YVCPRO", "OPTIONS", "HOST", YVcServer
SaveSetting "YVCPRO", "OPTIONS", "PORT", YVcPort
SaveSetting "YVCPRO", "OPTIONS", "EXE", YVcApp
SaveSetting "YVCPRO", "OPTIONS", "PATH", YVcPath
DeleteTrayIcon
End
End Sub

Private Sub Protect_on()
On Error Resume Next
VcHook1.StartService YVcServer, YVcPort
End Sub

Private Sub Protect_off()
On Error Resume Next
VcHook1.StopService
ModifyTrayIcon "Voice Protect - Disabled", Picture1
End Sub

Private Sub OpenClient_Click()
On Error Resume Next
EnableProtect_Click
Call OpenApplication(YVcPath)
End Sub

Private Sub SetClient_Click()
On Error Resume Next
Dim PFPath As String, PFName As String
CD1.DialogTitle = "Locate Your Chat/Voice Client"
CD1.Filter = "Chat/Voice Client EXE|*.exe|"
CD1.ShowOpen
PFName = CD1.FileTitle
PFPath = CD1.FileName
If PFName = "" Or PFPath = "" Then Exit Sub
DisableProtect_Click
YVcPath = PFPath
YVcApp = PFName
EnableProtect_Click
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
InsertDll YVcApp, LibPath
End Sub

Private Sub InsertDll(TheExe As String, TheDll As String)
On Error Resume Next
ProsH = GetHProcExe(TheExe)
If ProsH = 0 Then
Exit Sub
Else
Timer1.Enabled = False
InjectDll TheDll, ProsH
End If
End Sub

Private Sub YTCP1_ErrorOnVoice()
On Error Resume Next
ModifyTrayIcon "Voice Protect - Error", Picture3
End Sub

Private Sub YTCP1_LoggedInVoice()
On Error Resume Next
ModifyTrayIcon "Voice Protect - " & YVcApp, Picture2
End Sub

Private Sub YTCP1_LoggedOutVoice()
On Error Resume Next
ModifyTrayIcon "Voice Protect - Disabled", Picture1
End Sub
