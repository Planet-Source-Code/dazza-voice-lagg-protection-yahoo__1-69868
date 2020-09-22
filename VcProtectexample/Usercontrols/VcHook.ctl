VERSION 5.00
Begin VB.UserControl VcHook 
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   InvisibleAtRuntime=   -1  'True
   Picture         =   "VcHook.ctx":0000
   ScaleHeight     =   510
   ScaleWidth      =   495
   ToolboxBitmap   =   "VcHook.ctx":04A8
End
Attribute VB_Name = "VcHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event CurrentIP(ByVal IpAddress As String)
Private WithEvents Winsock1 As Winsock
Attribute Winsock1.VB_VarHelpID = -1
Private WithEvents Winsock2 As Winsock
Attribute Winsock2.VB_VarHelpID = -1
Private MyRoomServer As String, MySourceID As String

Public Sub StartService(ByVal VoiceServer As String, ByVal VoicePort As String)
On Error Resume Next
Winsock1.Close
Winsock2.Close
Winsock1.Protocol = sckTCPProtocol
Winsock2.Protocol = sckTCPProtocol
Winsock2.RemoteHost = VoiceServer
Winsock1.LocalPort = VoicePort
Winsock1.Listen
End Sub

Public Sub StopService()
On Error Resume Next
Winsock1.Close
Winsock2.Close
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
Set Winsock1 = New Winsock
Set Winsock2 = New Winsock
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
UserControl.Width = 420
UserControl.Height = 420
End Sub

Private Sub UserControl_Terminate()
On Error Resume Next
Set Winsock1 = Nothing
Set Winsock2 = Nothing
End Sub

Private Sub Winsock1_Close()
On Error Resume Next
Winsock1.Close
Winsock2.Close
Form1.YTCP1.StopService
Winsock1.Listen
End Sub

Private Sub Winsock2_Close()
On Error Resume Next
Winsock1.Close
Winsock2.Close
Form1.YTCP1.StopService
Winsock1.Listen
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
Winsock2.Close
Winsock2.RemotePort = Winsock1.LocalPort
Winsock2.Connect
Do
DoEvents
Loop Until Winsock2.State = sckConnected
Winsock1.Close
Winsock1.Accept requestID
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Winsock1.Close
Winsock2.Close
End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Winsock1.Close
Winsock2.Close
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String, DataLength As Integer, TmpData As String, HeaderLength As Integer
HeaderLength = 4
With Winsock1
While .BytesReceived >= HeaderLength
Call .PeekData(Data, vbString, HeaderLength)
DataLength = (256 * Asc(Mid(Data, 1, 1)) + Asc(Mid(Data, 2, 1)))
If DataLength <= .BytesReceived Then
Call .GetData(TmpData, vbString, DataLength)
ParseVoiceClient TmpData
Else
Exit Sub
End If
DoEvents
Wend
End With
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String, DataLength As Integer, TmpData As String, HeaderLength As Integer
HeaderLength = 4
With Winsock2
While .BytesReceived >= HeaderLength
Call .PeekData(Data, vbString, HeaderLength)
DataLength = (256 * Asc(Mid(Data, 1, 1)) + Asc(Mid(Data, 2, 1)))
If DataLength <= .BytesReceived Then
Call .GetData(TmpData, vbString, DataLength)
ParseVoiceServer TmpData
Else
Exit Sub
End If
DoEvents
Wend
End With
End Sub

Private Sub ParseVoiceClient(Data As String)
On Error Resume Next
Dim TcpLen As Integer, HdrLen As Integer, CommandType As Integer, PcktType As Integer
TcpLen = (256 * Asc(Mid(Data, 1, 1)) + Asc(Mid(Data, 2, 1)))
HdrLen = Asc(Mid(Data, 8, 1)) + 4
CommandType = Asc(Mid(Data, HdrLen + 2, 1))
Select Case CommandType
Case Is = 204
PcktType = Asc(Mid(Data, HdrLen + 14, 1))
Select Case PcktType
Case Is = 1
Winsock2.SendData Data
GoTo 1
Case Is = 7
Winsock2.SendData Data
GoTo 1
Case Is = 13
Winsock2.SendData Data
GoTo 1
Case Is = 15
Winsock2.SendData Data
GoTo 1
Case Is = 51
Winsock2.SendData Data
GoTo 1
Case Is = 0
Winsock2.SendData Data
GoTo 1
End Select
Case Is = 203
Winsock2.SendData Data
GoTo 1
Case Is = 202
Winsock2.SendData Data
GoTo 1
End Select
Winsock2.SendData Data
1
End Sub

Private Sub ParseVoiceServer(Data As String)
On Error Resume Next
Dim TcpLen As Integer, HdrLen As Integer, CommandType As Integer, PcktType As Integer, TempIP As String
TcpLen = (256 * Asc(Mid(Data, 1, 1)) + Asc(Mid(Data, 2, 1)))
HdrLen = Asc(Mid(Data, 8, 1)) + 4
CommandType = Asc(Mid(Data, HdrLen + 2, 1))
Select Case CommandType
Case Is = 204
PcktType = Asc(Mid(Data, HdrLen + 14, 1))
Select Case PcktType
Case Is = 10
TempIP = Mid(Data, HdrLen + 23, 4)
MyRoomServer = Asc(Mid(TempIP, 1, 1)) & Chr(46) & Asc(Mid(TempIP, 2, 1)) & Chr(46) & Asc(Mid(TempIP, 3, 1)) & Chr(46) & Asc(Mid(TempIP, 4, 1))
Winsock1.SendData Replace(Data, TempIP, Chr(127) & String(2, 0) & Chr(1))
RaiseEvent CurrentIP(MyRoomServer)
DoEvents
Winsock2.Close
Winsock1.Close
Winsock2.RemoteHost = MyRoomServer
Winsock1.LocalPort = 5001
Winsock1.Listen
GoTo 1
Case Is = 8
Winsock1.SendData Data
GoTo 1
Case Is = 7
Winsock1.SendData Data
GoTo 1
Case Is = 2
TempIP = Mid(Data, HdrLen + 23, 4)
Winsock1.SendData Replace(Data, TempIP, Chr(127) & String(2, 0) & Chr(1))
GoTo 1
Case Is = 4
MySourceID = Mid(Data, HdrLen + 5, 4)
Form1.YTCP1.SrartService MyRoomServer, MySourceID
TempIP = Mid(Data, HdrLen + 23, 4)
Winsock1.SendData Replace(Data, TempIP, Chr(127) & String(2, 0) & Chr(1))
GoTo 1
Case Is = 50
Winsock1.SendData Data
GoTo 1
Case Is = 14
Winsock1.SendData Data
GoTo 1
Case Is = 13
Winsock1.SendData Data
GoTo 1
Case Is = 15
Winsock1.SendData Data
GoTo 1
Case Is = 0
Winsock1.SendData Data
GoTo 1
End Select
Case Is = 202
Winsock1.SendData Data
GoTo 1
Case Is = 203
Winsock1.SendData Data
GoTo 1
End Select
Winsock1.SendData Data
1
End Sub
