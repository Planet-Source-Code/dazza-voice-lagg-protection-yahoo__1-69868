VERSION 5.00
Begin VB.UserControl YTCP 
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   InvisibleAtRuntime=   -1  'True
   Picture         =   "YUDP.ctx":0000
   ScaleHeight     =   525
   ScaleWidth      =   495
   ToolboxBitmap   =   "YUDP.ctx":04A3
End
Attribute VB_Name = "YTCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event LoggedInVoice()
Public Event LoggedOutVoice()
Public Event ErrorOnVoice()
Private WithEvents Winsock1 As Winsock
Attribute Winsock1.VB_VarHelpID = -1
Private WithEvents Winsock2 As Winsock
Attribute Winsock2.VB_VarHelpID = -1
Private WithEvents Winsock3 As Winsock
Attribute Winsock3.VB_VarHelpID = -1
Private MySourceID As String, intAudio As Long, intPacket As Long, VCdom As Integer

Public Sub SrartService(UDPIP As String, SrcID As String)
On Error Resume Next
Winsock1.Close
Winsock2.Close
Winsock3.Close
MySourceID = SrcID
If UDPIP = "" Then GoTo 1
Winsock2.Protocol = sckTCPProtocol
Winsock1.Protocol = sckTCPProtocol
Winsock3.Protocol = sckUDPProtocol
Winsock1.LocalPort = "5000"
Winsock3.LocalPort = "5000"
Winsock3.Bind
Winsock1.Listen
Winsock2.RemoteHost = UDPIP
Winsock2.RemotePort = "5000"
Winsock2.Connect
1
End Sub

Public Sub StopService()
On Error Resume Next
Winsock1.Close
Winsock2.Close
Winsock3.Close
RaiseEvent LoggedOutVoice
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
Set Winsock1 = New Winsock
Set Winsock2 = New Winsock
Set Winsock3 = New Winsock
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
Set Winsock3 = Nothing
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
ParseClient TmpData
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
ParseServer TmpData
Else
Exit Sub
End If
DoEvents
Wend
End With
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
If Winsock1.RemoteHostIP = Chr(49) & Chr(50) & Chr(55) & Chr(46) & Chr(48) & Chr(46) & Chr(48) & Chr(46) & Chr(49) Then
Winsock1.Close
Winsock1.Accept requestID
Winsock3.Close
Else
Winsock3.Close
Winsock1.Close
End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Winsock1.Close
Winsock2.Close
RaiseEvent ErrorOnVoice
End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Winsock1.Close
Winsock2.Close
RaiseEvent ErrorOnVoice
End Sub

Private Sub Winsock3_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Winsock3.Close
End Sub

Private Sub Winsock1_Close()
On Error Resume Next
Winsock2.Close
Winsock1.Close
RaiseEvent LoggedOutVoice
End Sub

Private Sub Winsock2_Close()
On Error Resume Next
Winsock1.Close
Winsock2.Close
RaiseEvent LoggedOutVoice
End Sub

Private Sub Winsock3_Close()
On Error Resume Next
Winsock3.Close
End Sub

Private Sub ParseClient(Data As String)
On Error Resume Next
Dim VcDt As String, VcEnc As String, PackLen As Integer, PacketType As Integer
Data = Mid(Data, 5)
PackLen = Len(Data)
PacketType = Asc(Mid(Data, 2, 1))
Select Case PacketType
Case Is = 34
If PackLen = 12 Then
If Mid(Data, 9, 4) = String(4, 0) Then
Winsock2.SendData TCPHeader(Data)
Else
Call VoiceDomination(Data)
End If
Else
Call VoiceDomination(Data)
End If
Case Is = 162
If PackLen = 12 Then
Winsock2.SendData TCPHeader(Data)
Else
Call VoiceDomination(Data)
End If
End Select
End Sub

Private Sub ParseServer(Data As String)
On Error Resume Next
Dim PackLen As Integer, PacketType As Integer
Data = Mid(Data, 5)
PackLen = Len(Data)
If Mid(Data, 1, 1) = Chr(128) Then
PacketType = Asc(Mid(Data, 2, 1))
Select Case PacketType
Case Is = 34
If PackLen = 12 Then
Winsock1.SendData TCPHeader(ReplaceSequence(Data))
Else
ParseVoiceDataRecv Data
End If
Case Is = 162
If PackLen = 12 Then
If Mid(Data, 9, 4) = String(4, 0) Then
Winsock1.SendData TCPHeader(Data)
Else
Winsock1.SendData TCPHeader(Data)
VCdom = 1
intAudio = 0
intPacket = 0
RaiseEvent LoggedInVoice
End If
Else
ParseVoiceDataRecv Data
End If
Case Is = 127
ParseVoiceDataRecv Data
Case Else
If PackLen = 12 Then
Winsock1.SendData TCPHeader(ReplaceSequence(Data))
Else
ParseVoiceDataRecv Data
End If
1
End Select
End If
End Sub

Private Sub ParseVoiceDataRecv(Data As String)
On Error Resume Next
Dim PackLen As Integer
PackLen = Len(Data)
If PackLen <= 12 Then Exit Sub
Winsock1.SendData TCPHeader(ReplaceTimeStamp(Data))
End Sub

Private Function ReplaceTimeStamp(VoicePacket As String) As String
On Error Resume Next
Dim PacketStart As String, PacketEnd As String
PacketStart = Mid(VoicePacket, 1, 2)
PacketEnd = Mid(VoicePacket, 9)
intAudio = intAudio + 720
intPacket = intPacket + 1
If intAudio > 4294836225# Then intAudio = 720
If intPacket > 65535 Then intPacket = 1
ReplaceTimeStamp = PacketStart & Chr(Int(intPacket / 256)) & Chr(Int(intPacket Mod 256)) & Chr((((intAudio And &HFF000000) / 256) / 256) / 256) & Chr(((intAudio And &HFF0000) / 256) / 256) & Chr((intAudio And &HFF00&) / 256) & Chr(intAudio And &HFF&) & PacketEnd
End Function

Private Function ReplaceSequence(VoicePacket As String) As String
On Error Resume Next
Dim PacketStart As String, PacketEnd As String
PacketStart = Mid(VoicePacket, 1, 2)
PacketEnd = Mid(VoicePacket, 9)
intPacket = intPacket + 1
If intPacket > 65535 Then intPacket = 1
ReplaceSequence = PacketStart & Chr(Int(intPacket / 256)) & Chr(Int(intPacket Mod 256)) & String(4, 0) & PacketEnd
End Function

Private Function TCPHeader(VoicePacket As String) As String
On Error Resume Next
Dim PackLen As Integer
PackLen = Len(VoicePacket)
TCPHeader = Chr(0) & Chr(PackLen + 4) & String(2, 0) & VoicePacket
End Function

Private Sub VoiceDomination(VoicePacket As String)
On Error Resume Next
Dim PackLen As Integer
PackLen = Len(VoicePacket)
If PackLen <= 12 Then GoTo 1
If Mid(VoicePacket, 1, 2) = Chr(128) & Chr(162) Then VCdom = 2
If VCdom >= 96 Then
VCdom = 2
Winsock2.SendData TCPHeader(Chr(128) & Chr(34) & Chr(0) & Chr(1) & String(4, 0) & MySourceID)
Winsock2.SendData TCPHeader(Chr(128) & Chr(162) & Chr(0) & Chr(VCdom) & Mid(VoicePacket, 5))
Else
Winsock2.SendData TCPHeader(Chr(128) & Chr(34) & Chr(0) & Chr(VCdom) & Mid(VoicePacket, 5))
End If
3
VCdom = VCdom + 1
Exit Sub
2
Winsock2.SendData TCPHeader(VoicePacket)
Exit Sub
1
Winsock2.SendData TCPHeader(Chr(128) & Chr(34) & Chr(0) & Chr(1) & String(4, 0) & MySourceID)
VCdom = 2
End Sub
