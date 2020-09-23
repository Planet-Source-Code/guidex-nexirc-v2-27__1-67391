VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmIRCServer 
   Caption         =   "NexIRC - Server"
   ClientHeight    =   2430
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIRCServer.frx":0000
   LinkTopic       =   "frmIRCServer"
   MDIChild        =   -1  'True
   ScaleHeight     =   2430
   ScaleWidth      =   6240
   Begin nexIRC.ctlTBox txtIncoming 
      Height          =   1455
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2566
   End
   Begin VB.ListBox lstUsers 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      IntegralHeight  =   0   'False
      Left            =   4320
      TabIndex        =   4
      Top             =   0
      Width           =   1815
   End
   Begin VB.TextBox txtOutgoing 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   230
      Left            =   0
      TabIndex        =   3
      Top             =   1560
      Width           =   6135
   End
   Begin VB.Timer tmrLinkPing 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   65535
      Left            =   1020
      Top             =   1920
   End
   Begin MSWinsockLib.Winsock Link 
      Index           =   0
      Left            =   60
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "dennis"
      RemotePort      =   6669
      LocalPort       =   6668
   End
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   60000
      Left            =   1440
      Top             =   1920
   End
   Begin VB.Timer tmrKlined 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10000
      Left            =   2520
      Top             =   1920
   End
   Begin VB.Timer tmrKill 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   200
      Left            =   2880
      Top             =   1920
   End
   Begin VB.Timer tmrFloodProt 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   250
      Left            =   2160
      Top             =   1920
   End
   Begin VB.Timer tmrSend 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   200
      Left            =   1800
      Top             =   1920
   End
   Begin MSWinsockLib.Winsock wsock 
      Index           =   0
      Left            =   480
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6667
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   4200
      X2              =   4200
      Y1              =   1440
      Y2              =   0
   End
   Begin VB.Label lblServer 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Label lblServerSocket 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label lblClientSocket 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1635
   End
End
Attribute VB_Name = "frmIRCServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text
Option Base 1

Public Sub ActivateResize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Form_Resize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ActivateResize()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long, FS As New FileSystemObject
Me.Icon = mdiNexIRC.Icon
Call AddTaskPanel("Server", 1)
lSettings.sIRCServerVisible = True
ReDim Users(4)
lUserCount = lUserCount + 4
Rehash
Link(0).LocalPort = LinkPort
Link(0).Listen
Caption = wsock(0).LocalIP
For i = LBound(Users) To UBound(Users): Set Users(i) = Nothing: Next i
For i = LBound(Channels) To UBound(Channels): Set Channels(i) = Nothing: Next i
Started = Now
If LogLevel > 3 Then WriteHeader
MaxGlobalUsers = 0
CurGlobalUsers = 0
Caption = "IRC Server (" & wsock(0).LocalIP & ":" & wsock(0).LocalPort & ") "
DoColor txtIncoming, "1• Ready for incoming connections"
DoColor txtIncoming, "1• Commands: help (command list), rehash (restart server), options (edit options), msg (message server), muser (message user), users (user list), dcu (disconnect user), clear (clear status text)"
ActivateResize
txtIncoming.SetBackColor lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor
txtOutgoing.BackColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor
txtOutgoing.ForeColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sTextColor
lstUsers.BackColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor
lstUsers.ForeColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sTextColor
If lSettings.sBorderlessObjects = True Then
    txtOutgoing.BorderStyle = 0
    txtIncoming.SetBorderStyle True
Else
    txtOutgoing.BorderStyle = 1
    txtIncoming.SetBorderStyle False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtOutgoing.Width = Me.ScaleWidth
txtOutgoing.Top = Me.ScaleHeight - txtOutgoing.Height
If txtIncoming.Left <> 0 Then txtIncoming.Left = 0
txtIncoming.Height = Me.ScaleHeight - txtOutgoing.Height
If txtOutgoing.Left <> 0 Then txtOutgoing.Left = 0
txtIncoming.Width = Me.ScaleWidth - lstUsers.Width
lstUsers.Left = txtOutgoing.Width - lstUsers.Width
If lstUsers.Top <> 0 Then lstUsers.Top = 0
lstUsers.Height = Me.ScaleHeight - txtOutgoing.Height
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Resize()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If LogLevel <> 0 Then WriteFooter
lSettings.sIRCServerVisible = False
RemoveTaskbar "Server"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub Link_Close(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
SendSvrMsg ServerName & " -- link closed -- " & Link(Index).Tag, False
SendLinks "DeadLink" & vbLf & ServerName & vbLf & Link(Index).Tag
If LogLevel = 1 Or LogLevel = 3 Then
    If LogFormat = 0 Then
        LogText "[LINK]<" & Now & " (LINK CLOSED " & ServerName & " -- " & Link(Index).Tag & ")> "
    Else
        LogHTML ServerName, "LINK CLOSED " & ServerName & " -- " & Link(Index).Tag
    End If
End If
Dim i As Long
For i = 1 To UBound(Users)
    DoEvents
    If Not Users(i) Is Nothing Then
        If Users(i).IsOnLink(Link(Index).Tag) Then
            SendQuit i, ServerName & " -- " & Link(Index).Tag
            Set Users(i) = Nothing
        End If
    End If
Next i
CurLinkCount = CurLinkCount - 1
Unload Link(Index)
Unload tmrLinkPing(Index)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Link_Close(Index As Integer)"
End Sub

Private Sub Link_Connect(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Wait 100
Dim i As Long, X As Long
For i = 1 To UBound(Users)
    If Not Users(i) Is Nothing Then
        SendLinks "NewUser" & vbLf & Users(i).Nick & vbLf & Users(i).Name & vbLf & Users(i).DNS & vbLf & Users(i).ident & vbLf & Users(i).Server & vbLf & Users(i).ServerDescritption & vbLf & Users(i).SignOn & vbLf & Users(i).GID & vbLf & Users(i).GetModes & vbLf & ServerName & " ", , Index
        Wait 10
        For X = 1 To Users(i).Onchannels.Count
            SendLinks "JoinChan" & vbLf & Users(i).Nick & vbLf & Users(i).Onchannels(X), , Index
        Next X
    End If
Next i
For i = 1 To UBound(Channels)
    If Not Channels(i) Is Nothing Then
        SendLinks "ChanMode" & vbLf & "ChanServ" & vbLf & "+" & vbLf & Channels(i).GetModesForFile & vbLf & Channels(i).Name, , Index
        If Channels(i).Key <> "" Then SendLinks "Key" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & Channels(i).Key, , Index
        If Channels(i).Limit <> 0 Then SendLinks "Limit" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & Channels(i).Limit, , Index
        Dim Y As Long
        SendLinks "SetTopic" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & "" & vbLf & Channels(i).Topic, , Index
        For Y = 1 To Channels(i).Bans.Count
            SendLinks "BanUser" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & "" & vbLf & Channels(i).Bans(Y), , Index
        Next Y
        For Y = 1 To Channels(i).Invites.Count
            SendLinks "InviteUser" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & "" & vbLf & Channels(i).Invites(Y), , Index
        Next Y
        For Y = 1 To Channels(i).Exceptions.Count
            SendLinks "ExceptUser" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & "" & vbLf & Channels(i).Exceptions(Y), , Index
        Next Y
    End If
Next i
SendLinks "Info" & vbLf & ServerName
On Local Error GoTo 0
Load tmrLinkPing(Index)
tmrLinkPing(Index).Enabled = True
tmrLinkPing(Index).Tag = 1
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Link_Connect(Index As Integer)"
End Sub

Private Sub Link_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim LinkCount As Long
LinkCount = Link.Count + 1
CurLinkCount = CurLinkCount + 1
MaxLinkCount = MaxLinkCount + 1
Index = LinkCount
Load Link(LinkCount)
Link(LinkCount).Close
Link(LinkCount).LocalPort = 30000 + LinkCount
Link(LinkCount).Accept requestID
Wait 100
Dim i As Long, X As Long
For i = 1 To UBound(Users)
    If Not Users(i) Is Nothing Then
        SendLinks "NewUser" & vbLf & Users(i).Nick & vbLf & Users(i).Name & vbLf & Users(i).DNS & vbLf & Users(i).ident & vbLf & Users(i).Server & vbLf & Users(i).ServerDescritption & vbLf & Users(i).SignOn & vbLf & Users(i).GID & vbLf & Users(i).GetModes & vbLf & ServerName & " ", , Index
        For X = 1 To Users(i).Onchannels.Count
            SendLinks "JoinChan" & vbLf & Users(i).Nick & vbLf & Users(i).Onchannels(X), , Index
        Next X
    End If
Next i
For i = 1 To UBound(Channels)
    If Not Channels(i) Is Nothing Then
        SendLinks "ChanMode" & vbLf & "ChanServ" & vbLf & "+" & vbLf & Channels(i).GetModesForFile & vbLf & Channels(i).Name, , Index
        If Channels(i).Key <> "" Then SendLinks "Key" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & Channels(i).Key, , Index
        If Channels(i).Limit <> 0 Then SendLinks "Limit" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & Channels(i).Limit, , Index
        Dim Y As Long
        SendLinks "SetTopic" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & "" & vbLf & Channels(i).Topic, , Index
        For Y = 1 To Channels(i).Bans.Count
            SendLinks "BanUser" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & "" & vbLf & Channels(i).Bans(Y), , Index
        Next Y
        For Y = 1 To Channels(i).Invites.Count
            SendLinks "InviteUser" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & "" & vbLf & Channels(i).Invites(Y), , Index
        Next Y
        For Y = 1 To Channels(i).Exceptions.Count
            SendLinks "ExceptUser" & vbLf & "ChanServ" & vbLf & Channels(i).Name & vbLf & "" & vbLf & Channels(i).Exceptions(Y), , Index
        Next Y
    End If
Next i
SendLinks "Info" & vbLf & ServerName
Load tmrLinkPing(Index)
tmrLinkPing(Index).Enabled = True
tmrLinkPing(Index).Tag = 1
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Link_ConnectionRequest(Index As Integer, ByVal requestID As Long)"
End Sub

Private Sub Link_DataArrival(Index As Integer, ByVal bytesTotal As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim strData As String, strcmd() As String, cmdArray() As String, i As Long, User As clsIRCServer_User, DontSendLink As Boolean, X As Long, NewUser As clsIRCServer_User, strRoute() As String
Link(Index).GetData strData, 8
If LogLevel = 1 Then
    If LogFormat = 0 Then
        LogText "[LINK]<" & Now & " (INCOMING " & Link(Index).Tag & ")> " & strData
    Else
        LogHTML Link(Index).Tag & "(Link) INCOMING", strData
    End If
End If
cmdArray = Split(strData, vbCrLf)
For X = LBound(cmdArray) To UBound(cmdArray)
    If cmdArray(X) = "" Then GoTo NextCmd
    strcmd = Split(cmdArray(X), vbLf)
    Select Case strcmd(0)
        Case "Info"
            Link(Index).Tag = strcmd(1)
            SendSvrMsg ServerName & " -- linked -- " & Link(Index).Tag, False
            If LogLevel = 1 Or LogLevel = 3 Then
                If LogFormat = 0 Then
                    LogText "[LINK]<" & Now & " (LINKED " & ServerName & " -- " & strcmd(1) & ")> " & strData
                Else
                    LogHTML ServerName, "LINKED " & ServerName & " -- " & strcmd(1)
                End If
            End If
        Case "NewUser"
            If strcmd(8) = "" Then GoTo NextCmd
            Set NewUser = NickToObject(strcmd(1))
            If NewUser Is Nothing Then
                Set User = GetFreeSlot
                lUserCount = lUserCount - 1
                User.DNS = strcmd(3)
                User.Nick = strcmd(1)
                User.Name = strcmd(2)
                User.ident = strcmd(4)
                User.Server = strcmd(5)
                User.ServerDescritption = strcmd(6)
                User.SignOn = strcmd(7)
                User.GID = strcmd(8)
                If InStr(1, strcmd(10), " ") = 0 Then strcmd(10) = strcmd(10) & " "
                User.Route = strcmd(10)
                User.Hops = CountSpaces(strcmd(10)) - 1
                User.AddModes strcmd(9)
            Else
                If strcmd(8) = NewUser.GID Then GoTo NextCmd
                If NewUser.SignOn < strcmd(7) Then
                    SendLinks "KillUser" & vbLf & strcmd(1) & vbLf & "Nick Collision, other nick signed on earlier"
                Else
                    SendWsock Index, ":" & Users(NewUser.Index).Nick & "!" & Users(NewUser.Index).ident & "@" & Users(NewUser.Index).DNS & " KILL :" & "Nick Collision, other nick signed on earlier"
                    SendWsock Index, "ERROR :Closing Link: " & "Nick Collision, other nick signed on earlier" & vbCrLf
                    Set User = GetFreeSlot
                    lUserCount = lUserCount - 1
                    User.DNS = strcmd(3)
                    User.Nick = strcmd(1)
                    User.Name = strcmd(2)
                    User.ident = strcmd(4)
                    User.Server = strcmd(5)
                    User.ServerDescritption = strcmd(6)
                    User.SignOn = strcmd(7)
                    User.GID = strcmd(8)
                    If InStr(1, strcmd(10), " ") = 0 Then strcmd(10) = strcmd(10) & " "
                    User.Route = strcmd(10)
                    User.Hops = CountSpaces(strcmd(10)) - 1
                    User.AddModes strcmd(9)
                End If
            End If
            SendLinks cmdArray(X) & " " & ServerName & " ", CLng(Index)
            DontSendLink = False
        Case "QuitUser"
            Set User = NickToObject(strcmd(1))
            SendQuit User.Index, strcmd(2), , False
            Set Users(User.Index) = Nothing
        Case "KillUser"
            Set User = NickToObject(strcmd(1))
            SendQuit User.Index, strcmd(2), True, False
            Set Users(User.Index) = Nothing
        Case "JoinChan"
            Set User = NickToObject(strcmd(1))
            If Not ChanExists(strcmd(2)) Then
                Dim NewChannel As clsIRCServer_Channel
                Set NewChannel = GetFreeChan
                NewChannel.Name = strcmd(2)
                NewChannel.Modes.Add "t", "t"
                NewChannel.Modes.Add "n", "n"
                NewChannel.Topic = DefTopic
                NotifyJoin User.Index, strcmd(2), False
                NewChannel.NormUsers.Add User.Nick, User.Nick
                NewChannel.All.Add User.Nick, User.Nick
                Users(Index).Onchannels.Add strcmd(2), strcmd(2)
            Else
                NotifyJoin User.Index, strcmd(2), False
                ChanToObject(strcmd(2)).All.Add User.Nick, User.Nick
                ChanToObject(strcmd(2)).NormUsers.Add User.Nick, User.Nick
            End If
            User.Onchannels.Add strcmd(2), strcmd(2)
        Case "PartUser"
            Set User = NickToObject(strcmd(1))
            SendPart User.Index, strcmd(2), strcmd(3), False
        Case "ModeUser"
            Set User = NickToObject(strcmd(1))
            Select Case strcmd(2)
                Case "+"
                    AddUserMode User.Index, strcmd(3), , False
                Case "-"
                    RemoveUsermode User.Index, strcmd(3), , False
            End Select
        Case "ChanMode"
            Set User = NickToObject(strcmd(1))
            Select Case strcmd(2)
                Case "+"
                    AddChanModes strcmd(3), strcmd(4), User, False
                Case "-"
                    RemoveChanModes strcmd(3), strcmd(4), User, False
            End Select
        Case "KickUser"
            Set User = NickToObject(strcmd(1))
            If strcmd(3) = "" Then strcmd(3) = strcmd(1)
            KickUser User.Nick, strcmd(2), strcmd(4), strcmd(3), True, False
        Case "KLine"
            Klines.Add strcmd(1), strcmd(1)
        Case "ServerMsg"
            SendSvrMsg strcmd(1), , Link(Index).Tag
        Case "Global"
            For i = LBound(Users) To UBound(Users)
                If Not Users(i) Is Nothing Then SendNotice "", "• Global -- " & strcmd(1), ServerName, , CInt(i), False
            Next i
        Case "PrivMsgChan"
            SendMsg strcmd(2), strcmd(3), strcmd(1), True, False
        Case "PrivMsgUser"
            SendMsg strcmd(2), strcmd(3), strcmd(1), False, False
        Case "NoticeUser"
            SendNotice strcmd(2), strcmd(3), strcmd(1), , , False
        Case "NoticeChan"
            SendNotice strcmd(2), strcmd(3), strcmd(1), True, , False
        Case "Nick"
            Set User = NickToObject(strcmd(1))
            ChangeNick User.Index, strcmd(2), False
        Case "OpUser"
            OpUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), True, False
        Case "DeOpUser"
            DeOpUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "VoiceUser"
            VoiceUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), True, False
        Case "DeVoiceUser"
            DeVoiceUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "BanUser"
            BanUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "UnBanUser"
            UnBanUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "ExceptUser"
            ExceptionUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "UnExceptUser"
            UnExceptionUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "InviteUser"
            InviteUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "UnInviteUser"
            UnInviteUser ChanToObject(strcmd(2)), strcmd(4), strcmd(1), False
        Case "Limit"
            AddChanModes "l " & strcmd(3), strcmd(2), NickToObject(strcmd(1)), False
        Case "Key"
            AddChanModes "k " & strcmd(3), strcmd(2), NickToObject(strcmd(1)), False
        Case "AddInvite"
            ChanToObject(strcmd(2)).Invited.Add strcmd(4), strcmd(4)
            If NickToObject(strcmd(4)).LocalUser Then SendWsock NickToObject(strcmd(4)).Index, ":" & strcmd(1) & " INVITE " & strcmd(4) & " " & strcmd(2)
        Case "SetTopic"
            SetTopic strcmd(2), strcmd(4), strcmd(1), False
        Case "DeadLink"
            SendSvrMsg strcmd(1) & " -- Link Closed -- " & strcmd(2)
            For i = 5 To UBound(Users)
                If Users(i).IsOnLink(strcmd(2)) Then
                    SendQuit i, strcmd(1) & " -- " & strcmd(2)
                    Set Users(i) = Nothing
                End If
            Next i
            SendLinks cmdArray(X), CLng(Index)
            DontSendLink = False
        Case "PING"
            SendLinks "PONG" & vbLf, , Index
            DontSendLink = True
        Case "PONG"
            tmrLinkPing(Index).Tag = 1
            DontSendLink = True
        Case Else
            DontSendLink = True
    End Select
    If Not DontSendLink Then SendLinks cmdArray(X), CLng(Index)
    DontSendLink = False
NextCmd:
Next X
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Link_DataArrival(Index As Integer, ByVal bytesTotal As Long)"
End Sub

Private Sub Link_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
SendSvrMsg ServerName & " -- Link Closed -- " & Link(Index).Tag, True
If LogLevel = 1 Or LogLevel = 3 Then
    If LogFormat = 0 Then
        LogText "[LINK]<" & Now & " (LINK CLOSED " & ServerName & " -- " & Link(Index).Tag & ")> "
    Else
        LogHTML ServerName, "LINK CLOSED " & ServerName & " -- " & Link(Index).Tag
    End If
End If
Link_Close (Index)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Link_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)"
End Sub

Private Sub tmrFloodProt_Timer(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Users(Index).HasRegistered = False Then
    Users(Index).HasRegistered = True
    Users(Index).MsgsSent = 0
    Exit Sub
End If
If Users(Index).IRCOp Then Exit Sub
If Users(Index).MsgsSent > 3000 Then
    SendQuit CLng(Index), "killed by sysadmin (excess flooding)", True
    SendWsock Users(Index).Index, ":Server!NexIRC@" & ServerName & " KILL " & Users(Index).Nick & " :Excess flooding", True
    SendWsock Users(Index).Index, "ERROR :Closing Link: " & Users(Index).Nick & "[" & frmIRCServer.wsock(Index).RemoteHostIP & ".] " & ServerName & " (excess flooding)", True
    If LogLevel = 1 Or LogLevel = 3 Then
        If LogFormat = 0 Then
            LogText "[LINK]<" & Now & " (FLOOD PROTECTION " & Users(Index).Nick & ")> "
        Else
            LogHTML ServerName, "FLOOD PROTECTION " & Users(Index).Nick
        End If
    End If
    Users(Index).Killed = True
    SendSvrMsg "Recieved Kill message for " & Users(Index).Nick & "!" & Users(Index).ident & "@" & Users(Index).DNS & " Path: " & Users(Index).Nick & " (excess flooding)", True
End If
Users(Index).FloodProt = Users(Index).FloodProt + 1
If Users(Index).FloodProt = 5 Then
    If GetPercent1(3000, Users(Index).MsgsSent) >= 80 Then SendSvrMsg "Flooding Alert for user " & Users(Index).Nick & "! (" & Users(Index).MsgsSent & "/3000 (" & GetPercent1(3000, Users(Index).MsgsSent) & "%)", True
    Users(Index).MsgsSent = 0
    Users(Index).FloodProt = 1
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrFloodProt_Timer(Index As Integer)"
End Sub

Private Sub tmrKill_Timer(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Index = 0 Then
    wsock_Close tmrKill(Index).Tag
    wsock(0).Listen
    tmrKill(0).Enabled = False
    tmrKill(0).interval = 200
Else
    wsock_Close tmrKill(Index).Tag
    Unload tmrKill(Index)
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrKill_Timer(Index As Integer)"
End Sub

Private Sub tmrKlined_Timer(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Klines.Remove tmrKlined(Index).Tag
Unload tmrKlined(Index)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrKlined_Timer(Index As Integer)"
End Sub

Private Sub tmrLinkPing_Timer(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Not CLng(tmrLinkPing(Index).Tag) = 1 Then
    SendQuit CLng(Index), "Ping Timeout"
    SendSvrMsg "No response from " & Link(Index).Tag & ", Closing Link", True, ServerName
    Link_Close (Index)
    Exit Sub
End If
tmrLinkPing(Index).Tag = 0
Link(Index).SendData "PING" & vbLf
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrLinkPing_Timer(Index As Integer)"
End Sub

Private Sub tmrSend_Timer(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Not wsock(Index).Tag = "" Then
    wsock(Index).Tag = Replace(wsock(Index).Tag, vbCrLf, "")
    Dim i As Long
    For i = 1 To Len(wsock(Index).Tag) Step MaxChunkSize
        wsock(Index).SendData Mid(wsock(Index).Tag, i, MaxChunkSize)
    Next i
    wsock(Index).SendData vbCrLf
    wsock(Index).Tag = ""
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrSend_Timer(Index As Integer)"
End Sub

Private Sub tmrTimeOut_Timer(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Not Users(Index).Ponged Then
    SendQuit CLng(Index), "Ping Timeout"
    wsock_Close (Index)
    Exit Sub
End If
Users(Index).Ponged = False
SendPing CLng(Index)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrTimeOut_Timer(Index As Integer)"
End Sub

Private Sub txtOutgoing_KeyPress(KeyAscii As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If KeyAscii = 13 Then
    Dim msg2 As String, msg As String, F As Integer, i As Integer, ub As Long
    msg = Trim(LCase(txtOutgoing.Text))
    KeyAscii = 0
    txtOutgoing.Text = ""
    Select Case msg
    Case "users"
        ub = UBound(Users)
        For i = 1 To ub
            If Not Users(i) Is Nothing Then
                DoColor txtIncoming, "" & Color.Normal & "• " & Users(i).Nick & " - " & Users(i).DNS & vbCrLf
                F = F + 1
            End If
        Next i
        DoColor txtIncoming, "" & Color.Normal & "• " & Str(F) & " Users total" & vbCrLf
    Case "muser"
        Dim lmsg As String
        msg2 = InputBox("Enter username:", "Message User")
        If Len(msg2) <> 0 Then
            ub = UBound(Users)
            F = 0
            For i = 1 To ub
                If Not Users(i) Is Nothing Then
                    If LCase(Users(i).Nick) = LCase(msg2) Then
                        lmsg = InputBox("Enter Message: ", "Message User")
                        If Len(lmsg) <> 0 Then
                            wsock(i).SendData "NOTICE " & Users(i).Nick & ":" & lmsg
                        End If
                        F = 1
                        Exit For
                    End If
                End If
            Next i
            If F = 0 Then
                DoColor txtIncoming, "" & Color.Normal & "• User not found"
            Else
                DoColor txtIncoming, "" & Color.Normal & "• Dcu success"
            End If
        End If
    Case "help"
        DoColor txtIncoming, "" & Color.Normal & "• Commands: help (command list), rehash (restart server), options (edit options), msg (message server), muser (message user), users (user list), dcu (disconnect user), clear (clear status text)"
        DoColor txtIncoming, "" & Color.Normal & "• End of command list" & vbCrLf
    Case "options"
        OpenTextFile App.Path & "\data\config\server\server.ini"
    Case "rehash"
        Restart
        DoColor txtIncoming, "" & Color.Normal & "• Server Restarted" & vbCrLf
    Case "msg"
        ub = UBound(Users)
        msg2 = InputBox("Enter global server message:", "Message Server")
        For i = 1 To ub
            If Not Users(i) Is Nothing Then
                frmIRCServer.wsock(i).SendData "NOTICE " & Users(i).Nick & ":" & msg2 & vbCrLf
                DoColor txtIncoming, "" & Color.Normal & "• " & Users(i).Nick & " - " & Users(i).DNS & vbCrLf
                F = F + 1
            End If
        Next i
        DoColor txtIncoming, "" & Color.Normal & "• Message sent to " & Str(F) & " Users total" & vbCrLf
    Case "dcu"
        msg2 = InputBox("Enter username:", "Disconnect user")
        If Len(msg2) <> 0 Then
            ub = UBound(Users)
            F = 0
            For i = 1 To ub
                If Not Users(i) Is Nothing Then
                    If LCase(Users(i).Nick) = LCase(msg2) Then
                        DoColor txtIncoming, "" & Color.Normal & "• " & msg2 & " was disconnected " & vbCrLf
                        wsock(i).Close
                        F = 1
                    End If
                End If
            Next i
            If F = 0 Then
                DoColor txtIncoming, "" & Color.Normal & "• User not found"
            Else
                DoColor txtIncoming, "" & Color.Normal & "• Dcu success"
            End If
        End If
    Case Else
        DoColor txtIncoming, "" & Color.Normal & "• " & msg & vbCrLf
    End Select
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtOutgoing_KeyPress(KeyAscii As Integer)"
End Sub

Private Sub txtIncoming_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtOutgoing.SetFocus
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtIncoming_GotFocus()"
End Sub

Public Sub wsock_Close(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Not Users(Index).SentQuit Then SendQuit CLng(Index), "Client exited"
Dim i As Long, CurChan As clsIRCServer_Channel
For i = 1 To Users(Index).Onchannels.Count
    Set CurChan = ChanToObject(Users(Index).Onchannels(i))
    CurChan.All.Remove Users(Index).Nick
    If CurChan.IsNorm(Users(Index).Nick) Then
        CurChan.NormUsers.Remove Users(Index).Nick
    ElseIf CurChan.IsVoice(Users(Index).Nick) Then
        CurChan.Voices.Remove Users(Index).Nick
    ElseIf CurChan.IsOp(Users(Index).Nick) Then
        CurChan.Ops.Remove Users(Index).Nick
    End If
Next i
Set Users(Index) = Nothing
For i = 1 To CloneControl.Count
    If CloneControl(i) = wsock(Index).RemoteHostIP Then: CloneControl.Remove (i): Exit For
Next i
wsock(Index).Close
Unload wsock(Index)
Unload tmrTimeOut(Index)
Unload tmrFloodProt(Index)
Unload tmrSend(Index)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub wsock_Close(Index As Integer)"
End Sub

Private Sub wsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim FS As clsIRCServer_User, WelcomeStr As String
If ClientsFromIP(wsock(0).RemoteHostIP) >= SessionLimit Then
    SendSvrMsg "Session limit exceeded: " & wsock(0).RemoteHostIP & " [" & ServerName & "]", True
    wsock(0).Close
    wsock(0).Accept requestID
    wsock(0).SendData ":Server!Server@" & ServerName & " KILL You :You have exceeded your session limit" & vbCrLf
    wsock(0).SendData "ERROR :Closing Link: Session limit exceeded" & vbCrLf
    Wait 100
    wsock(0).Close
    wsock(0).Listen
    Exit Sub
End If
Dim Killine As String
Killine = IsKlined(AddressToName(wsock(0).RemoteHostIP))
If Killine <> "" Then
    SendSvrMsg "K-Line active for: " & wsock(0).RemoteHostIP & " (" & Killine & ") [" & ServerName & "]", True
    wsock(0).Close
    wsock(0).Accept requestID
    SendWsock 0, ":" & ServerName & " NOTICE AUTH :•  " & ServerName & " -- Your Site (IP, Country, ISP..etc...) has been banned from this Server", , True
    SendWsock 0, ":" & ServerName & " NOTICE AUTH :•  " & ServerName & " -- This is not necessarily your fault, if you think you have been banned without any reason please send an email to the admin: " & AdminEmail, , True
    SendWsock 0, ":Server!Server@" & ServerName & " KILL You :Your Site (IP, Country, ISP..etc...) has been banned from this Server", , True
    SendWsock 0, "ERROR :Closing Link: Your Site (IP, Country, ISP..etc...) has been banned from this Server", , True
    Wait 150
    wsock(0).Close
    wsock(0).Listen
    Exit Sub
End If
Set FS = GetFreeSlot
If FS Is Nothing Then
    SendSvrMsg "Server is Full: " & wsock(0).RemoteHostIP & " [" & ServerName & "]", True
    wsock(0).Close
    wsock(0).Accept requestID
    wsock(0).SendData ":Server!Server@" & ServerName & " KILL You :Server is full, try again later" & vbCrLf
    wsock(0).SendData "ERROR :Closing Link: Server is Full" & vbCrLf
    Wait 100
    wsock(0).Close
    wsock(0).Listen
    Exit Sub
End If
If ((lUserCount - 4) > lMaxUser) Then lMaxUser = lMaxUser + 1
Load wsock(FS.Index)
wsock(FS.Index).Accept requestID
wsock(0).Close: wsock(0).Listen
Load tmrTimeOut(FS.Index)
Load tmrFloodProt(FS.Index)
Load tmrSend(FS.Index)
tmrTimeOut(FS.Index).Enabled = True
tmrSend(FS.Index).Enabled = True
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim strDNS As String
strDNS = mdlDNS.AddressToName(wsock(FS.Index).RemoteHostIP)
FS.DNS = IIf(strDNS = "", wsock(FS.Index).RemoteHostIP, strDNS)
FS.Server = ServerName
FS.NewUser = True
FS.SignOn = UnixTime
FS.Idle = UnixTime
FS.ServerDescritption = ServerDesc
FS.LocalUser = True
FS.GID = CreateGUID
WelcomeStr = ":" & ServerName & " NOTICE AUTH :• Welcome to " & ServerName & "!" & vbCrLf & ":" & ServerName & " NOTICE AUTH :• Looking up your Hostname...." & vbCrLf
SendWsock FS.Index, WelcomeStr
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub wsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)"
End Sub

Private Sub wsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Users(Index).Killed Then Exit Sub
Dim strMsg As String, strcmd() As String, lb As Long, ub As Long, i As Long
wsock(Index).GetData strMsg
strMsg = Replace(strMsg, vbCrLf, vbLf)
Debug.Print strMsg
If Index <= 0 Then Exit Sub
Users(Index).MsgsSent = Users(Index).MsgsSent + (bytesTotal * 1.5)
If LogLevel = 1 Or LogLevel = 2 Then
    If LogFormat = 0 Then
        LogText "[Client]<" & Now & " (from " & Users(Index).Nick & ")> " & strMsg
    Else
        LogHTML Users(Index).Nick, strMsg
    End If
End If
ServerTraffic = ServerTraffic + bytesTotal
strcmd = Split(strMsg, vbLf)
lb = LBound(strcmd)
ub = UBound(strcmd)
For i = lb To ub
    If Users(Index) Is Nothing Then Exit Sub
    If Users(Index).Killed Or strcmd(i) = "" Then Exit Sub
    If strcmd(i) Like "NICK*" Then
        Dim NewNick As String
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 150
        If InStr(1, strcmd(i), ":") <> 0 Then
            NewNick = Replace(strcmd(i), "NICK :", "")
        Else
            NewNick = Replace(strcmd(i), "NICK ", "")
        End If
        If InStr(1, NewNick, " ") <> 0 Then NewNick = Mid(NewNick, 1, InStr(1, NewNick, " ") - 1)
        If NewNick = Users(Index).Nick Then GoTo NextCmd
        If Len(NewNick) > Nicklen Then NewNick = Mid(NewNick, 1, Nicklen)
        If Not IsValidString(NewNick) Then SendWsock Index, ":" & ServerName & " 432 * " & NewNick & " :Erroneus nickname, Nickname has been cut"
        If Len(NewNick) > Nicklen Then NewNick = Left(NewNick, Nicklen)
        If Not (ChangeNick(CLng(Index), NewNick, (Not Users(Index).NewUser))) Then
            SendWsock Index, ":" & ServerName & " 433 * " & NewNick & " :Nickname is already in use"
            GoTo NextCmd
        Else
            SendWsock Index, "PING :" & GetRand
            Users(Index).Identified = False
            Users(Index).NR = False
            Users(Index).ClearOwnerShip
        End If
        lstUsers.AddItem NewNick
        DoColor txtIncoming, "" & Color.Normal & "• " & NewNick & "(" & Users(Index).DNS & ") logged in"
    ElseIf strcmd(i) Like "USERHOST*" Then
        Dim User As clsIRCServer_User
        Set User = NickToObject(Replace(strcmd(i), "USERHOST ", ""))
        If User Is Nothing Then
            SendWsock Index, ":" & ServerName & " 302 " & Users(Index).Nick & " :"
        Else
            SendWsock Index, ":" & ServerName & " 302 " & Users(Index).Nick & " :" & Replace(strcmd(i), "USERHOST ", "") & "=+" & User.ID
        End If
    ElseIf strcmd(i) Like "USER*" Then
        Dim ident As String, Email As String, Name As String, NewIdent As String * 10
        ident = Replace(strcmd(i), "USER ", "")
        ident = Mid(ident, 1, InStr(1, ident, " ") - 1)
        Email = Replace(strcmd(i), "USER " & ident, "")
        Email = Mid(Email, 3)
        Email = Mid(Email, 1, InStr(1, Email, " "))
        Email = Replace(Email, Chr(34), "")
        Email = Mid(Email, 1, Len(Email) - 1)
        Email = ident & "@" & Email
        Name = Mid(strcmd(i), InStr(1, strcmd(i), ":") + 1)
        Users(Index).ident = Mid(ident, 1, 10)
        ident = Mid(ident, 1, 10) & "@" & wsock(Index).RemoteHostIP
        Users(Index).Email = Email
        Users(Index).ID = ident
        Users(Index).Name = Name
    ElseIf strcmd(i) Like "QUIT*" Then
        Dim Quit As String
        Quit = Mid(strcmd(i), InStr(1, strcmd(i), " :") + 2)
        SendQuit CLng(Index), Quit
        wsock_Close (Index)
    ElseIf strcmd(i) Like "JOIN*" Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 75
        Dim Chan As String, ck As String, X As Long
        If Replace(strcmd(i), "JOIN ", "") = "0" Or Replace(strcmd(i), "JOIN ", "") = "#0" Then
            For X = 1 To Users(Index).Onchannels.Count
                SendPart CLng(Index), Users(Index).Onchannels.Item(1), ""
            Next X
            GoTo NextCmd
        End If
        If CountSpaces(strcmd(i)) = 3 Then
            Chan = Replace(strcmd(i), "JOIN ", "")
            Chan = Mid(Chan, 1, InStr(1, Chan, " ") - 1)
            ck = Replace(strcmd(i), "JOIN " & Chan & " ", "")
        Else
            Chan = Replace(strcmd(i), "JOIN ", "")
        End If
        Dim Chans() As String
        Chan = Replace(Chan, " ", "")
        Chans = Split(Chan, ",")
        For X = 0 To UBound(Chans)
            Chan = Chans(X)
            IsValidString Mid(Chan, 2)
            If Users(Index).Onchannels.Count >= MaxJoinChannels Then
                SendWsock Index, ":" & ServerName & " 432 * " & NewNick & ":You have joined too many Channels"
                GoTo NextCmd
            End If
            If Not Users(Index).IsOnChan(Chan) Then
                If Not ChanExists(Chan) Then
                    Dim NewChannel As clsIRCServer_Channel
                    Set NewChannel = GetFreeChan
                    NewChannel.Name = Chan
                    NewChannel.Modes.Add "t", "t"
                    NewChannel.Modes.Add "n", "n"
                    NewChannel.Topic = DefTopic
                    NewChannel.Ops.Add Users(Index).Nick, Users(Index).Nick
                    NewChannel.All.Add Users(Index).Nick, Users(Index).Nick
                    Users(Index).Onchannels.Add Chan, Chan
                    SendWsock Index, ":" & Users(Index).Nick & " JOIN " & Chan, True
                    SendWsock Index, ":" & ServerName & " 353 " & Users(Index).Nick & " = " & Chan & " :" & Replace(NewChannel.GetOps & " " & NewChannel.GetVoices & " " & NewChannel.GetNorms, "  ", " "), True
                    SendWsock Index, ":" & ServerName & " 366 " & Users(Index).Nick & " " & Chan & " :End of /NAMES list.", True
                    NotifyJoin CLng(Index), Chan, False
                    SendLinks "JoinChan" & vbLf & Users(Index).Nick & vbLf & Chan
                Else
                    If lSettings.sHandleErrors = True Then On Local Error Resume Next
                    Dim JoinChan As clsIRCServer_Channel
                    Set JoinChan = ChanToObject(Chan)
                    If (Not JoinChan.Key = "") Then
                        If Not JoinChan.Key = ck And Not Users(Index).IRCOp And (Not Users(Index).IsOwner(JoinChan.Name)) Then
                            SendWsock Index, ":" & ServerName & " 475 " & Users(Index).Nick & " " & Chan & " :Cannot join channel (+b)"
                            GoTo NextCmd
                        End If
                    End If
                    If (JoinChan.All.Count >= JoinChan.Limit And JoinChan.Limit <> 0) And Not Users(Index).IRCOp And (Not Users(Index).IsOwner(JoinChan.Name)) Then
                        SendWsock Index, ":" & ServerName & " 471 " & Users(Index).Nick & " " & Chan & " :Cannot join channel (+l)"
                        GoTo NextCmd
                    End If
                    If JoinChan.IsBanned(Users(Index)) And (Users(Index).IRCOp = False) And (JoinChan.IsException(Users(Index)) = False) And (Not Users(Index).IsOwner(JoinChan.Name)) Then
                        SendWsock Index, ":" & ServerName & " 474 " & Users(Index).Nick & " " & Chan & " :Cannot join channel (+b)"
                        GoTo NextCmd
                    End If
                    If JoinChan.IsMode("i") And (Users(Index).IRCOp = False) And (JoinChan.IsInvited2(Users(Index)) = False) And (JoinChan.IsInvited(Users(Index).Nick) = False) And (Not Users(Index).IsOwner(JoinChan.Name)) Then
                        SendWsock Index, ":" & ServerName & " 473 " & Users(Index).Nick & " " & Chan & " :Cannot join channel (+i)"
                        GoTo NextCmd
                    End If
                    NotifyJoin CLng(Index), Chan
                    JoinChan.NormUsers.Add Users(Index).Nick, Users(Index).Nick
                    JoinChan.All.Add Users(Index).Nick, Users(Index).Nick
                    Users(Index).Onchannels.Add Chan, Chan
                    SendWsock Index, ":" & Users(Index).Nick & " JOIN " & Chan, True
                    SendWsock Index, ":" & ServerName & " 353 " & Users(Index).Nick & " = " & Chan & " :" & FixNickList((Replace(JoinChan.GetOps & " " & JoinChan.GetVoices & " " & JoinChan.GetNorms, "  ", " "))), True
                    SendWsock Index, ":" & ServerName & " 366 " & Users(Index).Nick & " " & Chan & " :End of /NAMES list.", True
                    SendWsock Index, ":" & ServerName & " 332 " & Users(Index).Nick & " " & Chan & " :" & JoinChan.Topic, True
                    SendWsock Index, ":" & ServerName & " 333 " & JoinChan.TopicSetBy & " " & Chan & " " & JoinChan.TopicSetBy & " " & JoinChan.TopicSetOn, True
                End If
            End If
        Next X
    ElseIf strcmd(i) Like "PART*" Then
        Chan = Replace(strcmd(i), "PART ", "")
        If InStr(1, Chan, " ") Then Chan = Mid(Chan, 1, InStr(1, Chan, " ") - 1)
        Dim Reason As String
        Reason = Replace(strcmd(i), "PART " & Chan & " :", "")
        If Reason = strcmd(i) Then Reason = ""
        SendPart CLng(Index), Chan, Reason
    ElseIf (strcmd(i) Like "MODE*") = True Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 250
        Dim Mode As String, ToUser() As String, Modes() As String, OP As String, ToUsers As String, Channel As clsIRCServer_Channel, Y As Long
        Chan = Replace(strcmd(i), "MODE ", "")
        If InStr(1, Chan, " ") <> 0 Then
            Chan = Mid(Chan, 1, InStr(1, Chan, " ") - 1)
        End If
        Set Channel = ChanToObject(Chan)
        Set User = NickToObject(Chan)
        If Not User Is Nothing Then
            Dim UM As String
            UM = Mid(Replace(strcmd(i), "MODE " & User.Nick & " ", "", , , vbTextCompare), 1, 1)
            Select Case UM
                Case "+"
                    AddUserMode User.Index, Mid(Replace(strcmd(i), "MODE " & User.Nick & " ", "", , , vbTextCompare), 2)
                Case "-"
                    RemoveUsermode User.Index, Mid(Replace(strcmd(i), "MODE " & User.Nick & " ", "", , , vbTextCompare), 2)
            End Select
            GoTo NextCmd
        End If
        Dim cmdline() As String, UserMode As Boolean
        cmdline = Split(strcmd(i), " ")
        For X = LBound(cmdline) To UBound(cmdline)
            If Not NickToObject(cmdline(X)) Is Nothing Then UserMode = True
        Next X
        If InStr(1, strcmd(i), "*") <> 0 Then UserMode = True
        If strcmd(i) Like "MODE * +?" Then UserMode = False
        If UserMode Then
            Mode = Replace(strcmd(i), "MODE " & Chan & " ", "")
            OP = Mid(Mode, 1, 1)
            Mode = Mid(Mode, 2, InStr(1, Mode, " ") - 2)
            ToUsers = Mid(strcmd(i), InStr(1, strcmd(i), OP) + Len(Mode) + 2)
            If InStr(1, ToUsers, " ") <> 0 Then
                ToUser = Split(" " & ToUsers, " ")
            Else
                ReDim ToUser(1)
                ToUser(1) = ToUsers
            End If
            If Channel.IsOp(Users(Index).Nick) = False And (Not Users(Index).IsOwner(Channel.Name)) And Not Users(Index).IRCOp Then
                SendWsock Index, ":" & ServerName & " 482 " & Users(Index).Nick & " " & Chan & " :You're not channel operator"
                GoTo NextCmd
            End If
            For X = 1 To Len(Mode)
                ReDim Preserve Modes(X)
                Modes(X) = Mid(Mode, X, 1)
            Next X
            ReDim Preserve Modes(UBound(ToUser))
            For Y = LBound(Modes) To UBound(Modes)
                If Not ToUser(Y) = "" Then
                    Select Case Modes(IIf((Y = 0), Y + 1, Y))
                        Case "o"
                            Select Case OP
                                Case "+"
                                    OpUser Channel, ToUser(Y), Users(Index).Nick
                                Case "-"
                                    DeOpUser Channel, ToUser(Y), Users(Index).Nick
                            End Select
                        Case "v"
                            Select Case OP
                                Case "+"
                                    VoiceUser Channel, ToUser(Y), Users(Index).Nick
                                Case "-"
                                    DeVoiceUser Channel, ToUser(Y), Users(Index).Nick
                            End Select
                        Case "b"
                            Select Case OP
                                Case "+"
                                    BanUser Channel, ToUser(Y), Users(Index).Nick
                                Case "-"
                                    UnBanUser Channel, ToUser(Y), Users(Index).Nick
                            End Select
                        Case "e"
                            Select Case OP
                                Case "+"
                                    ExceptionUser Channel, ToUser(Y), Users(Index).Nick
                                Case "-"
                                    UnExceptionUser Channel, ToUser(Y), Users(Index).Nick
                            End Select
                        Case "I"
                            Select Case OP
                                Case "+"
                                    InviteUser Channel, ToUser(Y), Users(Index).Nick
                                Case "-"
                                    UnInviteUser Channel, ToUser(Y), Users(Index).Nick
                            End Select
                    End Select
                End If
            Next Y
        Else
            If InStr(1, strcmd(i), " +b", vbBinaryCompare) <> 0 Then
                For X = 1 To Channel.Bans.Count
                    SendWsock Index, ":" & ServerName & " 367 " & Users(Index).Nick & " " & Channel.Name & " " & Channel.Bans(X)
                Next X
                SendWsock Index, ":" & ServerName & " 368 " & Users(Index).Nick & " " & Channel.Name & " :End of Channel Ban List"
            ElseIf InStr(1, strcmd(i), " +e", vbBinaryCompare) <> 0 Then
                For X = 1 To Channel.Exceptions.Count
                    SendWsock Index, ":" & ServerName & " 348 " & Users(Index).Nick & " " & Channel.Name & " " & Channel.Exceptions(X)
                Next X
                SendWsock Index, ":" & ServerName & " 349 " & Users(Index).Nick & " " & Channel.Name & " :End of Channel Exceptions List"
            ElseIf InStr(1, strcmd(i), " +I", vbBinaryCompare) <> 0 Then
                For X = 1 To Channel.Invites.Count
                    SendWsock Index, ":" & ServerName & " 346 " & Users(Index).Nick & " " & Channel.Name & " " & Channel.Invites(X)
                Next X
                SendWsock Index, ":" & ServerName & " 347 " & Users(Index).Nick & " " & Channel.Name & " :End of Channel Invites List"
            ElseIf InStr(1, strcmd(i), " +w", vbBinaryCompare) <> 0 Then
                SendWsock Index, ":" & ServerName & " 472 " & Users(Index).Nick & " w :is unknown mode char to me"
            ElseIf InStr(1, strcmd(i), "+") <> 0 Then
                AddChanModes Mid(strcmd(i), InStr(1, strcmd(i), "+") + 1), Chan, Users(Index)
            ElseIf InStr(1, strcmd(i), "-") <> 0 Then
                RemoveChanModes Mid(strcmd(i), InStr(1, strcmd(i), "-") + 1), Chan, Users(Index)
            Else
                SendWsock Index, ":" & ServerName & " 324 " & Users(Index).Nick & " " & Channel.Name & " " & Channel.GetModes
            End If
        End If
    ElseIf (strcmd(i) Like "TOPIC*") = True Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 250
        If InStr(1, strcmd(i), " :") <> 0 Then
            Dim NewTopic As String
            Chan = Replace(strcmd(i), "TOPIC ", "")
            Chan = Mid(Chan, 1, InStr(1, Chan, " ") - 1)
            Set Channel = ChanToObject(Chan)
            If Channel Is Nothing Then GoTo NextCmd
            NewTopic = strcmd(i)
            NewTopic = Mid(NewTopic, InStr(1, NewTopic, ":") + 1)
            If Len(NewTopic) > TopicLen Then NewTopic = Left(NewTopic, TopicLen)
            If Channel.IsOp(Users(Index).Nick) = False And (Not Users(Index).IsOwner(Channel.Name)) And Channel.IsMode("t") Then
                SendWsock Index, ":" & ServerName & " 482 " & Users(Index).Nick & " :You're not channel operator"
                GoTo NextCmd
            End If
            SetTopic Chan, NewTopic, Users(Index).Nick
        Else
            Chan = Replace(strcmd(i), "TOPIC ", "")
            Set Channel = ChanToObject(Chan)
            If Channel Is Nothing Then GoTo NextCmd
            SendWsock Index, ":" & ServerName & " 332 " & Users(Index).Nick & " " & Chan & " :" & Channel.Topic
            SendWsock Index, ":" & ServerName & " 333 " & Channel.TopicSetBy & " " & Chan & " " & Users(Index).Nick & " " & Channel.TopicSetOn
        End If
    ElseIf (strcmd(i) Like "INVITE*") = True Then
        Dim Target As String
        Target = Replace(strcmd(i), "INVITE ", "")
        Target = Mid(Target, 1, InStr(1, Target, " ") - 1)
        Chan = Mid(strcmd(i), Len("INVITE " & Target & " ") + 1)
        Set Channel = ChanToObject(Chan)
        If Channel.IsMode("i") And Channel.IsOp(Users(Index).Nick) = False Then
            SendWsock Index, ":" & ServerName & " 482 " & Users(Index).Nick & " " & Chan & " :You're not channel operator"
            GoTo NextCmd
        End If
        If lSettings.sHandleErrors = True Then On Local Error Resume Next
        Channel.Invited.Add Target, Target
        SendWsock NickToObject(Target).Index, ":" & Users(Index).Nick & " INVITE " & Target & " " & Chan
        SendLinks "AddInvite" & vbLf & Users(Index).Nick & vbLf & Channel.Name & vbLf & "" & vbLf & Target
    ElseIf (strcmd(i) Like "KICK*") = True Then
        Dim Source As String
        Chan = Mid(strcmd(i), 6)
        Chan = Mid(Chan, 1, InStr(1, Chan, " ") - 1)
        Set Channel = ChanToObject(Chan)
        Source = Users(Index).Nick
        If Channel.IsOp(Source) = False And (Not Users(Index).IsOwner(Channel.Name)) Then
            SendWsock Index, ":" & ServerName & " 482 " & Users(Index).Nick & " " & Chan & " :You're not channel operator"
            GoTo NextCmd
        End If
        If InStr(1, strcmd(i), ":") <> 0 Then
            Reason = Mid(strcmd(i), InStr(1, strcmd(i), " :") + 2)
            If Len(Reason) > KickLen Then Reason = Left(Reason, KickLen)
            Target = Replace(strcmd(i), "KICK", "")
            Target = Mid(Target, 2)
            Target = Mid(Target, 1, InStr(1, Target, ":") - 2)
            Target = Replace(Target, Chan & " ", "")
            If Target = "ChanServ" Then
                SendSvrMsg Source & " tried to kick services[" & Chan & "]", True, ServerName
                SendWsock Index, ":" & ServerName & " 404 " & Source & " " & Chan & " :Cannot kick Services"
                GoTo NextCmd
            End If
            KickUser Source, Chan, Target, Reason, True
            GoTo NextCmd
        End If
        Target = Mid(strcmd(i), InStrRev(strcmd(i), " ", InStrRev(strcmd(i), " ")) + 1)
        If Target = "ChanServ" Then
            SendSvrMsg Source & " tried to kick services[" & Chan & "]", True, ServerName
            SendWsock Index, ":" & ServerName & " 404 " & Source & " " & Chan & " :Cannot kick Services"
            GoTo NextCmd
        End If
        KickUser Source, Chan, Target
    ElseIf (strcmd(i) Like "PONG*") = True Then
        Users(Index).Ponged = True
        If Users(Index).NewUser Then
            SendWsock Index, GetWelcome(CLng(Index))
            SendWsock Index, ReadMotd(Users(Index).Nick)
            Users(Index).NewUser = False
            SendLogonNews Index
            SendLinks "NewUser" & vbLf & Users(Index).Nick & vbLf & Users(Index).Name & vbLf & Users(Index).DNS & vbLf & Users(Index).ident & vbLf & Users(Index).Server & vbLf & Users(Index).ServerDescritption & vbLf & Users(Index).SignOn & vbLf & Users(Index).GID & vbLf & Users(Index).GetModes & vbLf & ServerName & " "
            CloneControl.Add wsock(Index).RemoteHostIP
            Users(Index).MsgsSent = 0
            tmrFloodProt(Index).Enabled = True
            If DefUserModes <> "" Then AddUserMode CLng(Index), DefUserModes
        End If
        ElseIf (strcmd(i) Like "PING*") = True Then
            SendWsock Index, "PONG " & Replace(strcmd(i), "PING ", ""), True
    ElseIf (strcmd(i) Like "PRIVMSG*") = True Then
        cmdline = Split(Mid(strcmd(i), 1, InStr(1, strcmd(i), " :")), " ")
        For X = LBound(cmdline) To UBound(cmdline)
            If Not NickToObject(cmdline(X)) Is Nothing Then UserMode = True
        Next X
        Target = Replace(strcmd(i), "PRIVMSG ", "")
        Target = Strings.Left(Target, InStr(1, Target, ":") - 2)
        cmdline = Split(Target, ",")
        For X = LBound(cmdline) To UBound(cmdline)
            Target = cmdline(X)
            If Not NickToObject(Target) Is Nothing Then UserMode = True
            Select Case LCase(Target)
                Case "chanserv"
                    UserMode = True
                Case "nickserv"
                   UserMode = True
                Case "memoserv"
                   UserMode = True
                Case "operserv"
                   UserMode = True
            End Select
            If (Not UserMode) Then
                Dim msgstr As String, msg As String
                Chan = Target
                Set Channel = ChanToObject(Chan)
                If Channel Is Nothing Then
                    SendWsock Index, ":" & ServerName & " 404 " & Users(Index).Nick & " " & Chan & " :Cannot send to channel"
                    GoTo NextCmd
                End If
                If Channel.IsMode("n") Then
                    If (Not Channel.IsOnChan(Users(Index).Nick) And (Not Users(Index).IsOwner(Channel.Name)) And Not Users(Index).IRCOp) Then
                        SendWsock Index, ":" & ServerName & " 404 " & Users(Index).Nick & " " & Chan & " :Cannot send to channel"
                        GoTo NextCmd
                    End If
                End If
                If Channel.IsBanned(Users(Index)) Then
                   If ((Not Users(Index).IsOwner(Channel.Name)) And Not Users(Index).IRCOp) Then
                        SendWsock Index, ":" & ServerName & " 404 " & Users(Index).Nick & " " & Chan & " :Cannot send to channel"
                        GoTo NextCmd
                    End If
                End If
                If Channel.IsMode("m") Then
                    If Channel.IsOp(Users(Index).Nick) Then
                    ElseIf Channel.IsVoice(Users(Index).Nick) Or (Users(Index).IsOwner(Channel.Name)) And Not Users(Index).IRCOp Then
                    Else
                        SendWsock Index, ":" & ServerName & " 404 " & Users(Index).Nick & " " & Chan & " :Cannot send to channel"
                        GoTo NextCmd
                    End If
                End If
                msg = strcmd(i)
                msg = Mid(msg, InStr(1, msg, ":") + 1)
                If Len(msg) > Msglen Then msg = Left(msg, Msglen)
                SendMsg Chan, msg, Users(Index).Nick
            Else
                Target = Replace(strcmd(i), "PRIVMSG ", "")
                Target = Strings.Left(Target, InStr(1, Target, ":") - 2)
                msg = strcmd(i)
                msg = Mid(msg, InStr(1, msg, ":") + 1)
                If Len(msg) > Msglen Then msg = Left(msg, Msglen)
                Set User = NickToObject(Target)
                If User Is Nothing Then
                    SendWsock Index, ":" & ServerName & " 401 " & Users(Index).Nick & " " & Target & " :No such nick/channel"
                    GoTo NextCmd
                End If
                SendMsg Target, msg, Users(Index).Nick, False
            End If
        Next X
    ElseIf (strcmd(i) Like "NOTICE*") = True Then
        Target = Replace(strcmd(i), "NOTICE ", "")
        Target = Replace(Target, ":*", " ")
        Target = Left(Target, InStr(1, Target, ":") - 2)
        Dim Targets() As String
        Targets = Split(Target, ",")
        msg = strcmd(i)
        msg = Mid(msg, InStr(1, msg, ":") + 1)
        If Len(msg) > Msglen Then msg = Left(msg, Msglen)
        For Y = LBound(Targets) To UBound(Targets)
            Target = Targets(Y)
            If InStr(1, Target, "#") = 0 Then
                If NickToObject(Target) Is Nothing Then
                    SendWsock Index, ":" & ServerName & " 401 " & Users(Index).Nick & " :No such nick/channel"
                    GoTo NextCmd
                End If
                SendNotice Target, msg, Users(Index).Nick
            Else
                Dim CurChan As clsIRCServer_Channel
                Set CurChan = ChanToObject(Target)
                If CurChan Is Nothing Then
                    SendWsock Index, ":" & ServerName & " 401 " & Users(Index).Nick & " :No such nick/channel"
                    GoTo NextCmd
                End If
                    If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + (10 * CurChan.All.Count)
                    SendNotice Target, msg, Users(Index).Nick, True
            End If
        Next Y
    ElseIf (strcmd(i) Like "MOTD") = True Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 1200
        SendWsock Index, ReadMotd(Users(Index).Nick)
    ElseIf (strcmd(i) Like "WHOIS*") = True Then
        Dim WhoisStr As String, Nick As String
        Set User = NickToObject(Replace(strcmd(i), "WHOIS ", ""))
        If Not User Is Nothing Then
            SendWsock Index, User.GetWhois(Users(Index).Nick)
        Else
            SendWsock Index, ":" & ServerName & " 401 " & Users(Index).Nick & " :No such nick/channel"
        End If
    ElseIf (strcmd(i) Like "AWAY*") = True Then
        If Not Users(Index).Away Then
            Users(Index).AwayMsg = Replace(strcmd(i), "AWAY :", "")
            Users(Index).Away = True
            SendWsock Index, ":" & ServerName & " 306 " & Users(Index).Nick & " :You have been marked as being away"
            Users(Index).Modes.Add "a", "a"
        Else
            Users(Index).Away = False
            Users(Index).AwayMsg = ""
            RemoveUsermode CLng(Index), "a", True
        End If
    ElseIf (strcmd(i) Like "WALLOPS*") = True Then
        If Users(Index).IRCOp Then
            WallOps Replace(strcmd(i), "WALLOPS ", ""), Index
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
    ElseIf (strcmd(i) Like "WALL*") = True Then
        If Users(Index).IRCOp Then
            Wall Replace(strcmd(i), "WALL ", ""), Index
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
    ElseIf (strcmd(i) Like "VERSION") = True Then
        SendWsock Index, GetWelcome(CLng(Index))
    ElseIf (strcmd(i) Like "TIME") = True Then
        SendWsock Index, ":" & ServerName & " 391" & Users(Index).Nick & " " & ServerName & " :" & Now
    ElseIf (strcmd(i) Like "VERSION") = True Then
        SendWsock Index, GetWelcome(CLng(Index))
    ElseIf (strcmd(i) Like "ISON*") = True Then
        Dim strIsOn As String, LoggedIn() As String, IsOnArr() As String
        ReDim LoggedIn(1)
        strIsOn = Replace(strcmd(i), "ISON ", "")
        IsOnArr = Split(strIsOn, " ")
        For X = LBound(IsOnArr) To UBound(IsOnArr)
            If Not NickToObject(IsOnArr(X)) Is Nothing Then
                ReDim Preserve LoggedIn(UBound(LoggedIn) + 1)
                LoggedIn(UBound(LoggedIn)) = IsOnArr(X)
            End If
        Next X
        strIsOn = Join(LoggedIn, " ")
        SendWsock Index, (":" & ServerName & " 303 " & Users(Index).Nick & " :" & strIsOn)
    ElseIf (strcmd(i) Like "LUSERS*") = True Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 400
        SendWsock Index, ":" & ServerName & " 252 " & Users(Index).Nick & " :" & Operators & " Operator(s) online" & vbCrLf & _
                                           ":" & ServerName & " 254 " & Users(Index).Nick & " :channels formed = " & lChanCount & vbCrLf & _
                                           ":" & ServerName & " 255 " & Users(Index).Nick & " :I have " & lUserCount - 4 & " clients and " & (CurLinkCount + 1) & " Servers" & vbCrLf & _
                                           ":" & ServerName & " 265 " & Users(Index).Nick & " :Current Local Users : " & (lUserCount - 4) & " Max Local Users : " & lMaxUser & vbCrLf & _
                                           ":" & ServerName & " 266 " & Users(Index).Nick & " :Current Global Users: " & CurGlobalUsers & " Max Global Users: " & MaxGlobalUsers & vbCrLf
    ElseIf (strcmd(i) Like "STATS*") = True Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 250
        Dim StatsParam As String
        StatsParam = Replace(strcmd(i), "STATS ", "")
        Select Case StatsParam
            Case "u"
                SendWsock Index, ":" & ServerName & " 242 " & Users(Index).Nick & " :" & CStr(Started)
                SendWsock Index, ":" & ServerName & " 250 " & Users(Index).Nick & " :Highest Connection Count: " & lMaxUser
                SendWsock Index, ":" & ServerName & " 219 " & Users(Index).Nick & " u :End of /STATS report"
        End Select
    ElseIf (strcmd(i) Like "INFO*") = True Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 250
        SendWsock Index, ":" & ServerName & " 371 " & Users(Index).Nick & " :" & ServerName & " running NexIRC " & App.Major & "." & App.Minor & "." & App.Revision
        SendWsock Index, ":" & ServerName & " 371 " & Users(Index).Nick & " :This server was created on Wednesday July 28th by Leon J Aiossa (guide_X@live.com"
        SendWsock Index, ":" & ServerName & " 374 " & Users(Index).Nick & " :End of INFO list"
    ElseIf (strcmd(i) Like "LINKS*") = True Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 200
        SendWsock Index, ":" & ServerName & " 364 " & Users(Index).Nick & " " & ServerName & " " & ServerName & " :0 " & ServerDesc
        SendWsock Index, ":" & ServerName & " 365 " & Users(Index).Nick & " * :End of /LINKS list"
    ElseIf strcmd(i) Like "NAMES*" Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 150
        Chan = Replace(strcmd(i), "NAMES ", "")
        Set Channel = ChanToObject(Chan)
        SendWsock Index, ":" & ServerName & " 353 " & Users(Index).Nick & " = " & Chan & " :" & FixNickList((Replace(Channel.GetOps & " " & Channel.GetVoices & " " & Channel.GetNorms, "  ", " ")))
        SendWsock Index, ":" & ServerName & " 366 " & Users(Index).Nick & " " & Chan & " :End of /NAMES list."
    ElseIf strcmd(i) Like "LIST*" Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 1000
        SendWsock Index, ":" & ServerName & " 321 " & Users(Index).Nick & " Channel :Users  Name"
        SendWsock Index, GetChanList(Users(Index).Nick)
        SendWsock Index, ":" & ServerName & " 323 " & Users(Index).Nick & " :End of /LIST"
    ElseIf strcmd(i) Like "ADMIN*" Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 250
        SendWsock Index, ":" & ServerName & " 256 " & Users(Index).Nick & " :Administrative info about " & ServerName
        SendWsock Index, ":" & ServerName & " 257 " & Users(Index).Nick & " :" & ServerDesc
        SendWsock Index, ":" & ServerName & " 258 " & Users(Index).Nick & " :" & AdminName
        SendWsock Index, ":" & ServerName & " 259 " & Users(Index).Nick & " :" & AdminEmail
    ElseIf strcmd(i) Like "WHO*" Then
        Dim strSearch As String
        strSearch = Replace(strcmd(i), "WHO ", "")
        For X = 1 To UBound(Users)
            DoEvents
            If Not Users(X) Is Nothing Then
                If (Users(X).Nick Like strSearch) Then
                    SendWsock Index, ":" & ServerName & " 352 * " & LCase(Users(Index).Nick) & " " & Users(X).Nick & " " & Users(X).DNS & " " & ServerName & " " & Users(X).Nick & " H :" & Users(X).Hops & " " & Users(X).Name
                End If
            End If
        Next X
        SendWsock Index, ":" & ServerName & " " & 315 & " " & Users(Index).Nick & " " & strSearch & " :END of /WHO list."
    ElseIf (strcmd(i) Like "OPER *") = True Then
        Dim PW As String, UserName As String
        UserName = Replace(strcmd(i), "OPER ", "")
        PW = Mid(UserName, InStr(1, UserName, " ") + 1)
        UserName = Mid(UserName, 1, InStr(1, UserName, " ") - 1)
        PW = Replace(PW, ":", "")
        If Not HasOline(Users(Index).Nick, Users(Index).GetMask) Then
            SendWsock Index, ":" & ServerName & " 491 " & Users(Index).Nick & " :No O-lines for your host"
            GoTo NextCmd
        End If
        With Olines(GetOline(Users(Index).DNS))
            If Not Users(Index).Nick = .UserName Then
                SendWsock Index, ":" & ServerName & " 491 " & Users(Index).Nick & " :your nickname must match the nickname with which the O-Line has been created"
                GoTo NextCmd
            End If
            If Not PW = .Password Then
                SendWsock Index, ":" & ServerName & " 464 " & Users(Index).Nick & " :Password incorrect"
                GoTo NextCmd
            End If
            SendWsock Index, ":" & ServerName & " 381 " & Users(Index).Nick & " :You are now an IRC operator"
            SendLinks "ModeUser" & vbLf & Users(Index).Nick & vbLf & "+" & vbLf & "o"
            SendWsock Index, ":" & Users(Index).Nick & " MODE " & Users(Index).Nick & " +o"
            If lSettings.sHandleErrors = True Then On Local Error Resume Next
            AddUserMode CLng(Index), "o"
            Users(Index).AddModes "o"
            Users(Index).IRCOp = True
            SendSvrMsg Users(Index).Nick & " is now Operator", True
            Users(Index).RealDNS = Users(Index).DNS
            Users(Index).DNS = ServerName
            Operators = Operators + 1
        End With
    ElseIf (strcmd(i) = "RESTART") = True Then
        If Users(Index).IRCOp Then
            For X = LBound(Users) To UBound(Users)
                If Not Users(X) Is Nothing Then SendNotice "", "• Global -- " & "Recieved RESTART command from " & Users(Index).Nick, "GLOBAL", , CInt(X)
            Next X
            Wait 2000
            Restart
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
    ElseIf strcmd(i) Like "DIE*" Then
        If Users(Index).IRCOp Then
            For X = LBound(Users) To UBound(Users)
                If Not Users(X) Is Nothing Then SendNotice "", "• Global -- " & "Recieved DIE command from " & Users(Index).Nick, "GLOBAL", , CInt(X)
            Next X
            Wait 2000
            Unload Me
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
    ElseIf (strcmd(i) Like "KLINE*") = True Then
        If Not Users(Index).IRCOp Then
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
            GoTo NextCmd
        End If
        Klines.Add Replace(strcmd(i), "KLINE ", "")
    ElseIf (strcmd(i) Like "KILL*") = True Then
        If Users(Index).IRCOp Then
            Dim NickName As String, Comment As String
            NickName = Replace(strcmd(i), "KILL ", "")
            If InStr(1, NickName, " ") = 0 Then
                Comment = Users(Index).Nick
            Else
                NickName = Mid(NickName, 1, InStr(1, NickName, " :") - 1)
                Comment = Replace(strcmd(i), "KILL " & NickName & " :", "")
            End If
            Set User = NickToObject(NickName, , True)
            If Not User Is Nothing Then
                SendLinks "KillUser" & vbLf & User.Nick & vbLf & Comment
                SendQuit User.Index, "Killed by " & Users(Index).Nick & " (" & Comment & ")", True
                If Not User.LocalUser Then GoTo NextCmd
                User.Killed = True
                SendWsock User.Index, ":" & Users(Index).Nick & "!" & Users(Index).ident & "@" & ServerName & " KILL " & User.Nick & " :" & Comment, True
                SendWsock User.Index, "ERROR :Closing Link: " & User.Nick & "[" & frmIRCServer.wsock(User.Index).RemoteHostIP & ".] " & ServerName & " (" & Comment & ")", True
                Dim Kline As Long
                Kline = GetRand
                Load tmrKlined(Kline)
                tmrKlined(Kline).Tag = wsock(User.Index).RemoteHostIP
                tmrKlined(Kline).Enabled = True
                If lSettings.sHandleErrors = True Then On Local Error Resume Next
                Klines.Add wsock(User.Index).RemoteHostIP, wsock(User.Index).RemoteHostIP
                SendNotice Users(Index).Nick, User.Nick & " has been removed from the network", "" & ServerName & ""
                SendSvrMsg "Recieved Kill message for " & User.Nick & "!" & User.ident & "@" & User.DNS & " Path: " & Users(Index).Nick & " (" & Comment & ")", True
            Else
                SendWsock Index, ":" & ServerName & " 401 " & Users(Index).Nick & " :No such nick/channel, using wildcards instead"
                For X = 1 To UBound(Users)
                    If Not Users(X) Is Nothing Then
                        If (Users(X).Nick & "!" & Users(X).ident & "@" & Users(X).DNS) Like NickName Then
                            Set User = Users(X)
                            SendLinks "KillUser" & vbLf & User.Nick & vbLf & Comment
                            SendQuit User.Index, "Killed by " & Users(Index).Nick & " (" & Comment & ")", True
                            If Not User.LocalUser Then GoTo NextCmd
                            User.Killed = True
                            SendWsock User.Index, ":" & Users(Index).Nick & "!" & Users(Index).ident & "@" & ServerName & " KILL " & User.Nick & " :" & Comment, True
                            SendWsock User.Index, "ERROR :Closing Link: " & User.Nick & "[" & frmIRCServer.wsock(User.Index).RemoteHostIP & ".] " & ServerName & " (" & Comment & ")", True
                            SendQuit User.Index, "Killed by " & Users(Index).Nick & " (" & Comment & ")", True
                            SendSvrMsg "Recieved Kill message for " & User.Nick & "!" & User.ident & "@" & User.DNS & " Path: " & Users(Index).Nick & " (" & Comment & ")", True
                        End If
                    End If
                Next X
            End If
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
    ElseIf (strcmd(i) Like "AKILL*") = True Then
        If Users(Index).IRCOp Then
            NickName = Replace(strcmd(i), "AKILL ", "")
            If InStr(1, NickName, " ") = 0 Then
                Comment = Users(Index).Nick
            Else
                NickName = Mid(NickName, 1, InStr(1, NickName, " ") - 1)
                Comment = Replace(strcmd(i), "AKILL " & NickName & " ", "")
            End If
            Set User = NickToObject(NickName, , True)
            If Not User Is Nothing Then
                SendLinks "KillUser" & vbLf & User.Nick & vbLf & Comment
                SendQuit User.Index, "AKilled by " & Users(Index).Nick & " (" & Comment & ")", True
                If Not User.LocalUser Then GoTo NextCmd
                User.Killed = True
                SendWsock User.Index, ":" & Users(Index).Nick & "!" & Users(Index).ident & "@" & ServerName & " KILL " & User.Nick & " :" & Comment, True
                SendWsock User.Index, "ERROR :Closing Link: " & User.Nick & "[" & frmIRCServer.wsock(User.Index).RemoteHostIP & ".] " & ServerName & " (" & Comment & ")", True
                Klines.Add wsock(User.Index).RemoteHostIP, wsock(User.Index).RemoteHostIP
                SendNotice Users(Index).Nick, User.Nick & " has been removed from the network", "" & ServerName & ""
                SendSvrMsg "Recieved AKill message for " & User.Nick & "!" & User.ident & "@" & User.DNS & " Path: " & Users(Index).Nick & " (" & Comment & ")", True
            Else
                SendWsock Index, ":" & ServerName & " 401 " & Users(Index).Nick & " :No such nick/channel, using wildcards instead"
                For X = 5 To UBound(Users)
                    If Not Users(X) Is Nothing Then
                        If (Users(X).Nick & "!" & Users(X).ident & "@" & Users(X).DNS) Like NickName Then
                            Set User = Users(X)
                            SendLinks "KillUser" & vbLf & User.Nick & vbLf & Comment
                            SendQuit User.Index, "AKilled by " & Users(Index).Nick & " (" & Comment & ")", True
                            If Not User.LocalUser Then GoTo NextCmd
                            User.Killed = True
                            SendWsock User.Index, ":" & Users(Index).Nick & "!" & Users(Index).ident & "@" & ServerName & " KILL " & User.Nick & " :" & Comment, True
                            SendWsock User.Index, "ERROR :Closing Link: " & User.Nick & "[" & frmIRCServer.wsock(User.Index).RemoteHostIP & ".] " & ServerName & " (" & Comment & ")", True
                            If lSettings.sHandleErrors = True Then On Local Error Resume Next
                            Klines.Add wsock(User.Index).RemoteHostIP, wsock(User.Index).RemoteHostIP
                            SendQuit User.Index, "AKilled by " & Users(Index).Nick & " (" & Comment & ")", True
                            SendSvrMsg "Recieved AKill message for " & User.Nick & "!" & User.ident & "@" & User.DNS & " Path: " & Users(Index).Nick & " (" & Comment & ")", True
                        End If
                    End If
                Next X
            End If
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
    ElseIf strcmd(i) Like "REHASH*" Then
        If Users(Index).IRCOp Then
            Rehash Users(Index).Nick
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
    ElseIf strcmd(i) Like "CLIENTINFO*" Then
        If Users(Index).IRCOp Then
            Set User = NickToObject(Replace(strcmd(i), "CLIENTINFO ", ""))
            SendWsock Index, ":NickName = " & User.Nick
            SendWsock Index, ":Ident = " & User.ident
            SendWsock Index, ":Name = " & User.Name
            SendWsock Index, ":Email = " & User.Email
            SendWsock Index, ":Modes = " & User.GetModes
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
    ElseIf strcmd(i) Like "DELETE*" Then
        If Users(Index).IRCOp Then
            Set Users(NickToObject(Replace(strcmd(i), "DELETE ", "")).Index) = Nothing
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
    ElseIf strcmd(i) Like "CONNECT*" Then
        Dim sName As String, sPort As String
        sName = Replace(strcmd(i), "CONNECT ", "")
        sPort = Mid(sName, InStr(1, sName, " ") + 1)
        sName = Replace(sName, " " & sPort, "")
        If sPort = 0 Then sPort = 6668
        If Users(Index).IRCOp Then
            Dim LinkCount As Long
            LinkCount = Link.Count + 1
            CurLinkCount = CurLinkCount + 1
            MaxLinkCount = MaxLinkCount + 1
            Load Link(LinkCount)
            Link(LinkCount).LocalPort = 0
            Link(LinkCount).Connect sName, sPort
            Link(LinkCount).Tag = sName
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
    ElseIf strcmd(i) Like "SQUIT*" Then
        If Users(Index).IRCOp Then
            Dim CloseLink As Long
            sName = Replace(strcmd(i), "SQUIT ", "")
            For X = 2 To Link.UBound
                If lSettings.sHandleErrors = True Then On Local Error Resume Next
                If Link(X).Tag = sName Then
                    SendSvrMsg "Link closed by " & Users(Index).Nick & "[" & ServerName & " -- " & Link(X).Tag & "]", True
                    Link(X).Close
                    Link_Close (X)
                    Unload Link(X)
                    Exit For
                End If
            Next X
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
    ElseIf (strcmd(i) Like "NS*") = True Then
        SendMsg "NickServ", Replace(strcmd(i), "NS ", ""), Users(Index).Nick, False
    ElseIf (strcmd(i) Like "NICKSERV*") = True Then
        SendMsg "NickServ", Replace(strcmd(i), "NickServ ", ""), Users(Index).Nick, False
    ElseIf (strcmd(i) Like "MS*") = True Then
        SendMsg "MemoServ", Replace(strcmd(i), "MS ", ""), Users(Index).Nick, False
    ElseIf (strcmd(i) Like "MEMOSERV*") = True Then
        SendMsg "MemoServ", Replace(strcmd(i), "MemoServ ", ""), Users(Index).Nick, False
    ElseIf (strcmd(i) Like "CS*") = True Then
        SendMsg "ChanServ", Replace(strcmd(i), "CS ", ""), Users(Index).Nick, False
    ElseIf (strcmd(i) Like "CHANSERV*") = True Then
        SendMsg "ChanServ", Replace(strcmd(i), "ChanServ ", ""), Users(Index).Nick, False
    ElseIf (strcmd(i) Like "OS*") = True Then
        ParseOSCmd strcmd(i), CLng(Index)
    ElseIf (strcmd(i) Like "OPERSERV*") = True Then
        ParseOSCmd Replace(strcmd(i), "OPERSERV", "OS"), CLng(Index)
    ElseIf strcmd(i) = "" Then
    Else
        If InStr(1, strcmd(i), " ") <> 0 Then strcmd(i) = Mid(strcmd(i), 1, InStr(1, strcmd(i), " ") - 1)
        SendWsock Index, ":" & ServerName & "  421 " & Users(Index).Nick & " :" & strcmd(i) & " Unknown command"
    End If
NextCmd:
Next i
parseerr:
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Index < 5 Then Exit Sub
If Not Users(Index) Is Nothing Then SendWsock Index, ":" & ServerName & " 421 " & Users(Index).Nick & " :Parsing error | Need more Parameters or wrong order of Parameters" & Err.Description
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub wsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)"
End Sub

Private Sub wsock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
SendQuit CLng(Index), "Connection error: " & Description, False
wsock_Close (Index)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub wsock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)"
End Sub

Private Sub wsock_SendComplete(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Users(Index).Killed Then wsock_Close (Index)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub wsock_SendComplete(Index As Integer)"
End Sub
