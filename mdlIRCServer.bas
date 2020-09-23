Attribute VB_Name = "mdlIRCServer"
Option Explicit
Option Base 1
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function CoCreateGuid Lib "ole32" (ID As Any) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hDC As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, lParam As Any) As Long
Global lChanCount As Long
Global lUserCount As Long
Global lMaxUser As Long
Global lChannelsUBound As Long
Global lUsersUBound As Long
Global MaxNickRegs As Long
Global MaxChanRegs As Long
Public CurLinkCount As Long
Public MaxLinkCount As Long
Global MaxChunkSize As Long
Global Users() As clsIRCServer_User
Global Channels() As clsIRCServer_Channel
Global Olines(100) As Oline
Global FS As New FileSystemObject
Global dB As New clsDatabase
Public ServerName As String
Public Started As Date
Public Klines As New Collection
Public CloneControl As New Collection
Public ServerTraffic As Double
Public OverAllMax As Long
Public DefTopic As String
Public DefUserModes As String
Public DefQuit As String
Public MaxChannels As String
Public ServerDesc As String
Public AdminName As String
Public AdminEmail As String
Public SessionLimit As Long
Public Nicklen As Integer
Public MaxJoinChannels As Integer
Public TopicLen As Integer
Public KickLen As Integer
Public Msglen As Integer
Public AwayLen As Integer
Public Operators As Integer
Public LinkPort As Long
Public LogFile As String
Public LogLevel As Integer
Public LogFormat As Integer
Public LogStatusHandle As Long
Dim StatusInterval As Long
Dim StatusFile As String
Public CurGlobalUsers As Long
Public MaxGlobalUsers As Long
Public Type Oline
    UserName As String
    Password As String
    mask As String
    InUse As Boolean
End Type

Public Function UnixTime() As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UnixTime = DateDiff("s", DateValue("1/1/1970"), Now)
End Function

Public Function ChanToObject(lChanName As String) As clsIRCServer_Channel
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
Set ChanToObject = Nothing
For i = 1 To UBound(Channels)
    If Not Channels(i) Is Nothing Then
        If UCase(lChanName) = UCase(Channels(i).Name) Then
            Set ChanToObject = Channels(i)
            Exit Function
        End If
    End If
Next i
End Function

Public Function NickToObject(NickName As String, Optional StartAt As Long = 1, Optional LocalsOnly As Boolean = False) As clsIRCServer_User
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long, ub As Long
ub = UBound(Users)
For i = 1 To ub
    If Not Users(i) Is Nothing Then
        If UCase(NickName) = UCase(Users(i).Nick) Then
            Set NickToObject = Users(i)
            Exit Function
        End If
    End If
Next i
End Function

Public Function GetFreeSlot() As clsIRCServer_User
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
If Not UBound(Users) >= lUsersUBound Then
    ReDim Preserve Users(UBound(Users) + 1)
    For i = 1 To UBound(Users)
        If (Users(i) Is Nothing) Then
            Set Users(i) = New clsIRCServer_User
            Users(i).Index = i
            Set GetFreeSlot = Users(i)
            Exit Function
        End If
    Next i
End If
Set GetFreeSlot = Nothing
End Function

Public Sub SendWsock(Index, Message, Optional SendImmediately As Boolean = True, Optional SendMainSocket As Boolean = False)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Users(Index).Server <> ServerName Then Exit Sub
If Not Users(Index).LocalUser Then Exit Sub
If Index = 0 And SendMainSocket = False Then Exit Sub
If SendMainSocket Then
    frmIRCServer.wsock(Index).SendData Message & vbCrLf
    ServerTraffic = ServerTraffic + Len(Message & vbCrLf)
    Exit Sub
End If
ServerTraffic = ServerTraffic + Len(Message & vbCrLf)
If LogLevel = 1 Or LogLevel = 2 Then
    If LogFormat = 0 Then
        LogText "[Server]<" & Now & " (to " & Users(Index).Nick & ")> " & Message
    Else
        LogHTML ServerName, Message
    End If
End If
If Not SendImmediately Then
    frmIRCServer.wsock(Index).Tag = frmIRCServer.wsock(Index).Tag & Message & vbCrLf
Else
    Dim i As Long
    For i = 1 To Len(Message) Step MaxChunkSize
        frmIRCServer.wsock(Index).SendData Mid(Message, i, MaxChunkSize)
    Next i
    frmIRCServer.wsock(Index).SendData vbCrLf
End If
Debug.Print Message
End Sub

Public Function NickInUse(NickName As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case UCase(NickName)
    Case "NICKSERV"
        NickInUse = True
    Case "CHANSERV"
        NickInUse = True
    Case "MEMOSERV"
        NickInUse = True
    Case "OPERSERV"
        NickInUse = True
End Select
If NickInUse Then Exit Function
Dim i As Long
For i = 1 To UBound(Users)
    If Not Users(i) Is Nothing Then
        If UCase(Users(i).Nick) = UCase(NickName) Then
            NickInUse = True
            Exit Function
        End If
    End If
Next i
End Function

Public Function GetRandom() As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Randomize
Dim MyValue As Long, i As Long, R As Long
For i = 1 To 8
    MyValue = Int((9 * Rnd) + 0)
    R = CLng(CStr(R) & CStr(MyValue))
Next i
GetRandom = R
End Function

Public Function ChanExists(lChannelName As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Not ChanToObject(lChannelName) Is Nothing Then ChanExists = True
End Function

Public Function GetFreeChan() As clsIRCServer_Channel
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To UBound(Channels)
    If (Channels(i) Is Nothing) Then
        Set Channels(i) = New clsIRCServer_Channel
        Channels(i).Index = i
        Set GetFreeChan = Channels(i)
        Exit Function
    End If
Next i
Set GetFreeChan = Nothing
End Function

Public Function ReadMotd(Nick As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If FS.FileExists(App.Path & "\data\config\fixed\motd.txt") Then
    With FS.OpenTextFile(App.Path & "\data\config\fixed\motd.txt", ForReading)
        ReadMotd = ":" & ServerName & " 375 " & Nick & " :- " & ServerName & " Message of the day, " & vbCrLf
        ReadMotd = ReadMotd & ":" & ServerName & " 372 " & Nick & " :- " & Now & vbCrLf
        Do While (Not .AtEndOfStream)
            DoEvents
            ReadMotd = ReadMotd & ":" & ServerName & " 372 " & Nick & " :- " & .ReadLine & vbCrLf
        Loop
        ReadMotd = ReadMotd & ":" & ServerName & " 376 " & Nick & " :End of /MOTD command." & vbCrLf
    End With
End If
End Function

Public Function CountSpaces(strCount As String) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Len(strCount)
    If (Mid(strCount, i, 1) = " ") Then CountSpaces = CountSpaces + 1
Next i
CountSpaces = CountSpaces + 1
End Function

Public Sub ParseModeNicks(Nicks As String, ByRef Nickarr() As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If InStr(1, Nicks, " ") <> 0 Then
    Nickarr = Split(Nicks, " ")
Else
    ReDim Nickarr(1)
    Nickarr(1) = Nicks
End If
End Sub

Public Sub Restart()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
For i = LBound(Users) To UBound(Users)
    DoEvents
    If Not Channels(i) Is Nothing Then Set Channels(i) = Nothing
    If Not Users(i) Is Nothing Then
        SendQuit i, "Recieved RESTART Command, " & ServerName, True
        Set Users(i) = Nothing
        Unload frmIRCServer.wsock(i)
        Unload frmIRCServer.tmrTimeOut(i)
        Unload frmIRCServer.tmrSend(i)
        Unload frmIRCServer.tmrFloodProt(i)
    End If
Next i
Rehash
End Sub

Public Function GetWelcome(Index As Long) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim User As clsIRCServer_User
Set User = Users(Index)
GetWelcome = ":" & ServerName & " 001 " & User.Nick & " :Welcome to the Team Nexgen IRC Network " & User.Nick & "!" & User.ID & vbCrLf
GetWelcome = GetWelcome & ":" & ServerName & " 002 " & User.Nick & " :Your server is " & ServerName & ", running version NexIRC v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
GetWelcome = GetWelcome & ":" & ServerName & " 003 " & User.Nick & " :This server was created on Wednesday July 28th" & vbCrLf
GetWelcome = GetWelcome & ":" & ServerName & " 004 " & User.Nick & " " & ServerName & " NexIRC" & vbCrLf
GetWelcome = GetWelcome & ":" & ServerName & " 005 " & User.Nick & " SAFELIST NICKLEN=" & Nicklen & " CHANTYPES=# CHANMODES=beI,k,l,imnpst MAXCHANNELS=" & MaxJoinChannels & " MAXBANS=1 NETWORK=NexIRC EXCEPTS=e INVEX=I CASEMAPPING=ascii TOPICLEN=" & TopicLen & " KICKLEN=" & KickLen & " CHARSET=UTF-8 :are available on this server" & vbCrLf
End Function

Public Function GetRand() As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Randomize
Dim MyValue As Long, i As Long, R As Long
For i = 1 To 4
    MyValue = Int((9 * Rnd) + 0)
    R = CLng(CStr(R) & CStr(MyValue))
Next i
GetRand = R
End Function

Public Function IsKlined(IP As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Klines.Count
    DoEvents
    If IP Like Klines(i) Then
        IsKlined = Klines(i)
        Exit Function
    End If
Next i
End Function

Public Sub LogText(LogStr As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If LogLevel = 0 Then Exit Sub
FS.OpenTextFile(LogFile, ForAppending, True).WriteLine LogStr
End Sub

Public Sub LogHTML(Originator As String, LogStr)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If LogLevel = 0 Then Exit Sub
With FS.OpenTextFile(LogFile, ForAppending, True)
    .WriteLine "<tr>"
    .WriteLine "<td width=20%>" & Now & "</td>"
    .WriteLine "<td width=20%>" & Originator & "</td>"
    .WriteLine "<td width=60%>" & LogStr & "</td>"
    .WriteLine "</tr>"
End With
End Sub

Public Function SizeString(strData As String, Size As Long) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Size <= Len(strData) Then
    SizeString = strData
    Exit Function
End If
strData = strData & Space(Size - Len(strData))
SizeString = strData
End Function

Public Sub Rehash(Optional Nick As String = "NexIRC.MN.org")
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim dB As New clsDatabase, i As Long, Kline As String
dB.FileName = App.Path & "\data\config\server\server.ini"
ServerName = dB.ReadEntry("Settings", "Servername", "NexIRC.MN.org")
ServerDesc = dB.ReadEntry("Settings", "Description", "NexIRC Server")
frmIRCServer.wsock(0).Close
frmIRCServer.wsock(0).LocalPort = dB.ReadEntry("Settings", "Port", "6667")
frmIRCServer.wsock(0).Listen
lUsersUBound = dB.ReadEntry("Settings", "MaxUsers", "100") + 4
MaxNickRegs = dB.ReadEntry("Settings", "MaxNickRegs", "100")
MaxChanRegs = dB.ReadEntry("Settings", "MaxChanRegs", "100")
MaxChannels = dB.ReadEntry("Settings", "MaxChannels", "100")
SessionLimit = dB.ReadEntry("Settings", "Session Limit", "3")
Nicklen = dB.ReadEntry("Settings", "MaxNickLength", "25")
MaxJoinChannels = dB.ReadEntry("Settings", "MaxJoinChannels", "7")
TopicLen = dB.ReadEntry("Settings", "TopicLen", "128")
KickLen = dB.ReadEntry("Settings", "KickLen", "64")
Msglen = dB.ReadEntry("Settings", "MsgLen", "512")
MaxChunkSize = dB.ReadEntry("Settings", "MaxChunkSize", "512")
LinkPort = dB.ReadEntry("Settings", "LinkPort", "8000")
frmIRCServer.lblClientSocket = "Clients: " & frmIRCServer.wsock.Count
frmIRCServer.lblServerSocket.Caption = "Link Port: " & LinkPort
frmIRCServer.lblServer = ServerName & " - " & ServerDesc
LogLevel = dB.ReadEntry("Settings", "LogLevel", "3")
LogFile = dB.ReadEntry("Settings", "LogFilename", "server.log")
If InStr(1, LogFile, "\") = 0 Then
    LogFile = App.Path & "\data\logs\server.log"
End If
Open LogFile For Output As #1: Close #1
StatusFile = FS.GetFile(LogFile).ParentFolder.Path & "\status.htm"
If LogLevel = 0 Then FS.DeleteFile (LogFile)
LogFormat = dB.ReadEntry("Settings", "LogFormat", "0")
StatusInterval = dB.ReadEntry("Settings", "StatusInterval", "0")
KillTimer frmIRCServer.hWnd, 0
If Not StatusInterval = 0 Then
    LogStatusHandle = SetTimer(frmIRCServer.hWnd, 0, (StatusInterval * 1000), AddressOf LogStatus)
    If LogStatusHandle = 0 Then LogHTML ServerName, "Unable to start StatusTimer"
End If
ReDim Preserve Channels(MaxChannels)
AdminName = dB.ReadEntry("Admin", "Name", "")
AdminEmail = dB.ReadEntry("Admin", "Email", "")
DefTopic = dB.ReadEntry("Channel Defaults", "Topic", "Unregistered Channel")
DefUserModes = dB.ReadEntry("Default User Settings", "UserModes", "w")
DefQuit = dB.ReadEntry("Default User Settings", "Default Quit Msg", "NexIRC")
For i = 1 To dB.ReadEntry("K-lines", "Count", "0")
    Kline = dB.ReadEntry("K-lines", CStr(i), "")
    Klines.Add Kline, Kline
Next i
Dim OLineCount As Long
OLineCount = dB.ReadEntry("O-Lines", "Count", "0")
For i = 1 To OLineCount
    Olines(i).UserName = dB.ReadEntry("O-Line " & i, "UserName", "")
    Olines(i).Password = dB.ReadEntry("O-Line " & i, "Password", "")
    Olines(i).mask = dB.ReadEntry("O-Line " & i, "Mask", "")
    Olines(i).InUse = True
Next i
SendSvrMsg "Server has rehashed on the request of " & Nick
End Sub

Public Function GetPercent1(Base As Long, Cur As Long) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim X As Long, z As Long, p2 As Long, BaseVal As Long, PercVal As Long, Percent As Long, Max
If Cur = 0 Then
    GetPercent1 = 0
    Exit Function
End If
X = Base
z = Cur
p2 = X / 100
BaseVal = X / p2
PercVal = z / p2
Percent = PercVal / BaseVal * 100
GetPercent1 = Percent
End Function

Public Sub SaveOlines()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long, X As Long
dB.FileName = App.Path & "\data\config\server\server.ini"
For i = 1 To UBound(Olines)
    If Olines(i).InUse Then
        dB.WriteINIEntry "O-Line " & CStr(i), "UserName", Olines(i).UserName
        dB.WriteINIEntry "O-Line " & CStr(i), "Password", Olines(i).Password
        dB.WriteINIEntry "O-Line " & CStr(i), "Mask", Olines(i).mask
        X = X + 1
    End If
Next i
dB.WriteINIEntry "O-Lines", "Count", CStr(X)
End Sub

Public Function GetFreeOLine() As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To UBound(Olines)
    If Not Olines(i).InUse Then
        GetFreeOLine = i
        Exit Function
    End If
Next i
End Function

Public Function HasOline(Nick As String, mask As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long, UIndex As Long
UIndex = NickToObject(Nick).Index
For i = 1 To UBound(Olines)
    If Olines(i).InUse Then
        If Users(UIndex).DNS Like Olines(i).mask Then
            HasOline = True
            Exit Function
        End If
    End If
Next i
End Function

Public Function GetOline(DNS As String) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To UBound(Olines)
    If Olines(i).InUse Then
        If DNS Like Olines(i).mask Then
            GetOline = i
            Exit Function
        End If
    End If
Next i
End Function

Public Sub SendSvrMsg(msg As String, Optional Glob As Boolean = False, Optional Origin As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
If Origin = "" Then Origin = ServerName
For i = 1 To UBound(Users)
    If Not Users(i) Is Nothing Then
        If Users(i).IsMode("s") Or Users(i).IRCOp Then SendNotice "", "• Notice -- " & msg, Origin, , CInt(i), False
    End If
Next i
If Glob Then SendLinks "ServerMsg" & vbLf & ServerName & vbLf & msg
End Sub

Public Sub Wall(msg As String, Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
WallOps msg, Index
End Sub

Public Sub WallOps(msg As String, Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim X As Long
For X = 1 To UBound(Users)
    If Not Users(X) Is Nothing Then
        If Users(X).IsMode("o") Or Users(X).IsMode("w") Then
            SendWsock X, ":" & Users(Index).Nick & "!" & Users(Index).ident & "@" & Users(Index).DNS & " WALLOPS " & msg
        End If
    End If
Next X
End Sub

Public Function ClientsFromIP(IP As String) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To CloneControl.Count
    If CloneControl(i) = IP Then ClientsFromIP = ClientsFromIP + 1
Next i
End Function

Public Function Wait(ByVal TimeToWait As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim EndTime As Long
EndTime = GetTickCount + TimeToWait
Do Until GetTickCount > EndTime
    Sleep 10
    DoEvents
Loop
End Function

Public Function StripTerminator(ByVal strString As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Public Sub StripColorCodes(ByRef msg As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim POS As Long, ColorCode As String
Do While InStr(1, msg, "") <> 0
    ColorCode = Mid(msg, InStr(1, msg, ""), 6)
    If ColorCode Like "##,##" Then
        msg = Replace(msg, ColorCode, "", , 1)
    ElseIf ColorCode Like "#,#*" Then
        msg = Replace(msg, Mid(ColorCode, 1, 4), "", , 1)
    ElseIf ColorCode Like "#,##*" Then
        msg = Replace(msg, Mid(ColorCode, 1, 5), "", , 1)
    ElseIf ColorCode Like "##,#*" Then
        msg = Replace(msg, Mid(ColorCode, 1, 5), "", , 1)
    ElseIf ColorCode Like "##*" Then
        msg = Replace(msg, Mid(ColorCode, 1, 3), "", , 1)
    ElseIf ColorCode Like "#*" Then
        msg = Replace(msg, Mid(ColorCode, 1, 2), "", , 1)
    ElseIf ColorCode Like "*" Then
        msg = Replace(msg, "", "", , 1)
    End If
Loop
End Sub

Public Function CreateGUID() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
    Dim ID(0 To 15) As Byte
    Dim Cnt As Long, GUID As String
    CoCreateGuid ID(0)
        For Cnt = 0 To 15
            CreateGUID = CreateGUID + IIf(ID(Cnt) < 16, "0", "") + Hex$(ID(Cnt))
        Next Cnt
        CreateGUID = Left$(CreateGUID, 8) + "-" + Mid$(CreateGUID, 9, 4) + "-" + Mid$(CreateGUID, 13, 4) + "-" + Mid$(CreateGUID, 17, 4) + "-" + Right$(CreateGUID, 12)
End Function

Public Function IsValidString(ByRef strString As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
    Dim i As Long, strAsc As String
    For i = 1 To Len(strString)
        strAsc = Asc(Mid(strString, i, 1))
        If (((strAsc < 65 Or (strAsc > 90 And strAsc < 97)) Or strAsc > 122) And (Mid(strString, i, 1) <> "_")) Then
            strString = Mid(strString, 1, i - 1)
            Exit Function
        End If
    Next i
    IsValidString = True
End Function

Public Function FN(Number, Optional MaxGroupLength As Long = 3, Optional Delimeter As String = ".")
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long, Num As String
Num = StrReverse(Number)
For i = 1 To Len(Num) Step MaxGroupLength
    FN = FN & Mid(Num, i, MaxGroupLength) & Delimeter
Next i
FN = StrReverse(Mid(FN, 1, Len(FN) - 1))
End Function

Public Sub LogStatus(ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
With FS.CreateTextFile(StatusFile, True)
    .WriteLine "<font face=Tahoma><body text=#FFFFFF bgcolor=#000000>"
    .WriteLine "<p align=center><b>LOG FILE</b></p>"
    .WriteLine "<p>Current Local Users:<i><b> " & lUserCount - 4 & "</b></i><br>"
    .WriteLine "Max Local Users:<i><b> " & lMaxUser & "</b></i></p>"
    .WriteLine "<p>Current Global Users:<i><b> " & CurGlobalUsers & "</b></i><br>"
    .WriteLine "Max Global Users:<i><b> " & MaxGlobalUsers & "</b></i></p>"
    .WriteLine "<p>Current Links:<i><b> " & CurLinkCount & "</b></i></br>"
    .WriteLine "Max Links:<i><b> " & MaxLinkCount & "</b></i></p>"
    .WriteLine "<hr>"
    For i = 2 To frmIRCServer.Link.UBound
        If lSettings.sHandleErrors = True Then On Local Error Resume Next
        If frmIRCServer.Link(i).Tag <> "" Then .WriteLine "<p>Link " & (i - 1) & ":<i><b> " & ServerName & " -- " & frmIRCServer.Link(i).Tag & "</b></i></p>"
    Next i
    .WriteLine "<hr>"
    .WriteLine "<p>Current Channels:<i><b> " & lChanCount & "</b></i><br>"
    .WriteLine "Max Channels:<i><b> " & lChannelsUBound & "</b></i></p>"
    .WriteLine "<p>Traffic:<i><b> " & FN(ServerTraffic, 3, ".") & " bytes" & "</b></i></p>"
    .WriteLine "<p>Client port:<i><b> " & frmIRCServer.wsock(0).LocalPort & "</b></i><br>"
    .WriteLine "Server port:<i><b> " & LinkPort & "</b></i></p></font>"
End With
End Sub

Public Sub CreateServices()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim GFS As clsIRCServer_User
Set GFS = GetFreeSlot
GFS.Nick = "ChanServ"
GFS.ID = "ChanServ@" & ServerName & ""
GFS.DNS = "" & ServerName & ""
GFS.Email = AdminEmail
GFS.IRCOp = True
GFS.Name = "Service"
GFS.SignOn = UnixTime
Set GFS = GetFreeSlot
GFS.Nick = "NickServ"
GFS.ID = "NickServ@" & ServerName & ""
GFS.DNS = "" & ServerName & ""
GFS.Email = AdminEmail
GFS.IRCOp = True
GFS.Name = "Service"
GFS.SignOn = UnixTime
Set GFS = GetFreeSlot
GFS.Nick = "MemoServ"
GFS.ID = "MemoServ@" & ServerName & ""
GFS.DNS = "" & ServerName & ""
GFS.Email = AdminEmail
GFS.IRCOp = True
GFS.Name = "Service"
GFS.SignOn = UnixTime
Set GFS = GetFreeSlot
GFS.Nick = "OperServ"
GFS.ID = "OperServ@" & ServerName & ""
GFS.DNS = "" & ServerName & ""
GFS.Email = AdminEmail
GFS.IRCOp = True
GFS.Name = "Service"
GFS.SignOn = UnixTime
End Sub

Public Sub WriteHeader()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
With FS.OpenTextFile(LogFile, ForAppending, True)
    .WriteLine "<html>"
    .WriteLine "<head>"
    '.WriteLine "<meta http-equiv=REFRESH content=2>"
    .WriteLine "<title>NexIRC Status File</title>"
    .WriteLine "</head>"
    .WriteLine "<body text=#00FF00 bgcolor=#000000>"
    .WriteLine "<p align=center><b>NexIRC LOG FILE</b></p>"
    .WriteLine "<table border=1 cellpadding=0 cellspacing=0 style=border-collapse: collapse bordercolor=#111111 width=100% id=AutoNumber1>"
    .WriteLine "    <tr>"
    .WriteLine "        <td width=20% align=center bgcolor=#C0C0C0><font color=#000000><b>Time</b></font></td>"
    .WriteLine "        <td width=20% align=center bgcolor=#C0C0C0><font color=#000000><b>Originator</b></font></td>"
    .WriteLine "        <td width=60% align=center bgcolor=#C0C0C0><font color=#000000><b>Message</b></font></td>"
    .WriteLine "    </tr>"
End With
End Sub

Public Sub WriteFooter()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
With FS.OpenTextFile(LogFile, ForAppending, True)
    .WriteLine "</table>"
    .WriteLine "</body>"
    .WriteLine "</html>"
End With
KillTimer frmIRCServer.hWnd, 0
End Sub

Public Function FixNickList(NS As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FixNickList = Trim(Replace(NS, "  ", " "))
End Function

Public Sub SendLinks(msg As String, Optional Index As Long, Optional OnlySendToLink)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Debug.Print msg
ServerTraffic = ServerTraffic + Len(msg & vbCrLf)
If LogLevel = 1 Then
    If LogFormat = 0 Then
        LogText "[LINK]<" & Now & " (OUTGOING " & msg & ")> " & msg
    Else
        LogHTML ServerName & "(Link) OUTGOING", msg
    End If
End If
If IsMissing(OnlySendToLink) Then
    Dim i As Long
    For i = 2 To frmIRCServer.Link.UBound
        If Not i = Index Then frmIRCServer.Link(i).SendData msg & vbCrLf
    Next i
Else
    frmIRCServer.Link(OnlySendToLink).SendData msg & vbCrLf
End If
End Sub
