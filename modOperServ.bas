Attribute VB_Name = "mdlOperServ"
Option Explicit
Option Compare Text
Private dB As New clsDatabase

Public Sub ParseOSCmd(Cmd As String, Index As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Not Users(Index).IRCOp And Not (Cmd = "OS stats") Then
    SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
    Exit Sub
End If
Dim msg As String, CMDStr As String, lcmd As Integer, arg1 As String, arg2 As String, cmd2 As String
Dim User As clsIRCServer_User
Set User = Users(Index)
msg = Replace(Cmd, "OS ", "")
If Not InStr(1, msg, " ") <> 0 Then
    CMDStr = msg
Else
    CMDStr = (Mid(msg, 1, InStr(1, msg, " ") - 1))
End If
msg = Replace(msg, CMDStr & " ", "")
Select Case LCase(CMDStr)
    Case "stats"
        Dim i As Long
        lcmd = 1
        SendNotice "", "STATISTICS FOR " & ServerName, "OperServ", , CInt(Index)
        SendNotice "", "Links", "OperServ", , CInt(Index)
        For i = frmIRCServer.Link.LBound To frmIRCServer.Link.UBound
            If lSettings.sHandleErrors = True Then On Local Error Resume Next
            If frmIRCServer.Link(i).Tag <> "" Then SendNotice "", "Link " & (i - 1) & "     " & ServerName & " -- " & frmIRCServer.Link(i).Tag, "OperServ", , CInt(Index)
        Next i
        SendNotice "", "---------------------------------------------------------------", "OperServ", , CInt(Index)
        SendNotice "", "Servertraffic: " & FN(ServerTraffic, 3, ".") & " bytes", "OperServ", , CInt(Index)
        SendNotice "", "Up Since: " & Started, "OperServ", , CInt(Index)
    Case "kill"
        lcmd = 4
        Set User = NickToObject(msg)
        If Not User Is Nothing Then
            SendSvrMsg "Recieved KILL message for " & User.Nick & "!" & User.ident & "@" & User.DNS & " Path: " & Users(Index).Nick & " (OperServ)"
            User.Killed = True
            SendWsock User.Index, ":" & Users(Index).Nick & "!" & Users(Index).ident & "@" & ServerName & " KILL " & User.Nick & " :OperServ", True
            SendWsock User.Index, "ERROR :Closing Link: " & User.Nick & "[" & frmIRCServer.wsock(User.Index).RemoteHostIP & ".] " & ServerName & " (OperServ)", True
            SendQuit User.Index, "Killed by " & Users(Index).Nick & " (OperServ)", True
        End If
    Case "akill"
        lcmd = 5
        Set User = NickToObject(msg)
        If Not User Is Nothing Then
            SendSvrMsg "Recieved AKILL message for " & User.Nick & "!" & User.ident & "@" & User.DNS & " Path: " & Users(Index).Nick & " (OperServ)"
            User.Killed = True
            SendWsock User.Index, ":" & Users(Index).Nick & "!" & Users(Index).ident & "@" & ServerName & " KILL " & User.Nick & " :OperServ", True
            SendWsock User.Index, "ERROR :Closing Link: " & User.Nick & "[" & frmIRCServer.wsock(User.Index).RemoteHostIP & ".] " & ServerName & " (OperServ)", True
            If lSettings.sHandleErrors = True Then On Local Error Resume Next
            Klines.Add frmIRCServer.wsock(User.Index)
            SendQuit User.Index, "AKilled by " & Users(Index).Nick & " (OperServ)", True
        End If
    Case "clear"
        lcmd = 6
        SendMsg "ChanServ", "clear " & msg, Users(Index).Nick, False
    Case "global"
        lcmd = 7
        For i = LBound(Users) To UBound(Users)
            If Not Users(i) Is Nothing Then SendNotice "", "• Global -- " & msg, "OperServ", , CInt(i)
        Next i
    Case "logonnews"
        lcmd = 8
        FS.CreateTextFile(App.Path & "\data\config\server\logon.ini", True).WriteLine msg
    Case "help"
        If msg = "" Or msg = "help" Then
            SendNotice "", "OperServ Commands", "OperServ", , CInt(Index)
            SendNotice "", "STATS (stats)", "OperServ", , CInt(Index)
            SendNotice "", "ADDSTAFF (addstaff [Nick] )", "OperServ", , CInt(Index)
            SendNotice "", "DELSTAFF (delstaff [Nick] )", "OperServ", , CInt(Index)
            SendNotice "", "KILL (kill [nick] [reason] )", "OperServ", , CInt(Index)
            SendNotice "", "AKILL (akill [nick] )", "OperServ", , CInt(Index)
            SendNotice "", "CLEAR (clear [channel] )", "OperServ", , CInt(Index)
            SendNotice "", "GLOBAL (global [message] )", "OperServ", , CInt(Index)
            SendNotice "", "LOGONNEWS (logonnews [news] )", "OperServ", , CInt(Index)
        Else
            Select Case LCase(msg)
                Case "stats"
                    SendNotice User.Nick, "Stats (stats)", "OperServ", , CInt(Index)
                Case "addstaff"
                    SendNotice User.Nick, "AddStaff (addstaff [Nick] )", "OperServ", , CInt(Index)
                Case "delstaff"
                    SendNotice User.Nick, "DelStaff (delstaff [Nick] )", "OperServ", , CInt(Index)
                Case "kill"
                    SendNotice User.Nick, "Kill (kill [nick] <Reason> )", "OperServ", , CInt(Index)
                Case "akill"
                    SendNotice User.Nick, "Akill (akill [nick] <Reason> )", "OperServ", , CInt(Index)
                Case "clear"
                    SendNotice User.Nick, "Clear (clear [channel] )", "OperServ", , CInt(Index)
                Case "global"
                    SendNotice User.Nick, "Global (global [message] )", "OperServ", , CInt(Index)
                Case "logonnews"
                    SendNotice User.Nick, "LogonNews (logonnews [news] )", "OperServ", , CInt(Index)
            End Select
        End If
    Case Else
        SendNotice User.Nick, "Command Unknown", "OperServ"
End Select
Exit Sub
parseerr:
Select Case lcmd
    Case 1
        SendNotice User.Nick, "Identify (identify [password] )", "OperServ"
    Case 2
        SendNotice User.Nick, "Drop (drop [password] )", "OperServ"
    Case 3
        SendNotice User.Nick, "Register (register [password] [email] )", "OperServ"
    Case 4
        SendNotice User.Nick, "Kill (kill [nick] [password] )", "OperServ"
    Case 5
        SendNotice User.Nick, "Info (info [nick] )", "OperServ"
    Case 6
        SendNotice User.Nick, "ChangeInfo (changeinfo [newpass] [newemail] )", "OperServ"
    Case Else
        SendNotice User.Nick, "Unknown Command or missing parameters", "OperServ"
End Select
End Sub

Public Sub SendLogonNews(Nick)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
SendNotice "", "•Logon News -- " & FS.OpenTextFile(App.Path & "\data\config\server\logon.ini").ReadLine, "OperServ", , CInt(Nick)
End Sub
