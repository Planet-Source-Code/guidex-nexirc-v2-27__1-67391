Attribute VB_Name = "mdlIRCInput"
Option Explicit
Public lDNS As New clsDNS

Sub ProcessInput(lData As String, ByVal lTBox As TBox, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lData = RTrim(lData)
Dim x, y As Integer, word() As String, parms As String, retDNS As String
word = Split(lData, " ")
For x = 1 To UBound(word)
    parms = parms & word(x) & Chr(32)
Next x
parms = RTrim(parms)
Select Case UCase(word(0))
Case "REGISTER"
    If lRegInfo.rRegistered = True Then
        DoColor lTBox, "" & Color.Normal & "Already registered"
        DoColorSep lTBox
        Exit Sub
    End If
    frmRegister.Show 0, mdiMain
    frmRegister.txtName.Text = parms
Case "lRaw"
    lForm.tcp.SendData parms & vbCrLf
    DoColor lTBox, "->Server: " & parms
    DoColorSep lTBox
Case "MSG"
    parms = ""
    For x = 2 To UBound(word)
        parms = parms & word(x) & Chr(32)
    Next x
    parms = RTrim(parms)
    lForm.tcp.SendData "PRIVMSG " & word(1) & " :" & parms & vbCrLf
    DoColor lTBox, "-> *" & word(1) & "* " & parms
    DoColorSep lTBox
Case "SERVER"
    Select Case UBound(word)
    Case 1
        ConnectToIRC word(1), "6667", lForm
    Case 2
        ConnectToIRC word(1), word(2), lForm
    End Select
Case "WHOIS"
    lForm.tcp.SendData "WHOIS " & word(1) & vbCrLf
Case "JOIN"
    If lSettings.sAddJoinedChannelsToChannelFolder = True Then
        AddtoChanFolder word(1): DoEvents
        SaveChanFolders
    End If
    If Left(word(1), 1) = "#" Then
        lForm.tcp.SendData "JOIN " & word(1) & vbCrLf
    Else
        lForm.tcp.SendData "JOIN #" & word(1) & vbCrLf
    End If
Case "PART"
    If Left(word(1), 1) = "#" Then
        lForm.tcp.SendData "PART " & word(1) & vbCrLf
    Else
        lForm.tcp.SendData "PART #" & word(1) & vbCrLf
    End If
Case "NICK"
    lForm.tcp.SendData "NICK " & word(1) & vbCrLf
Case "CHAT"
    Dim lIP As String
    For x = 1 To 25
        If mdiMain.CHATx(x).State = sckClosed Or mdiMain.CHATx(x).State = sckError Then
            mdiMain.CHATx(x).Close
            mdiMain.CHATx(x).LocalPort = mdiMain.CHATx(0).LocalPort
            mdiMain.CHATx(x).Listen
            Load ChatWindowx(x)
            ChatWindowx(x).Caption = word(1) & " - " & mdiMain.CHATx(x).LocalPort
            ChatWindowNamex(x) = word(1) & " - " & mdiMain.CHATx(x).LocalPort
            ChatWindowx(x).Show
            Call DoColor(ChatWindowx(x).txtIncoming, "" & Color.Notice & "• Connecting to " & ChatWindowNamex(x) & " (" & mdiMain.CHATx(x).RemoteHostIP & ":" & mdiMain.CHATx(x).RemotePort & ") on port " & mdiMain.CHATx(x).LocalPort)
            lForm.tcp.SendData "NOTICE " & word(1) & " :DCC CHAT (" & lForm.tcp.LocalIP & ")" & vbCrLf
            lIP = IrcGetLongIP(lForm.tcp.LocalIP)
            lForm.tcp.SendData "PRIVMSG " & word(1) & " :DCC CHAT chat " & lIP & " " & mdiMain.CHATx(x).LocalPort & "" & vbCrLf
            Exit For
        End If
    Next x
Case "NAMES"
    If word(1) <> "" Then
        lForm.tcp.SendData "NAMES " & word(1) & vbCrLf
    End If
Case "QUIT"
    lForm.tcp.SendData "QUIT :" & parms & vbCrLf
Case "ME"
    DoColor lTBox, "" & Color.Action & "*" & lSettings.sNickname & " " & parms
    DoColorSep lTBox
    lForm.tcp.SendData "PRIVMSG " & ACTION_CHANNEL & " :ACTION " & parms & "" & vbCrLf
Case "CLEAR"
    lTBox.InitializeAgain
Case "LIST"
    If parms <> "" Then
        lForm.tcp.SendData "LIST " & word(1) & vbCrLf
    Else
        lForm.tcp.SendData "LIST" & vbCrLf
    End If
Case "MOTD"
    lForm.tcp.SendData "MOTD" & vbCrLf
Case "LUSERS"
    lForm.tcp.SendData "LUSERS" & vbCrLf
Case "STATS"
    ShowStats lTBox
Case "ECHO"
    DoColor lTBox, "" & Color.Normal & parms
Case "DNSNAME"
    DoColor lTBox, "" & Color.Action & "• Looking up " & word(1)
    DoColorSep lTBox
    retDNS = DNS.NameToAddress(word(1))
    DoColor lTBox, "" & Color.Action & "• Resolved " & word(1) & " to " & retDNS
    DoColorSep lTBox
Case "DNSIP"
    DoColor lTBox, "" & Color.Action & "• Looking up " & word(1)
    DoColorSep lTBox
    retDNS = DNS.AddressToName(word(1))
    DoColor lTBox, "" & Color.Action & "• Resolved " & word(1) & " to " & retDNS
    DoColorSep lTBox
Case "DNS"
    DoColor lTBox, "" & Color.Action & "• Looking up " & word(1)
    If IsNumeric(Left(word(1), 1)) Then
        retDNS = DNS.AddressToName(word(1))
    Else
        retDNS = DNS.NameToAddress(word(1))
    End If
    DoColor lTBox, "" & Color.Action & "• Resolved " & word(1) & " to " & retDNS & vbCrLf & "•"
Case "PING"
    lForm.tcp.SendData "PRIVMSG " & word(1) & " :" & Chr$(1) & "PING " & Trim(str(DateDiff("s", CVDate("01/01/1970"), Now))) & Chr$(1) & vbCrLf
Case "REFRESHLIST"
    frmChannels.lvwChan.Refresh
Case Else
    DoColor lForm.txtIncoming, "->Server: " & lData
    DoColorSep lForm.txtIncoming
    lForm.tcp.SendData lData & vbCrLf
End Select
End Sub
