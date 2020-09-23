Attribute VB_Name = "mdlIRCStuff"
Option Explicit
Private lMenu As clsFMenu
Private Const lTCPUBound = 32
Private lConnected As Boolean
'Public CHAT_Index As Long
Public lChatWindow(1 To lTCPUBound) As New frmChat
Public lChatWindowName(1 To lTCPUBound) As String
Public lChatWindowx(1 To lTCPUBound) As New frmChat
Public lChatWindowNamex(1 To lTCPUBound) As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Public PingReply As Long
Private Type gNumericCode
    nNum As String
    nServer As String
    nNickName As String
    nParms As String
    nServerText As String
End Type
Private Type gCmdTrigger
    cParms As String
    cUserName As String
    cTarget As String
    cCommand As String
    cJoinChannel As String
    cChanPart As String
    cNickJoin As String
    cNickPart As String
End Type
Private lEvents As gCmdTrigger
Private lRaw As gNumericCode
Public FileIndex As Integer
Public FileListenPort As Integer
Public lMyCurrentModes As String

Public Sub SetConnected(lValue As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lConnected = lValue
End Sub

Public Function ReturnConnected() As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnConnected = lConnected
End Function

Public Function ReturnTCPUBound() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnTCPUBound = lTCPUBound
End Function

Public Function ReturnMaxTCP() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnMaxTCP = lTCPUBound
End Function

Public Sub AddTaskPanel(lCaption As String, lPicType As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If (mdiNexIRC.StatusBar.Panels.Count + 1) = 17 Then Exit Sub
If FindPanelIndex(lCaption, mdiNexIRC.StatusBar) = 0 Then
    mdiNexIRC.StatusBar.Panels.Add (mdiNexIRC.StatusBar.Panels.Count + 1), lCaption, lCaption
End If
'If err.number = 35602 Then Exit Sub
If lSettings.sAutosizeStatusbarItems = True Then
    mdiNexIRC.StatusBar.Panels.Item((mdiNexIRC.StatusBar.Panels.Count)).AutoSize = sbrSpring
Else
    mdiNexIRC.StatusBar.Panels.Item((mdiNexIRC.StatusBar.Panels.Count)).AutoSize = sbrNoAutoSize
End If
mdiNexIRC.StatusBar.Panels.Item((mdiNexIRC.StatusBar.Panels.Count)).Style = sbrText
mdiNexIRC.StatusBar.Panels.Item((mdiNexIRC.StatusBar.Panels.Count)).Bevel = sbrRaised
Select Case lPicType
Case 1
    mdiNexIRC.StatusBar.Panels.Item((mdiNexIRC.StatusBar.Panels.Count)).Picture = mdiNexIRC.imgTaskbar.ListImages(1).Picture
Case 2
    mdiNexIRC.StatusBar.Panels.Item((mdiNexIRC.StatusBar.Panels.Count)).Picture = mdiNexIRC.imgTaskbar.ListImages(2).Picture
End Select
End Sub

Public Sub RemoveTaskbar(lCaption As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To mdiNexIRC.StatusBar.Panels.Count
'    MsgBox mdiNexIRC.StatusBar.Panels.Item(i).Key & " - " & lCaption
    If LCase(mdiNexIRC.StatusBar.Panels.Item(i).Key) = LCase(lCaption) Then
        mdiNexIRC.StatusBar.Panels.Remove i
        Exit For
    End If
Next i
'If Err.Number <> 0 Then ProcessRuntimeError Err.Description, err.number, "Public Sub RemoveTaskbar(lCaption As String)"
End Sub

'Public Sub ShowStats(lTBox As ctlTBox)
'Dim lOperators As Integer, lVoiced As Integer, lUsers As Integer, i As Integer, j As Integer
'For i = 1 To ReturnChannelUBound
'    If LCase(ReturnActChannel) = LCase(ReturnChannelName(i)) Then
'        For j = 1 To ReturnChannelNamesCount(i)
'            Select Case Left(ReturnChannelWindowNamesColor(i, j), 1)
'            Case lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nOpColor
'                lOperators = lOperators + 1
'            Case lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nVoiceColor
'                lVoiced = lVoiced + 1
'            Case Else
'                lUsers = lUsers + 1
'            End Select
'        Next j
'        DoColor ReturnChannelIncomingTBox(i), "" & Color.Join & "• Operators: " & lOperators & " lVoiced: " & lVoiced & " lUsers: " & lUsers & " - Total: " & Trim(Str(lOperators + lVoiced + lUsers))
'        Exit For
'    End If
'Next i
'If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ShowStats(RTF as ctlTBox)"
'End Sub

Public Sub ClearVariables()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lRaw.nNickName = ""
lRaw.nNum = ""
lRaw.nParms = ""
lRaw.nServer = ""
lRaw.nServerText = ""
lEvents.cUserName = ""
lEvents.cTarget = ""
lEvents.cParms = ""
lEvents.cNickPart = ""
lEvents.cNickJoin = ""
lEvents.cCommand = ""
lEvents.cChanPart = ""
lEvents.cJoinChannel = ""
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ClearVariables()"
End Sub

Public Sub ParseIRCData(lData As String, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, j As Integer, Y As Integer, lWordArr() As String, lErr As String, lParams As String
lRaw.nServerText = lData
If lData = "" Then Exit Sub
lErr = lData
lWordArr = Split(lData, Chr(32))
For i = 3 To UBound(lWordArr)
    If Trim(lWordArr(i)) <> "" Then
        lParams = lParams & " " & lWordArr(i)
    End If
Next i
If IsNumeric(lWordArr(1)) = True Then
    lSettings.sNickname = lWordArr(2)
    Call numeric(lWordArr(0), lWordArr(1), lWordArr(2), Mid(lParams, 2), lForm)
    Exit Sub
End If
If lWordArr(0) = "PING" Then
    lForm.tcp.SendData "PONG " & Mid(lWordArr(1), 2) & vbCrLf
    lForm.Caption = lForm.Tag & " " & lSettings.sNickname & " on " & lSettings.sServer & ":" & lForm.tcp.RemotePort
    Exit Sub
End If
Call Command(lWordArr(0), lWordArr(1), lWordArr(2), Mid(lParams, 2), lForm)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ParseData(lData As String, lForm As Form)"
End Sub

Public Sub IsUserOnline(lTBox As ctlTBox, lNickName As String)
         If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, j As Integer
For i = 1 To ReturnChannelUBound
    If ReturnChannelName(i) <> "" Then
        For j = 1 To ReturnChannelNamesCount(i) - 1
            If LCase(ReturnChannelNames(i, j)) = LCase(lNickName) Or LCase(ReturnChannelNames(i, j)) = LCase("@" & lNickName) Or LCase(ReturnChannelNames(i, j)) = LCase("+" & lNickName) Then
                ProcessReplaceString sUserOnline, lTBox, lNickName, ReturnChannelName(i)
                Exit For
            End If
        Next j
    End If
Next i
End Sub

Public Sub ProcessInput(lData As String, ByVal lTBox As ctlTBox, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lData = RTrim(lData)
Dim X, Y As Integer, word() As String, parms As String, retDNS As String
word = Split(lData, " ")
For X = 1 To UBound(word)
    parms = parms & word(X) & Chr(32)
Next X
parms = RTrim(parms)
Select Case UCase(word(0))
'Case "REGISTER"
    'If lRegInfo.rRegistered = True Then
    '    ProcessReplaceString sAlreadyRegistered, lTBox
'        DoColorSep lTBox
'        Exit Sub
'    End If
'    frmRegister.Show 0, mdiNexIRC
'    frmRegister.txtName.Text = parms
Case "Raw"
    lForm.tcp.SendData parms & vbCrLf
    ProcessReplaceString sRaw, lTBox, parms
    DoColorSep lTBox
Case "MSG"
    parms = ""
    For X = 2 To UBound(word)
        parms = parms & word(X) & Chr(32)
    Next X
    parms = RTrim(parms)
    lForm.tcp.SendData "PRIVMSG " & word(1) & " :" & parms & vbCrLf
    ProcessReplaceString sOwnMessage, lTBox, word(1), parms
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
On Local Error Resume Next
    For X = 1 To 25
        
        If mdiNexIRC.wskChat2(X).State = sckClosed Or mdiNexIRC.wskChat2(X).State = sckError Then
            mdiNexIRC.wskChat2(X).Close
            mdiNexIRC.wskChat2(X).LocalPort = mdiNexIRC.wskChat2(0).LocalPort
            mdiNexIRC.wskChat2(X).Listen
            Load lChatWindowx(X)
            lChatWindowx(X).Caption = word(1) & " - " & mdiNexIRC.wskChat2(X).LocalPort
            lChatWindowNamex(X) = word(1) & " - " & mdiNexIRC.wskChat2(X).LocalPort
            lChatWindowx(X).Show
            ProcessReplaceString sInitiateDCCChat, lChatWindowx(X).txtIncoming, lChatWindowNamex(X), mdiNexIRC.wskChat2(X).RemoteHostIP, mdiNexIRC.wskChat2(X).RemotePort, mdiNexIRC.wskChat2(X).LocalPort
            lForm.tcp.SendData "NOTICE " & word(1) & " :DCC CHAT (" & lForm.tcp.LocalIP & ")" & vbCrLf
            lIP = IrcGetLongIP(lForm.tcp.LocalIP)
            lForm.tcp.SendData "PRIVMSG " & word(1) & " :DCC CHAT chat " & lIP & " " & mdiNexIRC.wskChat2(X).LocalPort & "" & vbCrLf
            Exit For
        End If
    Next X
Case "NAMES"
    If word(1) <> "" Then
        lForm.tcp.SendData "NAMES " & word(1) & vbCrLf
    End If
Case "QUIT"
    lForm.tcp.SendData "QUIT :" & parms & vbCrLf
Case "ME"
    ProcessReplaceString sAction, lTBox, lSettings.sNickname, parms
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
'    ShowStats lTBox
Case "ECHO"
    DoColor lTBox, "" & Color.Normal & parms
Case "DNSNAME"
    DoColor lTBox, "" & Color.Action & "• Looking up " & word(1)
    DoColorSep lTBox
    retDNS = NameToAddress(word(1))
    DoColor lTBox, "" & Color.Action & "• Resolved " & word(1) & " to " & retDNS
    DoColorSep lTBox
Case "DNSIP"
    DoColor lTBox, "" & Color.Action & "• Looking up " & word(1)
    DoColorSep lTBox
    retDNS = AddressToName(word(1))
    DoColor lTBox, "" & Color.Action & "• Resolved " & word(1) & " to " & retDNS
    DoColorSep lTBox
Case "lDNS"
    DoColor lTBox, "" & Color.Action & "• Looking up " & word(1)
    If IsNumeric(Left(word(1), 1)) Then
        retDNS = AddressToName(word(1))
    Else
        retDNS = NameToAddress(word(1))
    End If
    DoColor lTBox, "" & Color.Action & "• Resolved " & word(1) & " to " & retDNS & vbCrLf & "•"
Case "PING"
    lForm.tcp.SendData "PRIVMSG " & word(1) & " :" & Chr$(1) & "PING " & Trim(Str(DateDiff("s", CVDate("01/01/1970"), Now))) & Chr$(1) & vbCrLf
Case "REFRESHLIST"
    frmChannels.lvwChannels.Refresh
Case Else
    DoColor lForm.txtIncoming, "->Server: " & lData
    DoColorSep lForm.txtIncoming
    lForm.tcp.SendData lData & vbCrLf
End Select
End Sub

Public Sub Command(lUserName As String, lCommand As String, lTarget As String, lParameters As String, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, X As Integer, b As Boolean, s As Boolean, F As Integer, msg As String, msg2 As String, msg3 As String, lEmail As String
lEvents.cParms = lParameters
lEvents.cUserName = lUserName
lEvents.cTarget = lTarget
lEvents.cCommand = lCommand
msg3 = Now
If Left(lUserName, 1) = ":" Then lUserName = Right(lUserName, Len(lUserName) - 1)
lParameters = LTrim(lParameters)
If Left(lParameters, 1) = ":" Then
    lParameters = Mid(lParameters, 2)
End If
If Left(lTarget, 1) = ":" Then
    lTarget = Mid(lTarget, 2)
End If
s = True
For i = 1 To Len(lUserName)
    If Mid(lUserName, i, 1) = "!" Then
        lEmail = Mid(lUserName, i + 1)
        s = False
        lUserName = Mid(lUserName, 1, i - 1)
        If Left(lUserName, 1) = ":" Then
            lUserName = Mid(lUserName, 2)
        End If
    End If
Next i
Dim strCText As String
Select Case UCase(lCommand)
Case "JOIN"
    lEvents.cNickJoin = lUserName
    lEvents.cJoinChannel = lTarget
    For i = 1 To ReturnChannelUBound
'        MsgBox i & ReturnChannelName(i)
        If ReturnChannelName(i) = "" Then
            If LCase(lUserName) = LCase(lSettings.sNickname) Then
                LoadChannel i, lTarget
                lForm.tcp.SendData "MODE " & lTarget & vbCrLf
                Exit For
            End If
        End If
        If LCase(ReturnChannelName(i)) = LCase(lTarget) Then
            
            'MsgBox "Add user '" & lUserName & "' to '" & ReturnChannelName(i)
            AddUserToNicklist lUserName, ReturnChannelNamesListView(i)
            UpdateChannelCaption i
            strCText = ReturnReplacedString(sJoin, lUserName, lEmail, lTarget)
            strCText = Replace(strCText, "$time", msg3)
            If lSettings.sOptions.oShowJoinPart = True Then
                If lSettings.sOptions.oShowAddress = True Then
                    Call DoColor(ReturnChannelIncomingTBox(i), "" & Color.Join & strCText)
                Else
                    Call DoColor(ReturnChannelIncomingTBox(i), "" & Color.Join & strCText)
                End If
            Else
                If lSettings.sOptions.oShowAddress = True Then
                    DoColor lForm.txtIncoming, "" & Color.Join & strCText
                    DoColorSep lForm.txtIncoming
                Else
                    DoColor lForm.txtIncoming, "" & Color.Join & strCText
                    DoColorSep lForm.txtIncoming
                End If
                Exit For
            End If
            If IsInBlacklist(lUserName) = True Then
                Pause 0.2
                lForm.tcp.SendData "KICK " & lTarget & " " & lUserName & vbCrLf
                lForm.tcp.SendData "NAMES " & lTarget
            End If
        End If
    Next i
    LoadScript "nexirc\onjoin.txt"
Case "PART"
    lEvents.cChanPart = lTarget
    lEvents.cNickPart = lUserName
    For i = 1 To ReturnChannelUBound
        If ReturnChannelName(i) = LCase(lTarget) Then
            If LCase(lUserName) = LCase(lSettings.sNickname) Then
                UnloadChannel i
                SetChannelName i, ""
                Exit For
            End If
        End If
        If LCase(ReturnChannelName(i)) = LCase(lTarget) Then
            For X = 1 To ReturnChannelNamesCount(i) - 1
                If LCase(ReturnChannelListItemName(i, X)) = LCase(lUserName) Or LCase(ReturnChannelListItemName(i, X)) = LCase("@" & lUserName) Or LCase(ReturnChannelListItemName(i, X)) = LCase("+" & lUserName) Then
                    RemoveChannelName i, X
                    UpdateChannelCaption i
                    strCText = ReturnReplacedString(sPart, lUserName, lEmail, lTarget)
                    strCText = Replace(strCText, "$time", msg3)
                    If lSettings.sOptions.oShowJoinPart = True Then
                        If lSettings.sOptions.oShowAddress = True Then
                            Call DoColor(ReturnChannelIncomingTBox(i), "" & Color.Part & strCText)
                        Else
                            Call DoColor(ReturnChannelIncomingTBox(i), "" & Color.Part & strCText)
                        End If
                    Else
                        If lSettings.sOptions.oShowAddress = True Then
                            DoColor lForm.txtIncoming, "" & Color.Part & strCText
                            DoColorSep lForm.txtIncoming
                        Else
                            DoColor lForm.txtIncoming, "" & Color.Part & strCText
                            DoColorSep lForm.txtIncoming
                        End If
                        Exit For
                    End If
                End If
            Next X
        End If
    Next i
    LoadScript "nexirc\onpart.txt"
Case "PRIVMSG"
    If Left(lTarget, 1) = "#" Then
        For i = 1 To ReturnChannelUBound
            If ReturnChannelName(i) = LCase(lTarget) Then
                If LCase(Left(lParameters, 7)) = LCase("ACTION") Then
                    If Right(lParameters, 1) = "" Then
                        lParameters = Mid(lParameters, 1, Len(lParameters) - 1)
                    End If
                    Call DoColor(ReturnChannelIncomingTBox(i), "" & Color.Action & "• " & lUserName & " " & Mid(lParameters, 8))
                Else
                    If CheckIgnoreList(lUserName, lForm) = False Then
                        ProcessReplaceString sPm, ReturnChannelIncomingTBox(i), lUserName, lParameters
                        If InStr(lParameters, "!list") Then
                            If lSettings.sEnableList = True Then SendUserPlaylist lUserName, lForm
                        End If
                        If InStr(lParameters, "!") Then
                            If lSettings.sAudioServer = True Then
                                msg = lParameters
                                If Len(msg) <> 0 Then
                                    If lSettings.sAudioServer = True Then
                                        msg2 = Trim(Parse(LCase(msg), "!", ".mp3") & ".mp3")
                                        If Len(msg2) > 4 Then
                                            i = FindFileIndexByFilename(msg2)
                                            If i <> 0 Then
                                                If Trim(LCase(lPlayback.pCurrentFile)) = Trim(LCase(msg2)) Then
                                                    MenuStop
                                                    DoEvents
                                                End If
                                                If i <> 0 Then
                                                    mdiNexIRC.tmrSendUserPlaylist.Enabled = False
                                                    mdiNexIRC.ActiveateDCCSend lFiles.fFile(i).fFilename, lUserName
                                                End If
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
            End If
        Next i
    Else
        If LCase(Left(lParameters, 8)) = LCase("VERSION") Then
            Call CTCP("VERSION SENT", lForm.txtIncoming, lUserName, lForm)
            GoTo DONE
        End If
        b = False
        For i = 1 To 150
            If LCase(lUserName) = LCase(ReturnQueryName(i)) Then
                b = True
                Exit For
            End If
        Next i
        If b = False Then
            If CheckIgnoreList(lUserName, lForm) = False Then
                If lSettings.sSecureQuery = True Then
                    If InStr(LCase(lParameters), "[secure query]") Then
                        Exit Sub
                    End If
                    If IsUserInNotifyList(lUserName) = False Then
                        lForm.tcp.SendData "PRIVMSG " & lUserName & " :[Secure Query] Prompting User" & vbCrLf
                        frmSecureQuery.lblNickname.Caption = lUserName & "?"
                        frmSecureQuery.Show 1, mdiNexIRC
                    Else
                        lSecureQuery.sAccepted = True
                        lSecureQuery.sAddToIgnore = False
                        lSecureQuery.sAddToNotify = False
                    End If
                    If lSecureQuery.sAddToIgnore = True Then
                        lForm.tcp.SendData "PRIVMSG " & lUserName & " :[Secure Query] You have been placed on ignore" & vbCrLf
                        AddToIgnore lUserName
                    End If
                    If lSecureQuery.sAddToNotify = True Then
                        lForm.tcp.SendData "PRIVMSG " & lUserName & " :[Secure Query] " & lSettings.sNickname & " has added you to their notify list" & vbCrLf
                        AddNotify lUserName
                    End If
                    If lSecureQuery.sAccepted = False Then
                        lForm.tcp.SendData "PRIVMSG " & lUserName & " :[Secure Query] Query Declined" & vbCrLf
                        mdiNexIRC.SetFocus
                        Exit Sub
                    Else
                        lForm.tcp.SendData "PRIVMSG " & lUserName & " :[Secure Query] Query Accepted" & vbCrLf
                        mdiNexIRC.SetFocus
                    End If
                End If
                For i = 1 To 150
                    If ReturnQueryName(i) = "" Then
                        LoadQueryWindow i, lUserName, ""
                        Exit For
                    End If
                Next i
            End If
        End If
        If CheckIgnoreList(lUserName, lForm, True) = False Then
            For i = 1 To 150
                If Len(ReturnQueryName(i)) <> 0 Then
                    If LCase(ReturnQueryName(i)) = LCase(lUserName) Then
                        If Left(lParameters, 1) = "" Then
                            Call CTCP(lParameters, ReturnQueryIncomingTBox(i), lUserName, lForm)
                            SetQueryCaption i, lUserName & " [" & lEmail & "]"
                        Else
                            Call DoColor(ReturnQueryIncomingTBox(i), "" & Color.Normal & "<" & Color.Whois & "" & lUserName & "" & Color.Normal & "> " & lParameters)
                            SetQueryCaption i, lUserName & " [" & lEmail & "]"
                        End If
                    End If
                End If
            Next i
        End If
    End If
    LoadScript "nexirc\onmsg.txt"
Case "NICK"
    If LCase(lUserName) = LCase(lSettings.sNickname) Then
        lSettings.sNickname = lTarget
        lForm.Caption = lForm.Tag & ": " & lTarget & " on " & lSettings.sServer
    End If
    For i = 1 To ReturnChannelUBound
        If ReturnChannelName(i) <> "" Then
            For X = 1 To ReturnChannelNamesCount(i) - 1
                If LCase(ReturnChannelListItemName(i, X)) = LCase(lUserName) Then
                    RemoveChannelName i, X
                    'MsgBox "Add user '" & lUserName & "' to '" & ReturnChannelName(i)
                    AddUserToNicklist lTarget, ReturnChannelNamesListView(i)
                    F = FindListViewIndex(ReturnChannelNamesListView(i), lTarget)
                    SetChannelWindowNamesForeColor i, F, lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBottomBandsColor
                    DoColor ReturnChannelIncomingTBox(i), "" & Color.Nick & "• " & lUserName & " is now known as " & lTarget
                Else
                    If LCase(ReturnChannelListItemName(i, X)) = LCase("@" & lUserName) Then
                        RemoveChannelName i, X
                        'MsgBox "Add user '" & lUserName & "' to '" & ReturnChannelName(i)
                        AddUserToNicklist "@" & lTarget, ReturnChannelNamesListView(i)
                        F = FindListViewIndex(ReturnChannelNamesListView(i), "@" & lTarget)
                        SetChannelWindowNamesForeColor i, F, lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sTopBandsColor
                        DoColor ReturnChannelIncomingTBox(i), "" & Color.Nick & "• " & lUserName & " is now known as " & lTarget
                    Else
                        If LCase(ReturnChannelVisible(i, X)) = LCase("+" & lUserName) Then
                            RemoveChannelName i, X
                            'MsgBox "Add user '" & lUserName & "' to '" & ReturnChannelName(i)
                            AddUserToNicklist "+" & lTarget, ReturnChannelNamesListView(i)
                            F = FindListViewIndex(ReturnChannelNamesListView(i), "+" & lTarget)
                            SetChannelWindowNamesForeColor i, F, lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sPeaksColor
                            DoColor ReturnChannelIncomingTBox(i), "" & Color.Nick & "• " & lUserName & " is now known as " & lTarget
                        End If
                    End If
                End If
            Next X
        End If
        For X = 1 To 150
            If LCase(ReturnQueryName(X)) = LCase(lUserName) Then
                SetQueryName X, lTarget
                Call RemoveTaskbar(lUserName)
                Call AddTaskPanel(lTarget, 1)
                SetQueryCaption X, lTarget & " [" & lEmail & "]"
            End If
            Exit For
        Next X
    Next i
    LoadScript "nexirc\onnick.txt"
Case "NOTICE"
    If s Then
        Call DoColor(lForm.txtIncoming, "" & Color.Notice & "NOTICE: " & lParameters)
    End If
    If LCase(Left(lParameters, 1)) = LCase("") Then
        Call CTCP(lParameters, lForm.txtIncoming, lUserName, lForm)
        GoTo DONE
    End If
    If Left(lTarget, 1) = "#" Then
        For i = 1 To ReturnChannelUBound
            If LCase(lTarget) = LCase(ReturnChannelName(i)) Then DoColor ReturnChannelIncomingTBox(i), "" & Color.Notice & lTarget & ": <" & lUserName & "> " & lParameters
        Next i
    Else
        If s = False Then
            b = False
            If Left(LCase(mdiNexIRC.ActiveForm.Caption), 6) = LCase("Status") Then
                DoColor lForm.txtIncoming, "" & Color.Notice & lUserName & ": " & lParameters
                DoColorSep lForm.txtIncoming
            Else
                If Left(mdiNexIRC.ActiveForm.Caption, 1) = "#" Then
                    DoColor mdiNexIRC.ActiveForm.txtIncoming, "" & Color.Notice & lUserName & ": " & lParameters
                    DoColorSep mdiNexIRC.ActiveForm.txtIncoming
                Else
                    DoColor lForm.txtIncoming, "" & Color.Notice & lUserName & ": " & lParameters
                End If
            End If
        End If
    End If
    LoadScript "nexirc\onnotice.txt"
Case "QUIT"
    For i = 1 To ReturnChannelUBound
        If ReturnChannelName(i) <> "" Then
            For X = 1 To ReturnChannelNamesCount(i) - 1
                If LCase(ReturnChannelNames(i, X)) = LCase(lUserName) Or LCase(ReturnChannelNames(i, X)) = LCase("@" & lUserName) Or LCase(ReturnChannelNames(i, X)) = LCase("+" & lUserName) Then
                    RemoveChannelName i, X
                    UpdateChannelCaption i
                    strCText = ReturnReplacedString(sQuit, lUserName, lEmail, lParameters)
                    strCText = Replace(strCText, "$time", msg3)
                    If lSettings.sOptions.oShowQuit = True Then
                        Call DoColor(ReturnChannelIncomingTBox(i), "" & Color.Quit & strCText)
                    Else
                        DoColor lForm.txtIncoming, "" & Color.Quit & strCText
                        DoColorSep lForm.txtIncoming
                        Exit For
                    End If
                End If
            Next X
        End If
    Next i
    LoadScript "nexirc\onquit.txt"
Case "KICK"
    Dim word(1) As String
    i = InStr(2, lParameters, Chr(32))
    word(1) = Trim(Mid(lParameters, 1, i - 1))
    lParameters = Mid(lParameters, 2 + i)
    For i = 1 To ReturnChannelUBound
        If LCase(lTarget) = LCase(ReturnChannelName(i)) Then
            strCText = ReturnReplacedString(sKick, lUserName, word(1), lTarget, lParameters)
            strCText = Replace(strCText, "$time", msg3)
            If lSettings.sOptions.oShowKicks = True Then
                Call DoColor(ReturnChannelIncomingTBox(i), "" & Color.Kick & strCText)
            Else
                DoColor lForm.txtIncoming, "" & Color.Kick & strCText
                DoColorSep lForm.txtIncoming
            End If
            For X = 1 To ReturnChannelNamesCount(i) - 1
                If LCase(ReturnChannelNames(i, X)) = LCase(word(1)) Or LCase(ReturnChannelNames(i, X)) = LCase("@" & word(1)) Or LCase(ReturnChannelNames(i, X)) = LCase("+" & word(1)) Then
                    RemoveChannelName i, X
                    UpdateChannelCaption i
                End If
            Next X
            If lSettings.sOptions.oReJoin = True Then
                'lForm.tcp.SendData "JOIN " & lChannelName(i) & vbCrLf
                lForm.tcp.SendData "JOIN " & ReturnChannelName(i) & vbCrLf
            End If
        End If
    Next i
    LoadScript "nexirc\onkick.txt"
Case "MODE"
    LoadScript "nexirc\onmode.txt"
    Dim strWord() As String
    strWord = Split(lParameters, " ")
    If Len(strWord(0)) > 2 Then
        Dim strMode(1 To 4) As String
        Dim CurrentMode As String
        X = 1
        strWord(0) = Trim(strWord(0))
        For i = 1 To Len(strWord(0))
            If strWord(0) = "" Then Exit For
            If Left(strWord(0), 1) = "+" Or Left(strWord(0), 1) = "-" Then
                strMode(X) = Mid(strWord(0), 1, 2)
                CurrentMode = Mid(strWord(0), 1, 1)
                strWord(0) = Mid(strWord(0), 3)
            Else
                strMode(X) = CurrentMode & Mid(strWord(0), 1, 1)
                strWord(0) = Mid(strWord(0), 2)
            End If
            X = X + 1
            i = 1
        Next i
    Else
        Select Case LCase(Mid(strWord(0), 2))
            Case "o"
                Call SetOp(Left(strWord(0), 1), lUserName, lTarget, strWord(1), lForm)
            Case "v"
                Call SetVoice(Left(strWord(0), 1), lUserName, lTarget, strWord(1), lForm)
            Case "i"
                Call SetInvisible(Left(strWord(0), 1), lUserName, lForm)
            Case "r"
                Call Register(Left(strWord(0), 1), lUserName, lForm)
            Case "b"
                Call Ban(Left(strWord(0), 1), lUserName, lTarget, strWord(1), lForm)
            Case "l"
                Call Limit(Left(strWord(0), 1), lUserName, lTarget, strWord(1))
            Case "n"
                'If lSettings.sGeneralPrompts = True Then
                    'MsgBox Left(strWord(0), 1) & " " & lUserName & " " & lTarget & " " & strWord(1)
                'End If
            Case "t"
                'If lSettings.sGeneralPrompts = True Then
                    'MsgBox Left(strWord(0), 1) & " " & lUserName & " " & lTarget & " " & strWord(1)
                'End If
            Case Else
                DoColor lForm.txtIncoming, "" & Color.Mode & "• " & Replace(lUserName, ":", "") & " sets mode: " & lParameters
                DoColorSep lForm.txtIncoming
                If Left(strWord(0), 1) = "+" Then
                    lMyCurrentModes = lMyCurrentModes & Mid(strWord(0), 2)
                    lForm.Caption = lForm.Tag & ": [" & lMyCurrentModes & "] " & lSettings.sNickname & " on " & lSettings.sServer & ":" & lForm.tcp.RemotePort
                Else
                    lMyCurrentModes = Replace(lMyCurrentModes, Mid(strWord(0), 2), "")
                    lForm.Caption = lForm.Tag & ": [" & lMyCurrentModes & "] " & lSettings.sNickname & " on " & lSettings.sServer & ":" & lForm.tcp.RemotePort
                End If
            End Select
        End If
        For i = 1 To 4
            If strMode(i) <> "" Then
                Select Case LCase(Mid(strMode(i), 2))
                Case "o"
                    MsgBox "OP: " & strWord(i)
                    Call SetOp(Left(strMode(i), 1), lUserName, lTarget, strWord(i), lForm)
                Case "v"
                    MsgBox "Voice: " & strWord(i)
                    Call SetVoice(Left(strMode(i), 1), lUserName, lTarget, strWord(i), lForm)
                Case "i"
                    MsgBox "Invisible: " & strWord(i)
                    Call SetInvisible(Left(strMode(i), 1), lUserName, lForm)
                Case "b"
                    MsgBox "Ban: " & strWord(i)
                    Call Ban(Left(strMode(i), 1), lUserName, lTarget, strWord(i), lForm)
                Case "r"
                    Call Register(Left(strMode(i), 1), lUserName, lForm)
                Case Else
                    DoColor lForm.txtIncoming, "" & Color.Mode & "• " & Replace(lUserName, ":", "") & " sets mode: " & lParameters
                    DoColorSep lForm.txtIncoming
                    If Left(strMode(i), 1) = "+" Then
                        lMyCurrentModes = lMyCurrentModes & Mid(strMode(i), 2)
                        lForm.Caption = lForm.Tag & ": [" & lMyCurrentModes & "] " & lSettings.sNickname & " on " & lSettings.sServer & ":" & lForm.tcp.RemotePort
                    Else
                        lMyCurrentModes = Replace(lMyCurrentModes, Mid(strMode(i), 2), "")
                        lForm.Caption = lForm.Tag & ": [" & lMyCurrentModes & "] " & lSettings.sNickname & " on " & lSettings.sServer & ":" & lForm.tcp.RemotePort
                    End If
                    For X = 1 To ReturnChannelUBound
                        If LCase(ReturnChannelName(i)) = LCase(lTarget) Then
                        'If LCase(lChannelName(i)) = LCase(lTarget) Then
                            If Left(strMode(i), 1) = "-" Then
                            Else
                            End If
                            Exit For
                        End If
                    Next X
                End Select
            End If
        Next i
    Case "TOPIC"
        Call ChangeTopic(lUserName, lTarget, lParameters)
    Case "INVITE"
        DoColor lForm.txtIncoming, "" & Color.Invite & "• " & lUserName & " invites you to join " & lParameters
        DoColorSep lForm.txtIncoming
        If lSettings.sAutoJoinOnInvite = True Then lForm.tcp.SendData "JOIN " & lParameters & vbCrLf
    Case "ERROR:"
        DoColor lForm.txtIncoming, "" & Color.Kick & "• Error: " & lParameters
    Case "AUTH"
        DoColor lForm.txtIncoming, "" & Color.Notice & "• " & lParameters
        DoColorSep lForm.txtIncoming
    Case ":CLOSING"
        DoColor lForm.txtIncoming, "" & Color.Quit & "• " & lUserName & " " & lCommand & " " & lTarget & " " & lParameters
        DoColorSep lForm.txtIncoming
    Case "DLINE"
        DoColor lForm.txtIncoming, "" & Color.Quit & "• " & lUserName & " " & lCommand & " " & lTarget & " " & lParameters
        DoColorSep lForm.txtIncoming
    Case Else
        DoColor lForm.txtIncoming, "" & Color.Normal & "[04Username" & "" & Color.Normal & "]: " & lUserName & vbCrLf & "" & Color.Normal & "[04Command" & "" & Color.Normal & "]: " & lCommand & vbCrLf & "" & Color.Normal & "[04Target" & "" & Color.Normal & "]: " & lTarget & vbCrLf & "" & Color.Normal & "[04Paramters" & Color.Normal & "]: " & lParameters & vbCrLf
        DoColorSep lForm.txtIncoming
    End Select
DONE:
Exit Sub
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub Command(lUserName As String, lCommand As String, lTarget As String, lParameters As String, lForm As Form)"
End Sub

Sub CTCP(parms As String, RTF As ctlTBox, UserName As String, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim word() As String, X As Integer, Y As Integer, i As Integer, msg As String, lIP As String
parms = Mid(parms, 2)
If Right(parms, 1) = "" Then
    DoColor RTF, "" & Color.CTCP & "[" & UserName & " " & Mid(parms, 1, Len(parms) - 1) & "]"
    DoColorSep RTF
Else
    DoColor RTF, "" & Color.CTCP & "[" & UserName & " " & parms & "]"
    DoColorSep RTF
End If
word = Split(parms, Chr(32))
Select Case UCase(word(0))
Case "ACTION"
    parms = ""
    For X = 1 To UBound(word)
        parms = parms & word(X) & Chr(32)
    Next X
    Call DoColor(RTF, "" & Color.Action & "• " & UserName & " " & parms)
Case "DCC"
    Dim wDCC As New frmDCC_Accept, NewFile As frmDCCFILE
    If UCase(word(1)) = UCase("CHAT") Then
        If UCase(word(2)) = UCase("CHAT") Then
            DoColor RTF, "" & Color.CTCP & "• DCC Chat request from " & UserName
            DoColorSep RTF
            lIP = IrcGetIP(word(3))
            wDCC.Show 0, mdiNexIRC
            wDCC.lblNickname.Caption = UserName
            wDCC.lblIP.Caption = lIP
            wDCC.lblPort.Caption = Left(word(4), Len(word(4)) - 1)
        End If
    End If
    If UCase(word(1)) = UCase("SEND") Then
        lIP = IrcGetIP(Val(word(UBound(word) - 2)))
        Set NewFile = New frmDCCFILE
        FileIndex = FileIndex + 1
        If FileIndex > 999 Then FileIndex = 3
        NewFile.Tag = FreeFile
        Load NewFile.FILE(NewFile.Tag)
        NewFile.Show 0, mdiNexIRC
        For i = 2 To UBound(word)
            If IsNumeric(word(i)) = False Then
                word(2) = word(2) & word(i) & "_"
            Else
                Exit For
            End If
        Next i
        word(2) = Mid(word(2), 1, Len(word(2)) - 1)
        word(2) = Replace(word(2), """", "")
        word(2) = Left(word(2), Len(word(2)) - (Len(word(2)) / 2))
        NewFile.lblFile.Caption = word(2)
        NewFile.lblAddress.Caption = lIP & ":" & word(i + 1)
        NewFile.lblFileSize.Caption = Left(word(i + 2), Len(word(i + 2)) - 1)
        NewFile.lblNickname.Caption = UserName
        NewFile.ProgressBar.Min = 0
        NewFile.ProgressBar.Value = 0
        NewFile.picComplete.BackColor = vbWhite
        NewFile.ProgressBar.Max = Val(word(UBound(word)))
        NewFile.lblFilename.Caption = App.Path & "\data\downloads\" & word(2)
    End If
Case "VERSION"
    'lForm.tcp.SendData "NOTICE " & UserName & " :VERSION NexIRC v" & App.major & "." & App.minor & " by Team Nexgen. Programming by |guideX|" & vbCrLf
    lForm.tcp.SendData "NOTICE " & UserName & ":VERSION " & ReturnReplacedString(sVersion, App.Major, App.Minor)
    DoColor lSettings.sActiveServerForm.txtIncoming, "" & Color.Normal & "Version sent to " & UserName
Case "PING"
    Dim PingTime As String
    For i = 1 To Len(word(1))
        If IsNumeric(Mid(word(1), i, 1)) Then
            PingTime = PingTime & Mid(word(1), i, 1)
        Else
            Exit For
        End If
    Next i
    word(1) = PingTime
    If Right(word(1), 1) = Chr(1) Then word(1) = Mid(word(1), 1, Len(word(1)) - 1)
    X = Val(Trim(Str(DateDiff("s", CVDate("01/01/1970"), Now)))) - Val(Int(word(1)))
    i = Int(X)
    DoColor RTF, "" & Color.CTCP & "[" & UserName & " PING Reply]: " & i & " seconds "
    If LCase(lSettings.sNickname) <> LCase(UserName) Then
        lForm.tcp.SendData "NOTICE " & UserName & " :PING " & Val(Trim(Str(DateDiff("s", CVDate("07/28/1979"), Now)))) & Chr(1) & vbCrLf
        lForm.tcp.SendData "NOTICE " & UserName & " :I am running NexIRC©" & vbCrLf
    Else
        lForm.tcp.SendData "NOTICE " & UserName & " :I am running NexIRC©" & vbCrLf
    End If
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Sub ctcp(parms As String, lTextBox as ctlTBox, UserName As String, lForm As Form)"
End Sub

Public Sub numeric(strServer As String, Num As String, NickName As String, strLine As String, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lChanName() As String, mItem As Variant, xChannelName As String, xUsers As String, xTopic As String, cnumber As Long, F As Integer, mbox As VbMsgBoxResult, i As Integer, X As Integer, ix As Integer, word() As String, msg As String, strTemp As String
Static ChannelCount As Integer, ChannelPause As Integer
If Left(strServer, 1) = ":" Then strServer = Right(strServer, Len(strServer) - 1)
'Stop
lRaw.nNum = Num
lRaw.nServer = strServer
lRaw.nNickName = NickName
lRaw.nParms = strLine
lSettings.sServer = strServer
If lMyCurrentModes <> "" Then
    lForm.Caption = lForm.Tag & " [+" & lMyCurrentModes & "] " & NickName & " on " & lSettings.sServer & ":" & lForm.tcp.RemotePort
Else
    lForm.Caption = lForm.Tag & ": " & NickName & " on " & lSettings.sServer & ":" & lForm.tcp.RemotePort
End If
If Num = "311" Or Num = "312" Or Num = "317" Then
    strTemp = strLine
    If InStr(strTemp, Chr(32)) Then
        Do Until InStr(strTemp, Chr(32)) = 0
            X = InStr(strTemp, Chr(32))
            If X Then
                i = i + 1
                ReDim Preserve word(i)
                word(i) = Mid(strTemp, 1, X - 1)
                strTemp = Mid(strTemp, X + 1)
            End If
        Loop
        ReDim Preserve word(i + 1)
        word(i + 1) = strTemp
    End If
End If
Select Case Num
Case "367"
    frmChannelProporties.lstBans.AddItem Right(strLine, Len(strLine) - Len(Parse(strLine, "#", " ")) - 2)
Case "368"
    msg = Trim("#" & Parse(strLine, "#", " "))
    i = FindChannelIndex(msg)
    If i <> 0 And Len(msg) <> 0 Then
        frmChannelProporties.Show 0, mdiNexIRC
        frmChannelProporties.Tag = msg
        'frmChannelProporties.txtTopic.Text = lChannelTopic(i)
        frmChannelProporties.txtTopic.Text = ReturnChannelTopic(i)
        'If InStr(LCase(ReturnChannelModes(  FindChannelIndex(msg)  )), "t") Then
        If InStr(LCase(ReturnChannelModes(FindChannelIndex(msg))), "t") Then
        'If InStr(LCase(lChannelModes(FindChannelIndex(msg))), "t") Then
            frmChannelProporties.chkOnlyOpsSetTopic.Value = 1
        Else
            frmChannelProporties.chkOnlyOpsSetTopic.Value = 0
        End If
'        If InStr(LCase(lChannelModes(FindChannelIndex(msg))), "m") Then
        If InStr(LCase(ReturnChannelModes(FindChannelIndex(msg))), "m") Then
            frmChannelProporties.chkModerated.Value = 1
        Else
            frmChannelProporties.chkModerated.Value = 0
        End If
        If InStr(LCase(ReturnChannelModes(FindChannelIndex(msg))), "n") Then
'        If InStr(LCase(lChannelModes(FindChannelIndex(msg))), "n") Then
            frmChannelProporties.chkNoExternalMessages.Value = 1
        Else
            frmChannelProporties.chkNoExternalMessages.Value = 0
        End If
        If InStr(LCase(ReturnChannelModes(FindChannelIndex(msg))), "i") Then
        'If InStr(LCase(lChannelModes(FindChannelIndex(msg))), "i") Then
            frmChannelProporties.chkInviteOnly.Value = 1
        Else
            frmChannelProporties.chkInviteOnly.Value = 0
        End If
        If InStr(LCase(ReturnChannelModes(FindChannelIndex(msg))), "k") Then
        'If InStr(LCase(lChannelModes(FindChannelIndex(msg))), "k") Then
            frmChannelProporties.chkKey.Value = 1
        Else
            frmChannelProporties.chkKey.Value = 0
        End If
        If InStr(LCase(ReturnChannelModes(FindChannelIndex(msg))), "l") Then
'        If InStr(LCase(lChannelModes(FindChannelIndex(msg))), "l") Then
            frmChannelProporties.chkUserLimit.Value = 1
        Else
            frmChannelProporties.chkUserLimit.Value = 0
        End If
        If InStr(LCase(ReturnChannelModes(FindChannelIndex(msg))), "p") Then
'        If InStr(LCase(lChannelModes(FindChannelIndex(msg))), "p") Then
            frmChannelProporties.chkPrivate.Value = 1
        Else
            frmChannelProporties.chkPrivate.Value = 0
        End If
        If InStr(LCase(ReturnChannelModes(FindChannelIndex(msg))), "s") Then
        'If InStr(LCase(lChannelModes(FindChannelIndex(msg))), "s") Then
            frmChannelProporties.chkSecret.Value = 1
        Else
            frmChannelProporties.chkSecret.Value = 0
        End If
    End If
Case "254"
    RunAutoPerform lForm
Case "375"
    LoadScript "nexirc\numeric\375.txt"
    If lModes.mI = True Then
        lForm.tcp.SendData "mode " & NickName & " +i" & vbCrLf
    End If
    If lModes.mI = True Then
        lForm.tcp.SendData "mode " & NickName & " +s" & vbCrLf
    End If
    If lModes.mI = True Then
        lForm.tcp.SendData "mode " & NickName & " +w" & vbCrLf
    End If
Case "372"
    LoadScript "nexirc\numeric\372.txt"
    If lSettings.sOptions.oShowMOTD = True Then
        If lSettings.sMOTDVisible = False Then frmMOTD.Show
        DoColor frmMOTD.txtIncoming, "" & Color.Normal & Right(strLine, Len(strLine) - 1)
    Else
        Call DoColor(lForm.txtIncoming, "" & Color.Normal & strLine)
    End If
    OnConnectFunc
Case "376"
    LoadScript "nexirc\numeric\376.txt"
    DoColor lForm.txtIncoming, "" & Color.Normal & strLine
    DoColorSep lForm.txtIncoming
Case "353"
    LoadScript "nexirc\numeric\353.txt"
    lChanName = Split(strLine, " ")
    lChanName(2) = Mid(lChanName(2), 2)
    'Stop
    For i = 2 To UBound(lChanName)
        For ix = 1 To ReturnChannelUBound
            If LCase(ReturnChannelName(ix)) = LCase(lChanName(1)) Then
                lChanName(i) = Replace(lChanName(i), "%", "")
                AddUserToNicklist lChanName(i), ReturnChannelNamesListView(ix)
                UpdateChannelCaption ix
            End If
        Next ix
    Next i
    DoColor lForm.txtIncoming, "" & Color.Normal & strLine
Case "366"
    LoadScript "nexirc\numeric\366.txt"
    word = Split(strLine, " ")
    ACTION_CHANNEL = word(0)
    ProcessInput "STATS", lForm.txtIncoming, lForm
    DoColor lForm.txtIncoming, "" & Color.Normal & strLine
    DoColorSep lForm.txtIncoming
    
    'SortNicklist ReturnChannelNamesListView(FindChannelIndex(word(0)))
Case "321"
    frmChannels.Show
    LoadScript "nexirc\numeric\321.txt"
    DoColor lForm.txtIncoming, "" & Color.Normal & "Listing Channels"
    DoColorSep lForm.txtIncoming
    ChannelCount = 0
    ChannelPause = 0
    frmChannels.Show
    frmChannels.lvwChannels.Clear
    'frmChannels.lvwChannels.ListItems.Clear
    'lockwindowupdate frmChannels.lvwChannels.hwnd
Case "322"
    LoadScript "nexirc\numeric\322.txt"
    word = Split(strLine, " ")
    xChannelName = word(0)
    xUsers = word(1)
    For i = 2 To UBound(word)
        xTopic = xTopic & word(i) & " "
    Next i
    frmChannels.lvwChannels.ItemAdd 0, xChannelName, 0, 0
    DoEvents
    'Set mItem = frmChannels.lvwChannels.ListItems.Add(, , xChannelName)
'    mItem.SubItems(1) = xUsers
'    mItem.SubItems(2) = xTopic
    ChannelCount = ChannelCount + 1
    ChannelPause = ChannelPause + 1
    If ChannelPause >= 125 Then
        ChannelPause = 0
        'lockwindowupdate 0
        'frmChannels.lvwChannels.Refresh
        'lockwindowupdate frmChannels.lvwChannels.hwnd
    End If
Case "323"
    LoadScript "nexirc\numeric\323.txt"
    Call DoColor(lForm.txtIncoming, "" & Color.Normal & "End of channel list (" & ChannelCount & ")")
    DoColorSep lForm.txtIncoming
    frmChannels.Caption = "NexIRC - Channel List [" & ChannelCount & "]"
    frmChannels.lvwChannels.Sort 0, soAscending, stString
Case "332"
    LoadScript "nexirc\numeric\332.txt"
    Dim TopicArray() As String
    TopicArray = Split(strLine, " ")
    strLine = ""
    For i = 1 To UBound(TopicArray)
        strLine = strLine & " " & TopicArray(i)
    Next i
    strLine = LTrim(strLine)
    If Left(strLine, 1) = ":" Then
        strLine = Mid(strLine, 2)
        Dim ChanTopic As String
        ChanTopic = strLine
    End If
    For X = 1 To ReturnChannelUBound
        If LCase(ReturnChannelName(Int(X))) = LCase(TopicArray(0)) Then
            SetChannelTopicTextBox X, ""
            SetChannelTopic X, ""
            If Len(ReturnChannelTopic(X)) <> 0 Then DoColor ReturnChannelIncomingTBox(X), "" & Color.Normal & strLine
            SetChannelTopicToolTip X, strLine
            msg = Replace(ReturnChannelTopic(X), Chr(13), "")
            SetChannelTopic X, Replace(msg, Chr(10), "")
            SetChannelCaption X, ReturnChannelName(X) & " [+" & ReturnChannelModes(X) & "] :" & ReturnChannelTopic(X)
            SetChannelTag X, ReturnChannelName(X)
            SetChannelStatsTopic X, strLine
            Call DoColor(ReturnChannelIncomingTBox(X), "" & Color.Join & "• Topic is '" & ChanTopic & "'")
            Exit For
        End If
    Next X
Case 431
    ProcessReplaceString sNoNicknameGiven, lSettings.sActiveServerForm.txtIncoming
    frmNicknameError.Caption = "NexIRC - No Nickname Given"
    frmNicknameError.Show 1
    Exit Sub
Case 461
    If lSettings.sOptions.oShowNotifyInActiveWindow = True Then
        ActiveWindowDoColor "" & Color.Join & "• Nobody on your notify list is on IRC"
    Else
        DoColor lForm.txtIncoming, "" & Color.Join & "• Nobody on your notify list is on IRC"
    End If
    Exit Sub
Case "311"
    LoadScript "nexirc\numeric\311.txt"
    Dim lRealName As String
    For i = 4 To UBound(word)
        lRealName = lRealName & Chr(32) & word(i)
    Next i
    If lSettings.sRetrieveAddressFromWhoisForBan = True Then
        lForm.tcp.SendData "MODE " & lSettings.sBanChannel & " +b :" & word(2) & "@" & word(3) & vbCrLf: DoEvents
        If lSettings.sRetrieveAddressFromWhoisForKickBan = True Then
            lSettings.sActiveServerForm.tcp.SendData "KICK " & lSettings.sBanChannel & " " & lSettings.sBanNickname & vbCrLf: DoEvents
        End If
    Else
        Call DoColor(lForm.txtIncoming, "" & Color.Whois & word(1) & " is " & word(2) & "@" & word(3) & lRealName)
    End If
    lSettings.sBanNickname = ""
    lSettings.sBanChannel = ""
    lSettings.sRetrieveAddressFromWhoisForBan = False
    lSettings.sRetrieveAddressFromWhoisForKickBan = False
Case "378"
    LoadScript "nexirc\numeric\378.txt"
    Call DoColor(lForm.txtIncoming, "" & Color.Whois & strLine)
Case "312"
    LoadScript "nexirc\numeric\312.txt"
    Dim ServerQuote As String
    For i = 3 To UBound(word)
        ServerQuote = ServerQuote & Chr(32) & word(i)
    Next i
    Call DoColor(lForm.txtIncoming, "" & Color.Whois & word(1) & " using " & word(2) & " [" & ServerQuote & "]")
Case "317"
    LoadScript "nexirc\numeric\317.txt"
    Dim Seconds As Integer
    Dim Minutes As Variant
    Seconds = Val(word(2))
    Minutes = Seconds / 60
    For i = 1 To Len(Minutes)
        If Mid(Minutes, i, 1) = "." Then
            Minutes = Mid(Minutes, 1, i - 1)
        End If
    Next i
    Seconds = Seconds - (Val(Minutes) * 60)
    If Val(Minutes) > 0 Then
        Call DoColor(lForm.txtIncoming, "" & Color.Whois & word(1) & " has been idle " & Minutes & " mins " & Seconds & " secs")
    Else
        Call DoColor(lForm.txtIncoming, "" & Color.Whois & word(1) & " has been idle " & Seconds & " secs")
    End If
Case "318"
'    MsgBox "hey"
    LoadScript "nexirc\numeric\318.txt"
    Call DoColor(lForm.txtIncoming, "" & Color.Whois & strLine)
    DoColorSep lForm.txtIncoming
Case "307"
    LoadScript "nexirc\numeric\307.txt"
    Call DoColor(lForm.txtIncoming, "" & Color.Whois & strLine)
Case "319"
    LoadScript "nexirc\numeric\319.txt"
    Call DoColor(lForm.txtIncoming, "" & Color.Whois & Replace(strLine, ":", ""))
Case "001"
    LoadScript "nexirc\numeric\001.txt"
    DoColor lForm.txtIncoming, "" & Color.Server & strLine
    If lSettings.sOptions.oShowChannelFolder = True Then
        frmChannelFolder.Show 0, mdiNexIRC
    End If
    If lSettings.sOptions.oWhoisNotify = True Then
        lForm.tcp.SendData "ISON " & RTrim(ReturnNotifyList) & vbCrLf
    End If
    word = Split(strLine)
    DoColor lForm.txtIncoming, "" & Color.Normal & "• Your IP is " & word(UBound(word))
    OnConnectFunc
Case "002"
    LoadScript "nexirc\numeric\002.txt"
    DoColor lForm.txtIncoming, "" & Color.Server & strLine
Case "003"
    LoadScript "nexirc\numeric\003.txt"
    DoColor lForm.txtIncoming, "" & Color.Server & strLine
Case "004"
    LoadScript "nexirc\numeric\004.txt"
    DoColor lForm.txtIncoming, "" & Color.Server & strLine
Case "005"
    LoadScript "nexirc\numeric\005.txt"
    DoColor lForm.txtIncoming, "" & Color.Server & strLine
    DoColorSep lForm.txtIncoming
Case "251"
    LoadScript "nexirc\numeric\251.txt"
    strLine = Replace(strLine, ":", "")
    DoColor lForm.txtIncoming, "" & Color.Server & strLine
Case "252"
    LoadScript "nexirc\numeric\252.txt"
    strLine = Replace(strLine, ":", "")
    DoColor lForm.txtIncoming, "" & Color.Server & strLine
Case "254"
    LoadScript "nexirc\numeric\254.txt"
    strLine = Replace(strLine, ":", "")
    DoColor lForm.txtIncoming, "" & Color.Server & strLine
Case "255"
    LoadScript "nexirc\numeric\255.txt"
    strLine = Replace(strLine, ":", "")
    DoColor lForm.txtIncoming, "" & Color.Server & strLine & vbCrLf
    DoColorSep lForm.txtIncoming
Case "265"
    LoadScript "nexirc\numeric\265.txt"
    DoColor lForm.txtIncoming, "" & Color.Server & strLine
Case "266"
    LoadScript "nexirc\numeric\266.txt"
    DoColor lForm.txtIncoming, "" & Color.Server & strLine & vbCrLf
    DoColorSep lForm.txtIncoming
Case "401"
    LoadScript "nexirc\numeric\401.txt"
    DoColor lForm.txtIncoming, "" & Color.Server & strLine & vbCrLf
    DoColorSep lForm.txtIncoming
Case "432"
    frmNicknameError.Caption = "NexIRC - Change Nickname"
    frmNicknameError.Show 1
Case "433"
    Exit Sub
    ProcessReplaceString sNicknameInUse, lSettings.sActiveServerForm.txtIncoming, lSettings.sNickname
    If lSettings.sAutoSelectAlternateNickname = True And ReturnAlternateCount <> 0 Then
        SelectRandomAlternate lForm.tcp
    Else
        If lSettings.sGeneralPrompts = True Then
            mbox = MsgBox("The nickname you selected has already been selected by another user. Would you like to select another nickname?", vbYesNo + vbQuestion, App.Title)
            If mbox = vbYes Then
                frmNicknameError.Caption = "NexIRC - Nickname Taken"
                frmNicknameError.Show 1
            End If
        Else
            frmNicknameError.Caption = "NexIRC - Nickname Taken"
            frmNicknameError.Show 1
        End If
    End If
    LoadScript "nexirc\numeric\433.txt"
    DoColor lForm.txtIncoming, "" & Color.Server & strLine & vbCrLf
    DoColorSep lForm.txtIncoming
    Exit Sub
Case "482"
    LoadScript "nexirc\numeric\482.txt"
    DoColor lForm.txtIncoming, "" & Color.Server & strLine & vbCrLf
    DoColorSep lForm.txtIncoming
Case "303"
    LoadScript "nexirc\numeric\303.txt"
    Dim NName() As String
    Dim xON As String
    Static tempNames As String
    strLine = Replace(strLine, ":", "")
    strLine = strLine & " "
    xON = tempNames
    If tempNames = strLine Then
    Else
        NName = Split(strLine, " ")
        For i = 0 To UBound(NName)
            If NName(i) <> "" Then
                X = InStr(tempNames & " ", NName(i) & " ")
                If X Then
                    xON = Replace(xON, NName(i) & " ", "")
                Else
                    If lSettings.sOptions.oShowNotifyInActiveWindow = True Then
                        ActiveWindowDoColor "" & Color.Notify & "• " & NName(i) & " is on IRC"
                    Else
                        DoColor lForm.txtIncoming, "" & Color.Notify & "• " & NName(i) & " is on IRC"
                    End If
                    If lSettings.sOptions.oWhoisNotify = True Then
                        lForm.tcp.SendData "WHOIS " & NName(i) & vbCrLf
                    End If
                    mdiNexIRC.cboNotify.AddItem NName(i)
                    If lSettings.sShowNotifyWindow = True Then
                        If lSettings.sNotifyVisible = True Then
                            frmNotify.lstNotify.AddItem NName(i)
                        Else
                            frmNotify.Show
                            frmNotify.lstNotify.AddItem NName(i)
                        End If
                    End If
                End If
            End If
        Next i
        NName = Split(xON, " ")
        For i = 0 To UBound(NName)
            If Trim(NName(i)) <> "" Then
                DoColor lForm.txtIncoming, "" & Color.Notify & "• " & NName(i) & " has left IRC"
                DoColorSep lForm.txtIncoming
                If lSettings.sNotifyVisible = True Then
                    frmNotify.lstNotify.RemoveItem Trim(FindListBoxIndex(NName(i), frmNotify.lstNotify))
                    mdiNexIRC.cboNotify.RemoveItem FindComboBoxIndex(mdiNexIRC.cboNotify, NName(i))
                End If
            End If
        Next i
    End If
    tempNames = strLine
Case "328"
    LoadScript "nexirc\numeric\328.txt"
    word = Split(strLine, " ")
    word(1) = Mid(word(1), 2)
    For i = 1 To ReturnChannelUBound
        If LCase(ReturnChannelName(i)) = LCase(word(0)) Then
            DoColor ReturnChannelIncomingTBox(i), "" & Color.Join & "• " & word(0) & " homepage is " & word(1)
            DoColorSep ReturnChannelIncomingTBox(i)
        End If
    Next i
Case "333"
    LoadScript "nexirc\numeric\333.txt"
    Dim strTopicTime As String
    word = Split(strLine, " ")
    For i = 1 To ReturnChannelUBound
        If LCase(ReturnChannelName(i)) = LCase(word(0)) Then
            strTopicTime = ReturnIRCTime(word(2))
            DoColor ReturnChannelIncomingTBox(i), "" & Color.Join & "• Topic was set by " & word(1) & " [" & strTopicTime & "]"
            DoColorSep ReturnChannelIncomingTBox(i)
        End If
    Next i
Case "364"
    LoadScript "nexirc\numeric\364.txt"
Case "324"
    LoadScript "nexirc\numeric\324.txt"
    word = Split(strLine, " ")
    For i = 1 To ReturnChannelUBound
        If LCase(ReturnChannelName(i)) = LCase(word(0)) Then
            SetChannelModes i, Mid(word(1), 2)
            UpdateChannelCaption i
            Exit For
        End If
    Next i
Case "329"
    LoadScript "nexirc\numeric\329.txt"
    word = Split(strLine, " ")
    word(1) = ReturnIRCTime(word(1))
    For i = 1 To ReturnChannelUBound
        If LCase(ReturnChannelName(i)) = LCase(word(0)) Then
            DoColor ReturnChannelIncomingTBox(i), "" & Color.Join & "• " & ReturnChannelName(i) & " was created on " & word(1)
            DoColorSep ReturnChannelIncomingTBox(i)
            Exit For
        End If
    Next i
Case Else
    If DoesFileExist(App.Path & "\data\scripts\nexirc\numeric\" & Trim(Str(Num))) & ".txt" = True Then
        LoadScript "nexirc\numeric\" & Trim(Str(Num))
    End If
    DoColor lForm.txtIncoming, "" & Color.Normal & "[04" & Num & "" & Color.Normal & "]10" & strLine
    DoColorSep lForm.txtIncoming
End Select
End Sub

Public Sub SetOp(lMode As String, lUserName As String, lTarget As String, lModeName As String, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, n As Integer, msg As String
For i = 1 To ReturnChannelUBound
    If LCase(ReturnChannelName(i)) = LCase(lTarget) Then
        For n = 1 To ReturnChannelNamesCount(i) - 1
            If lMode = "+" Then
                If LCase(ReturnChannelNames(i, n)) = LCase(lModeName) Or LCase(ReturnChannelNames(i, n)) = LCase("+" & lModeName) Then
                    msg = Replace(lModeName, "@", "")
                    msg = Replace(msg, "+", "")
                    RemoveChannelName i, n
''                    MsgBox "Add user '" & lUserName & "' to '" & ReturnChannelName(i)
                    AddUserToNicklist "@" & msg, ReturnChannelNamesListView(i)
                    If lSettings.sOptions.oShowModes = True Then
                        ProcessReplaceString sUserOped, ReturnChannelIncomingTBox(i), lUserName, lModeName
                    Else
                        ProcessReplaceString sUserOped, lForm.txtIncoming, lUserName, lModeName
                        DoColorSep lForm.txtIncoming
                    End If
                End If
            Else
                If LCase(ReturnChannelNames(i, n)) = LCase(lModeName) Or LCase(ReturnChannelNames(i, n)) = LCase("+" & lModeName) Then
                    msg = Replace(lModeName, "@", "")
                    msg = Replace(msg, "+", "")
                    RemoveChannelName i, n
'                    MsgBox "Add user '" & lUserName & "' to '" & ReturnChannelName(i)
                    AddUserToNicklist msg, ReturnChannelNamesListView(i)
                    If lSettings.sOptions.oShowModes = True Then
                        ProcessReplaceString sUserOped, ReturnChannelIncomingTBox(i), lUserName, lModeName
                    Else
                        ProcessReplaceString sUserDeoped, lForm.txtIncoming, lUserName, lModeName
                        DoColorSep lForm.txtIncoming
                    End If
                End If
            End If
        Next n
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetOp(lMode As String, lUserName As String, lTarget As String, lModeName As String, lForm As Form)"
End Sub

Public Sub SetVoice(lMode As String, lUserName As String, lTarget As String, lModeName As String, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, n As Integer, msg As String
For i = 1 To ReturnChannelUBound
    If LCase(ReturnChannelName(i)) = LCase(lTarget) Then
        For n = 1 To ReturnChannelNamesCount(i) - 1
            If lMode = "+" Then
                If LCase(ReturnChannelNames(i, n)) = LCase(lModeName) Or LCase(ReturnChannelNames(i, n)) = LCase("@" & lModeName) Then
                    msg = Replace(lModeName, "@", "")
                    msg = Replace(msg, "+", "")
                    RemoveChannelName i, n
'                    MsgBox "Add user '" & lUserName & "' to '" & ReturnChannelName(i)
                    AddUserToNicklist "+" & msg, ReturnChannelNamesListView(i)
                    If lSettings.sOptions.oShowModes = True Then
                        ProcessReplaceString sUserVoiced, ReturnChannelIncomingTBox(i), lUserName, lModeName
                    Else
                        ProcessReplaceString sUserVoiced, lForm.txtIncoming, lUserName, lModeName
                        DoColorSep lForm.txtIncoming
                    End If
                End If
            Else
                If LCase(ReturnChannelNames(i, n)) = LCase(lModeName) Or LCase(ReturnChannelNames(i, n)) = LCase("@" & lModeName) Then
                    msg = Replace(lModeName, "@", "")
                    msg = Replace(msg, "+", "")
                    RemoveChannelName i, n
'                    MsgBox "Add user '" & lUserName & "' to '" & ReturnChannelName(i)
                    AddUserToNicklist msg, ReturnChannelNamesListView(i)
                    If lSettings.sOptions.oShowModes = True Then
                        ProcessReplaceString sUserDevoiced, ReturnChannelIncomingTBox(i), lUserName, lModeName
                    Else
                        ProcessReplaceString sUserDevoiced, lForm.txtIncoming, lUserName, lModeName
                        DoColorSep lForm.txtIncoming
                    End If
                End If
            End If
        Next n
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetVoice(lMode As String, lUserName As String, lTarget As String, lModeName As String, lForm As Form)"
End Sub

Public Sub SetInvisible(lMode As String, lUserName As String, lForm As Form)
'Public Sub SetInvisible(lMode As String, lUserName As String, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
If lMode = "+" Then
    lMyCurrentModes = lMyCurrentModes & "i"
Else
    lMyCurrentModes = Replace(lMyCurrentModes, "i", "")
End If
lForm.Caption = lForm.Tag & ": [" & lMyCurrentModes & "] " & lSettings.sNickname & " on " & lSettings.sServer & ":" & lForm.tcp.RemotePort
DoColor lForm.txtIncoming, "" & Color.Mode & "• " & lSettings.sNickname & " sets mode " & lMode & "i"
DoColorSep lForm.txtIncoming
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetInvisible(lMode As String, lUserName As String, lForm As Form)"
End Sub

Public Sub Ban(lMode As String, lUserName As String, lChannelName As String, lModeName As String, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
Dim i As Integer
For i = 1 To ReturnChannelUBound
    If LCase(ReturnChannelName(i)) = LCase(lChannelName) Then
        If lMode = "+" Then
            If lSettings.sOptions.oShowModes = True Then
                Call DoColor(ReturnChannelIncomingTBox(i), "" & Color.Mode & "• " & lUserName & " bans " & lModeName)
            Else
                DoColor lForm.txtIncoming, "" & Color.Mode & "• " & lUserName & " bans " & lModeName & " in " & lChannelName
                DoColorSep lForm.txtIncoming
            End If
        Else
            If lSettings.sOptions.oShowModes = True Then
                Call DoColor(ReturnChannelIncomingTBox(i), "" & Color.Mode & "• " & lUserName & " unbans " & lModeName)
            Else
                DoColor lForm.txtIncoming, "" & Color.Mode & "• " & lUserName & " unbans " & lModeName & " in " & lChannelName
                DoColorSep lForm.txtIncoming
            End If
        End If
    End If
Next i
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Sub Ban(lMode As String, lUserName As String, lChannelName As String, lModeName as String, lForm As Form)"
    Err.Clear
End Sub

Public Sub Limit(lMode As String, lUserName As String, lChannelName As String, lModeName As String)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
Dim i As Integer, n As Integer
Select Case lMode
Case "+"
    For i = 1 To ReturnChannelUBound
        If LCase(ReturnChannelName(i)) = LCase(lChannelName) Then
            SetChannelLimit i, Str(lModeName)
            SetChannelModes i, Replace(ReturnChannelModes(i), "l", "")
            SetChannelModes i, ReturnChannelModes(i) & "l"
            UpdateChannelCaption i
            Exit For
        End If
    Next i
Case "-"
End Select
ErrHandler:
ProcessRuntimeError Err.Description, Err.Number, "Public Sub Limit(lMode As String, lUserName As String, lChannelName As String, lModeName as String)"
End Sub

Public Sub ChangeTopic(lUserName As String, lChannel As String, lTopic As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To ReturnChannelUBound
    If LCase(ReturnChannelName(i)) = LCase(lChannel) Then
        SetChannelTopic i, lTopic
        SetChannelStatsTopic i, lTopic
        Call DoColor(ReturnChannelIncomingTBox(i), "" & Color.Topic & "• " & lUserName & " changes topic to '" & lTopic & "'")
        SetChannelCaption i, ReturnChannelName(i) & " [+" & ReturnChannelModes(i) & "] :" & lTopic
    End If
Next i
End Sub

Public Sub Register(lMode As String, lUserName As String, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lMode = "+" Then
    lMyCurrentModes = lMyCurrentModes & "r"
Else
    lMyCurrentModes = Replace(lMyCurrentModes, "r", "")
End If
lForm.Caption = lForm.Tag & ": [" & lMyCurrentModes & "] " & lSettings.sNickname & " on " & lSettings.sServer
DoColor lForm.txtIncoming, "" & Color.Mode & "• " & lSettings.sNickname & " sets mode " & lMode & "r"
DoColorSep lForm.txtIncoming
End Sub
