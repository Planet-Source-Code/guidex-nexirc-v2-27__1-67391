Attribute VB_Name = "mdlCommands"
Option Explicit

Public Sub Command(lUserName As String, lCommand As String, lTarget As String, lParameters As String, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lTimeStamp As String, i As Integer, x As Integer, Found As Boolean, UserEmail As String, blServer As Boolean, f As Integer
Dim msg As String, msg2 As String
lEvents.parms = lParameters
lEvents.UserName = lUserName
lEvents.Target = lTarget
lEvents.Command = lCommand
lTimeStamp = Now
If Left(lUserName, 1) = ":" Then lUserName = Right(lUserName, Len(lUserName) - 1)
lParameters = LTrim(lParameters)
If Left(lParameters, 1) = ":" Then
    lParameters = Mid(lParameters, 2)
End If
If Left(lTarget, 1) = ":" Then
    lTarget = Mid(lTarget, 2)
End If
blServer = True
For i = 1 To Len(lUserName)
    If Mid(lUserName, i, 1) = "!" Then
        UserEmail = Mid(lUserName, i + 1)
        blServer = False
        lUserName = Mid(lUserName, 1, i - 1)
        If Left(lUserName, 1) = ":" Then
            lUserName = Mid(lUserName, 2)
        End If
    End If
Next i
Dim strCText As String
Select Case UCase(lCommand)
Case "JOIN"
    lEvents.NickJoin = lUserName
    lEvents.JoinChannel = lTarget
    For i = 1 To lChannelUBound
        If lChannelName(i) = "" Then
            If LCase(lUserName) = LCase(lSettings.sNickname) Then
                Load lChannel(i)
                lChannel(i).Show
                lChannel(i).Tag = LCase(lTarget)
                lChannelName(i) = LCase(lTarget)
                lChannelModes(i) = ""
                UpdateCaption i
                chanstats(i).Name = lTarget
                lForm.tcp.SendData "MODE " & lTarget & vbCrLf
                DoColor lChannel(i).txtIncoming, "" & Color.Join & "• Now talking in " & lTarget
                Call AddTaskPanel(lTarget, 2)
                Exit For
            End If
        End If
        If LCase(lChannelName(i)) = LCase(lTarget) Then
            AddUserToNicklist lUserName, lChannel(i).lstNames
            UpdateCaption i
            strCText = ReturnReplacedString(sJoin, lUserName, UserEmail, lTarget)
            strCText = Replace(strCText, "$time", lTimeStamp)
            If lSettings.sOptions.oShowJoinPart = True Then
                If lSettings.sOptions.oShowAddress = True Then
                    Call DoColor(lChannel(i).txtIncoming, "" & Color.Join & strCText)
                Else
                    Call DoColor(lChannel(i).txtIncoming, "" & Color.Join & strCText)
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
    lEvents.ChanPart = lTarget
    lEvents.NickPart = lUserName
    For i = 1 To lChannelUBound
        If lChannelName(i) = LCase(lTarget) Then
            If LCase(lUserName) = LCase(lSettings.sNickname) Then
                Unload lChannel(i)
                lChannelName(i) = ""
                Exit For
            End If
        End If
        If LCase(lChannelName(i)) = LCase(lTarget) Then
            For x = 1 To lChannel(i).lstNames.ListItems.Count - 1
                If lChannel(i).lstNames.SelectedItem.Index Then
                    If LCase(lChannel(i).lstNames.ListItems(x).Text) = LCase(lUserName) Or LCase(lChannel(i).lstNames.ListItems(x).Text) = LCase("@" & lUserName) Or LCase(lChannel(i).lstNames.ListItems(x).Text) = LCase("+" & lUserName) Then
                        lChannel(i).lstNames.ListItems.Remove x
                        UpdateCaption i
                        strCText = ReturnReplacedString(sPart, lUserName, UserEmail, lTarget)
                        strCText = Replace(strCText, "$time", lTimeStamp)
                        If lSettings.sOptions.oShowJoinPart = True Then
                            If lSettings.sOptions.oShowAddress = True Then
                                Call DoColor(lChannel(i).txtIncoming, "" & Color.Part & strCText)
                            Else
                                Call DoColor(lChannel(i).txtIncoming, "" & Color.Part & strCText)
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
                End If
            Next x
        End If
    Next i
    LoadScript "nexirc\onpart.txt"
Case "PRIVMSG"
    If Left(lTarget, 1) = "#" Then
        For i = 1 To lChannelUBound
            If lChannelName(i) = LCase(lTarget) Then
                If LCase(Left(lParameters, 7)) = LCase("ACTION") Then
                    If Right(lParameters, 1) = "" Then
                        lParameters = Mid(lParameters, 1, Len(lParameters) - 1)
                    End If
                    Call DoColor(lChannel(i).txtIncoming, "" & Color.Action & "* " & lUserName & " " & Mid(lParameters, 8))
                Else
                    If CheckIgnoreList(lUserName, lForm) = False Then
                        ProcessReplaceString sPm, lChannel(i).txtIncoming, lUserName, lParameters
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
        Found = False
        For i = 1 To 150
            If LCase(lUserName) = LCase(lQueryName(i)) Then
                Found = True
                Exit For
            End If
        Next i
        If Found = False Then
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
                    If lQueryName(i) = "" Then
                        Load lQuery(i)
                        lQuery(i).Show
                        lQuery(i).Caption = lUserName & " [" & UserEmail & "]"
                        lQueryName(i) = lUserName
                        Call AddTaskPanel(lUserName, 1)
                        IsUserOnline lQuery(i).txtIncoming, lUserName
                        Exit For
                    End If
                Next i
            End If
        End If
        If CheckIgnoreList(lUserName, lForm, True) = False Then
            For i = 1 To 150
                If Len(lQueryName(i)) <> 0 Then
                    If LCase(lQueryName(i)) = LCase(lUserName) Then
                        If Left(lParameters, 1) = "" Then
                            Call CTCP(lParameters, lQuery(i).txtIncoming, lUserName, lForm)
                            lQuery(i).Caption = lUserName & " [" & UserEmail & "]"
                        Else
                            Call DoColor(lQuery(i).txtIncoming, "" & Color.Normal & "<" & Color.Whois & "" & lUserName & "" & Color.Normal & "> " & lParameters)
                            lQuery(i).Caption = lUserName & " [" & UserEmail & "]"
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
    For i = 1 To lChannelUBound
        If lChannelName(i) <> "" Then
            For x = 1 To lChannel(i).lstNames.ListItems.Count - 1
                If LCase(lChannel(i).lstNames.ListItems(x).Text) = LCase(lUserName) Then
                    lChannel(i).lstNames.ListItems.Remove x
                    AddUserToNicklist lTarget, lChannel(i).lstNames
                    f = FindListViewIndex(lChannel(i).lstNames, lTarget)
                    lChannel(i).lstNames.ListItems(f).ForeColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBottomBandsColor
                    DoColor lChannel(i).txtIncoming, "" & Color.Nick & "• " & lUserName & " is now known as " & lTarget
                Else
                    If LCase(lChannel(i).lstNames.ListItems(x).Text) = LCase("@" & lUserName) Then
                        lChannel(i).lstNames.ListItems.Remove x
                        AddUserToNicklist "@" & lTarget, lChannel(i).lstNames
                        f = FindListViewIndex(lChannel(i).lstNames, "@" & lTarget)
                        lChannel(i).lstNames.ListItems(f).ForeColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sTopBandsColor
                        DoColor lChannel(i).txtIncoming, "" & Color.Nick & "• " & lUserName & " is now known as " & lTarget
                    Else
                        If LCase(lChannel(i).lstNames.ListItems(x).Text) = LCase("+" & lUserName) Then
                            lChannel(i).lstNames.ListItems.Remove x
                            AddUserToNicklist "+" & lTarget, lChannel(i).lstNames
                            f = FindListViewIndex(lChannel(i).lstNames, "+" & lTarget)
                            lChannel(i).lstNames.ListItems(f).ForeColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sPeaksColor
                            DoColor lChannel(i).txtIncoming, "" & Color.Nick & "• " & lUserName & " is now known as " & lTarget
                        End If
                    End If
                End If
            Next x
        End If
        For x = 1 To 150
            If LCase(lQueryName(x)) = LCase(lUserName) Then
                lQueryName(x) = lTarget
                Call RemoveTaskbar(lUserName)
                Call AddTaskPanel(lTarget, 1)
                lQuery(x).Caption = lTarget & " [" & UserEmail & "]"
            End If
            Exit For
        Next x
    Next i
    LoadScript "nexirc\onnick.txt"
Case "NOTICE"
    If blServer Then
        Call DoColor(lForm.txtIncoming, "" & Color.Notice & "NOTICE: " & lParameters)
    End If
    If LCase(Left(lParameters, 1)) = LCase("") Then
        Call CTCP(lParameters, lForm.txtIncoming, lUserName, lForm)
        GoTo DONE
    End If
    If Left(lTarget, 1) = "#" Then
        For i = 1 To lChannelUBound
            If LCase(lTarget) = LCase(lChannelName(i)) Then
                DoColor lChannel(i).txtIncoming, "" & Color.Notice & lTarget & ": <" & lUserName & "> " & lParameters
            Else
            End If
        Next i
    Else
        If blServer = False Then
            Found = False
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
    For i = 1 To lChannelUBound
        If lChannelName(i) <> "" Then
            For x = 1 To lChannel(i).lstNames.ListItems.Count - 1
                If LCase(lChannel(i).lstNames.ListItems(x).Text) = LCase(lUserName) Or LCase(lChannel(i).lstNames.ListItems(x).Text) = LCase("@" & lUserName) Or LCase(lChannel(i).lstNames.ListItems(x).Text) = LCase("+" & lUserName) Then
                    lChannel(i).lstNames.ListItems.Remove x
                    UpdateCaption i
                    strCText = ReturnReplacedString(sQuit, lUserName, UserEmail, lParameters)
                    strCText = Replace(strCText, "$time", lTimeStamp)
                    If lSettings.sOptions.oShowQuit = True Then
                        Call DoColor(lChannel(i).txtIncoming, "" & Color.Quit & strCText)
                    Else
                        DoColor lForm.txtIncoming, "" & Color.Quit & strCText
                        DoColorSep lForm.txtIncoming
                        Exit For
                    End If
                End If
            Next x
        End If
    Next i
    LoadScript "nexirc\onquit.txt"
Case "KICK"
    Dim word(1) As String
    i = InStr(2, lParameters, Chr(32))
    word(1) = Trim(Mid(lParameters, 1, i - 1))
    lParameters = Mid(lParameters, 2 + i)
    For i = 1 To lChannelUBound
        If LCase(lTarget) = LCase(lChannelName(i)) Then
            strCText = ReturnReplacedString(sKick, lUserName, word(1), lTarget, lParameters)
            strCText = Replace(strCText, "$time", lTimeStamp)
            If lSettings.sOptions.oShowKicks = True Then
                Call DoColor(lChannel(i).txtIncoming, "" & Color.Kick & strCText)
            Else
                DoColor lForm.txtIncoming, "" & Color.Kick & strCText
                DoColorSep lForm.txtIncoming
            End If
            For x = 1 To lChannel(i).lstNames.ListItems.Count - 1
                If LCase(lChannel(i).lstNames.ListItems(x).Text) = LCase(word(1)) Or LCase(lChannel(i).lstNames.ListItems(x).Text) = LCase("@" & word(1)) Or LCase(lChannel(i).lstNames.ListItems(x).Text) = LCase("+" & word(1)) Then
                    lChannel(i).lstNames.ListItems.Remove x
                    UpdateCaption i
                End If
            Next x
            If lSettings.sOptions.oReJoin = True Then
                lForm.tcp.SendData "JOIN " & lChannelName(i) & vbCrLf
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
        x = 1
        strWord(0) = Trim(strWord(0))
        For i = 1 To Len(strWord(0))
            If strWord(0) = "" Then Exit For
            If Left(strWord(0), 1) = "+" Or Left(strWord(0), 1) = "-" Then
                strMode(x) = Mid(strWord(0), 1, 2)
                CurrentMode = Mid(strWord(0), 1, 1)
                strWord(0) = Mid(strWord(0), 3)
            Else
                strMode(x) = CurrentMode & Mid(strWord(0), 1, 1)
                strWord(0) = Mid(strWord(0), 2)
            End If
            x = x + 1
            i = 1
        Next i
    Else
        Select Case LCase(Mid(strWord(0), 2))
            Case "o"
                Call OP(Left(strWord(0), 1), lUserName, lTarget, strWord(1), lForm)
            Case "v"
                Call VOICE(Left(strWord(0), 1), lUserName, lTarget, strWord(1), lForm)
            Case "i"
                Call INVISIBLE(Left(strWord(0), 1), lUserName, lTarget, lForm)
            Case "r"
                Call REGISTER(Left(strWord(0), 1), lUserName, lTarget, lForm)
            Case "b"
                Call BAN(Left(strWord(0), 1), lUserName, lTarget, strWord(1), lForm)
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
                    MyModes = MyModes & Mid(strWord(0), 2)
                    lForm.Caption = lForm.Tag & ": [" & MyModes & "] " & lSettings.sNickname & " on " & lSettings.sServer & ":" & lForm.tcp.RemotePort
                Else
                    MyModes = Replace(MyModes, Mid(strWord(0), 2), "")
                    lForm.Caption = lForm.Tag & ": [" & MyModes & "] " & lSettings.sNickname & " on " & lSettings.sServer & ":" & lForm.tcp.RemotePort
                End If
            End Select
        End If
        For i = 1 To 4
            If strMode(i) <> "" Then
                Select Case LCase(Mid(strMode(i), 2))
                Case "o"
                    Call OP(Left(strMode(i), 1), lUserName, lTarget, strWord(i), lForm)
                Case "v"
                    Call VOICE(Left(strMode(i), 1), lUserName, lTarget, strWord(i), lForm)
                Case "i"
                    Call INVISIBLE(Left(strMode(i), 1), lUserName, lTarget, lForm)
                Case "b"
                    Call BAN(Left(strMode(i), 1), lUserName, lTarget, strWord(i), lForm)
                Case "r"
                    Call REGISTER(Left(strMode(i), 1), lUserName, lTarget, lForm)
                Case Else
                    DoColor lForm.txtIncoming, "" & Color.Mode & "• " & Replace(lUserName, ":", "") & " sets mode: " & lParameters
                    DoColorSep lForm.txtIncoming
                    If Left(strMode(i), 1) = "+" Then
                        MyModes = MyModes & Mid(strMode(i), 2)
                        lForm.Caption = lForm.Tag & ": [" & MyModes & "] " & lSettings.sNickname & " on " & lSettings.sServer & ":" & lForm.tcp.RemotePort
                    Else
                        MyModes = Replace(MyModes, Mid(strMode(i), 2), "")
                        lForm.Caption = lForm.Tag & ": [" & MyModes & "] " & lSettings.sNickname & " on " & lSettings.sServer & ":" & lForm.tcp.RemotePort
                    End If
                    For x = 1 To lChannelUBound
                        If LCase(lChannelName(i)) = LCase(lTarget) Then
                            If Left(strMode(i), 1) = "-" Then
                            Else
                            End If
                            Exit For
                        End If
                    Next x
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

Sub CTCP(parms As String, RTF As TBox, UserName As String, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim word() As String, x As Integer, y As Integer, i As Integer, msg As String, lIP As String
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
    For x = 1 To UBound(word)
        parms = parms & word(x) & Chr(32)
    Next x
    Call DoColor(RTF, "" & Color.Action & "* " & UserName & " " & parms)
Case "DCC"
    Dim wDCC As New frmDCCACCEPT, NewFile As frmDCCFILE
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
    lForm.tcp.SendData "NOTICE " & UserName & ":VERSION " & ReturnReplacedString(sVersion, App.major, App.minor)
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
    x = Val(Trim(Str(DateDiff("s", CVDate("01/01/1970"), Now)))) - Val(Int(word(1)))
    i = Int(x)
    DoColor RTF, "" & Color.CTCP & "[" & UserName & " PING Reply]: " & i & " seconds "
    If LCase(lSettings.sNickname) <> LCase(UserName) Then
        lForm.tcp.SendData "NOTICE " & UserName & " :PING " & Val(Trim(Str(DateDiff("s", CVDate("07/28/1979"), Now)))) & Chr(1) & vbCrLf
        lForm.tcp.SendData "NOTICE " & UserName & " :I am running NexIRC©" & vbCrLf
    Else
        lForm.tcp.SendData "NOTICE " & UserName & " :I am running NexIRC©" & vbCrLf
    End If
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Sub ctcp(parms As String, RTF as tBox, UserName As String, lForm As Form)"
End Sub
