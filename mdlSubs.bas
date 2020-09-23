Attribute VB_Name = "mdlSubs"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private lAboutInt As Integer

Public Function ReturnStrippedColorCodes(ByRef msg As String) As String
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
msg = Replace(msg, "", "")
msg = Replace(msg, "", "")
ReturnStrippedColorCodes = msg
End Function

Public Sub ActivateActiveFormResize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If mdiNexIRC.Visible = True Then
    mdiNexIRC.ActiveForm.ActivateResize
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ActivateActiveFormResize()"
End Sub

Public Sub CheckTextBox(lTextBox As TextBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lTextBox.SelStart = 0
lTextBox.SelStart = Len(lTextBox.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub 'checktextbox(lTextBox As TextBox)"
End Sub

'Public Sub SortNicklist(lListView As ctlListView)
'Dim i As Integer, lOpItems(1 To 1000) As ListItem, lVoiceItems(1 To 1000) As ListItem, lNormalItems(1 To 1000) As ListItem, lCount(1 To 3) As Integer, lOpIcon As StdPicture, lVoiceIcon As StdPicture, c As Integer
'Dim mColor(0 To 15) As Long
'mColor(0) = vbWhite
'mColor(1) = vbBlack
'mColor(2) = RGB(42, 42, 87)
'mColor(3) = RGB(33, 112, 33)
'mColor(4) = vbRed
'mColor(5) = RGB(109, 50, 50)
'mColor(6) = RGB(119, 33, 119)
'mColor(7) = RGB(252, 127, 0)
'mColor(8) = RGB(195, 195, 56)
'mColor(9) = RGB(0, 252, 0)
'mColor(10) = RGB(89, 167, 179)
'mColor(11) = RGB(0, 255, 255)
'mColor(12) = vbBlue
'mColor(13) = RGB(255, 0, 255)
'mColor(14) = RGB(127, 127, 127)
'mColor(15) = RGB(210, 210, 210)
''For i = 1 To lListView.ListItems.Count
'For i = 1 To lListView.Count
'    'If Len(lListView.ListItems(i).Text) <> 0 Then
'    If Len(lListView.ItemText(i)) <> 0 Then
'        'Select Case LCase(Trim(lListView.ListItems(i).Tag))
'        'Case "o"
'        '    lCount(1) = lCount(1) + 1
'        '    Set lOpItems(lCount(1)) = lListView.ListItems(i)
'        'Case "v"
'        '    lCount(2) = lCount(2) + 1
'        '    Set lVoiceItems(lCount(2)) = lListView.ListItems(i)
'        'Case "n"
'        '    lCount(3) = lCount(3) + 1
'        '    Set lNormalItems(lCount(3)) = lListView.ListItems(i)
'        'Case Else
'        '    lCount(3) = lCount(3) + 1
'        '    Set lNormalItems(lCount(3)) = lListView.ListItems(i)
'        'End Select
'    End If
'    If Err.Number <> 0 Then
'        Err.Clear
'        Exit For
'    End If
'Next i
''Next i
'lListView.ListItems.Clear
'For i = 1 To lCount(1)
'    Select Case lListView.BackColor
'    Case 0
'        lListView.ListItems.Add , , lOpItems(i).Text, 3, 3
'    Case Else
'        lListView.ListItems.Add , , lOpItems(i).Text, 1, 1
'    End Select
'    c = FindListViewIndex(lListView, lOpItems(i).Text)
'    lListView.ListItems(c).ForeColor = mColor(Int(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nOpColor) - 1)
'    lListView.ListItems(c).Tag = "o"
'    If Err.Number <> 0 Then
'        Err.Clear
'        Exit For
'    End If
'Next i
'For i = 1 To lCount(2)
'    Select Case lListView.BackColor
'    Case 0
'        lListView.ListItems.Add , , lVoiceItems(i).Text, 4, 4
'    Case Else
'        lListView.ListItems.Add , , lVoiceItems(i).Text, 2, 2
'    End Select
'    c = FindListViewIndex(lListView, lVoiceItems(i).Text)
'    lListView.ListItems(c).ForeColor = mColor(Int(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nVoiceColor) - 1)
'    lListView.ListItems(c).Tag = "v"
'    If Err.Number <> 0 Then
'        Err.Clear
'        Exit For
'    End If
'Next i
'For i = 1 To lCount(3)
'    Select Case lListView.BackColor
'    Case 0
'        lListView.ListItems.Add , , lNormalItems(i).Text
'    Case Else
'        lListView.ListItems.Add , , lNormalItems(i).Text
'    End Select
'    c = FindListViewIndex(lListView, lNormalItems(i).Text)
'    lListView.ListItems(c).ForeColor = mColor(Int(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nNormalColor) - 1)
'    lListView.ListItems(c).Tag = "n"
'    If Err.Number <> 0 Then
'        Err.Clear
'        Exit For
'    End If
'Next i
'If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SortNicklist(lListView As ListView)"
'End Sub

Public Sub UnloadProgram()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, F As Integer, o As Boolean
If lPlayback.pPlaying = True Then MenuStop: DoEvents
lSettings.sMainVisisble = False
mdiNexIRC.Visible = False
mdiNexIRC.tmrCheckButtonColors.Enabled = False
mdiNexIRC.tmrContinuousPlay.Enabled = False
mdiNexIRC.tmrPlaySoon.Enabled = False
mdiNexIRC.tmrSendUserPlaylist.Enabled = False
lSettings.sPlaySounds = False
For F = 0 To ReturnStatusWindowCount
    If Len(ReturnStatusWindowServer(F)) <> 0 Then
        If IsStatusTCPConnected(F) = True Then
            o = True
            StatusWindowSendData F, "QUIT :" & ReturnStringDataByType(sQuit): DoEvents
        End If
    End If
Next F
WriteINI GetINIFile(iIRC), mdiNexIRC.Name, "LEFT", mdiNexIRC.Left
WriteINI GetINIFile(iIRC), mdiNexIRC.Name, "TOP", mdiNexIRC.Top
WriteINI GetINIFile(iIRC), mdiNexIRC.Name, "WIDTH", mdiNexIRC.Width
WriteINI GetINIFile(iIRC), mdiNexIRC.Name, "HEIGHT", mdiNexIRC.Height
If lSettings.sAddMediaVisible = True Then Unload frmAddMedia
If lSettings.sChannelListVisible = True Then Unload frmChannelListing
If lSettings.sChannelFolderVisible = True Then Unload frmChannelFolder
If lSettings.sConnectionManagerVisible = True Then Unload frmConnectionManager
If lSettings.sCustomizeVisible = True Then Unload frmCustomize
If lSettings.sIRCServerVisible = True Then Unload frmIRCServer
If lSettings.sMobileMixerVisible = True Then Unload frmMobileMixer
If lSettings.sMOTDVisible = True Then Unload frmMOTD
If lSettings.sNotifyVisible = True Then Unload frmNotify
If lSettings.sPlaylistVisible = True Then Unload frmPlaylist
If lSettings.sSetupWizardVisible = True Then Unload frmSetupWizard
If lSettings.sWebVisible = True Then Unload frmWeb
If lSettings.sSplashVisible = True Then Unload frmSplash

Unload frmMenus
Unload frmGraphics
CloseAll
For i = 1 To ReturnTCPUBound
    Unload mdiNexIRC.wskChat2(i)
    Unload mdiNexIRC.wskChat(i)
Next i
Unload frmChannels
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub UnloadProgram()"
If o = False Then
    mdiNexIRC.tmrDIE.Enabled = True
Else
    mdiNexIRC.tmrEndSoon.Enabled = True
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub UnloadProgram()"
End Sub

Public Sub SendQuitMessage(lForm As Form, Optional lMessage As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
If Len(lMessage) <> 0 Then
    lForm.tcp.SendData "QUIT :" & lMessage & vbCrLf
    ProcessReplaceString sQuitMessage, lForm.txtIncoming, lMessage
Else
    msg = ReturnReplacedString(sQuitReason)
    If Len(msg) <> 0 Then
        lForm.tcp.SendData "QUIT:" & msg & vbCrLf
        ProcessReplaceString sQuitMessage, lForm.txtIncoming, msg
    Else
        lForm.tcp.SendData "QUIT:NexIRC" & lForm.txtIncoming, vbCrLf
        ProcessReplaceString sQuitMessage, lForm.txtIncoming, "NexIRC"
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SendQuitMessage(lStatus As Integer)"
End Sub

Public Sub ProcessRuntimeError(lErrorDescription As String, lErrorNumber As Integer, lSub As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Err = 0
Dim i As Integer
If Len(lErrorDescription) <> 0 And Len(lSub) <> 0 Then WriteINI GetINIFile(iErrorLog), lSub, Date & ": " & Time$ & "(" & GetRnd(10) & ")", Str(lErrorNumber) & ": " & lErrorDescription
End Sub

Public Sub ShowCustomize(lFrame As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
frmCustomize.Show 0, mdiNexIRC
For i = 0 To 9
    frmCustomize.fraSettings(i).Visible = False
    frmCustomize.optCheck(i).Value = False
Next i
frmCustomize.fraSettings(lFrame).Visible = True
frmCustomize.optCheck(lFrame).Value = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ShowCustomize(lFrame As Integer)"
End Sub

Public Sub Surf(lUrl As String, lhWnd As Long)
Dim msg As Long, mbox As VbMsgBoxResult
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Asc(Left(lUrl, 1)) = 39 Then lUrl = Right(lUrl, Len(lUrl) - 1)
If Asc(Right(lUrl, 1)) = 39 Then lUrl = Left(lUrl, Len(lUrl) - 1)
If Left(LCase(lUrl), 7) <> "http://" Then lUrl = "http://" & lUrl
If lSettings.sGeneralPrompts = True Then
    mbox = MsgBox("Would you like to navigate the url '" & lUrl & "'?", vbYesNo + vbQuestion, "NexIRC")
Else
    mbox = vbYes
End If
If mbox = vbNo Then Exit Sub
If lSettings.sBackgroundWebpage = False Then
    msg = ShellExecute(lhWnd, vbNullString, lUrl, vbNullString, "c:\", 1)
ElseIf lSettings.sBackgroundWebpage = True Then
    If lSettings.sWebVisible = True Then
        frmWeb.web.Navigate lUrl
    Else
        frmWeb.Show
        frmWeb.web.Navigate lUrl
    End If
    frmWeb.Visible = True
    lSettings.sWebVisible = True
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub Surf(lUrl As String, lHwnd As Long)"
End Sub

Public Sub Pause(interval)
'Exit Sub
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim Current
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub Pause(interval)"
End Sub

Public Sub ImageBoxMouseDown(lButton As Integer, lImageBox As Image, lImg2 As Image)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lButton = 1 Then
    lImageBox.Picture = lImg2.Picture
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ImageBoxMouseDown(lButton As Integer, lImageBox As Image, lImg2 As Image)"
End Sub

Public Sub ImageBoxMouseMove(lButton As Integer, lImageBox As Image, lImg1 As Image, lImg2 As Image, lX As Single, lY As Single, Optional lImg3 As Image)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lButton = 1 And lImageBox.Picture = lImg2.Picture Then
    If lX > lImageBox.Width Or lX < -1 Or lY > lImageBox.Height Or lY < -1 Then lImageBox.Picture = lImg1.Picture
ElseIf lButton = 1 And lImageBox.Picture = lImg1.Picture Then
    If lX < lImageBox.Width And lX > -1 And lY < lImageBox.Height And lY > -1 Then lImageBox.Picture = lImg2.Picture
ElseIf lButton = 0 Then
    lImageBox.Picture = lImg3.Picture
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ImageBoxMouseMove(lButton As Integer, lImageBox As Image, lImg1 As Image, lImg2 As Image, lX As Single, lY As Single, Optional lImg3 As Image)"
End Sub

Public Sub ImageBoxMouseUp(lButton As Integer, lImageBox As Image, lImg1 As Image, lImg2 As Image)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lButton = 1 And lImageBox.Picture = lImg2.Picture Then
    lImageBox.Picture = lImg1.Picture
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ImageBoxMouseUp(lButton As Integer, lImageBox As Image, lImg1 As Image, lImg2 As Image)"
End Sub

Public Sub PictureBoxMouseDown(lButton As Integer, lPictureBox As PictureBox, lPic2 As PictureBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lButton = 1 Then
    If lPictureBox.Picture <> lPic2.Picture Then
        lPictureBox.Picture = lPic2.Picture
        SetPictureColor lPictureBox, lRedColor, lBlueColor, lGreenColor, True
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub PictureBoxMouseDown(lButton As Integer, lPictureBox As PictureBox, lPic2 As PictureBox)"
End Sub

Public Sub PictureBoxMouseMove(lButton As Integer, lPictureBox As PictureBox, lPic1 As PictureBox, lPic2 As PictureBox, lX As Single, lY As Single, lPic3 As PictureBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lButton = 1 And lPictureBox.Picture = lPic2.Picture Then
    If lX > lPictureBox.Width Or lX < -1 Or lY > lPictureBox.Height Or lY < -1 Then
        If lPictureBox.Picture <> lPic1.Picture Then
            lPictureBox.Picture = lPic1.Picture
            SetPictureColor lPictureBox, lRedColor, lBlueColor, lGreenColor, True
        End If
    End If
ElseIf lButton = 1 And lPictureBox.Picture = lPic1.Picture Then
    If lX < lPictureBox.Width And lX > -1 And lY < lPictureBox.Height And lY > -1 Then
        If lPictureBox.Picture <> lPic2.Picture Then
            lPictureBox.Picture = lPic2.Picture
            SetPictureColor lPictureBox, lRedColor, lBlueColor, lGreenColor, True
        End If
    End If
ElseIf lButton = 0 Then
    If lPictureBox.Picture <> lPic3.Picture Then
        If lPictureBox.Picture <> lPic3.Picture Then
            lPictureBox.Picture = lPic3.Picture
            SetPictureColor lPictureBox, lRedColor, lBlueColor, lGreenColor, True
        End If
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub PictureBoxMouseMove(lButton As Integer, lPictureBox As PictureBox, lPic1 As PictureBox, lPic2 As PictureBox, lX As Single, lY As Single, Optional lPic3 As PictureBox)"
End Sub

Public Sub PictureBoxMouseUp(lButton As Integer, lPictureBox As PictureBox, lPic1 As PictureBox, lPic2 As PictureBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lButton = 1 And lPictureBox.Picture = lPic2.Picture Then
    lPictureBox.Picture = lPic1.Picture
    SetPictureColor lPictureBox, lRedColor, lBlueColor, lGreenColor, True
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub PictureBoxMouseUp(lButton As Integer, lPictureBox As PictureBox, lPic1 As PictureBox, lPic2 As PictureBox)"
End Sub

Public Sub OnConnectFunc()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sAutoJoinActivated = False Then
    lSettings.sAutoJoinActivated = True
    ActivateAutoJoin False
    If lSettings.sOptions.oShowChannelFolder = True Then
        frmChannelFolder.Show 0, mdiNexIRC
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub OnConnectFunc()"
End Sub

Public Sub SaveSettings()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer

WriteINI GetINIFile(iIRC), "Settings", "ShowExtraProgress", lSettings.sShowExtraProgress
WriteINI GetINIFile(iIRC), "Settings", "Balloons", lSettings.sBalloons
WriteINI GetINIFile(iIRC), "Settings", "ShowWhoisInChannel", lSettings.sShowWhoisInChannel
WriteINI GetINIFile(iIRC), "Settings", "AutoPortScanner", lSettings.sAutoPortScanner
WriteINI GetINIFile(iIRC), "Settings", "ExclusiveToMp3InPlaylist", lSettings.sExlusiveToMp3InPlaylist
WriteINI GetINIFile(iIRC), "Settings", "UseNickCompletor", lSettings.sUseNickCompletor
WriteINI GetINIFile(iIRC), "Settings", "UpdateCheck", lSettings.sUpdateCheck
WriteINI GetINIFile(iIRC), "Settings", "ByPassStartupScreen", lSettings.sByPassStartupScreen
WriteINI GetINIFile(iIRC), "Settings", "RepeatCurrentTrack", lPlayback.pRepeatCurrentTrack
WriteINI GetINIFile(iIRC), "Settings", "ServerMinimum", lSettings.sServerMinimum
WriteINI GetINIFile(iIRC), "Settings", "ShowSmallNetworks", lSettings.sShowSmallNetworks
WriteINI GetINIFile(iIRC), "Settings", "DownloadManager", lSettings.sDownloadManager
WriteINI GetINIFile(iIRC), "Settings", "BorderlessObjects", lSettings.sBorderlessObjects
WriteINI GetINIFile(iIRC), "Settings", "TimeStamping", lSettings.sTimeStamping
WriteINI GetINIFile(iIRC), "Settings", "ColoredNicklist", lSettings.sColoredNicklist
WriteINI GetINIFile(iIRC), "Settings", "SecureQuery", lSettings.sSecureQuery
WriteINI GetINIFile(iIRC), "Settings", "AutosizeStatusbarItems", lSettings.sAutosizeStatusbarItems
WriteINI GetINIFile(iIRC), "Settings", "ShowTips", lSettings.sShowTips
WriteINI GetINIFile(iIRC), "Settings", "EnableList", lSettings.sEnableList
WriteINI GetINIFile(iIRC), "Settings", "ShowQuickNotify", lSettings.sShowQuickNotify
WriteINI GetINIFile(iIRC), "Settings", "ReconnectOnDisconnect", lSettings.sReconnectOnDisconnect
WriteINI GetINIFile(iIRC), "Settings", "OfferWhenPlayed ", lSettings.sOfferWhenPlayed
WriteINI GetINIFile(iIRC), "Settings", "AudioServer", lSettings.sAudioServer
WriteINI GetINIFile(iIRC), "Settings", "FileOfferInChannel", lSettings.sFileOfferInChannel
WriteINI GetINIFile(iIRC), "Settings", "ButtonType", lSettings.sButtonType
WriteINI GetINIFile(iIRC), "Settings", "SaveIRCColorsToTheme", lSettings.sSaveIRCColorsToTheme
WriteINI GetINIFile(iIRC), "Settings", "ApplyThemeToIRCColors", lSettings.sApplyThemeToIRCColors
WriteINI GetINIFile(iIRC), "Settings", "AutoSelectAlternateNickname", lSettings.sAutoSelectAlternateNickname
WriteINI GetINIFile(iIRC), "Settings", "RefreshPictureColors", lSettings.sRefreshPictureColors
WriteINI GetINIFile(iIRC), "Settings", "ShowServerOnStartup", lSettings.sShowServerOnStartup
WriteINI GetINIFile(iIRC), "Settings", "ShowChannelFolder", lSettings.sOptions.oShowChannelFolder
WriteINI GetINIFile(iIRC), "Settings", "ShowNotifyWindow", lSettings.sShowNotifyWindow
WriteINI GetINIFile(iIRC), "Settings", "LogoTwitchOnPeaks", lSettings.sLogoTwitchOnPeaks
WriteINI GetINIFile(iIRC), "Settings", "AlwaysShowAudioSettings", lSettings.sAlwaysShowAudioSettings
WriteINI GetINIFile(iIRC), "Settings", "DCCEnabled", lSettings.sDCCEnabled
WriteINI GetINIFile(iIRC), "Settings", "ExclusiveToMp3InPlaylist", lPlayback.pCurrentEngine
WriteINI GetINIFile(iIRC), "Settings", "CurrentEngine", lPlayback.pCurrentEngine
WriteINI GetINIFile(iIRC), "Settings", "Homepage", lSettings.sHomepage
WriteINI GetINIFile(iIRC), "Settings", "SearchForMedia", lSettings.sSearchForMedia
WriteINI GetINIFile(iIRC), "Settings", "NavigateOnStartup", lSettings.sNavigateOnStartup
WriteINI GetINIFile(iIRC), "Settings", "AutojoinEnabled", lSettings.sAutoJoinEnabled
WriteINI GetINIFile(iIRC), "Settings", "ShowQuickMix", lSettings.sShowQuickmix
WriteINI GetINIFile(iIRC), "Settings", "AddJoinedChannelsToChannelFolder", lSettings.sAddJoinedChannelsToChannelFolder
WriteINI GetINIFile(iIRC), "Settings", "Shuffle", lSettings.sShuffle
WriteINI GetINIFile(iIRC), "Settings", "ContinuousPlay", lSettings.sContinuousPlay
WriteINI GetINIFile(iIRC), "Settings", "GeneralPrompts", lSettings.sGeneralPrompts
WriteINI GetINIFile(iIRC), "Settings", "DCCPrompts", lSettings.sDCCPrompts
WriteINI GetINIFile(iIRC), "Settings", "BackgroundWebpage", lSettings.sBackgroundWebpage
WriteINI GetINIFile(iIRC), "Settings", "BGColor", lSettings.sBGColor
WriteINI GetINIFile(iIRC), "Settings", "BGPicture", lSettings.sBGPicture
WriteINI GetINIFile(iIRC), "Settings", "PlaySounds", lSettings.sPlaySounds
WriteINI GetINIFile(iInitialValues), "Settings", "Bass", lInitialAudioValues.iBass
WriteINI GetINIFile(iInitialValues), "Settings", "CDAudio", lInitialAudioValues.iCDAudio
WriteINI GetINIFile(iInitialValues), "Settings", "LineIN", lInitialAudioValues.iLineIN
WriteINI GetINIFile(iInitialValues), "Settings", "Mic", lInitialAudioValues.iMic
WriteINI GetINIFile(iInitialValues), "Settings", "Treble", lInitialAudioValues.iTreble
WriteINI GetINIFile(iInitialValues), "Settings", "Wave", lInitialAudioValues.iWave
WriteINI GetINIFile(iInitialValues), "Settings", "InitialBassEnabled", lInitialAudioValues.iInitialBassEnabled
WriteINI GetINIFile(iInitialValues), "Settings", "InitialCDAudioEnabled", lInitialAudioValues.iInitialCDAudioEnabled
WriteINI GetINIFile(iInitialValues), "Settings", "InitialLineInEnabled", lInitialAudioValues.iInitialLineInEnabled
WriteINI GetINIFile(iInitialValues), "Settings", "InitialMicEnabled", lInitialAudioValues.iInitialMicEnabled
WriteINI GetINIFile(iInitialValues), "Settings", "InitialTrebleEnabled", lInitialAudioValues.iInitialTrebleEnabled
WriteINI GetINIFile(iInitialValues), "Settings", "InitialWaveEnabled", lInitialAudioValues.iInitialWaveEnabled
SaveTextStrings
WriteINI GetINIFile(iIRC), "Ident", "UserID", lSettings.sIdent.iUserID
WriteINI GetINIFile(iIRC), "Ident", "System", lSettings.sIdent.iSystem
WriteINI GetINIFile(iIRC), "Ident", "Port", lSettings.sIdent.iPort
WriteINI GetINIFile(iIRC), "Ident", "Ident", lSettings.sIdent.iEnabled
WriteINI GetINIFile(iIRC), "Ident", "Show", lSettings.sIdent.iShow
WriteINI GetINIFile(iIRC), "Ignore", "Enabled", ReturnIgnoreEnabled
WriteINI GetINIFile(iIRC), "Ignore", "Count", Str(ReturnIgnoreCount)
WriteINI GetINIFile(iIRC), "Info", "Network", lSettings.sNetwork
WriteINI GetINIFile(iIRC), "Info", "Server", lSettings.sServer
WriteINI GetINIFile(iIRC), "Info", "Port", lSettings.sPort
WriteINI GetINIFile(iIRC), "Info", "Password", lSettings.sPassword
WriteINI GetINIFile(iIRC), "Info", "Nickname", lSettings.sNickname
WriteINI GetINIFile(iIRC), "Info", "Username", lSettings.sEMail
WriteINI GetINIFile(iIRC), "Info", "Realname", lSettings.sRealName
WriteINI GetINIFile(iIRC), "Irc", "AutoJoinOnInvite", lSettings.sAutoJoinOnInvite
WriteINI GetINIFile(iIRC), "Irc", "Colors", Trim(lSettings.sColors)
WriteINI GetINIFile(iIRC), "Irc", "ConnectOnStartup", lSettings.sConnectOnStartup
WriteINI GetINIFile(iIRC), "Irc", "Invisable", lSettings.sOptions.oInvisable
WriteINI GetINIFile(iIRC), "Irc", "Rejoin", lSettings.sOptions.oReJoin
WriteINI GetINIFile(iIRC), "Irc", "SkipMotd", lSettings.sOptions.oSkipMOTD
WriteINI GetINIFile(iIRC), "Irc", "ShowMotd", lSettings.sOptions.oShowMOTD
WriteINI GetINIFile(iIRC), "Irc", "ShowOptionsOnStartup", lSettings.sShowOptionsOnStartup
WriteINI GetINIFile(iIRC), "Irc", "ShowAddress", lSettings.sOptions.oShowAddress
WriteINI GetINIFile(iIRC), "Irc", "ShowSplashOnStartup", lSettings.sShowSplashOnStartup
WriteINI GetINIFile(iIRC), "Irc", "ShowNotifyInActive", lSettings.sOptions.oShowNotifyInActiveWindow
WriteINI GetINIFile(iIRC), "Irc", "WhoIsNotify", lSettings.sOptions.oWhoisNotify
WriteINI GetINIFile(iIRC), "Irc", "Whois", lSettings.sOptions.oWhois
WriteINI GetINIFile(iIRC), "Ison", "Ison", lSettings.sISON
WriteINI GetINIFile(iIRC), "Notify", "Count", Str(ReturnNotifyCount)
WriteINI GetINIFile(iIRC), "Notify", "Enabled", ReturnNotifyEnabled
WriteINI GetINIFile(iIRC), "Show", "ServerMessages", lSettings.sOptions.oServerMessages
WriteINI GetINIFile(iIRC), "Show", "WallOps", lSettings.sOptions.oOpMessages
WriteINI GetINIFile(iIRC), "Show", "Quit", lSettings.sOptions.oShowQuit
WriteINI GetINIFile(iIRC), "Show", "JoinPart", lSettings.sOptions.oShowJoinPart
WriteINI GetINIFile(iIRC), "Show", "Modes", lSettings.sOptions.oShowModes
WriteINI GetINIFile(iIRC), "Show", "Topics", lSettings.sOptions.oShowTopics
WriteINI GetINIFile(iIRC), "Show", "Kicks", lSettings.sOptions.oShowKicks
WriteINI GetINIFile(iSpectrum), Trim(Str(lSpectrumThemes.sIndex)), "ProgressBarStyle", Trim(Str(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sProgressBarStyle))
WriteINI GetINIFile(iSpectrum), Trim(Str(lSpectrumThemes.sIndex)), "ProgressBarColor", Trim(Str(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sProgressBarColor))
SaveAutoPerform True
SaveBotCommands
SaveBots
SaveNotify
SaveAlternates
SaveChanFolders
SaveAutoConnect
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SaveSettings()"
End Sub

Public Sub PreloadSettings()
On Local Error Resume Next
Dim i As String, E As Integer, msg As String
SetINIFiles
E = Int(Trim(ReadINI(GetINIFile(iSpectrum), "Settings", "Index", "0")))
lSpectrumThemes.sStartupGraphic = App.Path & "\data\" & ReadINI(GetINIFile(iSpectrum), Str(E), "StartupGraphic", "")
lSettings.sColors = "0:01 1:01 2:09 3:09 4:15 5:03 6:03 7:07 8:08 9:00 10:10 11:08 12:08 13:08 14:10 15:14"
lSettings.sUpdateCheck = ReadINI(GetINIFile(iIRC), "Settings", "UpdateCheck", True)
lSettings.sGeneralPrompts = ReadINI(GetINIFile(iIRC), "Settings", "GeneralPrompts", False)
lSettings.sDCCPrompts = ReadINI(GetINIFile(iIRC), "Settings", "DCCPrompts", True)
lSettings.sShowSplashOnStartup = ReadINI(GetINIFile(iIRC), "IRC", "ShowSplashOnStartup", True)
If lSettings.sShowSplashOnStartup = True Then frmSplash.Show
lSettings.sNickname = ReadINI(GetINIFile(iIRC), "Info", "Nickname", "")
lSettings.sEMail = ReadINI(GetINIFile(iIRC), "Info", "Username", "")
lSettings.sRealName = ReadINI(GetINIFile(iIRC), "Info", "Realname", "")
lSettings.sHandleErrors = ReadINI(GetINIFile(iIRC), "Settings", "HandleErrors", True)
lSettings.sButtonType = ReadINI(GetINIFile(iIRC), "Settings", "ButtonType", 12)
lSettings.sTimeStamping = ReadINI(GetINIFile(iIRC), "Settings", "TimeStamping", True)
lSettings.sShowExtraProgress = CBool(ReadINI(GetINIFile(iIRC), "Settings", "ShowExtraProgress", True))
LoadSettings
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub PreloadSettings()"
End Sub

Public Sub CutRegion(ctlHwnd As Long, Ctl As ComboBox, bCut As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim hRgn As Long
If bCut = True Then
    hRgn = CreateRectRgn(1, 1, ((Ctl.Width / Screen.TwipsPerPixelX) - 3), ((Ctl.Height / Screen.TwipsPerPixelY) - 3))
Else
    hRgn = CreateRectRgn(0, 0, (Ctl.Width / Screen.TwipsPerPixelX), (Ctl.Height / Screen.TwipsPerPixelY))
End If
SetWindowRgn ctlHwnd, hRgn, True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub CutRegion(ctlHwnd As Long, Ctl As ComboBox, bCut As Boolean)"
End Sub

Public Sub RaiseAllStatusbarPanels(lStatusBar As StatusBar)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To lStatusBar.Panels.Count
    lStatusBar.Panels.Item(i).Bevel = sbrNoBevel
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub RaiseAllStatusbarPanels(lStatusbar As StatusBar)"
End Sub

Public Sub LoadServers()
Dim i As Integer, s As Integer, j As Integer, lWord() As String, msg As String
lServers.sNetworkUBound = 0
lServers.sNetworkCount = 0
lServers.sServerCount = 0
lServers.sServerUBound = 0
For i = 0 To 1000
    lServers.sNetwork(i).nDescription = ReadINI(GetINIFile(iServers), "ALLGROUPS", Trim(Str(i)), "")
    If Len(lServers.sNetwork(i).nDescription) <> 0 Then
        lServers.sNetworkUBound = lServers.sNetworkUBound + 1
        lServers.sNetworkCount = lServers.sNetworkCount + 1
    End If
Next i
For s = 0 To lServers.sNetworkCount
    For i = 0 To 1024
        msg = ReadINI(GetINIFile(iServers), lServers.sNetwork(s).nDescription, Str(i), "")
        If Len(msg) = 0 Then
            Exit For
        Else
            lServers.sServerUBound = lServers.sServerUBound + 1
            lServers.sServerCount = lServers.sServerCount + 1
            j = lServers.sServerCount
            lWord = Split(msg, "|")
            lServers.sServer(j).sDescription = lWord(1)
            lServers.sServer(j).sNetwork = s
            lServers.sServer(j).sPortRange = lWord(2)
            lServers.sServer(j).sServer = lWord(0)
        End If
    Next i
Next s
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub LoadServers()"
End Sub

Public Sub SetAboutProgress(lValue As Integer, Optional lText As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'If lSettings.sShowSplashOnStartup = True Then
If Trim(LCase(lText)) = "loading auto connect" And lSettings.sShowSplashOnStartup = True Then
    frmSplash.XP_ProgressBar1.Scrolling = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sProgressBarStyle
End If

frmSplash.AddToInfo lAboutInt, lText, ""
lAboutInt = lAboutInt + 1
frmSplash.XP_ProgressBar1.Value = lValue
'End If
'DoColor lSettings.sActiveServerForm, lText
'Exit Sub
'If lSettings.sShowSplashOnStartup = True Then
'    frmSplash.prgLoading.Value = lValue
'    If Len(lText) <> 0 Then
'        If lValue = 1 Then
'            frmSplash.txtLoading.Text = "• Loading NexIRC Options Please wait ..." & vbCrLf
'        Else
'            frmSplash.lblLoading.Caption = "Loading: " & Format(lValue, " ##") & "%" & " (" & lText & ")": DoEvents
'            If lValue <> 98 And lValue <> 99 Then
'                frmSplash.txtLoading.Text = frmSplash.txtLoading.Text & vbCrLf & "• " & lText
'            End If
'        End If
'        DoEvents
'    End If
'End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetAboutProgress(lValue As Integer, Optional lText As String)"
End Sub

Public Sub SetComboIndex(lCombo As ComboBox, lText As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To lCombo.ListCount
    If Trim(LCase(lCombo.List(i))) = Trim(LCase(lText)) Then
        lCombo.ListIndex = i
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetComboIndex(lCombo As ComboBox, lText As String)"
End Sub

Public Sub LoadScript(lFileName As String, Optional lPreserveFilename As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, F As clsFMenu
If Len(lFileName) <> 0 Then
    If LCase(Right(lFileName, 5)) = ".nirc" Then
        Set F = New clsFMenu
        If lPreserveFilename = False Then
            F.RunScriptFile App.Path & "\data\scripts\" & lFileName
        Else
            F.RunScriptFile lFileName
        End If
        Exit Sub
    End If
    If lPreserveFilename = False Then
        msg = ReadFile(App.Path & "\data\scripts\" & lFileName)
    Else
        msg = ReadFile(lFileName)
    End If
    If Len(msg) <> 0 Then
        'mdiNexIRC.ctlVBScript.ExecuteStatement msg
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub LoadScript(lFilename As String)"
End Sub

Public Sub LoadSettings()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim F As Integer, i As Integer, msg As String, lNow As String
SetAboutProgress 1, "Starting"

SetAboutProgress 2, "Loading Themes"
LoadSpectrumThemes
SetAboutProgress 15, "Loading Auto Connect"
LoadAutoConnect
SetAboutProgress 20, "Loading Custom Strings"
LoadStrings
SetAboutProgress 26, "Loading Playlist"
LoadPlaylist "", False
SetAboutProgress 31, "Loading Servers"
LoadServers
SetAboutProgress 35, "Loading Bots"
LoadBots
SetAboutProgress 36, "Loading Autoperform"
LoadAutoPerform
SetAboutProgress 38, "Loading Alternates"
LoadAlternates
SetAboutProgress 40, "Loading Autojoin"
LoadAutoJoin
SetAboutProgress 42, "Loading Notify"
LoadNotify
SetAboutProgress 44, "Loading Channel Folders"
LoadChanFolders
SetAboutProgress 46, "Loading Audio Settings"
lPlayback.pRepeatCurrentTrack = ReadINI(GetINIFile(iIRC), "Settings", "RepeatCurrentTrack", False)
lSettings.sShowQuickmix = ReadINI(GetINIFile(iIRC), "Settings", "ShowQuickMix", True)
lPlayback.pCurrentEngine = ReadINI(GetINIFile(iIRC), "Settings", "CurrentEngine", pMp3)
lSettings.sBalloons = ReadINI(GetINIFile(iIRC), "Settings", "Balloons", True)
lSettings.sShowWhoisInChannel = ReadINI(GetINIFile(iIRC), "Settings", "ShowWhoisInChannel", False)
lSettings.sLogoTwitchOnPeaks = ReadINI(GetINIFile(iIRC), "Settings", "LogoTwitchOnPeaks", False)
lSettings.sAlwaysShowAudioSettings = ReadINI(GetINIFile(iIRC), "Settings", "AlwaysShowAudioSettings", False)
lSettings.sPlaySounds = ReadINI(GetINIFile(iIRC), "Settings", "PlaySounds", False)
lSettings.sExlusiveToMp3InPlaylist = ReadINI(GetINIFile(iIRC), "Settings", "ExclusiveToMp3InPlaylist", True)
lSettings.sShuffle = ReadINI(GetINIFile(iIRC), "Settings", "Shuffle", True)
lSettings.sContinuousPlay = ReadINI(GetINIFile(iIRC), "Settings", "ContinuousPlay", False)
lSettings.sSearchForMedia = ReadINI(GetINIFile(iIRC), "Settings", "SearchForMedia", True)
SetAboutProgress 48, "Loading Offer Settings"
lSettings.sOfferWhenPlayed = ReadINI(GetINIFile(iIRC), "Settings", "OfferWhenPlayed", False)
lSettings.sAudioServer = ReadINI(GetINIFile(iIRC), "Settings", "AudioServer", False)
lSettings.sFileOfferInChannel = ReadINI(GetINIFile(iIRC), "Settings", "FileOfferInChannel", False)
lSettings.sEnableList = ReadINI(GetINIFile(iIRC), "Settings", "EnableList", False)
lSettings.sShowSmallNetworks = ReadINI(GetINIFile(iIRC), "Settings", "ShowSmallNetworks", True)
SetAboutProgress 50, "Loading Notify Settings"
lSettings.sShowNotifyWindow = ReadINI(GetINIFile(iIRC), "Settings", "ShowNotifyWindow", False)
lSettings.sShowQuickNotify = ReadINI(GetINIFile(iIRC), "Settings", "ShowQuickNotify", False)
lSettings.sOptions.oWhoisNotify = ReadINI(GetINIFile(iIRC), "IRC", "WhoisNotify", False)
SetNotifyEnabled ReadINI(GetINIFile(iIRC), "Notify", "Enabled", False)
lSettings.sOptions.oShowNotifyInActiveWindow = ReadINI(GetINIFile(iIRC), "IRC", "ShowNotifyInActive", True)
SetAboutProgress 52, "Loading DCC Settings"
lSettings.sDownloadManager = ReadINI(GetINIFile(iIRC), "Settings", "DownloadManager", True)
lSettings.sDCCEnabled = ReadINI(GetINIFile(iIRC), "Settings", "DCCEnabled", True)
SetAboutProgress 56, "Loading IRC Settings"
lSettings.sAutoPortScanner = ReadINI(GetINIFile(iIRC), "Settings", "AutoPortScanner", False)
lSettings.sByPassStartupScreen = ReadINI(GetINIFile(iIRC), "Settings", "ByPassStartupScreen", True)
lSettings.sServerMinimum = Int(ReadINI(GetINIFile(iIRC), "Settings", "ServerMinimum", 0))
lSettings.sAutoSelectAlternateNickname = ReadINI(GetINIFile(iIRC), "Settings", "AutoSelectAlternateNickname", True)
lSettings.sColoredNicklist = ReadINI(GetINIFile(iIRC), "Settings", "ColoredNicklist", True)
lSettings.sSecureQuery = ReadINI(GetINIFile(iIRC), "Settings", "SecureQuery", False)
lSettings.sReconnectOnDisconnect = ReadINI(GetINIFile(iIRC), "Settings", "ReconnectOnDisconnect", False)
lSettings.sAutoJoinEnabled = ReadINI(GetINIFile(iIRC), "Settings", "AutojoinEnabled", True)
lSettings.sAddJoinedChannelsToChannelFolder = ReadINI(GetINIFile(iIRC), "Settings", "AddJoinedChannelsToChannelFolder", True)
lSettings.sShowTips = ReadINI(GetINIFile(iIRC), "Settings", "ShowTips", False)
lSettings.sNetwork = ReadINI(GetINIFile(iIRC), "Info", "Network", "Undernet")
lSettings.sServer = ReadINI(GetINIFile(iIRC), "Info", "Server", "graz2.at.eu.undernet.org")
lSettings.sPort = ReadINI(GetINIFile(iIRC), "Info", "PORT", "6660-6669")
lSettings.sPassword = ReadINI(GetINIFile(iIRC), "Info", "PASSWORD", "")
lSettings.sAutoJoinOnInvite = ReadINI(GetINIFile(iIRC), "IRC", "AutoJoinOnInvite", False)
lSettings.sConnectOnStartup = ReadINI(GetINIFile(iIRC), "IRC", "ConnectOnStartup", False)
lSettings.sShowOptionsOnStartup = ReadINI(GetINIFile(iIRC), "IRC", "ShowOptionsOnStartup", True)
SetAboutProgress 60, "Loading Theme Settings"
lSettings.sRefreshPictureColors = ReadINI(GetINIFile(iIRC), "Settings", "RefreshPictureColors", False)
lSettings.sBGColor = ReadINI(GetINIFile(iIRC), "Settings", "BGColor", 0)
lSettings.sApplyThemeToIRCColors = ReadINI(GetINIFile(iIRC), "Settings", "ApplyThemeToIRCColors", True)
lSettings.sBGPicture = ReadINI(GetINIFile(iIRC), "Settings", "BGPicture", "")
lSettings.sSaveIRCColorsToTheme = ReadINI(GetINIFile(iIRC), "Settings", "SaveIRCColorsToTheme", True)
SetAboutProgress 64, "Loading Startup Settings"
lSettings.sShowServerOnStartup = ReadINI(GetINIFile(iIRC), "Settings", "ShowServerOnStartup", False)
SetAboutProgress 68, "Loading Interface Settings"
lSettings.sUseNickCompletor = ReadINI(GetINIFile(iIRC), "Settings", "UseNickCompletor", True)
lSettings.sAutosizeStatusbarItems = ReadINI(GetINIFile(iIRC), "Settings", "AutosizeStatusbarItems", False)
lSettings.sBorderlessObjects = ReadINI(GetINIFile(iIRC), "Settings", "BorderlessObjects", True)
SetAboutProgress 76, "Loading Browser Settings"
lSettings.sBackgroundWebpage = ReadINI(GetINIFile(iIRC), "Settings", "BackgroundWebpage", True)
lSettings.sNavigateOnStartup = ReadINI(GetINIFile(iIRC), "Settings", "NavigateOnStartup", True)
SetAboutProgress 80, "Loading Identd Settings"
lSettings.sIdent.iEnabled = ReadINI(GetINIFile(iIRC), "Ident", "Ident", True)
lSettings.sIdent.iPort = ReadINI(GetINIFile(iIRC), "Ident", "Port", "113")
lSettings.sIdent.iShow = ReadINI(GetINIFile(iIRC), "Ident", "Show", True)
lSettings.sIdent.iUserID = ReadINI(GetINIFile(iIRC), "Ident", "UserID", "NEXIRC")
lSettings.sIdent.iSystem = ReadINI(GetINIFile(iIRC), "Ident", "System", "WIN32")
SetAboutProgress 84, "Loading Options"
lSettings.sOptions.oShowChannelFolder = ReadINI(GetINIFile(iIRC), "Settings", "ShowChannelFolder", False)
lSettings.sOptions.oWhois = ReadINI(GetINIFile(iIRC), "IRC", "Whois", True)
lSettings.sOptions.oReJoin = ReadINI(GetINIFile(iIRC), "IRC", "Rejoin", True)
lSettings.sOptions.oSkipMOTD = ReadINI(GetINIFile(iIRC), "IRC", "SkipMOTD", False)
lSettings.sOptions.oShowMOTD = ReadINI(GetINIFile(iIRC), "IRC", "ShowMOTD", False)
lSettings.sOptions.oShowAddress = ReadINI(GetINIFile(iIRC), "Show", "Address", True)
lSettings.sOptions.oShowAddress = ReadINI(GetINIFile(iIRC), "IRC", "ShowAddress", True)
lSettings.sOptions.oInvisable = ReadINI(GetINIFile(iIRC), "IRC", "SetInvisible", False)
SetAboutProgress 88, "Loading Channel Settings"
lSettings.sOptions.oServerMessages = ReadINI(GetINIFile(iIRC), "Show", "ServerMessages", False)
lSettings.sOptions.oOpMessages = ReadINI(GetINIFile(iIRC), "Show", "Wallops", False)
lSettings.sOptions.oShowQuit = ReadINI(GetINIFile(iIRC), "Show", "Quit", True)
lSettings.sOptions.oShowJoinPart = ReadINI(GetINIFile(iIRC), "Show", "JoinPart", True)
lSettings.sOptions.oShowModes = ReadINI(GetINIFile(iIRC), "Show", "Modes", True)
lSettings.sOptions.oShowTopics = ReadINI(GetINIFile(iIRC), "Show", "Topics", True)
lSettings.sOptions.oShowKicks = ReadINI(GetINIFile(iIRC), "Show", "Kicks", True)
lSettings.sHomepage = ReadINI(GetINIFile(iIRC), "Settings", "Homepage", "http://www.team-nexgen.org")
SetAboutProgress 92, "Loading Initial Audio Values"
lInitialAudioValues.iBass = ReadINI(GetINIFile(iInitialValues), "Settings", "Bass", 0)
lInitialAudioValues.iCDAudio = ReadINI(GetINIFile(iInitialValues), "Settings", "CDAudio", 0)
lInitialAudioValues.iLineIN = ReadINI(GetINIFile(iInitialValues), "Settings", "LineIN", 0)
lInitialAudioValues.iMic = ReadINI(GetINIFile(iInitialValues), "Settings", "Mic", 0)
lInitialAudioValues.iTreble = ReadINI(GetINIFile(iInitialValues), "Settings", "Treble", 0)
lInitialAudioValues.iWave = ReadINI(GetINIFile(iInitialValues), "Settings", "Wave", 0)
lInitialAudioValues.iInitialWaveEnabled = ReadINI(GetINIFile(iInitialValues), "Settings", "InitialWaveEnabled", False)
lInitialAudioValues.iInitialBassEnabled = ReadINI(GetINIFile(iInitialValues), "Settings", "InitialBassEnabled", False)
lInitialAudioValues.iInitialCDAudioEnabled = ReadINI(GetINIFile(iInitialValues), "Settings", "InitialCDAudioEnabled", False)
lInitialAudioValues.iInitialLineInEnabled = ReadINI(GetINIFile(iInitialValues), "Settings", "InitialLineInEnabled", False)
lInitialAudioValues.iInitialMicEnabled = ReadINI(GetINIFile(iInitialValues), "Settings", "InitialMicEnabled", False)
lInitialAudioValues.iInitialTrebleEnabled = ReadINI(GetINIFile(iInitialValues), "Settings", "InitialTrebleEnabled", False)
SetAboutProgress 94, "Loading Window Settings"
lSettings.sLastWindowPos.lParentForm.fHeight = ReadINI(GetINIFile(iIRC), mdiNexIRC.Name, "HEIGHT", 8880)
lSettings.sLastWindowPos.lParentForm.fWidth = ReadINI(GetINIFile(iIRC), mdiNexIRC.Name, "WIDTH", 11280)
lSettings.sLastWindowPos.lParentForm.fLeft = ReadINI(GetINIFile(iIRC), mdiNexIRC.Name, "LEFT", 0)
lSettings.sLastWindowPos.lParentForm.fTop = ReadINI(GetINIFile(iIRC), mdiNexIRC.Name, "TOP", 0)
SetAboutProgress 95, "Loading Ignore"
LoadIgnore
SetAboutProgress 96, "Loading Blacklist"
LoadBlacklist
lSettings.sColors = Trim(ReadINI(GetINIFile(iIRC), "IRC", "COLORS", "0:01 1:00 2:09 3:09 4:15 5:03 6:03 7:07 8:08 9:00 10:10 11:08 12:08 13:08 14:10 15:14"))
SetAboutProgress 100, "Complete"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub LoadSettings()"
End Sub

Public Sub SendUserPlaylist(lUserName As String, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
mdiNexIRC.SetEUsername lUserName, lForm
mdiNexIRC.tmrSendUserPlaylist.Enabled = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SendUserPlaylist(lUsername As String)"
End Sub

Public Sub DisableIdent()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
mdiNexIRC.wskIdent.Close
ProcessReplaceString sIdentDDisabled, lSettings.sActiveServerForm.txtIncoming
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub DisableIdent()"
End Sub

Public Sub EnableIdent(lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sIdent.iEnabled = True Then
    mdiNexIRC.wskIdent.Close
    If mdiNexIRC.wskIdent.State <> sckListening Then
        mdiNexIRC.wskIdent.LocalPort = Val(lSettings.sIdent.iPort)
        mdiNexIRC.wskIdent.RemoteHost = lSettings.sServer
        mdiNexIRC.wskIdent.Listen
        If mdiNexIRC.wskIdent.State = sckListening Then
            ProcessReplaceString sIdentListening, lForm.txtIncoming, mdiNexIRC.wskIdent.LocalIP, mdiNexIRC.wskIdent.LocalPort
        Else
            ProcessReplaceString sIdentListenFailed, lForm.txtIncoming, mdiNexIRC.wskIdent.LocalIP, mdiNexIRC.wskIdent.LocalPort
        End If
    Else
        ProcessReplaceString sIdentListenFailed, lForm.txtIncoming, mdiNexIRC.wskIdent.LocalIP, mdiNexIRC.wskIdent.LocalPort
    End If
Else
    mdiNexIRC.wskIdent.Close
    ProcessReplaceString sIdentListenFailed, lForm.txtIncoming, mdiNexIRC.wskIdent.LocalIP, mdiNexIRC.wskIdent.LocalPort
End If
If Err.Number = 10048 Then
    ProcessReplaceString sIdentDDisabled, lSettings.sActiveServerForm.txtIncoming
    Err = 0
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub LoadSettings()"
End Sub

Public Sub NewQuery(lNickName As String, Optional lNormalFocus As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim ChatWith As String, i As Integer, xlfound As Boolean
xlfound = False
ChatWith = lNickName
ChatWith = Replace(ChatWith, "@", "")
ChatWith = Replace(ChatWith, "+", "")
ChatWith = Replace(ChatWith, "%", "")
For i = 1 To 150
    If LCase(ReturnQueryName(i)) = LCase(ChatWith) Then
        xlfound = True
        Exit For
    End If
Next i
If xlfound = False Then
    For i = 1 To 150
        If ReturnQueryName(i) = "" Then
        'If lQueryName(i) = "" Then
            'Load lQuery(i)
            'lQuery(i).Caption = ChatWith
            'lQueryName(i) = ChatWith
            'Call AddTaskPanel(ChatWith, 1)
            'lQuery(i).Visible = True
            'If lNormalFocus = True Then
            '    lQuery(i).WindowState = vbNormal
            'End If
            LoadQueryWindow i, ChatWith, ""
            Exit For
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub NewQuery(lNickname As String)"
End Sub

Public Sub ResetMainButtons()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
With mdiNexIRC
    If .picForward.Picture <> frmGraphics.picForward1.Picture Then
        If lSettings.sRefreshPictureColors = True Then
            mdiNexIRC.picForward.Picture = frmGraphics.picForward1.Picture
            SetPictureColor .picForward, lRedColor, lBlueColor, lGreenColor, True
        Else
            mdiNexIRC.picForward.Picture = frmGraphics.picForward1.Picture
        End If
    End If
    If .picPlay.Picture <> frmGraphics.picPlay1.Picture Then
        If lSettings.sRefreshPictureColors = True Then
            mdiNexIRC.picPlay.Picture = frmGraphics.picPlay1.Picture
            SetPictureColor .picPlay, lRedColor, lBlueColor, lGreenColor, True
        Else
            mdiNexIRC.picPlay.Picture = frmGraphics.picPlay1.Picture
        End If
    End If
    If .picPause.Picture <> frmGraphics.picPause1.Picture Then
        If lSettings.sRefreshPictureColors = True Then
            mdiNexIRC.picPause.Picture = frmGraphics.picPause1.Picture
            SetPictureColor .picPause, lRedColor, lBlueColor, lGreenColor, True
        Else
            mdiNexIRC.picPause.Picture = frmGraphics.picPause1.Picture
        End If
    End If
    If .picStop.Picture <> frmGraphics.picStop1.Picture Then
        If lSettings.sRefreshPictureColors = True Then
            mdiNexIRC.picStop.Picture = frmGraphics.picStop1.Picture
            SetPictureColor .picStop, lRedColor, lBlueColor, lGreenColor, True
        Else
            mdiNexIRC.picStop.Picture = frmGraphics.picStop1.Picture
        End If
    End If
    If .picBackward.Picture <> frmGraphics.picBackward1.Picture Then
        If lSettings.sRefreshPictureColors = True Then
            mdiNexIRC.picBackward.Picture = frmGraphics.picBackward1.Picture
            SetPictureColor .picChat, lRedColor, lBlueColor, lGreenColor, True
        Else
            mdiNexIRC.picBackward.Picture = frmGraphics.picBackward1.Picture
        End If
    End If
    If .picChat.Picture <> frmGraphics.picChat1.Picture Then
        If lSettings.sRefreshPictureColors = True Then
            mdiNexIRC.picChat.Picture = frmGraphics.picChat1.Picture
            SetPictureColor mdiNexIRC.picChat, lRedColor, lBlueColor, lGreenColor, True
        Else
            mdiNexIRC.picChat.Picture = frmGraphics.picChat1.Picture
        End If
    End If
    If .picExit.Picture <> frmGraphics.picExit1.Picture Then
        If lSettings.sRefreshPictureColors = True Then
            mdiNexIRC.picExit.Picture = frmGraphics.picExit1.Picture
            SetPictureColor mdiNexIRC.picExit, lRedColor, lBlueColor, lGreenColor, True
        Else
            mdiNexIRC.picExit.Picture = frmGraphics.picExit1.Picture
        End If
    End If
    If .picNexIRC.Picture <> frmGraphics.picNexIRC1.Picture Then
        If lSettings.sRefreshPictureColors = True Then
            mdiNexIRC.picNexIRC.Picture = frmGraphics.picNexIRC1.Picture
            SetPictureColor mdiNexIRC.picNexIRC, lRedColor, lBlueColor, lGreenColor, True
        Else
            mdiNexIRC.picNexIRC.Picture = frmGraphics.picNexIRC1.Picture
        End If
    End If
    If .picSend.Picture <> frmGraphics.picSend1.Picture Then
        If lSettings.sRefreshPictureColors = True Then
            mdiNexIRC.picSend.Picture = frmGraphics.picSend1.Picture
            SetPictureColor mdiNexIRC.picSend, lRedColor, lBlueColor, lGreenColor, True
        Else
            mdiNexIRC.picSend.Picture = frmGraphics.picSend1.Picture
        End If
    End If
    If .picDisconnect.Picture <> frmGraphics.picDisconnect1.Picture Then
        If lSettings.sRefreshPictureColors = True Then
            mdiNexIRC.picDisconnect.Picture = frmGraphics.picDisconnect1.Picture
            SetPictureColor mdiNexIRC.picDisconnect, lRedColor, lBlueColor, lGreenColor, True
        Else
            mdiNexIRC.picDisconnect.Picture = frmGraphics.picDisconnect1.Picture
        End If
    End If
    If .picConnect.Picture <> frmGraphics.picConnect1.Picture Then
        If lSettings.sRefreshPictureColors = True Then
            mdiNexIRC.picConnect.Picture = frmGraphics.picConnect1.Picture
            SetPictureColor mdiNexIRC.picConnect, lRedColor, lBlueColor, lGreenColor, True
        Else
            mdiNexIRC.picConnect.Picture = frmGraphics.picConnect1.Picture
        End If
    End If
    If .picAudio.Picture <> frmGraphics.picAudio1.Picture Then
        If lSettings.sRefreshPictureColors = True Then
            mdiNexIRC.picAudio.Picture = frmGraphics.picAudio1.Picture
            SetPictureColor mdiNexIRC.picAudio, lRedColor, lBlueColor, lGreenColor, True
        Else
            mdiNexIRC.picAudio.Picture = frmGraphics.picAudio1.Picture
        End If
    End If
    If .picOptions.Picture <> frmGraphics.picOptions1.Picture Then
        If lSettings.sRefreshPictureColors = True Then
            mdiNexIRC.picOptions.Picture = frmGraphics.picOptions1.Picture
            SetPictureColor mdiNexIRC.picOptions, lRedColor, lBlueColor, lGreenColor, True
        Else
            mdiNexIRC.picOptions.Picture = frmGraphics.picOptions1.Picture
        End If
    End If
    If .picChannelFolder.Picture <> frmGraphics.picChannelFolder1.Picture Then
        If lSettings.sRefreshPictureColors = True Then
            mdiNexIRC.picChannelFolder.Picture = frmGraphics.picChannelFolder1.Picture
            SetPictureColor mdiNexIRC.picChannelFolder, lRedColor, lBlueColor, lGreenColor, True
        Else
            mdiNexIRC.picChannelFolder.Picture = frmGraphics.picChannelFolder1.Picture
        End If
    End If
    If .picScript.Picture <> frmGraphics.picScript1.Picture Then
        If lSettings.sRefreshPictureColors = True Then
            mdiNexIRC.picScript.Picture = frmGraphics.picScript1.Picture
            SetPictureColor mdiNexIRC.picScript, lRedColor, lBlueColor, lGreenColor, True
        Else
            mdiNexIRC.picScript.Picture = frmGraphics.picScript1.Picture
        End If
    End If
End With
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ResetMainButtons()"
End Sub
