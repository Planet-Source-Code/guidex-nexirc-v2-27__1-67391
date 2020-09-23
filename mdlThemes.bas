Attribute VB_Name = "mdlThemes"
Option Explicit
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Type gButtons
    bConnect(3) As StdPicture
    bConnectLeft As Integer
    bConnectTop As Integer
    bDisconnect(3) As StdPicture
    bDisconnectTop As Integer
    bDisconnectLeft As Integer
    bAudio(3) As StdPicture
    bAudioLeft As Integer
    bAudioTop As Integer
    bOptions(3) As StdPicture
    bOptionsLeft As Integer
    bOptionsTop As Integer
    bChannelFolder(3) As StdPicture
    bChannelFolderLeft As Integer
    bChannelFolderTop As Integer
    bSend(3) As StdPicture
    bSendTop As Integer
    bSendLeft As Integer
    bScript(3) As StdPicture
    bScriptLeft As Integer
    bScriptTop As Integer
    bNexIRC(3) As StdPicture
    bNexIRCTop As Integer
    bNexIRCLeft As Integer
    bBackward(3) As StdPicture
    bBackwardTop As Integer
    bBackwardLeft As Integer
    bStop(3) As StdPicture
    bStopLeft As Integer
    bStopTop As Integer
    bForward(3) As StdPicture
    bForwardLeft As Integer
    bForwardTop As Integer
    bPause(3) As StdPicture
    bPauseLeft As Integer
    bPauseTop As Integer
    bPlay(3) As StdPicture
    bPlayLeft As Integer
    bPlayTop As Integer
    bExit(3) As StdPicture
    bExitTop As Integer
    bExitLeft As Integer
    bctlChat(3) As StdPicture
    bChatTop As Integer
    bChatLeft As Integer
End Type
Private Type gNicklist
    nOpColor As String
    nVoiceColor As String
    nNormalColor As String
    nOpPicIndex As Integer
    nVoicePicIndex As Integer
End Type
Private Type gSpectrumTheme
    sNicklistOptions As gNicklist
    sScreenShot As String
    sBottomBandsColor As String
    sTopBandsColor As String
    sDividerColor As String
    sLeftChanColor As String
    sPeaksColor As String
    sRightChanColor As String
    sSpectrumBackcolor As String
    sBackColor As String
    sTextColor As String
    sBottomToolbarColor As String
    sURLLeft As Integer
    sURLTop As Integer
    sURLWidth As Integer
    sURLHeight As Integer
    sDisableToolbarColors As Boolean
    sIRCColors As String
    sSpectrumWidth As Integer
    sSpectrumHeight As Integer
    sSpectrumLeft As Integer
    sSpectrumTop As Integer
    sName As String
    sBands As Integer
    sToolbarGraphic As String
    sRed As Long
    sGreen As Long
    sBlue As Long
    sButtonType As Integer
    sMode As Integer
    sButtons As gButtons
    sFontname As String
    sFontsize As Integer
    sOscilloType As Integer
    sBGTextColor As Integer
    sProgressBarStyle As Integer
    sProgressBarColor As Integer
End Type
Private Type gSpectrumThemes
    sHalfLogo As String
    sStartupGraphic As String
    sSpectrumTheme(150) As gSpectrumTheme
    sCount As Integer
    sIndex As Integer
End Type
Private Type gColors
    BGText As Integer
    Normal As Integer
    CTCP As Integer
    Notice As Integer
    Action As Integer
    Invite As Integer
    Join As Integer
    Kick As Integer
    Mode As Integer
    Nick As Integer
    Part As Integer
    Quit As Integer
    Topic As Integer
    Whois As Integer
    Server As Integer
    Notify As Integer
End Type
Global Color As gColors
Global lSpectrumThemes As gSpectrumThemes

Public Sub SaveSpectrumTheme(lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim s As String, i As Integer
s = GetINIFile(iSpectrum)
i = lIndex
With lSpectrumThemes.sSpectrumTheme(lIndex)
    WriteINI s, Trim(Str(i)), "BGTextColor", .sBGTextColor
    WriteINI s, Trim(Str(i)), "BackColor", .sBackColor
    WriteINI s, Trim(Str(i)), "Bands", .sBands
    WriteINI s, Trim(Str(i)), "Blue", .sBlue
    WriteINI s, Trim(Str(i)), "BottomBandsColor", .sBottomBandsColor
    WriteINI s, Trim(Str(i)), "BottomToolbarColor", .sBottomToolbarColor
    WriteINI s, Trim(Str(i)), "ButtonType", Str(.sButtonType)
    WriteINI s, Trim(Str(i)), "DisableToolbarColors", .sDisableToolbarColors
    WriteINI s, Trim(Str(i)), "DividerColor", .sDividerColor
    WriteINI s, Trim(Str(i)), "Fontname", .sFontname
    WriteINI s, Trim(Str(i)), "Fontsize", .sFontsize
    WriteINI s, Trim(Str(i)), "Green", .sGreen
    WriteINI s, Trim(Str(i)), "IRCColors", Trim(.sIRCColors)
    WriteINI s, Trim(Str(i)), "LeftChanColor", .sLeftChanColor
    WriteINI s, Trim(Str(i)), "Mode", .sMode
    WriteINI s, Trim(Str(i)), "Name", .sName
    WriteINI s, Trim(Str(i)), "OscilloType", .sOscilloType
    WriteINI s, Trim(Str(i)), "PeaksColor", .sPeaksColor
    WriteINI s, Trim(Str(i)), "Red", .sRed
    WriteINI s, Trim(Str(i)), "RightChanColor", .sRightChanColor
    WriteINI s, Trim(Str(i)), "SpectrumBackcolor", .sSpectrumBackcolor
    WriteINI s, Trim(Str(i)), "SpectrumHeight", .sSpectrumHeight
    WriteINI s, Trim(Str(i)), "SpectrumWidth", .sSpectrumWidth
    WriteINI s, Trim(Str(i)), "SpectrumTop", .sSpectrumTop
    WriteINI s, Trim(Str(i)), "SpectrumLeft", .sSpectrumLeft
    WriteINI s, Trim(Str(i)), "TextColor", .sTextColor
    WriteINI s, Trim(Str(i)), "ToolbarGraphic", .sToolbarGraphic
    WriteINI s, Trim(Str(i)), "TopBandsColor", .sTopBandsColor
    WriteINI s, Trim(Str(i)), "URLHeight", .sURLHeight
    WriteINI s, Trim(Str(i)), "URLWidth", .sURLWidth
    WriteINI s, Trim(Str(i)), "URLLeft", .sURLLeft
    WriteINI s, Trim(Str(i)), "URLTop", .sURLTop
End With
SaveNicklistOptions lIndex
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SaveSpectrumTheme(lIndex As Integer)"
End Sub

Public Sub SaveNicklistOptions(lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim s As String, i As Integer
s = GetINIFile(iSpectrum)
i = lIndex
With lSpectrumThemes.sSpectrumTheme(lIndex).sNicklistOptions
    WriteINI s, Trim(Str(i)), "NormalColor", .nNormalColor
    WriteINI s, Trim(Str(i)), "OpColor", .nOpColor
    WriteINI s, Trim(Str(i)), "VoiceColor", .nVoiceColor
    WriteINI s, Trim(Str(i)), "OpPicIndex", Trim(Str(.nOpPicIndex))
    WriteINI s, Trim(Str(i)), "VoicePicIndex", .nVoicePicIndex
End With
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SaveNicklistOptions()"
End Sub

Public Sub NewSpectrumTheme()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer, F As Integer, G As Boolean, m As Integer
For i = 1 To 100
    msg = "New Spectrum Theme " & Trim(Str(i))
    If Len(msg) <> 0 Then
        F = FindSpectrumThemeByName(msg)
        If F = 0 Then
            G = True
            Exit For
        End If
    End If
Next i
If G = True Then
    msg2 = InputBox("Please enter the name of your spectrum theme", "Spectrum Themes", msg)
Else
    msg2 = InputBox("Please enter the name of your spectrum theme", "Spectrum Themes", "")
End If
m = AddSpectrumTheme(msg2, "&H80000005", "&H8000000F", "0", "&H8000000D", "&H8000000C", "&H8000000D", "&H8000000D", 20, "xp\toolbar.gif", "0", "0:00 1:01 2:04 3:05 4:06 5:07 6:03 7:03 8:03 9:03 10:07 11:03 12:10 13:03 14:07 15:10", "0", "Tahoma", 8, 2, False, 0, 0, 100)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub NewSpectrumTheme()"
End Sub

Public Sub SetPictureColors(lRed As Long, lBlue As Long, lGreen As Long, Optional lSetColorOnly As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sDisableToolbarColors = True Then Exit Sub
If lSettings.sRefreshPictureColors = False Then Exit Sub
If lSetColorOnly = False Then
    mdiNexIRC.picConnect.Picture = frmGraphics.picConnect1.Picture
    mdiNexIRC.picDisconnect.Picture = frmGraphics.picDisconnect1.Picture
    mdiNexIRC.picAudio.Picture = frmGraphics.picAudio1.Picture
    mdiNexIRC.picOptions.Picture = frmGraphics.picOptions1.Picture
    mdiNexIRC.picChannelFolder.Picture = frmGraphics.picChannelFolder1.Picture
    mdiNexIRC.picSend.Picture = frmGraphics.picSend1.Picture
    mdiNexIRC.picNexIRC.Picture = frmGraphics.picNexIRC1.Picture
    mdiNexIRC.picScript.Picture = frmGraphics.picScript1.Picture
    mdiNexIRC.picPlay.Picture = frmGraphics.picPlay1.Picture
    mdiNexIRC.picPause.Picture = frmGraphics.picPause1.Picture
    mdiNexIRC.picStop.Picture = frmGraphics.picStop1.Picture
    mdiNexIRC.picForward.Picture = frmGraphics.picForward1.Picture
    mdiNexIRC.picBackward.Picture = frmGraphics.picBackward1.Picture
    DoEvents
End If
SetPictureColor mdiNexIRC.picTopToolbar, lRed, lBlue, lGreen, True
SetPictureColor mdiNexIRC.picConnect, lRed, lBlue, lGreen, True
SetPictureColor mdiNexIRC.picDisconnect, lRed, lBlue, lGreen, True
SetPictureColor mdiNexIRC.picAudio, lRed, lBlue, lGreen, True
SetPictureColor mdiNexIRC.picOptions, lRed, lBlue, lGreen, True
SetPictureColor mdiNexIRC.picChannelFolder, lRed, lBlue, lGreen, True
SetPictureColor mdiNexIRC.picSend, lRed, lBlue, lGreen, True
SetPictureColor mdiNexIRC.picNexIRC, lRed, lBlue, lGreen, True
SetPictureColor mdiNexIRC.picScript, lRed, lBlue, lGreen, True
SetPictureColor mdiNexIRC.picPlay, lRed, lBlue, lGreen, True
SetPictureColor mdiNexIRC.picPause, lRed, lBlue, lGreen, True
SetPictureColor mdiNexIRC.picForward, lRed, lBlue, lGreen, True
SetPictureColor mdiNexIRC.picBackward, lRed, lBlue, lGreen, True
SetPictureColor mdiNexIRC.picStop, lRed, lBlue, lGreen, True
DoEvents
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetPictureColors(lRed As Long, lBlue As Long, lGreen As Long, Optional lSetColorOnly As Boolean)"
End Sub

Public Function AddSpectrumTheme(lName As String, lBackColor As String, lBottomBandsColor As String, lDividerColor As String, lLeftChanColor As String, lPeaksColor As String, lRightChanColor As String, lTopBandsColor As String, lBands As Integer, lToolbarGraphic As String, lTextColor As String, Optional lIRCColors As String, Optional lMode As Integer, Optional lFontname As String, Optional lFontsize As Integer, Optional lButtonType As Integer, Optional lDisableToolbarColors As Boolean, Optional lRed As Integer, Optional lGreen As Integer, Optional lBlue As Integer, Optional lOscilloType As Integer, Optional lBottomToolbarColor As String, Optional lBGTextColor As Integer) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If FindSpectrumThemeByName(lName) <> 0 Then
    AddSpectrumTheme = FindSpectrumThemeByName(lName)
    Exit Function
End If
With lSpectrumThemes
    If Len(lName) <> 0 Then
        .sCount = (.sCount + 1)
        .sSpectrumTheme(.sCount).sBackColor = lBackColor
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "BackColor", lBackColor
        .sSpectrumTheme(.sCount).sBGTextColor = lBGTextColor
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "BGTextColor", lBGTextColor
        .sSpectrumTheme(.sCount).sBands = lBands
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "Bands", lBands
        .sSpectrumTheme(.sCount).sBottomBandsColor = lBottomBandsColor
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "BottomBandsColor", lBottomBandsColor
        .sSpectrumTheme(.sCount).sBottomToolbarColor = lBottomToolbarColor
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "BottomToolbarColor", lBottomToolbarColor
        .sSpectrumTheme(.sCount).sBlue = lBlue
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "Blue", lBlue
        .sSpectrumTheme(.sCount).sButtonType = lButtonType
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "ButtonType", lButtonType
        .sSpectrumTheme(.sCount).sDisableToolbarColors = lDisableToolbarColors
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "DisableToolbarColors", lDisableToolbarColors
        .sSpectrumTheme(.sCount).sDividerColor = lDividerColor
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "DividerColor", lDividerColor
        .sSpectrumTheme(.sCount).sFontname = lFontname
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "Fontname", lFontname
        .sSpectrumTheme(.sCount).sFontsize = lFontsize
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "Fontsize", lFontsize
        .sSpectrumTheme(.sCount).sGreen = lGreen
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "Green", lGreen
        .sSpectrumTheme(.sCount).sIRCColors = Trim(lIRCColors)
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "IRCColors", Trim(lIRCColors)
        .sSpectrumTheme(.sCount).sLeftChanColor = lLeftChanColor
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "LeftChanColor", lLeftChanColor
        .sSpectrumTheme(.sCount).sMode = lMode
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "Mode", lMode
        .sSpectrumTheme(.sCount).sName = lName
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "Name", lName
        .sSpectrumTheme(.sCount).sOscilloType = lOscilloType
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "OscilloType", lOscilloType
        .sSpectrumTheme(.sCount).sRed = lRed
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "Red", lRed
        .sSpectrumTheme(.sCount).sRightChanColor = lRightChanColor
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "RightChanColor", lRightChanColor
        .sSpectrumTheme(.sCount).sPeaksColor = lPeaksColor
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "PeaksColor", lPeaksColor
        .sSpectrumTheme(.sCount).sTextColor = lTextColor
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "TextColor", lTextColor
        .sSpectrumTheme(.sCount).sToolbarGraphic = lToolbarGraphic
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "ToolbarGraphic", lToolbarGraphic
        .sSpectrumTheme(.sCount).sTopBandsColor = lTopBandsColor
        WriteINI GetINIFile(iSpectrum), Str(.sCount), "TopBandsColor", lTopBandsColor
        WriteINI GetINIFile(iSpectrum), "Settings", "Count", Trim(Str(.sCount))
        mdiNexIRC.cboSpectrumThemes.AddItem lName
        AddSpectrumTheme = .sCount
    End If
End With
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function AddSpectrumTheme(lName As String, lBackColor As String, lBottomBandsColor As String, lDividerColor As String, lLeftChanColor As String, lPeaksColor As String, lRightChanColor As String, lTopBandsColor As String, lBands As Integer, lToolbarGraphic As String, lTextColor As String, lIRCColors As String) As Integer"
End Function

Public Sub ApplySpectrumTheme(lName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, F As Integer
i = FindSpectrumThemeByName(lName)
If i <> 0 Then
    If lSpectrumThemes.sSpectrumTheme(i).sBackColor = "-214748363" Then lSpectrumThemes.sSpectrumTheme(i).sBackColor = &H8000000F
    If lSpectrumThemes.sSpectrumTheme(i).sBottomBandsColor = "-214748363" Then lSpectrumThemes.sSpectrumTheme(i).sBottomBandsColor = &H8000000F
    If lSpectrumThemes.sSpectrumTheme(i).sBottomToolbarColor = "-214748363" Then lSpectrumThemes.sSpectrumTheme(i).sBottomToolbarColor = &H8000000F
    If lSpectrumThemes.sSpectrumTheme(i).sDividerColor = "-214748363" Then lSpectrumThemes.sSpectrumTheme(i).sDividerColor = &H8000000F
    If lSpectrumThemes.sSpectrumTheme(i).sLeftChanColor = "-214748363" Then lSpectrumThemes.sSpectrumTheme(i).sLeftChanColor = &H8000000F
    If lSpectrumThemes.sSpectrumTheme(i).sPeaksColor = "-214748363" Then lSpectrumThemes.sSpectrumTheme(i).sPeaksColor = &H8000000F
    If lSpectrumThemes.sSpectrumTheme(i).sRightChanColor = "-214748363" Then lSpectrumThemes.sSpectrumTheme(i).sRightChanColor = &H8000000F
    If lSpectrumThemes.sSpectrumTheme(i).sSpectrumBackcolor = "-214748363" Then lSpectrumThemes.sSpectrumTheme(i).sSpectrumBackcolor = &H8000000F
    If lSpectrumThemes.sSpectrumTheme(i).sTextColor = "-214748363" Then lSpectrumThemes.sSpectrumTheme(i).sTextColor = &H8000000F
    If lSpectrumThemes.sSpectrumTheme(i).sTopBandsColor = "-214748363" Then lSpectrumThemes.sSpectrumTheme(i).sTopBandsColor = &H8000000F
    For F = 0 To ReturnStatusWindowCount
        If Len(ReturnStatusWindowServer(F)) <> 0 Then SetStatusWindowColors F, lSpectrumThemes.sSpectrumTheme(i).sBackColor, lSpectrumThemes.sSpectrumTheme(i).sTextColor
    Next F
    For F = 1 To 150
        If Len(ReturnQueryName(F)) <> 0 Then SetQueryWindowColors F, lSpectrumThemes.sSpectrumTheme(i).sBackColor, lSpectrumThemes.sSpectrumTheme(i).sTextColor
    Next F
    For F = 1 To lSettings.sChannelCount
        If lSettings.sChannelCount <> 0 And LCase(ReturnChannelCaption(F)) <> "nexirc - channel" Then SetChannelWindowColors F, lSpectrumThemes.sSpectrumTheme(i).sBackColor, lSpectrumThemes.sSpectrumTheme(i).sTextColor
    Next F
    If lSettings.sAddMediaVisible = True Then
        frmAddMedia.ctlDir.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBackColor
        frmAddMedia.ctlDir.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
        frmAddMedia.ctlDrive.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBackColor
        frmAddMedia.ctlDrive.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
        frmAddMedia.ctlFiles.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBackColor
        frmAddMedia.ctlFiles.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
    End If
    If lSettings.sChannelListVisible = True Then
        frmChannels.lvwChannels.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBackColor
        frmChannels.lvwChannels.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
    End If
    If lSettings.sIRCServerVisible = True Then
        frmIRCServer.txtOutgoing.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBackColor
        frmIRCServer.txtOutgoing.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
        frmIRCServer.txtIncoming.SetBackColor lSpectrumThemes.sSpectrumTheme(i).sBackColor
        frmIRCServer.lstUsers.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBackColor
        frmIRCServer.lstUsers.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
    End If
    If lSettings.sNotifyVisible = True Then
        frmNotify.lstNotify.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBackColor
        frmNotify.lstNotify.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
    End If
    If lSettings.sConnectionManagerVisible = True Then
        frmConnectionManager.lstConnections.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBackColor
        frmConnectionManager.lstConnections.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
    End If
    If lSettings.sPlaylistVisible = True Then
        frmPlaylist.lstPlaylist.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBackColor
        frmPlaylist.lstPlaylist.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
    End If
    If lSettings.sMOTDVisible = True Then
        frmMOTD.txtIncoming.SetBackColor lSpectrumThemes.sSpectrumTheme(i).sBackColor
    End If
    'With mdiNexIRC.ctlMP3OCX
    '    .BackColor = lSpectrumThemes.sSpectrumTheme(i).sSpectrumBackcolor
    '    If Len(lSpectrumThemes.sSpectrumTheme(i).sBottomBandsColor) <> 0 Then .BottomBandsColor = lSpectrumThemes.sSpectrumTheme(i).sBottomBandsColor
    '    .DividerColor = lSpectrumThemes.sSpectrumTheme(i).sDividerColor
    '    .PeaksColor = lSpectrumThemes.sSpectrumTheme(i).sPeaksColor
    '    .TopBandsColor = lSpectrumThemes.sSpectrumTheme(i).sTopBandsColor
    '    .RightChanColor = lSpectrumThemes.sSpectrumTheme(i).sRightChanColor
    '    .LeftChanColor = lSpectrumThemes.sSpectrumTheme(i).sLeftChanColor
    '    .Bands = lSpectrumThemes.sSpectrumTheme(i).sBands
    '    .SpectrumMode = lSpectrumThemes.sSpectrumTheme(i).sMode
    '    Select Case lSpectrumThemes.sSpectrumTheme(i).sOscilloType
    '    Case 0
    '        .OscilloType = otNone
    '    Case 1
    '        .OscilloType = otWave
    '    Case 2
    '        .OscilloType = otSpectrum
    '    End Select
    'End With
    lRedColor = lSpectrumThemes.sSpectrumTheme(i).sRed
    lGreenColor = lSpectrumThemes.sSpectrumTheme(i).sRed
    lBlueColor = lSpectrumThemes.sSpectrumTheme(i).sBlue
    With lSpectrumThemes.sSpectrumTheme(i).sButtons
        frmGraphics.picChat1.Picture = .bctlChat(1)
        frmGraphics.picChat2.Picture = .bctlChat(2)
        frmGraphics.picChat3.Picture = .bctlChat(3)
        frmGraphics.picChannelFolder1.Picture = .bChannelFolder(1)
        frmGraphics.picChannelFolder2.Picture = .bChannelFolder(2)
        frmGraphics.picChannelFolder3.Picture = .bChannelFolder(3)
        frmGraphics.picConnect1.Picture = .bConnect(1)
        frmGraphics.picConnect2.Picture = .bConnect(2)
        frmGraphics.picConnect3.Picture = .bConnect(3)
        frmGraphics.picAudio1.Picture = .bAudio(1)
        frmGraphics.picAudio2.Picture = .bAudio(2)
        frmGraphics.picAudio3.Picture = .bAudio(3)
        frmGraphics.picNexIRC1.Picture = .bNexIRC(1)
        frmGraphics.picNexIRC2.Picture = .bNexIRC(2)
        frmGraphics.picNexIRC3.Picture = .bNexIRC(3)
        frmGraphics.picDisconnect1.Picture = .bDisconnect(1)
        frmGraphics.picDisconnect2.Picture = .bDisconnect(2)
        frmGraphics.picDisconnect3.Picture = .bDisconnect(3)
        frmGraphics.picScript1.Picture = .bScript(1)
        frmGraphics.picScript2.Picture = .bScript(2)
        frmGraphics.picScript3.Picture = .bScript(3)
        frmGraphics.picSend1.Picture = .bSend(1)
        frmGraphics.picSend2.Picture = .bSend(2)
        frmGraphics.picSend3.Picture = .bSend(3)
        frmGraphics.picOptions1.Picture = .bOptions(1)
        frmGraphics.picOptions2.Picture = .bOptions(2)
        frmGraphics.picOptions3.Picture = .bOptions(3)
        frmGraphics.picStop1.Picture = .bStop(1)
        frmGraphics.picStop2.Picture = .bStop(2)
        frmGraphics.picStop3.Picture = .bStop(3)
        frmGraphics.picBackward1.Picture = .bBackward(1)
        frmGraphics.picBackward2.Picture = .bBackward(2)
        frmGraphics.picBackward3.Picture = .bBackward(3)
        frmGraphics.picForward1.Picture = .bForward(1)
        frmGraphics.picForward2.Picture = .bForward(2)
        frmGraphics.picForward3.Picture = .bForward(3)
        frmGraphics.picPlay1.Picture = .bPlay(1)
        frmGraphics.picPlay2.Picture = .bPlay(2)
        frmGraphics.picPlay3.Picture = .bPlay(3)
        frmGraphics.picPause1.Picture = .bPause(1)
        frmGraphics.picPause2.Picture = .bPause(2)
        frmGraphics.picPause3.Picture = .bPause(3)
        frmGraphics.picExit1.Picture = .bExit(1)
        frmGraphics.picExit2.Picture = .bExit(2)
        frmGraphics.picExit3.Picture = .bExit(3)
        DoEvents
        mdiNexIRC.picChat.Picture = .bctlChat(1)
        mdiNexIRC.picChat.Left = .bChatLeft
        mdiNexIRC.picChat.Top = .bChatTop
        mdiNexIRC.picChat.Visible = True
        mdiNexIRC.picBackward.Left = .bBackwardLeft
        mdiNexIRC.picBackward.Top = .bBackwardTop
        mdiNexIRC.picBackward.Picture = .bBackward(1)
        mdiNexIRC.picBackward.Visible = True
        mdiNexIRC.picStop.Left = .bStopLeft
        mdiNexIRC.picStop.Top = .bStopTop
        mdiNexIRC.picStop.Visible = True
        mdiNexIRC.picStop.Picture = .bStop(1)
        mdiNexIRC.picPause.Left = .bPauseLeft
        mdiNexIRC.picPause.Top = .bPauseTop
        mdiNexIRC.picPause.Picture = .bPause(1)
        mdiNexIRC.picPause.Visible = True
        mdiNexIRC.picPlay.Picture = .bPlay(1)
        mdiNexIRC.picPlay.Top = .bPlayTop
        mdiNexIRC.picPlay.Left = .bPlayLeft
        mdiNexIRC.picPlay.Visible = True
        mdiNexIRC.picForward.Picture = .bForward(1)
        mdiNexIRC.picForward.Left = .bForwardLeft
        mdiNexIRC.picForward.Top = .bForwardTop
        mdiNexIRC.picForward.Visible = True
        mdiNexIRC.picDisconnect.Picture = .bDisconnect(1)
        mdiNexIRC.picDisconnect.Left = .bDisconnectLeft
        mdiNexIRC.picDisconnect.Top = .bDisconnectTop
        mdiNexIRC.picDisconnect.Visible = True
        mdiNexIRC.picConnect.Picture = .bConnect(1)
        mdiNexIRC.picConnect.Left = .bConnectLeft
        mdiNexIRC.picConnect.Top = .bConnectTop
        mdiNexIRC.picConnect.Visible = True
        mdiNexIRC.picAudio.Picture = .bAudio(1)
        mdiNexIRC.picAudio.Left = .bAudioLeft
        mdiNexIRC.picAudio.Top = .bAudioTop
        mdiNexIRC.picAudio.Visible = True
        mdiNexIRC.picOptions.Picture = .bOptions(1)
        mdiNexIRC.picOptions.Left = .bOptionsLeft
        mdiNexIRC.picOptions.Top = .bOptionsTop
        mdiNexIRC.picOptions.Visible = True
        mdiNexIRC.picChannelFolder.Picture = .bChannelFolder(1)
        mdiNexIRC.picChannelFolder.Left = .bChannelFolderLeft
        mdiNexIRC.picChannelFolder.Top = .bChannelFolderTop
        mdiNexIRC.picChannelFolder.Visible = True
        mdiNexIRC.picSend.Picture = .bSend(1)
        mdiNexIRC.picSend.Top = .bSendTop
        mdiNexIRC.picSend.Left = .bSendLeft
        mdiNexIRC.picSend.Visible = True
        mdiNexIRC.picScript.Picture = .bScript(1)
        mdiNexIRC.picScript.Left = .bScriptLeft
        mdiNexIRC.picScript.Top = .bScriptTop
        mdiNexIRC.picScript.Visible = True
        mdiNexIRC.picNexIRC.Picture = .bNexIRC(1)
        mdiNexIRC.picNexIRC.Left = .bNexIRCLeft
        mdiNexIRC.picNexIRC.Top = .bNexIRCTop
        mdiNexIRC.picNexIRC.Visible = True
        mdiNexIRC.picExit.Picture = .bExit(1)
        mdiNexIRC.picExit.Left = .bExitLeft
        mdiNexIRC.picExit.Top = .bExitTop
        mdiNexIRC.picExit.Visible = True
    End With
    mdiNexIRC.txtUrl.Width = lSpectrumThemes.sSpectrumTheme(i).sURLWidth
    mdiNexIRC.txtUrl.Height = lSpectrumThemes.sSpectrumTheme(i).sURLHeight
    mdiNexIRC.txtUrl.Left = lSpectrumThemes.sSpectrumTheme(i).sURLLeft
    mdiNexIRC.txtUrl.Top = lSpectrumThemes.sSpectrumTheme(i).sURLTop
    'mdiNexIRC.ctlMP3OCX.Left = lSpectrumThemes.sSpectrumTheme(i).sSpectrumLeft
    'mdiNexIRC.ctlMP3OCX.Top = lSpectrumThemes.sSpectrumTheme(i).sSpectrumTop
    'mdiNexIRC.ctlMP3OCX.Width = lSpectrumThemes.sSpectrumTheme(i).sSpectrumWidth
    'mdiNexIRC.ctlMP3OCX.Height = lSpectrumThemes.sSpectrumTheme(i).sSpectrumHeight
    SetPictureColors lSpectrumThemes.sSpectrumTheme(i).sRed, lSpectrumThemes.sSpectrumTheme(i).sBlue, lSpectrumThemes.sSpectrumTheme(i).sGreen
    If Len(lSpectrumThemes.sSpectrumTheme(i).sTextColor) <> 0 Then
        For F = 0 To 7
            frmMobileMixer.lblMixer(F).ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
        Next F
        frmMenus.ResetBarCheck
        frmMenus.mnuBarIndex(lSpectrumThemes.sSpectrumTheme(i).sBands).Checked = True
        mdiNexIRC.lblKHZ.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
        mdiNexIRC.lblSendMessage.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
        mdiNexIRC.lblNotify.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
        mdiNexIRC.lblFilename2.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
        mdiNexIRC.lblFrames.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
        frmMobileMixer.chkContinuous.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
        frmMobileMixer.chkShuffle.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
        frmMobileMixer.chkMute.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
    End If
    mdiNexIRC.txtUrl.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBackColor
    mdiNexIRC.txtUrl.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
    mdiNexIRC.txtUrl.Visible = True
    mdiNexIRC.txtMessage.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBackColor
    mdiNexIRC.txtMessage.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
    
    mdiNexIRC.cboProporties.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBackColor
    mdiNexIRC.cboProporties.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
    mdiNexIRC.cboProporties.Visible = True
    mdiNexIRC.cboValue.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBackColor
    mdiNexIRC.cboValue.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
    mdiNexIRC.cboColors.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBackColor
    mdiNexIRC.cboNotify.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBackColor
    mdiNexIRC.cboSpectrumThemes.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBackColor
    mdiNexIRC.cboSpectrumThemes.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
    mdiNexIRC.cboColors.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
    mdiNexIRC.cboNotify.ForeColor = lSpectrumThemes.sSpectrumTheme(i).sTextColor
    lSettings.sButtonType = lSpectrumThemes.sSpectrumTheme(i).sButtonType
    WriteINI lSettings.sIRCServerVisible, "Settings", "ButtonType", Trim(Str(lSettings.sButtonType))
    mdiNexIRC.UpdateMainButtonTypes
    If Len(lSpectrumThemes.sSpectrumTheme(i).sBottomToolbarColor) <> 0 Then frmMobileMixer.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBottomToolbarColor
    If Len(lSpectrumThemes.sSpectrumTheme(i).sBottomToolbarColor) <> 0 Then mdiNexIRC.picMP3OCX.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBottomToolbarColor
    If Len(lSpectrumThemes.sSpectrumTheme(i).sBottomToolbarColor) <> 0 Then mdiNexIRC.picNotify.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBottomToolbarColor
    If Len(lSpectrumThemes.sSpectrumTheme(i).sBottomToolbarColor) <> 0 Then frmMobileMixer.chkContinuous.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBottomToolbarColor
    If Len(lSpectrumThemes.sSpectrumTheme(i).sBottomToolbarColor) <> 0 Then frmMobileMixer.chkShuffle.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBottomToolbarColor
    If Len(lSpectrumThemes.sSpectrumTheme(i).sBottomToolbarColor) <> 0 Then frmMobileMixer.chkMute.BackColor = lSpectrumThemes.sSpectrumTheme(i).sBottomToolbarColor
    If Len(lSpectrumThemes.sSpectrumTheme(i).sToolbarGraphic) <> 0 Then
        msg = App.Path & "\data\images\" & lSpectrumThemes.sSpectrumTheme(i).sToolbarGraphic
        If DoesFileExist(msg) = True Then
            mdiNexIRC.picTopToolbar.Picture = LoadPicture(msg): mdiNexIRC.picTopToolbar.Refresh
        End If
    Else
        RefreshColors
    End If
    msg = ""
    lSpectrumThemes.sIndex = i
    If lSettings.sApplyThemeToIRCColors = True Then
        If Len(lSpectrumThemes.sSpectrumTheme(i).sIRCColors) <> 0 Then
            ApplyIRCColors Trim(lSpectrumThemes.sSpectrumTheme(i).sIRCColors)
        End If
    End If
    WriteINI GetINIFile(iSpectrum), "Settings", "Index", Trim(Str(i))
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ApplySpectrumTheme(lName As String)"
End Sub

Public Sub SetButtonType(lButton As ctlXPButton)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
Dim i As Integer
i = lSettings.sButtonType + 1
If i > 9 Then i = i + 1
lButton.ButtonType = i
'lButton.ButtonType = [Windows 32-bit]
If Len(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sFontname) <> 0 Then
    lButton.Font.Name = "Tahoma"
'    lButton.Font.Name = lSpectrumThemes.sSpectrumTheme(i).sFontname
End If
If lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sFontsize <> 0 Then
    'lButton.Font.Size = lSpectrumThemes.sSpectrumTheme(i).sFontsize
    lButton.Font.Size = 8
End If
lButton.SpecialEffect = cbShadowed
'lButton.SpecialEffect = cbEmbossed
lButton.UseGreyscale = True
'lButton.SoftBevel = False
lButton.UseMaskColor = True
lButton.Refresh
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetButtonType(lButton As ctlXPButton)"
    Err.Clear
End Sub

Public Sub ApplyIRCColors(lColors As String, Optional lApplyToMemoryOnly As Boolean, Optional lSaveToTheme As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg() As String, i As Integer, msg2() As String
If Len(lColors) = 0 Then Exit Sub
If lSaveToTheme = True Then
    WriteINI GetINIFile(iSpectrum), Trim(Str(lSpectrumThemes.sIndex)), "IRCColors", Trim(lColors)
End If
If lApplyToMemoryOnly = True Then
    lSettings.sColors = Trim(lColors)
    Exit Sub
Else
    lSettings.sColors = Trim(lColors)
    lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sIRCColors = Trim(lColors)
    msg = Split(lColors, " ")
    For i = 0 To UBound(msg)
        msg2 = Split(msg(i), ":")
        Select Case i
        Case 0
            If Len(msg2(1)) <> 0 Then Color.BGText = msg2(1)
        Case 1
            If Len(msg2(1)) <> 0 Then Color.Normal = msg2(1)
        Case 2
            If Len(msg2(1)) <> 0 Then Color.CTCP = msg2(1)
        Case 3
            If Len(msg2(1)) <> 0 Then Color.Notice = msg2(1)
        Case 4
            If Len(msg2(1)) <> 0 Then Color.Action = msg2(1)
        Case 5
            If Len(msg2(1)) <> 0 Then Color.Invite = msg2(1)
        Case 6
            If Len(msg2(1)) <> 0 Then Color.Join = msg2(1)
        Case 7
            If Len(msg2(1)) <> 0 Then Color.Kick = msg2(1)
        Case 8
            If Len(msg2(1)) <> 0 Then Color.Mode = msg2(1)
        Case 9
            If Len(msg2(1)) <> 0 Then Color.Nick = msg2(1)
        Case 10
            If Len(msg2(1)) <> 0 Then Color.Notify = msg2(1)
        Case 11
            If Len(msg2(1)) <> 0 Then Color.Part = msg2(1)
        Case 12
            If Len(msg2(1)) <> 0 Then Color.Quit = msg2(1)
        Case 13
            If Len(msg2(1)) <> 0 Then Color.Topic = msg2(1)
        Case 14
            If Len(msg2(1)) <> 0 Then Color.Whois = msg2(1)
        Case 15
            If Len(msg2(1)) <> 0 Then Color.Server = msg2(1)
        End Select
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ApplyIRCColors(lColors As String)"
End Sub

Public Sub SetProgressBarColor(lProgressBar As XP_ProgressBar)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lColor As Integer
lColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sProgressBarColor
Select Case lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sProgressBarColor
Case 0
    lProgressBar.Color = 0
'    lProgressBar.Color = &HFFFFFF
Case 1
    lProgressBar.Color = 0
Case 2
    lProgressBar.Color = &H800000
Case 3
    lProgressBar.Color = &H4040&
Case 4
    lProgressBar.Color = &H404080
Case 5
    lProgressBar.Color = &HC0&
Case 6
Case 7
Case 8
Case 9
Case 10
Case Else
    lProgressBar.Color = 0
End Select
If lProgressBar.Color = vbWhite Then lProgressBar.Color = 0
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetProgressBarColor(lProgressBar As XP_ProgressBar)"
End Sub

Public Sub LoadSpectrumThemes()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
lSpectrumThemes.sCount = Trim(ReadINI(GetINIFile(iSpectrum), "Settings", "Count", 0))
If lSpectrumThemes.sCount <> 0 Then
    lSpectrumThemes.sIndex = ReadINI(GetINIFile(iSpectrum), "Settings", "Index", 0)
    For i = 1 To lSpectrumThemes.sCount
        With lSpectrumThemes.sSpectrumTheme(i)
            .sName = ReadINI(GetINIFile(iSpectrum), Str(i), "Name", "")
            If Len(.sName) <> 0 Then
                mdiNexIRC.cboSpectrumThemes.AddItem .sName
                .sSpectrumLeft = ReadINI(GetINIFile(iSpectrum), Str(i), "SpectrumLeft", 0)
                .sSpectrumTop = ReadINI(GetINIFile(iSpectrum), Str(i), "SpectrumTop", 0)
                .sSpectrumHeight = ReadINI(GetINIFile(iSpectrum), Str(i), "SpectrumHeight", 0)
                .sSpectrumWidth = ReadINI(GetINIFile(iSpectrum), Str(i), "SpectrumWidth", 0)
                .sScreenShot = App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Str(i), "Screenshot", "")
                .sNicklistOptions.nNormalColor = ReadINI(GetINIFile(iSpectrum), Str(i), "NormalColor", "0")
                .sNicklistOptions.nVoiceColor = ReadINI(GetINIFile(iSpectrum), Str(i), "VoiceColor", "0")
                .sNicklistOptions.nOpColor = ReadINI(GetINIFile(iSpectrum), Str(i), "OpColor", "0")
                .sSpectrumBackcolor = ReadINI(GetINIFile(iSpectrum), Str(i), "SpectrumBackcolor", "")
                .sBGTextColor = ReadINI(GetINIFile(iSpectrum), Str(i), "BGTextColor", 0)
                .sBackColor = ReadINI(GetINIFile(iSpectrum), Str(i), "BackColor", "")
                .sBottomBandsColor = ReadINI(GetINIFile(iSpectrum), Str(i), "BottomBandsColor", "")
                .sDividerColor = ReadINI(GetINIFile(iSpectrum), Str(i), "DividerColor", "")
                .sLeftChanColor = ReadINI(GetINIFile(iSpectrum), Str(i), "LeftChanColor", "")
                .sPeaksColor = ReadINI(GetINIFile(iSpectrum), Str(i), "PeaksColor", "")
                .sRightChanColor = ReadINI(GetINIFile(iSpectrum), Str(i), "RightChanColor", "")
                .sTopBandsColor = ReadINI(GetINIFile(iSpectrum), Str(i), "RightChanColor", "")
                .sBands = ReadINI(GetINIFile(iSpectrum), Str(i), "Bands", 15)
                .sToolbarGraphic = ReadINI(GetINIFile(iSpectrum), Str(i), "ToolbarGraphic", "")
                .sBottomToolbarColor = ReadINI(GetINIFile(iSpectrum), Str(i), "BottomToolbarColor", "")
                .sTextColor = ReadINI(GetINIFile(iSpectrum), Str(i), "TextColor", "16777215")
                .sRed = ReadINI(GetINIFile(iSpectrum), Str(i), "Red", 0)
                .sGreen = ReadINI(GetINIFile(iSpectrum), Str(i), "Green", 0)
                .sBlue = ReadINI(GetINIFile(iSpectrum), Str(i), "Blue", 0)
                .sIRCColors = Trim(ReadINI(GetINIFile(iSpectrum), Str(i), "IRCColors", ""))
                .sMode = ReadINI(GetINIFile(iSpectrum), Str(i), "Mode", 0)
                .sButtonType = ReadINI(GetINIFile(iSpectrum), Str(i), "ButtonType", 12)
                .sDisableToolbarColors = ReadINI(GetINIFile(iSpectrum), Str(i), "DisableToolbarColors", False)
                .sFontname = ReadINI(GetINIFile(iSpectrum), Str(i), "Fontname", "Tahoma")
                .sFontsize = Int(ReadINI(GetINIFile(iSpectrum), Str(i), "Fontsize", 8))
                .sOscilloType = Int(ReadINI(GetINIFile(iSpectrum), Str(i), "OscilloType", 2))
                .sButtons.bExitLeft = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "ExitLeft", -400))
                .sButtons.bExitTop = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "ExitTop", 0))
                .sURLHeight = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "URLHeight", 0))
                .sURLLeft = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "URLLeft", 0))
                .sURLTop = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "URLTop", 0))
                .sURLWidth = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "URLWidth", 0))
                .sProgressBarStyle = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "ProgressBarStyle", 0))
                .sProgressBarColor = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "ProgressBarColor", 0))
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Exit1", ""), .sButtons.bExit(1)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Exit2", ""), .sButtons.bExit(2)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Exit3", ""), .sButtons.bExit(3)
                .sButtons.bBackwardLeft = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "BackwardLeft", -400))
                .sButtons.bBackwardTop = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "BackwardTop", 0))
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Backward1", ""), .sButtons.bBackward(1)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Backward2", ""), .sButtons.bBackward(2)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Backward3", ""), .sButtons.bBackward(3)
                .sButtons.bStopLeft = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "StopLeft", -400))
                .sButtons.bStopTop = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "StopTop", 0))
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Stop1", ""), .sButtons.bStop(1)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Stop2", ""), .sButtons.bStop(2)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Stop3", ""), .sButtons.bStop(3)
                .sButtons.bPauseLeft = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "PauseLeft", -400))
                .sButtons.bPauseTop = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "PauseTop", 0))
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Pause1", ""), .sButtons.bPause(1)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Pause2", ""), .sButtons.bPause(2)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Pause3", ""), .sButtons.bPause(3)
                .sButtons.bForwardLeft = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "ForwardLeft", -400))
                .sButtons.bForwardTop = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "ForwardTop", 0))
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Forward1", ""), .sButtons.bForward(1)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Forward2", ""), .sButtons.bForward(2)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Forward3", ""), .sButtons.bForward(3)
                .sButtons.bPlayLeft = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "PlayLeft", -400))
                .sButtons.bPlayTop = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "PlayTop", 0))
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Play1", ""), .sButtons.bPlay(1)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Play2", ""), .sButtons.bPlay(2)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Play3", ""), .sButtons.bPlay(3)
                .sButtons.bConnectLeft = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "ConnectLeft", -400))
                .sButtons.bConnectTop = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "ConnectTop", 0))
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Connect1", ""), .sButtons.bConnect(1)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Connect2", ""), .sButtons.bConnect(2)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Connect3", ""), .sButtons.bConnect(3)
                .sButtons.bSendLeft = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "SendLeft", -400))
                .sButtons.bSendTop = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "SendTop", 0))
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Send1", ""), .sButtons.bSend(1)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Send2", ""), .sButtons.bSend(2)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Send3", ""), .sButtons.bSend(3)
                .sButtons.bDisconnectLeft = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "DisconnectLeft", -400))
                .sButtons.bDisconnectTop = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "DisconnectTop", 0))
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Disconnect1", ""), .sButtons.bDisconnect(1)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Disconnect2", ""), .sButtons.bDisconnect(2)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Disconnect3", ""), .sButtons.bDisconnect(3)
                .sButtons.bAudioLeft = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "AudioLeft", -400))
                .sButtons.bAudioTop = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "AudioTop", 0))
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Audio1", ""), .sButtons.bAudio(1)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Audio2", ""), .sButtons.bAudio(2)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Audio3", ""), .sButtons.bAudio(3)
                .sButtons.bNexIRCLeft = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "NexIRCLeft", -400))
                .sButtons.bNexIRCTop = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "NexIRCTop", 0))
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "NexIRC1", ""), .sButtons.bNexIRC(1)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "NexIRC2", ""), .sButtons.bNexIRC(2)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "NexIRC3", ""), .sButtons.bNexIRC(3)
                .sButtons.bScriptLeft = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "ScriptLeft", -400))
                .sButtons.bScriptTop = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "ScriptTop", 0))
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Script1", ""), .sButtons.bScript(1)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Script2", ""), .sButtons.bScript(2)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Script3", ""), .sButtons.bScript(3)
                .sButtons.bSendTop = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "SendTop", 0))
                .sButtons.bSendLeft = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "SendLeft", -400))
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Send1", ""), .sButtons.bSend(1)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Send2", ""), .sButtons.bSend(2)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Send3", ""), .sButtons.bSend(3)
                .sButtons.bOptionsLeft = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "OptionsLeft", -400))
                .sButtons.bOptionsTop = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "OptionsTop", 0))
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Options1", ""), .sButtons.bOptions(1)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Options2", ""), .sButtons.bOptions(2)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Options3", ""), .sButtons.bOptions(3)
                .sButtons.bChannelFolderLeft = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "ChannelFolderLeft", -400))
                .sButtons.bChannelFolderTop = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "ChannelFolderTop", 0))
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "ChannelFolder1", ""), .sButtons.bChannelFolder(1)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "ChannelFolder2", ""), .sButtons.bChannelFolder(2)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "ChannelFolder3", ""), .sButtons.bChannelFolder(3)
                .sButtons.bChatLeft = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "ChatLeft", -400))
                .sButtons.bChatTop = Int(ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "ChatTop", 0))
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Chat1", ""), .sButtons.bctlChat(1)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Chat2", ""), .sButtons.bctlChat(2)
                ApplyImageToPictureBox App.Path & "\data\images\" & ReadINI(GetINIFile(iSpectrum), Trim(Str(i)), "Chat3", ""), .sButtons.bctlChat(3)
            End If
        End With
    Next i
    If lSpectrumThemes.sIndex <> 0 Then
        mdiNexIRC.cboSpectrumThemes.Text = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sName
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub LoadSpectrumThemes()"
End Sub

Public Function FindSpectrumThemeByName(lName As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To lSpectrumThemes.sCount
    If LCase(lName) = LCase(lSpectrumThemes.sSpectrumTheme(i).sName) Then
        FindSpectrumThemeByName = i
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindSpectrumThemeByName(lName As String) As Integer"
End Function

Public Function GetRed(cValue As Long) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
GetRed = cValue Mod 256
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function GetRed(cValue As Long) As Long"
End Function

Public Function GetGreen(cValue As Long) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
GetGreen = Int((cValue / 256)) Mod 256
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function GetGreen(cValue As Long) As Long"
End Function

Public Function GetBlue(cValue As Long) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
GetBlue = Int(cValue / 65536)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function GetBlue(cValue As Long) As Long"
End Function

Public Sub RefreshColors()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
mdiNexIRC.tmrCheckButtonColors.Enabled = False
If lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sDisableToolbarColors = True Then Exit Sub
If lSettings.sRefreshPictureColors = True Then
    If Len(mdiNexIRC.picBackward.Tag) <> 0 Then
        If GetPixel(mdiNexIRC.picBackward.hDC, 12, 12) <> mdiNexIRC.picBackward.Tag Then
            SetPictureColor mdiNexIRC.picBackward, lRedColor, lBlueColor, lGreenColor, True
        End If
    End If
    If Len(mdiNexIRC.picForward.Tag) <> 0 Then
        If GetPixel(mdiNexIRC.picForward.hDC, 12, 12) <> mdiNexIRC.picForward.Tag Then
            SetPictureColor mdiNexIRC.picForward, lRedColor, lBlueColor, lGreenColor, True
        End If
    End If
    If Len(mdiNexIRC.picStop.Tag) <> 0 Then
        If GetPixel(mdiNexIRC.picStop.hDC, 12, 12) <> mdiNexIRC.picStop.Tag Then
            SetPictureColor mdiNexIRC.picStop, lRedColor, lBlueColor, lGreenColor, True
        End If
    End If
    If Len(mdiNexIRC.picPause.Tag) <> 0 Then
        If GetPixel(mdiNexIRC.picPause.hDC, 12, 12) <> mdiNexIRC.picPause.Tag Then
            SetPictureColor mdiNexIRC.picPause, lRedColor, lBlueColor, lGreenColor, True
        End If
    End If
    If Len(mdiNexIRC.picPlay.Tag) <> 0 Then
        If GetPixel(mdiNexIRC.picPlay.hDC, 12, 12) <> mdiNexIRC.picPlay.Tag Then
            SetPictureColor mdiNexIRC.picPlay, lRedColor, lBlueColor, lGreenColor, True
        End If
    End If
    If Len(mdiNexIRC.picConnect.Tag) <> 0 Then
        If GetPixel(mdiNexIRC.picConnect.hDC, 12, 12) <> mdiNexIRC.picConnect.Tag Then
            SetPictureColor mdiNexIRC.picConnect, lRedColor, lBlueColor, lGreenColor, True
        End If
    End If
    If Len(mdiNexIRC.picDisconnect.Tag) <> 0 Then
        If GetPixel(mdiNexIRC.picDisconnect.hDC, 12, 12) <> mdiNexIRC.picDisconnect.Tag Then
            SetPictureColor mdiNexIRC.picDisconnect, lRedColor, lBlueColor, lGreenColor, True
        End If
    End If
    If Len(mdiNexIRC.picChannelFolder.Tag) <> 0 Then
        If GetPixel(mdiNexIRC.picChannelFolder.hDC, 12, 12) <> mdiNexIRC.picChannelFolder.Tag Then
            SetPictureColor mdiNexIRC.picChannelFolder, lRedColor, lBlueColor, lGreenColor, True
        End If
    End If
    If Len(mdiNexIRC.picOptions.Tag) <> 0 Then
        If GetPixel(mdiNexIRC.picOptions.hDC, 12, 12) <> mdiNexIRC.picOptions.Tag Then
            SetPictureColor mdiNexIRC.picOptions, lRedColor, lBlueColor, lGreenColor, True
        End If
    End If
    If Len(mdiNexIRC.picAudio.Tag) <> 0 Then
        If GetPixel(mdiNexIRC.picAudio.hDC, 12, 12) <> mdiNexIRC.picAudio.Tag Then
            SetPictureColor mdiNexIRC.picAudio, lRedColor, lBlueColor, lGreenColor, True
        End If
    End If
    If Len(mdiNexIRC.picSend.Tag) <> 0 Then
        If GetPixel(mdiNexIRC.picSend.hDC, 12, 12) <> mdiNexIRC.picSend.Tag Then
            SetPictureColor mdiNexIRC.picSend, lRedColor, lBlueColor, lGreenColor, True
        End If
    End If
    If Len(mdiNexIRC.picScript.Tag) <> 0 Then
        If GetPixel(mdiNexIRC.picScript.hDC, 12, 12) <> mdiNexIRC.picScript.Tag Then
            SetPictureColor mdiNexIRC.picScript, lRedColor, lBlueColor, lGreenColor, True
        End If
    End If
    If Len(mdiNexIRC.picNexIRC.Tag) <> 0 Then
        If GetPixel(mdiNexIRC.picNexIRC.hDC, 12, 12) <> mdiNexIRC.picNexIRC.Tag Then
            SetPictureColor mdiNexIRC.picNexIRC, lRedColor, lBlueColor, lGreenColor, True
        End If
    End If
    If Len(mdiNexIRC.picChat.Tag) <> 0 Then
        If GetPixel(mdiNexIRC.picChat.hDC, 12, 12) <> mdiNexIRC.picChat.Tag Then
            SetPictureColor mdiNexIRC.picChat, lRedColor, lBlueColor, lGreenColor, True
        End If
    End If
End If
mdiNexIRC.tmrCheckButtonColors.Enabled = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub RefreshColors()"
End Sub

Public Sub SetPictureColor(lPictureBox As PictureBox, lRed As Long, lBlue As Long, lGreen As Long, lMainForm As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim pX As Long, pY As Long, X As Long, Y As Long, colorval As Long, red As Long, green As Long, blue As Long, red2 As Long, green2 As Long, blue2 As Long
If lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sDisableToolbarColors = True Then Exit Sub
If lSettings.sRefreshPictureColors = False Then Exit Sub
If lMainForm = True Then mdiNexIRC.tmrCheckButtonColors.Enabled = False
pX = lPictureBox.Width / Screen.TwipsPerPixelX - 1
pY = lPictureBox.Height / Screen.TwipsPerPixelY - 1
If lRed = 0 Or lBlue = 0 Or lGreen = 0 Then
    Exit Sub
End If
lRedColor = lRed
lGreenColor = lGreen
lBlueColor = lBlue
For X = 0 To pX
    For Y = 0 To pY
        colorval = GetPixel(lPictureBox.hDC, X, Y)
        red = GetRed(colorval)
        green = GetGreen(colorval)
        blue = GetBlue(colorval)
        red2 = red + Int(lRed / 100 * red)
        green2 = green + Int(lGreen / 100 * green)
        blue2 = blue + Int(lBlue / 100 * blue)
        If red2 >= 255 Then red2 = 255
        If green2 >= 255 Then green2 = 255
        If blue2 >= 255 Then blue2 = 255
        If red2 <= 0 Then red2 = 0
        If green2 <= 0 Then green2 = 0
        If blue2 <= 0 Then blue2 = 0
        SetPixel lPictureBox.hDC, X, Y, RGB(red2, green2, blue2): DoEvents
        If Y = 12 And X = 12 Then
            lPictureBox.Tag = GetPixel(lPictureBox.hDC, X, Y)
        End If
    Next Y
Next X
If lMainForm = True Then mdiNexIRC.tmrCheckButtonColors.Enabled = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetPictureColor(lPictureBox As PictureBox, lRed As Long, lBlue As Long, lGreen As Long)"
End Sub
