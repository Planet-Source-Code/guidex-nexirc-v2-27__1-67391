Attribute VB_Name = "mdlMultimediaInterface"
Option Explicit
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Enum sndConst
    SND_ASYNC = &H1
    SND_LOOP = &H8
    SND_MEMORY = &H4
    SND_NODEFAULT = &H2
    SND_NOSTOP = &H10
    SND_SYNC = &H0
End Enum

Public Function RemoveFromPlaylist(lFileName As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
i = FindFileIndexByFilename(lFileName)
If lSettings.sPlaylistVisible = True Then frmPlaylist.lstPlaylist.RemoveItem FindListBoxIndex(lFileName, frmPlaylist.lstPlaylist)
If i <> 0 Then
    lFiles.fFile(i).fFilename = ""
    WriteINI GetINIFile(iPlaylist), "Settings", Str(i), ""
Else
    ProcessReplaceString sProgrammingError, lSettings.sActiveServerForm.txtIncoming, "Entry not Found", "404"
End If
End Function

Public Sub LoadM3U(lFileName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
If Len(lFileName) <> 0 Then
    msg = ReadFile(App.Path & "\data\playlists\" & lFileName)
    If Len(msg) <> 0 Then
        'Read M3U
    Else
        ClearPlaylistMemory
        If lSettings.sGeneralPrompts = True Then
            MsgBox "This playlist is empty", vbInformation
        End If
    End If
End If
End Sub

Public Sub DisplayPlaylists()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, m As Integer
mdiNexIRC.file1.Path = App.Path & "\data\playlists\"
If (mdiNexIRC.mnuPlaylistCollection.Count - 1) <> 0 Then
    For i = 0 To mdiNexIRC.mnuPlaylistCollection.Count - 1
        If i <> 0 Then Unload mdiNexIRC.mnuPlaylistCollection(i)
    Next i
End If
For i = 0 To mdiNexIRC.file1.ListCount
    If Len(mdiNexIRC.file1.List(i)) <> 0 Then
        If Right(LCase(mdiNexIRC.file1.List(i)), 4) = ".m3u" Or Right(LCase(mdiNexIRC.file1.List(i)), 4) = ".ini" Then
            m = m + 1
            Load mdiNexIRC.mnuPlaylistCollection(m)
            mdiNexIRC.mnuPlaylistCollection(m).Visible = True
            mdiNexIRC.mnuPlaylistCollection(m).Caption = mdiNexIRC.file1.List(i)
        End If
    End If
Next i
End Sub

Public Sub ClearPlaylistMemory()
Dim i As Integer
For i = 0 To 10000
    With lFiles.fFile(i)
        .fFilename = ""
    End With
    lFiles.fCount = 0
    lFiles.fIndex = 0
Next i
End Sub

Public Sub DefragPlaylist()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg(10000) As String, i As Integer, F As Integer
For i = 0 To 10000
    If Len(lFiles.fFile(i).fFilename) <> 0 Then
        If DoesFileExist(lFiles.fFile(i).fFilename) = True Then
            F = F + 1
            msg(F) = lFiles.fFile(i).fFilename
        End If
    End If
Next i
If F <> 0 Then
    ClearPlaylistMemory
    lFiles.fCount = F
    For i = 0 To lFiles.fCount
        If Len(msg(i)) <> 0 Then
            If FindFileIndex(msg(i)) = 0 Then
                lFiles.fFile(i).fFilename = msg(i)
            End If
        End If
    Next i
    Kill GetINIFile(iPlaylist)
    SavePlaylist
End If
End Sub

Public Sub LoadPlaylist(Optional lFileName As String, Optional lShowPlaylist As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, F As Integer
If lShowPlaylist = True Then Unload frmPlaylist
If Len(lFileName) <> 0 Then SetPlaylistINIFile lFileName
Select Case LCase(Right(GetINIFile(iPlaylist), 4))
Case ".m3u"
    LoadM3U GetINIFile(iPlaylist)
    If lShowPlaylist = True Then frmPlaylist.Show
Case ".ini"
    lFiles.fCount = ReadINI(GetINIFile(iPlaylist), "Settings", "Count", 0)
    If lFiles.fCount <> 0 Then
        For i = 0 To lFiles.fCount
            F = F + 1
            lFiles.fFile(F).fFilename = ReadINI(GetINIFile(iPlaylist), "Settings", Str(i), "")
            If Len(lFiles.fFile(F).fFilename) <> 0 Then
                If DoesFileExist(lFiles.fFile(F).fFilename) = False Then
                    F = F - 1
                End If
            End If
        Next i
    End If
    DefragPlaylist
    If lShowPlaylist = True Then frmPlaylist.Show
End Select
End Sub

Public Sub PromptAddToPlaylist()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
frmAddFolderToPlaylist.Show 1
msg = lDirReturn
lDirReturn = ""
If Len(msg) <> 0 Then
    AddFolderToPlaylist msg
End If
End Sub

'Public Sub SwitchPlaybackEngine(lType As ePlaybackEngine)
'Dim i As Integer, lOldEngine As ePlaybackEngine
'lOldEngine = lPlayback.pCurrentEngine
'If lOldEngine <> lType Then
'    If lPlayback.pPlaying = True Then
'        MenuStop
'        lPlayback.pPlaying = False
'    End If
'End If
'lPlayback.pCurrentEngine = lType
'Select Case lType
'Case pMediaPlayer
'    'For i = 1 To mdiNexIRC.Count
'        'ActivateActiveFormResize
'    'Next i
'    With mdiNexIRC
'        .picMP3OCX.Visible = False
''        .ctlMP3OCX.Visible = False
'        .tmrPlaySoon.Enabled = False
'        .tmrPlaySoon.interval = 0
'    End With
'Case pMp3
'    With mdiNexIRC
'        If lSettings.sAlwaysShowAudioSettings = True Then
'            .picMP3OCX.Visible = True
'        End If
'    End With
'    If lSettings.sContinuousPlay = True Then ActivateContinuousPlay
'End Select
'End Sub

Public Function AddToPlaylist(lFileName As String, Optional lSavePlaylist As Boolean) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String
If lSettings.sExlusiveToMp3InPlaylist = True And LCase(Right(lFileName, 4)) <> ".mp3" Then Exit Function
If FindFileIndex(lFileName) <> 0 Then Exit Function
If lSettings.sExlusiveToMp3InPlaylist = True Then
    If IsMP3(lFileName) = True Then
        i = lFiles.fCount + 1
        lFiles.fCount = i
        lFiles.fFile(i).fFilename = lFileName
        If lSavePlaylist = True Then SavePlaylist
        If lSettings.sPlaylistVisible = True Then
        End If
    Else
        Exit Function
    End If
ElseIf lSettings.sExlusiveToMp3InPlaylist = False Then
    i = lFiles.fCount + 1
    lFiles.fCount = i
    lFiles.fFile(i).fFilename = lFileName
    If lSavePlaylist = True Then SavePlaylist
End If
If lSettings.sPlaylistVisible = True Then
    msg = lFileName
    msg = GetFileTitle(msg)
    frmPlaylist.lstPlaylist.AddItem msg
End If
End Function

Public Sub AddFolderToPlaylist(lFolder As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, lTmp1 As Long, cCol As tSearch, msg As String
GetFiles lFolder, "*.mp3", vbNormal, cCol
For lTmp1 = 1 To cCol.Count
    If Len(cCol.Path(lTmp1)) <> 0 And DoesFileExist(cCol.Path(lTmp1)) = True And FindFileIndex(cCol.Path(lTmp1)) = 0 Then
        If lSettings.sExlusiveToMp3InPlaylist = True Then
            If IsMP3(cCol.Path(lTmp1)) = True Then
                lFiles.fCount = lFiles.fCount + 1
                lFiles.fFile(lFiles.fCount).fFilename = cCol.Path(lTmp1)
                If lSettings.sPlaylistVisible = True Then
                    msg = lFiles.fFile(lFiles.fCount).fFilename
                    msg = GetFileTitle(msg)
                    frmPlaylist.lstPlaylist.AddItem msg
                    DoEvents
                End If
            End If
        Else
            lFiles.fCount = lFiles.fCount + 1
            lFiles.fFile(lFiles.fCount).fFilename = cCol.Path(lTmp1)
            If lSettings.sPlaylistVisible = True Then
                msg = lFiles.fFile(lFiles.fCount).fFilename
                msg = GetFileTitle(msg)
                frmPlaylist.lstPlaylist.AddItem msg
                DoEvents
            End If
        End If
    End If
Next lTmp1
SavePlaylist
End Sub

Public Sub SavePlaylist()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String
WriteINI GetINIFile(iPlaylist), "Settings", "Count", Str(lFiles.fCount)
For i = 0 To lFiles.fCount
    If Len(lFiles.fFile(i).fFilename) <> 0 And DoesFileExist(lFiles.fFile(i).fFilename) = True Then
        WriteINI GetINIFile(iPlaylist), "Settings", Str(i), lFiles.fFile(i).fFilename
    Else
        lFiles.fFile(i).fFilename = ""
    End If
Next i
End Sub

Public Function FindFileIndexByFilename(lFileName As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, msg2 As String
msg = lFileName
If Len(lFileName) <> 0 Then
    For i = 0 To lFiles.fCount
        If i = 1001 Then
            FindFileIndexByFilename = 0
            Exit For
        End If
        msg2 = lFiles.fFile(i).fFilename
        If Len(msg2) <> 0 Then
            If InStr(1, LCase(msg2), LCase(msg)) Then
                FindFileIndexByFilename = i
                Exit For
            End If
        End If
    Next i
End If
End Function

Public Function FindFileIndex(lFullPath As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lFullPath) <> 0 Then
    For i = 0 To lFiles.fCount
        If i = 1001 Then
            FindFileIndex = 0
            Exit For
        End If
        If LCase(lFiles.fFile(i).fFilename) = LCase(lFullPath) Then
            FindFileIndex = i
            Exit For
        Else
        End If
    Next i
End If
End Function

Public Function PlayFile(lFileName As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer
If DoesFileExist(lFileName) Then
    i = AddToPlaylist(lFileName, True)
    lFiles.fIndex = i
    msg = lFileName
    msg2 = msg
    msg2 = GetFileTitle(msg2)
    msg = Left(msg, Len(msg) - Len(msg2))
    AddFolderToPlaylist msg
    frmPlayer.Show
    frmPlayer.ctlMovieX1.OpenMovie lFileName
    frmPlayer.ctlMovieX1.PlayMovie
Else
    Exit Function
End If
'Dim typeDevice As String, result, msg As String, msg2 As String
'If GetStatusMultimedia(LCase(Trim(mdiNexIRC.lblMultimedia.Caption))) = "playing" Then
'    StopMultimedia mdiNexIRC.lblMultimedia.Caption
'    CloseMultimedia mdiNexIRC.lblMultimedia.Caption
'End If
'mdiNexIRC.lblMultimedia.Caption = Time$ & Date$ & GetRnd(10000)
'If Right(LCase(lFileName), 4) = ".avi" Then
'    If lSettings.sGeneralPrompts = True Then MsgBox "Format not supported"
'    Exit Function
'ElseIf Right(LCase(lFileName), 4) = ".rmi" Or Right(LCase(lFileName), 4) = ".mid" Then
'    typeDevice = "sequencer"
'ElseIf Right(LCase(lFileName), 4) = ".mp3" Then
''    If lPlayback.pCurrentEngine = pMp3 Then
''        PlayMP3 lFileName
''    Else
'    typeDevice = "MPEGVideo"
''    End If
'ElseIf Right(LCase(lFileName), 4) = ".vob" Then
'    If lSettings.sGeneralPrompts = True Then MsgBox "Format not supported"
'    Exit Function
'ElseIf Right(LCase(lFileName), 4) = ".cda" Then
'    If lSettings.sGeneralPrompts = True Then MsgBox "CDAudio Format not currently supported", vbInformation
'    Exit Function
'ElseIf Right(LCase(lFileName), 4) = ".wav" Then
'Else
'    typeDevice = "MPEGVideo"
'End If
''result = OpenMultimedia(frmVideo.fraVideo.hWnd, mdiNexIRC.lblMultimedia.Caption, lFilename, typeDevice)
'result = OpenMultimedia(mdiNexIRC.hWnd, mdiNexIRC.lblMultimedia.Caption, lFileName, typeDevice)
'If result = "Success" Then
'    SwitchPlaybackEngine pMediaPlayer
'    lPlayback.pPlaying = True
'    msg = PlayMultimedia(mdiNexIRC.lblMultimedia.Caption, 0, GetTotalframes(mdiNexIRC.lblMultimedia.Caption))
'    lPlayback.pCurrentFile = lFileName
'    msg = ""
'    msg = lFileName
'    msg2 = lFileName
'    msg2 = GetFileTitle(msg2)
'    lFiles.fIndex = FindFileIndexByFilename(msg2)
'    msg = Left(msg, Len(msg) - Len(msg2))
'    AddFolderToPlaylist msg
'ElseIf Left(LCase(result), 18) = "a problem occurred" Then
'    If lSettings.sGeneralPrompts = True Then
'        MsgBox "There was a problem playing the file '" & lFileName & "'", vbExclamation
'    End If
'    Exit Function
'End If
'PlayFile = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function PlayFile(lFileName As String) As Boolean"
End Function

Public Sub ToggleMixer(lEnabled As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lEnabled = True Then
    lSettings.sShowQuickmix = True
    frmMenus.mnuShowHideMixer.Caption = "Hide"
    mdiNexIRC.picMobileMixer.Visible = True
    frmMobileMixer.Show
    frmMobileMixer.Move -4 * Screen.TwipsPerPixelX, -6 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
Else
    lSettings.sShowQuickmix = False
    frmMenus.mnuShowHideMixer.Caption = "Show"
    If lSettings.sMobileMixerVisible = True Then
        mdiNexIRC.picMobileMixer.Visible = False
        Unload frmMobileMixer
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ToggleMixer(lEnabled As Boolean)"
End Sub

Public Sub ActivateContinuousPlay(Optional lToggleOff As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lFiles.fCount = 0 Then Exit Sub
Dim i As Integer
If lToggleOff = True Then
    lSettings.sContinuousPlay = False
    WriteINI GetINIFile(iIRC), "Settings", "ContinuousPlay", lSettings.sContinuousPlay
    mdiNexIRC.tmrContinuousPlay.Enabled = False
    frmMobileMixer.chkContinuous.Value = 0
    Exit Sub
End If
If lSettings.sContinuousPlay = True Then
    If lPlayback.pCurrentEngine = pMediaPlayer And lPlayback.pPlaying = True Then
        i = GetPercent(mdiNexIRC.lblMultimedia.Caption)
        If i = 100 Or i > 101 Then lPlayback.pPlaying = False
    End If
    If lPlayback.pPlaying = False Then
        If frmMobileMixer.chkShuffle.Value = 1 Then
RandomFile:
            i = GetRnd(lFiles.fCount)
            If DoesFileExist(lFiles.fFile(i).fFilename) = True Then
                lFiles.fIndex = i
            Else
                GoTo RandomFile
            End If
        ElseIf frmMobileMixer.chkShuffle.Value = 0 Then
            If lFiles.fIndex + 1 = lFiles.fCount Then lFiles.fIndex = 0
            lFiles.fIndex = lFiles.fIndex + 1
        End If
        PlayFile lFiles.fFile(lFiles.fIndex).fFilename
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ActivateContinuousPlay(Optional lToggleOff As Boolean)"
End Sub

'Public Sub PlayMP3(lFileName As String)
'Dim msg As String, i As Integer
'If lPlayback.pCurrentEngine = pMp3 Then
'    mdiNexIRC.tmrContinuousPlay.Enabled = False
    'mdiNexIRC.ctlMP3OCX.Stop
'    Select Case lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sOscilloType
'    Case 0
'    '    mdiNexIRC.ctlMP3OCX.OscilloType = 0
'    Case 1
'    '    mdiNexIRC.ctlMP3OCX.OscilloType = 1
'    Case 2
'    '    mdiNexIRC.ctlMP3OCX.OscilloType = 2
'    End Select
'    'mdiNexIRC.ctlMP3OCX.Visible = True
'    msg = lFileName
'    msg = GetFileTitle(msg)
'    lPlayback.pCurrentFile = msg
'    lFiles.fIndex = FindFileIndexByFilename(msg)
'    lSettings.sPlayNext = lFileName
'    mdiNexIRC.tmrPlaySoon.interval = 1000
'    mdiNexIRC.tmrPlaySoon.Enabled = True
'End If
'If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub PlayMP3(lFilename As String)"
'End Sub

Public Sub ActivatePlayback()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer
If Len(lSettings.sPlayNext) <> 0 Then
    If DoesFileExist(lSettings.sPlayNext) = True Then
'        Stop
'        mdiNexIRC.ctlMP3OCX.Visible = True
'        mdiNexIRC.ctlMP3OCX.Play lSettings.sPlayNext
'        MsgBox lSettings.sPlayNext
'        GetMP3Info lSettings.sPlayNext
        msg = lSettings.sPlayNext
        msg = GetFileTitle(msg)
        If lSettings.sAudioServer = True Then
            If lSettings.sFileOfferInChannel = True Then
                If lSettings.sOfferWhenPlayed = True Then
                    If lSettings.sChannelCount <> 0 Then
                        For i = 1 To lSettings.sChannelCount
                            'ProcessReplaceString sFileOffer, lChannel(i).txtIncoming, lChannelName(i), Format(FileLen(lSettings.sPlayNext), "###,###,###")
                            ProcessReplaceString sFileOffer, ReturnChannelIncomingTBox(i), ReturnChannelName(i), Format(FileLen(lSettings.sPlayNext), "###,###,###")
'                            lSettings.sActiveServerForm.tcp.SendData "PRIVMSG " &      lChannelName(i) & " :" &
                        Next i
                    End If
                End If
            End If
        End If
        lSettings.sPlayNext = ""
    End If
End If
If lSettings.sContinuousPlay = True Then mdiNexIRC.tmrContinuousPlay.Enabled = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ActivatePlayback()"
End Sub

Public Function PlayWav(strPath As String, sndVal As sndConst)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sPlaySounds = True Then
    sndPlaySound strPath, sndVal
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function PlayWav(strPath As String, sndVal As sndConst)"
End Function

'Public Sub GetMP3Info(lFileName As String)
'Dim msg As String
'If Len(lFileName) <> 0 And DoesFileExist(lFileName) = True Then
    'mdiNexIRC.ctlMP3OCX.GetFileInfo lFileName: DoEvents
    'If Len(mdiNexIRC.ctlMP3OCX.Artist) <> 0 And Len(mdiNexIRC.ctlMP3OCX.Title) <> 0 Then
        'mdiNexIRC.lblFilename2.Caption = mdiNexIRC.ctlMP3OCX.Artist & ": " & mdiNexIRC.ctlMP3OCX.Title
'    Else
'        msg = lFileName
'        msg = GetFileTitle(msg)
'        mdiNexIRC.lblFilename2.Caption = msg
'    End If
'End If
'If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub GetMP3Info(lFilename As String)"
'End Sub

Public Sub MenuPlay()
If lSettings.sHandleErrors = True Then On Local Error Resume Next

frmPlayer.ctlMovieX1.PlayMovie
'Select Case lPlayback.pCurrentEngine
'Case pMediaPlayer
'    PlayMultimedia mdiNexIRC.lblMultimedia.Caption, 0, GetTotalframes(mdiNexIRC.lblMultimedia.Caption)
'Case pMp3
    'Select Case mdiNexIRC.ctlMP3OCX.PlayState
    'Case 0
    '    mdiNexIRC.ctlMP3OCX.Play lFiles.fFile(lFiles.fIndex).fFilename
    '    mdiNexIRC.ctlMP3OCX.Visible = True
'    Case 1
'        Exit Sub
'    Case 2
    '    mdiNexIRC.ctlMP3OCX.Pause
'    End Select
'End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub MenuPlay()"
End Sub

Public Sub MenuPause()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmPlayer.ctlMovieX1.PauseMovie
'Select Case lPlayback.pCurrentEngine
'Case pMp3
'    'mdiNexIRC.ctlMP3OCX.Pause
'Case pMediaPlayer
'    PauseMultimedia mdiNexIRC.lblMultimedia.Caption
'End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub MenuPause()"
End Sub

Public Sub MenuStop()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmPlayer.ctlMovieX1.StopMovie
'Select Case lPlayback.pCurrentEngine
'Case pMp3
'    'mdiNexIRC.ctlMP3OCX.Stop
'Case pMediaPlayer
'    StopMultimedia mdiNexIRC.lblMultimedia.Caption
'    lPlayback.pPlaying = False
'End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub MenuStop()"
End Sub

Public Sub PromptPlayback(lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = OpenDialog(lForm, "Audio (*.wav;*.mp3;*.wma;*.snd;*.au;*.ogg)|*.wav;*.mp3;*.wma;*.snd;*.au;*.ogg|All Files (*.*)|*.*|", "Open Audio", CurDir)
If Len(msg) <> 0 Then PlayFile msg
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub PromptPlayback(lForm As Form)"
End Sub
