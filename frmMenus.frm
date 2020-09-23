VERSION 5.00
Begin VB.Form frmMenus 
   Caption         =   "NexIRC (Menus)"
   ClientHeight    =   360
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   7290
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenus.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   360
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Menu mnuQuery123 
      Caption         =   "Query"
      Begin VB.Menu mnuWhoisQuery 
         Caption         =   "Whois"
      End
      Begin VB.Menu mnuSep8392638 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOffer123 
         Caption         =   "Offer"
         Begin VB.Menu mnuOfferCurrentlyPlaying 
            Caption         =   "Currently Playing"
         End
         Begin VB.Menu mnuSelectFileToOffer123 
            Caption         =   "Select"
         End
      End
      Begin VB.Menu mnuMessages 
         Caption         =   "Messages"
         Begin VB.Menu mnuNotice1234 
            Caption         =   "Notice"
         End
         Begin VB.Menu mnuInviteToChannel 
            Caption         =   "Invite"
         End
         Begin VB.Menu mnuSendPlaylist 
            Caption         =   "Playlist"
         End
      End
      Begin VB.Menu mnuDCC123 
         Caption         =   "DCC"
         Begin VB.Menu mnuDCCSend123 
            Caption         =   "Send"
         End
         Begin VB.Menu mnuDCCChat123 
            Caption         =   "Chat"
         End
      End
      Begin VB.Menu mnuSep38923692634 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
         Begin VB.Menu mnuAddToNotify 
            Caption         =   "Notify"
         End
         Begin VB.Menu mnuIgnoreQueryUser 
            Caption         =   "Ignore"
         End
         Begin VB.Menu mnuAddToBots 
            Caption         =   "Bots"
         End
      End
   End
   Begin VB.Menu mnuBackground 
      Caption         =   "Background"
      Begin VB.Menu mnuBG 
         Caption         =   "Background"
         Begin VB.Menu mnuBGColor 
            Caption         =   "Color"
            Begin VB.Menu mnuSelectBGColor 
               Caption         =   "Select"
            End
         End
         Begin VB.Menu mnuPicture 
            Caption         =   "Picture"
            Begin VB.Menu mnuSelectBGPicture 
               Caption         =   "Select"
            End
            Begin VB.Menu mnuSelectBGPictureNone 
               Caption         =   "None"
            End
         End
      End
      Begin VB.Menu mnuSep2893263798 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolbar 
         Caption         =   "Toolbar"
         Begin VB.Menu mnuShowHideToolbar 
            Caption         =   "Hide"
         End
         Begin VB.Menu mnuSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuToolbarPicture 
            Caption         =   "Picture"
         End
      End
      Begin VB.Menu mnuStatusBar 
         Caption         =   "Statusbar"
         Begin VB.Menu mnuShowHideStatusbar 
            Caption         =   "Hide"
         End
      End
      Begin VB.Menu mnuMenus 
         Caption         =   "Menus"
         Begin VB.Menu mnuShowHideMenus 
            Caption         =   "Hide"
         End
      End
      Begin VB.Menu mnuNOTIFY 
         Caption         =   "Notify"
         Begin VB.Menu mnuShowHideNotify 
            Caption         =   "Show"
         End
      End
      Begin VB.Menu mnuMixer 
         Caption         =   "Mixer"
         Begin VB.Menu mnuShowHideMixer 
            Caption         =   "Hide"
         End
      End
      Begin VB.Menu mnuSpectrum 
         Caption         =   "Spectrum"
         Begin VB.Menu mnuShowHideSpectrum 
            Caption         =   "Hide"
         End
         Begin VB.Menu mnuSep2389727896 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBars4 
            Caption         =   "Bands"
            Begin VB.Menu mnuBarIndex 
               Caption         =   ""
               Index           =   0
            End
         End
      End
      Begin VB.Menu mnuAll 
         Caption         =   "All"
         Begin VB.Menu mnuShowHideAll 
            Caption         =   "Show"
         End
      End
   End
   Begin VB.Menu mnuPlaylist 
      Caption         =   "Playlist"
      Begin VB.Menu mnuFileProporties 
         Caption         =   "Proporties"
      End
      Begin VB.Menu mnuSep8032798037 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlayback 
         Caption         =   "Playback"
         Begin VB.Menu mnuPlay1 
            Caption         =   "Play"
         End
         Begin VB.Menu mnuPause 
            Caption         =   "Pause"
         End
         Begin VB.Menu mnuStop 
            Caption         =   "Stop"
         End
      End
      Begin VB.Menu mnuAddPlaylist 
         Caption         =   "Add"
         Begin VB.Menu mnuAddToPlaylist 
            Caption         =   "Directories"
         End
         Begin VB.Menu mnuAddFolder 
            Caption         =   "Directory"
         End
         Begin VB.Menu mnuAddFile 
            Caption         =   "File"
         End
      End
      Begin VB.Menu mnuContinuousPlayMenu 
         Caption         =   "Continuous"
         Begin VB.Menu mnuSelectRandom 
            Caption         =   "Random"
         End
         Begin VB.Menu mnuSep89037289063892 
            Caption         =   "-"
         End
         Begin VB.Menu mnuToggleContinuousPlayON 
            Caption         =   "On"
         End
         Begin VB.Menu mnuContinuousPlayOff 
            Caption         =   "Off"
         End
      End
      Begin VB.Menu mnuSep38072890372 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuSavePlaylist 
         Caption         =   "Save"
      End
   End
   Begin VB.Menu mnuChannel 
      Caption         =   "Channel (Nicklist)"
      Begin VB.Menu mnuWhois 
         Caption         =   "Whois"
      End
      Begin VB.Menu mnuSep803278032 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOffer 
         Caption         =   "Offer"
         Begin VB.Menu mnuOfferCurrentlyPlayingMedia 
            Caption         =   "Currently Playing"
         End
         Begin VB.Menu mnuSelectFileToOffer 
            Caption         =   "Select"
         End
      End
      Begin VB.Menu mnuMessages278983 
         Caption         =   "Messages"
         Begin VB.Menu mnuQuery 
            Caption         =   "Query"
         End
         Begin VB.Menu mnuNotice 
            Caption         =   "Notice"
         End
         Begin VB.Menu mnuInvite 
            Caption         =   "Invite"
         End
         Begin VB.Menu mnuPlaylist123 
            Caption         =   "Playlist"
         End
      End
      Begin VB.Menu DCC 
         Caption         =   "DCC"
         Begin VB.Menu mnuDCCSend 
            Caption         =   "Send"
         End
         Begin VB.Menu mnuDCCChat 
            Caption         =   "Chat"
         End
      End
      Begin VB.Menu mnuSep327987302368 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKick 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuBan 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuKickBan 
         Caption         =   "Kick/Ban"
      End
      Begin VB.Menu mnuSep8932736293 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOps3 
         Caption         =   "Ops"
         Begin VB.Menu mnuOp3 
            Caption         =   "Op"
         End
         Begin VB.Menu mnuDeop 
            Caption         =   "Deop"
         End
         Begin VB.Menu mnuVoice3 
            Caption         =   "Voice"
         End
         Begin VB.Menu mnuDevocie 
            Caption         =   "Devoice"
         End
      End
      Begin VB.Menu mnuBots3 
         Caption         =   "Bots"
         Begin VB.Menu mnuUndernetX 
            Caption         =   "Undernet X"
            Begin VB.Menu mnuLOGIN 
               Caption         =   "LOGIN"
            End
         End
         Begin VB.Menu mnuEggdrop 
            Caption         =   "Eggdrop"
            Begin VB.Menu mnuOp 
               Caption         =   "OP"
            End
            Begin VB.Menu mnuIDENT 
               Caption         =   "IDENT"
            End
            Begin VB.Menu mnuVoice 
               Caption         =   "VOICE"
            End
         End
      End
      Begin VB.Menu mnuAddUser 
         Caption         =   "Add"
         Begin VB.Menu mnuAddNotify 
            Caption         =   "Notify"
         End
         Begin VB.Menu mnuAddIgnore 
            Caption         =   "Ignore"
         End
         Begin VB.Menu mnuAddBot 
            Caption         =   "Botlist"
         End
      End
   End
   Begin VB.Menu mnuConnectionManager 
      Caption         =   "Connection Manager"
      Begin VB.Menu mnuShowConnection 
         Caption         =   "SetFocus"
      End
      Begin VB.Menu mnuSep3827892632 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewConnection 
         Caption         =   "Connection"
      End
      Begin VB.Menu mnuSep38923626392627936 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSendMessageToConnection 
         Caption         =   "Send Message"
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuAddText 
         Caption         =   "Add Text"
      End
      Begin VB.Menu mnuSep389262673672 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMaximizeConnection 
         Caption         =   "Maximize"
      End
      Begin VB.Menu mnuMinimizeConnection 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuSep937289036927863 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseConnection 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuNotify4 
      Caption         =   "Notify"
      Begin VB.Menu mnuNotifySendMessage 
         Caption         =   "Message"
      End
      Begin VB.Menu mnuSep38926789362 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveFromNotify 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuShowNotifyList 
         Caption         =   "List"
      End
      Begin VB.Menu mnUSep93729863262390 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSendFileNotify 
         Caption         =   "DCC"
         Begin VB.Menu mnuNOTIFYDCCFILE 
            Caption         =   "File"
         End
         Begin VB.Menu mnuDCCChatNotify 
            Caption         =   "Chat"
         End
      End
   End
End
Attribute VB_Name = "frmMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Function GetQueryNickname(lNickName As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(lNickName) <> 0 Then
    If Left(lNickName, 1) = "@" Or Left(lNickName, 1) = "+" Then
        lNickName = Right(lNickName, Len(lNickName) - 1)
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function GetQueryNickname() As String"
End Function

Public Sub ResetBarCheck()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 9 To 48
    mnuBarIndex(i).Checked = False
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ResetBarCheck()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
Me.Icon = mdiNexIRC.Icon
For i = 8 To 48
    Load mnuBarIndex(i)
    mnuBarIndex(i).Caption = Str(i)
Next i
mnuBarIndex(0).Visible = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub mnuAddBot_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
'msg = mdiNexIRC.ActiveForm.lvwNames.SelectedItem.Text
msg = mdiNexIRC.ActiveForm.ReturnSelectedItem()
If Left(msg, 1) = "@" Or Left(msg, 1) = "+" Then
    msg = Right(msg, Len(msg) - 1)
End If
frmAddBotCommand.Show 0, Me
frmAddBotCommand.txtNickname.Text = msg
frmAddBotCommand.cboNicknameType.ListIndex = 1
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAddBot_Click()"
End Sub

Private Sub mnuAddFile_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = OpenDialog(Me, "All Files (*.*)|*.*|", "Add Media File", CurDir)
If Len(msg) <> 0 And DoesFileExist(msg) = True Then AddToPlaylist msg, True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAddFile_Click()"
End Sub

Private Sub mnuAddFolder_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PromptAddToPlaylist
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAddFolder_Click()"
End Sub

Private Sub mnuAddIgnore_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'If Len(Trim(mdiNexIRC.ActiveForm.lvwNames.SelectedItem.Text)) <> 0 Then AddToIgnore Trim(mdiNexIRC.ActiveForm.lvwNames.SelectedItem.Text)
If Len(Trim(mdiNexIRC.ActiveForm.ReturnSelectedItem())) <> 0 Then AddToIgnore Trim(mdiNexIRC.ActiveForm.ReturnSelectedItem())
'Dim i As Integer
'If Len(mdiNexIRC.ActiveForm.lvwNames.SelectedItem.Text) <> 0 Then
    
'    lIgnore.iCount = ReturnIgnoreCount + 1
'    lIgnore.iIgnore(i).iNickname = mdiNexIRC.ActiveForm.lvwNames.SelectedItem.Text
'    WriteINI lINIFiles.iIRC, "Ignore", "Count", Trim(Str(lIgnore.iCount))
'    WriteINI lINIFiles.iIRC, "Ignore", Trim(Str(lIgnore.iCount)), lIgnore.iIgnore(i).iNickname
'End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAddIgnore_Click()"
End Sub

Private Sub mnuAddNotify_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer
'msg = mdiNexIRC.ActiveForm.lvwNames.SelectedItem.Text
msg = mdiNexIRC.ActiveForm.ReturnSelectedItem()
If Len(msg) <> 0 Then i = AddNotify(msg)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAddNotify_Click()"
End Sub

Private Sub mnuAddText_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String
msg = InputBox("Message:")
DoColor ReturnStatusWindowTBox(frmConnectionManager.lstConnections.ListIndex + 1), msg
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAddText_Click()"
End Sub

Private Sub mnuAddToBots_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = GetQueryNickname(mdiNexIRC.ActiveForm.Caption)
If Len(msg) <> 0 Then
    frmAddBotCommand.Show 0, Me
    frmAddBotCommand.txtNickname.Text = msg
    frmAddBotCommand.cboNicknameType.ListIndex = 1
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAddToBots_Click()"
End Sub

Private Sub mnuAddToNotify_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer, j As Integer, n As Integer
msg = GetQueryNickname(mdiNexIRC.ActiveForm.Caption)
If Len(msg) <> 0 Then AddNotify msg
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAddToNotify_Click()"
End Sub

Private Sub mnuAddToPlaylist_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmAddMedia.Show
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAddToPlaylist_Click()"
End Sub

Private Sub mnuBan_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
'msg = mdiNexIRC.ActiveForm.lvwNames.SelectedItem.Text
msg = mdiNexIRC.ActiveForm.ReturnSelectedItem()
If Left(msg, 1) = "@" Or Left(msg, 1) = "+" Then
    msg = Right(msg, Len(msg) - 1)
End If
If Len(mdiNexIRC.ActiveForm.ReturnSelectedItem()) <> 0 And Len(mdiNexIRC.ActiveForm.Tag) <> 0 Then
    lSettings.sRetrieveAddressFromWhoisForBan = True
    lSettings.sBanChannel = mdiNexIRC.ActiveForm.Tag
    lSettings.sActiveServerForm.tcp.SendData "WHOIS " & msg & vbCrLf
End If
End Sub

Private Sub mnuBarIndex_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 8 To 48
    mnuBarIndex(i).Checked = False
Next i
'mdiNexIRC.ctlMP3OCX.Bands = Index
mnuBarIndex(Index).Checked = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuBarIndex_Click(Index As Integer)"
End Sub

Private Sub mnuCloseConnection_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UnloadStatusWindow frmConnectionManager.lstConnections.ListIndex + 1
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuCloseConnection_Click()"
End Sub

Private Sub mnuContinuousPlayOff_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmMobileMixer.chkContinuous.Enabled = False
mdiNexIRC.tmrContinuousPlay.Enabled = False
lSettings.sContinuousPlay = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuContinuousPlayOff_Click()"
End Sub

Private Sub mnuDCCChat_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = mdiNexIRC.ActiveForm.ReturnSelectedItem()
If Left(msg, 1) = "@" Or Left(msg, 1) = "+" Then
    msg = Right(msg, Len(msg) - 1)
End If
Call ProcessInput("CHAT " & msg, lSettings.sActiveServerForm.txtIncoming, lSettings.sActiveServerForm)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuDCCChat_Click()"
End Sub

Private Sub mnuDCCChatNotify_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmDCC_Chat.Show 0, Me
frmDCC_Chat.txtName.Text = frmNotify.lstNotify.Text
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuDCCChatNotify_Click()"
End Sub

Private Sub mnuDCCSend_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = mdiNexIRC.ActiveForm.ReturnSelectedItem()
If Left(msg, 1) = "@" Or Left(msg, 1) = "+" Then msg = Right(msg, Len(msg) - 1)
If Len(msg) <> 0 Then
    mdiNexIRC.ActivateDCCSendByNickname msg
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuDCCSend_Click()"
End Sub

Private Sub mnuDeop_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String
msg = mdiNexIRC.ActiveForm.ReturnSelectedItem()
If Left(msg, 1) = "@" Or Left(msg, 1) = "+" Then msg = Right(msg, Len(msg) - 1)
If Len(msg) <> 0 Then
    lSettings.sActiveServerForm.tcp.SendData "MODE " & mdiNexIRC.ActiveForm.Tag & " -o :" & msg & vbCrLf
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuOp3_Click()"
End Sub

Private Sub mnuDevocie_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String
msg = mdiNexIRC.ActiveForm.ReturnSelectedItem()
If Left(msg, 1) = "@" Or Left(msg, 1) = "+" Then msg = Right(msg, Len(msg) - 1)
If Len(msg) <> 0 Then
    lSettings.sActiveServerForm.tcp.SendData "MODE " & mdiNexIRC.ActiveForm.Tag & " -v :" & msg & vbCrLf
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuOp3_Click()"
End Sub

Private Sub mnuDisconnect_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DisconnectStatusWindow frmConnectionManager.lstConnections.ListIndex + 1
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuDisconnect_Click()"
End Sub

Private Sub mnuFileProporties_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To frmPlaylist.lstPlaylist.ListCount
    If frmPlaylist.lstPlaylist.ListCount <> i Then
        If frmPlaylist.lstPlaylist.Selected(i) = True Then
            Call ShowFileProperties(mdiNexIRC.hWnd, lFiles.fFile(FindFileIndexByFilename(frmPlaylist.lstPlaylist.List(i))).fFilename)
        End If
    Else
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuFileProporties_Click()"
End Sub

Private Sub mnuIDENT_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, msg3 As String
msg = mdiNexIRC.ActiveForm.ReturnSelectedItem()
If Len(msg) <> 0 Then
    msg2 = InputBox("Enter Password:")
    msg3 = InputBox("Enter Username:")
    If Len(msg2) <> 0 And Len(msg3) <> 0 Then
        lSettings.sActiveServerForm.SendData "PRIVMSG " & msg & " :IDENT " & msg2 & " " & msg3 & vbCrLf
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuOP_Click()"
End Sub

Private Sub mnuIgnoreQueryUser_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = GetQueryNickname(mdiNexIRC.ActiveForm.Caption)
If Len(msg) <> 0 Then
    AddToIgnore msg
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuIgnoreQueryUser_Click()"
End Sub

Private Sub mnuInvite_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String
msg = mdiNexIRC.ActiveForm.ReturnSelectedItem()
If Left(msg, 1) = "@" Or Left(msg, 1) = "+" Then msg = Right(msg, Len(msg) - 1)
If Len(msg) <> 0 Then
    msg2 = InputBox("Enter Channel:")
    lSettings.sActiveServerForm.tcp.SendData "INVITE " & msg & " :" & msg2 & vbCrLf & vbCrLf
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuInvite_Click()"
End Sub

Private Sub mnuInviteToChannel_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String
msg = GetQueryNickname(mdiNexIRC.ActiveForm.Caption)
If Len(msg) <> 0 Then
    msg2 = InputBox("Enter Channel:")
    lSettings.sActiveServerForm.tcp.SendData "INVITE " & msg & " :" & msg2 & vbCrLf & vbCrLf
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuInviteToChannel_Click()"
End Sub

Private Sub mnuKick_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = mdiNexIRC.ActiveForm.ReturnSelectedItem()
If Left(msg, 1) = "@" Then msg = Right(msg, Len(msg) - 1)
If Left(msg, 1) = "+" Then msg = Right(msg, Len(msg) - 1)
If Len(msg) <> 0 Then
    lSettings.sActiveServerForm.tcp.SendData "KICK " & mdiNexIRC.ActiveForm.Tag & " " & msg & vbCrLf
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuKick_Click()"
End Sub

Private Sub mnuKickBan_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = mdiNexIRC.ActiveForm.ReturnSelectedItem()
If Left(msg, 1) = "@" Or Left(msg, 1) = "+" Then
    msg = Right(msg, Len(msg) - 1)
End If
If Len(msg) <> 0 And Len(mdiNexIRC.ActiveForm.Tag) <> 0 Then
    lSettings.sRetrieveAddressFromWhoisForBan = True
    lSettings.sRetrieveAddressFromWhoisForKickBan = True
    lSettings.sBanChannel = mdiNexIRC.ActiveForm.Tag
    lSettings.sBanNickname = msg
    lSettings.sActiveServerForm.tcp.SendData "WHOIS " & msg & vbCrLf
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuKick_Click()"
End Sub

Private Sub mnuLOGIN_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, msg3 As String
msg2 = InputBox("Enter Username:", "Undernet X Login", GetSetting(App.Title, "Settings", "Undernet X Name", ""))
msg3 = InputBox("Enter Password:", "Undernet X Login", GetSetting(App.Title, "Settings", "Undernet X Password", ""))
If Len(msg2) <> 0 And Len(msg3) <> 0 Then
    SaveSetting App.Title, "Settings", "Undernet X Name", msg2
    SaveSetting App.Title, "Settings", "Undernet X Password", msg3
    lSettings.sActiveServerForm.SendData "PRIVMSG x@channels.undernet.org LOGIN " & msg2 & " " & msg3 & vbCrLf
    ProcessReplaceString sUndernetLogin, lSettings.sActiveServerForm.txtIncoming
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuOP_Click()"
End Sub

Private Sub mnuMaximizeConnection_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'lStatusWindows.sStatusWindow(frmConnectionManager.lstConnections.ListIndex + 1).sForm.WindowState = vbMaximized
SetStatusWindowState frmConnectionManager.lstConnections.ListIndex + 1, vbMaximized
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuMaximizeConnection_Click()"
End Sub

Private Sub mnuMenus_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If mdiNexIRC.mnuFile.Visible = True Then
    mnuShowHideMenus.Caption = "Hide"
Else
    mnuShowHideMenus.Caption = "Show"
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuMenus_Click()"
End Sub

Private Sub mnuMinimizeConnection_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'lStatusWindows.sStatusWindow(frmConnectionManager.lstConnections.ListIndex + 1).sForm.WindowState = vbMinimized
SetStatusWindowState frmConnectionManager.lstConnections.ListIndex + 1, vbMinimized
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuMinimizeConnection_Click()"
End Sub

Private Sub mnuMixer_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sShowQuickmix = True Then
    mnuShowHideMixer.Caption = "Hide"
Else
    mnuShowHideMixer.Caption = "Show"
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuMixer_Click()"
End Sub

Private Sub mnuNewConnection_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmQuickConnect.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuNewConnection_Click()"
End Sub

Private Sub mnuNotice_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = InputBox("Enter Notice:")
If Len(msg) <> 0 Then
    lSettings.sActiveServerForm.tcp.SendData "NOTICE " & mdiNexIRC.ActiveForm.ReturnSelectedItem() & " :" & msg & vbCrLf
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuNotice_Click()"
End Sub

Private Sub mnuNotify_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If mdiNexIRC.picNotify.Visible = True Then
    mnuShowHideNotify.Caption = "Hide"
Else
    mnuShowHideNotify.Caption = "Show"
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuNotify_Click()"
End Sub

Private Sub mnuNOTIFYDCCFILE_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim NewSendFileWin As frmSendFile
If Len(frmNotify.lstNotify.Text) <> 0 Then
    Set NewSendFileWin = New frmSendFile
    FileListenPort = FileListenPort + 1
    If FileListenPort > 9000 Then FileListenPort = 1560
    Load NewSendFileWin.tcpSend(FileListenPort)
    NewSendFileWin.Tag = Str(FileListenPort)
    NewSendFileWin.tcpSend(NewSendFileWin.Tag).LocalPort = NewSendFileWin.Tag
    NewSendFileWin.tcpSend(NewSendFileWin.Tag).Listen
    NewSendFileWin.Show 0, Me
    NewSendFileWin.txtNickname.Text = frmNotify.lstNotify.Text
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuNOTIFYDCCFILE_Click()"
End Sub

Private Sub mnuNotifySendMessage_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = InputBox("Send Message: ")
If Len(msg) <> 0 Then
    If Len(frmNotify.lstNotify.Text) <> 0 Then
        lSettings.sActiveServerForm.tcp.SendData "PRIVMSG " & frmNotify.lstNotify.Text & " :" & msg & vbCrLf
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuNotifySendMessage_Click()"
End Sub

Private Sub mnuOfferCurrentlyPlaying_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer
msg = lPlayback.pCurrentFile
If Len(msg) <> 0 Then
    i = FindFileIndexByFilename(msg)
    If i <> 0 Then
        msg2 = "PRIVMSG " & GetQueryNickname(mdiNexIRC.ActiveForm.Caption) & " :• (File Offer) Type: !" & msg & " (To Recieve File) (" & Format(FileLen(lFiles.fFile(i).fFilename), "###,###,###") & " KB)" & "(" & mdiNexIRC.lblKHZ.Caption & ")"
        lSettings.sActiveServerForm.tcp.SendData msg2 & vbCrLf
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuOfferCurrentlyPlaying_Click()"
End Sub

Private Sub mnuOfferCurrentlyPlayingMedia_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer
msg = lPlayback.pCurrentFile
If Len(msg) <> 0 Then
    i = FindFileIndexByFilename(msg)
    If i <> 0 Then
        msg2 = "PRIVMSG " & mdiNexIRC.ActiveForm.ReturnSelectedItem() & " :• (File Offer) Type: !" & msg & " (To Recieve File) (" & Format(FileLen(lFiles.fFile(i).fFilename), "###,###,###") & " KB)" & "(" & mdiNexIRC.lblKHZ.Caption & ")"
        lSettings.sActiveServerForm.tcp.SendData msg2 & vbCrLf
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuOfferCurrentlyPlayingMedia_Click()"
End Sub

Private Sub mnuOP_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String
msg = mdiNexIRC.ActiveForm.ReturnSelectedItem()
If Left(msg, 1) = "@" Or Left(msg, 1) = "+" Then msg = Right(msg, Len(msg) - 1)
If Len(msg) <> 0 Then
    msg2 = InputBox("Enter Password:")
    If Len(msg2) <> 0 Then lSettings.sActiveServerForm.tcp.SendData "PRIVMSG " & msg & " : SetOp " & msg2 & vbCrLf
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuOP_Click()"
End Sub

Private Sub mnuOp3_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String
msg = mdiNexIRC.ActiveForm.ReturnSelectedItem()
If Left(msg, 1) = "@" Or Left(msg, 1) = "+" Then msg = Right(msg, Len(msg) - 1)
If Len(msg) <> 0 Then
    lSettings.sActiveServerForm.tcp.SendData "MODE " & mdiNexIRC.ActiveForm.Tag & " +o :" & msg & vbCrLf
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuOp3_Click()"
End Sub

Private Sub mnuPause_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
MenuPause
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuPause_Click()"
End Sub

Private Sub mnuPlay1_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, F As Integer
For i = 0 To frmPlaylist.lstPlaylist.ListCount
    If frmPlaylist.lstPlaylist.ListCount <> i Then
        If frmPlaylist.lstPlaylist.Selected(i) = True Then F = F + 1
    Else
        Exit For
    End If
Next i
If F = 0 Then
    Exit Sub
ElseIf F = 1 Then
    PlayFile lFiles.fFile(FindFileIndex(frmPlaylist.lstPlaylist.Text)).fFilename
Else
    If lSettings.sGeneralPrompts = True Then
        MsgBox "NexIRC can only play one file at a time.", vbExclamation
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuPlay1_Click()"
End Sub

Private Sub mnuPlaylist123_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
mdiNexIRC.tmrSendUserPlaylist.Enabled = True
mdiNexIRC.SetEUsername mdiNexIRC.ActiveForm.ReturnSelectedItem(), lSettings.sActiveServerForm
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuQuery_Click()"
End Sub

Private Sub mnuQuery_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
NewQuery mdiNexIRC.ActiveForm.ReturnSelectedItem()
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuQuery_Click()"
End Sub

Private Sub mnuRemove_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
'lockwindowupdate frmPlaylist.hwnd
For i = 0 To frmPlaylist.lstPlaylist.ListCount
    If frmPlaylist.lstPlaylist.ListCount <> i Then
        If frmPlaylist.lstPlaylist.Selected(i) = True Then
            RemoveFromPlaylist frmPlaylist.lstPlaylist.List(i)
            frmPlaylist.lstPlaylist.RemoveItem i
            i = i - 1
        End If
    Else
        Exit For
    End If
Next i
SavePlaylist
'lockwindowupdate 0
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuRemove_Click()"
End Sub

Private Sub mnuRemoveFromNotify_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(frmNotify.lstNotify.Text) <> 0 Then
    frmNotify.lstNotify.RemoveItem frmNotify.lstNotify.ListIndex
    RemoveFromNotify frmNotify.lstNotify.Text
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuRemoveFromNotify_Click()"
End Sub

Private Sub mnuSavePlaylist_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ClearPlaylistMemory
SavePlaylist
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuSavePlaylist_Click()"
End Sub

Private Sub mnuSelectBGColor_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmSelectShape.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuSelectBGColor_Click()"
End Sub

Private Sub mnuSelectBGPicture_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = OpenDialog(mdiNexIRC, "Jpeg Files (*.jpg)|*.jpg|Compuserv GIF Files (*.gif)|*.gif|All Files (*.*)|*.*|", App.Title, App.Path & "\data\images\")
If Len(msg) <> 0 Then
    If DoesFileExist(msg) = True Then
        lSettings.sBGPicture = msg
        mdiNexIRC.Picture = LoadPicture(msg)
        WriteINI GetINIFile(iIRC), "Settings", "BGPicture", lSettings.sBGPicture
    Else
        If lSettings.sGeneralPrompts = True Then
            MsgBox "Could not locate " & msg, vbExclamation
        End If
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuSelectBGPicture_Click()"
End Sub

Private Sub mnuSelectBGPictureNone_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sBGPicture = ""
WriteINI GetINIFile(iIRC), "Settings", "BGPicture", lSettings.sBGPicture
mdiNexIRC.Picture = Nothing
mdiNexIRC.BackColor = vbBlack
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuSelectBGPictureNone_Click()"
End Sub

Private Sub mnuSelectFileToOffer_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = OpenDialog(mdiNexIRC, "All Files (*.*)|*.*|", "Select Media", CurDir)
If Len(msg) <> 0 Then
    If FindFileIndexByFilename(msg) <> 0 Then
        msg = GetFileTitle(msg)
        lSettings.sActiveServerForm.tcp.SendData "PRIVMSG " & mdiNexIRC.ActiveForm.ReturnSelectedItem() & " :15• 4(File Offer) Type: 7!" & msg & " 4(To Recieve File) (" & Format(FileLen(lFiles.fFile(FindFileIndexByFilename(msg)).fFilename), "###,###,###") & " KB)" & "(" & mdiNexIRC.lblKHZ.Caption & ")" & vbCrLf
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuOfferCurrentlyPlayingMedia_Click()"
End Sub

Private Sub mnuSelectFileToOffer123_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = OpenDialog(mdiNexIRC, "All Files (*.*)|*.*|", "Select Media", CurDir)
If Len(msg) <> 0 Then
    If FindFileIndexByFilename(msg) <> 0 Then
        msg = GetFileTitle(msg)
        lSettings.sActiveServerForm.tcp.SendData "PRIVMSG " & GetQueryNickname(mdiNexIRC.ActiveForm.Caption) & " :15• 4(File Offer) Type: 7!" & msg & " 4(To Recieve File) (" & Format(FileLen(lFiles.fFile(FindFileIndexByFilename(msg)).fFilename), "###,###,###") & " KB)" & "(" & mdiNexIRC.lblKHZ.Caption & ")" & vbCrLf
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuSelectFileToOffer123_Click()"
End Sub

Private Sub mnuSelectRandom_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
GetR:
i = GetRnd(lFiles.fCount)
If DoesFileExist(lFiles.fFile(i).fFilename) = True Then
    MenuStop
    PlayFile lFiles.fFile(i).fFilename
Else
    GoTo GetR
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuSelectRandom_Click()"
End Sub

Private Sub mnuSendMessageToConnection_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String
msg = InputBox("Message:")
If ReturnStatusWindowTCPState(frmConnectionManager.lstConnections.ListIndex + 1) = sckConnected Then StatusWindowSendData frmConnectionManager.lstConnections.ListIndex + 1, msg & vbCrLf
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuSendMessageToConnection_Click()"
End Sub

Private Sub mnuShowConnection_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim re
re = ShowWindow(ReturnStatusWindowHwnd(frmConnectionManager.lstConnections.ListIndex + 1), 9)
SetStatusWindowFocus frmConnectionManager.lstConnections.ListIndex + 1
SetStatusWindowState frmConnectionManager.lstConnections.ListIndex + 1, vbNormal
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuShowConnection_Click()"
End Sub

Private Sub mnuShowHideAll_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If mnuShowHideAll.Caption = "Show" Then
    mnuShowHideAll.Caption = "Hide"
    With mdiNexIRC
        .picNotify.Visible = True
        .picMP3OCX.Visible = True
        .picTopToolbar.Visible = True
        .StatusBar.Visible = True
        .mnuFile.Visible = True
        .mnuTools.Visible = True
        .mnuWindow.Visible = True
        .mnuHelp.Visible = True
    End With
ElseIf mnuShowHideAll.Caption = "Hide" Then
    mnuShowHideAll.Caption = "Show"
    With mdiNexIRC
        .picNotify.Visible = False
        .picMP3OCX.Visible = False
        .picTopToolbar.Visible = False
        .StatusBar.Visible = False
        .mnuFile.Visible = False
        .mnuTools.Visible = False
        .mnuWindow.Visible = False
        .mnuHelp.Visible = False
    End With
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuShowHideAll_Click()"
End Sub

Private Sub mnuShowHideMenus_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Trim(LCase(mnuShowHideMenus.Caption)) = "show" Then
    mnuShowHideMenus.Caption = "Hide"
    mdiNexIRC.mnuFile.Visible = True
    mdiNexIRC.mnuTools.Visible = True
    mdiNexIRC.mnuHelp.Visible = True
    mdiNexIRC.mnuWindow.Visible = True
ElseIf Trim(LCase(mnuShowHideMenus.Caption)) = "hide" Then
    mnuShowHideMenus.Caption = "Show"
    mdiNexIRC.mnuFile.Visible = False
    mdiNexIRC.mnuTools.Visible = False
    mdiNexIRC.mnuHelp.Visible = False
    mdiNexIRC.mnuWindow.Visible = False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuShowHideMenus_Click()"
End Sub

Private Sub mnuShowHideMixer_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sMobileMixerVisible = True Then
    ToggleMixer False
Else
    ToggleMixer True
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuShowHideMixer_Click()"
End Sub

Private Sub mnuShowHideNotify_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If mdiNexIRC.picNotify.Visible = True Then
    mdiNexIRC.picNotify.Visible = False
    lSettings.sShowQuickNotify = False
Else
    mdiNexIRC.picNotify.Visible = True
    lSettings.sShowQuickNotify = True
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuShowHideNotify_Click()"
End Sub

Private Sub mnuShowHideSpectrum_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If mdiNexIRC.picMP3OCX.Visible = True Then
    mdiNexIRC.picMP3OCX.Visible = False
    mnuShowHideSpectrum.Caption = "Show"
Else
    mdiNexIRC.picMP3OCX.Visible = True
    mnuShowHideSpectrum.Caption = "Hide"
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuShowHideSpectrum_Click()"
End Sub

Private Sub mnuShowHideStatusbar_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If mdiNexIRC.StatusBar.Visible = True Then
    mdiNexIRC.StatusBar.Visible = False
    mnuShowHideStatusbar.Caption = "Show"
Else
    mdiNexIRC.StatusBar.Visible = True
    mnuShowHideStatusbar.Caption = "Hide"
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuShowHideStatusbar_Click()"
End Sub

Private Sub mnuShowHideToolbar_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If mdiNexIRC.picTopToolbar.Visible = True Then
    mdiNexIRC.picTopToolbar.Visible = False
    mnuShowHideToolbar.Caption = "Show"
Else
    mdiNexIRC.picTopToolbar.Visible = True
    mnuShowHideToolbar.Caption = "Hide"
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuShowHideToolbar_Click()"
End Sub

Private Sub mnuShowNotifyList_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
frmCustomize.Show 0, Me
For i = 0 To 9
    frmCustomize.optCheck(i).Value = False
    frmCustomize.fraSettings(i).Visible = False
Next i
frmCustomize.optCheck(3).Value = True
frmCustomize.fraSettings(3).Visible = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuShowNotifyList_Click()"
End Sub

Private Sub mnuSpectrum_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If mdiNexIRC.picMP3OCX.Visible = True Then
    mnuShowHideSpectrum.Caption = "Hide"
Else
    mnuShowHideSpectrum.Caption = "Show"
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuSpectrum_Click()"
End Sub

Private Sub mnuStatusBar_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If mdiNexIRC.StatusBar.Visible = True Then
    mnuShowHideStatusbar.Caption = "Hide"
Else
    mnuShowHideStatusbar.Caption = "Show"
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuStatusBar_Click()"
End Sub

Private Sub mnuStop_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
MenuStop
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuStop_Click()"
End Sub

Private Sub mnuToggleContinuousPlayON_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmMobileMixer.chkContinuous.Enabled = True
mdiNexIRC.tmrContinuousPlay.Enabled = True
lSettings.sContinuousPlay = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuToggleContinuousPlayON_Click()"
End Sub

Private Sub mnuToolbar_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If mdiNexIRC.picTopToolbar.Visible = True Then
    mnuShowHideToolbar.Caption = "Hide"
Else
    mnuShowHideToolbar.Caption = "Show"
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuToolbar_Click()"
End Sub

Private Sub mnuToolbarPicture_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = OpenDialog(mdiNexIRC, "Compuserv Gif (*.gif)|*.gif|All Files (*.*)|*.*|", "Select Toolbar images", App.Path & "\data\images\")
If Len(msg) <> 0 Then
    If DoesFileExist(msg) = True Then
        mdiNexIRC.picTopToolbar.Picture = LoadPicture(msg)
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuToolbarPicture_Click()"
End Sub

Private Sub MNUVOICE_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String
msg = mdiNexIRC.ActiveForm.ReturnSelectedItem()
If Len(msg) <> 0 Then
    msg2 = InputBox("Enter Password:")
    If Len(msg2) <> 0 Then
        lSettings.sActiveServerForm.SendData "PRIVMSG " & msg & " :SetVoice " & msg2 & vbCrLf
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuOP_Click()"
End Sub

Private Sub mnuVoice3_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String
msg = mdiNexIRC.ActiveForm.ReturnSelectedItem()
If Left(msg, 1) = "@" Or Left(msg, 1) = "+" Then msg = Right(msg, Len(msg) - 1)
If Len(msg) <> 0 Then
    lSettings.sActiveServerForm.tcp.SendData "MODE " & mdiNexIRC.ActiveForm.Tag & " +v :" & msg & vbCrLf
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuOp3_Click()"
End Sub

Private Sub mnuWallpaperAL_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuWallpaperAL_Click(Index As Integer)"
End Sub

Private Sub mnuWhois_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = mdiNexIRC.ActiveForm.ReturnSelectedItem()
If Left(msg, 1) = "@" Then msg = Right(msg, Len(msg) - 1)
If Left(msg, 1) = "+" Then msg = Right(msg, Len(msg) - 1)
If Len(msg) <> 0 Then
    lSettings.sActiveServerForm.tcp.SendData "WHOIS " & msg & vbCrLf
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuWhois_Click()"
End Sub

Private Sub mnuWhoisQuery_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = GetQueryNickname(mdiNexIRC.ActiveForm.Caption)
If Len(msg) <> 0 Then
    lSettings.sActiveServerForm.tcp.SendData "WHOIS " & msg & vbCrLf
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuWhois_Click()"
End Sub
