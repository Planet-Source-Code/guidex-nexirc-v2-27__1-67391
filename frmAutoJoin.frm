VERSION 5.00
Begin VB.Form frmAutoJoin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Auto Join"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAutoJoin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstChannels 
      BackColor       =   &H00FFFFFF&
      Height          =   2205
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.ComboBox cboNetworks 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1320
      MousePointer    =   1  'Arrow
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin nexIRC.ctlXPButton cmdHelp 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Help"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmAutoJoin.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdJoin 
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Join"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmAutoJoin.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmAutoJoin.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Add"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmAutoJoin.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdRemove 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Remove"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmAutoJoin.frx":007C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdClear 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Clear"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmAutoJoin.frx":0098
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdActivate 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Activate"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmAutoJoin.frx":00B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblNetwork 
      Caption         =   "&Network:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActivate 
         Caption         =   "A&ctivate"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowToUseThisWindow 
         Caption         =   "H&ow to use this Window"
      End
   End
End
Attribute VB_Name = "frmAutoJoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboNetworks_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
FillListBoxWithAutoJoin lstChannels
For i = 0 To ReturnAutoJoinCount
    If Len(ReturnAutoJoinChannel(i)) <> 0 Then
        If LCase(ReturnAutoJoinNetwork(i)) = LCase(cboNetworks.Text) Then
            lstChannels.AddItem ReturnAutoJoinChannel(i)
        End If
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cboNetworks_Click()"
End Sub

Private Sub cmdActivate_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ActivateAutoJoin True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdActivate_Click()"
End Sub

Private Sub cmdAdd_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lNetwork As String, lChannel As String, i As Integer
lNetwork = cboNetworks.Text
If Len(lNetwork) <> 0 Then
    lChannel = InputBox("Enter channel name:")
    If Len(lChannel) <> 0 Then
        i = AddAutoJoin(lChannel, lNetwork)
        lstChannels.AddItem lChannel
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAdd_Click()"
End Sub

Private Sub cmdCancel_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdCancel_Click()"
End Sub

Private Sub cmdClear_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mbox As VbMsgBoxResult
If lSettings.sGeneralPrompts = True Then
    mbox = MsgBox("Clear all auto join channels?", vbYesNo + vbQuestion)
Else
    mbox = vbYes
End If
If mbox = vbYes Then
    lstChannels.Clear
    Kill GetINIFile(iAutoJoin)
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdClear_Click()"
End Sub

Private Sub cmdHelp_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DisplayHelpInformation 6
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdHelp_Click()"
End Sub

Private Sub cmdJoin_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.SendData "JOIN :" & lstChannels.Text & vbCrLf
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdJoin_Click()"
End Sub

'Private Sub cmdJoin_Click()
'Unload Me
'If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdOK_Click()"
'End Sub

Private Sub cmdRemove_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
lstChannels.RemoveItem lstChannels.ListIndex
i = FindAutoJoinIndex(lstChannels.Text)
If i <> 0 Then
    DeleteAutoJoin i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdRemove_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
Me.Icon = mdiNexIRC.Icon
SetButtonType cmdAdd
SetButtonType cmdRemove
SetButtonType cmdClear
SetButtonType cmdActivate
SetButtonType cmdHelp
SetButtonType cmdJoin
SetButtonType cmdCancel
For i = 0 To lServers.sNetworkCount
    If Len(lServers.sNetwork(i).nDescription) <> 0 Then cboNetworks.AddItem lServers.sNetwork(i).nDescription
Next i
cboNetworks.Text = GetSetting(App.Title, "Settings", "AutoJoinLastNetwork", "Undernet")
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
SaveSetting App.Title, "Settings", "AutoJoinLastNetwork", cboNetworks.Text
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub mnuActivate_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdActivate_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuActivate_Click()"
End Sub

Private Sub mnuAdd_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdAdd_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAdd_Click()"
End Sub

Private Sub mnuClear_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdClear_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuClear_Click()"
End Sub

Private Sub mnuExit_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuExit_Click()"
End Sub

Private Sub mnuHowToUseThisWindow_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdHelp_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuHowToUseThisWindow_Click()"
End Sub

Private Sub mnuRemove_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdRemove_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuRemove_Click()"
End Sub
