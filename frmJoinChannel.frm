VERSION 5.00
Begin VB.Form frmJoinChannel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Join Channel"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   4365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmJoinChannel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4365
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtKey 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtJoinChannel 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.CheckBox chkAddToAutoJoin 
      Appearance      =   0  'Flat
      Caption         =   "Add to &Auto Join"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CheckBox chkAddToChanFolder 
      Appearance      =   0  'Flat
      Caption         =   "Add to Channel &Folder"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   840
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin nexIRC.ctlXPButton cmdHelp 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1440
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
      MICON           =   "frmJoinChannel.frx":000C
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
      TabIndex        =   7
      Top             =   1440
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
      MICON           =   "frmJoinChannel.frx":0028
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
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   1440
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
      MICON           =   "frmJoinChannel.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblKey 
      Caption         =   "&Key:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblChannel 
      Caption         =   "&Channel:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuAdditional 
      Caption         =   "&Additional"
      Begin VB.Menu mnuChannelFolder 
         Caption         =   "Channel Folder"
      End
      Begin VB.Menu mnuAutoJoin 
         Caption         =   "Auto Join"
      End
      Begin VB.Menu mnuChannelProporties 
         Caption         =   "Channel &Listing"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuAddToChannelFolder 
         Caption         =   "&Add to Channel Folder"
      End
      Begin VB.Menu mnuAddToAutoJoin 
         Caption         =   "&Add to Auto Join"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowToUseThisWindow 
         Caption         =   "H&ow to use this Window"
      End
   End
End
Attribute VB_Name = "frmJoinChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdCancel_Click()"
End Sub

Private Sub cmdHelp_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DisplayHelpInformation 21
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdHelp_Click()"
End Sub

Private Sub cmdJoin_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If chkAddToAutoJoin.Value = 1 Then AddAutoJoin txtJoinChannel.Text, lSettings.sNetwork
If chkAddToChanFolder.Value = 1 Then AddtoChanFolder txtJoinChannel.Text
If Len(txtKey.Text) = 0 Then
    lSettings.sActiveServerForm.tcp.SendData "JOIN " & txtJoinChannel.Text & vbCrLf
Else
    lSettings.sActiveServerForm.tcp.SendData "JOIN " & txtJoinChannel.Text & " " & txtKey.Text & vbCrLf
End If
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdJoin_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = mdiNexIRC.Icon
SetButtonType cmdHelp
SetButtonType cmdJoin
SetButtonType cmdCancel
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub mnuAutoJoin_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmAutoJoin.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAutoJoin_Click()"
End Sub

Private Sub mnuChannelFolder_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmChannelFolder.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuChannelFolder_Click()"
End Sub

Private Sub mnuChannelProporties_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmChannelListing.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuChannelProporties_Click()"
End Sub

Private Sub mnuExit_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub mnuHowToUseThisWindow_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdHelp_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuHowToUseThisWindow_Click()"
End Sub
