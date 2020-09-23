VERSION 5.00
Begin VB.Form frmPlayer 
   Caption         =   "NexIRC - Video Player"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlayer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraControls 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   6615
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   300
         Left            =   120
         TabIndex        =   14
         Top             =   440
         Width           =   735
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         Height          =   300
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   300
         Left            =   960
         TabIndex        =   12
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdRewind 
         Caption         =   "Rewind"
         Height          =   300
         Left            =   1800
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdForward 
         Caption         =   "Forward"
         Height          =   300
         Left            =   1800
         TabIndex        =   10
         Top             =   435
         Width           =   735
      End
      Begin VB.CommandButton cmdMute 
         Caption         =   "Mute"
         Height          =   300
         Left            =   2640
         TabIndex        =   9
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdFullScreen 
         Caption         =   "Full Scr."
         Height          =   300
         Left            =   2640
         TabIndex        =   8
         Top             =   435
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Open CD Door"
         Height          =   300
         Left            =   4320
         TabIndex        =   7
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdCloseCDDoor 
         Caption         =   "Close CD Door"
         Height          =   300
         Left            =   4320
         TabIndex        =   6
         Top             =   440
         Width           =   1215
      End
      Begin VB.CommandButton cmdChangePlayRate 
         Caption         =   "Rate"
         Height          =   300
         Left            =   3480
         TabIndex        =   5
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pause"
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   435
         Width           =   735
      End
      Begin VB.HScrollBar scrVolume 
         Height          =   255
         LargeChange     =   25
         Left            =   5640
         Max             =   100
         Min             =   1
         SmallChange     =   5
         TabIndex        =   3
         Top             =   120
         Value           =   1
         Width           =   855
      End
      Begin VB.HScrollBar scrProgress 
         Height          =   255
         LargeChange     =   25
         Left            =   5640
         Max             =   100
         Min             =   1
         SmallChange     =   5
         TabIndex        =   2
         Top             =   480
         Value           =   1
         Width           =   855
      End
   End
   Begin nexIRC.ctlMovieX ctlMovieX1 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2566
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private lScrolling As Boolean

'Private Sub cmdAuthorize_Click()
'Dim msg As String, msg2 As String
'msg = GetSetting(App.Title, "Settings", "UserName", "")
'msg2 = GetSetting(App.Title, "Settings", "Password", "")
'If Len(msg) = 0 And Len(msg2) = 0 Then
'    msg = InputBox("Enter UserName:")
'    If Len(msg) = 0 Then Exit Sub
'    msg2 = InputBox("Enter Password:")
'    If Len(msg2) = 0 Then Exit Sub
'End If
'ctlMovieX1.Authorize msg, msg2
'End Sub

Private Sub cmdChangePlayRate_Click()
Dim l As Long
l = CLng(InputBox("Enter Playrate:"))
ctlMovieX1.ChangePlayRate l
End Sub

Private Sub cmdCloseCDDoor_Click()
ctlMovieX1.CloseDoor
End Sub

'Private Sub cmdForward_Click()
'If tmrFastForward.Enabled = True Then
'    tmrFastForward.Enabled = False
'Else
'    tmrFastForward.Enabled = True
'End If
'End Sub

Private Sub cmdFullScreen_Click()
ctlMovieX1.FullScreen
End Sub

Private Sub cmdMute_Click()
If Len(cmdMute.Tag) = 0 Then
    cmdMute.Tag = "Muted"
    ctlMovieX1.Mute True
Else
    cmdMute.Tag = ""
    ctlMovieX1.Mute False
End If
End Sub

Private Sub cmdOpen_Click()
ctlMovieX1.OpenMovieDialog Me, "Supported Files (*.m4a;*.avi;*.mpg;*.mpeg;*.mpe;*.mp3;*.mp2;*.mp1;*.wav;*.aif;*.aiff;*.aifc;*.au;*.mv1;*.mov;*.mpa;*.qt;*.snd;*.mpm;*.mpv;*.enc;*.mid;*.rmi;*.vob;*.wma;*.wmv)|*.m4a;*.avi;*.mpg;*.mpeg;*.mpe;*.mp3;*.mp2;*.mp1;*.wav;*.aif;*.aiff;*.aifc;*.au;*.mv1;*.mov;*.mpa;*.qt;*.snd;*.mpm;*.mpv;*.enc;*.mid;*.rmi;*.vob;*.wma;*.wmv|", "MovieX Sample", CurDir
End Sub

Private Sub cmdPause_Click()
ctlMovieX1.PauseMovie
End Sub

Private Sub cmdPlay_Click()
Dim l As Long
ctlMovieX1.PlayMovie
Form_Resize
l = ctlMovieX1.ReturnTotalSeconds()
scrProgress.Max = l
End Sub

'Private Sub cmdRewind_Click()
'If tmrRewind.Enabled = True Then
'    tmrRewind.Enabled = False
'Else
'    tmrRewind.Enabled = True
'End If
'End Sub

Private Sub cmdStop_Click()
ctlMovieX1.StopMovie
End Sub

Private Sub Command2_Click()
ctlMovieX1.OpenCDDoor
End Sub

Private Sub Form_Load()
Form_Resize
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
fraControls.Top = Me.ScaleHeight - fraControls.Height
fraControls.Width = Me.ScaleWidth
ctlMovieX1.Width = Me.ScaleWidth
ctlMovieX1.Height = Me.ScaleHeight - (fraControls.Height)
ctlMovieX1.SetSize 0, 0, Me.ScaleWidth / Screen.TwipsPerPixelX, Me.ScaleHeight / Screen.TwipsPerPixelY - (fraControls.Height / Screen.TwipsPerPixelY)
End Sub

Private Sub Form_Unload(Cancel As Integer)
ctlMovieX1.StopMovie
ctlMovieX1.CloseMovie
End Sub

Private Sub scrProgress_Change()
ctlMovieX1.ChangeMoviePosition scrProgress.Value
End Sub

Private Sub scrVolume_Change()
ctlMovieX1.SetVolume scrVolume.Value * 10
End Sub

Private Sub scrVolume_Scroll()
ctlMovieX1.SetVolume scrVolume.Value * 10
End Sub

Private Sub tmrFastForward_Timer()
ctlMovieX1.ForwardFrames 80
End Sub

Private Sub tmrRewind_Timer()
ctlMovieX1.RewindFrames 80
End Sub

