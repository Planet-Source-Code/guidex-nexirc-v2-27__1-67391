VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiNexIRC 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "NexIRC"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   870
   ClientWidth     =   11385
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   ScrollBars      =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox picMobileMixer 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7260
      Left            =   0
      MousePointer    =   9  'Size W E
      ScaleHeight     =   7260
      ScaleWidth      =   975
      TabIndex        =   33
      Top             =   450
      Visible         =   0   'False
      Width           =   975
      Begin VB.Timer tmrUnloadSplashDelay 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   480
         Top             =   0
      End
      Begin MSWinsockLib.Winsock wskChat2 
         Index           =   0
         Left            =   0
         Top             =   3720
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wskChat 
         Index           =   0
         Left            =   0
         Top             =   3360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wskIdent 
         Left            =   0
         Top             =   3000
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer timeParse 
         Left            =   0
         Top             =   1080
      End
      Begin VB.Timer tmrNotify 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   0
         Top             =   0
      End
      Begin VB.FileListBox file1 
         Height          =   480
         Left            =   0
         TabIndex        =   36
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Timer tmrContinuousPlay 
         Enabled         =   0   'False
         Interval        =   2500
         Left            =   0
         Top             =   1440
      End
      Begin VB.FileListBox filLogs 
         Height          =   870
         Left            =   0
         TabIndex        =   34
         Top             =   6480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Timer tmrPlaySoon 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   720
      End
      Begin VB.Timer tmrCheckButtonColors 
         Interval        =   900
         Left            =   0
         Top             =   360
      End
      Begin VB.Timer tmrEndSoon 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   0
         Top             =   1800
      End
      Begin VB.Timer tmrSendUserPlaylist 
         Enabled         =   0   'False
         Interval        =   1500
         Left            =   0
         Top             =   2160
      End
      Begin MSComctlLib.ImageList imgTaskbar 
         Left            =   0
         Top             =   4800
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   14
         ImageHeight     =   12
         MaskColor       =   16711935
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":0CCA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":1D1C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox ctlVBScript 
         Height          =   480
         Left            =   0
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   40
         Top             =   4200
         Width           =   1200
      End
      Begin MSComctlLib.ImageList imlToolbarIcons2 
         Left            =   0
         Top             =   5880
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   18
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":2D6E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":3108
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":34A2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":383C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":3BD6
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":3F70
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":430A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":46A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":4A3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":4DD8
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":5E2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":6B04
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":6E9E
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":7238
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":75D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":7744
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":78B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":7C50
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Timer tmrDIE 
         Enabled         =   0   'False
         Interval        =   400
         Left            =   0
         Top             =   2520
      End
   End
   Begin VB.PictureBox picNotify 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   11385
      TabIndex        =   7
      Top             =   8025
      Visible         =   0   'False
      Width           =   11385
      Begin nexIRC.ctlXPButton cmdSend 
         Height          =   330
         Left            =   10440
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   582
         BTYPE           =   2
         TX              =   ""
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
         MICON           =   "mdiMain.frx":7FEA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtMessage 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4200
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   6135
      End
      Begin VB.ComboBox cboNotify 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "mdiMain.frx":8006
         Left            =   720
         List            =   "mdiMain.frx":8008
         Style           =   2  'Dropdown List
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   30
         Width           =   2175
      End
      Begin VB.Label lblSendMessage 
         BackStyle       =   0  'Transparent
         Caption         =   "Send Message:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   60
         Width           =   1455
      End
      Begin VB.Label lblNotify 
         BackStyle       =   0  'Transparent
         Caption         =   "Notify:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.PictureBox picMP3OCX 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   11385
      TabIndex        =   4
      Top             =   7710
      Visible         =   0   'False
      Width           =   11385
      Begin nexIRC.ctlXPButton cmdDelete 
         Height          =   330
         Left            =   10440
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   582
         BTYPE           =   2
         TX              =   ""
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
         MICON           =   "mdiMain.frx":800A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdSave 
         Height          =   330
         Left            =   9480
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   582
         BTYPE           =   2
         TX              =   ""
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
         MICON           =   "mdiMain.frx":8026
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox cboValue 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "mdiMain.frx":8042
         Left            =   4560
         List            =   "mdiMain.frx":8044
         Style           =   2  'Dropdown List
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   0
         Width           =   20000
      End
      Begin VB.ComboBox cboProporties 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "mdiMain.frx":8046
         Left            =   2280
         List            =   "mdiMain.frx":8086
         Style           =   2  'Dropdown List
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   0
         Width           =   20000
      End
      Begin VB.ComboBox cboColors 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "mdiMain.frx":819B
         Left            =   0
         List            =   "mdiMain.frx":8238
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Colors"
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox cboSpectrumThemes 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Current Theme"
         Top             =   0
         Width           =   20000
      End
   End
   Begin VB.PictureBox picTopToolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   11385
      TabIndex        =   1
      Top             =   0
      Width           =   11385
      Begin VB.PictureBox picExit 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         MouseIcon       =   "mdiMain.frx":8732
         MousePointer    =   99  'Custom
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Exit NexIRC"
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picNexIRC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3000
         MouseIcon       =   "mdiMain.frx":8884
         MousePointer    =   99  'Custom
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Show About"
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picForward 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3360
         MouseIcon       =   "mdiMain.frx":89D6
         MousePointer    =   99  'Custom
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Forward in Playlist/URL"
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picPlay 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         MouseIcon       =   "mdiMain.frx":8B28
         MousePointer    =   99  'Custom
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Play"
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picPause 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4080
         MouseIcon       =   "mdiMain.frx":8C7A
         MousePointer    =   99  'Custom
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Pause"
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picStop 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         MouseIcon       =   "mdiMain.frx":8DCC
         MousePointer    =   99  'Custom
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Stop"
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picBackward 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         MouseIcon       =   "mdiMain.frx":8F1E
         MousePointer    =   99  'Custom
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Backward in Playlist/URL"
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picChat 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         MouseIcon       =   "mdiMain.frx":9070
         MousePointer    =   99  'Custom
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "DCC Chat"
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picSend 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4800
         MouseIcon       =   "mdiMain.frx":91C2
         MousePointer    =   99  'Custom
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Send File"
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picScript 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5160
         MouseIcon       =   "mdiMain.frx":9314
         MousePointer    =   99  'Custom
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "New Script"
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picChannelFolder 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         MouseIcon       =   "mdiMain.frx":9466
         MousePointer    =   99  'Custom
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Channel Folder"
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picOptions 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         MouseIcon       =   "mdiMain.frx":95B8
         MousePointer    =   99  'Custom
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Settings"
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picAudio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         MouseIcon       =   "mdiMain.frx":970A
         MousePointer    =   99  'Custom
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Play Music"
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtUrl 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5520
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picConnect 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         MouseIcon       =   "mdiMain.frx":985C
         MousePointer    =   99  'Custom
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picDisconnect 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         MouseIcon       =   "mdiMain.frx":99AE
         MousePointer    =   99  'Custom
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Disconnect"
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblFrames 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   12000
         TabIndex        =   23
         Top             =   120
         Width           =   45
      End
      Begin VB.Label lblFilename2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   11160
         TabIndex        =   22
         Top             =   120
         Width           =   45
      End
      Begin VB.Label lblKHZ 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4800
         TabIndex        =   13
         ToolTipText     =   "Bitrate, Number of channels, Samples per second"
         Top             =   225
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   15
         Left            =   8040
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7260
      Left            =   10755
      ScaleHeight     =   7260
      ScaleWidth      =   630
      TabIndex        =   35
      Top             =   450
      Visible         =   0   'False
      Width           =   630
      Begin VB.Label lblMultimedia 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   8385
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   529
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Begin VB.Menu mnuNewConnection 
            Caption         =   "&Connection to IRC"
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnuNewStatusWindow 
            Caption         =   "&Status Window"
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuNewChannel 
            Caption         =   "&Channel"
            Shortcut        =   {F9}
         End
         Begin VB.Menu mnuNewQuery 
            Caption         =   "&Query"
         End
         Begin VB.Menu mnuSep38962937 
            Caption         =   "-"
         End
         Begin VB.Menu mnuNewBot 
            Caption         =   "Bot"
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuNewBotCommand2 
            Caption         =   "Bot Command"
            Shortcut        =   ^K
         End
         Begin VB.Menu mnuSep89326362789 
            Caption         =   "-"
         End
         Begin VB.Menu mnuNewScriptFile 
            Caption         =   "Script &File"
         End
         Begin VB.Menu mnuNewScriptFileRange 
            Caption         =   "Script File &Range"
         End
         Begin VB.Menu mnuSep2378936 
            Caption         =   "-"
         End
         Begin VB.Menu mnuNewTheme 
            Caption         =   "&Theme"
         End
      End
      Begin VB.Menu mnuOpenMenu 
         Caption         =   "Open"
         Begin VB.Menu mnuOpenScript 
            Caption         =   "Script"
            Begin VB.Menu mnuOpenScriptFile 
               Caption         =   "VB Script"
               Shortcut        =   ^O
            End
            Begin VB.Menu mnuNexIRCScriptFile 
               Caption         =   "nIRC Script"
            End
            Begin VB.Menu mnuSep8937289638 
               Caption         =   "-"
            End
            Begin VB.Menu mnuOpenMenuFile 
               Caption         =   "Menu"
            End
         End
         Begin VB.Menu mnuWebSlashHtml 
            Caption         =   "Web/HTML"
            Begin VB.Menu mnuOpenWebsiteURL 
               Caption         =   "Enter Website URL"
            End
            Begin VB.Menu mnuOpenHTMLFile 
               Caption         =   "Open File"
            End
         End
         Begin VB.Menu mnuOpenLogFile 
            Caption         =   "Log File"
            Shortcut        =   ^L
         End
         Begin VB.Menu mnUSep3296397263 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOpenSupportedMedia 
            Caption         =   "Audio"
         End
      End
      Begin VB.Menu mnuSaveMenu 
         Caption         =   "Save"
         Begin VB.Menu mnuSave 
            Caption         =   "Save"
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save as ..."
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuSep8396298736 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSavePlaylist 
            Caption         =   "Save Playlist"
            Shortcut        =   ^P
         End
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Import"
         Begin VB.Menu mnuImportmIRCServers 
            Caption         =   "mIRC Servers.ini"
         End
      End
      Begin VB.Menu mnuSep378927398263 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuConnection 
         Caption         =   "Connection"
         Begin VB.Menu mnuQuickConnect 
            Caption         =   "Quick Connect"
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuConnectionManager 
            Caption         =   "Connection Manager"
            Shortcut        =   ^M
         End
         Begin VB.Menu mnuAutoConnect 
            Caption         =   "Auto Connect"
            Shortcut        =   ^U
         End
         Begin VB.Menu mnuNexIRCServer 
            Caption         =   "IRC Server"
         End
         Begin VB.Menu mnUSep3789263792 
            Caption         =   "-"
         End
         Begin VB.Menu mnuConnect 
            Caption         =   "Connect"
         End
         Begin VB.Menu mnuDisconnect 
            Caption         =   "Disconnect"
         End
         Begin VB.Menu mnuSep4380723869264 
            Caption         =   "-"
         End
         Begin VB.Menu mnuQuits 
            Caption         =   "Quits"
            Begin VB.Menu mnuSendQuitMessage 
               Caption         =   "Default Quit"
            End
            Begin VB.Menu mnuSep327893629873 
               Caption         =   "-"
            End
            Begin VB.Menu mnuFakeKline 
               Caption         =   "Fake K-Line"
            End
            Begin VB.Menu mnuRebooting 
               Caption         =   "Rebooting"
            End
            Begin VB.Menu mnuChangingServers 
               Caption         =   "Changing Servers"
            End
            Begin VB.Menu mnuGoingToSleep 
               Caption         =   "Going to sleep"
            End
            Begin VB.Menu mnuBusy 
               Caption         =   "Busy"
            End
         End
         Begin VB.Menu mnuConnectNetworks 
            Caption         =   "Networks"
            Begin VB.Menu mnuConnectToUndernet 
               Caption         =   "Connect to Undernet"
               Shortcut        =   +^{F1}
            End
            Begin VB.Menu mnuConnectToNewnet 
               Caption         =   "Connect to Newnet"
               Shortcut        =   +^{F2}
            End
            Begin VB.Menu mnuConnectToEfnet 
               Caption         =   "Connect to Efnet"
               Shortcut        =   +^{F3}
            End
         End
      End
      Begin VB.Menu mnuChannels 
         Caption         =   "Channels"
         Begin VB.Menu mnuAutojoin3 
            Caption         =   "Autojoin"
            Begin VB.Menu mnuEditAutojoin 
               Caption         =   "Edit"
               Shortcut        =   {F6}
            End
            Begin VB.Menu mnuActivateAutojoin 
               Caption         =   "Activate"
               Shortcut        =   {F7}
            End
         End
         Begin VB.Menu mnuJoinChannels 
            Caption         =   "Join Channel"
            Begin VB.Menu mnuJoinNEXGEN 
               Caption         =   "#nexgen"
            End
            Begin VB.Menu mnuJoinAcidmax 
               Caption         =   "#acidmax"
            End
            Begin VB.Menu mnuJoinNEXGENTRIVIA 
               Caption         =   "#nexgentrivia"
            End
         End
         Begin VB.Menu mnuSep03270923686293 
            Caption         =   "-"
         End
         Begin VB.Menu mnuJoinChannelName 
            Caption         =   "Join"
            Shortcut        =   ^J
         End
         Begin VB.Menu mnuSep892379873692 
            Caption         =   "-"
         End
         Begin VB.Menu mnuListChannels 
            Caption         =   "List"
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu mnuFindChannels 
            Caption         =   "Find"
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnuChannelFolder 
            Caption         =   "Folder"
            Shortcut        =   ^{F3}
         End
      End
      Begin VB.Menu mnuServer 
         Caption         =   "Server"
         Begin VB.Menu mnuMessageServer 
            Caption         =   "Send Message"
         End
         Begin VB.Menu mnuMOTD 
            Caption         =   "MOTD"
         End
         Begin VB.Menu mnuLUSERS 
            Caption         =   "LUSERS"
         End
         Begin VB.Menu mnuTIME 
            Caption         =   "TIME"
         End
         Begin VB.Menu mnuINFO 
            Caption         =   "INFO"
         End
         Begin VB.Menu mnuIRCOPS 
            Caption         =   "IRCOPS"
         End
         Begin VB.Menu mnuKLINES 
            Caption         =   "K-LINES"
         End
      End
      Begin VB.Menu mnuDCC 
         Caption         =   "DCC"
         Begin VB.Menu mnuDCCSend 
            Caption         =   "Send"
            Shortcut        =   ^F
         End
         Begin VB.Menu mnuDCCChat 
            Caption         =   "Chat"
            Shortcut        =   ^W
         End
         Begin VB.Menu mnuSep372986392639 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDownloads 
            Caption         =   "Downloads"
         End
      End
      Begin VB.Menu mnuScripts 
         Caption         =   "Scripts"
         Begin VB.Menu mnuMenuEditor 
            Caption         =   "Menu Editor"
            Shortcut        =   ^D
         End
         Begin VB.Menu mnuExecuteScript 
            Caption         =   "Execute Script"
            Shortcut        =   {F8}
         End
         Begin VB.Menu mnuSep3209369263972 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuScriptEditor 
            Caption         =   "Script Editor"
         End
         Begin VB.Menu mnuScriptBrowser 
            Caption         =   "Script Browser"
            Shortcut        =   {F11}
         End
      End
      Begin VB.Menu mnuAudio 
         Caption         =   "Audio"
         Begin VB.Menu mnuQuickPlay 
            Caption         =   "Quick Play"
            Shortcut        =   ^G
         End
         Begin VB.Menu mnuVideoWindow 
            Caption         =   "Video Window"
         End
         Begin VB.Menu mnuSep3892697362 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMixer 
            Caption         =   "Mixer"
            Shortcut        =   {F12}
         End
         Begin VB.Menu mnuAlarm 
            Caption         =   "Alarm"
         End
         Begin VB.Menu mnuSep839726987362 
            Caption         =   "-"
         End
         Begin VB.Menu mnuControls 
            Caption         =   "Controls"
            Begin VB.Menu mnuPlay 
               Caption         =   "Play"
            End
            Begin VB.Menu mnuPause 
               Caption         =   "Pause"
            End
            Begin VB.Menu mnuStop 
               Caption         =   "Stop"
            End
         End
         Begin VB.Menu mnuPlaylist3 
            Caption         =   "Playlist"
            Begin VB.Menu mnuShowPlaylist 
               Caption         =   "Show"
               Shortcut        =   +{F1}
            End
            Begin VB.Menu mnuAddtoPlaylist 
               Caption         =   "Add Folder"
            End
            Begin VB.Menu mnuAddmedia 
               Caption         =   "Add Media"
            End
            Begin VB.Menu mnuSearchwithinPlaylist 
               Caption         =   "Search"
               Shortcut        =   {F3}
            End
            Begin VB.Menu mnuSep89236362 
               Caption         =   "-"
            End
            Begin VB.Menu mnuRefreshPlaylist 
               Caption         =   "Refresh"
            End
            Begin VB.Menu mnuPlaylistSep 
               Caption         =   "-"
            End
            Begin VB.Menu mnuPlaylistINI 
               Caption         =   "playlist.ini"
            End
            Begin VB.Menu mnuPlaylistCollection 
               Caption         =   ""
               Index           =   0
            End
         End
      End
      Begin VB.Menu mnuBrowser 
         Caption         =   "Browser"
         Begin VB.Menu mnuOpenURL 
            Caption         =   "Open URL"
            Shortcut        =   ^R
         End
         Begin VB.Menu mnuLinks 
            Caption         =   "Links"
            Begin VB.Menu mnuDALnet 
               Caption         =   "DALnet"
            End
            Begin VB.Menu mnuNewnet 
               Caption         =   "Newnet"
            End
            Begin VB.Menu mnuAustnet 
               Caption         =   "Austnet"
            End
         End
         Begin VB.Menu mnuSep980372986392 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBack 
            Caption         =   "Back"
         End
         Begin VB.Menu mnuForward 
            Caption         =   "Forward"
         End
         Begin VB.Menu mnuHome 
            Caption         =   "Home"
         End
         Begin VB.Menu mnuSearch 
            Caption         =   "Search"
         End
         Begin VB.Menu mnuRefresh 
            Caption         =   "Refresh"
         End
         Begin VB.Menu mnuClose 
            Caption         =   "Close"
         End
      End
      Begin VB.Menu mnuBotControl 
         Caption         =   "Bots"
         Begin VB.Menu mnuShowBotControl 
            Caption         =   "Run Command"
            Shortcut        =   ^{F4}
         End
         Begin VB.Menu mnuNewBotCommand 
            Caption         =   "Add Bot/Command"
            Shortcut        =   ^{F5}
         End
         Begin VB.Menu mnuAutoPreform 
            Caption         =   "Auto Perform"
            Shortcut        =   ^{F6}
         End
      End
      Begin VB.Menu mnuSystemStats 
         Caption         =   "System Stats"
         Begin VB.Menu mnuAdvancedSystemStats 
            Caption         =   "Advanced System Stats"
         End
         Begin VB.Menu mnuSep3879264 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSystemStatsConsole 
            Caption         =   "System Stats Console"
         End
         Begin VB.Menu mnuSep3890283692836 
            Caption         =   "-"
         End
         Begin VB.Menu mnuShowAll 
            Caption         =   "Show All"
         End
         Begin VB.Menu mnuShowProcessors 
            Caption         =   "Processors"
         End
         Begin VB.Menu mnuScsiDevices 
            Caption         =   "SCSI Devices"
         End
         Begin VB.Menu mnuOperatingSystem 
            Caption         =   "Operating System"
         End
      End
      Begin VB.Menu mnuSep83729639 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustomize 
         Caption         =   "Customize"
      End
      Begin VB.Menu mnuSetupWizard 
         Caption         =   "Setup Wizard"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      Begin VB.Menu mnuTileHorizontal 
         Caption         =   "Tile Horizontal"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuTileVerticle 
         Caption         =   "Tile Verticle"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "Cascade"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuArrangeIcons 
         Caption         =   "Arrange Icons"
      End
      Begin VB.Menu mnuSep836293678926 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoMax 
         Caption         =   "Auto Maximize"
      End
      Begin VB.Menu mnuSep389623627 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMinimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuMinimizeAll 
         Caption         =   "Minimize All"
      End
      Begin VB.Menu mnuSep8392697365279358 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseWindow 
         Caption         =   "Close Window"
      End
      Begin VB.Menu mnuSep38927632673 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWList 
         Caption         =   "List"
         WindowList      =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp1 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuShowTips 
         Caption         =   "Tip"
      End
      Begin VB.Menu mnuSep380926973862734582 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRegister 
         Caption         =   "Register"
      End
      Begin VB.Menu mnuNexgen 
         Caption         =   "Team Nexgen"
         Begin VB.Menu mnuSoftware 
            Caption         =   "Software"
         End
         Begin VB.Menu mnuSep8936296392 
            Caption         =   "-"
         End
         Begin VB.Menu mnuNexgenHomepage 
            Caption         =   "Homepage"
         End
         Begin VB.Menu mnuScripts2 
            Caption         =   "mIRC Scripts"
         End
         Begin VB.Menu mnuMessageForums 
            Caption         =   "Message Forums"
         End
         Begin VB.Menu mnuStaff 
            Caption         =   "Members"
         End
         Begin VB.Menu mnuGuestBook 
            Caption         =   "Guest Book"
         End
         Begin VB.Menu mnuArtwork 
            Caption         =   "Artwork"
         End
      End
      Begin VB.Menu mnuSep903872983629 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "mdiNexIRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private eUsername As String
Private eForm As Form
Private lPlaylistSendIndex As Integer
Private bonk As Integer
Private PicDir As Integer
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hDC As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, lParam As Any) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const CFM_BACKCOLOR = &H4000000
Private colString As New Collection
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private lSplashVisible As Boolean
Private lSplashDelay As Integer

Public Sub SetSplashVisible(lValue As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSplashVisible = lValue
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetSplashVisible(lValue As Boolean)"
End Sub

Public Function IsSplashVisible() As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
IsSplashVisible = lSplashVisible
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function IsSplashVisible() As Boolean"
End Function

Public Sub SetEUsername(lUserName As String, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lPlaylistSendIndex = 0
eUsername = lUserName
Set eForm = lForm
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetEUsername(lUsername As String)"
End Sub

Public Sub UpdateMainButtonTypes()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
SetButtonType cmdDelete
SetButtonType frmMobileMixer.cmdAdd
SetButtonType frmMobileMixer.cmdAudioSettings
SetButtonType frmMobileMixer.cmdPlaylist
SetButtonType frmMobileMixer.cmdRandom
SetButtonType cmdSave
SetButtonType cmdSend
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub UpdateMainButtonTypes()"
End Sub

Public Sub ActivateDCCSendByNickname(lNickName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim NewSendFileWin As frmSendFile
If Len(lNickName) <> 0 Then
    Set NewSendFileWin = New frmSendFile
    FileListenPort = FileListenPort + 1
    If FileListenPort > 9000 Then FileListenPort = 1560
    Load NewSendFileWin.tcpSend(FileListenPort)
    NewSendFileWin.Tag = Str(FileListenPort)
    NewSendFileWin.tcpSend(NewSendFileWin.Tag).LocalPort = NewSendFileWin.Tag
    NewSendFileWin.tcpSend(NewSendFileWin.Tag).Listen
    NewSendFileWin.Show 0, Me
    NewSendFileWin.txtNickname.Text = lNickName
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ActivateDCCSendByNickname(lNickname As String)"
End Sub

Public Sub ActiveateDCCSend(lFileName As String, lNickName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim NewSendFileWin As frmSendFile
If Len(lFileName) <> 0 And Len(lNickName) <> 0 Then
    Set NewSendFileWin = New frmSendFile
    FileListenPort = FileListenPort + 1
    If FileListenPort > 9000 Then FileListenPort = 1560
    Load NewSendFileWin.tcpSend(FileListenPort)
    NewSendFileWin.Tag = Str(FileListenPort)
    NewSendFileWin.tcpSend(NewSendFileWin.Tag).LocalPort = NewSendFileWin.Tag
    NewSendFileWin.tcpSend(NewSendFileWin.Tag).Listen
    NewSendFileWin.Show 0, Me
    NewSendFileWin.txtFileName.Text = lFileName
    NewSendFileWin.txtNickname.Text = lNickName
    NewSendFileWin.SetStrFullPath lFileName
    NewSendFileWin.tmrSendFile.Enabled = True
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ActiveateDCCSend()"
End Sub

Public Sub ActivateResize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
MDIForm_Resize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ActivateResize()"
End Sub

Private Sub cboProporties_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, hDC As Long, msg As String, msg2 As String, m As Integer
cboValue.Clear
Select Case cboProporties.Text
Case "BottomToolbarColor"
    cboValue.AddItem "Selected (" & lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBottomToolbarColor & ")"
    For i = 0 To cboColors.ListCount
        If Len(cboColors.List(i)) <> 0 Then cboValue.AddItem cboColors.List(i)
    Next i
    cboValue.AddItem "(Custom)"
    cboValue.ListIndex = 0
Case "SpectrumBackcolor"
    cboValue.AddItem "Selected (" & lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sSpectrumBackcolor & ")"
    For i = 0 To cboColors.ListCount
        If Len(cboColors.List(i)) <> 0 Then cboValue.AddItem cboColors.List(i)
    Next i
    cboValue.AddItem "(Custom)"
    cboValue.ListIndex = 0
Case "BackColor"
    cboValue.AddItem "Selected (" & lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor & ")"
    For i = 0 To cboColors.ListCount
        If Len(cboColors.List(i)) <> 0 Then cboValue.AddItem cboColors.List(i)
    Next i
    cboValue.AddItem "(Custom)"
    cboValue.ListIndex = 0
Case "BottomBandsColor"
    cboValue.AddItem "Selected (" & lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBottomBandsColor & ")"
    For i = 0 To cboColors.ListCount
        If Len(cboColors.List(i)) <> 0 Then cboValue.AddItem cboColors.List(i)
    Next i
    cboValue.AddItem "(Custom)"
    cboValue.ListIndex = 0
Case "ButtonTypes"
    cboValue.AddItem "Windows 16 Bit"
    cboValue.AddItem "Windows 32 Bit"
    cboValue.AddItem "Windows XP"
    cboValue.AddItem "Macintosh"
    cboValue.AddItem "Java"
    cboValue.AddItem "Netscape"
    cboValue.AddItem "Flat"
    cboValue.AddItem "Highlight"
    cboValue.AddItem "Office XP"
    cboValue.AddItem "Transparent"
    cboValue.AddItem "3D Hover"
    cboValue.AddItem "Oval Flat"
    cboValue.AddItem "KDE/2"
    cboValue.ListIndex = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sButtonType
Case "DividerColor"
    cboValue.AddItem "Selected (" & lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sDividerColor & ")"
    For i = 0 To cboColors.ListCount
        If Len(cboColors.List(i)) <> 0 Then cboValue.AddItem cboColors.List(i)
    Next i
    cboValue.AddItem "(Custom)"
    cboValue.ListIndex = 0
Case "Fontname"
    hDC = GetDC(cboValue.hWnd)
    ShowFontType = 4
    EnumFontFamilies hDC, vbNullString, AddressOf EnumFontFamTypeProc, cboValue
    cboValue.ListIndex = FindComboBoxIndex(cboValue, lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sFontname)
Case "Fontsize"
    For i = 6 To 120
        cboValue.AddItem Str(i)
    Next i
    cboValue.ListIndex = FindComboBoxIndex(cboValue, Str(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sFontsize))
Case "IRCColors"
    For i = 1 To lSpectrumThemes.sCount
        If Len(lSpectrumThemes.sSpectrumTheme(i).sIRCColors) <> 0 Then
            cboValue.AddItem Trim(lSpectrumThemes.sSpectrumTheme(i).sIRCColors)
            If i = lSpectrumThemes.sIndex Then
                m = FindComboBoxIndex(cboValue, Trim(lSpectrumThemes.sSpectrumTheme(i).sIRCColors))
            End If
        End If
    Next i
    cboValue.AddItem "(Custom)"
    cboValue.ListIndex = m
Case "LeftChanColor"
    cboValue.AddItem "Selected (" & lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sLeftChanColor & ")"
    For i = 0 To cboColors.ListCount
        If Len(cboColors.List(i)) <> 0 Then cboValue.AddItem cboColors.List(i)
    Next i
    cboValue.AddItem "(Custom)"
    cboValue.ListIndex = 0
Case "Name"
    cboValue.AddItem lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sName
    cboValue.AddItem "(Change)"
    cboValue.ListIndex = 0
Case "PeaksColor"
    cboValue.AddItem "Selected (" & lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sPeaksColor & ")"
    For i = 0 To cboColors.ListCount
        If Len(cboColors.List(i)) <> 0 Then cboValue.AddItem cboColors.List(i)
    Next i
    cboValue.AddItem "(Custom)"
    cboValue.ListIndex = 0
Case "RightChanColor"
    cboValue.AddItem "Selected (" & lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sRightChanColor & ")"
    For i = 0 To cboColors.ListCount
        If Len(cboColors.List(i)) <> 0 Then cboValue.AddItem cboColors.List(i)
    Next i
    cboValue.AddItem "(Custom)"
    cboValue.ListIndex = 0
Case "ShowPeaks"
    cboValue.AddItem "True"
    cboValue.AddItem "False"
    'If ctlMP3OCX.ShowPeaks = True Then
     '   cboValue.ListIndex = 0
    'E 'lse
     '   cboValue.ListIndex = 1
    'End If
Case "SpectrumBands"
    For i = 8 To 48
        cboValue.AddItem Str(i)
    Next i
    cboValue.ListIndex = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBands - 8
Case "SpectrumMode"
    For i = 0 To 3
        cboValue.AddItem Str(i)
    Next i
    cboValue.ListIndex = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sMode
Case "SpectrumVisual"
    cboValue.AddItem "None"
    cboValue.AddItem "Wave"
    cboValue.AddItem "Spectrum"
    'Select Case ctlMP3OCX.OscilloType
    'Case otNone
    '    cboValue.ListIndex = 0
    '    lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sOscilloType = 0
    'Case otSpectrum
    '    cboValue.ListIndex = 2
    '    lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sOscilloType = 1
    'Case otWave
    '    cboValue.ListIndex = 1
    '    lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sOscilloType = 2
    'End Select
Case "TextColor"
    cboValue.AddItem "Selected (" & lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sTextColor & ")"
    For i = 0 To cboColors.ListCount
        If Len(cboColors.List(i)) <> 0 Then cboValue.AddItem cboColors.List(i)
    Next i
    cboValue.AddItem "(Custom)"
    cboValue.ListIndex = 0
Case "ToolbarGraphic"
    For i = 1 To lSpectrumThemes.sCount
        If Len(lSpectrumThemes.sSpectrumTheme(i).sToolbarGraphic) <> 0 Then
            cboValue.AddItem lSpectrumThemes.sSpectrumTheme(i).sToolbarGraphic
            If i = lSpectrumThemes.sIndex Then
                m = FindComboBoxIndex(cboValue, lSpectrumThemes.sSpectrumTheme(i).sToolbarGraphic)
            End If
        End If
    Next i
    cboValue.AddItem "(Custom)"
    cboValue.ListIndex = m
Case "TopBandsColor"
    cboValue.AddItem "Selected (" & lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sTopBandsColor & ")"
    For i = 0 To cboColors.ListCount
        If Len(cboColors.List(i)) <> 0 Then cboValue.AddItem cboColors.List(i)
    Next i
    cboValue.AddItem "(Custom)"
    cboValue.ListIndex = 0
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cboProporties_Change()"
End Sub

Private Sub cboSpectrumThemes_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ApplySpectrumTheme cboSpectrumThemes.Text: DoEvents
cboValue.Clear
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cboSpectrumThemes_Click()"
End Sub

Private Sub cboValue_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer, l As Long
Select Case cboProporties.Text
Case "BottomToolbarColor"
    msg = Trim(Parse(cboValue.Text, "(", ")"))
    If Len(msg) <> 0 Then
        If LCase(msg) = "custom" Then
            frmColorSelector.Show 1
            msg = lReturnColor
            If Len(msg) = 0 Then
                Exit Sub
            Else
                cboValue.AddItem "User (" & msg & ")", 0
                cboValue.ListIndex = 0
            End If
        End If
        lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBottomToolbarColor = Trim(msg)
    End If
Case "SpectrumBackcolor"
    msg = Trim(Parse(cboValue.Text, "(", ")"))
    If Len(msg) <> 0 Then
        If LCase(msg) = "custom" Then
            frmColorSelector.Show 1
            msg = lReturnColor
            If Len(msg) = 0 Then
                Exit Sub
            Else
                cboValue.AddItem "User (" & msg & ")", 0
                cboValue.ListIndex = 0
            End If
        End If
        lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sSpectrumBackcolor = Trim(msg)
    End If
Case "BackColor"
    msg = Trim(Parse(cboValue.Text, "(", ")"))
    If Len(msg) <> 0 Then
        If LCase(msg) = "custom" Then
            frmColorSelector.Show 1
            msg = lReturnColor
            If Len(msg) = 0 Then
                Exit Sub
            Else
                cboValue.AddItem "User (" & msg & ")", 0
                cboValue.ListIndex = 0
            End If
        End If
        lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor = Trim(msg)
    End If
Case "BottomBandsColor"
    msg = Trim(Parse(cboValue.Text, "(", ")"))
    If Len(msg) <> 0 Then
        If LCase(msg) = "custom" Then
            frmColorSelector.Show 1
            msg = lReturnColor
            If Len(msg) = 0 Then
                Exit Sub
            Else
                cboValue.AddItem "User (" & msg & ")", 0
                cboValue.ListIndex = 0
            End If
        End If
        lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBottomBandsColor = Trim(msg)
    End If
Case "ButtonTypes"
    i = cboValue.ListIndex
    lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sButtonType = i
    UpdateMainButtonTypes
Case "DividerColor"
    msg = Trim(Parse(cboValue.Text, "(", ")"))
    If Len(msg) <> 0 Then
        If LCase(msg) = "custom" Then
            frmColorSelector.Show 1
            msg = lReturnColor
            If Len(msg) = 0 Then
                Exit Sub
            Else
                cboValue.AddItem "User (" & msg & ")", 0
                cboValue.ListIndex = 0
            End If
        End If
        lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sDividerColor = Trim(msg)
    End If
Case "FontName"
    lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sFontname = cboValue.Text
Case "FontSize"
    lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sFontsize = Int(cboValue.Text)
Case "IRCColors"
    lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sIRCColors = Trim(cboValue.Text)
Case "LeftChanColor"
    msg = Trim(Parse(cboValue.Text, "(", ")"))
    If Len(msg) <> 0 Then
        If LCase(msg) = "custom" Then
            frmColorSelector.Show 1
            msg = lReturnColor
            If Len(msg) = 0 Then
                Exit Sub
            Else
                cboValue.AddItem "User (" & msg & ")", 0
                cboValue.ListIndex = 0
            End If
        End If
        lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sLeftChanColor = Trim(msg)
    End If
Case "Name"
    If cboValue.Text = "(Change)" Then
        msg = InputBox("Enter name of Spectrum Theme:", "Spectrum Themes")
        If Len(msg) <> 0 Then
            lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sName = msg
            cboValue.Clear
            cboValue.AddItem msg
            cboValue.AddItem "(Change)"
        End If
    End If
Case "PeaksColor"
    msg = Trim(Parse(cboValue.Text, "(", ")"))
    If Len(msg) <> 0 Then
        If LCase(msg) = "custom" Then
            frmColorSelector.Show 1
            msg = lReturnColor
            If Len(msg) = 0 Then
                Exit Sub
            Else
                cboValue.AddItem "User (" & msg & ")", 0
                cboValue.ListIndex = 0
            End If
        End If
        lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sPeaksColor = Trim(msg)
    End If
Case "RightChanColor"
    msg = Trim(Parse(cboValue.Text, "(", ")"))
    If Len(msg) <> 0 Then
        If LCase(msg) = "custom" Then
            frmColorSelector.Show 1
            msg = lReturnColor
            If Len(msg) = 0 Then
                Exit Sub
            Else
                cboValue.AddItem "User (" & msg & ")", 0
                cboValue.ListIndex = 0
            End If
        End If
        lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sRightChanColor = Trim(msg)
    End If
Case "ShowPeaks"
    Select Case cboValue.Text
    Case "True"
'        ctlMP3OCX.ShowPeaks = True
    Case "False"
 '       ctlMP3OCX.ShowPeaks = False
    End Select
Case "SpectrumBands"
'    ctlMP3OCX.Bands = Int(cboValue.Text)
    lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBands = Int(cboValue.Text)
Case "SpectrumMode"
'    ctlMP3OCX.SpectrumMode = Int(cboValue.Text)
    lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sMode = Int(cboValue.Text)
Case "SpectrumVisual"
    Select Case cboValue.Text
    Case "None"
'        ctlMP3OCX.OscilloType = otNone
        lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sOscilloType = 0
    Case "Spectrum"
'        ctlMP3OCX.OscilloType = otSpectrum
        lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sOscilloType = 2
    Case "Wave"
'        ctlMP3OCX.OscilloType = otWave
        lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sOscilloType = 1
    End Select
Case "TextColor"
    msg = Trim(Parse(cboValue.Text, "(", ")"))
    If Len(msg) <> 0 Then
        If LCase(msg) = "custom" Then
            frmColorSelector.Show 1
            msg = lReturnColor
            If Len(msg) = 0 Then
                Exit Sub
            Else
                cboValue.AddItem "User (" & msg & ")", 0
                cboValue.ListIndex = 0
            End If
        End If
        lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sTextColor = Trim(msg)
    End If
Case "ToolbarGraphic"
    lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sToolbarGraphic = cboValue.Text
Case "TopBandsColor"
    msg = Trim(Parse(cboValue.Text, "(", ")"))
    If Len(msg) <> 0 Then
        If LCase(msg) = "custom" Then
            frmColorSelector.Show 1
            msg = lReturnColor
            If Len(msg) = 0 Then
                Exit Sub
            Else
                cboValue.AddItem "User (" & msg & ")", 0
                cboValue.ListIndex = 0
            End If
        End If
        lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sTopBandsColor = Trim(msg)
    End If
End Select
ApplySpectrumTheme lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sName
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cboValue_Change()"
End Sub

Private Sub CHATx_Connect(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ProcessReplaceString sNowConnected, mdiNexIRC.ActiveForm.txtIncoming
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub CHATx_Connect(Index As Integer)"
End Sub

Private Sub CHATx_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ProcessReplaceString sProgrammingError, lChatWindowx(Index).txtIncoming, Description, Str(Number)
wskChat2(Index).Close
lChatWindowx(Index).txtOutgoing.Enabled = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub CHATx_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)"
End Sub

Private Sub cmdSave_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
SaveSpectrumTheme lSpectrumThemes.sIndex
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdSave_Click()"
End Sub

Private Sub cmdSend_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.SendData "PRIVMSG " & cboNotify.Text & " :" & txtMessage.Text & vbCrLf
txtMessage.Text = ""
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdSend_Click()"
End Sub

Private Sub Command1_Click()
'frmPlayer.Show
End Sub

'Private Sub ctlMP3OCX_FrameNotify(ByVal Frame As Long)
'If Frame = frmMobileMixer.sldMixer(2).Max Then
'    frmMobileMixer.sldMixer(2).Value = 0
'    mdiNexIRC.ctlMP3OCX.Visible = False
'Else
'    frmMobileMixer.sldMixer(2).Value = Frame
'    If ProgressScrolling = False Then frmMobileMixer.sldMixer(2).Value = Frame
'    lblFrames.Caption = "Frame: " & Frame & " of " & frmMobileMixer.sldMixer(2).Max
'End If
'If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub ctlMP3OCX_FrameNotify(ByVal Frame As Long)"
'End Sub

'Private Sub ctlMP3OCX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 1 Then
'    FormDrag Me
'Else
'    PopupMenu frmMenus.mnuBackground
'End If
'If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub ctlMP3OCX_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
'End Sub

'Private Sub ctlMP3OCX_PeakFound()
'If lSettings.sLogoTwitchOnPeaks = True Then
'    If PicDir = 1 Then
'        mdiNexIRC.picNexIRC.Left = mdiNexIRC.picNexIRC.Left - 10
'        mdiNexIRC.picNexIRC.Top = mdiNexIRC.picNexIRC.Top - 10
'        Sleep 1
'        mdiNexIRC.picNexIRC.Left = mdiNexIRC.picNexIRC.Left + 10
'        mdiNexIRC.picNexIRC.Top = mdiNexIRC.picNexIRC.Top + 10
'    ElseIf PicDir = 2 Then
'        mdiNexIRC.picNexIRC.Left = mdiNexIRC.picNexIRC.Left + 10
'        mdiNexIRC.picNexIRC.Top = mdiNexIRC.picNexIRC.Top + 10
'        Sleep 1
'        mdiNexIRC.picNexIRC.Left = mdiNexIRC.picNexIRC.Left - 10
'        mdiNexIRC.picNexIRC.Top = mdiNexIRC.picNexIRC.Top - 10
'    ElseIf PicDir = 3 Then
'        mdiNexIRC.picNexIRC.Left = mdiNexIRC.picNexIRC.Left - 10
'        mdiNexIRC.picNexIRC.Top = mdiNexIRC.picNexIRC.Top + 10
'        Sleep 1
'        mdiNexIRC.picNexIRC.Left = mdiNexIRC.picNexIRC.Left + 10
'        mdiNexIRC.picNexIRC.Top = mdiNexIRC.picNexIRC.Top - 10
'    ElseIf PicDir = 4 Then
'        mdiNexIRC.picNexIRC.Left = mdiNexIRC.picNexIRC.Left + 10
'        mdiNexIRC.picNexIRC.Top = mdiNexIRC.picNexIRC.Top - 10
'        Sleep 1
'        mdiNexIRC.picNexIRC.Left = mdiNexIRC.picNexIRC.Left - 10
'        mdiNexIRC.picNexIRC.Top = mdiNexIRC.picNexIRC.Top + 10
'    End If
'End If
'If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub ctlMP3OCX_PeakFound()"
'End Sub

'Private Sub ctlMP3OCX_Started(ByVal Frames As Long)
'Dim msg As String
'If Frames <> 0 Then
'    frmMobileMixer.sldMixer(2).Max = Frames
'End If
'lPlayback.pPlaying = True
'If ctlMP3OCX.NumberOfChannels = 1 Then
'    msg = "M"
'ElseIf ctlMP3OCX.NumberOfChannels = 2 Then
'    msg = "ST"
'End If
'lblKHZ.Caption = ctlMP3OCX.BitRate & "/" & msg & "/" & ctlMP3OCX.SamplesPerSecond
'If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub ctlMP3OCX_Started(ByVal Frames As Long)"
'End Sub

'Private Sub ctlMP3OCX_ThreadEnded(ByVal ExitCode As MP3OCXLib.ThreadErrors)
'lPlayback.pPlaying = False
'If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub ctlMP3OCX_ThreadEnded(ByVal ExitCode As MP3OCXLib.ThreadErrors)"
'End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If IsSplashVisible = True Then Unload frmSplash
Me.Visible = False
lSettings.sMainVisisble = False
Unload frmGraphics
End
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub MDIForm_Unload(Cancel As Integer)"
End Sub

Private Sub mnuAdvancedSystemStats_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmStats.Show
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAdvancedSystemStats_Click()"
End Sub

Private Sub mnuAutoConnect_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmAutoConnect.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAutoConnect_Click()"
End Sub

Private Sub mnuAutoMax_Click()
If mnuAutoMax.Checked = False Then
    mnuAutoMax.Checked = True
Else
    mnuAutoMax.Checked = False
End If
End Sub

Private Sub mnuConnect_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
i = FindStatusWindowIndexByhWND(mdiNexIRC.ActiveForm.hWnd)
If i <> 0 And Len(ReturnStatusWindowServer(i)) <> 0 Then ConnectToIRC ReturnStatusWindowServer(i), ReturnStatusWindowPort(i), ReturnStatusWindow(i)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuConnect_Click()"
End Sub

Private Sub mnuDALnet_Click()
Surf "http://www.dal.net", Me.hWnd
End Sub

Private Sub mnuImportmIRCServers_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg9 As String, msg8 As String, msg As String, msg7 As String, msg6 As String, s1() As String, s2() As String, i As Integer, n As Integer, c As Integer, msg2 As String, msg3 As String, msg4 As String, msg5 As String, d As Integer
msg6 = OpenDialog(Me, "Servers.ini (*.ini)|*.ini|", "Import mIRC 'servers.ini'", App.Path & "\")
If Len(msg6) = 0 Then Exit Sub
msg7 = SaveDialog(Me, "INI Files (*.ini)|*.ini|", "Save file as ...", App.Path & "\data\fixed\")
msg7 = Left(msg7, Len(msg7) - 1) & ".ini"
If Len(msg7) = 0 Then Exit Sub
'Exit Sub
If DoesFileExist(msg7) = True Then Kill msg7
msg = ReadFile(msg6)
s1 = Split(msg, vbCrLf)
frmImportAndExportProgress.Show
frmImportAndExportProgress.XP_ProgressBar1.Max = UBound(s1)
For i = 0 To UBound(s1)
    If Len(s1(i)) <> 0 Then
        s1(i) = Replace(s1(i), "GROUP:", ":")
        s1(i) = Replace(s1(i), "SERVER:", ":")
        s2 = Split(s1(i), ":")
        For n = 0 To UBound(s2)
            Select Case n
            Case 0
                msg2 = Trim(s2(n))
                For c = 0 To Len(msg2)
                    If Left(msg2, 1) <> "=" Then
                        msg2 = Right(msg2, Len(msg2) - 1)
                    Else
                        msg2 = Right(msg2, Len(msg2) - 1)
                        Exit For
                    End If
                Next c
                For d = 0 To 500
                    msg9 = ReadINI(msg7, "AllGroups", Trim(Str(d)), "")
                    If msg2 = msg9 Then Exit For
                    If Len(msg9) = 0 Then
                        If ReadINI(msg7, "AllGroups", Trim(Str(d - 1)), "") <> msg2 Then
                            WriteINI msg7, "AllGroups", Trim(Str(d)), msg2
                            Exit For
                        End If
                    End If
                    'Exit Sub
                Next d
            Case 1
                msg3 = Trim(s2(n))
            Case 2
                msg4 = Trim(s2(n))
            Case 3
                msg5 = Trim(s2(n))
            End Select
        Next n
        frmImportAndExportProgress.lblProgress.Caption = "Importing: " & msg4 & " (" & msg2 & ")"
        For d = 0 To 100
            msg8 = msg4 & "|" & msg3 & "|" & msg5
            If Len(ReadINI(msg7, msg2, Trim(Str(d)), "")) = 0 Then
                WriteINI msg7, msg2, Trim(Str(d)), msg8
                Exit For
            End If
        Next d
        DoEvents
        frmImportAndExportProgress.XP_ProgressBar1.Value = i
    End If
Next i
frmImportAndExportProgress.lblProgress.Caption = "Import and Export Complete"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuImportmIRCServers_Click()"
End Sub

Private Sub mnuJoinAcidmax_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.SendData "JOIN #acidmax" & vbCrLf
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuJoinAcidmax_Click()"
End Sub

Private Sub mnuJoinNEXGENTRIVIA_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.SendData "JOIN #nexgentrivia" & vbCrLf
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuJoinNEXGENTRIVIA_Click()"
End Sub

Private Sub mnuMenuEditor_Click()
frmMenuEditor.Show 0, mdiNexIRC
End Sub

Private Sub mnuNewnet_Click()
Surf "http://www.newnet.net/", Me.hWnd
End Sub

Private Sub mnuNewQuery_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, colorX(0 To 15) As Long, ChatWith As String, i As Integer, xlfound As Boolean
colorX(0) = vbWhite
colorX(1) = vbBlack
colorX(2) = RGB(0, 0, 140)
colorX(3) = RGB(0, 140, 0)
colorX(4) = vbRed
colorX(5) = RGB(110, 65, 0)
colorX(6) = RGB(140, 0, 140)
colorX(7) = RGB(248, 146, 0)
colorX(8) = vbYellow
colorX(9) = vbGreen
colorX(10) = RGB(0, 140, 140)
colorX(11) = RGB(0, 255, 255)
colorX(12) = vbBlue
colorX(13) = vbMagenta
colorX(14) = RGB(140, 140, 140)
colorX(15) = RGB(200, 200, 200)
xlfound = False
ChatWith = InputBox("Nickname to Query:", "Enter Query", "")
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
            LoadQueryWindow i, ChatWith, ""
            'Load lQuery(i)
            'lQuery(i).txtOutgoing.BackColor = colorX(Color.BGText)
            'lQuery(i).txtOutgoing.ForeColor = colorX(Color.Normal)
            'lQuery(i).txtIncoming.SetBackColor colorX(Color.BGText)
            'lQuery(i).Caption = ChatWith
            'lQueryName(i) = ChatWith
            'Call AddTaskPanel(ChatWith, 1)
            Exit For
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuNewQuery_Click()"
End Sub

Private Sub mnuOpenHTMLFile_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = OpenDialog(Me, "HTML Files (*.html)|*.html|HTM Files (*.htm)|*.htm|ASP Files (*.asp)|*.asp|VB Script Files (*.scr;*.txt)|*.scr;*.txt|All Files (*.*)|*.*|", "Open HTML", CurDir)
If Len(msg) <> 0 And DoesFileExist(msg) = True Then Surf msg, Me.hWnd
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuOpenHTMLFile_Click()"
End Sub

Private Sub mnuOperatingSystem_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ShowSystemStats ActiveForm, True, False, False, False, False, False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuOperatingSystem_Click()"
End Sub

Private Sub mnuScsiDevices_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ShowSystemStats ActiveForm, False, False, False, False, False, True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuScsiDevices_Click()"
End Sub

Private Sub mnuShowAll_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ShowSystemStats ActiveForm, True, True, True, True, True, True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuShowAll_Click()"
End Sub

Private Sub mnuShowProcessors_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ShowSystemStats ActiveForm, False, False, True, True, True, False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuShowProcessors_Click()"
End Sub

Private Sub mnuSystemStatsConsole_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'frmSystemStatsConsole.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuSystemStatsConsole_Click()"
End Sub

Private Sub mnuVideoWindow_Click()
frmPlayer.Show
End Sub

Private Sub picBackward_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PictureBoxMouseDown Button, picBackward, frmGraphics.picBackward2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picChannelFolder_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picBackward_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If picStop.Picture <> frmGraphics.picStop1.Picture Then picStop.Picture = frmGraphics.picStop1.Picture
PictureBoxMouseMove Button, picBackward, frmGraphics.picBackward1, frmGraphics.picBackward2, X, Y, frmGraphics.picBackward3
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picChannelFolder_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picBackward_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sWebVisible = True Then
    If lSettings.sBackgroundWebpage = True Then
        frmWeb.web.GoBack
    End If
Else
    If lPlayback.pPlaying = True Then
        If (lFiles.fIndex - 1) <> 0 Then
            MenuStop
            PlayFile lFiles.fFile(lFiles.fIndex - 1).fFilename
        End If
    End If
End If
PictureBoxMouseUp Button, picBackward, frmGraphics.picBackward1, frmGraphics.picBackward2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picBackward_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picChannelFolder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PictureBoxMouseDown Button, picChannelFolder, frmGraphics.picChannelFolder2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picChannelFolder_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picChannelFolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If picOptions.Picture <> frmGraphics.picOptions1.Picture Then
    picOptions.Picture = frmGraphics.picOptions1.Picture
    SetPictureColor picOptions, lRedColor, lBlueColor, lGreenColor, True
End If
If picSend.Picture <> frmGraphics.picSend1.Picture Then
    picSend.Picture = frmGraphics.picSend1.Picture
    SetPictureColor picSend, lRedColor, lBlueColor, lGreenColor, True
End If
PictureBoxMouseMove Button, picChannelFolder, frmGraphics.picChannelFolder1, frmGraphics.picChannelFolder2, X, Y, frmGraphics.picChannelFolder3
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picChannelFolder_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picChannelFolder_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    frmChannelFolder.Show
End If
PictureBoxMouseDown Button, picChannelFolder, frmGraphics.picChannelFolder2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picChannelFolder_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picChat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PictureBoxMouseDown Button, picChat, frmGraphics.picChat2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picChat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picChat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If picSend.Picture <> frmGraphics.picSend1.Picture Then picSend.Picture = frmGraphics.picSend1.Picture
If picScript.Picture <> frmGraphics.picScript1.Picture Then picScript.Picture = frmGraphics.picScript1.Picture
'If picScript.Picture <> frmGraphics.picScript1.Picture Then picScript.Picture = frmGraphics.picScript1.Picture
PictureBoxMouseMove Button, picChat, frmGraphics.picChat1, frmGraphics.picChat2, X, Y, frmGraphics.picChat3
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picChat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picChat_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmDCC_Chat.Show 0, Me
PictureBoxMouseUp Button, picChat, frmGraphics.picChat1, frmGraphics.picChat2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picChat_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PictureBoxMouseMove Button, picExit, frmGraphics.picExit1, frmGraphics.picExit2, X, Y, frmGraphics.picExit3
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mbox As VbMsgBoxResult
If Button = 1 Then
    If lSettings.sGeneralPrompts = True Then
        mbox = MsgBox("End NexIRC, are you sure?", vbYesNo + vbQuestion, "Exit?")
        If mbox = vbYes Then
            Unload Me
        Else
            Exit Sub
        End If
    Else
        Unload Me
        Exit Sub
    End If
ElseIf Button = 2 Then
    PopupMenu frmMenus.mnuBackground
End If
PictureBoxMouseUp Button, picExit, frmGraphics.picExit1, frmGraphics.picExit2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picForward_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PictureBoxMouseDown Button, picForward, frmGraphics.picForward2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picForward_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picForward_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If picPlay.Picture <> frmGraphics.picPlay1.Picture Then picPlay.Picture = frmGraphics.picPlay1.Picture
PictureBoxMouseMove Button, picForward, frmGraphics.picForward1, frmGraphics.picForward2, X, Y, frmGraphics.picForward3
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picForward_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picForward_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sWebVisible = True Then
    If lSettings.sBackgroundWebpage = True Then
        frmWeb.web.GoForward
    End If
Else
    If lPlayback.pPlaying = True Then
        If (lFiles.fIndex + 1) <> lFiles.fCount Then
            MenuStop
            PlayFile lFiles.fFile(lFiles.fIndex + 1).fFilename
        End If
    End If
End If
PictureBoxMouseUp Button, picForward, frmGraphics.picForward1, frmGraphics.picForward2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picForward_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picMobileMixer_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If bDocked = True Then
    frmMobileMixer.Move -4 * Screen.TwipsPerPixelX, -6 * Screen.TwipsPerPixelY, mdiNexIRC.picMobileMixer.ScaleWidth + (8 * Screen.TwipsPerPixelX), mdiNexIRC.picMobileMixer.ScaleHeight + (8 * Screen.TwipsPerPixelY)
End If
ActivateResize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picMobileMixer_Resize()"
End Sub

Private Sub picNexIRC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PictureBoxMouseDown Button, picNexIRC, frmGraphics.picNexIRC2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picNexIRC_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picNexIRC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PictureBoxMouseMove Button, picNexIRC, frmGraphics.picNexIRC1, frmGraphics.picNexIRC2, X, Y, frmGraphics.picNexIRC3
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picNexIRC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picNexIRC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case Button
Case 1
    'frmCustomize.Show
    frmCustomize.ShowAbout
Case 2
    PopupMenu frmMenus.mnuBackground
End Select
PictureBoxMouseUp Button, picNexIRC, frmGraphics.picNexIRC1, frmGraphics.picNexIRC2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picNexIRC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picPause_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PictureBoxMouseDown Button, picPause, frmGraphics.picPause2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picPause_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picPause_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If picPlay.Picture <> frmGraphics.picPlay1.Picture Then picPlay.Picture = frmGraphics.picPlay1.Picture
If picStop.Picture <> frmGraphics.picStop1.Picture Then picStop.Picture = frmGraphics.picStop1.Picture
PictureBoxMouseMove Button, picPause, frmGraphics.picPause1, frmGraphics.picPause2, X, Y, frmGraphics.picPause3
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picPause_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picPause_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
MenuPause
PictureBoxMouseUp Button, picPause, frmGraphics.picPause1, frmGraphics.picPause2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picPause_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PictureBoxMouseDown Button, picPlay, frmGraphics.picPlay2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If picForward.Picture <> frmGraphics.picForward1.Picture Then picForward.Picture = frmGraphics.picForward1.Picture
If picPause.Picture <> frmGraphics.picPause1.Picture Then picPause.Picture = frmGraphics.picPause1.Picture
PictureBoxMouseMove Button, picPlay, frmGraphics.picPlay1, frmGraphics.picPlay2, X, Y, frmGraphics.picPlay3
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
MenuPlay
PictureBoxMouseUp Button, picPlay, frmGraphics.picPlay1, frmGraphics.picPlay2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picPlay_Click()"
End Sub

Private Sub picExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PictureBoxMouseDown Button, picExit, frmGraphics.picExit2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picExit_Click()"
End Sub

Private Sub picDisconnect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PictureBoxMouseDown Button, picDisconnect, frmGraphics.picDisconnect2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picDisconnect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picDisconnect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If picAudio.Picture <> frmGraphics.picAudio1.Picture Then
    picAudio.Picture = frmGraphics.picAudio1.Picture
    SetPictureColor picAudio, lRedColor, lBlueColor, lGreenColor, True
End If
PictureBoxMouseMove Button, picDisconnect, frmGraphics.picDisconnect1, frmGraphics.picDisconnect2, X, Y, frmGraphics.picDisconnect3
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picDisconnect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picDisconnect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    picConnect.Visible = True
    If lSettings.sActiveServerForm.tcp.State = sckConnected Then SendQuitMessage lSettings.sActiveServerForm
    picDisconnect.Visible = False
    picDisconnect.Picture = frmGraphics.picDisconnect1.Picture
    Button = 0
    lSettings.sActiveServerForm.tmrReconnect.Enabled = False
End If
PictureBoxMouseUp Button, picDisconnect, frmGraphics.picDisconnect1, frmGraphics.picDisconnect2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picDisconnect_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picConnect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PictureBoxMouseDown Button, picConnect, frmGraphics.picConnect2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picConnect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picConnect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If picAudio.Picture <> frmGraphics.picAudio1.Picture Then picAudio.Picture = frmGraphics.picAudio1.Picture
PictureBoxMouseMove Button, picConnect, frmGraphics.picConnect1, frmGraphics.picConnect2, X, Y, frmGraphics.picConnect3
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picConnect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picConnect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim xServer As String, xPort As String, i As Integer, PortFound As Boolean
If Button = 1 Then
    ConnectToIRC lSettings.sServer, lSettings.sPort, lSettings.sActiveServerForm
End If
PictureBoxMouseUp Button, picConnect, frmGraphics.picConnect1, frmGraphics.picConnect2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picConnect_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picAudio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PictureBoxMouseDown Button, picAudio, frmGraphics.picAudio2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picAudio_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picAudio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If picConnect.Picture <> frmGraphics.picConnect1.Picture Then
    picConnect.Picture = frmGraphics.picConnect1.Picture
    SetPictureColor picConnect, lRedColor, lBlueColor, lGreenColor, True
End If
If picDisconnect.Picture <> frmGraphics.picDisconnect1.Picture Then
    picDisconnect.Picture = frmGraphics.picDisconnect1.Picture
    SetPictureColor picDisconnect, lRedColor, lBlueColor, lGreenColor, True
End If
If picOptions.Picture <> frmGraphics.picOptions1.Picture Then
    picOptions.Picture = frmGraphics.picOptions1.Picture
    SetPictureColor picOptions, lRedColor, lBlueColor, lGreenColor, True
End If
PictureBoxMouseMove Button, picAudio, frmGraphics.picAudio1, frmGraphics.picAudio2, X, Y, frmGraphics.picAudio3
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picAudio_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picAudio_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
If Button = 1 Then
    PromptPlayback Me
End If
PictureBoxMouseUp Button, picAudio, frmGraphics.picAudio1, frmGraphics.picAudio2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picAudio_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picScript_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PictureBoxMouseDown Button, picScript, frmGraphics.picScript3
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picScript_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picScript_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If picSend.Picture <> frmGraphics.picSend1.Picture Then picSend.Picture = frmGraphics.picSend1.Picture
If picChat.Picture <> frmGraphics.picChat1.Picture Then picChat.Picture = frmGraphics.picChat1.Picture
PictureBoxMouseMove Button, picScript, frmGraphics.picScript1, frmGraphics.picScript2, X, Y, frmGraphics.picScript2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picScript_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picScript_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
NewScriptFile
PictureBoxMouseUp Button, picScript, frmGraphics.picScript1, frmGraphics.picScript2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picScript_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picSend_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PictureBoxMouseDown Button, picSend, frmGraphics.picSend3
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picSend_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picSend_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If picSend.Picture <> frmGraphics.picSend1.Picture Then picSend.Picture = frmGraphics.picSend1.Picture
If picChat.Picture <> frmGraphics.picChat1.Picture Then picChat.Picture = frmGraphics.picChat1.Picture
If picChannelFolder.Picture <> frmGraphics.picChannelFolder1.Picture Then
    picChannelFolder.Picture = frmGraphics.picChannelFolder1.Picture
    SetPictureColor picChannelFolder, lRedColor, lBlueColor, lGreenColor, True
End If
If picOptions.Picture <> frmGraphics.picOptions1.Picture Then
    picOptions.Picture = frmGraphics.picOptions1.Picture
    SetPictureColor picOptions, lRedColor, lBlueColor, lGreenColor, True
End If
If picScript.Picture <> frmGraphics.picScript1.Picture Then
    picScript.Picture = frmGraphics.picScript1.Picture
    SetPictureColor picScript, lRedColor, lBlueColor, lGreenColor, True
End If
PictureBoxMouseMove Button, picSend, frmGraphics.picSend1, frmGraphics.picSend2, X, Y, frmGraphics.picSend2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picSend_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picSend_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim NewSendFileWin As frmSendFile
Set NewSendFileWin = New frmSendFile
FileListenPort = FileListenPort + 1
If FileListenPort > 9000 Then FileListenPort = 1560
Load NewSendFileWin.tcpSend(FileListenPort)
NewSendFileWin.Tag = Str(FileListenPort)
NewSendFileWin.tcpSend(NewSendFileWin.Tag).LocalPort = NewSendFileWin.Tag
NewSendFileWin.tcpSend(NewSendFileWin.Tag).Listen
NewSendFileWin.Show 0, Me
PictureBoxMouseUp Button, picSend, frmGraphics.picSend1, frmGraphics.picSend2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picSend_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picOptions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PictureBoxMouseDown Button, picOptions, frmGraphics.picOptions2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picOptions_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If picAudio.Picture <> frmGraphics.picAudio1.Picture Then
    picAudio.Picture = frmGraphics.picAudio1.Picture
    SetPictureColor picAudio, lRedColor, lBlueColor, lGreenColor, True
End If
If picSend.Picture <> frmGraphics.picSend1.Picture Then
    picSend.Picture = frmGraphics.picSend1.Picture
    SetPictureColor picSend, lRedColor, lBlueColor, lGreenColor, True
End If
If picChannelFolder.Picture <> frmGraphics.picChannelFolder1.Picture Then
    picChannelFolder.Picture = frmGraphics.picChannelFolder1.Picture
    SetPictureColor picChannelFolder, lRedColor, lBlueColor, lGreenColor, True
End If
PictureBoxMouseMove Button, picOptions, frmGraphics.picOptions1, frmGraphics.picOptions2, X, Y, frmGraphics.picOptions3
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picOptions_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub picOptions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    frmCustomize.Show
End If
PictureBoxMouseUp Button, picOptions, frmGraphics.picOptions1, frmGraphics.picOptions2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picOptions_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub CHAT_Close(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload wskChat(Index)
lChatWindowName(Index) = ""
ProcessReplaceString sConnectionTerminated, lChatWindow(Index).txtIncoming
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub CHAT_Close(Index As Integer)"
End Sub

Private Sub CHAT_Connect(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ProcessReplaceString sConnectionEstablished, mdiNexIRC.ActiveForm.txtIncoming
'Call DoColor(lChatWindow(Index).txtIncoming, "4* Connection established")
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub CHAT_Connect(Index As Integer)"
End Sub

Private Sub CHAT_DataArrival(Index As Integer, ByVal bytesTotal As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim strData As String, msg2 As String
On Local Error Resume Next
wskChat(Index).GetData strData
ProcessReplaceString sPm, lChatWindow(Index).txtIncoming, lChatWindowName(Index), strData
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub CHAT_DataArrival(Index As Integer, ByVal bytesTotal As Long)"
End Sub

Private Sub CHATx_Close(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ProcessReplaceString sConnectionClosed, lChatWindow(Index).txtIncoming
'DoColor lChatWindowx(Index).txtIncoming, "" & Color.Notice & "• Connection Closed"
wskChat2(Index).Close
lChatWindowNamex(Index) = ""
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub CHATx_Close(Index As Integer)"
End Sub

Private Sub CHATx_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ProcessReplaceString sEnterDCCChat, lChatWindowx(Index).txtIncoming, lChatWindowNamex(Index), wskChat2(Index).RemoteHostIP, wskChat2(Index).RemotePort
lChatWindowx(Index).txtOutgoing.Enabled = True
wskChat2(Index).Close
wskChat2(Index).Accept requestID
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub CHATx_ConnectionRequest(Index As Integer, ByVal requestID As Long)"
End Sub

Private Sub CHATx_DataArrival(Index As Integer, ByVal bytesTotal As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim strData As String
wskChat2(Index).GetData strData
If Right(strData, 1) = Chr(10) Or Right(strData, 1) = Chr(13) Then
    strData = Left(strData, Len(strData) - 1)
ElseIf Right(strData, 2) = vbCrLf Then
    strData = Left(strData, Len(strData) - 2)
End If
ProcessReplaceString sPm, lChatWindowx(Index).txtIncoming, lChatWindowNamex(Index), strData
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub CHATx_DataArrival(Index As Integer, ByVal bytesTotal As Long)"
End Sub

Private Sub picStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PictureBoxMouseDown Button, picStop, frmGraphics.picStop2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If picPause.Picture <> frmGraphics.picPause1.Picture Then picPause.Picture = frmGraphics.picPause1.Picture
If picBackward.Picture <> frmGraphics.picBackward1.Picture Then picBackward.Picture = frmGraphics.picBackward1.Picture
PictureBoxMouseMove Button, picStop, frmGraphics.picStop1, frmGraphics.picStop2, X, Y, frmGraphics.picStop3
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
MenuStop
PictureBoxMouseUp Button, picStop, frmGraphics.picStop1, frmGraphics.picStop2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub tmrNotify_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If ReturnNotifyList = "" Then
    tmrNotify.Enabled = False
Else
    If lSettings.sActiveServerForm.tcp.State = sckConnected Then
        lSettings.sActiveServerForm.tcp.SendData "ISON " & RTrim(ReturnNotifyList) & vbCrLf
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrNotify_Timer()"
End Sub

Private Sub lblFilename2_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lblFrames.Left = 11150 + lblFilename2.Width + 100
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lblFilename2_Change()"
End Sub

Private Sub lblFilename2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 2 Then PopupMenu frmMenus.mnuBackground
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lblFrames_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblFrames_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 2 Then PopupMenu frmMenus.mnuBackground
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lblFrames_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblKHZ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 2 Then PopupMenu frmMenus.mnuBackground
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lblKHZ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

'Private Sub AddVBScriptObjects()
'ctlVBScript.AddObject "frmChannelFolder", frmChannelFolder, True
'ctlVBScript.AddObject "frmIRCServer", frmIRCServer, True
'ctlVBScript.AddObject "frmAutoConnect", frmAutoConnect, True
'ctlVBScript.AddObject "mdiNexIRC", mdiNexIRC, True
'ctlVBScript.AddObject "frmChat", frmChat, True
'ctlVBScript.AddObject "frmChannels", frmChannels, True
'ctlVBScript.AddObject "frmQuery", frmQuery, True
'ctlVBScript.AddObject "frmTextEditor", frmTextEditor, True
'ctlVBScript.AddObject "frmWeb", frmWeb, True
'ctlVBScript.AddObject "frmSendFile", frmSendFile, True
'ctlVBScript.AddObject "frmJoinChannel", frmJoinChannel, True
'ctlVBScript.AddObject "frmDownloadManager", frmDownloadManager, True
'ctlVBScript.AddObject "frmMenuEditor", frmMenuEditor, True
'ctlVBScript.AddObject "frmScriptBrowser", frmScriptManager, True
'ctlVBScript.AddObject "frmCustomize", frmCustomize, True
'ctlVBScript.AddObject "frmAddBotCommand", frmAddBotCommand, True
'ctlVBScript.AddObject "frmBots", frmBots, True
'ctlVBScript.AddObject "frmAddMedia", frmAddMedia, True
'ctlVBScript.AddObject "frmAddFolderToPlaylist", frmAddFolderToPlaylist, True
'ctlVBScript.AddObject "frmSetupWizard", frmSetupWizard, True
'ctlVBScript.AddObject "frmAlarm", frmAlarm, True
'ctlVBScript.AddObject "frmAddServer", frmAddServer, True
'ctlVBScript.AddObject "SCRIPT", ctlVBScript, True
'End Sub

Private Sub MDIForm_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
PreloadSettings
If lSettings.sShowSplashOnStartup = True Then tmrUnloadSplashDelay.Enabled = True
If lSettings.sBorderlessObjects = True Then
    CutRegion cboSpectrumThemes.hWnd, cboSpectrumThemes, True
    CutRegion cboProporties.hWnd, cboProporties, True
    CutRegion cboValue.hWnd, cboValue, True
    CutRegion cboNotify.hWnd, cboNotify, True
End If
If lSettings.sLastWindowPos.lParentForm.fWidth = 0 Then lSettings.sLastWindowPos.lParentForm.fWidth = 11280
If lSettings.sLastWindowPos.lParentForm.fHeight = 0 Then lSettings.sLastWindowPos.lParentForm.fHeight = 8880
'AddVBScriptObjects
Me.Width = lSettings.sLastWindowPos.lParentForm.fWidth
Me.Height = lSettings.sLastWindowPos.lParentForm.fHeight
Me.Left = lSettings.sLastWindowPos.lParentForm.fLeft
Me.Top = lSettings.sLastWindowPos.lParentForm.fTop
'If lRegInfo.rRegistered = True Then
'    mnuRegister.Visible = False
'    lRegInfo.rRegistered = True
'    Caption = "NexIRC Professional"
'Else
'    mnuRegister.Visible = True
'    lRegInfo.rRegistered = False
'    Caption = "NexIRC"
'End If
LoadIgnore
If ReturnNotifyEnabled = True Then
    mdiNexIRC.tmrNotify.Enabled = True
Else
    mdiNexIRC.tmrNotify.Enabled = False
End If
SetConnected False
For i = 1 To ReturnTCPUBound
    Load wskChat2(i)
    Load wskChat(i)
Next i
lMyCurrentModes = "+"
FileIndex = 3
FileListenPort = 1559
For i = 1 To ReturnChannelUBound
    SetChannelModes i, ""
Next i
DoEvents
DisplayPlaylists
mdiNexIRC.BackColor = QBColor(lSettings.sBGColor)
DoEvents
If lSettings.sShowQuickNotify = True Then picNotify.Visible = True
If lSettings.sShuffle = True Then frmMobileMixer.chkShuffle.Value = 1
'SwitchPlaybackEngine lPlayback.pCurrentEngine
If lSettings.sContinuousPlay = True Then
    ActivateContinuousPlay
    frmMobileMixer.chkContinuous.Value = 1
End If
If Len(lSettings.sBGPicture) <> 0 Then
    If DoesFileExist(lSettings.sBGPicture) = True Then
        mdiNexIRC.Picture = LoadPicture(lSettings.sBGPicture)
    End If
End If
ActivateResize
mdiNexIRC.Arrange vbCascade
For i = 0 To frmMobileMixer.sldMixer.Count
    Select Case i
    Case 7
        If lInitialAudioValues.iInitialWaveEnabled = True Then
            frmMobileMixer.sldMixer(7).Value = lInitialAudioValues.iWave
        End If
    Case 0
        If lInitialAudioValues.iInitialBassEnabled = True Then
            frmMobileMixer.sldMixer(0).Value = lInitialAudioValues.iBass
        End If
    Case 1
        If lInitialAudioValues.iInitialTrebleEnabled = True Then
            frmMobileMixer.sldMixer(1).Value = lInitialAudioValues.iTreble
        End If
    Case 4
        If lInitialAudioValues.iInitialCDAudioEnabled = True Then
            frmMobileMixer.sldMixer(4).Value = lInitialAudioValues.iCDAudio
        End If
    Case 3
        If lInitialAudioValues.iInitialLineInEnabled = True Then
            frmMobileMixer.sldMixer(3).Value = lInitialAudioValues.iLineIN
        End If
    Case 5
        If lInitialAudioValues.iInitialMicEnabled = True Then
            frmMobileMixer.sldMixer(5).Value = lInitialAudioValues.iMic
        End If
    End Select
Next i
UpdateMainButtonTypes
mnuPlaylistCollection(0).Visible = False
If lSettings.sShowQuickmix = True Then
    ToggleMixer True
End If
If lSettings.sShowTips = True Then frmTip.Show 1, Me
If lSettings.sBackgroundWebpage = True Then
    If lSettings.sNavigateOnStartup = True Then
        frmWeb.Show
        frmWeb.ActivateResize
    End If
End If
If lSettings.sShowServerOnStartup = True Then frmIRCServer.Show
If lSettings.sConnectOnStartup = True Then ConnectToIRC lSettings.sServer, lSettings.sPort, lSettings.sActiveServerForm
If lSpectrumThemes.sIndex = 0 Then lSpectrumThemes.sIndex = 1
ApplySpectrumTheme lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sName: DoEvents
NewStatusWindow lSettings.sServer, lSettings.sPort, False
If lSettings.sShowOptionsOnStartup = True Then frmCustomize.Show 0, Me
lSettings.sMainVisisble = True
mdiNexIRC.Arrange vbCascade
PerformAutoConnect
DoEvents
Me.Visible = True
'frmPlayer.Show
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub MDIForm_Load()"
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 2 Then
    PopupMenu frmMenus.mnuBackground
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ResetMainButtons
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sEnding = False Then
    Cancel = 1
    UnloadProgram
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)"
End Sub

Private Sub MDIForm_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Me.ScaleWidth <> 0 And Me.ScaleWidth > 3990 Then txtMessage.Width = Me.ScaleWidth - 4000
If Me.ScaleWidth <> 0 Then cboSpectrumThemes.Width = ((Me.ScaleWidth / 3) + 320) - ((cmdDelete.Width / 2.6) * 2)
If Me.ScaleWidth <> 0 Then cboProporties.Width = ((Me.ScaleWidth / 3) + 320) - ((cmdDelete.Width / 2.6) * 2)
If Me.ScaleWidth <> 0 Then cboValue.Width = ((Me.ScaleWidth / 3) + 320) - ((cmdDelete.Width / 2.6) * 2)
cmdSend.Left = Me.ScaleWidth + 350
cboProporties.Left = cboSpectrumThemes.Width + 40
cboValue.Left = (cboSpectrumThemes.Width * 2) + 80
CutRegion cboSpectrumThemes.hWnd, cboSpectrumThemes, True
CutRegion cboProporties.hWnd, cboProporties, True
CutRegion cboValue.hWnd, cboValue, True
cmdDelete.Left = Me.ScaleWidth - cmdDelete.Width + 900
cmdSave.Left = Me.ScaleWidth - cmdSave.Width
If lSettings.sBackgroundWebpage = True Then
    If lSettings.sWebVisible = True Or frmWeb.Visible = True Then
        frmWeb.Width = Me.ScaleWidth + 40
        frmWeb.Height = Me.ScaleHeight + 40
        If frmWeb.Left <> -20 Then frmWeb.Left = -40
        If frmWeb.Top <> -20 Then frmWeb.Top = -40
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub MDIForm_Resize()"
End Sub

Private Sub mnuAbout_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
frmCustomize.Show 0, Me
frmCustomize.ShowAbout
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAbout_Click()"
End Sub

Private Sub mnuAddmedia_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmAddMedia.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAddmedia_Click()"
End Sub

Private Sub mnuAddToPlaylist_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PromptAddToPlaylist
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAddToPlaylist_Click()"
End Sub

Private Sub mnuAlarm_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmAlarm.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAlarm_Click()"
End Sub

Private Sub mnuArrangeIcons_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
mdiNexIRC.Arrange vbArrangeIcons
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuArrangeIcons_Click()"
End Sub

Private Sub mnuActivateAutojoin_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ActivateAutoJoin True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuActivateAutojoin_Click()"
End Sub

Private Sub mnuArtwork_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Surf "http://www.team-nexgen.org/immagination", Me.hWnd
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuArtwork_Click()"
End Sub

Private Sub mnuBusy_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'lSettings.sActiveServerForm.tcp.SendData "QUIT •: Busy" & vbCrLf
SendQuitMessage lSettings.sActiveServerForm, "Busy"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuBusy_Click()"
End Sub

Private Sub mnuCascade_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
mdiNexIRC.Arrange vbCascade
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuCascade_Click()"
End Sub

Private Sub mnuChangingServers_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'lSettings.sActiveServerForm.tcp.SendData "QUIT •: Changing Servers" & vbCrLf
SendQuitMessage lSettings.sActiveServerForm, "Changing servers"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuChangingServers_Click()"
End Sub

Private Sub mnuChannelFolder_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmChannelFolder.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuChannelFolder_Click()"
End Sub

Private Sub mnuClose_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sBackgroundWebpage = True Then
    Unload frmWeb
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuClose_Click()"
End Sub

Private Sub mnuCloseWindow_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload ActiveForm
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuCloseWindow_Click()"
End Sub

Private Sub mnuConnectionManager_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmConnectionManager.Show
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuConnectionManager_Click()"
End Sub

Private Sub mnuConnectToEfnet_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
NewStatusWindow "irc.efnet.net", "6667", True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuConnectToEfnet_Click()"
End Sub

Private Sub mnuConnectToNewnet_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
NewStatusWindow "irc.eskimo.com", "6667", True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuConnectToNewnet_Click()"
End Sub

Private Sub mnuConnectToUndernet_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
NewStatusWindow "irc.undernet.org", "6667", True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuConnectToUndernet_Click()"
End Sub

Private Sub mnuDCCChat_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmDCC_Chat.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuDCCChat_Click()"
End Sub

Public Sub DCCSend(lFileName As String, lNickName As String, lClickSendButton As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim NewSendFileWin As frmSendFile, lIP As String, TempFileName As String
If Len(lFileName) <> 0 And Len(lNickName) <> 0 Then
    Set NewSendFileWin = New frmSendFile
    FileListenPort = FileListenPort + 1
    If FileListenPort > 9000 Then FileListenPort = 1560
    Load NewSendFileWin.tcpSend(FileListenPort)
    NewSendFileWin.Tag = Str(FileListenPort)
    NewSendFileWin.tcpSend(NewSendFileWin.Tag).LocalPort = NewSendFileWin.Tag
    NewSendFileWin.tcpSend(NewSendFileWin.Tag).Listen
    NewSendFileWin.Show 0, Me
    If Len(lFileName) <> 0 Then NewSendFileWin.txtFileName.Text = lFileName
    If Len(lNickName) <> 0 Then NewSendFileWin.txtNickname.Text = lNickName
    DoEvents
    NewSendFileWin.lblFileSize.Caption = Format(FileLen(NewSendFileWin.txtFileName.Text), "###,###,###,###") & " KB"
    NewSendFileWin.SetStrFullPath lFileName
    If lClickSendButton = True Then NewSendFileWin.ClickSendButton
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub DCCSend(Optional lFilename As String, Optional lNickname As String, Optional lClickSendButton As Boolean)"
End Sub

Private Sub mnuDCCSend_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim NewSendFileWin As frmSendFile
Set NewSendFileWin = New frmSendFile
FileListenPort = FileListenPort + 1
If FileListenPort > 9000 Then FileListenPort = 1560
Load NewSendFileWin.tcpSend(FileListenPort)
NewSendFileWin.Tag = Str(FileListenPort)
NewSendFileWin.tcpSend(NewSendFileWin.Tag).LocalPort = NewSendFileWin.Tag
NewSendFileWin.tcpSend(NewSendFileWin.Tag).Listen
NewSendFileWin.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuDCCSend_Click()"
End Sub

Private Sub mnuDisconnect_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.Close
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuDisconnect_Click()"
End Sub

Private Sub mnuDownloads_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmDownloadManager.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuDownloads_Click()"
End Sub

'Private Sub mnuExecuteScript_Click()
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
'Dim F As Form, msg As String, msg2 As String
'Set F = ActiveForm
'Select Case F.Name
'Case "frmTextEditor"
'    'mdiNexIRC.ctlVBScript.ExecuteStatement F.txtIncoming.Text
'Case Else
'    msg = OpenDialog(Me, "Text Files (*.txt)|*.txt|All Files (*.*)|*.*|", "Open Script", App.Path & "\data\scripts\")
'    If Len(msg) <> 0 Then
'        msg2 = ReadFile(msg)
'        'If Len(msg2) <> 0 Then mdiNexIRC.ctlVBScript.ExecuteStatement msg2
'    End If
'End Select
'If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuExecuteScript_Click()"
'End Sub

Private Sub mnuCustomize_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmCustomize.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuCustomize_Click()"
End Sub

Private Sub mnuEditAutojoin_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmAutoJoin.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuEditAutojoin_Click()"
End Sub

Private Sub mnuExit_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UnloadProgram
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuExit_Click()"
End Sub

Private Sub mnuFakeKline_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.SendData "QUIT •: K-Lined" & vbCrLf
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuFakeKline_Click()"
End Sub

Private Sub mnuGoingToSleep_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.SendData "QUIT •: Going to sleep" & vbCrLf
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuGoingToSleep_Click()"
End Sub

Private Sub mnuGuestBook_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Surf "http://www.team-nexgen.org/guest/", Me.hWnd
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuGuestBook_Click()"
End Sub

Private Sub mnuHelp1_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DisplayHelpInformationByName mdiNexIRC.ActiveForm.Name
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuHelp1_Click()"
End Sub

Private Sub mnuHome_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Surf "http://www.team-nexgen.org", Me.hWnd
ActivateResize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuHome_Click()"
End Sub

Private Sub mnuInfo_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.SendData "Info" & vbCrLf
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuInfo_Click()"
End Sub

Private Sub mnuIRCOPS_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.SendData "IRCOPS" & vbCrLf
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuIRCOPS_Click()"
End Sub

Private Sub mnuJoinChannelName_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmJoinChannel.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuJoinChannelName_Click()"
End Sub

Private Sub MNUKLINES_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.SendData "KLINES" & vbCrLf
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub MNUKLINES_Click()"
End Sub

Private Sub mnuFindChannels_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmChannelListing.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuFindChannels_Click()"
End Sub

Private Sub mnuJoinNexgen_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.SendData "JOIN #nexgen" & vbCrLf
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuJoinNexgen_Click()"
End Sub

Private Sub mnuListChannels_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.SendData "LIST" & vbCrLf
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuListChannels_Click()"
End Sub

Private Sub mnuLUsers_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.SendData "LUSERS" & vbCrLf
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuLUsers_Click()"
End Sub

Private Sub mnuMessageForums_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Surf "http://www.team-nexgen.org", Me.hWnd
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuMessageForums_Click()"
End Sub

Private Sub mnuMessageServer_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmMessageServer.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuMessageServer_Click()"
End Sub

Private Sub mnuMinimize_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ActiveForm.WindowState = vbMinimized
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuMinimize_Click()"
End Sub

Private Sub mnuMinimizeAll_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To mdiNexIRC.Count
    ActiveForm.WindowState = vbMinimized
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuMinimizeAll_Click()"
End Sub

Private Sub mnuMixer_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sShowQuickmix = True Then
    lSettings.sShowQuickmix = False
    Unload frmMobileMixer
Else
    ToggleMixer True
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuMessageUser_Click()"
End Sub

Private Sub mnuMOTD_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.SendData "MOTD" & vbCrLf
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuMOTD_Click()"
End Sub

Private Sub mnuNewBot_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer
msg = InputBox("Enter Nickname:")
If Len(msg) <> 0 Then AddBot msg, bEggdrop
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuNewBot_Click()"
End Sub

Private Sub mnuNewBotCommand_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmAddBotCommand.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuNewBotCommand_Click()"
End Sub

Private Sub mnuNewChannel_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmJoinChannel.Show 0, Me
frmJoinChannel.Caption = "NexIRC - New Channel"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuNewChannel_Click()"
End Sub

Private Sub mnuNewConnection_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmQuickConnect.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuNewConnection_Click()"
End Sub

Private Sub mnuNewScriptFile_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
NewScriptFile
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuNewScriptFile_Click()"
End Sub

Private Sub mnuNewScriptFileRange_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmScriptRange.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuNewScriptFileRange_Click()"
End Sub

Private Sub mnuNewStatusWindow_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
NewStatusWindow lSettings.sServer, lSettings.sPort, False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuNewStatusWindow_Click()"
End Sub

Private Sub mnuNewTheme_Click()
NewSpectrumTheme
End Sub

Private Sub mnuNexgenHomepage_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmWeb.Visible = True
frmWeb.web.Navigate "http://www.team-nexgen.org"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuNexgenHomepage_Click()"
End Sub

Private Sub mnuNexIRCServer_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmIRCServer.Show
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuNexIRCServer_Click()"
End Sub

Private Sub mnuOpenLogFile_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim F As Form, msg As String, msg2 As String
msg = OpenDialog(Me, "LOG Files (*.log)|*.log|All Files (*.*)|*.*|", "Open LOG File ...", App.Path & "\data\logs\")
If Len(msg) <> 0 Then
    Set F = New frmTextEditor
    F.Show
    F.txtIncoming.Text = ReadFile(msg)
    msg = GetFileTitle(msg)
    F.Caption = msg
    F.Tag = ""
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuOpenLogFile_Click()"
End Sub

Private Sub mnuOpenSupportedMedia_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PromptPlayback Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuOpenSupportedMedia_Click()"
End Sub

Private Sub mnuOpenURL_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = InputBox("Enter URL:")
Surf msg, Me.hWnd
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuOpenURL_Click()"
End Sub

Private Sub mnuPause_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
MenuPause
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuPause_Click()"
End Sub

Private Sub mnuPlay_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
MenuPlay
End Sub

Private Sub mnuPlaylistCollection_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
With mnuPlaylistCollection(Index)
    If Len(.Caption) <> 0 Then
        SetPlaylistINIFile App.Path & "\data\playlists\" & Trim(.Caption)
        LoadPlaylist "", True
    End If
End With
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuPlaylistCollection_Click(Index As Integer)"
End Sub

Private Sub mnuOpenScriptFile_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
OpenScript PromptOpenScriptFile(Me)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuOpenScriptFile_Click()"
End Sub

Private Sub mnuOpenWebsiteURL_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = InputBox("Enter URL:")
If Len(msg) <> 0 Then
    frmWeb.Show
    frmWeb.web.Navigate msg
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuOpenWebsiteURL_Click()"
End Sub

Private Sub mnuQuickConnect_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmQuickConnect.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuQuickConnect_Click()"
End Sub

Private Sub mnuQuickPlay_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PromptPlayback Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuQuickPlay_Click()"
End Sub

Private Sub mnuRebooting_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.SendData "QUIT •: Rebooting" & vbCrLf
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuRebooting_Click()"
End Sub

Private Sub mnuRefresh_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sBackgroundWebpage = True Then
    frmWeb.web.Refresh
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuRefresh_Click()"
End Sub

Private Sub mnuRefreshPlaylist_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmPlaylist.RefreshPlaylist
DisplayPlaylists
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuRefreshPlaylist_Click()"
End Sub

'Private Sub mnuRegister_Click()
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
'frmRegister.Show 0, Me
'mdiNexIRC.SetFocus
'If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuRegister_Click()"
'End Sub

Private Sub mnuSave_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, F As Form, i As Integer
Set F = ActiveForm
Select Case F.Name
Case "frmWeb"
    If lSettings.sBackgroundWebpage = True Then
        msg = frmWeb.web.LocationURL
    End If
Case "frmChannels"
    msg2 = "Channels-" & Date$ & ".log"
    For i = 0 To F.lvwChannels.Count
        If Len(F.lvwChannels.ItemText(i)) <> "" Then
            msg = msg & vbCrLf & F.lvwChannels.ItemText(i)
        End If
    Next i
    If Len(msg) <> 0 Then
        SaveFile msg2, msg
    End If
Case "frmMOTD"
    msg = "MOTD-" & Date$ & ".log"
    If Len(msg) <> 0 Then
        SaveFile App.Path & "\data\logs\" & msg, F.txtMOTD.Text
        ProcessReplaceString sSaveLog, F.txtMOTD, msg
    End If
Case "frmQuery"
    msg = F.Caption & "-" & Date$ & ".log"
    If Len(msg) <> 0 Then
        SaveFile App.Path & "\data\logs\" & msg, F.txtIncoming.Text
        ProcessReplaceString sSaveLog, F.txtMOTD, msg
    End If
Case "frmChat"
    msg = F.Caption & "-" & Date$ & ".log"
    If Len(msg) <> 0 Then
        SaveFile App.Path & "\data\logs\" & msg, F.txtIncoming.Text
        ProcessReplaceString sSaveLog, F.txtMOTD, msg
    End If
Case "frmChannel"
    msg = F.Caption & "-" & Date$ & ".log"
    If Len(msg) <> 0 Then
        SaveFile App.Path & "\data\logs\" & msg, F.txtIncoming.Text
        ProcessReplaceString sSaveLog, F.txtMOTD, msg
    End If
Case "frmStatus"
    msg = "Status-" & Date$ & ".log"
    If Len(msg) <> 0 Then
        SaveFile App.Path & "\data\logs\" & msg, F.txtIncoming.Text
        ProcessReplaceString sSaveLog, F.txtMOTD, msg
    End If
Case "frmTextEditor"
    msg = Trim(F.txtIncoming.Tag)
    If Len(msg) = 0 Then
        msg = SaveDialog(Me, "Text Files (*.txt)|*.txt|", "Save as ...", CurDir)
        msg = Left(msg, Len(msg) - 1) & ".txt"
    End If
    If Len(msg) <> 0 Then
        msg2 = msg
        msg2 = GetFileTitle(msg2)
        If DoesFileExist(msg) = True Then Kill msg
        If Len(msg) <> 0 Then
            SaveFile msg, F.txtIncoming.Text
            F.Tag = ""
            F.txtIncoming.Tag = msg
            F.Caption = msg2
            If lSettings.sGeneralPrompts = True Then
                MsgBox msg2 & " saved", vbInformation
            End If
        End If
    End If
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuSave_Click()"
End Sub

Private Sub mnuSaveAs_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, F As Form, i As Integer
Set F = ActiveForm
Select Case F.Name
Case "frmStatus"
    msg = SaveDialog(Me, "LOG Files (*.log)|*.log|", "Save as ...", CurDir)
    If Len(msg) <> 0 Then
        msg = Left(msg, Len(msg) - 1) & ".log"
        SaveFile msg, F.txtIncoming.Text
        ProcessReplaceString sSaveLog, F.txtIncoming, msg
    End If
Case "frmTextEditor"
    msg = SaveDialog(Me, "Text Files (*.txt)|*.txt|", "Save as ...", App.Path & "\data\scripts\")
    If Len(msg) <> 0 Then
        msg = Left(msg, Len(msg) - 1) & ".txt"
        SaveFile msg, F.txtIncoming.Text
        F.Tag = ""
        msg2 = msg
        msg2 = GetFileTitle(msg2)
        RemoveTaskbar F.Caption
        F.Caption = msg2
        AddTaskPanel msg2, 1
    End If
Case "frmWeb"
    If lSettings.sBackgroundWebpage = True Then
        msg = frmWeb.web.LocationURL
    End If
Case "frmChannels"
    msg2 = SaveDialog(Me, "LOG Files (*.log)|*.log|", "Save as ...", CurDir)
    If Len(msg2) <> 0 Then
        msg2 = Left(msg2, Len(msg2) - 1) & ".log"
        For i = 0 To F.lvwChannels.Count
           If Len(F.lvwChannels.ItemText(i)) <> "" Then
                msg = msg & vbCrLf & F.lvwChannels.ItemText(i)
            End If
        Next i
    End If
    If Len(msg) <> 0 Then
        SaveFile msg2, msg
    End If
Case "frmMOTD"
    msg = SaveDialog(Me, "LOG Files (*.log)|*.log|", "Save as ...", CurDir)
    If Len(msg) <> 0 Then
        msg = Left(msg, Len(msg) - 1) & ".log"
        SaveFile App.Path & "\data\logs\" & msg, F.txtMOTD.Text
        ProcessReplaceString sSaveLog, F.txtIncoming, msg
    End If
Case "frmQuery"
    msg = SaveDialog(Me, "LOG Files (*.log)|*.log|", "Save as ...", CurDir)
    If Len(msg) <> 0 Then
        msg = Left(msg, Len(msg) - 1) & ".log"
        SaveFile App.Path & "\data\logs\" & msg, F.txtIncoming.Text
        ProcessReplaceString sSaveLog, F.txtIncoming, msg
    End If
Case "frmChat"
    msg = SaveDialog(Me, "LOG Files (*.log)|*.log|", "Save as ...", CurDir)
    If Len(msg) <> 0 Then
        msg = Left(msg, Len(msg) - 1) & ".log"
        SaveFile App.Path & "\data\logs\" & msg, F.txtIncoming.Text
        ProcessReplaceString sSaveLog, F.txtIncoming, msg
    End If
Case "frmChannel"
    msg = SaveDialog(Me, "LOG Files (*.log)|*.log|", "Save as ...", CurDir)
    If Len(msg) <> 0 Then
        msg = Left(msg, Len(msg) - 1) & ".log"
        SaveFile App.Path & "\data\logs\" & msg, F.txtIncoming.Text
        ProcessReplaceString sSaveLog, F.txtIncoming, msg
        
    End If
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuSaveAs_Click()"
End Sub

Private Sub mnuSavePlaylist_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer, msg2 As String
msg = InputBox("Enter Description:")
If Len(msg) <> 0 Then
    SetPlaylistINIFile App.Path & "\data\playlists\" & msg & ".ini"
    SavePlaylist
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuSavePlaylist_Click()"
End Sub

Private Sub mnuScriptBrowser_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmScriptManager.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuScriptBrowser_Click()"
End Sub

Private Sub mnuScriptEditor_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
NewScriptFile
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuScriptEditor_Click()"
End Sub

Private Sub mnuScripts2_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Surf "http://www.team-nexgen.org/scripts.shtml", Me.hWnd
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuScripts2_Click()"
End Sub

Private Sub mnuSearch_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Surf "http://www.metacrawler.com", Me.hWnd
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuSearch_Click()"
End Sub

Private Sub mnuSearchwithinPlaylist_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmSearchWithinPlaylist.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuSearchwithinPlaylist_Click()"
End Sub

Private Sub mnuSendQuitMessage_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
SendQuitMessage lSettings.sActiveServerForm
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuSendQuitMessage_Click()"
End Sub

Private Sub mnuSetupWizard_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmSetupWizard.Show 1, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuSetupWizard_Click()"
End Sub

Private Sub mnuShowBotControl_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmBots.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuShowBotControl_Click()"
End Sub

Private Sub mnuShowPlaylist_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmPlaylist.Show
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuShowPlaylist_Click()"
End Sub

Private Sub mnuShowTips_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmTip.Show 0, mdiNexIRC
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuShowTips_Click()"
End Sub

Private Sub mnuSoftware_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Surf "http://www.team-nexgen.org/software.shtml", Me.hWnd
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuSoftware_Click()"
End Sub

Private Sub mnuStaff_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Surf "http://www.team-nexgen.org/members.shtml", Me.hWnd
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuStaff_Click()"
End Sub

Private Sub mnuStop_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
MenuStop
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuStop_Click()"
End Sub

Private Sub mnuTileHorizontal_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
mdiNexIRC.Arrange vbTileHorizontal
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuTileHorizontal_Click()"
End Sub

Private Sub mnuTileVerticle_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
mdiNexIRC.Arrange vbTileVertical
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuTileVerticle_Click()"
End Sub

Private Sub mnuTime_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.SendData "TIME" & vbCrLf
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuTime_Click()"
End Sub

Private Sub picMP3OCX_DblClick()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
picMP3OCX.Visible = False
ActivateResize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picMP3OCX_DblClick()"
End Sub

Private Sub picMP3OCX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 2 Then PopupMenu frmMenus.mnuBackground
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picMP3OCX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picNotify_DblClick()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
MDIForm_Resize
picNotify.Visible = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picNotify_DblClick()"
End Sub

Private Sub picTopToolbar_DblClick()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
MDIForm_Resize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picTopToolbar_DblClick()"
End Sub

Private Sub picTopToolbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 2 Then
    PopupMenu frmMenus.mnuBackground
Else
    FormDrag Me
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picTopToolbar_DblClick()"
End Sub

Private Sub picTopToolbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ResetMainButtons
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picTopToolbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

'Private Sub script_Error()
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
'ProcessReplaceString sProgrammingError, lSettings.sActiveServerForm.txtIncoming, ctlVBScript.Error.Description, "Line: " & ctlVBScript.Error.Line & "(" & Str(ctlVBScript.Error.Number) & ")"
'PlayWav App.Path & "\data\sounds\err.wav", &H1
'If lSettings.sGeneralPrompts = True Then
'    MsgBox "Error: " & vbCrLf & ctlVBScript.Error.Description & vbCrLf & vbCrLf & "Line: " & ctlVBScript.Error.Line & vbCrLf & "Column: " & ctlVBScript.Error.Column & vbCrLf & "Source: " & ctlVBScript.Error.Text, vbCritical
'End If
'If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub script_Error()"
'End Sub

Private Sub StatusBar_PanelClick(ByVal Panel As MSComctlLib.Panel)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, X As Integer, re As Integer, F As Integer, msg As String
For i = 1 To StatusBar.Panels.Count
    If StatusBar.Panels.Item(i).Key = Panel Then
        StatusBar.Panels.Item(i).Bevel = sbrInset
        If Left(LCase(Panel), 6) = "status" Then
            For F = 1 To ReturnStatusWindowCount
                msg = LCase(Left(ReturnStatusWindowCaption(F), 8))
                If msg = LCase(Panel) Then
                    If LCase(Left(mdiNexIRC.ActiveForm.Caption, 8)) = LCase(Panel) Then
                        re = ShowWindow(ReturnStatusWindowHwnd(F), 9)
                        SetStatusWindowFocus F
                        SetStatusWindowState F, vbNormal
                    Else
                        re = ShowWindow(ReturnStatusWindowHwnd(F), 9)
                        SetStatusWindowFocus F
                        SetStatusWindowState F, vbNormal
                    End If
                End If
            Next F
        End If
        If LCase(Panel) = "notify" Then
            frmNotify.SetFocus
            frmNotify.WindowState = vbNormal
        End If
        If LCase(Panel) = "list" Then
'            frmChannels.SetFocus
            frmChannels.WindowState = vbNormal
            frmChannels.Visible = True
        End If
        If LCase(Panel) = "motd" Then
            frmMOTD.SetFocus
            frmMOTD.WindowState = vbNormal
        End If
        If LCase(Panel) = "server" Then
            frmIRCServer.SetFocus
            frmIRCServer.WindowState = vbNormal
        End If
        If LCase(Panel) = "manager" Then
            re = ShowWindow(frmConnectionManager.hWnd, 9)
            frmConnectionManager.SetFocus
            frmConnectionManager.WindowState = vbNormal
        End If
        If LCase(Panel) = "playlist" Then
            re = ShowWindow(frmPlaylist.hWnd, 9)
            frmPlaylist.SetFocus
        End If
        If Left(Panel, 1) = "#" Then
            For X = 1 To ReturnChannelUBound
                If LCase(ReturnChannelName(X)) = LCase(Panel) Then
                'If LCase(lChannelName(x)) = LCase(Panel) Then
                    If LCase(Me.ActiveForm.Caption) = LCase(ReturnChannelCaption(X)) Then
                    'If LCase(Me.ActiveForm.Caption) = LCase(lChannel(x).Caption) Then
                        re = ShowWindow(ReturnChannelHwnd(X), 6)
                        're = ShowWindow(lChannel(x).hWnd, 6)
                    Else
                        re = ShowWindow(ReturnChannelHwnd(X), 6)
                        're = ShowWindow(lChannel(x).hWnd, 9)
                        
                        SetFocusOnChannel X
                        SetChannelWindowState X, 0
                        'lChannel(x).WindowState = 0
                        
                    End If
                End If
            Next X
        Else
            For X = 1 To 150
                If LCase(ReturnQueryName(X)) = LCase(Panel) Then
                'If LCase(lQueryName(x)) = LCase(Panel) Then
                    If LCase(Me.ActiveForm.Caption) = LCase(ReturnQueryCaption(X)) Then
                    'If LCase(Me.ActiveForm.Caption) = LCase(lQuery(x).Caption) Then
                        're = ShowWindow(lQuery(x).hWnd, 6)
                        re = ShowWindow(ReturnQueryHwnd(X), 6)
                    Else
                        re = ShowWindow(ReturnQueryHwnd(X), 6)
                        're = ShowWindow(lQuery(x).hWnd, 9)
                        'lQuery(x).SetFocus
                        SetFocusOnQueryWindow X
                        SetQueryWindowState X, 0
                        
                    End If
                End If
            Next X
        End If
    Else
        StatusBar.Panels.Item(i).Bevel = sbrRaised
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub StatusBar_PanelClick(ByVal Panel As MSComctlLib.Panel)"
End Sub

Private Sub timeParse_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim word() As String, strType As String, parms As String, intcount As Integer
intcount = 1
Do Until intcount > colString.Count
    ParseIRCData colString.Item(intcount), lSettings.sActiveServerForm
    colString.Remove intcount
    intcount = intcount + 1
Loop
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub timeParse_Timer()"
End Sub

Private Sub tmrCheckButtonColors_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
RefreshColors
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrCheckButtonColors_Timer()"
End Sub

Private Sub tmrContinuousPlay_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sContinuousPlay = True Then
    ActivateContinuousPlay
End If
PicDir = GetRnd(4)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrContinuousPlay_Timer()"
End Sub

Private Sub tmrDIE_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sEnding = True
Unload mdiNexIRC
End Sub

Private Sub tmrEndSoon_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
CloseConnections
tmrDIE.Enabled = True
End Sub

Private Sub tmrPlaySoon_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
tmrPlaySoon.Enabled = False
ActivatePlayback
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrContinuousPlay_Timer()"
End Sub

Private Sub tmrSendUserPlaylist_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg  As String
lPlaylistSendIndex = lPlaylistSendIndex + 1
If lPlaylistSendIndex = lFiles.fCount Then tmrSendUserPlaylist.Enabled = False
If Len(lFiles.fFile(lPlaylistSendIndex).fFilename) <> 0 Then
    msg = lFiles.fFile(lPlaylistSendIndex).fFilename
    msg = GetFileTitle(msg)
    If eForm.tcp.State = sckConnected Then
        eForm.tcp.SendData "PRIVMSG " & eUsername & " :!" & msg & vbCrLf
    Else
        lPlaylistSendIndex = 0
        eUsername = ""
        Set eForm = Nothing
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrSendUserPlaylist_Timer()"
End Sub

Private Sub tmrUnloadSplashDelay_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSplashDelay = lSplashDelay + 1
If lSettings.sShowSplashOnStartup = True Then
    If lSettings.sByPassStartupScreen = True Then
        If lSplashDelay = 2 Then
            Unload frmSplash
            tmrUnloadSplashDelay.Enabled = False
        End If
    Else
        frmSplash.imgSplash.Visible = True
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrUnloadSplashDelay_Timer()"
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If KeyAscii = 13 Then
    cmdSend_Click
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtMessage_KeyPress(KeyAscii As Integer)"
End Sub

Private Sub txtUrl_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtUrl.SelStart = 0
txtUrl.SelLength = Len(txtUrl.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtUrl_GotFocus()"
End Sub

Private Sub txtUrl_KeyPress(KeyAscii As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If KeyAscii = 13 Then
    If Len(txtUrl.Text) <> 0 Then
        Surf txtUrl.Text, Me.hWnd
    End If
    KeyAscii = 0
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtUrl_KeyPress(KeyAscii As Integer)"
End Sub

Private Sub txtUrl_LostFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sBackgroundWebpage = True Then
    If frmWeb.Visible = True Then
        frmWeb.Visible = True
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtUrl_LostFocus()"
End Sub

Private Sub wskIdent_Close()

If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sIdent.iShow = True Then ProcessReplaceString sIdentClosed, lSettings.sActiveServerForm
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub wskIdent_Close()"
End Sub

Private Sub wskIdent_Connect()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sIdent.iShow = True Then ProcessReplaceString sIdentConnect, lSettings.sActiveServerForm
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub wskIdent_Close()"
End Sub

Private Sub wskIdent_ConnectionRequest(ByVal requestID As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sIdent.iEnabled = True Then
    wskIdent.Accept requestID
    wskIdent.SendData wskIdent.LocalPort & ", " & requestID & ":USERID:WIN32:" & lSettings.sIdent.iUserID & vbCrLf
    If lSettings.sIdent.iShow = True Then ProcessReplaceString sIdentConnection, lSettings.sActiveServerForm.txtIncoming, Trim(Str(requestID))
Else
    If lSettings.sIdent.iShow = True Then ProcessReplaceString sIdentRequestDenied, lSettings.sActiveServerForm, Trim(Str(requestID))
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub wskIdent_ConnectionRequest(ByVal requestID As Long)"
End Sub
