VERSION 5.00
Begin VB.Form frmSetupWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Setup Wizard"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetupWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin nexIRC.ctlXPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   103
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
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
      MICON           =   "frmSetupWizard.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdNext 
      Default         =   -1  'True
      Height          =   495
      Left            =   3600
      TabIndex        =   104
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "Next"
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
      MICON           =   "frmSetupWizard.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdFinish 
      Height          =   495
      Left            =   4800
      TabIndex        =   105
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "Finish"
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
      MICON           =   "frmSetupWizard.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdBack 
      Height          =   495
      Left            =   2400
      TabIndex        =   107
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "Back"
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
      MICON           =   "frmSetupWizard.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraSetupWizard 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   5
      Left            =   120
      TabIndex        =   57
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.OptionButton optEditAlternate 
         Appearance      =   0  'Flat
         Caption         =   "N&o"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   4680
         TabIndex        =   63
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optEditAlternate 
         Appearance      =   0  'Flat
         Caption         =   "&Yes"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   62
         Top             =   360
         Width           =   735
      End
      Begin VB.ListBox lstAlternateNicknames 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   2010
         Left            =   600
         TabIndex        =   61
         Top             =   720
         Width           =   4095
      End
      Begin nexIRC.ctlXPButton cmdAddAlternate 
         Height          =   375
         Left            =   4800
         TabIndex        =   112
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
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
         MICON           =   "frmSetupWizard.frx":007C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdRemoveAlternate 
         Height          =   375
         Left            =   4800
         TabIndex        =   113
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
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
         MICON           =   "frmSetupWizard.frx":0098
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label5 
         Caption         =   "Would you like to edit 'Alternate Nicknames'?"
         Height          =   255
         Left            =   600
         TabIndex        =   58
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label22 
         Caption         =   "Alternate Nicknames"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   94
         Top             =   0
         Width           =   4575
      End
      Begin VB.Image imgAlternate 
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Frame fraSetupWizard 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   2
      Left            =   120
      TabIndex        =   46
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.OptionButton optIgnoreEnabled 
         Appearance      =   0  'Flat
         Caption         =   "N&o"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   4680
         TabIndex        =   48
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.ListBox lstIgnore 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   2010
         Left            =   600
         TabIndex        =   50
         Top             =   720
         Width           =   4095
      End
      Begin VB.OptionButton optIgnoreEnabled 
         Appearance      =   0  'Flat
         Caption         =   "&Yes"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   49
         Top             =   360
         Width           =   735
      End
      Begin nexIRC.ctlXPButton cmdAddIgnore 
         Height          =   375
         Left            =   4800
         TabIndex        =   114
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
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
         MICON           =   "frmSetupWizard.frx":00B4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdRemoveIgnore 
         Height          =   375
         Left            =   4800
         TabIndex        =   115
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
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
         MICON           =   "frmSetupWizard.frx":00D0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label23 
         Caption         =   "Would you like to edit your ignore list?"
         Height          =   255
         Left            =   600
         TabIndex        =   95
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   "Ignore List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   47
         Top             =   0
         Width           =   5175
      End
      Begin VB.Image imgNicklist 
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Frame fraSetupWizard 
      BorderStyle     =   0  'None
      Caption         =   "Welcome"
      Height          =   2895
      Index           =   3
      Left            =   120
      TabIndex        =   51
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.ComboBox cboBotType 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmSetupWizard.frx":00EC
         Left            =   600
         List            =   "frmSetupWizard.frx":00FF
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   720
         Width           =   4095
      End
      Begin VB.ListBox lstBotlist 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   1620
         Left            =   600
         TabIndex        =   55
         Top             =   1080
         Width           =   4095
      End
      Begin VB.OptionButton optEditBots 
         Appearance      =   0  'Flat
         Caption         =   "N&o"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   4680
         TabIndex        =   53
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optEditBots 
         Appearance      =   0  'Flat
         Caption         =   "&Yes"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   52
         Top             =   360
         Width           =   735
      End
      Begin nexIRC.ctlXPButton cmdAddBot 
         Height          =   375
         Left            =   4800
         TabIndex        =   116
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
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
         MICON           =   "frmSetupWizard.frx":0140
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdRemoveBot 
         Height          =   375
         Left            =   4800
         TabIndex        =   117
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
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
         MICON           =   "frmSetupWizard.frx":015C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdEditBot 
         Height          =   375
         Left            =   4800
         TabIndex        =   118
         Top             =   1680
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Edit"
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
         MICON           =   "frmSetupWizard.frx":0178
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label24 
         Caption         =   "Would you like to edit your Bot list?"
         Height          =   255
         Left            =   600
         TabIndex        =   96
         Top             =   360
         Width           =   5175
      End
      Begin VB.Image imgDog 
         Height          =   480
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "IRC Bots"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   54
         Top             =   0
         Width           =   5175
      End
   End
   Begin VB.Frame fraSetupWizard 
      BorderStyle     =   0  'None
      Caption         =   "Welcome"
      Height          =   2895
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5775
      Begin VB.CheckBox chkFinishAfterProfile 
         Appearance      =   0  'Flat
         Caption         =   "&Finish after applying profile"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   2520
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.ComboBox cboProfile 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2160
         Width           =   2295
      End
      Begin VB.OptionButton optSetupWizardContinue 
         Appearance      =   0  'Flat
         Caption         =   "&Use this profile:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   4
         Top             =   1920
         Width           =   1455
      End
      Begin VB.OptionButton optSetupWizardContinue 
         Appearance      =   0  'Flat
         Caption         =   "&No"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   3
         Top             =   1680
         Width           =   735
      End
      Begin VB.OptionButton optSetupWizardContinue 
         Appearance      =   0  'Flat
         Caption         =   "&Yes"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   2
         Top             =   1680
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "Setup Wizard"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   0
         Top             =   0
         Width           =   3135
      End
      Begin VB.Image imgNexgen2 
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lblWelcomeInformation 
         Caption         =   $"frmSetupWizard.frx":0194
         Height          =   1215
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Frame fraSetupWizard 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.OptionButton optAutoJoinEnabled 
         Appearance      =   0  'Flat
         Caption         =   "N&o"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   4680
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.ComboBox cboAutojoinNetwork 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   720
         Width           =   4095
      End
      Begin VB.ListBox lstAutojoin 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   1620
         Left            =   600
         TabIndex        =   13
         Top             =   1080
         Width           =   4095
      End
      Begin VB.OptionButton optAutoJoinEnabled 
         Appearance      =   0  'Flat
         Caption         =   "&Yes"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin nexIRC.ctlXPButton cmdAddAutojoin 
         Height          =   375
         Left            =   4800
         TabIndex        =   106
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
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
         MICON           =   "frmSetupWizard.frx":0316
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdRemoveAutoJoin 
         Height          =   375
         Left            =   4800
         TabIndex        =   108
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
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
         MICON           =   "frmSetupWizard.frx":0332
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label6 
         Caption         =   "Would you like to edit your 'Autojoin List'?"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   360
         Width           =   5175
      End
      Begin VB.Image imgAutojoin 
         Height          =   480
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label26 
         Caption         =   "Autojoin List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Frame fraSetupWizard 
      BorderStyle     =   0  'None
      Caption         =   "Welcome"
      Height          =   2895
      Index           =   7
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.ComboBox cboTheme 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   2520
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.CheckBox chkColoredNicklist 
         Appearance      =   0  'Flat
         Caption         =   "&Colored Nicklist"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   2160
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkThemeToolbar 
         Appearance      =   0  'Flat
         Caption         =   "&Show Theme Toolbar"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox chkShowMixer 
         Appearance      =   0  'Flat
         Caption         =   "S&how Mixer"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   21
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkShowQuickNotify 
         Appearance      =   0  'Flat
         Caption         =   "Sh&ow Quick Notify"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkAutosizeTaskbar 
         Appearance      =   0  'Flat
         Caption         =   "&Autosize Taskbar"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox chkTimeStamping 
         Appearance      =   0  'Flat
         Caption         =   "&Time Stamping"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.OptionButton optTheme 
         Appearance      =   0  'Flat
         Caption         =   "&Default"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   17
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optTheme 
         Appearance      =   0  'Flat
         Caption         =   "&Everything"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optTheme 
         Appearance      =   0  'Flat
         Caption         =   "&Nothing"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin nexIRC.ctlXPButton cmdPreviewTheme 
         Height          =   375
         Left            =   4800
         TabIndex        =   109
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Preview"
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
         MICON           =   "frmSetupWizard.frx":034E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label31 
         Caption         =   "Theme:"
         Height          =   255
         Left            =   600
         TabIndex        =   26
         Top             =   2520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label27 
         Caption         =   "Would you like to select Interface Options?"
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   360
         Width           =   5175
      End
      Begin VB.Image imgSetup 
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label8 
         Caption         =   "Interface Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.Frame fraSetupWizard 
      BorderStyle     =   0  'None
      Caption         =   "Welcome"
      Height          =   2895
      Index           =   6
      Left            =   120
      TabIndex        =   64
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         Caption         =   "&Skip tips"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4080
         TabIndex        =   34
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optEverything 
         Appearance      =   0  'Flat
         Caption         =   "&Everything"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   33
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optNothing 
         Appearance      =   0  'Flat
         Caption         =   "&Nothing"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   32
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optStartupProfile 
         Appearance      =   0  'Flat
         Caption         =   "&Default"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   31
         Top             =   600
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CheckBox chkOptionsOnStartup 
         Appearance      =   0  'Flat
         Caption         =   "Show Options on Startup"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   39
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkShowTips 
         Appearance      =   0  'Flat
         Caption         =   "Show Tips"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   38
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkShowHomepageOnStartup 
         Appearance      =   0  'Flat
         Caption         =   "Show Homepage"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   37
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkByPassSplashScreen 
         Appearance      =   0  'Flat
         Caption         =   "By-pass Splash Screen"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   36
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkConnectOnStartup 
         Appearance      =   0  'Flat
         Caption         =   "Connect on Startup"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   35
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label28 
         Caption         =   "Startup Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   29
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Label7 
         Caption         =   "Would you like to edit your 'Startup Options'?"
         Height          =   255
         Left            =   600
         TabIndex        =   30
         Top             =   360
         Width           =   5175
      End
      Begin VB.Image imgSetup2 
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Frame fraSetupWizard 
      BorderStyle     =   0  'None
      Caption         =   "Welcome"
      Height          =   2895
      Index           =   1
      Left            =   120
      TabIndex        =   45
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.ListBox lstNotifyList 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   2010
         Left            =   600
         TabIndex        =   44
         Top             =   720
         Width           =   4095
      End
      Begin VB.OptionButton optNotifyEnabled 
         Appearance      =   0  'Flat
         Caption         =   "&Yes"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   42
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optNotifyEnabled 
         Appearance      =   0  'Flat
         Caption         =   "N&o"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   4680
         TabIndex        =   43
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
      Begin nexIRC.ctlXPButton cmdAddNotify 
         Height          =   375
         Left            =   4800
         TabIndex        =   110
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
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
         MICON           =   "frmSetupWizard.frx":036A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdRemoveNotify 
         Height          =   375
         Left            =   4800
         TabIndex        =   111
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
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
         MICON           =   "frmSetupWizard.frx":0386
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label29 
         Caption         =   "Notify List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   40
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Would you like to edit your 'Notify List?"
         Height          =   255
         Left            =   600
         TabIndex        =   41
         Top             =   360
         Width           =   5175
      End
      Begin VB.Image imgNotify 
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Frame fraSetupWizard 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   8
      Left            =   120
      TabIndex        =   65
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.OptionButton optEditServer 
         Appearance      =   0  'Flat
         Caption         =   "N&o"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   4680
         TabIndex        =   67
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.ComboBox cboKickLength 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox cboMaxNicklen 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   2520
         Width           =   855
      End
      Begin VB.ComboBox cboPort 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtDefaultQuit 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   85
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtDefaultTopic 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   83
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ComboBox cboAwayLen 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   1080
         Width           =   855
      End
      Begin VB.ComboBox cboSessions 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   73
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtServerName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   71
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox chkShowServerOnStartup 
         Appearance      =   0  'Flat
         Caption         =   "Show Server on Startup"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3480
         TabIndex        =   69
         Top             =   1800
         Width           =   2055
      End
      Begin VB.OptionButton optEditServer 
         Appearance      =   0  'Flat
         Caption         =   "&Yes"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   68
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Nick Length:"
         Height          =   255
         Left            =   600
         TabIndex        =   76
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Port:"
         Height          =   255
         Left            =   600
         TabIndex        =   74
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Kick Length:"
         Height          =   255
         Left            =   3480
         TabIndex        =   86
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Default Quit:"
         Height          =   255
         Left            =   600
         TabIndex        =   84
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "Default Topic:"
         Height          =   255
         Left            =   600
         TabIndex        =   82
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Away Length:"
         Height          =   255
         Left            =   3480
         TabIndex        =   80
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Max Sessions:"
         Height          =   255
         Left            =   3480
         TabIndex        =   78
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Description:"
         Height          =   255
         Left            =   600
         TabIndex        =   72
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Name:"
         Height          =   255
         Left            =   600
         TabIndex        =   70
         Top             =   720
         Width           =   975
      End
      Begin VB.Image imgServer 
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label9 
         Caption         =   "Would you like to edit 'Server Options'?"
         Height          =   255
         Left            =   600
         TabIndex        =   66
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label Label30 
         Caption         =   "Server Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   97
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.Frame fraSetupWizard 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   9
      Left            =   120
      TabIndex        =   59
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox txtProfileName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   91
         Top             =   1080
         Width           =   3735
      End
      Begin VB.OptionButton optProfile 
         Appearance      =   0  'Flat
         Caption         =   "&Save a Profile"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   89
         Top             =   720
         Width           =   3015
      End
      Begin VB.OptionButton optProfile 
         Appearance      =   0  'Flat
         Caption         =   "&Do not save Profile"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   88
         Top             =   720
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.Label lblEditAutoPerform 
         Caption         =   "Edit Auto Preform"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   600
         MouseIcon       =   "frmSetupWizard.frx":03A2
         MousePointer    =   99  'Custom
         TabIndex        =   119
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label33 
         Caption         =   "Register NexIRC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   600
         MouseIcon       =   "frmSetupWizard.frx":04F4
         MousePointer    =   99  'Custom
         TabIndex        =   102
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblMenuEditor 
         Caption         =   "Edit Menus"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   600
         MouseIcon       =   "frmSetupWizard.frx":0646
         MousePointer    =   99  'Custom
         TabIndex        =   101
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblViewHelp 
         Caption         =   "Edit Auto Connect"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   600
         MouseIcon       =   "frmSetupWizard.frx":0798
         MousePointer    =   99  'Custom
         TabIndex        =   100
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label lblAddAudioToYourPlaylist 
         Caption         =   "Add Audio to Playlist"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   600
         MouseIcon       =   "frmSetupWizard.frx":08EA
         MousePointer    =   99  'Custom
         TabIndex        =   99
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label32 
         Caption         =   "Additional Setup Tasks:"
         Height          =   255
         Left            =   600
         TabIndex        =   98
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label20 
         Caption         =   "If you wish to save these settings, click save profile"
         Height          =   255
         Left            =   600
         TabIndex        =   92
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label14 
         Caption         =   "Profile &Name:"
         Height          =   255
         Left            =   600
         TabIndex        =   90
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Image imgNexgen 
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Setup Wizard Complete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   60
         Top             =   0
         Width           =   5175
      End
   End
   Begin VB.Label Label21 
      Caption         =   "Alternate Nicknames"
      Height          =   375
      Left            =   720
      TabIndex        =   93
      Top             =   120
      Width           =   5175
   End
   Begin VB.Image Image6 
      Height          =   975
      Left            =   120
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "frmSetupWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lFrame As Integer
Private lFrameCount As Integer
Private lSpectrumThemeIndex As Integer

Public Sub ResetSetupWizardFrames(lFrameIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To lFrameCount
    fraSetupWizard(i).Visible = False
Next i
fraSetupWizard(lFrameIndex).Visible = True
lFrame = lFrameIndex
If lFrame = lFrameCount Then
    cmdFinish.Enabled = True
    cmdNext.Enabled = False
    cmdBack.Enabled = False
    cmdCancel.Enabled = False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ResetSetupWizardFrames(lFrameIndex As Integer)"
End Sub

Private Sub cboAutojoinNetwork_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
FillListBoxWithAutoJoin lstAutojoin
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cboAutojoinNetwork_Click()"
End Sub

Private Sub cboBotType_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
lstBotlist.Clear
For i = 0 To 150
    If Len(ReturnBotNickname(i)) <> 0 Then
        If ReturnBotType(i) = cboBotType.ListIndex Then lstBotlist.AddItem ReturnBotNickname(i)
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cboBotType_Change()"
End Sub

Private Sub cboTheme_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSpectrumThemeIndex = cboTheme.ListIndex
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cboTheme_Click()"
End Sub

Private Sub chkAutosizeTaskbar_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sAutosizeStatusbarItems = GetCheckboxValue(chkAutosizeTaskbar)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkAutosizeTaskbar_Click()"
End Sub

Private Sub chkByPassSplashScreen_Click()
lSettings.sByPassStartupScreen = GetCheckboxValue(chkByPassSplashScreen)
End Sub

Private Sub chkColoredNicklist_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sColoredNicklist = GetCheckboxValue(chkColoredNicklist)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkColoredNicklist_Click()"
End Sub

Private Sub chkConnectOnStartup_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sConnectOnStartup = GetCheckboxValue(chkConnectOnStartup)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkConnectOnStartup_Click()"
End Sub

Private Sub chkOptionsOnStartup_Click()
lSettings.sShowOptionsOnStartup = GetCheckboxValue(chkOptionsOnStartup)
End Sub

Private Sub chkShowHomepageOnStartup_Click()
lSettings.sNavigateOnStartup = GetCheckboxValue(chkShowHomepageOnStartup)
End Sub

Private Sub chkShowMixer_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sShowQuickmix = GetCheckboxValue(chkShowMixer)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkShowMixer_Click()"
End Sub

Private Sub chkShowQuickNotify_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sShowQuickNotify = GetCheckboxValue(chkShowQuickNotify)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkShowQuickNotify_Click()"
End Sub

Private Sub chkShowServerOnStartup_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sShowServerOnStartup = GetCheckboxValue(chkShowServerOnStartup)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkShowServerOnStartup_Click()"
End Sub

Private Sub chkShowTips_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sShowTips = GetCheckboxValue(chkShowTips)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkShowTips_Click()"
End Sub

Private Sub chkThemeToolbar_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sAlwaysShowAudioSettings = GetCheckboxValue(chkThemeToolbar)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkThemeToolbar_Click()"
End Sub

Private Sub chkTimeStamping_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sTimeStamping = GetCheckboxValue(chkTimeStamping)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkTimeStamping_Click()"
End Sub

Private Sub cmdAddAlternate_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = InputBox("Enter Alternate Nickname:", "Setup Wizard", "")
If msg = "@guidex" Then
    AddAlternate "|guide|"
    AddAlternate "|gdx300|"
    AddAlternate "|guidX|"
    AddAlternate "|gu!deX|"
    lstAlternateNicknames.AddItem "|guide|"
    lstAlternateNicknames.AddItem "|gdx300|"
    lstAlternateNicknames.AddItem "|guidX|"
    lstAlternateNicknames.AddItem "|gu!deX|"
    Exit Sub
Else
    lstAlternateNicknames.AddItem msg
    AddAlternate msg
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAddAlternate_Click()"
End Sub

Private Sub cmdAddAutojoin_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = InputBox("Enter Channel name:", "NexIRC - Setup Wizard")
If LCase(msg) = "@nexgen" Then
    lstAutojoin.AddItem "#nexgen"
    lstAutojoin.AddItem "#nexgentrivia"
    lstAutojoin.AddItem "#acidmax"
    lstAutojoin.AddItem "#nexirc"
    AddAutoJoin "#nexgen", cboAutojoinNetwork.Text
    AddAutoJoin "#nexgentrivia", cboAutojoinNetwork.Text
    AddAutoJoin "#acidmax", cboAutojoinNetwork.Text
    AddAutoJoin "#nexirc", cboAutojoinNetwork.Text
    Exit Sub
End If
If Len(msg) <> 0 Then
    lstAutojoin.AddItem msg
    AddAutoJoin msg, cboAutojoinNetwork.Text
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAddAutojoin_Click()"
End Sub

Private Sub cmdAddBot_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer
msg = InputBox("Enter nickname:", "NexIRC - Setup Wizard")
If msg = "@undernet" Then msg = "x@channels.undernet.org"
If Len(msg) <> 0 Then
    Select Case cboBotType.ListIndex
    Case 0
        AddBot msg, bUnknownBot
    Case 1
        AddBot msg, bEggdrop
    Case 2
        AddBot msg, bX
    Case 3
        AddBot msg, bChanServ
    Case 4
        AddBot msg, bMemoServ
    End Select
    lstBotlist.AddItem msg
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAddNotify_Click()"
End Sub

Private Sub cmdAddIgnore_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = InputBox("Enter nickname:", "NexIRC - Setup Wizard")
If msg = "@newnet" Then
    AddToIgnore "nBouncer"
    lstIgnore.AddItem "nBouncer"
    Exit Sub
End If
If Len(msg) <> 0 Then
    AddToIgnore msg
    lstIgnore.AddItem msg
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAddNotify_Click()"
End Sub

Private Sub cmdAddNotify_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer, j As Integer, n As Integer
msg = InputBox("Enter nickname:", "NexIRC - Setup Wizard")
If msg = "@nexgen" Then
    AddNotify "KnightFal"
    lstNotifyList.AddItem "KnightFal"
    AddNotify "Magique"
    lstNotifyList.AddItem "Magique"
    AddNotify "|guideX|"
    lstNotifyList.AddItem "|guideX|"
    AddNotify "byte187"
    lstNotifyList.AddItem "byte187"
    AddNotify "Alien1"
    lstNotifyList.AddItem "Alien1"
    AddNotify "Cheesewiz"
    lstNotifyList.AddItem "Cheesewiz"
    AddNotify "PC_Tech"
    lstNotifyList.AddItem "PC_Tech"
    AddNotify "Thorne^"
    lstNotifyList.AddItem "Thorne^"
    Exit Sub
End If
If Len(msg) <> 0 Then
    j = 101
    For i = 1 To 150
        If Len(ReturnNotifyNickname(i)) = 0 Then
            If j = 101 Then
                j = i
                SetNotifyNickname j, msg
                n = n + 1
            End If
        Else
            n = n + 1
        End If
    Next i
    SaveNotify
    lstNotifyList.AddItem msg
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAddNotify_Click()"
End Sub

Private Sub cmdBack_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lFrame <> 0 Then
    lFrame = lFrame - 1
    ResetSetupWizardFrames lFrame
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdBack_Click()"
End Sub

Private Sub cmdCancel_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lFrame = 0
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdCancel_Click()"
End Sub

Private Sub cmdEditBot_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer
msg = lstBotlist.Text
i = FindBotIndex(msg)
msg2 = InputBox("Edit Bot Nickname:", "Bots", msg)
If Len(msg2) <> 0 Then
    If msg = msg2 Then
        GoTo NDing
    Else
        lstBotlist.RemoveItem lstBotlist.ListIndex
        lstBotlist.AddItem msg2
        RemoveBot msg
        AddBot msg2, cboBotType.ListIndex
    End If
End If
NDing:
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdEditBot_Click()"
End Sub

Private Sub cmdFinish_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, F As Integer, d As Integer
If optProfile(0).Enabled = True Then
    If lSpectrumThemeIndex <> 0 Then
        lSpectrumThemes.sIndex = lSpectrumThemeIndex + 1
        WriteINI GetINIFile(iSpectrum), "Settings", "Index", lSpectrumThemeIndex + 1
    End If
    If Len(txtProfileName.Text) <> 0 Then
        F = Int(ReadINI(App.Path & "\data\config\profiles\0.ini", "Settings", "Count", 0)) + 1
        If F <> 0 Then
            msg = App.Path & "\data\config\profiles\" & Trim(Str(F)) & ".ini"
            If Len(msg) <> 0 Then
                If Len(txtProfileName.Text) <> 0 Then
                    WriteINI App.Path & "\data\config\profiles\0.ini", "Settings", "Count", Trim(Str(F))
                    d = d + 1
                    WriteINI msg, Trim(Str(d)), "Data", txtProfileName.Text
                    WriteINI msg, Trim(Str(d)), "Type", "0"
                    If lstAlternateNicknames.ListCount <> 0 Then
                        For i = 0 To lstAlternateNicknames.ListCount
                            If Len(lstAlternateNicknames.List(i)) <> 0 Then
                                d = d + 1
                                WriteINI msg, Trim(Str(d)), "Data", lstAlternateNicknames.List(i)
                                WriteINI msg, Trim(Str(d)), "Type", "1"
                            End If
                        Next i
                    End If
                    If ReturnAutoJoinCount <> 0 Then
                        For i = 0 To ReturnAutoJoinCount
                            If CheckAutoJoin(i) = True Then
                                d = d + 1
                                WriteINI msg, Trim(Str(d)), "Data", ReturnAutoJoinChannel(i) & "(" & ReturnAutoJoinNetwork(i) & ")"
                                WriteINI msg, Trim(Str(d)), "Type", "2"
                            End If
                        Next i
                    End If
                    If lstNotifyList.ListCount <> 0 Then
                        For i = 0 To lstNotifyList.ListCount
                            If Len(lstNotifyList.List(i)) <> 0 Then
                                d = d + 1
                                WriteINI msg, Trim(Str(d)), "Data", lstNotifyList.List(i)
                                WriteINI msg, Trim(Str(d)), "Type", "3"
                            End If
                        Next i
                    End If
                    If lstIgnore.ListCount <> 0 Then
                        For i = 0 To lstIgnore.ListCount
                            If Len(lstIgnore.List(i)) <> 0 Then
                                d = d + 1
                                WriteINI msg, Trim(Str(d)), "Data", lstIgnore.List(i)
                                WriteINI msg, Trim(Str(d)), "Type", "4"
                            End If
                        Next i
                    End If
                    If ReturnBotCount <> 0 Then
                        For i = 0 To ReturnBotCount
                            If Len(ReturnBotNickname(i)) <> 0 Then
                                d = d + 1
                                Select Case ReturnBotNickname(i)
                                Case bUnknownBot
                                    WriteINI msg, Trim(Str(d)), "Data", ReturnBotNickname(i) & "(Unknown)"
                                Case bEggdrop
                                    WriteINI msg, Trim(Str(d)), "Data", ReturnBotNickname(i) & "(Eggdrop)"
                                Case bX
                                    WriteINI msg, Trim(Str(d)), "Data", ReturnBotNickname(i) & "(Undernet)"
                                Case bChanServ
                                    WriteINI msg, Trim(Str(d)), "Data", ReturnBotNickname(i) & "(Chanserv)"
                                Case bMemoServ
                                    WriteINI msg, Trim(Str(d)), "Data", ReturnBotNickname(i) & "(Memoserv)"
                                End Select
                                WriteINI msg, Trim(Str(d)), "Type", "5"
                            End If
                        Next i
                    End If
                    If optTheme(0).Value = True Then
                        d = d + 1
                        WriteINI msg, Trim(Str(d)), "Data", "0"
                        WriteINI msg, Trim(Str(d)), "Type", "6"
                    ElseIf optTheme(1).Value = True Then
                        d = d + 1
                        WriteINI msg, Trim(Str(d)), "Data", "1"
                        WriteINI msg, Trim(Str(d)), "Type", "6"
                    ElseIf optTheme(2).Value = True Then
                        d = d + 1
                        WriteINI msg, Trim(Str(d)), "Data", "2"
                        WriteINI msg, Trim(Str(d)), "Type", "6"
                    End If
                    If optStartupProfile.Value = True Then
                        d = d + 1
                        WriteINI msg, Trim(Str(d)), "Data", "0"
                        WriteINI msg, Trim(Str(d)), "Type", "7"
                    ElseIf optNothing.Value = True Then
                        d = d + 1
                        WriteINI msg, Trim(Str(d)), "Data", "1"
                        WriteINI msg, Trim(Str(d)), "Type", "7"
                    ElseIf optEverything.Value = True Then
                        d = d + 1
                        WriteINI msg, Trim(Str(d)), "Data", "2"
                        WriteINI msg, Trim(Str(d)), "Type", "7"
                    ElseIf Option1.Value = True Then
                        d = d + 1
                        WriteINI msg, Trim(Str(d)), "Data", "3"
                        WriteINI msg, Trim(Str(d)), "Type", "7"
                    End If
                    'If lRegInfo.rRegistered = True Then
                    'd = d + 1
                    'WriteINI msg, Trim(Str(d)), "Data", Trim(lRegInfo.rName)
                    'WriteINI msg, Trim(Str(d)), "Type", "50"
                    'd = d + 1
                    ''WriteINI msg, Trim(Str(d)), "Data", Trim(lRegInfo.rPassword)
                    'writeINI msg, Trim(Str(d)), "Type", "51"
                    '''End If
                    d = d + 1
                    WriteINI msg, Trim(Str(d)), "Data", "SaveSettings"
                    WriteINI msg, Trim(Str(d)), "Type", "100"
                    WriteINI msg, "Settings", "Count", Trim(Str(d))
                End If
            End If
        End If
    Else
        If optProfile(0).Value = True Then
            PlayWav App.Path & "\data\sounds\err.wav", SND_ASYNC
            MsgBox "If you wish to save a profile, you must enter a profile name!", vbExclamation
            txtProfileName.SetFocus
            Exit Sub
        End If
    End If
End If
lFrame = 0
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdFinish_Click()"
End Sub

Private Sub cmdMoreThemes_Click()

End Sub

Private Sub cmdNext_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim c As Integer, E As Integer, b As Integer, i As Integer, F As Integer, msg As String, msg2 As String, msg3 As String, lCurrent As Boolean, lReg As String
If lFrameCount <> lFrame Then
    i = lFrame + 1
    Select Case i
    Case 1
        If optSetupWizardContinue(2).Value = True Then
            F = Int(ReadINI(App.Path & "\data\config\profiles\0.ini", "Settings", "Count", 0))
            For b = 1 To F
                msg = App.Path & "\data\config\profiles\" & Trim(Str(b)) & ".ini"
                If DoesFileExist(msg) Then
                    For E = 0 To Int(ReadINI(msg, "Settings", "Count", 0))
                        c = Int(ReadINI(msg, Trim(Str(E)), "Type", -1))
                        Select Case c
                        Case 0
                            msg2 = ReadINI(msg, Trim(Str(E)), "Data", "")
                            If Trim(LCase(msg2)) = Trim(LCase(cboProfile.Text)) Then
                                lCurrent = True
                            Else
                                lCurrent = False
                            End If
                        Case 1
                            If lCurrent = True Then
                                msg2 = ReadINI(msg, Trim(Str(E)), "Data", "")
                                If Len(msg2) <> 0 Then
                                    AddAlternate msg2
                                    lstAlternateNicknames.AddItem msg2
                                End If
                            End If
                        Case 2
                            If lCurrent = True Then
                                msg2 = ReadINI(msg, Trim(Str(E)), "Data", "")
                                If Len(msg2) <> 0 Then
                                    msg3 = Parse(msg2, "(", ")")
                                    If Len(msg3) <> 0 Then
                                        msg2 = Left(msg2, Len(msg2) - Len(msg3) - 2)
                                        AddAutoJoin msg2, msg3
                                        If LCase(msg3) = LCase(cboAutojoinNetwork.Text) Then lstAutojoin.AddItem msg2
                                    End If
                                End If
                            End If
                        Case 3
                            If lCurrent = True Then
                                msg2 = ReadINI(msg, Trim(Str(E)), "Data", "")
                                If Len(msg2) <> 0 Then
                                    AddNotify msg2
                                    lstNotifyList.AddItem msg2
                                End If
                            End If
                        Case 4
                            If lCurrent = True Then
                                msg2 = ReadINI(msg, Trim(Str(E)), "Data", "")
                                If Len(msg2) <> 0 Then
                                    AddToIgnore msg2
                                    lstIgnore.AddItem msg2
                                End If
                            End If
                        Case 5
                            If lCurrent = True Then
                                msg2 = ReadINI(msg, Trim(Str(E)), "Data", "")
                                If Len(msg2) <> 0 Then
                                    msg3 = Parse(msg2, "(", ")")
                                    If Len(msg3) <> 0 Then
                                        msg2 = Left(msg2, Len(msg2) - Len(msg3) - 2)
                                        Select Case LCase(msg3)
                                        Case "undernet"
                                            If LCase(cboBotType.Text) = "undernet x" Then
                                                lstBotlist.AddItem msg2
                                            End If
                                            AddBot msg2, bX
                                        Case "eggdrop"
                                            If LCase(cboBotType.Text) = "eggdrop" Then
                                                lstBotlist.AddItem msg2
                                            End If
                                            AddBot msg2, bEggdrop
                                        Case "chanserv"
                                            If LCase(cboBotType.Text) = "chanserv" Then
                                                lstBotlist.AddItem msg2
                                            End If
                                            AddBot msg2, bChanServ
                                        Case "custom"
                                            If LCase(cboBotType.Text) = "unknown/custom bot" Then
                                                lstBotlist.AddItem msg2
                                            End If
                                            AddBot msg2, bUnknownBot
                                        Case "memoserv"
                                            If LCase(cboBotType.Text) = "memoserv" Then
                                                lstBotlist.AddItem msg2
                                            End If
                                            AddBot msg2, bMemoServ
                                        End Select
                                    End If
                                End If
                            End If
                        Case 6
                            If lCurrent = True Then
                                msg2 = ReadINI(msg, Trim(Str(E)), "Data", "")
                                If Len(msg2) <> 0 Then
                                    Select Case Int(msg2)
                                    Case 0
                                        optStartupProfile.Value = True
                                    Case 1
                                        optNothing.Value = True
                                    Case 2
                                        optEverything.Value = True
                                    Case 3
                                        Option1.Value = True
                                    End Select
                                End If
                            End If
                        Case 7
                            If lCurrent = True Then
                                msg2 = ReadINI(msg, Trim(Str(E)), "Data", "")
                                If Len(msg2) <> 0 Then optTheme(Int(msg2)).Value = True
                            End If
                        'Case 50
                        '    If lCurrent = True Then
                        '        msg2 = ReadINI(msg, Trim(Str(E)), "Data", "")
                        '        If Len(msg2) <> 0 Then
                        '            lRegInfo.rName = msg2
                        '        End If
                        '    End If
                        'Case 51
                        '    If lCurrent = True Then
                        '        msg2 = ReadINI(msg, Trim(Str(E)), "Data", "")
                        '        If Len(msg2) <> 0 Then
                        '            lRegInfo.rPassword = msg2
                        '            If Len(lRegInfo.rName) <> 0 And Len(lRegInfo.rPassword) <> 0 Then
                        '                lReg = KeyGen(lRegInfo.rName, "pickles", 1)
                        '                If lReg = lRegInfo.rPassword Then
                        '                    lRegInfo.rRegistered = True
                        '                    WriteINI GetINIFile(iIRC), "REGInfo", "NAME", lRegInfo.rName
                        '                    WriteINI GetINIFile(iIRC), "REGInfo", "PASSWORD", lRegInfo.rPassword
                        '                    mdiNexIRC.Caption = "NexIRC (Registered Version)"
                        '                End If
                        '            Else
                        '                lRegInfo.rRegistered = False
                        '            End If
                        '        End If
                        '    End If
                        Case Else
                            If lCurrent = True Then
                                msg2 = ReadINI(msg, Trim(Str(E)), "Data", "")
                                If Len(msg2) <> 0 Then
                                    Select Case LCase(msg2)
                                    Case "savesettings"
                                        SaveSettings
                                    End Select
                                End If
                            End If
                        End Select
                    Next E
                End If
            Next b
            If chkFinishAfterProfile.Value = 1 Then
                lFrame = 9
                cmdNext.Enabled = False
                cmdFinish.Enabled = True
                cmdBack.Enabled = False
                cmdCancel.Enabled = False
                fraSetupWizard(0).Visible = False
                fraSetupWizard(1).Visible = False
                fraSetupWizard(2).Visible = False
                fraSetupWizard(3).Visible = False
                fraSetupWizard(4).Visible = False
                fraSetupWizard(5).Visible = False
                fraSetupWizard(6).Visible = False
                fraSetupWizard(7).Visible = False
                fraSetupWizard(8).Visible = False
                fraSetupWizard(9).Visible = True
                cmdFinish_Click
                Exit Sub
            End If
        ElseIf optSetupWizardContinue(1).Value = True Then
            cmdFinish.Caption = "Exit"
            Label1.Caption = "Setup Wizard did not complete"
            Label20.Caption = "Click 'Exit'"
            txtProfileName.Enabled = False
            optProfile(0).Enabled = False
            optProfile(1).Enabled = False
            lFrame = 9
            cmdNext.Enabled = False
            cmdFinish.Enabled = True
            cmdBack.Enabled = False
            cmdCancel.Enabled = False
            fraSetupWizard(0).Visible = False
            fraSetupWizard(1).Visible = False
            fraSetupWizard(2).Visible = False
            fraSetupWizard(3).Visible = False
            fraSetupWizard(4).Visible = False
            fraSetupWizard(5).Visible = False
            fraSetupWizard(6).Visible = False
            fraSetupWizard(7).Visible = False
            fraSetupWizard(8).Visible = False
            fraSetupWizard(9).Visible = True
            Exit Sub
        End If
    Case 2
        If lstNotifyList.ListCount = 0 And optNotifyEnabled(1).Value = True Then
            MsgBox "You must at least add one entry to your notify list for it to be enabled.", vbExclamation
            Exit Sub
        Else
            If lstNotifyList.ListCount <> 0 Then SetNotifyEnabled True
        End If
    Case 3
        If lstIgnore.ListCount = 0 And optIgnoreEnabled(1).Value = True Then
            MsgBox "You must at least add one entry to your ignore list for it to be enabled.", vbExclamation
            Exit Sub
        Else
            If lstIgnore.ListCount <> 0 Then SetIgnoreEnabled True
        End If
    Case 9
        WriteINI GetINIFile(iIRCServer), "Settings", "ServerName", txtServerName.Text
        WriteINI GetINIFile(iIRCServer), "Settings", "Description", txtDescription.Text
        WriteINI GetINIFile(iIRCServer), "Channel Defaults", "Topic", txtDefaultTopic.Text
        WriteINI GetINIFile(iIRCServer), "Default User Settings", "Default Quit Msg", txtDefaultQuit.Text
        WriteINI GetINIFile(iIRCServer), "Settings", "Port", cboPort.Text
        WriteINI GetINIFile(iIRCServer), "Settings", "MaxNickLength", cboMaxNicklen.Text
        WriteINI GetINIFile(iIRCServer), "Settings", "Session Limit", cboSessions.Text
        WriteINI GetINIFile(iIRCServer), "Settings", "AwayLen", cboAwayLen.Text
        WriteINI GetINIFile(iIRCServer), "Settings", "KickLen", cboKickLength.Text
    End Select
    ResetSetupWizardFrames i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdRemoveNotify_Click()"
End Sub

Private Sub cmdPreviewTheme_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmQuickImage.picImage.Picture = LoadPicture(lSpectrumThemes.sSpectrumTheme(FindSpectrumThemeByName(cboTheme.Text)).sScreenShot)
frmQuickImage.Show 1
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdPreviewTheme_Click()"
End Sub

Private Sub cmdRemoveAutojoin_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
i = FindAutoJoinIndex(lstAutojoin.Text)
If i <> 0 Then
    DeleteAutoJoin i
    lstAutojoin.RemoveItem lstAutojoin.ListIndex
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdRemoveAutojoin_Click()"
End Sub

Private Sub cmdRemoveBot_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
RemoveBot lstBotlist.Text
lstBotlist.RemoveItem lstBotlist.ListIndex
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdRemoveBot_Click()"
End Sub

Private Sub cmdRemoveNotify_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lstNotifyList.RemoveItem lstNotifyList.ListIndex
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdRemoveNotify_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, p As Integer, c As Integer, E As Integer, msg2 As String, msg As String, msg3 As String
lSettings.sSetupWizardVisible = True
imgServer.Picture = LoadPicture(App.Path & "\data\images\server.gif")
imgNexgen.Picture = LoadPicture(App.Path & "\data\icons\nexgen.ico")
imgAlternate.Picture = LoadPicture(App.Path & "\data\icons\alternates.ico")
imgNicklist.Picture = LoadPicture(App.Path & "\data\icons\nicklist.ico")
imgDog.Picture = LoadPicture(App.Path & "\data\images\dog.gif")
imgNexgen2.Picture = LoadPicture(App.Path & "\data\icons\nexgen.ico")
imgAutojoin.Picture = LoadPicture(App.Path & "\data\icons\channel.ico")
imgSetup.Picture = LoadPicture(App.Path & "\data\images\setup.gif")
imgSetup2.Picture = LoadPicture(App.Path & "\data\images\setup.gif")
imgNotify.Picture = LoadPicture(App.Path & "\data\icons\alternates.ico")
For i = 0 To lSpectrumThemes.sCount
    If Len(lSpectrumThemes.sSpectrumTheme(i).sName) <> 0 Then
        cboTheme.AddItem lSpectrumThemes.sSpectrumTheme(i).sName
    End If
Next i
cboTheme.ListIndex = 0
For i = 1 To 512
    cboKickLength.AddItem Str(i)
Next i
For i = 1 To 256
    cboAwayLen.AddItem Str(i)
Next i
For i = 1 To 100
    cboSessions.AddItem Str(i)
Next i
For i = 1 To 100
    cboMaxNicklen.AddItem Str(i)
Next i
For i = 81 To 9999
    cboPort.AddItem Str(i)
Next i
txtServerName.Text = ReadINI(GetINIFile(iIRCServer), "Settings", "ServerName", "")
txtDescription.Text = ReadINI(GetINIFile(iIRCServer), "Settings", "Description", "")
txtDefaultTopic.Text = ReadINI(GetINIFile(iIRCServer), "Channel Defaults", "Topic", "")
txtDefaultQuit.Text = ReadINI(GetINIFile(iIRCServer), "Default User Settings", "Default Quit Msg", "")
msg3 = ReadINI(GetINIFile(iIRCServer), "Settings", "Port", "")
cboPort.ListIndex = FindComboBoxIndex(cboPort, msg3)
msg3 = ReadINI(GetINIFile(iIRCServer), "Settings", "MaxNickLength", "")
cboMaxNicklen.ListIndex = FindComboBoxIndex(cboMaxNicklen, msg3)
msg3 = ReadINI(GetINIFile(iIRCServer), "Settings", "Session Limit", "")
cboSessions.ListIndex = FindComboBoxIndex(cboSessions, msg3)
msg3 = ReadINI(GetINIFile(iIRCServer), "Settings", "AwayLen", "")
cboAwayLen.ListIndex = FindComboBoxIndex(cboAwayLen, msg3)
msg3 = ReadINI(GetINIFile(iIRCServer), "Settings", "KickLen", "")
cboKickLength.ListIndex = FindComboBoxIndex(cboKickLength, msg3)
Me.Icon = mdiNexIRC.Icon
p = Int(ReadINI(App.Path & "\data\config\profiles\0.ini", "Settings", "Count", 0))
For i = 1 To p
    msg2 = App.Path & "\data\config\profiles\" & Trim(Str(i)) & ".ini"
    If DoesFileExist(msg2) Then
        E = Int(ReadINI(msg2, "Settings", "Count", 0))
        If E <> 0 Then
            For c = 0 To E
                If ReadINI(msg2, Trim(Str(c)), "Type", "-1") = "0" Then
                    msg = ReadINI(msg2, Trim(Str(c)), "Data", "")
                    If Len(msg) <> 0 Then cboProfile.AddItem msg
                End If
            Next c
        End If
    End If
Next i
cboProfile.ListIndex = 0
cboBotType.ListIndex = 1
'SetButtonType cmdMoreThemes
SetButtonType cmdPreviewTheme
SetButtonType cmdAddAlternate
SetButtonType cmdCancel
SetButtonType cmdBack
SetButtonType cmdNext
SetButtonType cmdFinish
SetButtonType cmdAddAutojoin
SetButtonType cmdRemoveAutoJoin
SetButtonType cmdAddBot
SetButtonType cmdAddIgnore
SetButtonType cmdAddNotify
SetButtonType cmdRemoveNotify
SetButtonType cmdEditBot
SetButtonType cmdRemoveBot
SetButtonType cmdRemoveAlternate
lFrameCount = 9
For i = 0 To lServers.sNetworkCount
    If Len(lServers.sNetwork(i).nDescription) <> 0 Then cboAutojoinNetwork.AddItem lServers.sNetwork(i).nDescription
Next i
FillListBoxWithAlternates lstAlternateNicknames
FillListBoxWithNotify lstNotifyList
FillListBoxWithIgnore lstIgnore
cboAutojoinNetwork.ListIndex = FindComboBoxIndex(cboAutojoinNetwork, "Undernet")
SetCheckBoxValue chkConnectOnStartup, lSettings.sConnectOnStartup
SetCheckBoxValue chkByPassSplashScreen, lSettings.sByPassStartupScreen
SetCheckBoxValue chkShowHomepageOnStartup, lSettings.sNavigateOnStartup
SetCheckBoxValue chkShowTips, lSettings.sShowTips
SetCheckBoxValue chkOptionsOnStartup, lSettings.sShowOptionsOnStartup
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sSetupWizardVisible = False
lFrame = 0
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub Label33_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'frmRegister.Show 1
End Sub

Private Sub Label34_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmScriptManager.Show 1
End Sub

Private Sub lblAddAudioToYourPlaylist_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PromptAddToPlaylist
End Sub

Private Sub lblEditAutoPerform_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmAutoPerform.Show 1
End Sub

Private Sub lblMenuEditor_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmMenuEditor.Show 1
End Sub

Private Sub lblViewHelp_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmAutoConnect.Show 1
End Sub

Private Sub optAutoJoinEnabled_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case Index
Case 0
    cboAutojoinNetwork.Enabled = False
    cboAutojoinNetwork.BackColor = &H8000000F
    lstAutojoin.BackColor = &H8000000F
    lstAutojoin.Enabled = False
    cmdAddAutojoin.Enabled = False
    cmdRemoveAutoJoin.Enabled = False
Case 1
    cboAutojoinNetwork.BackColor = vbWhite
    lstAutojoin.BackColor = vbWhite
    cboAutojoinNetwork.Enabled = True
    lstAutojoin.Enabled = True
    cmdAddAutojoin.Enabled = True
    cmdRemoveAutoJoin.Enabled = True
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub optAutoJoinEnabled_Click(Index As Integer)"
End Sub

Private Sub optEditAlternate_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case Index
Case 0
    lstAlternateNicknames.Enabled = False
    cmdAddAlternate.Enabled = False
    cmdRemoveAlternate.Enabled = False
    lstAlternateNicknames.BackColor = &H8000000F
Case 1
    lstAlternateNicknames.Enabled = True
    cmdAddAlternate.Enabled = True
    cmdRemoveAlternate.Enabled = True
    lstAlternateNicknames.BackColor = vbWhite
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub optEditAlternate_Click(Index As Integer)"
End Sub

Private Sub optEditBots_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case Index
Case 0
    cboBotType.Enabled = False
    cboBotType.BackColor = -2147483633
    lstBotlist.Enabled = False
    lstBotlist.BackColor = -2147483633
    cmdAddBot.Enabled = False
    cmdRemoveBot.Enabled = False
    cmdEditBot.Enabled = False
Case 1
    cboBotType.Enabled = True
    cboBotType.BackColor = vbWhite
    lstBotlist.Enabled = True
    lstBotlist.BackColor = vbWhite
    cmdAddBot.Enabled = True
    cmdRemoveBot.Enabled = True
    cmdEditBot.Enabled = True
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub optEditBots_Click(Index As Integer)"
End Sub

Private Sub optEditServer_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case Index
Case 0
    cboPort.Enabled = False
    cboPort.BackColor = &H8000000F
    cboMaxNicklen.Enabled = False
    cboMaxNicklen.BackColor = &H8000000F
    cboSessions.Enabled = False
    cboSessions.BackColor = &H8000000F
    cboAwayLen.Enabled = False
    cboAwayLen.BackColor = &H8000000F
    cboKickLength.Enabled = False
    cboKickLength.BackColor = &H8000000F
    txtServerName.Enabled = False
    txtServerName.BackColor = &H8000000F
    txtDescription.Enabled = False
    txtDescription.BackColor = &H8000000F
    txtDefaultQuit.Enabled = False
    txtDefaultQuit.BackColor = &H8000000F
    txtDefaultTopic.Enabled = False
    txtDefaultTopic.BackColor = &H8000000F
    chkShowServerOnStartup.Enabled = False
Case 1
    cboPort.Enabled = True
    cboPort.BackColor = vbWhite
    cboMaxNicklen.Enabled = True
    cboMaxNicklen.BackColor = vbWhite
    cboSessions.Enabled = True
    cboSessions.BackColor = vbWhite
    cboAwayLen.Enabled = True
    cboAwayLen.BackColor = vbWhite
    cboKickLength.Enabled = True
    cboKickLength.BackColor = vbWhite
    txtServerName.Enabled = True
    txtServerName.BackColor = vbWhite
    txtDescription.Enabled = True
    txtDescription.BackColor = vbWhite
    txtDefaultQuit.Enabled = True
    txtDefaultQuit.BackColor = vbWhite
    txtDefaultTopic.Enabled = True
    txtDefaultTopic.BackColor = vbWhite
    chkShowServerOnStartup.Enabled = True
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub optEditServer_Click(Index As Integer)"
End Sub

Private Sub optEverything_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
chkConnectOnStartup.Value = 1
chkByPassSplashScreen.Value = 1
chkShowHomepageOnStartup.Value = 1
chkShowTips.Value = 1
chkOptionsOnStartup.Value = 1
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub optStartupProfile_Click()"
End Sub

Private Sub optIgnoreEnabled_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case Index
Case 0
    lstIgnore.Enabled = False
    cmdAddIgnore.Enabled = False
    cmdRemoveIgnore.Enabled = False
    lstIgnore.BackColor = -2147483633
Case 1
    lstIgnore.Enabled = True
    cmdAddIgnore.Enabled = True
    cmdRemoveIgnore.Enabled = True
    lstIgnore.BackColor = vbWhite
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Option1_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
chkConnectOnStartup.Value = 0
chkByPassSplashScreen.Value = 1
chkShowHomepageOnStartup.Value = 1
chkShowTips.Value = 0
chkOptionsOnStartup.Value = 1
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub optStartupProfile_Click()"
End Sub

Private Sub optNothing_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
chkConnectOnStartup.Value = 0
chkByPassSplashScreen.Value = 0
chkShowHomepageOnStartup.Value = 0
chkShowTips.Value = 0
chkOptionsOnStartup.Value = 0
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub optStartupProfile_Click()"
End Sub

Private Sub optNotifyEnabled_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case Index
Case 0
    lstNotifyList.BackColor = -2147483633
    lstNotifyList.Enabled = False
    cmdAddNotify.Enabled = False
    cmdRemoveNotify.Enabled = False
Case 1
    lstNotifyList.BackColor = vbWhite
    lstNotifyList.Enabled = True
    cmdAddNotify.Enabled = True
    cmdRemoveNotify.Enabled = True
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub optNotifyEnabled_Click(Index As Integer)"
End Sub

Private Sub optProfile_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case Index
Case 1
    txtProfileName.BackColor = &H8000000F
    txtProfileName.Enabled = False
Case 0
    txtProfileName.BackColor = vbWhite
    txtProfileName.Enabled = True
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub optProfile_Click(Index As Integer)"
End Sub

Private Sub optSetupWizardContinue_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case Index
Case 0
    cboProfile.BackColor = &H8000000F
    cboProfile.Enabled = False
    chkFinishAfterProfile.Enabled = False
Case 1
    cboProfile.BackColor = &H8000000F
    cboProfile.Enabled = False
    chkFinishAfterProfile.Enabled = False
Case 2
    cboProfile.BackColor = vbWhite
    cboProfile.Enabled = True
    chkFinishAfterProfile.Enabled = True
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub optSetupWizardContinue_Click(Index As Integer)"
End Sub

Private Sub optStartupProfile_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
chkConnectOnStartup.Value = 0
chkByPassSplashScreen.Value = 1
chkShowHomepageOnStartup.Value = 1
chkShowTips.Value = 1
chkOptionsOnStartup.Value = 1
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub optStartupProfile_Click()"
End Sub

Private Sub optTheme_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case Index
Case 0
    chkThemeToolbar.Value = 0
    chkShowMixer.Value = 1
    chkShowQuickNotify.Value = 0
    chkAutosizeTaskbar.Value = 0
    chkTimeStamping.Value = 1
    chkColoredNicklist.Value = 1
Case 1
    chkThemeToolbar.Value = 1
    chkShowMixer.Value = 1
    chkShowQuickNotify.Value = 1
    chkAutosizeTaskbar.Value = 1
    chkTimeStamping.Value = 1
    chkColoredNicklist.Value = 1
Case 2
    chkThemeToolbar.Value = 0
    chkShowMixer.Value = 0
    chkShowQuickNotify.Value = 0
    chkAutosizeTaskbar.Value = 0
    chkTimeStamping.Value = 0
    chkColoredNicklist.Value = 0
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub optTheme_Click(Index As Integer)"
End Sub
