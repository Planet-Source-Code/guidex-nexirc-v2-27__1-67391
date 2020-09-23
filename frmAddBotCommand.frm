VERSION 5.00
Begin VB.Form frmAddBotCommand 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Add Bot/Command"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   510
   ClientWidth     =   3960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddBotCommand.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3960
   StartUpPosition =   1  'CenterOwner
   Begin nexIRC.ctlXPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   16
      ToolTipText     =   "Cancel/Hide this Window"
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "  Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmAddBotCommand.frx":000C
      PICN            =   "frmAddBotCommand.frx":016E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdAdd 
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      ToolTipText     =   "Add Bot to List"
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "    &Add"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmAddBotCommand.frx":070A
      PICN            =   "frmAddBotCommand.frx":086C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdHelp 
      Height          =   375
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Show Help Topics"
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "     &Help"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmAddBotCommand.frx":0E06
      PICN            =   "frmAddBotCommand.frx":0F68
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cboAddType 
      Height          =   315
      ItemData        =   "frmAddBotCommand.frx":1504
      Left            =   1200
      List            =   "frmAddBotCommand.frx":150E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select Add Bot to Add a Bot to your Bot list, select Add Bot Command to setup a command with the bot"
      Top             =   120
      Width           =   2655
   End
   Begin VB.OptionButton optAddOption 
      Appearance      =   0  'Flat
      Caption         =   "Add Bot Command"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   13
      Top             =   720
      Width           =   1815
   End
   Begin VB.OptionButton optAddOption 
      Appearance      =   0  'Flat
      Caption         =   "Add Bot"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   12
      Top             =   480
      Width           =   1815
   End
   Begin VB.Frame fraOption 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   3855
      Begin VB.ComboBox cboCommandType 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtCommand 
         Height          =   315
         Left            =   1080
         TabIndex        =   9
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "&Bot Type:"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblCommand 
         Caption         =   "&Command:"
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame fraOption 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3855
      Begin VB.ComboBox cboNicknameType 
         Height          =   315
         ItemData        =   "frmAddBotCommand.frx":152C
         Left            =   1080
         List            =   "frmAddBotCommand.frx":152E
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Select the Bot type"
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtNickname 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         ToolTipText     =   "Type the Nickname of the Bot"
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label lblType 
         Caption         =   "&Type:"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         ToolTipText     =   "Select the Bot type"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblNickname 
         Caption         =   "&Nickname:"
         Height          =   255
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "Type the Nickname of the Bot"
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Label lblAdd 
      Caption         =   "&Add:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Select Add Bot to Add a Bot to your Bot list, select Add Bot Command to setup a command with the bot"
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAddBot 
         Caption         =   "&Add Bot"
      End
      Begin VB.Menu mnuAddCommand 
         Caption         =   "A&dd Command"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuBotType1 
      Caption         =   "&Type"
      Begin VB.Menu mnuBotType 
         Caption         =   "&Unknown"
         Index           =   0
      End
      Begin VB.Menu mnuBotType 
         Caption         =   "&Eggdrop"
         Index           =   1
      End
      Begin VB.Menu mnuBotType 
         Caption         =   "&Undernet X"
         Index           =   2
      End
      Begin VB.Menu mnuBotType 
         Caption         =   "&ChanServ"
         Index           =   3
      End
      Begin VB.Menu mnuBotType 
         Caption         =   "&MemoServ"
         Index           =   4
      End
   End
   Begin VB.Menu mnuHelp1 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "H&ow to use this Window"
      End
   End
End
Attribute VB_Name = "frmAddBotCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboAddType_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case cboAddType.ListIndex
Case 0
    fraOption(0).Visible = True
    fraOption(1).Visible = False
Case 1
    fraOption(0).Visible = False
    fraOption(1).Visible = True
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cboAddType_Click()"
End Sub

Private Sub cmdAdd_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
Select Case cboAddType.ListIndex
Case 0
    AddBot txtNickname.Text, cboNicknameType.ListIndex
    GoTo ErrCHK
Case 1
    i = AddBotCommand(txtCommand.Text, cboCommandType.ListIndex)
    GoTo ErrCHK
End Select
ErrCHK:
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdOK_Click()"
Unload Me
End Sub

Private Sub cmdCancel_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdCancel_Click()"
End Sub

Private Sub cmdHelp_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DisplayHelpInformation 0
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdHelp_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = mdiNexIRC.Icon
cboAddType.ListIndex = 0
SetButtonType cmdHelp
SetButtonType cmdAdd
SetButtonType cmdCancel
cboNicknameType.AddItem "0 - Unknown/Custom Bot"
cboNicknameType.AddItem "1 - Eggdrop"
cboNicknameType.AddItem "2 - Undernet X"
cboNicknameType.AddItem "3 - ChanServ"
cboNicknameType.AddItem "4 - MemoServ"
cboCommandType.AddItem "0 - Unknown/Custom Bot"
cboCommandType.AddItem "1 - Eggdrop"
cboCommandType.AddItem "2 - Undernet X"
cboCommandType.AddItem "3 - ChanServ"
cboCommandType.AddItem "4 - MemoServ"
End Sub

Private Sub mnuAddBot_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cboAddType.ListIndex = 0
fraOption(0).Visible = True
fraOption(1).Visible = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAddBot_Click()"
End Sub

Private Sub mnuAddCommand_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cboAddType.ListIndex = 1
fraOption(0).Visible = False
fraOption(1).Visible = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAddCommand_Click()"
End Sub

Private Sub mnuBotType_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case cboAddType.ListIndex
Case 0
    cboNicknameType.ListIndex = Index
Case 1
    cboCommandType.ListIndex = Index
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuBotType_Click(Index As Integer)"
End Sub

Private Sub mnuExit_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuExit_Click()"
End Sub

Private Sub mnuHelp_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DisplayHelpInformation 0
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuHelp_Click()"
End Sub

Private Sub optAddOption_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
fraOption(0).Visible = False
fraOption(1).Visible = False
fraOption(Index).Visible = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub optAddOption_Click(Index As Integer)"
End Sub
