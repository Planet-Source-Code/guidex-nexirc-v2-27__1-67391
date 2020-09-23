VERSION 5.00
Begin VB.Form frmNoNicknameGiven 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - No Nickname Given"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNoNicknameGiven.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNickname 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin NexIRC.ctlXPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      ToolTipText     =   "Cancel/Hide this Window"
      Top             =   720
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
      MICON           =   "frmNoNicknameGiven.frx":000C
      PICN            =   "frmNoNicknameGiven.frx":016E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin NexIRC.ctlXPButton cmdChangeNick 
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "OK"
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
      MICON           =   "frmNoNicknameGiven.frx":070A
      PICN            =   "frmNoNicknameGiven.frx":086C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin NexIRC.ctlXPButton cmdHelp 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Show Help Topics"
      Top             =   720
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
      MICON           =   "frmNoNicknameGiven.frx":0E06
      PICN            =   "frmNoNicknameGiven.frx":0F68
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   4560
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblNickname 
      Caption         =   "Nickname:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmNoNicknameGiven"
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

Private Sub cmdChangeNick_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.SendData "NICK " & txtNickname.Text & vbCrLf
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdChangeNick_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = mdiNexIRC.Icon
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub
