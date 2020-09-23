VERSION 5.00
Begin VB.Form frmDCC_Accept 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - DCC Chat"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   3330
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDCC_Accept.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   3330
   StartUpPosition =   1  'CenterOwner
   Begin nexIRC.ctlXPButton cmdHelp 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "frmDCC_Accept.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdChat 
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Chat"
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
      MICON           =   "frmDCC_Accept.frx":0028
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
      Left            =   2280
      TabIndex        =   5
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "frmDCC_Accept.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblNickName 
      Caption         =   "NickName"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblIP 
      Caption         =   "IP"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblPort 
      Caption         =   "Port"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAcceptDCCChat 
         Caption         =   "&Accept DCC Chat"
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
Attribute VB_Name = "frmDCC_Accept"
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

Private Sub cmdChat_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To ReturnTCPUBound
    If mdiNexIRC.wskChat(i).State = sckError Or mdiNexIRC.wskChat(i).State = sckClosed Then
        Load lChatWindow(i)
        lChatWindow(i).Show
        mdiNexIRC.wskChat(i).Close
        lChatWindow(i).Caption = lblNickname
        lChatWindowName(i) = lblNickname
        mdiNexIRC.wskChat(i).Connect lblIP.Caption, lblPort.Caption
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdChat_Click()"
Unload Me
End Sub

Private Sub cmdHelp_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DisplayHelpInformation 15
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdHelp_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = mdiNexIRC.Icon
SetButtonType cmdHelp
SetButtonType cmdChat
SetButtonType cmdCancel
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub mnuAcceptDCCChat_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdChat_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAcceptDCCChat_Click()"
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
