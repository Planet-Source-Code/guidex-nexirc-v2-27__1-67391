VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmAccept 
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
   Icon            =   "frmAccept.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   3330
   StartUpPosition =   1  'CenterOwner
   Begin OsenXPCntrl.OsenXPButton cmdHelp 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&Help"
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
      MICON           =   "frmAccept.frx":000C
      PICN            =   "frmAccept.frx":016E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&Cancel"
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
      MICON           =   "frmAccept.frx":070A
      PICN            =   "frmAccept.frx":086C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdChat 
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "C&hat"
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
      MICON           =   "frmAccept.frx":0E08
      PICN            =   "frmAccept.frx":0F6A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   -1  'True
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
Attribute VB_Name = "frmAccept"
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
For i = 1 To maxtcp
    If mdiMain.CHAT(i).State = sckError Or mdiMain.CHAT(i).State = sckClosed Then
        Load ChatWindow(i)
        ChatWindow(i).Show
        mdiMain.CHAT(i).Close
        ChatWindow(i).Caption = lblNickname
        ChatWindowName(i) = lblNickname
        mdiMain.CHAT(i).Connect lblIP.Caption, lblPort.Caption
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
Me.Icon = mdiMain.Icon
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
