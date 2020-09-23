VERSION 5.00
Begin VB.Form frmSecureQuery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Secure Query"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   3585
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSecureQuery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   3585
   StartUpPosition =   1  'CenterOwner
   Begin nexIRC.ctlXPButton cmdAccept 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Accept"
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
      MICON           =   "frmSecureQuery.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer tmrAutoClose 
      Interval        =   1000
      Left            =   120
      Top             =   720
   End
   Begin VB.CheckBox chkAddToNotify 
      Appearance      =   0  'Flat
      Caption         =   "Add to &Notify"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CheckBox chkAddToIgnore 
      Appearance      =   0  'Flat
      Caption         =   "Add to &Ignore"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin nexIRC.ctlXPButton cmdDecline 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Decline"
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
      MICON           =   "frmSecureQuery.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblAutoclose 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Image imgNexgen 
      Height          =   480
      Left            =   120
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblNickname 
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblAccept 
      Caption         =   "Accept query from "
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuAddToIgnore 
         Caption         =   "Add to &Ignore"
      End
      Begin VB.Menu mnuAddToNotify 
         Caption         =   "Add to &Notify"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowToUseThisWindow 
         Caption         =   "H&ow to use this Window"
      End
   End
End
Attribute VB_Name = "frmSecureQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lClosingIn As Integer

Private Sub chkAddToIgnore_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case chkAddToIgnore.Value
Case 0
    mnuAddToIgnore.Checked = False
Case 1
    mnuAddToIgnore.Checked = True
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkAddToIgnore_Click()"
End Sub

Private Sub cmdAccept_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSecureQuery.sAccepted = True
lSecureQuery.sAddToIgnore = GetCheckboxValue(chkAddToIgnore)
lSecureQuery.sAddToNotify = GetCheckboxValue(chkAddToNotify)
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdDecline_Click()"
End Sub

Private Sub cmdDecline_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSecureQuery.sAddToIgnore = GetCheckboxValue(chkAddToIgnore)
lSecureQuery.sAddToNotify = GetCheckboxValue(chkAddToNotify)
lSecureQuery.sAccepted = False
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdDecline_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = mdiNexIRC.Icon
imgNexgen.Picture = LoadPicture(App.Path & "\data\icons\nexgen.ico")
lClosingIn = 20
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub mnuAddToIgnore_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case chkAddToIgnore.Value
Case 0
    chkAddToIgnore.Value = 1
    mnuAddToIgnore.Checked = True
Case 1
    chkAddToIgnore.Value = 0
    mnuAddToIgnore.Checked = False
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAddToIgnore_Click()"
End Sub

Private Sub mnuAddToNotify_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case chkAddToNotify.Value
Case 0
    chkAddToNotify.Value = 1
    mnuAddToNotify.Checked = True
Case 1
    chkAddToNotify.Value = 0
    mnuAddToNotify.Checked = False
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAddToNotify_Click()"
End Sub

Private Sub mnuExit_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuExit_Click()"
End Sub

Private Sub mnuHowToUseThisWindow_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
MsgBox "No help for this topic is available", vbExclamation
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuHowToUseThisWindow_Click()"
End Sub

Private Sub tmrAutoClose_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lClosingIn = lClosingIn - 1
lblAutoclose.Caption = "Timeout: " & lClosingIn
If lClosingIn = 0 Then
    lSecureQuery.sAccepted = False
    lSecureQuery.sAddToIgnore = False
    lSecureQuery.sAddToNotify = False
    Unload Me
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrAutoClose_Timer()"
End Sub
