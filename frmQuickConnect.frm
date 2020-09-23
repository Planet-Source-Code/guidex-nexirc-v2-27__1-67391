VERSION 5.00
Begin VB.Form frmQuickConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Quick Connect"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmQuickConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4815
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkNewStatus 
      Appearance      =   0  'Flat
      Caption         =   "&New Status Window"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   960
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      Text            =   "6667"
      Top             =   840
      Width           =   3735
   End
   Begin VB.ComboBox cmbNetwork 
      Height          =   315
      Left            =   960
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Text            =   "cmbNetwork"
      Top             =   120
      Width           =   3735
   End
   Begin VB.ComboBox cmbServer 
      Height          =   315
      Left            =   960
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Text            =   "cmbServer"
      Top             =   480
      Width           =   3735
   End
   Begin nexIRC.ctlXPButton cmdOK 
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "OK"
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
      MICON           =   "frmQuickConnect.frx":000C
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
      Left            =   3720
      TabIndex        =   8
      Top             =   1320
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
      MICON           =   "frmQuickConnect.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdHelp 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1320
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
      MICON           =   "frmQuickConnect.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblPort 
      Caption         =   "&Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblServer 
      Caption         =   "&Server:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblNetwork 
      Caption         =   "&Network:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
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
Attribute VB_Name = "frmQuickConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbNetwork_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, j As Integer, mItem As ListItem, word() As String
cmbServer.Clear
j = FindNetworkIndex(cmbNetwork.Text)
If j <> 0 Then
    For i = 1 To lServers.sServerCount
        If lServers.sServer(i).sNetwork = j Then
            cmbServer.AddItem lServers.sServer(i).sServer
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmbNetwork_Click()"
End Sub

Private Sub cmdCancel_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdCancel_Click()"
End Sub

Private Sub cmdHelp_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DisplayHelpInformation 28
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdHelp_Click()"
End Sub

Private Sub cmdOK_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lServer As String, lPort As String, mbox As VbMsgBoxResult
lServer = cmbServer.Text
lPort = txtPort.Text
If Len(lServer) <> 0 And Len(lPort) <> 0 Then
    If chkNewStatus.Value = 1 Then
        NewStatusWindow lServer, lPort, True
    Else
        ConnectToIRC lServer, lPort, lSettings.sActiveServerForm
    End If
ElseIf Len(lServer) = 0 Then
    If lSettings.sGeneralPrompts = True Then
        mbox = MsgBox("You did not enter a server, would you like to enter one now?", vbYesNoCancel + vbExclamation)
    Else
        mbox = vbYes
    End If
    If mbox = vbYes Then
        cmbServer.SetFocus
        Exit Sub
    ElseIf mbox = vbNo Then
        Unload Me
        Exit Sub
    ElseIf mbox = vbCancel Then
        Exit Sub
    End If
End If
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdOK_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = mdiNexIRC.Icon
SetButtonType cmdCancel
SetButtonType cmdHelp
SetButtonType cmdOK
Dim i As Integer, F As Integer
cmbServer.Text = lSettings.sServer
For i = 0 To lServers.sNetworkCount
    If Len(lServers.sNetwork(i).nDescription) <> 0 Then cmbNetwork.AddItem lServers.sNetwork(i).nDescription
Next i
If Len(lSettings.sNetwork) <> 0 Then cmbNetwork.ListIndex = FindComboBoxIndex(cmbNetwork, lSettings.sNetwork)
cmbServer.Text = lSettings.sServer
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
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
