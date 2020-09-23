VERSION 5.00
Begin VB.Form frmAddServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Add Server"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4800
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   1680
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1680
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1680
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   480
      Width           =   3015
   End
   Begin nexIRC.ctlXPButton cmdHelp 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1200
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
      MICON           =   "frmAddServer.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "frmAddServer.frx":0028
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
      Left            =   3600
      TabIndex        =   6
      Top             =   1200
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
      MICON           =   "frmAddServer.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblPortRange 
      AutoSize        =   -1  'True
      Caption         =   "Port Range:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   870
   End
   Begin VB.Label lblServerName 
      AutoSize        =   -1  'True
      Caption         =   "Server Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   990
   End
   Begin VB.Label lblServerDescription 
      AutoSize        =   -1  'True
      Caption         =   "&Server Description:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1380
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
Attribute VB_Name = "frmAddServer"
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
DisplayHelpInformation 4
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdHelp_Click()"
End Sub

Private Sub cmdOK_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim word() As String, t As Integer, i As Integer, l As Integer, bool As Boolean, mItem As ListItem, msg As String
If Len(frmCustomize.cmbNetwork.Text) <> 0 And Len(txtName.Text) <> 0 And Len(txtPort.Text) <> 0 Then
    Set mItem = frmCustomize.lvwServers.ListItems.Add(, , txtDescription.Text)
    mItem.SubItems(1) = txtName.Text
    mItem.SubItems(2) = txtPort.Text
    For t = 0 To 1000
        msg = ReadINI(GetINIFile(iServers), frmCustomize.cmbNetwork.Text, Str(t), "")
        If Len(msg) <> 0 Then
            l = l + 1
        Else
            Exit For
        End If
    Next t
    WriteINI GetINIFile(iServers), frmCustomize.cmbNetwork.Text, Str(l), txtName.Text & "|" & txtDescription.Text & "|" & txtPort.Text
    lServers.sServerUBound = lServers.sServerUBound + 1
    lServers.sServerCount = lServers.sServerCount + 1
    lServers.sServer(lServers.sServerCount).sDescription = txtDescription.Text
    lServers.sServer(lServers.sServerCount).sNetwork = FindNetworkIndex(frmCustomize.cmbNetwork.Text)
    lServers.sServer(lServers.sServerCount).sPortRange = txtPort.Text
    lServers.sServer(lServers.sServerCount).sServer = txtName.Text
    'If chkConnect.Value = 1 Then
    '    ConnectToIRC txtName.Text, txtPort.Text, lSettings.sActiveServerForm
    'End If
Else
    If Len(txtDescription.Text) <> 0 Then
        Beep
        If lSettings.sGeneralPrompts = True Then MsgBox "Unable to add a server without a description.", vbExclamation
        txtDescription.SetFocus
        Exit Sub
    End If
    If Len(txtName.Text) <> 0 Then
        Beep
        If lSettings.sGeneralPrompts = True Then MsgBox "Unable to add a server without a name.", vbExclamation
        txtName.SetFocus
        Exit Sub
    End If
    If Len(txtPort.Text) <> 0 Then
        Beep
        If lSettings.sGeneralPrompts = True Then MsgBox "Unable to add a server without a port.", vbExclamation
        txtPort.SetFocus
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
SetButtonType cmdOK
SetButtonType cmdHelp
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtDescription_GotFocus()"
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

Private Sub txtDescription_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtDescription.SelStart = 0
txtDescription.SelLength = Len(txtDescription.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtDescription_GotFocus()"
End Sub

Private Sub txtName_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtName.SelStart = 0
txtName.SelLength = Len(txtName.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtName_GotFocus()"
End Sub

Private Sub txtport_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtPort.SelStart = 0
txtPort.SelLength = Len(txtPort.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtPort_GotFocus()"
End Sub
