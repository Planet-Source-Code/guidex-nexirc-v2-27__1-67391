VERSION 5.00
Begin VB.Form frmEditServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Edit Server"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4830
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboNetwork 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1560
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1560
      MousePointer    =   1  'Arrow
      TabIndex        =   7
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   1560
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   480
      Width           =   3135
   End
   Begin nexIRC.ctlXPButton cmdHelp 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1560
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
      MICON           =   "frmEditServer.frx":000C
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
      TabIndex        =   9
      Top             =   1560
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
      MICON           =   "frmEditServer.frx":0028
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
      TabIndex        =   10
      Top             =   1560
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
      MICON           =   "frmEditServer.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblNetwork 
      Caption         =   "&Network:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblServerDescription 
      AutoSize        =   -1  'True
      Caption         =   "&Server Description:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1380
   End
   Begin VB.Label lblServerName 
      AutoSize        =   -1  'True
      Caption         =   "Server &Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Port Range:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   870
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOther 
      Caption         =   "&Other"
      Begin VB.Menu mnuAddServer 
         Caption         =   "&Add Server"
      End
      Begin VB.Menu mnuAddNetwork 
         Caption         =   "&Add Network"
      End
      Begin VB.Menu mnuServerList 
         Caption         =   "&Server List"
      End
      Begin VB.Menu mnuIRCServer 
         Caption         =   "&IRC Server"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowToUseThisWindow 
         Caption         =   "H&ow to use this Window"
      End
   End
End
Attribute VB_Name = "frmEditServer"
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
DisplayHelpInformation 19
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdHelp_Click()"
End Sub

Private Sub cmdOK_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer, msg2 As String, d As Integer, m As Integer, E As Integer, s As Integer, lItem As ListItem, lFound As Boolean
d = FindNetworkIndex(cboNetwork.Text)
If d <> 0 Then
    msg2 = frmCustomize.lvwServers.SelectedItem.Text
    s = FindServerIndex(frmCustomize.lvwServers.SelectedItem.SubItems(1))
    E = FindListViewIndex(frmCustomize.lvwServers, msg2)
    If E <> 0 Then
        frmCustomize.lvwServers.ListItems(E).Text = txtDescription.Text
        frmCustomize.lvwServers.ListItems(E).SubItems(1) = txtName.Text
        frmCustomize.lvwServers.ListItems(E).SubItems(2) = txtPort.Text
    End If
    For m = 0 To 150
        msg = ReadINI(GetINIFile(iServers), cboNetwork.Text, Trim(Str(m)), "")
        If InStr(LCase(msg), LCase(txtName.Text)) Then
            lFound = True
            WriteINI GetINIFile(iServers), cboNetwork.Text, Trim(Str(m)), txtName.Text & "|" & txtDescription.Text & "|" & txtPort.Text
            Exit For
        End If
    Next m
    If lFound = False Then
        For m = 0 To 150
            msg = ReadINI(GetINIFile(iServers), cboNetwork.Text, Trim(Str(m)), "")
            If Len(msg) = 0 Then
                WriteINI GetINIFile(iServers), cboNetwork.Text, Trim(Str(m)), txtName.Text & "|" & txtDescription.Text & "|" & txtPort.Text
            End If
        Next m
    End If
    If s <> 0 Then
        lServers.sServer(s).sDescription = txtDescription.Text
        lServers.sServer(s).sNetwork = FindNetworkIndex(cboNetwork.Text)
        lServers.sServer(s).sPassword = ""
        lServers.sServer(s).sPortRange = txtPort.Text
        lServers.sServer(s).sServer = txtName.Text
    End If
End If
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdOK_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To 1000
    If Len(lServers.sNetwork(i).nDescription) <> 0 Then
        cboNetwork.AddItem lServers.sNetwork(i).nDescription
    End If
Next i
cboNetwork.ListIndex = FindComboBoxIndex(cboNetwork, frmCustomize.cmbNetwork.Text)
SetButtonType cmdHelp
SetButtonType cmdOK
SetButtonType cmdCancel
txtDescription = frmCustomize.lvwServers.SelectedItem.Text
txtName = frmCustomize.lvwServers.SelectedItem.SubItems(1)
txtPort = frmCustomize.lvwServers.SelectedItem.SubItems(2)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub mnuAddNetwork_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmCustomize.Show
frmAddNetwork.Show 0, frmCustomize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAddServer_Click()"
End Sub

Private Sub mnuAddServer_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmCustomize.Show
frmAddServer.Show 0, frmCustomize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuAddServer_Click()"
End Sub

Private Sub mnuIRCServer_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmIRCServer.Show
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuIRCServer_Click()"
End Sub

Private Sub mnuServerList_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmCustomize.Show
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuServerList_Click()"
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
