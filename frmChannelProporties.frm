VERSION 5.00
Begin VB.Form frmChannelProporties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Channel Proporties"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChannelProporties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5520
   StartUpPosition =   1  'CenterOwner
   Begin nexIRC.ctlXPButton cmdAdd 
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   840
      Width           =   855
      _ExtentX        =   1508
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
      MICON           =   "frmChannelProporties.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstBans 
      Height          =   1230
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4335
   End
   Begin VB.CheckBox chkSecret 
      Appearance      =   0  'Flat
      Caption         =   "Secret"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   2880
      Width           =   855
   End
   Begin VB.CheckBox chkPrivate 
      Appearance      =   0  'Flat
      Caption         =   "Private"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtUserLimit 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   2400
      Width           =   975
   End
   Begin VB.CheckBox chkUserLimit 
      Appearance      =   0  'Flat
      Caption         =   "User Limit:"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtKey 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   2160
      Width           =   975
   End
   Begin VB.CheckBox chkKey 
      Appearance      =   0  'Flat
      Caption         =   "Key:"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   2160
      Width           =   735
   End
   Begin VB.CheckBox chkModerated 
      Appearance      =   0  'Flat
      Caption         =   "Moderated"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CheckBox chkInviteOnly 
      Appearance      =   0  'Flat
      Caption         =   "Invite Only"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CheckBox chkNoExternalMessages 
      Appearance      =   0  'Flat
      Caption         =   "No External Messages"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CheckBox chkOnlyOpsSetTopic 
      Appearance      =   0  'Flat
      Caption         =   "Topic Changes by Ops"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtTopic 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   5295
   End
   Begin nexIRC.ctlXPButton cmdDelete 
      Height          =   375
      Left            =   4560
      TabIndex        =   15
      Top             =   1320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Delete"
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
      MICON           =   "frmChannelProporties.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdEdit 
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   1800
      Width           =   855
      _ExtentX        =   1508
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
      MICON           =   "frmChannelProporties.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdAutoJoin 
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   2280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Auto Join"
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
      MICON           =   "frmChannelProporties.frx":0060
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
      Left            =   4560
      TabIndex        =   18
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
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
      MICON           =   "frmChannelProporties.frx":007C
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
      Height          =   375
      Left            =   3600
      TabIndex        =   19
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
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
      MICON           =   "frmChannelProporties.frx":0098
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
      TabIndex        =   20
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
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
      MICON           =   "frmChannelProporties.frx":00B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblBanList 
      Caption         =   "Banlist:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblTopic 
      Caption         =   "Topic:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmChannelProporties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lSelectedChannel As String

Private Sub chkInviteOnly_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case chkOnlyOpsSetTopic.Value
Case 0
    lSettings.sActiveServerForm.tcp.SendData "MODE " & lSelectedChannel & " -i" & vbCrLf
Case 1
    lSettings.sActiveServerForm.tcp.SendData "MODE " & lSelectedChannel & " +i" & vbCrLf
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkInviteOnly_Click()"
End Sub

Private Sub chkKey_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case chkKey.Value
Case 0
    lSettings.sActiveServerForm.tcp.SendData "MODE " & lSelectedChannel & " -k " & txtKey.Text & vbCrLf
Case 1
    If Len(txtKey.Text) <> 0 Then
        lSettings.sActiveServerForm.tcp.SendData "MODE " & lSelectedChannel & " +k " & txtKey.Text & vbCrLf
    End If
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkKey_Click()"
End Sub

Private Sub chkModerated_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case chkModerated.Value
Case 0
    lSettings.sActiveServerForm.tcp.SendData "MODE " & lSelectedChannel & " -m" & vbCrLf
Case 1
    lSettings.sActiveServerForm.tcp.SendData "MODE " & lSelectedChannel & " +m" & vbCrLf
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkModerated_Click()"
End Sub

Private Sub chkNoExternalMessages_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case chkNoExternalMessages.Value
Case 0
    lSettings.sActiveServerForm.tcp.SendData "MODE " & lSelectedChannel & " -n" & vbCrLf
Case 1
    lSettings.sActiveServerForm.tcp.SendData "MODE " & lSelectedChannel & " +n" & vbCrLf
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkOnlyOpsSetTopic_Click()"
End Sub

Private Sub chkOnlyOpsSetTopic_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case chkOnlyOpsSetTopic.Value
Case 0
    lSettings.sActiveServerForm.tcp.SendData "MODE " & lSelectedChannel & " -t" & vbCrLf
Case 1
    lSettings.sActiveServerForm.tcp.SendData "MODE " & lSelectedChannel & " +t" & vbCrLf
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkOnlyOpsSetTopic_Click()"
End Sub

Private Sub chkPrivate_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case chkPrivate.Value
Case 0
    lSettings.sActiveServerForm.tcp.SendData "MODE " & lSelectedChannel & " -p" & vbCrLf
Case 1
    lSettings.sActiveServerForm.tcp.SendData "MODE " & lSelectedChannel & " +p" & vbCrLf
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkPrivate_Click()"
End Sub

Private Sub chkSecret_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case chkSecret.Value
Case 0
    lSettings.sActiveServerForm.tcp.SendData "MODE " & lSelectedChannel & " -s" & vbCrLf
Case 1
    lSettings.sActiveServerForm.tcp.SendData "MODE " & lSelectedChannel & " +s" & vbCrLf
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkSecret_Click()"
End Sub

Private Sub chkUserLimit_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case chkUserLimit.Value
Case 0
    lSettings.sActiveServerForm.tcp.SendData "MODE " & lSelectedChannel & " -l " & txtUserLimit.Text & vbCrLf
Case 1
    If Len(txtUserLimit.Text) <> 0 Then
        lSettings.sActiveServerForm.tcp.SendData "MODE " & lSelectedChannel & " +l " & txtUserLimit.Text & vbCrLf
    End If
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkUserLimit_Click()"
End Sub

Private Sub cmdAutoJoin_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
AddAutoJoin Me.Tag, lSettings.sNetwork
frmAutoJoin.Show
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAutoJoin_Click()"
End Sub

Private Sub cmdCancel_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdCancel_Click()"
End Sub

Private Sub cmdOK_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer, F As Integer
msg = mdiNexIRC.ActiveForm.Tag
If Len(msg) <> 0 Then
    For i = 1 To ReturnChannelUBound
        If LCase(ReturnChannelName(i)) = LCase(msg) Then
            F = i
            Exit For
        End If
    Next i
    If LCase(Trim(ReturnChannelTopic(F)) <> LCase(Trim(txtTopic.Text))) Then
        lSettings.sActiveServerForm.tcp.SendData "TOPIC " & msg & " " & txtTopic.Text & vbCrLf
    End If
End If
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdOK_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSelectedChannel = mdiNexIRC.ActiveForm.Tag
Me.Icon = mdiNexIRC.Icon
SetButtonType cmdAdd
SetButtonType cmdAutoJoin
SetButtonType cmdCancel
SetButtonType cmdDelete
SetButtonType cmdEdit
SetButtonType cmdOK
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub
