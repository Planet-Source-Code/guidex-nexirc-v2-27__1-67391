VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSendFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - DCC Send"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSendFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4305
   StartUpPosition =   1  'CenterOwner
   Begin nexIRC.ctlXPButton cmdHelp 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1920
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
      MICON           =   "frmSendFile.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer tmrSendFile 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   1920
   End
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtNickname 
      Height          =   285
      Left            =   1080
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "..."
      Height          =   285
      Left            =   3720
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   1080
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin MSWinsockLib.Winsock tcpSend 
      Index           =   0
      Left            =   1200
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin nexIRC.ctlXPButton cmdSend 
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Send"
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
      MICON           =   "frmSendFile.frx":0028
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
      Left            =   3240
      TabIndex        =   14
      Top             =   1920
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
      MICON           =   "frmSendFile.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblBytesSent 
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label lblSent 
      Caption         =   "Sent:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblPercent 
      Caption         =   "0%"
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label lblPercentDesc 
      Caption         =   "Percent:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblSendTo 
      AutoSize        =   -1  'True
      Caption         =   "&Send to:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "File:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   300
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Size:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   345
   End
   Begin VB.Label lblFileSize 
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   840
      Width           =   3135
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
Attribute VB_Name = "frmSendFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFullPath As String, i, fLength, ret, buffer As String, bSize As Long, ByteSent As Long

Public Sub SetStrFullPath(lFullPath As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
strFullPath = lFullPath
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetStrFullPath(lFullPath As String)"
End Sub

Public Sub ClickSendButton()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdSend_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ClickSendButton()"
End Sub

Function SendFile() As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
i = (Val(Me.Tag) - 1530)
If ByteSent < Progress.Max Then
    lblPercent.Caption = Int((ByteSent * 100) / Progress.Max) & "%"
    Me.Caption = "DCC File Transfer: " & lblPercent.Caption
    Progress.Value = ByteSent: DoEvents
Else
    Progress.Value = Progress.Max: DoEvents
End If
bSize = 1024
If Progress.Value = Progress.Max Then GoTo EndOfTrans
fLength = LOF(i)
If ByteSent >= fLength Then
    tcpSend(Me.Tag).SendData ""
    Exit Function
End If
If ByteSent + bSize > fLength Then
    bSize = fLength - ByteSent
End If
buffer = Space$(bSize)
Get i, , buffer
tcpSend(Me.Tag).SendData buffer
ByteSent = ByteSent + bSize
lblBytesSent.Caption = Format(ByteSent, "###,###,###")
EndOfTrans:
If ByteSent >= Val(FileLen(txtFileName.Text)) Then
    Close #i
    cmdCancel.Caption = "Close"
    Me.Caption = "NexIRC - File Sent"
    lblPercent.Caption = ""
    Progress.Value = 0
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Function SendFile() As Boolean"
End Function

Private Sub cmdCancel_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdCancel_Click()"
End Sub

Private Sub cmdFile_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = OpenDialog(Me, "All Files (*.*)|*.*|Mp3 Files (*.mp3)|*.mp3|Zip Files (*.zip)|*.zip|Wave Files (*.wav)|*.wav|Rar Files (*.rar)|*.rar|Jpeg Files (*.jpg)|*.jpg|Gif Files (*.gif)|*.gif|", "DCC Send", CurDir)
If Len(msg) <> 0 Then
    txtFileName.Text = msg
    lblFileSize.Caption = Format(FileLen(txtFileName.Text), "###,###,###,###") & " KB"
    strFullPath = msg
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdFile_Click()"
End Sub

Private Sub cmdHelp_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DisplayHelpInformation 32
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdHelp_Click()"
End Sub

Private Sub cmdSend_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lIP As String, TempFileName As String
cmdSend.Enabled = False
TempFileName = Replace(txtFileName, " ", "_")
lIP = IrcGetLongIP(lSettings.sActiveServerForm.tcp.LocalIP)
txtFileName.Enabled = False
txtNickname.Enabled = False
If Len(lIP) <> 0 Then
    Caption = "NexIRC - Sending"
    If lSettings.sActiveServerForm.tcp.State = sckConnected Then
        lSettings.sActiveServerForm.tcp.SendData "NOTICE " & txtNickname.Text & " :DCC SEND " & txtFileName.Text & "(" & tcpSend(Me.Tag).LocalIP & ")" & vbCrLf
        lSettings.sActiveServerForm.tcp.SendData "PRIVMSG " & txtNickname.Text & " :DCC SEND " & TempFileName & " " & lIP & " " & tcpSend(Me.Tag).LocalPort & " " & " " & FileLen(txtFileName.Text) & "" & vbCrLf
        Progress.Min = 0
        Progress.Max = Val(FileLen(txtFileName.Text))
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdSend_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = mdiNexIRC.Icon
SetButtonType cmdCancel
SetButtonType cmdHelp
SetButtonType cmdSend
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

Private Sub tcpSend_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If tcpSend(Me.Tag).State <> sckClosed Then tcpSend(Me.Tag).Close
tcpSend(Me.Tag).Close
tcpSend(Me.Tag).Accept requestID
Me.Caption = "NexIRC - Connection Request"
i = (Val(Me.Tag) - 1530)
Open strFullPath For Binary Access Read As i
bSize = 1024
fLength = LOF(i)
If fLength - Loc(i) <= bSize Then
    bSize = fLength - Loc(i)
End If
If bSize = 0 Then Exit Sub
ByteSent = ByteSent + bSize
buffer = Space$(bSize)
Get i, , buffer
tcpSend(Me.Tag).SendData buffer
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tcpSend_ConnectionRequest(Index As Integer, ByVal requestID As Long)"
End Sub

Private Sub tcpSend_DataArrival(Index As Integer, ByVal bytesTotal As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim strData As String
tcpSend(Me.Tag).GetData strData
SendFile
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tcpSend_DataArrival(Index As Integer, ByVal bytesTotal As Long)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub tmrSendFile_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lblFileSize.Caption = Format(FileLen(txtFileName.Text), "###,###,###,###") & " KB"
strFullPath = txtFileName.Text
tmrSendFile.Enabled = False
ClickSendButton
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtFileName_GotFocus()"
End Sub

Private Sub txtFileName_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtFileName.SelStart = 0
txtFileName.SelLength = Len(txtFileName.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtFileName_GotFocus()"
End Sub

Private Sub txtNickname_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtNickname.SelStart = 0
txtNickname.SelLength = Len(txtNickname.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtNickname_GotFocus()"
End Sub
