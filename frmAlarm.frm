VERSION 5.00
Begin VB.Form frmAlarm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alarm"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   3720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAlarm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   3720
   StartUpPosition =   1  'CenterOwner
   Begin nexIRC.ctlXPButton cmdHelp 
      Height          =   375
      Left            =   120
      TabIndex        =   6
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
      MICON           =   "frmAlarm.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtAudio 
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtAlarmDate 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Timer tmrAlarm 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   2040
      Top             =   1200
   End
   Begin nexIRC.ctlXPButton cmdSelect 
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Select"
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
      MICON           =   "frmAlarm.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdClose 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Close"
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
      MICON           =   "frmAlarm.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdON 
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "ON"
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
      MICON           =   "frmAlarm.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdOFF 
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "OFF"
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
      MICON           =   "frmAlarm.frx":007C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "&Time:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "&Audio:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "&Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuToggle 
      Caption         =   "&Toggle"
      Begin VB.Menu mnuToggleOn 
         Caption         =   "&On"
      End
      Begin VB.Menu mnuToggleOff 
         Caption         =   "O&ff"
      End
   End
   Begin VB.Menu mnuMedia 
      Caption         =   "&Media"
      Begin VB.Menu mnuSelectFile 
         Caption         =   "&Select File"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowToUseThisWindow 
         Caption         =   "H&ow to use this Window"
      End
   End
End
Attribute VB_Name = "frmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdClose_Click()"
End Sub

Private Sub cmdHelp_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DisplayHelpInformation 5
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdHelp_Click()"
End Sub

Private Sub cmdOFF_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Caption = "Alarm"
tmrAlarm.Enabled = False
cmdON.Value = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdON_Click()"
End Sub

Private Sub cmdON_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ToggleAlarm True, txtTime.Text, txtAlarmDate.Text, txtAudio.Text
cmdOFF.Value = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdON_Click()"
End Sub

Private Sub cmdSelect_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer, msg2 As String
msg = OpenDialog(Me, "Mp3 Files (*.mp3)|*.mp3|", "Select MP3 File", App.Path & "\data\sounds\")
msg2 = msg
msg2 = GetFileTitle(msg2)
If Len(msg) <> 0 And DoesFileExist(msg) = True Then
    i = FindFileIndex(msg)
    If i <> 0 Then
        txtAudio.Text = msg2
    Else
        AddToPlaylist msg, True
        txtAudio.Text = msg2
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdSelect_Click()"
End Sub

Private Sub ctlXPButton1_Click()

End Sub

Private Sub ctlHelp_Click()

End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, msg2 As String
Me.Icon = mdiNexIRC.Icon
SetButtonType cmdON
SetButtonType cmdOFF
SetButtonType cmdSelect
SetButtonType cmdClose
SetButtonType cmdHelp
txtAlarmDate.Text = Date
txtTime.Text = Time
Start:
If lFiles.fCount <> 0 Then
    msg = lFiles.fFile(GetRnd(lFiles.fCount)).fFilename
    If Len(msg) <> 0 And DoesFileExist(msg) = True Then
        msg2 = msg
        msg2 = GetFileTitle(msg2)
        txtAudio.Text = msg2
    Else
        GoTo Start
    End If
End If
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

Private Sub mnuSelectFile_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdSelect_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuSelectFile_Click()"
End Sub

Private Sub mnuToggleOff_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdOFF.Value = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuToggleOff_Click()"
End Sub

Private Sub mnuToggleOn_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdON.Value = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuToggleOn_Click()"
End Sub

Private Sub tmrAlarm_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
SetTime
CheckAlarm
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrAlarm_Timer()"
End Sub

