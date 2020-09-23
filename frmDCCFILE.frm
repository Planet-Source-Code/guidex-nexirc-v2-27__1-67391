VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmDCCFILE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - DCC Get"
   ClientHeight    =   3120
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
   Icon            =   "frmDCCFILE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4305
   StartUpPosition =   1  'CenterOwner
   Begin nexIRC.XP_ProgressBar ProgressBar 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   6956042
   End
   Begin VB.PictureBox picComplete 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin nexIRC.ctlXPButton cmdHelp 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2640
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
      MICON           =   "frmDCCFILE.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdGet 
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   2640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Get"
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
      MICON           =   "frmDCCFILE.frx":0028
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
      Left            =   3240
      TabIndex        =   13
      Top             =   2640
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "frmDCCFILE.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSWinsockLib.Winsock FILE 
      Index           =   0
      Left            =   3720
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Filename:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nickname:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Filesize:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   570
   End
   Begin VB.Label lblNickName 
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblFile 
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblFileSize 
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label lblAddress 
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label lblRCV 
      Caption         =   "0"
      Height          =   195
      Left            =   1200
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label lblFilename 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOther 
      Caption         =   "&Other"
      Begin VB.Menu mnuDownloadManager 
         Caption         =   "&Download Manager"
      End
      Begin VB.Menu mnuDCCChat 
         Caption         =   "&DCC Chat"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowToUseThisWindow 
         Caption         =   "H&ow to use this Window"
      End
   End
End
Attribute VB_Name = "frmDCCFILE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload FILE(Me.Tag)
Close #Me.Tag
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdClose_Click()"
End Sub

Private Sub cmdGET_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim cc() As String
cc = Split(lblAddress, ":")
FILE(Me.Tag).Close
If cc(1) <> "-" Then
    FILE(Me.Tag).Connect cc(0), cc(1)
End If
cmdGet.Enabled = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdGET_Click()"
End Sub

Private Sub cmdHelp_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DisplayHelpInformation 17
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdHelp_Click()"
End Sub

Private Sub FILE_Close(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'Call DoColor(lSettings.sActiveServerForm.txtIncoming, "4* File Connection Closed (" & Me.Tag & ")")
ProcessReplaceString sConnectionClosed, lSettings.sActiveServerForm.txtIncoming
Close #Me.Tag
cmdClose.Caption = "Close"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub FILE_Close(Index As Integer)"
End Sub

Private Sub FILE_Connect(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String
Call DoColor(lSettings.sActiveServerForm.txtIncoming, "4* Ready for file transfer")
msg = Me.lblFilename.Caption
Open msg For Binary As Me.Tag
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub FILE_Connect(Index As Integer)"
End Sub

Private Sub FILE_DataArrival(Index As Integer, ByVal bytesTotal As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim Hexdata As String, i As Integer
Me.ProgressBar.Value = Me.ProgressBar.Value + bytesTotal
Dim ReadBuffer() As Byte
Dim RetVal As Long
FILE(Me.Tag).GetData ReadBuffer, vbByte
Put #Me.Tag, , ReadBuffer
Hexdata = Hex(LOF(Me.Tag))
Hexdata = String$(8 - Len(Hexdata), "0") & Hexdata
ReDim SendBackData(3) As Byte
For i = 1 To Len(Hexdata) Step 2
    SendBackData((i - 1) / 2) = Val("&H" & Mid(Hexdata, i, 2))
Next
FILE(Me.Tag).SendData SendBackData
If Me.ProgressBar.Value = Me.ProgressBar.Max Then
    Me.picComplete.BackColor = vbBlue
    cmdClose.Caption = "Close"
    Caption = "Transfer Complete"
    If lSettings.sDownloadManager = True Then frmDownloadManager.Show 0, Me
Else
    Me.picComplete.BackColor = vbWhite
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub FILE_DataArrival(Index As Integer, ByVal bytesTotal As Long)"
End Sub

Private Sub FILE_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdClose.Caption = "Close"
Close #Me.Tag
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub FILE_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = mdiNexIRC.Icon
SetButtonType cmdHelp
SetButtonType cmdGet
SetButtonType cmdClose
Close #Me.Tag
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub mnuDCCChat_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmDCC_Chat.Show
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuDCCChat_Click()"
End Sub

Private Sub mnuDownloadManager_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmDownloadManager.Show
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuDownloadManager_Click()"
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
