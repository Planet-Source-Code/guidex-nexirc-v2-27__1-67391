VERSION 5.00
Begin VB.Form frmAddFolderToPlaylist 
   Caption         =   "NexIRC - Add to Playlist"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   4500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddFolderToPlaylist.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   194
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboPaths 
      Height          =   315
      ItemData        =   "frmAddFolderToPlaylist.frx":000C
      Left            =   120
      List            =   "frmAddFolderToPlaylist.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   4215
   End
   Begin nexIRC.ctlXPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   2400
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
      MICON           =   "frmAddFolderToPlaylist.frx":0025
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdAdd 
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   2400
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
      MICON           =   "frmAddFolderToPlaylist.frx":0041
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
      TabIndex        =   2
      Top             =   2400
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
      MICON           =   "frmAddFolderToPlaylist.frx":005D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   4215
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowToUse 
         Caption         =   "H&ow to Use this Window"
      End
   End
End
Attribute VB_Name = "frmAddFolderToPlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboPaths_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dir1.Path = GetMyDocumentsDir()
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cboPaths_Click()"
End Sub

Private Sub cmdAdd_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lDirReturn = Dir1.Path
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAdd_Click()"
End Sub

Private Sub cmdCancel_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdCancel_Click()"
End Sub

Private Sub cmdHelp_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DisplayHelpInformation 1
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdOK_Click()"
End Sub

Private Sub Drive1_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dir1.Path = Drive1.Drive
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Drive1_Change()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = mdiNexIRC.Icon
SetButtonType cmdCancel
SetButtonType cmdAdd
SetButtonType cmdHelp
lDirReturn = ""
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_Resize()
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
cboPaths.Width = frmAddFolderToPlaylist.ScaleWidth - 18
Dir1.Width = frmAddFolderToPlaylist.ScaleWidth - 18
Dir1.Height = frmAddFolderToPlaylist.ScaleHeight - (70 + cboPaths.Height)
Drive1.Width = frmAddFolderToPlaylist.ScaleWidth - 18
Drive1.Top = frmAddFolderToPlaylist.ScaleHeight - 60
cmdCancel.Left = frmAddFolderToPlaylist.ScaleWidth - (cmdCancel.Width + 10)
cmdCancel.Top = frmAddFolderToPlaylist.ScaleHeight - (30)
cmdAdd.Left = frmAddFolderToPlaylist.ScaleWidth - (cmdAdd.Width + 93)
cmdAdd.Top = frmAddFolderToPlaylist.ScaleHeight - 30
cmdHelp.Top = frmAddFolderToPlaylist.ScaleHeight - 30
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Resize()"
    Err.Clear
End Sub

Private Sub mnuExit_Click()
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
Unload Me
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuExit_Click()"
    Err.Clear
End Sub

Private Sub mnuHowToUse_Click()
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
DisplayHelpInformation 1
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuHowToUse_Click()"
    Err.Clear
End Sub
