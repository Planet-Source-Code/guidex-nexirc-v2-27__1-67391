VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Tips"
   ClientHeight    =   3300
   ClientLeft      =   2355
   ClientTop       =   2685
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   5160
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Appearance      =   0  'Flat
      Caption         =   "&Show Tips at Startup"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2940
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2715
      Left            =   120
      OLEDropMode     =   2  'Automatic
      ScaleHeight     =   2715
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know..."
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3495
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   2115
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3495
      End
   End
   Begin nexIRC.ctlXPButton cmdClose 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      ToolTipText     =   "Cancel/Hide this Window"
      Top             =   120
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
      MICON           =   "frmTip.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdNextTip 
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      ToolTipText     =   "Cancel/Hide this Window"
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Next Tip"
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
      MICON           =   "frmTip.frx":0028
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
      Left            =   3960
      TabIndex        =   6
      ToolTipText     =   "Cancel/Hide this Window"
      Top             =   2480
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
      MICON           =   "frmTip.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   2760
      Left            =   90
      Top             =   105
      Width           =   3795
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNextTip 
         Caption         =   "&Next Tip"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuShowTipsAtStartup 
         Caption         =   "&Show Tips at Startup"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowToUseThisWindow 
         Caption         =   "H&ow to use this Window"
      End
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tips As New Collection
Dim CurrentTip As Long

Private Sub DoNextTip()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CurrentTip = Int((Tips.Count * Rnd) + 1)
frmTip.DisplayCurrentTip
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub DoNextTip()"
End Sub

Function LoadTips(sFile As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim NextTip As String, InFile As Integer
InFile = FreeFile
If sFile = "" Then
    LoadTips = False
    Exit Function
End If
If Dir(sFile) = "" Then
    LoadTips = False
    Exit Function
End If
Open sFile For Input As InFile
While Not EOF(InFile)
    Line Input #InFile, NextTip
    Tips.Add NextTip
Wend
Close InFile
DoNextTip
LoadTips = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Function LoadTips(sFile As String) As Boolean"
End Function

Private Sub chkLoadTipsAtStartup_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sShowTips = GetCheckboxValue(chkLoadTipsAtStartup)
WriteINI GetINIFile(iIRC), "Settings", "ShowTips", lSettings.sShowTips
If lSettings.sCustomizeVisible = True Then frmCustomize.chkShowTips.Value = chkLoadTipsAtStartup.Value
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Function LoadTips(sFile As String) As Boolean"
End Sub

Private Sub cmdClose_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdOK_Click()"
End Sub

Private Sub cmdHelp_Click()
MsgBox "No help is available for this topic", vbExclamation
End Sub

Private Sub cmdNextTip_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DoNextTip
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdNextTip_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = mdiNexIRC.Icon
Me.chkLoadTipsAtStartup.Value = vbChecked
If chkLoadTipsAtStartup.Value = vbChecked Then mnuShowTipsAtStartup.Checked = True
Randomize
If LoadTips(App.Path & "\data\config\fixed\tips.ini") = False Then
    Unload Me
End If
SetButtonType cmdClose
SetButtonType cmdNextTip
SetButtonType cmdHelp
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Public Sub DisplayCurrentTip()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Tips.Count > 0 Then
    lblTipText.Caption = Tips.Item(CurrentTip)
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub DisplayCurrentTip()"
End Sub

Private Sub mnuExit_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdClose_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuExit_Click()"
End Sub

Private Sub mnuHowToUseThisWindow_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdHelp_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuHowToUseThisWindow_Click()"
End Sub

Private Sub mnuNextTip_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdNextTip_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuNextTip_Click()"
End Sub

Private Sub mnuShowTipsAtStartup_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case mnuShowTipsAtStartup.Checked
Case True
    mnuShowTipsAtStartup.Checked = False
    chkLoadTipsAtStartup.Value = 0
Case False
    mnuShowTipsAtStartup.Checked = True
    chkLoadTipsAtStartup.Value = 1
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuShowTipsAtStartup_Click()"
End Sub
