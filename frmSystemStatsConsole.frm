VERSION 5.00
Begin VB.Form frmSystemStatsConsole 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - System Stats"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSystemStatsConsole.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   3255
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSCSIDevices 
      Appearance      =   0  'Flat
      Caption         =   "SCSI Devices"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   2760
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox chkPCName 
      Appearance      =   0  'Flat
      Caption         =   "PC Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   2520
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkProcessor2 
      Appearance      =   0  'Flat
      Caption         =   "Processor 2"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkProcessor1 
      Appearance      =   0  'Flat
      Caption         =   "Processor 1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkProcessorCount 
      Appearance      =   0  'Flat
      Caption         =   "Processor Count"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox chkOS 
      Appearance      =   0  'Flat
      Caption         =   "OS"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox txtSystemStats 
      Height          =   2295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin nexIRC.ctlXPButton cmdDisplay 
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   7
      ToolTipText     =   "Cancel/Hide this Window"
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Display"
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
      MICON           =   "frmSystemStatsConsole.frx":000C
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
      Left            =   1920
      TabIndex        =   8
      ToolTipText     =   "Cancel/Hide this Window"
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmSystemStatsConsole.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmSystemStatsConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkOS_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdDisplay_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkOS_Click()"
End Sub

Private Sub chkPCName_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdDisplay_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkPCName_Click()"
End Sub

Private Sub chkProcessor1_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdDisplay_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkProcessor1_Click()"
End Sub

Private Sub chkProcessor2_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdDisplay_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkProcessor2_Click()"
End Sub

Private Sub chkProcessorCount_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdDisplay_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkProcessorCount_Click()"
End Sub

Private Sub chkSCSIDevices_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdDisplay_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkSCSIDevices_Click()"
End Sub

Private Sub cmdClose_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdClose_Click()"
End Sub

Private Sub cmdDisplay_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ShowSystemStats lSettings.sActiveServerForm, GetCheckboxValue(chkOS), GetCheckboxValue(chkPCName), GetCheckboxValue(chkProcessorCount), GetCheckboxValue(chkProcessor1), GetCheckboxValue(chkProcessor2), GetCheckboxValue(chkSCSIDevices), True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdDisplay_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sSystemStatsConsoleVisible = True
SetButtonType cmdClose
SetButtonType cmdDisplay
Me.Icon = mdiNexIRC.Icon
ShowSystemStats lSettings.sActiveServerForm, GetCheckboxValue(chkOS), GetCheckboxValue(chkPCName), GetCheckboxValue(chkProcessorCount), GetCheckboxValue(chkProcessor1), GetCheckboxValue(chkProcessor2), GetCheckboxValue(chkSCSIDevices), True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
lSettings.sSystemStatsConsoleVisible = False
End Sub
