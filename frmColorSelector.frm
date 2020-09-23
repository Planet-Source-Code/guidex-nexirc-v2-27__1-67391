VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColorSelector 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Color Selector"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   3210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmColorSelector.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   3210
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picColor 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   840
      ScaleHeight     =   555
      ScaleWidth      =   2235
      TabIndex        =   6
      Top             =   1200
      Width           =   2295
   End
   Begin MSComctlLib.Slider sldRed 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   75
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   661
      _Version        =   393216
      Max             =   255
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider sldGreen 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   435
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   661
      _Version        =   393216
      Max             =   255
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider sldBlue 
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   795
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   661
      _Version        =   393216
      Max             =   255
      TickStyle       =   3
   End
   Begin nexIRC.ctlXPButton cmdCancel 
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   1920
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
      MICON           =   "frmColorSelector.frx":000C
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
      Left            =   840
      TabIndex        =   9
      Top             =   1920
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
      MICON           =   "frmColorSelector.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      Caption         =   "Preview:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Blue:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Green:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Red:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmColorSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UpdateRGB()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mColor As Long, HexColor As String, R As Byte, G As Byte, b As Byte
mColor = RGB(Int(sldRed.Value), Int(sldGreen), Int(sldBlue.Value))
picColor.BackColor = mColor
R = (mColor And &HFF&)
b = (mColor And &HFF00&) / &H100&
G = (mColor And &HFF0000) / &H10000
lReturnColor = "&H" & Right("0" & Hex(G), 2) & Right("0" & Hex(b), 2) & Right("0" & Hex(R), 2) & "&"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub UpdateRGB()"
End Sub

Private Sub cmdCancel_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lReturnColor = ""
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdCancel_Click()"
End Sub

Private Sub cmdOK_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdOK_Click()"
End Sub

Private Sub Form_Load()
Me.Icon = mdiNexIRC.Icon
End Sub

Private Sub mnuExit_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuExit_Click()"
End Sub

Private Sub sldBlue_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UpdateRGB
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub sldBlue_Change()"
End Sub

Private Sub sldGreen_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UpdateRGB
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub sldGreen_Change()"
End Sub

Private Sub sldRed_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UpdateRGB
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub sldRed_Change()"
End Sub

Private Sub sldRed_Scroll()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UpdateRGB
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub sldRed_Scroll()"
End Sub
