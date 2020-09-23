VERSION 5.00
Begin VB.Form frmSelectShape 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Color"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelectShape.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3120
   StartUpPosition =   1  'CenterOwner
   Begin nexIRC.ctlXPButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   18
      Top             =   3120
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
      MICON           =   "frmSelectShape.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox picCurrentColor 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   17
      Top             =   2400
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   14
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   14
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   13
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   12
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   12
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   11
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   11
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   10
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   9
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   8
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   7
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   6
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   5
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   4
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   3
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   2
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   1
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   0
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
   Begin nexIRC.ctlXPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   19
      Top             =   3120
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
      MICON           =   "frmSelectShape.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   3120
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label2 
      Caption         =   "Current Color:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Choices:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmSelectShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lBackColor As Integer

Private Sub cmdOK_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sBGColor = lBackColor
SaveSettings
mdiNexIRC.BackColor = QBColor(lBackColor)
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdOK_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = mdiNexIRC.Icon
SetButtonType cmdOK
Dim i As Integer
For i = 0 To 14
    picColor(i).BackColor = QBColor(i)
Next i
picCurrentColor.BackColor = QBColor(lSettings.sBGColor)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub picColor_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To 14
    picColor(i).Enabled = False
Next i
picColor(Index).Enabled = True
lBackColor = Index
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picColor_Click(Index As Integer)"
End Sub
