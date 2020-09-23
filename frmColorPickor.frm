VERSION 5.00
Begin VB.Form frmColorPickor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Choose Color"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3165
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmColorPickor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3165
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   0
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   14
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   1
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   2
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   3
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   4
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   5
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   6
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   8
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   7
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   8
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   9
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   10
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   11
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   12
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   13
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox picColor 
      Height          =   495
      Index           =   14
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   1320
      Width           =   495
   End
End
Attribute VB_Name = "frmColorPickor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To 14
    picColor(i).BackColor = QBColor(i)
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub picColor_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmColorEditor.txtRetValue.Text = str(Index)
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picColor_Click(Index As Integer)"
End Sub
