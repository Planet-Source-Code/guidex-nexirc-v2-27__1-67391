VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin NexIRC.ctlXPButton d 
      Height          =   615
      Index           =   1
      Left            =   2760
      TabIndex        =   11
      Top             =   1440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "ctlXPButton2"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "Form1.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin NexIRC.ctlXPButton d 
      Height          =   855
      Index           =   0
      Left            =   2640
      TabIndex        =   12
      Top             =   360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "ctlXPButton1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "Form1.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin VB.PictureBox optCheck 
      Height          =   360
      Index           =   8
      Left            =   840
      ScaleHeight     =   300
      ScaleWidth      =   1035
      TabIndex        =   10
      Top             =   4320
      Width           =   1095
   End
   Begin VB.PictureBox optCheck 
      Height          =   360
      Index           =   6
      Left            =   840
      ScaleHeight     =   300
      ScaleWidth      =   1035
      TabIndex        =   9
      Top             =   4680
      Width           =   1095
   End
   Begin VB.PictureBox optCheck 
      Height          =   360
      Index           =   0
      Left            =   840
      ScaleHeight     =   300
      ScaleWidth      =   1035
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.PictureBox optCheck 
      Height          =   360
      Index           =   1
      Left            =   840
      ScaleHeight     =   300
      ScaleWidth      =   1035
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.PictureBox optCheck 
      Height          =   360
      Index           =   2
      Left            =   840
      ScaleHeight     =   300
      ScaleWidth      =   1035
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.PictureBox optCheck 
      Height          =   360
      Index           =   9
      Left            =   840
      ScaleHeight     =   300
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.PictureBox optCheck 
      Height          =   360
      Index           =   3
      Left            =   840
      ScaleHeight     =   300
      ScaleWidth      =   1035
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.PictureBox optCheck 
      Height          =   360
      Index           =   4
      Left            =   840
      ScaleHeight     =   300
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.PictureBox optCheck 
      Height          =   360
      Index           =   5
      Left            =   840
      ScaleHeight     =   300
      ScaleWidth      =   1035
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.PictureBox optCheck 
      Height          =   360
      Index           =   7
      Left            =   840
      ScaleHeight     =   300
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin VB.PictureBox optCheck 
      Height          =   360
      Index           =   10
      Left            =   840
      ScaleHeight     =   300
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   3480
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub d_Click(Index As Integer)
If Index = 0 Then
    d(1).Value = False
ElseIf Index = 1 Then
    d(0).Value = False
End If



End Sub
