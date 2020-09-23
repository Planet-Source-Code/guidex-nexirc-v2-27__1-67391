VERSION 5.00
Begin VB.Form frmGraphics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Graphics"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1245
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGraphics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   1245
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picChat1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   44
      Top             =   5160
      Width           =   300
   End
   Begin VB.PictureBox picChat2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   43
      Top             =   5160
      Width           =   300
   End
   Begin VB.PictureBox picChat3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   42
      Top             =   5160
      Width           =   300
   End
   Begin VB.PictureBox picExit3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   41
      Top             =   4800
      Width           =   300
   End
   Begin VB.PictureBox picExit2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   40
      Top             =   4800
      Width           =   300
   End
   Begin VB.PictureBox picExit1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   39
      Top             =   4800
      Width           =   300
   End
   Begin VB.PictureBox picNexIRC3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   38
      Top             =   4440
      Width           =   300
   End
   Begin VB.PictureBox picNexIRC2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   37
      Top             =   4440
      Width           =   300
   End
   Begin VB.PictureBox picForward3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   36
      Top             =   4080
      Width           =   300
   End
   Begin VB.PictureBox picPlay3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   35
      Top             =   3720
      Width           =   300
   End
   Begin VB.PictureBox picPause3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   34
      Top             =   3360
      Width           =   300
   End
   Begin VB.PictureBox picStop3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   33
      Top             =   3000
      Width           =   300
   End
   Begin VB.PictureBox picBackward3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   32
      Top             =   2640
      Width           =   300
   End
   Begin VB.PictureBox picForward2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   31
      Top             =   4080
      Width           =   300
   End
   Begin VB.PictureBox picForward1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   30
      Top             =   4080
      Width           =   300
   End
   Begin VB.PictureBox picPlay2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   29
      Top             =   3720
      Width           =   300
   End
   Begin VB.PictureBox picPlay1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   28
      Top             =   3720
      Width           =   300
   End
   Begin VB.PictureBox picPause2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   27
      Top             =   3360
      Width           =   300
   End
   Begin VB.PictureBox picPause1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   26
      Top             =   3360
      Width           =   300
   End
   Begin VB.PictureBox picStop2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   25
      Top             =   3000
      Width           =   300
   End
   Begin VB.PictureBox picStop1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   24
      Top             =   3000
      Width           =   300
   End
   Begin VB.PictureBox picBackward2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   23
      Top             =   2640
      Width           =   300
   End
   Begin VB.PictureBox picBackward1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   22
      Top             =   2640
      Width           =   300
   End
   Begin VB.PictureBox picScript3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   21
      Top             =   2280
      Width           =   300
   End
   Begin VB.PictureBox picScript2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   20
      Top             =   2280
      Width           =   300
   End
   Begin VB.PictureBox picScript1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   19
      Top             =   2280
      Width           =   300
   End
   Begin VB.PictureBox picNexIRC1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   18
      Top             =   4440
      Width           =   300
   End
   Begin VB.PictureBox picSend3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   17
      Top             =   1920
      Width           =   300
   End
   Begin VB.PictureBox picSend2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   16
      Top             =   1920
      Width           =   300
   End
   Begin VB.PictureBox picSend1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   15
      Top             =   1920
      Width           =   300
   End
   Begin VB.PictureBox picChannelFolder3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   14
      Top             =   1560
      Width           =   300
   End
   Begin VB.PictureBox picChannelFolder2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   13
      Top             =   1560
      Width           =   300
   End
   Begin VB.PictureBox picChannelFolder1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   12
      Top             =   1560
      Width           =   300
   End
   Begin VB.PictureBox picDisconnect2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   300
   End
   Begin VB.PictureBox picDisconnect3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   300
   End
   Begin VB.PictureBox picDisconnect1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   300
   End
   Begin VB.PictureBox picOptions2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   120
      Width           =   300
   End
   Begin VB.PictureBox picOptions3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   120
      Width           =   300
   End
   Begin VB.PictureBox picOptions1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   120
      Width           =   300
   End
   Begin VB.PictureBox picAudio2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   840
      Width           =   300
   End
   Begin VB.PictureBox picAudio3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   840
      Width           =   300
   End
   Begin VB.PictureBox picAudio1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   840
      Width           =   300
   End
   Begin VB.PictureBox picConnect2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   480
      Width           =   300
   End
   Begin VB.PictureBox picConnect3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   480
      Width           =   300
   End
   Begin VB.PictureBox picConnect1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   480
      Width           =   300
   End
End
Attribute VB_Name = "frmGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
