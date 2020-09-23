VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmHelpSystem 
   Caption         =   "NexIRC - Help"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3705
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHelpSystem.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin OsenXPCntrl.OsenXPButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   4560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "OK"
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
      MICON           =   "frmHelpSystem.frx":0CCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cancel"
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
      MICON           =   "frmHelpSystem.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraHelpTopic 
      BorderStyle     =   0  'None
      Caption         =   "Help Topics"
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3495
      Begin OsenXPCntrl.OsenXPButton cmdPrint 
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   3600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Print"
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
         MICON           =   "frmHelpSystem.frx":0D02
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton cmdViewExample 
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   3600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Example"
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
         MICON           =   "frmHelpSystem.frx":0D1E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtText 
         Height          =   3255
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lblHelpDescription 
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   0
         Width           =   2175
      End
      Begin VB.Label labell1 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.ComboBox cboHelpTopic 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblTopic 
      Caption         =   "Topic:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmHelpSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboHelpTopic_Click()
On Local Error Resume Next
If Len(cboHelpTopic.Text) <> 0 Then
    lblHelpDescription.Caption = cboHelpTopic.Text
    Select Case cboHelpTopic.Text
    Case "Using irc"
        txtText.Text = "Using irc is an incredibly great and easy way of chatting with other users online."
    Case "Using help topics"
        txtText.Text = "To use the help topics windows, simply click the topic box, and select your topic information. Once you have read the description, you may click example to view a visual representation of this in action. If you need to have this information handy, simply click 'print' and the text will feed to your printer. Once you are done using the help features, click 'close' at the bottom of this window."
    End Select
End If
End Sub

Private Sub cmdCancel_Click()
On Local Error Resume Next
Unload Me
End Sub

Private Sub cmdOK_Click()
On Local Error Resume Next
Unload Me
End Sub

Private Sub Form_Load()
On Local Error Resume Next
cboHelpTopic.AddItem "Using help topics"
cboHelpTopic.AddItem "Using irc"
End Sub
