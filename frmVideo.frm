VERSION 5.00
Begin VB.Form frmVideo 
   Caption         =   "NexIRC - Media"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVideo.frx":0000
   LinkTopic       =   "frmVideo"
   MDIChild        =   -1  'True
   ScaleHeight     =   2985
   ScaleWidth      =   4680
   Visible         =   0   'False
   Begin VB.Frame fraVideo 
      BackColor       =   &H00800000&
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
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
fraVideo.Width = Me.ScaleWidth
fraVideo.Height = Me.ScaleHeight
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Resize()"
End Sub
