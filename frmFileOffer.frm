VERSION 5.00
Object = "{CECDEAC1-E92C-11D2-B1AA-300962C10000}#1.0#0"; "VFmp3player.ocx"
Begin VB.Form frmAudioServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Audio Server"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFileOffer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   3540
   Begin VB.CheckBox chkVisFX 
      Appearance      =   0  'Flat
      Caption         =   "VisFX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VFMP3PLAYERLib.VFmp3level level1 
      Height          =   2655
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
      _ExtentY        =   4683
      _StockProps     =   0
      ActiveColor1    =   14737632
      ActiveColor2    =   12632256
      ActiveColor3    =   8421504
      InactiveColor1  =   16711680
      InactiveColor2  =   12582912
      InactiveColor3  =   8388608
      GapHeight       =   1
      BorderStyle     =   1
   End
   Begin VB.ListBox lstConnections 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2700
      IntegralHeight  =   0   'False
      ItemData        =   "frmFileOffer.frx":0CCA
      Left            =   0
      List            =   "frmFileOffer.frx":0CCC
      TabIndex        =   2
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmAudioServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next

If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub
