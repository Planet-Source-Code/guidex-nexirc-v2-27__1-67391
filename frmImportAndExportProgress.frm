VERSION 5.00
Begin VB.Form frmImportAndExportProgress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Import and Export"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5355
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImportAndExportProgress.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin nexIRC.XP_ProgressBar XP_ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   6956042
   End
   Begin VB.Label lblProgress 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmImportAndExportProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = mdiNexIRC.Icon
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub
