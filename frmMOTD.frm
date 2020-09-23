VERSION 5.00
Begin VB.Form frmMOTD 
   Caption         =   "NexIRC - Message of the day"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4035
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMOTD.frx":0000
   LinkTopic       =   "frmMOTD"
   MDIChild        =   -1  'True
   ScaleHeight     =   1530
   ScaleWidth      =   4035
   Begin nexIRC.ctlTBox txtIncoming 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   2566
   End
End
Attribute VB_Name = "frmMOTD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ActivateResize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Form_Resize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ActivateResize()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = mdiNexIRC.Icon
lSettings.sMOTDVisible = True
Call AddTaskPanel("MOTD", 1)
txtIncoming.SetBackColor lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If txtIncoming.Left <> 0 Then txtIncoming.Left = 0
txtIncoming.Width = Me.ScaleWidth
txtIncoming.Height = Me.ScaleHeight
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Resize()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Call RemoveTaskbar("MOTD")
lSettings.sMOTDVisible = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub
