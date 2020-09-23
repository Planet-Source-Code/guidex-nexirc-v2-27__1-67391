VERSION 5.00
Begin VB.Form frmQuickImage 
   BorderStyle     =   0  'None
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picImage 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2295
      ScaleWidth      =   2535
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmQuickImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub picImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picImage_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Width = picImage.ScaleWidth
Me.Height = picImage.ScaleHeight
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picImage_Resize()"
End Sub
