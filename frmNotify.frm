VERSION 5.00
Begin VB.Form frmNotify 
   Caption         =   "NexIRC - Notify"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNotify.frx":0000
   LinkTopic       =   "frmNotify"
   MDIChild        =   -1  'True
   ScaleHeight     =   2535
   ScaleWidth      =   5205
   Begin VB.ListBox lstNotify 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2190
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmNotify"
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
lSettings.sNotifyVisible = True
Call AddTaskPanel("Notify", 1)
lstNotify.BackColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor
lstNotify.ForeColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sTextColor
ActivateResize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lstNotify.Left <> 0 Then lstNotify.Left = 0
lstNotify.Width = Me.ScaleWidth
lstNotify.Height = Me.ScaleHeight
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Resize()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sNotifyVisible = False
Call RemoveTaskbar("Notify")
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub lstNotify_DblClick()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
NewQuery lstNotify.Text, True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lstNotify_DblClick()"
End Sub

Private Sub lstNotify_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(lstNotify.Text) <> 0 And Button = 2 Then
    PopupMenu frmMenus.mnuNotify4
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lstNotify_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

