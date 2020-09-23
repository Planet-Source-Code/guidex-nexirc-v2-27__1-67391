VERSION 5.00
Begin VB.Form frmTextEditor 
   Caption         =   "NexIRC - Text Editor"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4665
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTextEditor.frx":0000
   LinkTopic       =   "frmTextEditor"
   MDIChild        =   -1  'True
   ScaleHeight     =   3765
   ScaleWidth      =   4665
   Begin VB.TextBox txtIncoming 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmTextEditor"
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
ActivateResize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Call RemoveTaskbar(Me.Caption)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)"
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
Dim mbox As VbMsgBoxResult, msg As String, msg2 As String
If Me.Tag = "*" Then
    If lSettings.sGeneralPrompts = True Then
        mbox = MsgBox(Me.Caption & " has been changed, would you like to save changes?", vbYesNo + vbQuestion)
    Else
        mbox = vbYes
    End If
    If mbox = vbYes Then
        msg = SaveDialog(Me, "Text Files (*.txt)|*.txt|", "Save as ...", CurDir)
        If Len(msg) <> 0 Then
            msg = Left(msg, Len(msg) - 1) & ".txt"
            SaveFile msg, txtIncoming.Text
            Me.Tag = ""
            msg2 = msg
            msg2 = GetFileTitle(msg2)
            Me.Caption = msg2
        End If
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub txtIncoming_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Tag = "*"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtIncoming_Change()"
End Sub

Private Sub txtIncoming_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To mdiNexIRC.StatusBar.Panels.Count
    mdiNexIRC.StatusBar.Panels(i).Bevel = sbrRaised
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtIncoming_GotFocus()"
End Sub

Private Sub txtIncoming_KeyPress(KeyAscii As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If KeyAscii = 27 Then Me.WindowState = vbMinimized
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtIncoming_KeyPress(KeyAscii As Integer)"
End Sub
