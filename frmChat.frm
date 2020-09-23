VERSION 5.00
Begin VB.Form frmChat 
   Caption         =   "NexIRC - DCC Chat:"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "frmChat"
   MDIChild        =   -1  'True
   ScaleHeight     =   2850
   ScaleWidth      =   4965
   Begin nexIRC.ctlTBox txtIncoming 
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   2566
   End
   Begin VB.TextBox txtOutgoing 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   0
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   1560
      Width           =   3975
   End
End
Attribute VB_Name = "frmChat"
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
txtIncoming.SetBackColor lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor
txtOutgoing.BackColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor
txtOutgoing.ForeColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sTextColor
If lSettings.sBorderlessObjects = True Then
    txtIncoming.SetBorderStyle False
    txtOutgoing.BorderStyle = 0
Else
    txtIncoming.SetBorderStyle True
    txtOutgoing.BorderStyle = 1
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtIncoming.Width = Me.ScaleWidth
txtIncoming.Height = Me.ScaleHeight - txtOutgoing.Height
txtOutgoing.Top = txtIncoming.Height
txtOutgoing.Width = Me.ScaleWidth
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Resize()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To UBound(lChatWindowName)
    If LCase(Me.Caption) = LCase(lChatWindowName(i)) Then
        Unload mdiNexIRC.wskChat(i)
    End If
Next i
For i = 1 To UBound(lChatWindowName)
    If LCase(Me.Caption) = LCase(lChatWindowName(i)) Then
        Unload mdiNexIRC.wskChat2(i)
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub
                                                    
Private Sub txtIncoming_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtOutgoing.SetFocus
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtIncoming_GotFocus()"
End Sub

Private Sub txtOutgoing_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtOutgoing.SelStart = 0
txtOutgoing.SelLength = Len(txtOutgoing.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtIncoming_Change()"
End Sub

Private Sub txtOutgoing_KeyPress(KeyAscii As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, msg2 As String
If KeyAscii = 13 Then
    msg = txtOutgoing.Text
    txtOutgoing.Text = ""
    KeyAscii = 0
    For i = 1 To ReturnMaxTCP
        If Trim(LCase(lChatWindowName(i))) = Trim(LCase(Me.Caption)) Or Trim(LCase(lChatWindowx(i).Caption)) = Trim(LCase(Me.Caption)) Then
            If mdiNexIRC.wskChat2(i).State = sckConnected Then mdiNexIRC.wskChat2(i).SendData msg & vbCrLf
                ProcessReplaceString sPm, txtIncoming, lChatWindowName(i), msg
            Exit For
        End If
    Next i
End If
If KeyAscii = 27 Then Me.WindowState = vbMinimized
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtOutgoing_KeyPress(KeyAscii As Integer)"
End Sub
