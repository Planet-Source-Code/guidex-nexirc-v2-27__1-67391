VERSION 5.00
Begin VB.Form frmQuery 
   Caption         =   "NexIRC - Query"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmQuery.frx":0000
   LinkTopic       =   "frmQuery"
   MDIChild        =   -1  'True
   ScaleHeight     =   2910
   ScaleWidth      =   4935
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
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Line linTextSep 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   3960
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lQueryNickname As String

Public Sub SetQueryNickname(lNickName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lQueryNickname = lNickName
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetQueryNickname(lNickname As String)"
End Sub

Public Function ReturnQueryNickname() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnQueryNickname = lQueryNickname
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnQueryNickname() As String"
End Function

Public Sub ActivateResize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Form_Resize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ActivateResize()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = mdiNexIRC.Icon
Call Form_Resize
Me.Height = 5000
Me.Width = 6000
txtIncoming.SetTag "query"
txtIncoming.SetBackColor lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor
txtOutgoing.BackColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor
txtOutgoing.ForeColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sTextColor
If lSettings.sBorderlessObjects = True Then
    txtIncoming.SetBorderStyle True
    txtOutgoing.BorderStyle = 0
Else
    txtIncoming.SetBorderStyle False
    txtOutgoing.BorderStyle = 1
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Me.ScaleHeight <> 0 And Me.ScaleWidth <> 0 Then
    txtIncoming.Height = Me.ScaleHeight - txtOutgoing.Height - 20
    txtIncoming.Width = Me.ScaleWidth
    txtOutgoing.Width = Me.ScaleWidth
    txtOutgoing.Top = Me.ScaleHeight - txtOutgoing.Height + 30
    If txtIncoming.Left <> 0 Then txtIncoming.Left = 0
    If txtOutgoing.Left <> 0 Then txtOutgoing.Left = 0
    linTextSep.x2 = Me.ScaleWidth
    linTextSep.y1 = Me.txtIncoming.Height + 10
    linTextSep.y2 = Me.txtIncoming.Height + 10
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Resize()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To mdiNexIRC.StatusBar.Panels.Count
    If Trim(LCase(mdiNexIRC.StatusBar.Panels(i))) = lQueryNickname Then
        SetQueryName i, ""
        Call RemoveTaskbar(lQueryNickname)
        Exit For
    End If
Next i
'Dim i As Integer
'Dim msg() As String
'msg = Split(Me.Caption, " ")
'For i = 0 To UBound(msg)
'    MsgBox i & ": " & LCase(msg(i)) & "-" & LCase(ReturnQueryName(i))
'    If LCase(msg(i)) = LCase(ReturnQueryName(i)) Then
'        SetQueryName i, ""
'        Call RemoveTaskbar(msg(i))
'    End If
'Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub txtIncoming_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtOutgoing.SetFocus
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtIncoming_Change()"
End Sub

Private Sub txtOutgoing_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtOutgoing.SelStart = 0
txtOutgoing.SelLength = Len(txtOutgoing.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtOutgoing_GotFocus()"
End Sub

Private Sub txtOutgoing_KeyPress(KeyAscii As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim word() As String
word = Split(Me.Caption, Chr(32))
If KeyAscii = 13 Then
    If Left(txtOutgoing.Text, 1) = "/" Then
        ACTION_CHANNEL = word(0)
        Call ProcessInput(Mid(txtOutgoing, 2), Me.txtIncoming, lSettings.sActiveServerForm)
    Else
        lSettings.sActiveServerForm.tcp.SendData "PRIVMSG " & word(0) & " :" & txtOutgoing & vbCrLf
        DoColor txtIncoming, "" & Color.Normal & "<" & Color.Action & "" & lSettings.sNickname & "" & Color.Normal & "> " & txtOutgoing.Text
    End If
    txtOutgoing = ""
    KeyAscii = 0
End If
If KeyAscii = 27 Then Me.WindowState = vbMinimized
If KeyAscii = 11 Then
    Dim starttext As Integer
    starttext = txtOutgoing.SelStart
    txtOutgoing = Mid(txtOutgoing, 1, txtOutgoing.SelStart) & "" & Mid(txtOutgoing, txtOutgoing.SelStart + 1)
    txtOutgoing.SelStart = starttext + 1
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtOutgoing_KeyPress(KeyAscii As Integer)"
End Sub
