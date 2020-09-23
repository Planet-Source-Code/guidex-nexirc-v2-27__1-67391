VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStatus 
   Caption         =   "NexIRC - Status"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "frmStatus"
   MDIChild        =   -1  'True
   ScaleHeight     =   2700
   ScaleWidth      =   4545
   Visible         =   0   'False
   Begin nexIRC.ctlTBox txtIncoming 
      Height          =   1455
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2566
   End
   Begin MSComctlLib.Toolbar tlbStatus 
      Align           =   4  'Align Right
      Height          =   2700
      Left            =   4185
      TabIndex        =   1
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   4763
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iltStatus"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Connect to IRC Server"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Send Quit Message"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Previous typed word"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Next Typed word"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "List item up"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "List Item Down"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save Log"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iltStatus 
      Left            =   120
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatus.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatus.frx":00FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatus.frx":01F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatus.frx":02EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatus.frx":03DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatus.frx":047D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatus.frx":0540
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstSent 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer tmrReconnect 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   1200
      Top             =   2040
   End
   Begin VB.CheckBox chkMOTDActivated 
      Caption         =   "MOTD"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtOutgoing 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   4095
   End
   Begin MSWinsockLib.Winsock tcp 
      Left            =   720
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Line linTextSep 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   4080
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private cFM As clsFMenu
Private lScriptData As String
Private lScriptFile As String
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub Reconnect()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mbox As VbMsgBoxResult, lServer As String, lPort As String, i As Integer
If lSettings.sReconnectOnDisconnect = True Then
    i = Int(Trim(Right(Me.Tag, Len(Me.Tag) - 7)))
    If i <> 0 And Len(ReturnStatusWindowServer(i)) <> 0 And Len(ReturnStatusWindowPort(i)) <> 0 Then
        lServer = ReturnStatusWindowServer(i)
        lPort = ReturnStatusWindowPort(i)
    Else
        lServer = lSettings.sServer
        lPort = lSettings.sPort
    End If
    If lSettings.sGeneralPrompts = True Then
        mbox = MsgBox("Connection to " & lServer & ": " & lPort & " has been closed. Would you like to reconnect?", vbYesNo + vbQuestion, "Reconnect?")
        If mbox = vbYes Then
            ConnectToIRC lServer, lPort, Me
        Else
            tmrReconnect.Enabled = False
        End If
    Else
        ConnectToIRC lServer, lPort, Me
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub Reconnect()"
End Sub

Public Sub ActivateResize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Form_Resize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ActivateResize()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim fm As clsFMenu
'Me.Visible = True
txtIncoming.SetTag "status"
Me.Icon = mdiNexIRC.Icon
If Len(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor) <> 0 Then Me.BackColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor
If Len(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor) <> 0 Then txtIncoming.SetBackColor lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor
If Len(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor) <> 0 Then txtOutgoing.BackColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor
If Len(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sTextColor) <> 0 Then txtOutgoing.ForeColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sTextColor
Set lSettings.sActiveServerForm = Me
Call AddTaskPanel("Status " & ReturnStatusWindowCount, 1)
Call Form_Resize
If lSettings.sBorderlessObjects = True Then
    txtOutgoing.BorderStyle = 0
    txtIncoming.SetBorderStyle True
Else
    txtOutgoing.BorderStyle = 1
    txtIncoming.SetBorderStyle False
End If
Set fm = New clsFMenu
fm.RunScriptFile App.Path & "\data\scripts\nexirc\on(status_show).nirc"
Me.Visible = True
txtOutgoing.Visible = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Call RemoveTaskbar(Me.Tag)
Set lSettings.sActiveServerForm = Nothing
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)"
End Sub

Private Sub Form_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Me.ScaleHeight <> 0 And Me.ScaleWidth <> 0 Then
    txtIncoming.Height = (Me.ScaleHeight - txtOutgoing.Height) - 20
    txtIncoming.Width = Me.ScaleWidth - tlbStatus.Width
    txtOutgoing.Width = Me.ScaleWidth - tlbStatus.Width
    txtOutgoing.Top = Me.ScaleHeight - txtOutgoing.Height
    If txtIncoming.Left <> 0 Then txtIncoming.Left = 0
    If txtOutgoing.Left <> 0 Then txtOutgoing.Left = 0
    lstSent.Left = txtIncoming.Left
    lstSent.Top = txtIncoming.Top
    lstSent.Width = txtIncoming.Width
    lstSent.Height = txtIncoming.Height
    linTextSep.x2 = Me.ScaleWidth
    linTextSep.y1 = Me.txtIncoming.Height
    linTextSep.y2 = Me.txtIncoming.Height
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Resize()"
End Sub

Private Sub lstSent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtOutgoing.Text = lstSent.Text
lstSent.Visible = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lstSent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub tlbStatus_ButtonClick(ByVal Button As MSComctlLib.Button)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
Select Case Button.Index
Case 1
    ConnectToIRC lSettings.sServer, lSettings.sPort, Me
Case 2
    SendQuitMessage Me, ReturnReplacedString(sQuitReason)
Case 4
    lstSent.Visible = False
    lstSent.ListIndex = lstSent.ListIndex - 1
    txtOutgoing.Text = lstSent.Text
Case 5
    lstSent.Visible = False
    lstSent.ListIndex = lstSent.ListIndex + 1
    txtOutgoing.Text = lstSent.Text
Case 6
    lstSent.Visible = True
    lstSent.ListIndex = lstSent.ListIndex - 1
    txtOutgoing.Text = lstSent.Text
Case 7
    lstSent.Visible = True
    lstSent.ListIndex = lstSent.ListIndex + 1
    txtOutgoing.Text = lstSent.Text
Case 9
    msg = "Status-" & Date$ & ".log"
    If Len(msg) <> 0 Then
        SaveFile App.Path & "\data\logs\" & msg, mdiNexIRC.ActiveForm.txtIncoming.Text
        ProcessReplaceString sSaveLog, mdiNexIRC.ActiveForm.txtIncoming, msg
    End If
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tlbStatus_ButtonClick(ByVal Button As MSComctlLib.Button)"
End Sub

Private Sub tmrReconnect_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Reconnect
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrReconnect_Timer()"
End Sub

Private Sub txtIncoming_GotFocus()
txtOutgoing.SetFocus
End Sub

Private Sub txtOutgoing_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
txtOutgoing.SelStart = 0
txtOutgoing.SelLength = Len(txtOutgoing.Text)
Set lSettings.sActiveServerForm = Me
For i = 1 To mdiNexIRC.StatusBar.Panels.Count
    mdiNexIRC.StatusBar.Panels(i).Bevel = sbrRaised
Next i
mdiNexIRC.StatusBar.Panels(FindPanelIndex(Me.Tag, mdiNexIRC.StatusBar)).Bevel = sbrInset
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtOutgoing_GotFocus()"
End Sub

Private Sub txtOutgoing_KeyDown(KeyCode As Integer, Shift As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = txtOutgoing.Text
If ProcessKeyDown(KeyCode, Shift, msg, lstSent, txtOutgoing) = True Then
    KeyCode = 0
End If
If KeyCode = 13 Then
    txtOutgoing.Text = ""
    If Left(msg, 1) = "/" Then
        ACTION_CHANNEL = ""
        Call ProcessInput(Mid(msg, 2), txtIncoming, Me)
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtOutgoing_KeyDown(KeyCode As Integer, Shift As Integer)"
End Sub

Private Sub txtOutgoing_KeyPress(KeyAscii As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim F As clsFMenu, m As Boolean, k As New frmTextEditor, msg As String
If KeyAscii = 27 Then
    Me.WindowState = vbMinimized
End If
If KeyAscii = 13 Then
    txtOutgoing = LTrim(txtOutgoing)
    If Left(txtOutgoing.Text, 1) = "/" Then
        Call ProcessInput(Mid(txtOutgoing, 2), Me.txtIncoming, Me)
    ElseIf Left(txtOutgoing.Text, 1) = "?" Then
        If Left(LCase(txtOutgoing.Text), 5) = "?save" Then
            If Len(lScriptFile) <> 0 Then
                SaveFile lScriptFile, lScriptData
            Else
                msg = InputBox("Enter Filename:", "NexIRC", "default.nirc")
                If Len(msg) <> 0 Then
                    lScriptFile = App.Path & "\data\scripts\" & msg
                    SaveFile lScriptFile, lScriptData
                End If
            End If
        End If
        If Left(LCase(txtOutgoing.Text), 4) = "?asc" Then
            MsgBox Asc(Right(txtOutgoing.Text, 1))
        End If
        If Left(LCase(txtOutgoing.Text), 6) = "?clear" Then
            lScriptData = ""
            lScriptFile = ""
            ProcessReplaceString sScriptCleared, txtIncoming
        End If
        If Left(LCase(txtOutgoing.Text), 5) = "?view" Then
            If Len(lScriptFile) <> 0 Then
                Set k = New frmTextEditor
                k.Show
                k.txtIncoming.Text = ReadFile(lScriptFile)
                msg = lScriptFile
                msg = GetFileTitle(msg)
                k.Caption = msg
                k.Tag = ""
            Else
                MsgBox lScriptData
            End If
        End If
    ElseIf Len(txtOutgoing.Text) = 0 Then
        'Dim i As Integer
        If Len(lScriptData) <> 0 Then
            DoColorLines txtIncoming, lScriptData
        Else
            'DoColor txtIncoming, "3NexIRC [14Version: " & App.major & "." & App.minor & "3]"
            ProcessReplaceString sVersion, txtIncoming, App.Major, App.Minor
        End If
    Else
        DoColor txtIncoming, "3Script Command [14" & txtOutgoing.Text & "3]"
        DoColorSep txtIncoming
        Set F = New clsFMenu
        m = F.RunCommand(txtOutgoing.Text, Me)
        If m = True Then
            If Len(lScriptData) <> 0 Then
                lScriptData = lScriptData & vbCrLf & txtOutgoing.Text
            Else
                lScriptData = txtOutgoing.Text
            End If
        End If
    End If
    txtOutgoing = ""
    KeyAscii = 0
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtOutgoing_KeyPress(KeyAscii As Integer)"
End Sub

Private Sub txtIncoming_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtOutgoing.SetFocus
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtIncoming_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub tcp_Close()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
lSettings.sAutoJoinActivated = False
mdiNexIRC.picConnect.Visible = True
mdiNexIRC.picDisconnect.Visible = False
SetConnected False
ProcessReplaceString sConnectionClosed, txtIncoming
'DoColor txtIncoming, "2• Connection Closed"
DoColorSep txtIncoming
lMyCurrentModes = ""
Me.Caption = Me.Tag
For i = 1 To ReturnChannelUBound
    UnloadChannel i
    SetChannelName i, ""
Next i
chkMOTDActivated.Value = 0
If lSettings.sReconnectOnDisconnect = True Then tmrReconnect.Enabled = True
DisableIdent
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tcp_Close()"
End Sub

Private Sub tcp_Connect()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String
tmrReconnect.Enabled = False
mdiNexIRC.picConnect.Visible = False
mdiNexIRC.picDisconnect.Visible = True
lMyCurrentModes = ""
SetConnected True
DoColor txtIncoming, "2• Connected" & vbCrLf
msg = Left(lSettings.sEMail, 1) & Parse(lSettings.sEMail, Left(lSettings.sEMail, 1), "@")
msg2 = Parse(lSettings.sEMail, "@", Right(lSettings.sEMail, 2)) & Right(lSettings.sEMail, 2)
tcp.SendData "USER " & msg & " " & Chr(34) & msg2 & Chr(34) & " " & Chr(34) & tcp.LocalIP & Chr(34) & " :" & lSettings.sRealName & vbCrLf
tcp.SendData "NICK " & lSettings.sNickname & vbCrLf
Me.Caption = Me.Tag & ": " & lSettings.sNickname
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tcp_Connect()"
End Sub

Private Sub tcp_DataArrival(ByVal bytesTotal As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim IRCLine() As String, strData As String, i As Integer
'PingReply = GetTickCount
Static RestLine As String
If tcp.State = sckConnected Then tcp.GetData strData
If Len(RestLine) > 0 Then
    strData = RestLine & strData
End If
IRCLine = Split(strData, Chr(10))
If Right(strData, 1) = Chr(10) Or Right(strData, 1) = Chr(13) Then
    RestLine = ""
    For i = 0 To UBound(IRCLine)
        If IRCLine(i) <> "" Then
            IRCLine(i) = Replace(IRCLine(i), Chr(22), "")
            IRCLine(i) = Replace(IRCLine(i), Chr(13), "")
            IRCLine(i) = Replace(IRCLine(i), Chr(10), "")
            ParseIRCData IRCLine(i), Me
        End If
    Next i
Else
    If UBound(IRCLine) <> -1 Then RestLine = IRCLine(UBound(IRCLine))
    For i = 0 To UBound(IRCLine) - 1
        If IRCLine(i) <> "" Then
            IRCLine(i) = Replace(IRCLine(i), Chr(22), "")
            IRCLine(i) = Replace(IRCLine(i), Chr(13), "")
            IRCLine(i) = Replace(IRCLine(i), Chr(10), "")
            ParseIRCData IRCLine(i), Me
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tcp_DataArrival(ByVal bytesTotal As Long)"
End Sub

Private Sub tcp_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ProcessReplaceString sConnectionError, txtIncoming, Description, Trim(CStr(Number))
'DoColor txtIncoming, "2• Connection Closed " & Number & " (" & Description & ")" & vbCrLf
DoColorSep txtIncoming
lMyCurrentModes = ""
Me.Caption = Me.Tag
Dim i As Integer
For i = 1 To ReturnChannelUBound
    UnloadChannel i
    SetChannelName i, ""
Next i
SetConnected False
If lSettings.sReconnectOnDisconnect = True Then tmrReconnect.Enabled = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tcp_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)"
End Sub

