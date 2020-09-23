VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmChannel 
   Caption         =   "NexIRC - Channel"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChannel.frx":0000
   LinkTopic       =   "frmChannel"
   MDIChild        =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   7215
   Begin nexIRC.ctlListView lvwNames 
      Height          =   1455
      Left            =   3600
      TabIndex        =   4
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2566
   End
   Begin MSComctlLib.ImageList imgChanIcons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannel.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannel.frx":0360
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannel.frx":06B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannel.frx":0A08
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstSent 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      IntegralHeight  =   0   'False
      ItemData        =   "frmChannel.frx":0D5C
      Left            =   0
      List            =   "frmChannel.frx":0D5E
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtOutgoing 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   4815
   End
   Begin VB.TextBox txtTopic 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin nexIRC.ctlTBox txtIncoming 
      Height          =   1455
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2566
   End
   Begin VB.Line linTextSep 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   4800
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line linTextSep2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   3480
      X2              =   3480
      Y1              =   1440
      Y2              =   0
   End
End
Attribute VB_Name = "frmChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cFM As clsFMenu
Private lItemSelected As String

Public Function ReturnSelectedItem() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnSelectedItem = lItemSelected
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnSelectedItem() As String"
End Function

Public Sub ActivateResize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Form_Resize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ActivateResize()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtIncoming.SetTag "channel"
Me.Icon = mdiNexIRC.Icon
lvwNames.Initialize
lvwNames.BorderStyle = bsThin
lvwNames.ViewMode = vmDetails
lvwNames.ColumnAdd 0, "Items", 447, [caLeft]
lvwNames.Font.Name = "Tahoma"
lvwNames.FullRowSelect = True
lvwNames.HeaderFlat = False
lvwNames.HeaderHide = False
'lvwNames.Font = "Tahoma"
With lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex)
    If Len(.sBackColor) <> 0 Then Me.BackColor = .sBackColor
    If Len(.sBackColor) <> 0 Then txtIncoming.SetBackColor .sBackColor
    If Len(.sBackColor) <> 0 Then txtOutgoing.BackColor = .sBackColor
    If Len(.sTextColor) <> 0 Then txtOutgoing.ForeColor = .sTextColor
    If Len(.sBackColor) <> 0 Then lvwNames.BackColor = .sBackColor
    If Len(.sTextColor) <> 0 Then lvwNames.ForeColor = .sTextColor
    'CHECKED
    If Len(.sBackColor) <> 0 Then lstSent.BackColor = .sBackColor
    If Len(.sTextColor) <> 0 Then lstSent.ForeColor = .sTextColor
End With
lSettings.sChannelCount = lSettings.sChannelCount + 1
If lSettings.sBorderlessObjects = True Then
    txtOutgoing.BorderStyle = 0
    txtIncoming.SetBorderStyle True
    'lvwNames.BorderStyle = ccNone
    lvwNames.BorderStyle = bsNone
Else
    txtOutgoing.BorderStyle = 1
    txtIncoming.SetBorderStyle False
    
    'lvwNames.BorderStyle = ccFixedSingle
    lvwNames.BorderStyle = bsThin
End If
Call Form_Resize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If txtTopic.Visible = True Then txtTopic.Visible = False
If txtIncoming.Left <> 0 Then txtIncoming.Left = 0
If Me.ScaleHeight <> 0 Then txtIncoming.Height = Me.ScaleHeight - txtOutgoing.Height - 20
If txtOutgoing.Left <> 0 Then txtOutgoing.Left = 0
If Me.ScaleWidth <> 0 Then txtIncoming.Width = Me.ScaleWidth - lvwNames.Width
If Me.ScaleWidth <> 0 Then txtTopic.Width = Me.ScaleWidth - lvwNames.Width
lvwNames.Left = txtTopic.Width + 20
If Me.ScaleHeight <> 0 Then lvwNames.Height = Me.ScaleHeight - txtOutgoing.Height - 40
txtOutgoing.Width = Me.ScaleWidth
txtOutgoing.Top = Me.ScaleHeight - txtOutgoing.Height
lstSent.Width = txtIncoming.Width
lstSent.Height = txtIncoming.Height
lstSent.Top = txtIncoming.Top - 20
lstSent.Left = txtIncoming.Left = 20
linTextSep.x2 = Me.ScaleWidth
linTextSep.y1 = Me.txtIncoming.Height
linTextSep.y2 = Me.txtIncoming.Height
linTextSep2.x1 = txtIncoming.Width
linTextSep2.x2 = txtIncoming.Width
linTextSep2.y1 = 0
linTextSep2.y2 = lvwNames.Height
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Resize()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, word() As String
lSettings.sChannelCount = lSettings.sChannelCount - 1
word = Split(Me.Caption, " ")
If lSettings.sActiveServerForm.tcp.State = sckConnected Then lSettings.sActiveServerForm.tcp.SendData "PART " & word(0) & vbCrLf
For i = 1 To ReturnChannelUBound
    If LCase(ReturnChannelName(i)) = LCase(word(0)) Then
    'If LCase(lChannelName(i)) = LCase(word(0)) Then
        SetChannelName i, ""
        SetChannelModes i, ""
        'lChannelName(i) = ""
        'lChannelModes(i) = ""
        'lChannelTopic(i) = ""
        Call RemoveTaskbar(word(0))
    End If
Next i
'Set lvwNames = Nothing
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub lvwNames_ItemClick(Item As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next

lItemSelected = lvwNames.ItemText(Item)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lvwNames_ItemClick(Item As Integer)"
End Sub

'Private Sub lvwNames_DblClick()
'Dim colorX(0 To 15) As Long, ChatWith As String, i As Integer, xlfound As Boolean
'colorX(0) = vbWhite
'colorX(1) = vbBlack
'colorX(2) = RGB(0, 0, 140)
'colorX(3) = RGB(0, 140, 0)
'colorX(4) = vbRed
'colorX(5) = RGB(110, 65, 0)
'colorX(6) = RGB(140, 0, 140)
'colorX(7) = RGB(248, 146, 0)
'colorX(8) = vbYellow
'colorX(9) = vbGreen
'colorX(10) = RGB(0, 140, 140)
'colorX(11) = RGB(0, 255, 255)
'colorX(12) = vbBlue
'colorX(13) = vbMagenta
'colorX(14) = RGB(140, 140, 140)
'colorX(15) = RGB(200, 200, 200)
'xlfound = False



'ChatWith = lvwNames.ListItems(lvwNames.SelectedItem.Index).Text
'ChatWith = Replace(ChatWith, "@", "")
'ChatWith = Replace(ChatWith, "+", "")
'ChatWith = Replace(ChatWith, "%", "")
'For i = 1 To 150
'    If LCase(ReturnQueryName(i)) = LCase(ChatWith) Then
    'If LCase(lQueryName(i)) = LCase(ChatWith) Then
'        xlfound = True
'        Exit For
'    End If
'Next i
'If xlfound = False Then
'    For i = 1 To 150
        
'        If ReturnQueryName(i) = "" Then
        'If lQueryName(i) = "" Then
            'Load lQuery(i)
            'lQuery(i).txtOutgoing.BackColor = colorX(Color.BGText)
            'lQuery(i).txtOutgoing.ForeColor = colorX(Color.Normal)
            'lQuery(i).txtIncoming.SetBackColor Str(colorX(Color.BGText))
            'lQuery(i).Caption = ChatWith
            'lQueryName(i) = ChatWith
            'Call AddTaskPanel(ChatWith, 1)
'            Exit For
'        End If
'    Next i
'End If
'If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lstNames_DblClick()"
'End Sub

Private Sub lvwNames_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lTopCorrection As Integer, lLeftCorrection As Integer
If Button = 2 Then
    If mdiNexIRC.picTopToolbar.Visible = True Then
        lTopCorrection = (mdiNexIRC.picTopToolbar.Height / Screen.TwipsPerPixelY)
    End If
    lTopCorrection = lTopCorrection + (mdiNexIRC.Top / Screen.TwipsPerPixelY) + (Me.Top / Screen.TwipsPerPixelY) + 81
    If mdiNexIRC.mnuFile.Visible = False Then
        lTopCorrection = lTopCorrection - 20
    End If
    If mdiNexIRC.picMobileMixer.Visible = True Then
        lLeftCorrection = (mdiNexIRC.picMobileMixer.Width / Screen.TwipsPerPixelX)
    End If
    lLeftCorrection = lLeftCorrection + (mdiNexIRC.Left / Screen.TwipsPerPixelX) + (Me.Left / Screen.TwipsPerPixelX) + (txtIncoming.Width / Screen.TwipsPerPixelX) + 10
    If Button = 2 Then
        If Button = 2 Then
            If DoesFileExist(GetINIFile(iNicklistMenu)) = True Then
                Set cFM = New clsFMenu
                With cFM
                    .OwnerHWND = Me.hWnd
                    Call .LoadMenus(GetINIFile(iNicklistMenu))
                    Call .ShowMenu((X / Screen.TwipsPerPixelX) + lLeftCorrection, (Y / Screen.TwipsPerPixelY) + lTopCorrection, mdiNexIRC.ActiveForm)
                End With
                Set cFM = Nothing
            End If
        End If
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lstNames_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub lstSent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtOutgoing.Text = lstSent.Text
lstSent.Visible = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lstSent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
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
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtOutgoing_GotFocus()"
End Sub

Private Sub txtOutgoing_KeyDown(KeyCode As Integer, Shift As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = txtOutgoing.Text
If ProcessKeyDown(KeyCode, Shift, msg, lstSent, txtOutgoing) = True Then
    KeyCode = 0
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtOutgoing_KeyDown(KeyCode As Integer, Shift As Integer)"
End Sub

Private Sub txtOutgoing_KeyPress(KeyAscii As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim word() As String, starttext As Integer, msg As String, i As Integer
word = Split(Me.Caption, Chr(32))
If KeyAscii = 13 Then
    txtOutgoing.Text = LTrim(txtOutgoing)
    If Left(txtOutgoing.Text, 1) = "/" Then
        SetActChannel LCase(word(0))
        Call ProcessInput(Mid(txtOutgoing.Text, 2), Me.txtIncoming, lSettings.sActiveServerForm)
    Else
        If Len(txtOutgoing.Text) <> 0 Then
            If lSettings.sActiveServerForm.tcp.State = sckConnected Then lSettings.sActiveServerForm.tcp.SendData "PRIVMSG " & word(0) & " :" & txtOutgoing.Text & vbCrLf
            DoColor txtIncoming, "" & Color.Normal & "<" & Color.Action & "" & lSettings.sNickname & "" & Color.Normal & "> " & txtOutgoing.Text
        End If
    End If
    txtOutgoing = ""
    KeyAscii = 0
End If
If KeyAscii = 27 Then
    Me.WindowState = vbMinimized
End If
If KeyAscii = 11 Then
    starttext = txtOutgoing.SelStart
    txtOutgoing = Mid(txtOutgoing, 1, txtOutgoing.SelStart) & "" & Mid(txtOutgoing, txtOutgoing.SelStart + 1)
    txtOutgoing.SelStart = starttext + 1
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtOutgoing_KeyPress(KeyAscii As Integer)"
End Sub

Private Sub txtOutgoing_KeyUp(KeyCode As Integer, Shift As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer, F As Integer, l As Integer, msg2 As String
If lSettings.sUseNickCompletor = True Then
    If Left(txtOutgoing.Text, 1) = "\" Then
        msg = Right(txtOutgoing.Text, Len(txtOutgoing.Text) - 1)
        If Len(msg) <> 0 Then
            For i = 1 To lvwNames.Count
    '        For i = 1 To lvwNames.ListItems.Count
                If Len(lvwNames.ItemText(i)) <> 0 Then
    '            If Len(lvwNames.ListItems(i).Text) <> 0 Then
                    msg2 = lvwNames.ItemText(i)
    '                msg2 = lvwNames.ListItems(i).Text
                    If Left(msg2, 1) = "@" Or Left(msg2, 1) = "+" Then msg2 = Right(msg2, Len(msg2) - 1)
                    If LCase(msg) = LCase(Left(msg2, Len(msg))) Then
                        If F = 0 Then l = i
                        F = F + 1
                    End If
                End If
            Next i
        End If
        If F = 1 Then
            
            msg2 = lvwNames.ItemText(l)
            'msg2 = lvwNames.ListItems(l).Text
            If Left(msg2, 1) = "@" Or Left(msg2, 1) = "+" Then msg2 = Right(msg2, Len(msg2) - 1)
            txtOutgoing.Text = ReturnStringDataByType(sNickCompletor1) & msg2 & ReturnStringDataByType(sNickCompletor2) & " "
            txtOutgoing.SelStart = Len(msg2) + 3
        End If
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtOutgoing_KeyUp(KeyCode As Integer, Shift As Integer)"
End Sub

Private Sub txtTopic_KeyPress(KeyAscii As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If KeyAscii = 13 Then
    lSettings.sActiveServerForm.tcp.SendData "TOPIC " & Me.Caption & " :" & txtTopic.Text & vbCrLf
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtTopic_KeyPress(KeyAscii As Integer)"
End Sub
