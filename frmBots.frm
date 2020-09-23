VERSION 5.00
Begin VB.Form frmBots 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Bot Control"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   630
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
   Icon            =   "frmBots.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4965
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtValue2 
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox txtValue1 
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Top             =   840
      Width           =   3735
   End
   Begin VB.ComboBox cboCommand 
      Height          =   315
      ItemData        =   "frmBots.frx":000C
      Left            =   1080
      List            =   "frmBots.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   3735
   End
   Begin VB.ComboBox cboNickname 
      Height          =   315
      ItemData        =   "frmBots.frx":0010
      Left            =   1080
      List            =   "frmBots.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.CheckBox chkSaveToAutoPerform 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
   Begin nexIRC.ctlXPButton cmdHelp 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Help"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmBots.frx":0014
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmBots.frx":0030
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmBots.frx":004C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblValue2 
      Caption         =   "Value 2:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblValue1 
      Caption         =   "Value 1:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblCommand 
      Caption         =   "Command:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblNickname 
      Caption         =   "&Nickname:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowToUseThisWindow 
         Caption         =   "H&ow to use this Window"
      End
   End
End
Attribute VB_Name = "frmBots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lValue1Caption As String
Dim lValue2Caption As String

Private Sub cboCommand_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim b As Integer, msg As String, msg2 As String, msg3 As String
b = FindBotIndex(cboNickname.Text)
If b <> 0 Then
    txtValue1.Text = ""
    txtValue2.Text = ""
    lValue1Caption = ""
    lValue2Caption = ""
    txtValue1.PasswordChar = ""
    msg = Parse(cboCommand.Text, "$", Right(cboCommand.Text, 3)) & Right(cboCommand.Text, 3)
    If InStr(msg, "$") Then
        lblValue2.Visible = True
        msg2 = Parse(msg, "$", Right(msg, 3)) & Right(msg, 3)
        msg = Left(msg, 1) & Parse(msg, Left(msg, 1), " ")
        lblValue1.Caption = msg & ": "
        lblValue2.Caption = msg2 & ": "
        txtValue2.Visible = True
        lValue1Caption = Trim(msg)
        lValue2Caption = Trim(msg2)
        msg3 = ReadINI(GetINIFile(iBots), "Bot " & Trim(Str(b)), msg, "")
        If Len(msg3) <> 0 Then txtValue1.Text = msg3
        msg3 = ReadINI(GetINIFile(iBots), "Bot " & Trim(Str(b)), msg2, "")
        If Len(msg3) <> 0 Then txtValue2.Text = msg3
        If Trim(LCase(lValue1Caption)) = "password" Then txtValue1.PasswordChar = "*"
        If Trim(LCase(lValue2Caption)) = "password" Then txtValue2.PasswordChar = "*"
    Else
        lblValue1.Caption = msg & ": "
        txtValue2.Visible = False
        lblValue2.Visible = False
        lValue1Caption = Trim(msg)
        lValue2Caption = ""
        txtValue1.Text = ReadINI(GetINIFile(iBots), "Bot " & Trim(Str(b)), msg, "")
        If Trim(LCase(lValue1Caption)) = "password" Then txtValue1.PasswordChar = "*"
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cboCommand_Click()"
End Sub

Private Sub cboNickname_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, F As Integer
cboCommand.Clear
i = FindBotIndex(cboNickname.Text)
If Len(ReturnBotNickname(i)) <> 0 Then
    FillComboWithBotCommands cboCommand, ReturnBotType(i)
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdCancel_Click()"
End Sub

Private Sub cmdCancel_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdCancel_Click()"
End Sub

Private Sub cmdHelp_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DisplayHelpInformation 7
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdHelp_Click()"
End Sub

Private Sub cmdOK_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer
If Len(cboNickname.Text) <> 0 And Len(cboCommand.Text) <> 0 Then
    i = FindBotIndex(cboNickname.Text)
    msg = cboCommand.Text
    If Len(msg) <> 0 And i <> 0 Then
        If Len(lValue1Caption) <> 0 Then
            If Len(txtValue1.Text) <> 0 Then
                msg = Replace(msg, "$" & lValue1Caption, txtValue1.Text)
            End If
        End If
        If Len(lValue2Caption) Then
            If Len(txtValue2.Text) <> 0 Then
                msg = Replace(msg, "$" & lValue2Caption, txtValue2.Text)
            End If
        End If
        WriteINI GetINIFile(iBots), "Bot " & Trim(Str(i)), Left(lblValue1.Caption, Len(lblValue1.Caption) - 2), txtValue1.Text
        If Len(lblValue2.Caption) <> 0 Then WriteINI GetINIFile(iBots), "Bot " & Trim(Str(i)), Left(lblValue2.Caption, Len(lblValue2.Caption) - 2), txtValue2.Text
        If chkSaveToAutoPerform.Value = 1 Then
            PerformBotCommand lSettings.sActiveServerForm, i, msg, txtValue1.Text, txtValue2.Text, True
        Else
            PerformBotCommand lSettings.sActiveServerForm, i, msg, txtValue1.Text, txtValue2.Text, False
        End If
    End If
End If
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdOK_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
Me.Icon = mdiNexIRC.Icon
SetButtonType cmdOK
SetButtonType cmdCancel
SetButtonType cmdHelp
For i = 0 To ReturnBotCount
    If Len(ReturnBotNickname(i)) <> 0 Then
        cboNickname.AddItem ReturnBotNickname(i)
    End If
Next i
If cboNickname.ListCount <> 0 Then cboNickname.ListIndex = 0
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub mnuExit_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuExit_Click()"
End Sub

Private Sub mnuHowToUseThisWindow_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdHelp_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuHowToUseThisWindow_Click()"
End Sub
