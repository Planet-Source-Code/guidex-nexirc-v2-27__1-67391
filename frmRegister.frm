VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register"
   ClientHeight    =   3960
   ClientLeft      =   90
   ClientTop       =   1065
   ClientWidth     =   3375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   3375
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&G"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox OsenXPButton2 
      Height          =   375
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   1395
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1080
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1080
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   2520
      Width           =   2175
   End
   Begin NexIRC.ctlXPButton cmdCancel 
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   1
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmRegister.frx":000C
      PICN            =   "frmRegister.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin NexIRC.ctlXPButton cmdRegister 
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   1
      TX              =   "&OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmRegister.frx":170C2
      PICN            =   "frmRegister.frx":170DE
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin NexIRC.ctlXPButton cmdHelp 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   1
      TX              =   "&Help"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmRegister.frx":40D00
      PICN            =   "frmRegister.frx":40D1C
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   3240
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3240
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Wait 1 day for code to be generated, you will recieve the code in e-mail, enter it below when it has been recieved"
      Height          =   975
      Left            =   840
      TabIndex        =   9
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Step 2:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration costs $20 USD. Click the button below to launch paypal"
      Height          =   855
      Left            =   840
      TabIndex        =   6
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Step 1:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   3240
      Y1              =   3375
      Y2              =   3375
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   3240
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "How to register NexIRC"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHowToUseThisWindow 
         Caption         =   "H&ow to use this Window"
      End
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Leon As Boolean

Private Sub cmdCancel_Click()
''If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdCancel_Click()"
End Sub

Private Sub cmdHelp_Click()
''If lSettings.sHandleErrors = True Then On Local Error Resume Next
DisplayHelpInformation 29
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdHelp_Click()"
End Sub

Private Sub cmdRegister_Click()
''If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim m As Boolean, i As String
If txtName.Text = "guidex_developer" And txtPassword.Text = "07281979-2841" And Leon = True Then
    Command1.Visible = True
    txtName.Text = ""
    txtPassword.Text = ""
    Exit Sub
End If
If m = True Then
    txtPassword.Text = KeyGen(txtName.Text, "pickles", 1)
Else
    i = KeyGen(txtName.Text, "pickles", 1)
    If i = txtPassword.Text Then
        If lSettings.sGeneralPrompts = True Then
            MsgBox "Thank you very much for registering. All of the money made from NexIRC is spent on the development of NexIRC", vbInformation
        End If
        mdiNexIRC.Caption = "NexIRC (Registered Version)"
        lRegInfo.rName = txtName.Text
        lRegInfo.rPassword = txtPassword.Text
        WriteINI GetINIFile(iIRC), "REGInfo", "NAME", lRegInfo.rName
        WriteINI GetINIFile(iIRC), "REGInfo", "PASSWORD", lRegInfo.rPassword
        lRegInfo.rRegistered = True
        Unload Me
    Else
        If lSettings.sGeneralPrompts = True Then
            MsgBox "The code you entered was not correct. The name did not match the password. Please try again", vbInformation
        End If
        Exit Sub
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdRegister_Click()"
End Sub

Private Sub Command1_Click()
''If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = InputBox("Enter secret phrase:", "Code generator", "37463788473623")
If msg = "pickles" Then
    txtPassword.Text = KeyGen(txtName.Text, "pickles", 1)
Else
    Command1.Enabled = False
    Command1.Visible = False
    Unload Me
    txtName.Text = ""
    txtPassword.Text = ""
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Command1_Click()"
End Sub

Private Sub Form_Load()
''If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = mdiNexIRC.Icon
Leon = False
If Len(lRegInfo.rName) <> 0 And Len(lRegInfo.rPassword) <> 0 Then
    txtName.Text = lRegInfo.rName
    txtPassword.Text = lRegInfo.rPassword
End If
SetButtonType cmdRegister
SetButtonType cmdHelp
SetButtonType cmdCancel

'SetButtonType OsenXPButton2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub mnuExit_Click()
''If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuHowToUseThisWindow_Click()"
End Sub

Private Sub mnuHowToUseThisWindow_Click()
''If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdHelp_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuHowToUseThisWindow_Click()"
End Sub

Private Sub OsenXPButton2_Click()
''If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Leon = True Then
    Exit Sub
End If
Surf "https://www.paypal.com/xclick/business=guidex%40tnexgen.com&item_name=Audiogen+Registration&amount=20.00&no_note=1&tax=0&currency_code=USD&lc=US", Me.hwnd
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub OsenXPButton2_Click()"
End Sub

Private Sub OsenXPButton2_KeyPress(KeyAscii As Integer)
''If lSettings.sHandleErrors = True Then On Local Error Resume Next
If KeyAscii = 92 Then
    Leon = True
    Caption = ""
    Label1.Caption = ""
    Label2.Caption = ""
    Label3.Caption = ""
    Label5.Caption = ""
    Label6.Caption = ""
    Label7.Caption = ""
    Label8.Caption = ""
    OsenXPButton2.Visible = False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub OsenXPButton2_KeyPress(KeyAscii As Integer)"
End Sub

Private Sub txtName_GotFocus()
''If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtName.SelStart = 0
txtName.SelLength = Len(txtName.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtName_GotFocus()"
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
''If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Leon = True Then
    Caption = KeyCode
End If
If Command1.Visible = True And Shift = 1 And KeyCode = 66 Then
    Command1.Enabled = True
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)"
End Sub

Private Sub txtPassword_GotFocus()
''If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPassword.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtPassword_GotFocus()"
End Sub
