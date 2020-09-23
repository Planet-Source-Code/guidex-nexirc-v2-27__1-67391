VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmColorEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Color Editor"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmColorEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   3825
   StartUpPosition =   1  'CenterOwner
   Begin OsenXPCntrl.OsenXPButton cmdHelp 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Help"
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
      MPTR            =   99
      MICON           =   "frmColorEditor.frx":0CCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   2640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Cancel"
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
      MPTR            =   99
      MICON           =   "frmColorEditor.frx":0E2C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstObjects 
      BackColor       =   &H00FFFFFF&
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.OptionButton optColorType 
      Caption         =   "RGB"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optColorType 
      Caption         =   "QBColors"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtRetValue 
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin OsenXPCntrl.OsenXPButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   2640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "OK"
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
      MPTR            =   99
      MICON           =   "frmColorEditor.frx":0F8E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdForeground 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Foreground"
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
      MPTR            =   99
      MICON           =   "frmColorEditor.frx":10F0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdBackColor 
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Background"
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
      MPTR            =   99
      MICON           =   "frmColorEditor.frx":1252
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdWindowColors 
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Settings"
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
      MPTR            =   99
      MICON           =   "frmColorEditor.frx":13B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmColorEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lBackground As Boolean

Private Sub cmdBackColor_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
If optColorType(0).Value = True Then
    lBackground = True
    frmColorPickor.Show 0, Me
Else
    msg = InputBox("Enter RGB Value:")
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdBackColor_Click()"
End Sub

Private Sub cmdCancel_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdCancel_Click()"
End Sub

Private Sub cmdHelp_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DisplayHelpInformation 13
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdHelp_Click()"
End Sub

Private Sub cmdOK_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdOK_Click()"
End Sub

Private Sub cmdForeground_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lBackground = False
frmColorPickor.Show 0, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdForeground_Click()"
End Sub

Private Sub cmdWindowColors_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
frmOptions.Show 0, Me
For i = 0 To 8
    frmOptions.fraSettings(i).Visible = False
Next i
frmOptions.optCheck(4).Value = True
frmOptions.fraSettings(4).Visible = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdWindowColors_Click()"
End Sub

Private Sub Form_Load()
SetButtonType cmdHelp
SetButtonType cmdOK
SetButtonType cmdCancel
SetButtonType cmdBackColor
SetButtonType cmdForeground
SetButtonType cmdWindowColors
End Sub

Private Sub txtRetValue_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case LCase(lstObjects.Text)
Case "visfxactivecolor1"
    If lBackground = True Then
        mdiMain.ActiveForm.level1.ActiveColor1 = QBColor(Int(txtRetValue.Text))
    End If
Case "visfxactivecolor2"
    If lBackground = True Then
        mdiMain.ActiveForm.level1.ActiveColor2 = QBColor(Int(txtRetValue.Text))
    End If
Case "visfxactivecolor3"
    If lBackground = True Then
        mdiMain.ActiveForm.level1.ActiveColor3 = QBColor(Int(txtRetValue.Text))
    End If
Case "visfxinactivecolor1"
    If lBackground = True Then
        mdiMain.ActiveForm.level1.InactiveColor1 = QBColor(Int(txtRetValue.Text))
    End If
Case "visfxinactivecolor2"
    If lBackground = True Then
        mdiMain.ActiveForm.level1.InactiveColor2 = QBColor(Int(txtRetValue.Text))
    End If
Case "visfxinactivecolor3"
    If lBackground = True Then
        mdiMain.ActiveForm.level1.InactiveColor3 = QBColor(Int(txtRetValue.Text))
    End If
Case "outgoing"
    If lBackground = True Then
        mdiMain.ActiveForm.txtSend.BackColor = QBColor(Int(txtRetValue.Text))
    Else
        mdiMain.ActiveForm.txtSend.ForeColor = QBColor(Int(txtRetValue.Text))
    End If
Case "chat"
    If lBackground = True Then
        mdiMain.ActiveForm.txtStatus.BackColor = QBColor(Int(txtRetValue.Text))
    Else
        'mdiMain.ActiveForm.txtStatus.ForeColor = QBColor(Int(txtRetValue.Text))
    End If
End Select
If Len(mdiMain.ActiveForm.Tag) = 0 Then
    SaveWindowSettings mdiMain.ActiveForm.Caption, mdiMain.ActiveForm
Else
    SaveWindowSettings mdiMain.ActiveForm.Tag, mdiMain.ActiveForm
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtRetValue_Change()"
End Sub
