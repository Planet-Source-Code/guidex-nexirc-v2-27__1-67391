VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Object = "{FDFCF4A3-AD96-11D4-9959-0050BACD4F4C}#1.0#0"; "MDec.ocx"
Begin VB.Form frmDecode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Decode"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4530
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDecode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOutputFilename 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txtInputFilename 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MDECLib.MDec MDec1 
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   0
   End
   Begin OsenXPCntrl.OsenXPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   840
      Width           =   735
      _ExtentX        =   1296
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
      MPTR            =   0
      MICON           =   "frmDecode.frx":0CCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   735
      _ExtentX        =   1296
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
      MPTR            =   0
      MICON           =   "frmDecode.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblInputFilename 
      Caption         =   "Input Filename:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Output Filename:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frmDecode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
End Sub

Private Sub cmdOK_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(txtInputFilename.Text) <> 0 And DoesFileExist(txtInputFilename.Text) = True Then
    If Len(txtOutputFilename.Text) <> 0 Then
        If DoesFileExist(txtOutputFilename.Text) = False Then
            MDec1.OPENFILENAME = txtInputFilename.Text
            MDec1.savefilename = txtOutputFilename.Text
            MDec1.Decode
        Else
            If lSettings.sGeneralPrompts = True Then
                MsgBox "File " & txtOutputFilename.Text & " exists.", vbExclamation
            End If
        End If
    End If
End If
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
SetButtonType cmdOK
SetButtonType cmdCancel
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub MDec1_PercentDone(ByVal nPercent As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If nPercent <> 100 Then
    If ProgressBar1.Value <> nPercent Then
        ProgressBar1.Value = nPercent
    End If
Else
    If lSettings.sGeneralPrompts = True Then MsgBox "Decode complete", vbInformation
    Unload Me
End If
End Sub
