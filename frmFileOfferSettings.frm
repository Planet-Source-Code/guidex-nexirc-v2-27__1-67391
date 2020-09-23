VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmFileOfferSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Audio Server Settings"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFileOfferSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   3345
   StartUpPosition =   2  'CenterScreen
   Begin OsenXPCntrl.OsenXPButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1680
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "frmFileOfferSettings.frx":0CCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdChanFolder 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Channels"
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
      MICON           =   "frmFileOfferSettings.frx":0E2C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chkAudioServer 
      Appearance      =   0  'Flat
      Caption         =   "Audio Server Enabled"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.CheckBox chkFileOfferInChannel 
      Appearance      =   0  'Flat
      Caption         =   "File Offer in Channel"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1935
   End
   Begin VB.CheckBox chkLogAudioDownloads 
      Appearance      =   0  'Flat
      Caption         =   "Log Audio Downloads"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CheckBox chkOfferWhenPlayed 
      Appearance      =   0  'Flat
      Caption         =   "Offer when Played"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CheckBox chkEnableListmedia 
      Appearance      =   0  'Flat
      Caption         =   "Enable List"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CheckBox chkEnableFind 
      Appearance      =   0  'Flat
      Caption         =   "Enable Search"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin OsenXPCntrl.OsenXPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1680
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "frmFileOfferSettings.frx":0F8E
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
Attribute VB_Name = "frmFileOfferSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdCancel_Click()"
End Sub

Private Sub cmdChanFolder_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmChannelFolder.Show
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdChanFolder_Click()"
End Sub

Private Sub cmdOK_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sAudioServer = GetCheckboxValue(chkAudioServer)
lSettings.sFileOfferInChannel = GetCheckboxValue(chkFileOfferInChannel)
lSettings.sOfferWhenPlayed = GetCheckboxValue(chkOfferWhenPlayed)
lSettings.sLogAudioDownloads = GetCheckboxValue(chkLogAudioDownloads)
lSettings.sEnableSearch = GetCheckboxValue(chkEnableFind)
lSettings.sEnableList = GetCheckboxValue(chkEnableListmedia)
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdCancel_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
SetCheckBoxValue chkAudioServer, lSettings.sAudioServer
SetCheckBoxValue chkFileOfferInChannel, lSettings.sFileOfferInChannel
SetCheckBoxValue chkLogAudioDownloads, lSettings.sLogAudioDownloads
SetCheckBoxValue chkOfferWhenPlayed, lSettings.sOfferWhenPlayed
SetCheckBoxValue chkEnableFind, lSettings.sEnableSearch
SetCheckBoxValue chkEnableListmedia, lSettings.sEnableList
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub
