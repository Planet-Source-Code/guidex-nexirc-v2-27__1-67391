VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin nexIRC.ctlListView lvwLoading 
      Height          =   3495
      Left            =   1920
      TabIndex        =   0
      Top             =   2280
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6165
   End
   Begin nexIRC.XP_ProgressBar XP_ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   6956042
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
   Begin VB.Image imgSplash 
      Height          =   4230
      Left            =   5880
      Picture         =   "frmSplash1.frx":000C
      Top             =   5880
      Visible         =   0   'False
      Width           =   6750
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lStopUnload As Boolean

Public Sub AddToInfo(lIndex As Integer, lText As String, lSubItem As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lIndex = 0
If Len(lText) <> 0 Then
    lblStatus.Caption = lText
    'lvwLoading.ItemAdd lIndex, lText, 0, 0
End If
'lvwLoading.ItemSelected(lIndex) = True
DoEvents
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = mdiNexIRC.Icon
mdiNexIRC.SetSplashVisible True
XP_ProgressBar1.Scrolling = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sProgressBarStyle
SetProgressBarColor XP_ProgressBar1
lvwLoading.Initialize
lvwLoading.BorderStyle = bsThin
lvwLoading.ViewMode = vmDetails
lvwLoading.ColumnAdd 0, "Startup Item", 450, [caLeft]
'lvwLoading.ColumnAdd 1, "Progress", 70, [caLeft]
lvwLoading.Font.Name = "Tahoma"
lvwLoading.FullRowSelect = True
lvwLoading.HeaderFlat = False
lvwLoading.HeaderHide = False
lvwLoading.Font = "Tahoma"
'-
'lvwServers.Initialize
'lvwServers.BorderStyle = bsThin
'lvwServers.ViewMode = vmDetails
'lvwServers.ColumnAdd 0, "Server", 100, [caLeft]
'lvwServers.ColumnAdd 1, "Port", 50, [caLeft]
'lvwServers.Font.Name = "Tahoma"
'lvwServers.FullRowSelect = True
'lvwServers.HeaderFlat = False
'lvwServers.HeaderHide = False
'lvwServers.Font = "Tahoma"
'-
'lvwUpdates.Initialize
'lvwUpdates.BorderStyle = bsThin
'lvwUpdates.ViewMode = vmDetails
'lvwUpdates.ColumnAdd 0, "Server", 487, [caLeft]
'lvwUpdates.Font.Name = "Tahoma"
'lvwUpdates.FullRowSelect = True
'lvwUpdates.HeaderFlat = False
'lvwUpdates.HeaderHide = False
'lvwUpdates.Font = "Tahoma"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
mdiNexIRC.SetSplashVisible False
End Sub

Private Sub tabSplash_Click()
lStopUnload = True
End Sub
