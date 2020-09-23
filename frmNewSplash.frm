VERSION 5.00
Object = "{1ABC71B2-B0F7-4C1D-9870-3DED8934B20B}#1.0#0"; "prjXTab.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NexIRC"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjXTab.XTab tabSplash 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6800
      TabCount        =   4
      TabCaption(0)   =   "Loading"
      TabContCtrlCnt(0)=   2
      Tab(0)ContCtrlCap(1)=   "lvwLoading"
      Tab(0)ContCtrlCap(2)=   "XP_ProgressBar1"
      TabCaption(1)   =   "Setup"
      TabContCtrlCnt(1)=   14
      Tab(1)ContCtrlCap(1)=   "cboNetwork"
      Tab(1)ContCtrlCap(2)=   "lvwServers"
      Tab(1)ContCtrlCap(3)=   "cmdDefaults"
      Tab(1)ContCtrlCap(4)=   "chkShowWizard"
      Tab(1)ContCtrlCap(5)=   "chkShowSplashOnStartup"
      Tab(1)ContCtrlCap(6)=   "chkShowOptionsOnStartup"
      Tab(1)ContCtrlCap(7)=   "chkConnectOnStartup"
      Tab(1)ContCtrlCap(8)=   "txtNickname"
      Tab(1)ContCtrlCap(9)=   "txtEMail"
      Tab(1)ContCtrlCap(10)=   "txtRealname"
      Tab(1)ContCtrlCap(11)=   "Shape1"
      Tab(1)ContCtrlCap(12)=   "Label2"
      Tab(1)ContCtrlCap(13)=   "Label1"
      Tab(1)ContCtrlCap(14)=   "lblNickName"
      TabCaption(2)   =   "Updates"
      TabContCtrlCnt(2)=   1
      Tab(2)ContCtrlCap(1)=   "lvwUpdates"
      TabCaption(3)   =   "Credits"
      TabContCtrlCnt(3)=   10
      Tab(3)ContCtrlCap(1)=   "Image3"
      Tab(3)ContCtrlCap(2)=   "lblLeonAiossa"
      Tab(3)ContCtrlCap(3)=   "Label14"
      Tab(3)ContCtrlCap(4)=   "Label15"
      Tab(3)ContCtrlCap(5)=   "Label16"
      Tab(3)ContCtrlCap(6)=   "Label17"
      Tab(3)ContCtrlCap(7)=   "Label18"
      Tab(3)ContCtrlCap(8)=   "Label19"
      Tab(3)ContCtrlCap(9)=   "Label20"
      Tab(3)ContCtrlCap(10)=   "Label21"
      TabStyle        =   1
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   -2147483627
      Begin NexIRC.ctlListView lvwUpdates 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   25
         Top             =   480
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4895
      End
      Begin VB.ComboBox cboNetwork 
         Height          =   315
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   480
         Width           =   2295
      End
      Begin NexIRC.ctlListView lvwServers 
         Height          =   2895
         Left            =   -69840
         TabIndex        =   23
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   5106
      End
      Begin NexIRC.ctlListView lvwLoading 
         Height          =   2775
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4895
      End
      Begin NexIRC.isButton cmdDefaults 
         Height          =   330
         Left            =   -71040
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         Icon            =   "frmNewSplash.frx":0000
         Style           =   8
         Caption         =   "Defaults"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin VB.CheckBox chkShowWizard 
         Appearance      =   0  'Flat
         Caption         =   "Show 'Setup &Wizard'"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   12
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CheckBox chkShowSplashOnStartup 
         Appearance      =   0  'Flat
         Caption         =   "Show '&Splash'"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   11
         Top             =   2040
         Width           =   2655
      End
      Begin VB.CheckBox chkShowOptionsOnStartup 
         Appearance      =   0  'Flat
         Caption         =   "Show '&Customize'"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   10
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox chkConnectOnStartup 
         Appearance      =   0  'Flat
         Caption         =   "Connect to &IRC"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   9
         Top             =   1560
         Width           =   4215
      End
      Begin VB.TextBox txtNickname 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   230
         Left            =   -73905
         MaxLength       =   9
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         ToolTipText     =   "8"
         Top             =   495
         Width           =   3975
      End
      Begin VB.TextBox txtEMail 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   230
         Left            =   -73905
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         ToolTipText     =   "10"
         Top             =   735
         Width           =   3975
      End
      Begin VB.TextBox txtRealname 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   230
         Left            =   -73905
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Top             =   975
         Width           =   3975
      End
      Begin NexIRC.XP_ProgressBar XP_ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   3360
         Width           =   7335
         _ExtentX        =   12938
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
         BrushStyle      =   0
         Color           =   4210752
         Scrolling       =   9
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1560
         Left            =   -74880
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1875
      End
      Begin VB.Label lblLeonAiossa 
         BackStyle       =   0  'Transparent
         Caption         =   "(Leon Aiossa)"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Programming:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -72840
         TabIndex        =   20
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Leon J Aiossa, Jamie Cabral"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -71160
         TabIndex        =   19
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Idea's/Inspiration:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -72840
         TabIndex        =   18
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Tim Hueller, Jamie Cabral, Amy Fleischhacker"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -71160
         TabIndex        =   17
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Graphics:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -72840
         TabIndex        =   16
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Leon J Aiossa, Colin Foss"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -71160
         TabIndex        =   15
         Top             =   960
         Width           =   3735
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Additional Credits:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -72840
         TabIndex        =   14
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Abby Ganyaw, Janis Perry, Eric Bishop, Clete Lindsay, 'Crossroad', John Scholten, Dasmius, SupraX, Magique"
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   -71160
         TabIndex        =   13
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         Height          =   750
         Left            =   -73920
         Top             =   480
         Width           =   4005
      End
      Begin VB.Label Label2 
         Caption         =   "&Real Name:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblNickName 
         Caption         =   "Nickname:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
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
''If lSettings.sHandleErrors = True Then On Local Error Resume Next
lIndex = 0
If Len(lText) <> 0 Then
    lvwLoading.ItemAdd lIndex, lText, 0, 0
End If
lvwLoading.ItemSelected(lIndex) = True
DoEvents
End Sub

Private Sub Form_Load()
''If lSettings.sHandleErrors = True Then On Local Error Resume Next
mdiNexIRC.SetSplashVisible True
lvwLoading.Initialize
lvwLoading.BorderStyle = bsThin
lvwLoading.ViewMode = vmDetails
lvwLoading.ColumnAdd 0, "Startup Item", 416, [caLeft]
lvwLoading.ColumnAdd 1, "Progress", 70, [caLeft]
lvwLoading.Font.Name = "Tahoma"
lvwLoading.FullRowSelect = True
lvwLoading.HeaderFlat = False
lvwLoading.HeaderHide = False
lvwLoading.Font = "Tahoma"
'-
lvwServers.Initialize
lvwServers.BorderStyle = bsThin
lvwServers.ViewMode = vmDetails
lvwServers.ColumnAdd 0, "Server", 100, [caLeft]
lvwServers.ColumnAdd 1, "Port", 50, [caLeft]
lvwServers.Font.Name = "Tahoma"
lvwServers.FullRowSelect = True
lvwServers.HeaderFlat = False
lvwServers.HeaderHide = False
lvwServers.Font = "Tahoma"
'-
lvwUpdates.Initialize
lvwUpdates.BorderStyle = bsThin
lvwUpdates.ViewMode = vmDetails
lvwUpdates.ColumnAdd 0, "Server", 487, [caLeft]
lvwUpdates.Font.Name = "Tahoma"
lvwUpdates.FullRowSelect = True
lvwUpdates.HeaderFlat = False
lvwUpdates.HeaderHide = False
lvwUpdates.Font = "Tahoma"
End Sub

Private Sub Form_Unload(Cancel As Integer)
''If lSettings.sHandleErrors = True Then On Local Error Resume Next
mdiNexIRC.SetSplashVisible False
End Sub

Private Sub tabSplash_Click()
lStopUnload = True
End Sub
