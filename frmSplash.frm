VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{1ABC71B2-B0F7-4C1D-9870-3DED8934B20B}#1.0#0"; "prjXTab.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5070
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrActivate 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5160
      Top             =   4560
   End
   Begin prjXTab.XTab XTab1 
      Height          =   3615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6376
      TabCount        =   5
      TabCaption(0)   =   "Loading"
      TabContCtrlCnt(0)=   6
      Tab(0)ContCtrlCap(1)=   "prgLoading"
      Tab(0)ContCtrlCap(2)=   "txtLoading"
      Tab(0)ContCtrlCap(3)=   "lblVersion"
      Tab(0)ContCtrlCap(4)=   "Shape2"
      Tab(0)ContCtrlCap(5)=   "lblLoading"
      Tab(0)ContCtrlCap(6)=   "lblInformation"
      TabCaption(1)   =   "Setup"
      TabContCtrlCnt(1)=   18
      Tab(1)ContCtrlCap(1)=   "cmdDefaults"
      Tab(1)ContCtrlCap(2)=   "lstServers"
      Tab(1)ContCtrlCap(3)=   "cboNetwork"
      Tab(1)ContCtrlCap(4)=   "chkConnectOnStartup"
      Tab(1)ContCtrlCap(5)=   "txtRealname"
      Tab(1)ContCtrlCap(6)=   "txtEMail"
      Tab(1)ContCtrlCap(7)=   "txtNickname"
      Tab(1)ContCtrlCap(8)=   "chkShowOptionsOnStartup"
      Tab(1)ContCtrlCap(9)=   "chkShowSplashOnStartup"
      Tab(1)ContCtrlCap(10)=   "chkShowWizard"
      Tab(1)ContCtrlCap(11)=   "Shape1"
      Tab(1)ContCtrlCap(12)=   "Image1"
      Tab(1)ContCtrlCap(13)=   "lblServer"
      Tab(1)ContCtrlCap(14)=   "lblNetwork"
      Tab(1)ContCtrlCap(15)=   "lblRealName"
      Tab(1)ContCtrlCap(16)=   "lblEMail"
      Tab(1)ContCtrlCap(17)=   "Label1"
      Tab(1)ContCtrlCap(18)=   "lblOnStartup"
      TabCaption(2)   =   "Updates"
      TabContCtrlCnt(2)=   2
      Tab(2)ContCtrlCap(1)=   "txtUpdate"
      Tab(2)ContCtrlCap(2)=   "Shape3"
      TabCaption(3)   =   "About"
      TabContCtrlCnt(3)=   11
      Tab(3)ContCtrlCap(1)=   "Image4"
      Tab(3)ContCtrlCap(2)=   "lblLeonAiossa"
      Tab(3)ContCtrlCap(3)=   "Label21"
      Tab(3)ContCtrlCap(4)=   "Label20"
      Tab(3)ContCtrlCap(5)=   "Label19"
      Tab(3)ContCtrlCap(6)=   "Label18"
      Tab(3)ContCtrlCap(7)=   "Label17"
      Tab(3)ContCtrlCap(8)=   "Label16"
      Tab(3)ContCtrlCap(9)=   "Label15"
      Tab(3)ContCtrlCap(10)=   "Label14"
      Tab(3)ContCtrlCap(11)=   "Image3"
      TabCaption(4)   =   "Register"
      TabContCtrlCnt(4)=   13
      Tab(4)ContCtrlCap(1)=   "cmdBuy"
      Tab(4)ContCtrlCap(2)=   "cmdChange"
      Tab(4)ContCtrlCap(3)=   "txtName"
      Tab(4)ContCtrlCap(4)=   "txtPassword"
      Tab(4)ContCtrlCap(5)=   "lblClickStartMsg"
      Tab(4)ContCtrlCap(6)=   "lblRegisterNexIRC"
      Tab(4)ContCtrlCap(7)=   "lblName"
      Tab(4)ContCtrlCap(8)=   "lblPassword"
      Tab(4)ContCtrlCap(9)=   "lblStep1"
      Tab(4)ContCtrlCap(10)=   "lblStep1D"
      Tab(4)ContCtrlCap(11)=   "lblStep2"
      Tab(4)ContCtrlCap(12)=   "lblStep2D"
      Tab(4)ContCtrlCap(13)=   "Image2"
      TabTheme        =   2
      InActiveTabBackStartColor=   -2147483626
      InActiveTabBackEndColor=   -2147483626
      InActiveTabForeColor=   -2147483631
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   -2147483628
      TabStripBackColor=   -2147483626
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   -2147483627
      Begin NexIRC.isButton cmdBuy 
         Height          =   345
         Left            =   -73920
         TabIndex        =   50
         Top             =   960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   609
         Icon            =   "frmSplash.frx":000C
         Style           =   9
         Caption         =   "Buy"
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
      Begin VB.PictureBox cmdChange 
         Height          =   375
         Left            =   -68640
         ScaleHeight     =   315
         ScaleWidth      =   915
         TabIndex        =   41
         Top             =   2520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   230
         Left            =   -73920
         MousePointer    =   1  'Arrow
         TabIndex        =   40
         Top             =   1920
         Width           =   6255
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   230
         Left            =   -73920
         MousePointer    =   1  'Arrow
         TabIndex        =   39
         Top             =   2160
         Width           =   6255
      End
      Begin NexIRC.ctlTBox txtUpdate 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   29
         Top             =   480
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   5318
      End
      Begin NexIRC.isButton cmdDefaults 
         Height          =   330
         Left            =   -70920
         TabIndex        =   26
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         Icon            =   "frmSplash.frx":0028
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
      Begin VB.ListBox lstServers 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2565
         Left            =   -69720
         TabIndex        =   19
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox cboNetwork 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -69720
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox chkConnectOnStartup 
         Appearance      =   0  'Flat
         Caption         =   "Connect to &IRC"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   17
         Top             =   1560
         Width           =   4215
      End
      Begin VB.TextBox txtRealname 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   230
         Left            =   -73800
         MousePointer    =   1  'Arrow
         TabIndex        =   16
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txtEMail 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   230
         Left            =   -73800
         MousePointer    =   1  'Arrow
         TabIndex        =   15
         ToolTipText     =   "10"
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox txtNickname 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   230
         Left            =   -73800
         MaxLength       =   9
         MousePointer    =   1  'Arrow
         TabIndex        =   14
         ToolTipText     =   "8"
         Top             =   480
         Width           =   3975
      End
      Begin VB.CheckBox chkShowOptionsOnStartup 
         Appearance      =   0  'Flat
         Caption         =   "Show '&Customize'"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   13
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox chkShowSplashOnStartup 
         Appearance      =   0  'Flat
         Caption         =   "Show '&Splash'"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   12
         Top             =   2040
         Width           =   2655
      End
      Begin VB.CheckBox chkShowWizard 
         Appearance      =   0  'Flat
         Caption         =   "Show 'Setup &Wizard'"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   11
         Top             =   1320
         Width           =   2655
      End
      Begin MSComctlLib.ProgressBar prgLoading 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3240
         Width           =   7770
         _ExtentX        =   13705
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MousePointer    =   13
      End
      Begin VB.TextBox txtLoading 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   600
         Width           =   7755
      End
      Begin VB.Label lblClickStartMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Click 'Start' when complete"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -73920
         TabIndex        =   49
         Top             =   2640
         Width           =   3855
      End
      Begin VB.Label lblRegisterNexIRC 
         BackStyle       =   0  'Transparent
         Caption         =   "How to register NexIRC"
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
         Left            =   -74880
         TabIndex        =   48
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "&Name:"
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
         Left            =   -74880
         TabIndex        =   47
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
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
         Left            =   -74880
         TabIndex        =   46
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblStep1 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1:"
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
         Left            =   -74880
         TabIndex        =   45
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblStep1D 
         BackStyle       =   0  'Transparent
         Caption         =   "Registration costs $20 USD. Click the button below to launch paypal"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -73920
         TabIndex        =   44
         Top             =   720
         Width           =   6495
      End
      Begin VB.Label lblStep2 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2:"
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
         Left            =   -74880
         TabIndex        =   43
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblStep2D 
         BackStyle       =   0  'Transparent
         Caption         =   "Wait 1 day for code to be generated, you will recieve the code in e-mail, enter it below when it has been recieved"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   -73920
         TabIndex        =   42
         Top             =   1440
         Width           =   6375
      End
      Begin VB.Image Image2 
         Height          =   1170
         Left            =   -75000
         Top             =   2880
         Width           =   3555
      End
      Begin VB.Image Image4 
         Height          =   1170
         Left            =   -75000
         Top             =   2520
         Width           =   3555
      End
      Begin VB.Label lblLeonAiossa 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(Leon Aiossa)"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74640
         TabIndex        =   38
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Abby Ganyaw, Janis Perry, Eric Bishop, Clete Lindsay, 'Crossroad', John Scholten, Dasmius, SupraX, Magique"
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   -71160
         TabIndex        =   37
         Top             =   1080
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
         Left            =   -72960
         TabIndex        =   36
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Leon J Aiossa, Colin Foss"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -71160
         TabIndex        =   35
         Top             =   840
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
         Left            =   -72960
         TabIndex        =   34
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Tim Hueller, Jamie Cabral, Amy Fleischhacker"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -71160
         TabIndex        =   33
         Top             =   600
         Width           =   3735
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
         Left            =   -72960
         TabIndex        =   32
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Leon J Aiossa, Jamie Cabral"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -71160
         TabIndex        =   31
         Top             =   360
         Width           =   3615
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
         Left            =   -72960
         TabIndex        =   30
         Top             =   360
         Width           =   1335
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1560
         Left            =   -74880
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1875
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00808080&
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Height          =   3050
         Left            =   -74900
         Top             =   460
         Width           =   7730
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "Version 2.0"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   5295
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         Height          =   2565
         Left            =   105
         Top             =   585
         Width           =   7785
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         Height          =   750
         Left            =   -73810
         Top             =   470
         Width           =   4000
      End
      Begin VB.Image Image1 
         Height          =   1170
         Left            =   -75000
         Top             =   2520
         Width           =   3555
      End
      Begin VB.Label lblServer 
         BackStyle       =   0  'Transparent
         Caption         =   "&Server:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -71280
         TabIndex        =   25
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblNetwork 
         BackStyle       =   0  'Transparent
         Caption         =   "&Network:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -71280
         TabIndex        =   24
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblRealName 
         BackStyle       =   0  'Transparent
         Caption         =   "&Real name:"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   -74880
         TabIndex        =   23
         ToolTipText     =   "11"
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblEMail 
         BackStyle       =   0  'Transparent
         Caption         =   "&E-Mail:"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   -74880
         TabIndex        =   22
         ToolTipText     =   "9"
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nickname:"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   -74880
         TabIndex        =   21
         ToolTipText     =   "7"
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblOnStartup 
         BackStyle       =   0  'Transparent
         Caption         =   "On &Startup:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   20
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblLoading 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading ..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         MousePointer    =   13  'Arrow and Hourglass
         TabIndex        =   10
         Top             =   3360
         Width           =   7095
      End
      Begin VB.Label lblInformation 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   2295
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   7215
      End
   End
   Begin VB.Timer tmrEnableStartButton 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5640
      Top             =   4560
   End
   Begin VB.PictureBox cmdUpdate 
      Height          =   315
      Left            =   2040
      ScaleHeight     =   255
      ScaleWidth      =   795
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox cmdRegister 
      Height          =   315
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   795
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox cmdAdjustButtons 
      Height          =   315
      Left            =   4080
      ScaleHeight     =   255
      ScaleWidth      =   795
      TabIndex        =   2
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox cmdAbout 
      Height          =   315
      Left            =   1080
      ScaleHeight     =   255
      ScaleWidth      =   795
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox cmdSettings 
      Height          =   315
      Left            =   3000
      ScaleHeight     =   255
      ScaleWidth      =   795
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSWinsockLib.Winsock wskLatestVersion 
      Left            =   6120
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox cmdClose 
      Height          =   315
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   795
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox cmdStart1 
      Height          =   315
      Left            =   6000
      ScaleHeight     =   255
      ScaleWidth      =   795
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Visible         =   0   'False
      X1              =   3960
      X2              =   3960
      Y1              =   3960
      Y2              =   3720
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lSeconds As Integer
Private lTimeLeft As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub ActiveateAbout()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdAbout.Visible = False
cmdAdjustButtons.Visible = False
cmdClose.Visible = False
cmdRegister.Visible = False
cmdSettings.Visible = False
cmdUpdate.Visible = False
cmdStart.Caption = "OK"
fraAbout.Visible = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ActiveateAbout()"
End Sub

Public Sub ActiveateUpdateCheck()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sUpdateCheck = True Then
    Call DoColor(txtUpdate, "5• Checking for latest version")
    wskLatestVersion.Close
    wskLatestVersion.Connect "www.tnexgen.com", 80
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ActiveateUpdateCheck()"
End Sub

Private Function CheckVitalInfo() As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(lSettings.sServer) = 0 Then
    If Len(lstServers.Text) <> 0 Then
        lSettings.sServer = lstServers.Text
    Else
        lSettings.sSetupActivated = True
        fraRegister.Visible = False
        fraAbout.Visible = False
        fraUpdates.Visible = False
        fraSetup.Visible = True
        lstServers.SetFocus
        Beep
        Exit Function
    End If
End If
If Len(lSettings.sNetwork) = 0 Then
    If Len(cboNetwork.Text) <> 0 Then
        lSettings.sNetwork = cboNetwork.Text
    Else
        lSettings.sSetupActivated = True
        fraRegister.Visible = False
        fraAbout.Visible = False
        fraUpdates.Visible = False
        fraSetup.Visible = True
        cboNetwork.SetFocus
        Beep
        Exit Function
    End If
End If
If Len(lSettings.sNickname) = 0 Then
    If Len(txtNickname.Text) <> 0 Then
        lSettings.sNickname = txtNickname.Text
        WriteINI GetINIFile(iIRC), "Info", "NICKNAME", lSettings.sNickname
    Else
        lSettings.sSetupActivated = True
        fraRegister.Visible = False
        fraAbout.Visible = False
        fraUpdates.Visible = False
        fraSetup.Visible = True
        txtNickname.SetFocus
        Beep
        Exit Function
    End If
End If
If Len(lSettings.sEMail) = 0 Then
    If Len(txtEMail.Text) <> 0 Then
        lSettings.sEMail = txtEMail.Text
        WriteINI GetINIFile(iIRC), "Info", "USERNAME", lSettings.sEMail
    Else
        lSettings.sSetupActivated = True
        fraSetup.Visible = True
        txtEMail.SetFocus
        Beep
        Exit Function
    End If
End If
If Len(lSettings.sRealName) = 0 Then
    If Len(txtRealname.Text) <> 0 Then
        lSettings.sRealName = txtRealname.Text
        WriteINI GetINIFile(iIRC), "Info", "REALNAME", lSettings.sRealName
    Else
        lSettings.sSetupActivated = True
        fraSetup.Visible = True
        txtRealname.SetFocus
        Beep
        Exit Function
    End If
End If
CheckVitalInfo = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Function CheckVitalInfo() As Boolean"
End Function

Private Sub CheckShowNextTime()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If chkShowOptionsOnStartup.Value = 1 Then
    lSettings.sShowOptionsOnStartup = True
Else
    lSettings.sShowOptionsOnStartup = False
End If
WriteINI GetINIFile(iIRC), "IRC", "ShowSplashOnStartup", lSettings.sShowOptionsOnStartup
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub CheckShowNextTime()"
End Sub

Private Sub cboNetwork_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, j As Integer, mItem As ListItem, word() As String
lstServers.Clear
i = FindNetworkIndex(cboNetwork.Text)
If i <> 0 Then
    For j = 1 To lServers.sServerCount
        If lServers.sServer(j).sNetwork = i Then
            lstServers.AddItem lServers.sServer(j).sServer
        End If
    Next j
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cboNetwork_Click()"
End Sub

Private Sub cmdAbout_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lblVersion.Caption = "About NexIRC"
fraUpdate.Visible = False
fraSetup.Visible = False
fraRegister.Visible = False
fraUpdates.Visible = False
fraAbout.Visible = True
cmdRegister.Value = False
cmdAbout.Value = True
cmdUpdate.Value = False
cmdSettings.Value = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAbout_Click()"
End Sub

Private Sub cmdAdjustButtons_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
tmrActivate.Enabled = False
If cmdAdjustButtons.Caption = "&More" Then
    lblInformation.Caption = "To start NexIRC, click 'Start', If you wish to change your username, server, network and other settings, click 'Setup'. If you wish to check for updates, click 'Update'. To register NexIRC, click 'Register', to see the about screen, click 'About'"
    cmdAdjustButtons.Caption = "&Less"
    Line1.Visible = True
    Sleep 50: DoEvents
    cmdSettings.Visible = True
    Sleep 50: DoEvents
    cmdUpdate.Visible = True
    Sleep 50: DoEvents
    cmdAbout.Visible = True
    Sleep 50: DoEvents
    cmdRegister.Visible = True
ElseIf cmdAdjustButtons.Caption = "&Less" Then
    lblInformation.Caption = ""
    cmdAdjustButtons.Caption = "&More"
    cmdRegister.Visible = False
    Sleep 100: DoEvents
    cmdAbout.Visible = False
    Sleep 100: DoEvents
    cmdUpdate.Visible = False
    Sleep 100: DoEvents
    cmdSettings.Visible = False
    Sleep 100: DoEvents
    Line1.Visible = False
End If
fraUpdate.Visible = False
fraRegister.Visible = False
fraSetup.Visible = False
fraUpdates.Visible = False
cmdRegister.Value = False
cmdAbout.Value = False
cmdUpdate.Value = False
cmdSettings.Value = False
lblVersion.Caption = "Version: " & App.major & "." & App.minor
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAdjustButtons_Click()"
End Sub

Private Sub cmdBUY_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Surf "https://www.paypal.com/xclick/business=guidex%40tnexgen.com&item_name=Audiogen+Registration&amount=20.00&no_note=1&tax=0&currency_code=USD&lc=US", Me.hwnd
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdBUY_Click()"
End Sub

Private Sub cmdChange_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtName.Enabled = True
txtPassword.Enabled = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdChange_Click()"
End Sub

Private Sub cmdClose_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckShowNextTime: DoEvents
End
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdClose_Click()"
End Sub

Private Sub cmdDefaults_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtNickname.Text = "NexIRC"
txtEMail.Text = "nexirc@tnexgen.com"
txtRealname = "NexIRC User"
chkShowSplashOnStartup.Value = 1
chkShowOptionsOnStartup.Value = 0
chkConnectOnStartup.Value = 0
cboNetwork.Text = "undernet"
lstServers.ListIndex = 3
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdDefaults_Click()"
End Sub

Private Sub cmdRegister_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lRegInfo.rRegistered = True Then
    lblVersion.Caption = "Your registration Information"
Else
    lblVersion.Caption = "Register"
End If
fraUpdate.Visible = False
fraSetup.Visible = False
fraRegister.Visible = True
fraUpdates.Visible = False
fraAbout.Visible = False
cmdRegister.Value = True
cmdAbout.Value = False
cmdUpdate.Value = False
cmdSettings.Value = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdRegister_Click()"
End Sub

Private Sub cmdSettings_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sSetupActivated = True
fraUpdate.Visible = False
fraSetup.Visible = True
fraAbout.Visible = False
fraRegister.Visible = False
fraUpdates.Visible = False
cmdRegister.Value = False
cmdAbout.Value = False
cmdUpdate.Value = False
lblVersion.Caption = "Setup"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdSettings_Click()"
End Sub

Private Sub cmdStart_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim m As Boolean, i As String
CheckShowNextTime
lSettings.sShowSplashOnStartup = GetCheckboxValue(chkShowSplashOnStartup)
lSettings.sShowOptionsOnStartup = GetCheckboxValue(chkShowOptionsOnStartup)
lSettings.sConnectOnStartup = GetCheckboxValue(chkConnectOnStartup)
lSettings.sNetwork = cboNetwork.Text
If Len(lstServers.Text) <> 0 Then lSettings.sServer = lstServers.Text
lRegInfo.rName = txtName.Text
lRegInfo.rPassword = txtPassword.Text
If lRegInfo.rRegistered = False And Len(txtName.Text) <> 0 And Len(txtPassword.Text) <> 0 Then
    i = KeyGen(lRegInfo.rName, "pickles", 1)
    If i = lRegInfo.rPassword Then
        If lSettings.sGeneralPrompts = True Then
            MsgBox "Thank you very much for registering. All of the money made from NexIRC is spent on the development of NexIRC", vbInformation
        End If
        WriteINI GetINIFile(iIRC), "REGInfo", "NAME", lRegInfo.rName
        WriteINI GetINIFile(iIRC), "REGInfo", "PASSWORD", lRegInfo.rPassword
        lRegInfo.rRegistered = True
    Else
        Beep
        txtPassword.Text = ""
        txtName.Text = ""
        If lSettings.sGeneralPrompts = True Then
            MsgBox "The code you entered was not correct. The name did not match the password. Please try again", vbInformation
        End If
        fraAbout.Visible = False
        fraRegister.Visible = True
        fraSetup.Visible = False
        fraUpdates.Visible = False
        txtName.SetFocus
        Exit Sub
    End If
End If
If lSettings.sSetupActivated = True Then
    lSettings.sSetupActivated = False
    SaveSettings
End If
If CheckVitalInfo = True Then Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdStart_Click()"
End Sub

Private Sub cmdUpdate_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
fraUpdate.Visible = True
fraSetup.Visible = False
fraRegister.Visible = False
fraUpdates.Visible = False
fraAbout.Visible = False
cmdRegister.Value = False
cmdAbout.Value = False
cmdUpdate.Value = False
cmdSettings.Value = False
ActiveateUpdateCheck
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdUpdate_Click()"
End Sub

Private Sub cmdWhyRegister_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
MsgBox "To help ensure that NexIRC stays relatively free, registration unlocks no key features, however when you register, you are able to forego the two second delay caused by the splash screen. You are also able to change the 'Show Splash Screen on Startup' value to false. Registration also helps Team Nexgen to continue to make great software such as NexIRC. Users within the United States will recieve a CD by mail featuring NexIRC and other great Team Nexgen software, and music.", vbInformation, "Registration Benifits"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdWhyRegister_Click()"
End Sub

Private Sub Form_Load()
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
'Me.SetFocus
Exit Sub
Dim b As Boolean
txtUpdate.SetBackColor "0"
Me.Icon = mdiNexIRC.Icon
b = ReadINI(GetINIFile(iIRC), "Settings", "UsedSetupWizard", True)
If b = False Then chkShowWizard.Value = 1
lTimeLeft = 3
lblVersion.Caption = "Version: " & App.major & "." & App.minor
Me.Width = 7425
Me.Height = 4665
CutRegion cboNetwork.hwnd, cboNetwork, True
'picSplash.Picture = LoadPicture(lSpectrumThemes.sStartupGraphic)
Image1.Picture = LoadPicture(App.Path & "\data\images\halflogo.gif")
Image4.Picture = Image1.Picture
Image2.Picture = Image1.Picture
Image3.Picture = LoadPicture(App.Path & "\data\images\leon.gif")
lSettings.sSplashVisible = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As Boolean, mbox As VbMsgBoxResult
If chkShowWizard.Value = 1 Then
    Me.Visible = False
    frmSetupWizard.Show 1
    WriteINI GetINIFile(iIRC), "Settings", "UsedSetupWizard", "True"
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sSplashVisible = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub fraAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub fraAbout_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub fraSetup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub fraSetup_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub imgSplash_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub imgSplash_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub lblClickStartMsg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lblClickStartMsg_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub lblInformation_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    FormDrag Me
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lblLeonAiossa_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub lblLeonAiossa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lblLeonAiossa_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub lblRegisterNexIRC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lblRegisterNexIRC_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub lblSetupInformation_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lblSetupInformation_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub lblStep1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lblStep1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub lblStep1D_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lblStep1D_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub lblStep2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lblStep2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub lblStep2D_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lblStep2D_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub lblVersion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lblVersion_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub lstServers_Click()
chkConnectOnStartup.Caption = "Connect to " & lstServers.Text
End Sub

Private Sub picLoading_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picLoading_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picSplash_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picSplash_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub tmrActivate_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, F As Integer
tmrActivate.Enabled = False
If lSettings.sShowSplashOnStartup = True Then
    'cmdAbout.Value = True
    If lRegInfo.rRegistered = True Then
        lblStep1.Visible = False
        lblStep1D.Visible = False
        lblStep2.Visible = False
        lblStep2D.Visible = False
        lblClickStartMsg.Visible = False
        lblRegisterNexIRC.Visible = False
        cmdBuy.Visible = False
'        cmdWhyRegister.Visible = False
        'cmdRegister.Caption = "RegInfo"
        cmdChange.Visible = True
        chkShowSplashOnStartup.Enabled = True
'        cmdStart.Enabled = True
        txtName.Text = lRegInfo.rName
        txtPassword.Text = lRegInfo.rPassword
        txtName.Enabled = False
        txtPassword.Enabled = False
        LoadSettings
    Else
        chkShowSplashOnStartup.Enabled = False
        LoadSettings
    End If
    If Len(lSettings.sNetwork) = 0 Then lSettings.sNetwork = "undernet"
    If Len(lSettings.sServer) = 0 Then lSettings.sServer = "graz.at.eu.undernet.org"
    F = Int(FindComboBoxIndex(cboNetwork, lSettings.sNetwork))
    If cboNetwork.ListCount <> 0 Then cboNetwork.ListIndex = F
    lstServers.Text = lSettings.sServer
    'SetButtonType cmdWhyRegister
    'SetButtonType cmdAbout
    'SetButtonType cmdAdjustButtons
    'SetButtonType cmdChange
    'SetButtonType cmdClose
    'SetButtonType cmdDefaults
    'SetButtonType cmdRegister
    'SetButtonType cmdSettings
    'SetButtonType cmdStart
    'SetButtonType cmdUpdate
    txtNickname.Text = lSettings.sNickname
    txtEMail.Text = lSettings.sEMail
    txtRealname.Text = lSettings.sRealName
    SetCheckBoxValue chkShowSplashOnStartup, lSettings.sShowSplashOnStartup
    SetCheckBoxValue chkConnectOnStartup, lSettings.sConnectOnStartup
    SetCheckBoxValue chkShowOptionsOnStartup, lSettings.sShowOptionsOnStartup
    For i = 0 To lServers.sNetworkCount
        If Len(lServers.sNetwork(i).nDescription) <> 0 Then cboNetwork.AddItem lServers.sNetwork(i).nDescription
    Next i
    cboNetwork.ListIndex = FindComboBoxIndex(cboNetwork, lSettings.sNetwork)
    For i = 0 To lstServers.ListCount
        If LCase(lstServers.List(i)) = LCase(lSettings.sServer) Then
            lstServers.ListIndex = i
            Exit For
        End If
    Next i
Else
    Unload Me
End If
'picLoading.Visible = False
cmdAdjustButtons.Visible = True
cmdClose.Visible = True
'cmdStart.Visible = True
'If lSettings.sShowSplashOnStartup = True Then SetPictureColor frmSplash.picSplash, lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sRed, lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBlue, lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sGreen, False
If lSettings.sByPassStartupScreen = True Then
    cmdStart_Click
    tmrActivate.Enabled = False
    tmrEnableStartButton.Enabled = False
    Exit Sub
Else
    If lSettings.sUpdateCheck = True Then
        ActiveateUpdateCheck
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub txtEMail_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtEMail.SelStart = 0
txtEMail.SelLength = Len(txtEMail.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtEMail_GotFocus()"
End Sub

Private Sub txtLoading_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckTextBox txtLoading
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtLoading_Change()"
End Sub

Private Sub txtNickname_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtNickname.SelStart = 0
txtNickname.SelLength = Len(txtNickname.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtNickname_GotFocus()"
End Sub

Private Sub txtNickname_KeyDown(KeyCode As Integer, Shift As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Shift = 1 And KeyCode = 114 Then
    txtNickname.Text = "guest"
    txtEMail.Text = ""
    txtRealname.Text = ""
End If
If Shift = 1 And KeyCode = 119 Then
    txtNickname.Text = "Magique"
    txtEMail.Text = "magique@tnexgen.com"
    txtRealname.Text = "Mandi Mcdonald"
End If
If Shift = 1 And KeyCode = 120 Then
    txtNickname.Text = "KnightFal"
    txtEMail.Text = "knightfal@tnexgen.com"
    txtRealname.Text = "Colin Foss"
End If
If Shift = 1 And KeyCode = 122 Then
    cmdStart_Click
End If
If Shift = 1 And KeyCode = 123 Then
    txtNickname.Text = "|guideX|"
    txtEMail.Text = "guidex@tnexgen.com"
    txtRealname.Text = "Leon Aiossa"
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtNickname_KeyPress(KeyAscii As Integer)"
End Sub

Private Sub txtNickname_KeyPress(KeyAscii As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If KeyAscii = 13 Then
    cmdClose_Click
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtNickname_KeyPress(KeyAscii As Integer)"
End Sub

Private Sub txtRealname_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtRealname.SelStart = 0
txtRealname.SelLength = Len(txtRealname.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtRealname_GotFocus()"
End Sub

Private Sub wskLatestVersion_Close()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Call DoColor(txtUpdate, "4• Connection closed")
wskLatestVersion.Close
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub wskLatestVersion_Close()"
End Sub

Private Sub wskLatestVersion_Connect()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim GetString As String, ShortWebSite As String
wskLatestVersion.Tag = "OPEN"
Call DoColor(txtUpdate, "3• Connecting to NexIRC update site")
ShortWebSite = "http://www.tnexgen.com/nircupdate.ini"
GetString = "GET " + ShortWebSite + " HTTP/1.0" + vbCrLf
GetString = GetString + "Accept: */*" + vbCrLf
GetString = GetString + "Accept: text/html" + vbCrLf
GetString = GetString + vbCrLf
wskLatestVersion.SendData GetString
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub wskLatestVersion_Connect()"
End Sub

Private Sub wskLatestVersion_DataArrival(ByVal bytesTotal As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim buffer As String, msg As String, minor As String, major As String, mbox As VbMsgBoxResult
If wskLatestVersion.Tag = "OPEN" Then wskLatestVersion.GetData buffer
lSettings.sLatestVersionData = lSettings.sLatestVersionData & buffer
msg = Right(lSettings.sLatestVersionData, 4)
If Trim(msg) = App.major & "." & App.minor Then
    Call DoColor(txtUpdate, "3• This version of NexIRC is up to date")
Else
    major = Left(Trim(msg), 1)
    minor = Right(Trim(msg), 2)
    If App.major = Int(major) Then
        If App.minor < Int(minor) Then
            
            lblVersion.Caption = "Version: " & App.major & "." & App.minor & " (Out of date)"
            Call DoColor(txtUpdate, "4• This version of NexIRC is out of date")
            Pause 0.2
            Call DoColor(txtUpdate, "4• It is recommended you download the update")
            Pause 0.2
            Call DoColor(txtUpdate, "4• Your Version: " & App.major & "." & App.minor)
            Pause 0.2
            Call DoColor(txtUpdate, "4• New Version: " & major & "." & minor)
            Pause 0.2
            Call DoColor(txtUpdate, "4• Download location: http://www.tnexgen.com/downloads/nirc" & major & "." & minor & ".exe")
            Pause 0.2
            fraSetup.Visible = False
            fraUpdates.Visible = True
            cmdUpdate.Value = True
            WebBrowser1.Navigate "http://www.tnexgen.com/nexirc2dl.html"
            mbox = MsgBox("A newer version of NexIRC is now available." & vbCrLf & "Your Version: " & App.major & "." & App.minor & vbCrLf & "New Version: " & msg & vbCrLf & "Download it now?", vbYesNo + vbQuestion, "NexIRC - Update Monitor")
            If mbox = vbYes Then Surf "http://www.tnexgen.com/downloads/nirc" & major & minor & ".exe", mdiNexIRC.hwnd
        Else
            Call DoColor(txtUpdate, "5• This version of NexIRC is currently in development, or is a alpha/beta of NexIRC.")
            Call DoColor(txtUpdate, "5• E-Mail all bugs to guidex@tnexgen.com")
            lblVersion.Caption = "Version: " & App.major & "." & App.minor & " (In development)"
        End If
    ElseIf App.major < Int(major) Then
        lblVersion.Caption = "Version: " & App.major & "." & App.minor & " (Completely out of date)"
        fraSetup.Visible = False
        fraUpdates.Visible = True
        cmdUpdate.Value = True
        WebBrowser1.Navigate "http://www.tnexgen.com/nexirc2dl.html"
    ElseIf App.major > Int(major) Then
        lblVersion.Caption = "Version: " & App.major & "." & App.minor & " (Completely out of date)"
        Call DoColor(txtUpdate, "3• This version of NexIRC is completely out of date, it is strongly recommended that you update to the new version")
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub wskLatestVersion_DataArrival(ByVal bytesTotal As Long)"
End Sub

Private Sub wskLatestVersion_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Call DoColor(txtUpdate, "3• Unable to check for updates, check your internet connection (4" & Description & "3)")
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub wskLatestVersion_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)"
End Sub

