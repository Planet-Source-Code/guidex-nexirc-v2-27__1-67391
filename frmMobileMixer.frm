VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMobileMixer 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   885
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   517
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   59
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin nexIRC.ctlXPButton cmdPlaylist 
      Height          =   375
      Left            =   30
      TabIndex        =   21
      Top             =   6600
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Playlist"
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
      MICON           =   "frmMobileMixer.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdRandom 
      Height          =   375
      Left            =   30
      TabIndex        =   20
      Top             =   6240
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Random"
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
      MICON           =   "frmMobileMixer.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chkContinuous 
      Appearance      =   0  'Flat
      Caption         =   "Cont."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   15
      ToolTipText     =   "Continuous Playback (Starts one song after another)"
      Top             =   5400
      Width           =   975
   End
   Begin VB.CheckBox chkMute 
      Appearance      =   0  'Flat
      Caption         =   "Mute"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   17
      ToolTipText     =   "Mute volume"
      Top             =   5880
      Width           =   975
   End
   Begin VB.CheckBox chkShuffle 
      Appearance      =   0  'Flat
      Caption         =   "Shuffle"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   16
      ToolTipText     =   "Shuffle Mode"
      Top             =   5640
      Width           =   975
   End
   Begin nexIRC.ctlFormDragger FormDragger1 
      Align           =   1  'Align Top
      Height          =   285
      Left            =   0
      Top             =   0
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   503
   End
   Begin MSComctlLib.Slider sldMixer 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Bass Balance"
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      Max             =   65535
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider sldMixer 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "Treble Balance"
      Top             =   2520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      Max             =   65535
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider sldMixer 
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   10
      ToolTipText     =   "Line In Volume"
      Top             =   3720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      Max             =   65535
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider sldMixer 
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   8
      ToolTipText     =   "CDAudio Volume"
      Top             =   3120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      Max             =   65535
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider sldMixer 
      Height          =   375
      Index           =   5
      Left            =   0
      TabIndex        =   12
      ToolTipText     =   "Mic Volume"
      Top             =   4320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      Max             =   65535
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider sldMixer 
      Height          =   375
      Index           =   6
      Left            =   0
      TabIndex        =   14
      ToolTipText     =   "Aux Volume"
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      Max             =   65535
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider sldMixer 
      Height          =   375
      Index           =   7
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Wave Volume"
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      Max             =   65535
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider sldMixer 
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Progress of File"
      Top             =   720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      Max             =   65535
      TickStyle       =   3
   End
   Begin VB.TextBox txtMixerText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   6600
      Visible         =   0   'False
      Width           =   615
   End
   Begin nexIRC.ctlXPButton cmdAdd 
      Height          =   375
      Left            =   30
      TabIndex        =   22
      Top             =   6960
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Add"
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
      MICON           =   "frmMobileMixer.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdAudioSettings 
      Height          =   375
      Left            =   30
      TabIndex        =   23
      Top             =   7320
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Settings"
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
      MICON           =   "frmMobileMixer.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblMixer 
      BackStyle       =   0  'Transparent
      Caption         =   "Bass:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblMixer 
      BackStyle       =   0  'Transparent
      Caption         =   "Treble:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblMixer 
      BackStyle       =   0  'Transparent
      Caption         =   "CD Audio:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblMixer 
      BackStyle       =   0  'Transparent
      Caption         =   "Line In:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lblMixer 
      BackStyle       =   0  'Transparent
      Caption         =   "Mic:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblMixer 
      BackStyle       =   0  'Transparent
      Caption         =   "Aux:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   13
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label lblMixer 
      BackStyle       =   0  'Transparent
      Caption         =   "Wave:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblMixer 
      BackStyle       =   0  'Transparent
      Caption         =   "&Progress:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmMobileMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Private Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Private Declare Function mixerGetDevCaps Lib "winmm.dll" Alias "mixerGetDevCapsA" (ByVal uMxId As Long, pmxcaps As MIXERCAPS, ByVal cbmxcaps As Long) As Long
Private Declare Function mixerGetID Lib "winmm.dll" (ByVal hmxobj As Long, pumxID As Long, ByVal fdwId As Long) As Long
Private Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Private Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Private Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function mixerMessage Lib "winmm.dll" (ByVal hmx As Long, ByVal uMsg As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long) As Long
Private Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Private Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Private Declare Function RegisterDLL Lib "Regist10.dll" Alias "REGISTERDLL" (ByVal DllPath As String, bRegister As Boolean) As Boolean
Private Declare Function mmioClose Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Private Declare Function mmioDescend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKInfo, lpckParent As MMCKInfo, ByVal uFlags As Long) As Long
Private Declare Function mmioDescendParent Lib "winmm.dll" Alias "mmioDescend" (ByVal hmmio As Long, lpck As MMCKInfo, ByVal X As Long, ByVal uFlags As Long) As Long
Private Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioInfo As mmioInfo, ByVal dwOpenFlags As Long) As Long
Private Declare Function mmioRead Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As Long, ByVal cch As Long) As Long
Private Declare Function mmioReadFormat Lib "winmm.dll" Alias "mmioRead" (ByVal hmmio As Long, ByRef pch As waveFormat, ByVal cch As Long) As Long
Private Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Private Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKInfo, ByVal uFlags As Long) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Private Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Const MMSYSERR_NOERROR = 0
Private Const MAXPNAMELEN = 32
Private Const MIXER_LONG_NAME_CHARS = 64
Private Const MIXER_SHORT_NAME_CHARS = 16
Private Const MIXER_GETLINEInfoF_COMPONENTTYPE = &H3&
Private Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Private Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
Private Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Private Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Private Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Private Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Private Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Private Const MIXERCONTROL_CT_CLASS_CUSTOM = &H0&
Private Const MIXERCONTROL_CT_CLASS_LIST = &H70000000
Private Const MIXERCONTROL_CT_CLASS_MASK = &HF0000000
Private Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Private Const MIXERCONTROL_CT_CLASS_NUMBER = &H30000000
Private Const MIXERCONTROL_CT_CLASS_SLIDER = &H40000000
Private Const MIXERCONTROL_CT_CLASS_TIME = &H60000000
Private Const MIXERCONTROL_CT_SC_LIST_MULTIPLE = &H1000000
Private Const MIXERCONTROL_CT_SC_LIST_SINGLE = &H0&
Private Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
Private Const MIXERCONTROL_CT_SC_SWITCH_BUTTON = &H1000000
Private Const MIXERCONTROL_CT_SC_TIME_MICROSECS = &H0&
Private Const MIXERCONTROL_CT_SC_TIME_MILLISECS = &H1000000
Private Const MIXERCONTROL_CT_SUBCLASS_MASK = &HF000000
Private Const MIXERCONTROL_CT_UNITS_CUSTOM = &H0&
Private Const MIXERCONTROL_CT_UNITS_DECIBELS = &H40000
Private Const MIXERCONTROL_CT_UNITS_MASK = &HFF0000
Private Const MIXERCONTROL_CT_UNITS_PERCENT = &H50000
Private Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Private Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Private Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Private Const MIXERLINE_COMPONENTTYPE_SRC_ANALOG = &H1000& + 10
Private Const MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY = &H1000& + 9
Private Const MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC = &H1000& + 5
Private Const MIXERLINE_COMPONENTTYPE_SRC_DIGITAL = &H1000& + 1
Private Const MIXERLINE_COMPONENTTYPE_SRC_LAST = &H1000& + 10
Private Const MIXERLINE_COMPONENTTYPE_SRC_LINE = &H1000& + 2
Private Const MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER = &H1000& + 4
Private Const MIXERLINE_COMPONENTTYPE_SRC_UNDEFINED = &H1000& + 0
Private Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = &H1000& + 8
Private Const MIXERLINE_COMPONENTTYPE_SRC_I25InVol = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 1)
Private Const MIXERLINE_COMPONENTTYPE_SRC_LINEVol = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
Private Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Private Const MIXERLINE_COMPONENTTYPE_SRC_MIDIVol = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 4)
Private Const MIXERLINE_COMPONENTTYPE_SRC_CDVol = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)
Private Const MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 6)
Private Const MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 7)
Private Const MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Private Const MIXERLINE_COMPONENTTYPE_SRC_AUXVol = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 0)
Private Const MIXERLINE_COMPONENTTYPE_DST_UNDEFINED = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 0)
Private Const MIXERLINE_COMPONENTTYPE_DST_DIGITAL = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 1)
Private Const MIXERLINE_COMPONENTTYPE_DST_LINE = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 2)
Private Const MIXERLINE_COMPONENTTYPE_DST_MONITOR = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 3)
Private Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Private Const MIXERLINE_COMPONENTTYPE_DST_HEADPHONES = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 5)
Private Const MIXERLINE_COMPONENTTYPE_DST_TELEPHONE = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 6)
Private Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Private Const MIXERLINE_COMPONENTTYPE_DST_LAST = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)
Private Const MIXERLINE_COMPONENTTYPE_DST_VOICEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)
Private Const MMIO_READ = &H0
Private Const MMIO_FINDCHUNK = &H10
Private Const MMIO_FINDRIFF = &H20
Private Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_BASS = (MIXERCONTROL_CONTROLTYPE_FADER + 2)
Private Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_TREBLE = (MIXERCONTROL_CONTROLTYPE_FADER + 3)
Private Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Private Const MIXERCONTROL_CONTROLTYPE_MUTE = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
Private Const MIXERCONTROL_CONTROLTYPE_EQUALIZER = (MIXERCONTROL_CONTROLTYPE_FADER + 4)
Private Const MIXERCONTROL_CONTROLTYPE_LOUDNESS = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 4)
Private Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)
Private Const MIXERCONTROL_CONTROLTYPE_BOOLEANMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_BUTTON = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BUTTON Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_CUSTOM = (MIXERCONTROL_CT_CLASS_CUSTOM Or MIXERCONTROL_CT_UNITS_CUSTOM)
Private Const MIXERCONTROL_CONTROLTYPE_DECIBELS = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_DECIBELS)
Private Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_MICROTIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MICROSECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_MILLITIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MILLISECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_MIXER = (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
Private Const MIXERCONTROL_CONTROLTYPE_MONO = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 3)
Private Const MIXERCONTROL_CONTROLTYPE_SLIDER = (MIXERCONTROL_CT_CLASS_SLIDER Or MIXERCONTROL_CT_UNITS_SIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_STEREOENH = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 5)
Private Const MIXERCONTROL_CONTROLTYPE_UNSIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_UNSIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_SINGLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_SINGLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_MUX = (MIXERCONTROL_CONTROLTYPE_SINGLESELECT + 1)
Private Const MIXERCONTROL_CONTROLTYPE_PAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 1)
Private Const MIXERCONTROL_CONTROLTYPE_PERCENT = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_PERCENT)
Private Const MIXERCONTROL_CONTROLTYPE_QSOUNDPAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 2)
Private Const MIXERCONTROL_CONTROLTYPE_SIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_SIGNED)
Private Const MIXERLINE_LINEF_ACTIVE = &H1&
Private Const MIXERLINE_LINEF_DISCONNECTED = &H8000&
Private Const MIXERLINE_LINEF_SOURCE = &H80000000
Private Const MIXERLINE_TARGETTYPE_AUX = 5
Private Const MIXERLINE_TARGETTYPE_MIDIIN = 4
Private Const MIXERLINE_TARGETTYPE_MIDIOUT = 3
Private Const MIXERLINE_TARGETTYPE_UNDEFINED = 0
Private Const MIXERLINE_TARGETTYPE_WAVEIN = 2
Private Const MIXERLINE_TARGETTYPE_WAVEOUT = 1
Private Const MIXERR_INVALLINE = 1024 + 0
Private Const MIXERR_BASE = 1024
Private Const MIXERR_INVALCONTROL = 1024 + 1
Private Const MIXERR_INVALVALUE = 1024 + 2
Private Const MIXERR_LASTERROR = 1024 + 2
Private Type MIXERCAPS
     wMid As Integer
     wPid As Integer
     vDriverVersion As Long
     szPname As String * MAXPNAMELEN
     fdwSupport As Long
     cDestinations As Long
End Type
Private Type Target
     dwType As Long
     dwDeviceID As Long
     wMid As Integer
     wPid As Integer
     vDriverVersion As Long
     szPname As String * MAXPNAMELEN
End Type
Private Type MIXERLINE
     cbStruct As Long
     dwDestination As Long
     dwSource As Long
     dwLineID As Long
     fdwLine As Long
     dwUser As Long
     dwComponentType As Long
     cChannels As Long
     cConnections As Long
     cControls As Long
     szShortName As String * MIXER_SHORT_NAME_CHARS
     szName As String * MIXER_LONG_NAME_CHARS
     lpTarget As Target
End Type
Private Type MIXERLINECONTROLS
     cbStruct As Long
     dwLineID As Long
     dwControl As Long
     cControls As Long
     cbmxctrl As Long
     pamxctrl As Long
End Type
Private Type MIXERCONTROL
     cbStruct As Long
     dwControlID As Long
     dwControlType As Long
     fdwControl As Long
     cMultipleItems As Long
     szShortName(1 To MIXER_SHORT_NAME_CHARS) As Byte
     szName(1 To MIXER_LONG_NAME_CHARS) As Byte
     Bounds(1 To 6) As Long
     Metrics(1 To 6) As Long
End Type
Private Type MIXERCONTROLDETAILS
     cbStruct As Long
     dwControlID As Long
     cChannels As Long
     Item As Long
     cbDetails As Long
     paDetails As Long
End Type
Private Type MIXERCONTROLDETAILS_BOOLEAN
     fValue As Long
End Type
Private Type MIXERCONTROLDETAILS_LISTTEXT
     dwParam1 As Long
     dwParam2 As Long
     szName As String * MIXER_LONG_NAME_CHARS
End Type
Private Type MIXERCONTROLDETAILS_SIGNED
     lValue As Long
End Type
Private Type MIXERCONTROLDETAILS_UNSIGNED
     dwValue As Long
End Type
Private Type waveFormat
   wFormatTag As Integer
   nChannels As Integer
   nSamplesPerSec As Long
   nAvgBytesPerSec As Long
   nBlockAlign As Integer
   wBitsPerSample As Integer
   cbSize As Integer
End Type
Private Type mmioInfo
   dwFlags As Long
   fccIOProc As Long
   pIOProc As Long
   wErrorRet As Long
   htask As Long
   cchBuffer As Long
   pchBuffer As String
   pchNext As String
   pchEndRead As String
   pchEndWrite As String
   lBufOffset As Long
   lDiskOffset As Long
   adwInfo(4) As Long
   dwReserved1 As Long
   dwReserved2 As Long
   hmmio As Long
End Type
Private Type MMCKInfo
    ckid As Long
    ckSize As Long
    fccType As Long
    dwDataOffset As Long
    dwFlags As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'Private Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Any) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal lLeft As Long, ByVal lTop As Long) As Long
Private Declare Function SetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Const SWW_HPARENT = -8
Private Const HTRIGHT = 11
'Private Const WM_NCLBUTTONDOWN = &HA1
Dim bMoving As Boolean
Dim lFloatingLeft As Long
Dim lFloatingTop As Long
Dim lFloatingWidth As Long
Dim lFloatingHeight As Long
Dim volR As Long
Dim volL As Long
Dim volume As Long
Dim Mute As MIXERCONTROL
Dim unmute As MIXERCONTROL
Dim ONOFF As MIXERCONTROL
Dim hmixer As Long
Dim VolCtrl As MIXERCONTROL
Dim WavCtrl As MIXERCONTROL
Dim CDVol As MIXERCONTROL
Dim LineVol As MIXERCONTROL
Dim MICROPHONE As MIXERCONTROL
Dim PCSPEAKER As MIXERCONTROL
Dim AUXVol As MIXERCONTROL
Dim TELEPHONE As MIXERCONTROL
Dim MIDIVol As MIXERCONTROL
Dim rc As Long
Dim I25InVol As MIXERCONTROL
Dim Treble As MIXERCONTROL
Dim Bass As MIXERCONTROL
Dim ok As Boolean

Private Function GetMixerControl(ByVal hmixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, ByRef mxc As MIXERCONTROL) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mxlc As MIXERLINECONTROLS, mxl As MIXERLINE, hMem As Long, rc As Long
mxl.cbStruct = Len(mxl)
mxl.dwComponentType = componentType
rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEInfoF_COMPONENTTYPE)
If (MMSYSERR_NOERROR = rc) Then
    mxlc.cbStruct = Len(mxlc)
    mxlc.dwLineID = mxl.dwLineID
    mxlc.dwControl = ctrlType
    mxlc.cControls = 1
    mxlc.cbmxctrl = Len(mxc)
    hMem = GlobalAlloc(&H40, Len(mxc))
    mxlc.pamxctrl = GlobalLock(hMem)
    mxc.cbStruct = Len(mxc)
    rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
    If (MMSYSERR_NOERROR = rc) Then
        GetMixerControl = True
        CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
    Else
        GetMixerControl = False
    End If
    GlobalFree (hMem)
    Exit Function
End If
GetMixerControl = False
End Function

Private Function SetVolumeControl(ByVal hmixer As Long, mxc As MIXERCONTROL, ByVal volume As Long) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mxcd As MIXERCONTROLDETAILS, vol As MIXERCONTROLDETAILS_UNSIGNED, rc As Long, hMem As Long
mxcd.cbStruct = Len(mxcd)
mxcd.dwControlID = mxc.dwControlID
mxcd.cChannels = 1
mxcd.Item = 0
mxcd.cbDetails = Len(vol)
hMem = GlobalAlloc(&H40, Len(vol))
mxcd.paDetails = GlobalLock(hMem)
vol.dwValue = volume
CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
GlobalFree (hMem)
If (MMSYSERR_NOERROR = rc) Then
    SetVolumeControl = True
Else
    SetVolumeControl = False
End If
End Function

Private Function SetPANControl(ByVal hmixer As Long, mxc As MIXERCONTROL, ByVal volL As Long, ByVal volR As Long) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mxcd As MIXERCONTROLDETAILS, vol(1) As MIXERCONTROLDETAILS_UNSIGNED, rc As Long, hMem As Long
mxcd.Item = mxc.cMultipleItems
mxcd.dwControlID = mxc.dwControlID
mxcd.cbStruct = Len(mxcd)
mxcd.cbDetails = Len(vol(1))
mxcd.cChannels = 2
hMem = GlobalAlloc(&H40, Len(vol(1)))
mxcd.paDetails = GlobalLock(hMem)
vol(1).dwValue = volR
vol(0).dwValue = volL
CopyPtrFromStruct mxcd.paDetails, vol(1).dwValue, Len(vol(0)) * mxcd.cChannels
CopyPtrFromStruct mxcd.paDetails, vol(0).dwValue, Len(vol(1)) * mxcd.cChannels
rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
GlobalFree (hMem)
If (MMSYSERR_NOERROR = rc) Then
    SetPANControl = True
Else
    SetPANControl = False
End If
End Function

Private Function unSetMuteControl(ByVal hmixer As Long, mxc As MIXERCONTROL, ByVal unmute As Long) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mxcd As MIXERCONTROLDETAILS, vol As MIXERCONTROLDETAILS_UNSIGNED, rc As Long, hMem As Long
mxcd.cbStruct = Len(mxcd)
mxcd.dwControlID = mxc.dwControlID
mxcd.cChannels = 1
mxcd.Item = 0
mxcd.cbDetails = Len(vol)
hMem = GlobalAlloc(&H40, Len(vol))
mxcd.paDetails = GlobalLock(hMem)
vol.dwValue = unmute
CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
GlobalFree (hMem)
If (MMSYSERR_NOERROR = rc) Then
    unSetMuteControl = True
Else
    unSetMuteControl = False
End If
End Function

Private Function SetMuteControl(ByVal hmixer As Long, mxc As MIXERCONTROL, Mute As Boolean) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mxcd As MIXERCONTROLDETAILS, vol As MIXERCONTROLDETAILS_UNSIGNED, rc As Long, hMem As Long
mxcd.cbStruct = Len(mxcd)
mxcd.dwControlID = mxc.dwControlID
mxcd.cChannels = 1
mxcd.Item = 0
mxcd.cbDetails = Len(vol)
hMem = GlobalAlloc(&H40, Len(vol))
mxcd.paDetails = GlobalLock(hMem)
CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
GlobalFree (hMem)
If (MMSYSERR_NOERROR = rc) Then
    SetMuteControl = True
Else
    SetMuteControl = False
End If
End Function

Private Function GetVolumeControlValue(ByVal hmixer As Long, mxc As MIXERCONTROL) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mxcd As MIXERCONTROLDETAILS, vol As MIXERCONTROLDETAILS_UNSIGNED, rc As Long, hMem As Long
mxcd.cbStruct = Len(mxcd)
mxcd.dwControlID = mxc.dwControlID
mxcd.cChannels = 1
mxcd.Item = 0
mxcd.cbDetails = Len(vol)
mxcd.paDetails = 0
hMem = GlobalAlloc(&H40, Len(vol))
mxcd.paDetails = GlobalLock(hMem)
rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
CopyStructFromPtr vol, mxcd.paDetails, Len(vol)
GlobalFree (hMem)
If (MMSYSERR_NOERROR = rc) Then
    GetVolumeControlValue = vol.dwValue
Else
    GetVolumeControlValue = -1
End If
End Function

Public Sub InitQuickMixer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
rc = mixerOpen(hmixer, 0, 0, 0, 0)
If ((MMSYSERR_NOERROR <> rc)) Then
    ProcessReplaceString sProgrammingError, lSettings.sActiveServerForm.txtIncoming, "Mixer could not initialize", Str(MMSYSERR_NOERROR)
    Exit Sub
End If
For i = 0 To 11
Select Case i
Case 1
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol, MIXERCONTROL_CONTROLTYPE_VOLUME, WavCtrl)
    If (ok = True) Then
        txtMixerText(7) = GetVolumeControlValue(hmixer, WavCtrl)
        frmMobileMixer.sldMixer(7).Value = txtMixerText(7).Text
    Else
        frmMobileMixer.sldMixer(7).Enabled = False
    End If
Case 2
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, MIXERCONTROL_CONTROLTYPE_VOLUME, MICROPHONE)
    If (ok = True) Then
        txtMixerText(5) = GetVolumeControlValue(hmixer, MICROPHONE)
        frmMobileMixer.sldMixer(5).Value = txtMixerText(5).Text
    Else
        frmMobileMixer.sldMixer(5).Enabled = False
    End If
Case 3
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_CDVol, MIXERCONTROL_CONTROLTYPE_VOLUME, CDVol)
    If (ok = True) Then
        txtMixerText(4) = GetVolumeControlValue(hmixer, CDVol)
        frmMobileMixer.sldMixer(4).Value = txtMixerText(4).Text
    Else
        frmMobileMixer.sldMixer(4).Enabled = False
    End If
Case 4
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_AUXVol, MIXERCONTROL_CONTROLTYPE_VOLUME, AUXVol)
    If (ok = True) Then
        txtMixerText(6) = GetVolumeControlValue(hmixer, AUXVol)
        frmMobileMixer.sldMixer(6).Value = txtMixerText(6).Text
    Else
        frmMobileMixer.sldMixer(6).Enabled = False
    End If
Case 9
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_LINEVol, MIXERCONTROL_CONTROLTYPE_VOLUME, LineVol)
    If (ok = True) Then
        txtMixerText(3) = GetVolumeControlValue(hmixer, LineVol)
        frmMobileMixer.sldMixer(3).Value = txtMixerText(3).Text
    Else
        frmMobileMixer.sldMixer(3).Enabled = False
    End If
Case 10
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_BASS, Bass)
    If (ok = True) Then
        txtMixerText(0) = GetVolumeControlValue(hmixer, Bass)
        frmMobileMixer.sldMixer(0).Value = txtMixerText(0).Text
    Else
        frmMobileMixer.sldMixer(0).Enabled = False
    End If
Case 11
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_TREBLE, Treble)
    If (ok = True) Then
        txtMixerText(1) = GetVolumeControlValue(hmixer, Treble)
        frmMobileMixer.sldMixer(1).Value = txtMixerText(1).Text
    Else
        frmMobileMixer.sldMixer(1).Enabled = False
    End If
End Select
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub InitQuickMixer()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sMobileMixerVisible = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub sldMixer_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Index = 2 Then
    Select Case lPlayback.pCurrentEngine
    Case pMp3
        'mdiNexIRC.ctlMP3OCX.Seek sldMixer(2).Value
        'mdiNexIRC.ctlMP3OCX.Tag = "CHANGE"
    End Select
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub sldMixer_Click(Index As Integer)"
End Sub

Private Sub sldMixer_Scroll(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, v As String
Select Case Index
Case 0
    SetVolumeControl hmixer, Bass, volume
    volume = CLng(sldMixer(Index).Value)
    txtMixerText(Index).Text = volume
Case 1
    SetVolumeControl hmixer, Treble, volume
    volume = CLng(sldMixer(Index).Value)
    txtMixerText(Index).Text = volume
Case 2
    If lPlayback.pCurrentEngine = pMp3 Then
        'If mdiNexIRC.ctlMP3OCX.Tag = "CHANGE" Then
        '    mdiNexIRC.ctlMP3OCX.Tag = ""
        '    ProgressScrolling = False
        'Else
        '    ProgressScrolling = True
        'End If
    End If
Case 3
    volume = CLng(sldMixer(Index).Value)
    txtMixerText(Index).Text = volume
    SetVolumeControl hmixer, LineVol, volume
Case 4
    volume = CLng(sldMixer(Index).Value)
    txtMixerText(Index).Text = volume
    SetVolumeControl hmixer, CDVol, volume
Case 5
    volume = CLng(sldMixer(Index).Value)
    txtMixerText(Index).Text = volume
    SetVolumeControl hmixer, MICROPHONE, volume
Case 6
    volume = CLng(sldMixer(Index).Value)
    txtMixerText(Index).Text = volume
    SetVolumeControl hmixer, AUXVol, volume
Case 7
    volume = CLng(sldMixer(Index).Value)
    txtMixerText(Index).Text = volume
    SetVolumeControl hmixer, WavCtrl, volume
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub sldMixer_Scroll(Index As Integer)"
End Sub

Private Sub chkContinuous_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If chkContinuous.Value = 1 Then
    mdiNexIRC.tmrContinuousPlay.Enabled = True
    lSettings.sContinuousPlay = True
    WriteINI GetINIFile(iIRC), "Settings", "ContinuousPlay", "True"
ElseIf chkContinuous.Value = 0 Then
    mdiNexIRC.tmrContinuousPlay.Enabled = False
    lSettings.sContinuousPlay = False
    WriteINI GetINIFile(iIRC), "Settings", "ContinuousPlay", "False"
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkContinuous_Click()"
End Sub

Private Sub chkMute_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If chkMute.Value = 0 Then
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol, MIXERCONTROL_CONTROLTYPE_MUTE, Mute)
    SetMuteControl hmixer, Mute, 1
ElseIf chkMute.Value = 1 Then
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
    unSetMuteControl hmixer, unmute, 1
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkMute_Click()"
End Sub

Private Sub chkShuffle_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If chkShuffle.Value = 1 Then
    lSettings.sShuffle = True
Else
    lSettings.sShuffle = False
End If
WriteINI GetINIFile(iIRC), "Settings", "Shuffle", lSettings.sShuffle
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkShuffle_Click()"
End Sub

Private Sub cmdAdd_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'frmChannels.lvwChannels.Sort 0, soAscending, stString
frmAddMedia.Show
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAdd_Click()"
End Sub

Private Sub cmdAudioSettings_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Button = 1 Then
    Select Case lPlayback.pCurrentEngine
    Case pMp3
        If mdiNexIRC.picMP3OCX.Visible = False Then
            mdiNexIRC.picMP3OCX.Visible = True
        Else
            mdiNexIRC.picMP3OCX.Visible = False
        End If
    End Select
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAudioSettings_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub cmdPlaylist_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sPlaylistVisible = True Then
    frmPlaylist.Hide
    lSettings.sPlaylistVisible = False
Else
    frmPlaylist.Show
    frmPlaylist.Visible = True
    If frmPlaylist.WindowState = vbMinimized Then frmPlaylist.WindowState = vbNormal
    lSettings.sPlaylistVisible = True
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdPlaylist_Click()"
End Sub

Private Sub cmdRandom_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
GetR:
i = GetRnd(lFiles.fCount)
If DoesFileExist(lFiles.fFile(i).fFilename) = True Then
    MenuStop
    PlayFile lFiles.fFile(i).fFilename
Else
    GoTo GetR
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdRandom_Click()"
End Sub

Private Sub StoreFormDimensions()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Not bMoving Then
    If bDocked Then
        lDockedWidth = Me.Width
        lDockedHeight = Me.Height
    Else
        lFloatingLeft = Me.Left
        lFloatingTop = Me.Top
        lFloatingWidth = Me.Width
        lFloatingHeight = Me.Height
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub StoreFormDimensions()"
End Sub

Private Sub Form_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
lDockedWidth = mdiNexIRC.picMobileMixer.ScaleWidth + (8 * Screen.TwipsPerPixelX)
lDockedHeight = mdiNexIRC.picMobileMixer.ScaleHeight + (8 * Screen.TwipsPerPixelY)
lFloatingLeft = Me.Left
lFloatingTop = Me.Top
lFloatingWidth = Me.Width
lFloatingHeight = Me.Height
bDocked = True
SetParent Me.hWnd, mdiNexIRC!picMobileMixer.hWnd
lSettings.sMobileMixerVisible = True
For i = 1 To 11
    Load txtMixerText(i)
Next i
InitQuickMixer
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub txtMixerText_Change(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Index <> 2 And Index <> 8 And Index <> 9 And Index <> 10 And Index <> 11 Then frmMobileMixer.sldMixer(Index).Value = Int(txtMixerText(Index).Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtMixerText_Change(Index As Integer)"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Call SetWindowWord(Me.hWnd, SWW_HPARENT, 0&)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub Form_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Me.WindowState <> vbMinimized Then
    StoreFormDimensions
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Resize()"
End Sub

Private Sub FormDragger1_DblClick()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
bMoving = True
If bDocked = True Then
    Me.Visible = False
    bDocked = False
    SetParent Me.hWnd, 0
    Me.Move lFloatingLeft, lFloatingTop, lFloatingWidth, lFloatingHeight
    mdiNexIRC!picMobileMixer.Visible = False
    Me.Visible = True
    Call SetWindowWord(Me.hWnd, SWW_HPARENT, mdiNexIRC.hWnd)
Else
    bDocked = True
    SetParent Me.hWnd, mdiNexIRC!picMobileMixer.hWnd
    Me.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
    mdiNexIRC!picMobileMixer.Visible = True
End If
bMoving = False
mdiNexIRC.ActivateResize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub FormDragger1_DblClick()"
End Sub

Private Sub FormDragger1_FormDropped(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim rct As RECT
GetWindowRect mdiNexIRC!picMobileMixer.hWnd, rct
With rct
    .Left = .Left - 4
    .Top = .Top - 4
    .Right = .Right + 4
    .Bottom = .Bottom + 4
End With
If PtInRect(rct, FormLeft, FormTop) Then
    bDocked = True
    SetParent Me.hWnd, mdiNexIRC!picMobileMixer.hWnd
    Me.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
    mdiNexIRC!picMobileMixer.Visible = True
Else
    Me.Visible = False
    bDocked = False
    SetParent Me.hWnd, 0
    Me.Move FormLeft * Screen.TwipsPerPixelX, FormTop * Screen.TwipsPerPixelY, lFloatingWidth, lFloatingHeight
    mdiNexIRC!picMobileMixer.Visible = False
    Me.Visible = True
    Call SetWindowWord(Me.hWnd, SWW_HPARENT, mdiNexIRC.hWnd)
End If
bMoving = False
StoreFormDimensions
mdiNexIRC.ActivateResize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub FormDragger1_FormDropped(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)"
End Sub

Private Sub FormDragger1_FormMoved(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim rct As RECT
bMoving = True
GetWindowRect mdiNexIRC!picMobileMixer.hWnd, rct
With rct
    .Left = .Left - 4
    .Top = .Top - 4
    .Right = .Right + 4
    .Bottom = .Bottom + 4
End With
If PtInRect(rct, FormLeft, FormTop) Then
    FormWidth = lDockedWidth / Screen.TwipsPerPixelX
    FormHeight = lDockedHeight / Screen.TwipsPerPixelY
Else
    FormWidth = lFloatingWidth / Screen.TwipsPerPixelX
    FormHeight = lFloatingHeight / Screen.TwipsPerPixelY
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub FormDragger1_FormMoved(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)"
End Sub
