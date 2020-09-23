Attribute VB_Name = "mdlMixer"
Option Explicit
#If Win32 Then
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
    Private Declare Function ShellExecute Lib "shell.dll" (ByVal hWnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
#End If
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
Private Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long
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
Type MIXERCAPS
     wMid As Integer
     wPid As Integer
     vDriverVersion As Long
     szPname As String * MAXPNAMELEN
     fdwSupport As Long
     cDestinations As Long
End Type
Type Target
     dwType As Long
     dwDeviceID As Long
     wMid As Integer
     wPid As Integer
     vDriverVersion As Long
     szPname As String * MAXPNAMELEN
End Type
Type MIXERLINE
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
Type MIXERLINECONTROLS
     cbStruct As Long
     dwLineID As Long
     dwControl As Long
     cControls As Long
     cbmxctrl As Long
     pamxctrl As Long
End Type
Type MIXERCONTROL
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
Type MIXERCONTROLDETAILS
     cbStruct As Long
     dwControlID As Long
     cChannels As Long
     Item As Long
     cbDetails As Long
     paDetails As Long
End Type
Type MIXERCONTROLDETAILS_BOOLEAN
     fValue As Long
End Type
Type MIXERCONTROLDETAILS_LISTTEXT
     dwParam1 As Long
     dwParam2 As Long
     szName As String * MIXER_LONG_NAME_CHARS
End Type
Type MIXERCONTROLDETAILS_SIGNED
     lValue As Long
End Type
Type MIXERCONTROLDETAILS_UNSIGNED
     dwValue As Long
End Type
Type waveFormat
   wFormatTag As Integer
   nChannels As Integer
   nSamplesPerSec As Long
   nAvgBytesPerSec As Long
   nBlockAlign As Integer
   wBitsPerSample As Integer
   cbSize As Integer
End Type
Type mmioInfo
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
Type MMCKInfo
    ckid As Long
    ckSize As Long
    fccType As Long
    dwDataOffset As Long
    dwFlags As Long
End Type

Function GetMixerControl(ByVal hmixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, ByRef mxc As MIXERCONTROL) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mxlc As MIXERLINECONTROLS, mxl As MIXERLINE, hmem As Long, rc As Long
mxl.cbStruct = Len(mxl)
mxl.dwComponentType = componentType
rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEInfoF_COMPONENTTYPE)
If (MMSYSERR_NOERROR = rc) Then
    mxlc.cbStruct = Len(mxlc)
    mxlc.dwLineID = mxl.dwLineID
    mxlc.dwControl = ctrlType
    mxlc.cControls = 1
    mxlc.cbmxctrl = Len(mxc)
    hmem = GlobalAlloc(&H40, Len(mxc))
    mxlc.pamxctrl = GlobalLock(hmem)
    mxc.cbStruct = Len(mxc)
    rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
    If (MMSYSERR_NOERROR = rc) Then
        GetMixerControl = True
        CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
    Else
        GetMixerControl = False
    End If
    GlobalFree (hmem)
    Exit Function
End If
GetMixerControl = False
End Function

Function SetVolumeControl(ByVal hmixer As Long, mxc As MIXERCONTROL, ByVal volume As Long) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mxcd As MIXERCONTROLDETAILS, vol As MIXERCONTROLDETAILS_UNSIGNED, rc As Long, hmem As Long
mxcd.cbStruct = Len(mxcd)
mxcd.dwControlID = mxc.dwControlID
mxcd.cChannels = 1
mxcd.Item = 0
mxcd.cbDetails = Len(vol)
hmem = GlobalAlloc(&H40, Len(vol))
mxcd.paDetails = GlobalLock(hmem)
vol.dwValue = volume
CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
GlobalFree (hmem)
If (MMSYSERR_NOERROR = rc) Then
    SetVolumeControl = True
Else
    SetVolumeControl = False
End If
End Function

Function SetPANControl(ByVal hmixer As Long, mxc As MIXERCONTROL, ByVal volL As Long, ByVal volR As Long) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mxcd As MIXERCONTROLDETAILS, vol(1) As MIXERCONTROLDETAILS_UNSIGNED, rc As Long, hmem As Long
mxcd.Item = mxc.cMultipleItems
mxcd.dwControlID = mxc.dwControlID
mxcd.cbStruct = Len(mxcd)
mxcd.cbDetails = Len(vol(1))
mxcd.cChannels = 2
hmem = GlobalAlloc(&H40, Len(vol(1)))
mxcd.paDetails = GlobalLock(hmem)
vol(1).dwValue = volR
vol(0).dwValue = volL
CopyPtrFromStruct mxcd.paDetails, vol(1).dwValue, Len(vol(0)) * mxcd.cChannels
CopyPtrFromStruct mxcd.paDetails, vol(0).dwValue, Len(vol(1)) * mxcd.cChannels
rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
GlobalFree (hmem)
If (MMSYSERR_NOERROR = rc) Then
    SetPANControl = True
Else
    SetPANControl = False
End If
End Function

Function unSetMuteControl(ByVal hmixer As Long, mxc As MIXERCONTROL, ByVal unmute As Long) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mxcd As MIXERCONTROLDETAILS, vol As MIXERCONTROLDETAILS_UNSIGNED, rc As Long, hmem As Long
mxcd.cbStruct = Len(mxcd)
mxcd.dwControlID = mxc.dwControlID
mxcd.cChannels = 1
mxcd.Item = 0
mxcd.cbDetails = Len(vol)
hmem = GlobalAlloc(&H40, Len(vol))
mxcd.paDetails = GlobalLock(hmem)
vol.dwValue = unmute
CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
GlobalFree (hmem)
If (MMSYSERR_NOERROR = rc) Then
    unSetMuteControl = True
Else
    unSetMuteControl = False
End If
End Function

Function SetMuteControl(ByVal hmixer As Long, mxc As MIXERCONTROL, mute As Boolean) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mxcd As MIXERCONTROLDETAILS, vol As MIXERCONTROLDETAILS_UNSIGNED, rc As Long, hmem As Long
mxcd.cbStruct = Len(mxcd)
mxcd.dwControlID = mxc.dwControlID
mxcd.cChannels = 1
mxcd.Item = 0
mxcd.cbDetails = Len(vol)
hmem = GlobalAlloc(&H40, Len(vol))
mxcd.paDetails = GlobalLock(hmem)
CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
GlobalFree (hmem)
If (MMSYSERR_NOERROR = rc) Then
    SetMuteControl = True
Else
    SetMuteControl = False
End If
End Function

Function GetVolumeControlValue(ByVal hmixer As Long, mxc As MIXERCONTROL) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mxcd As MIXERCONTROLDETAILS, vol As MIXERCONTROLDETAILS_UNSIGNED, rc As Long, hmem As Long
mxcd.cbStruct = Len(mxcd)
mxcd.dwControlID = mxc.dwControlID
mxcd.cChannels = 1
mxcd.Item = 0
mxcd.cbDetails = Len(vol)
mxcd.paDetails = 0
hmem = GlobalAlloc(&H40, Len(vol))
mxcd.paDetails = GlobalLock(hmem)
rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
CopyStructFromPtr vol, mxcd.paDetails, Len(vol)
GlobalFree (hmem)
If (MMSYSERR_NOERROR = rc) Then
    GetVolumeControlValue = vol.dwValue
Else
    GetVolumeControlValue = -1
End If
End Function
