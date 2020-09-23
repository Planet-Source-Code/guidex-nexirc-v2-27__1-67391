VERSION 5.00
Begin VB.UserControl ctlMovieX 
   ClientHeight    =   930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1020
   ScaleHeight     =   930
   ScaleWidth      =   1020
   Begin VB.Frame fraPlayback 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Menu mnuControlMenu 
      Caption         =   "<Control Menu>"
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
      End
      Begin VB.Menu mnuResume 
         Caption         =   "Resume"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMute 
         Caption         =   "Mute"
      End
      Begin VB.Menu mnuUnmute 
         Caption         =   "Un-Mute"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenCDDoor 
         Caption         =   "Open CD Door"
      End
      Begin VB.Menu mnuCloseCDDoor 
         Caption         =   "Close CD Door"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuForward 
         Caption         =   "Forward"
      End
      Begin VB.Menu mnuRewind 
         Caption         =   "Rewind"
      End
      Begin VB.Menu mnuPosition 
         Caption         =   "Position"
      End
   End
End
Attribute VB_Name = "ctlMovieX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
Private lMovieX As New clsMovieX
Public Event AGDoubleClick()
Public Event AGRightClick()
Public Event AGMovieOpened(lFileName As String)
Public Event AGMovieClosed()
Public Event AGMovieStopped()
Public Event AGMoviePaused()
Public Event AGMovieResumed()
Public Event AGMoviePlay()
Public Event AGMouseMove(lLeft As Integer, lTop As Integer)
Enum eStatus
    sIdle = 0
    sPlaying = 1
    sOpen = 2
    sPaused = 3
    sStopped = 4
    sError = 5
End Enum
Private lStatus As eStatus
Private lBlackDropSide As Integer
Private lBlackDropSide2 As Integer
Private lBlackDropBottom As Integer
Private lBlackDropTop As Integer
'Private lRegistered As Boolean
Private lKeyRetries As Integer
Private lUseLess As Boolean

Public Sub Surf(lUrl As String, lhWnd As Long)
On Local Error GoTo ErrHandler
Dim msg As String, c As Integer, i As Integer, l As Long
l = ShellExecute(lhWnd, vbNullString, lUrl, vbNullString, "C:\", SW_SHOWNORMAL)
Exit Sub
ErrHandler:
    Err.Clear
End Sub

Public Sub OpenMovieDialog(lForm As Form, lFilter As String, lTitle As String, lInitDir As String)
On Local Error GoTo ErrHandler
Dim msg As String
msg = OpenDialog(lForm, lFilter, lTitle, lInitDir)
If Len(msg) <> 0 And DoesFileExist(msg) = True Then
    OpenMovie msg
End If
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub

'Public Sub Authorize(lUserName As String, lPassword As String)
'On Local Error GoTo ErrHandler
'Dim m As Boolean
'lKeyRetries = lKeyRetries + 1
'If lKeyRetries = 4 Then
'    lUseLess = True
'    lRegistered = False
'    lKeyRetries = 3
'    MsgBox "You have reached the maximum number of authorize retries, unable to try again.", vbExclamation
'    Exit Sub
'End If
'm = TestKeyValid(lUserName, lPassword)
'If m = False Then
'    lRegistered = False
'    MsgBox "Authorize failed, invalid key provided", vbExclamation
'Else
'    lRegistered = True
'End If
'Exit Sub
'ErrHandler:
'    MsgBox Err.Description
'    Err.Clear
'End Sub

Public Sub SetMovieWindowByhWnd(lhWnd As Long)
On Local Error GoTo ErrHandler
'If lUseLess = True Then Exit sub
If lhWnd <> 0 Then lMovieX.SetMovieWindow lhWnd
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub

Public Function GetStatus() As eStatus
On Local Error GoTo ErrHandler
'If lUseLess = True Then Exit Function
GetStatus = lStatus
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Sub Mute(lToggle As Boolean)
On Local Error GoTo ErrHandler
'If lUseLess = True Then Exit sub
If lToggle = True Then
    lMovieX.SetAudioOn
Else
    lMovieX.SetAudioOff
End If
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub

Public Sub OpenCDDoor()
On Local Error GoTo ErrHandler
'If lUseLess = True Then Exit sub
lMovieX.SetDoorOpen
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub

Public Sub CloseDoor()
On Local Error GoTo ErrHandler
'If lUseLess = True Then Exit sub
lMovieX.SetDoorClosed
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub

Public Function GetVolume() As String
On Local Error GoTo ErrHandler
'If lUseLess = True Then Exit Function
GetVolume = lMovieX.GetVolume
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function GetFramesPerSecond() As String
On Local Error GoTo ErrHandler
'If lUseLess = True Then Exit Function
GetFramesPerSecond = lMovieX.GetFramePerSecRate
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function RewindSeconds(lNumSeconds As Long) As String
On Local Error GoTo ErrHandler
'If lUseLess = True Then Exit Function
lMovieX.RewindBySeconds lNumSeconds
RewindSeconds = lMovieX.GetStatus
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function ForwardSeconds(lNumSeconds As Long) As String
On Local Error Resume Next
'If lUseLess = True Then Exit Function
lMovieX.ForwardBySeconds lNumSeconds
ForwardSeconds = lMovieX.GetStatus
End Function

Public Function ForwardFrames(lNumFrames As Long) As String
On Local Error Resume Next
'If lUseLess = True Then Exit Function
lMovieX.ForwardByFrames lNumFrames
ForwardFrames = lMovieX.GetStatus
End Function

Public Function RewindFrames(lNumFrames As Long) As String
On Local Error Resume Next
'If lUseLess = True Then Exit Function
lMovieX.RewindByFrames lNumFrames
RewindFrames = lMovieX.GetStatus
End Function

Public Sub FullScreen()
On Local Error Resume Next
'If lUseLess = True Then Exit sub
lMovieX.PlayFullScreen
End Sub

Public Function SetSize(lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long) As String
On Local Error Resume Next
'If lUseLess = True Then Exit Function
lMovieX.SizeLocateMovie lLeft, lTop, lWidth, lHeight
SetSize = lMovieX.CheckError
End Function

Public Function CloseMovie() As String
On Local Error GoTo ErrHandler
'If lUseLess = True Then Exit Function
lMovieX.CloseMovie
CloseMovie = lMovieX.GetStatus
lStatus = sIdle
RaiseEvent AGMovieClosed
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function OpenMovie(lFileName As String) As String
On Local Error GoTo ErrHandler
'DisplayReginfo fraPlayback.hWnd
'If lUseLess = True Then Exit Function
If Len(lFileName) <> 0 Then
    lMovieX.lCurrentFile = lFileName
    lMovieX.OpenMovieWindow fraPlayback.hWnd, "child"
    OpenMovie = lMovieX.GetStatus
    OpenMovie = lMovieX.GetStatus
    RaiseEvent AGMovieOpened(lFileName)
    lStatus = sOpen
End If
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Sub SetBlackDrop(lTopMargin As Integer, lSideMargin As Integer, lSideMargin2 As Integer, lBottomMargin As Integer)
On Local Error GoTo ErrHandler
'If lUseLess = True Then Exit sub
lBlackDropBottom = lBottomMargin
lBlackDropTop = lTopMargin
lBlackDropSide = lSideMargin
lBlackDropSide2 = lSideMargin2
UserControl_Resize
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub

Public Function ResumeMovie() As String
On Local Error GoTo ErrHandler
'If lUseLess = True Then Exit Function
lMovieX.ResumeMovie
lMovieX.TimeOut 1
ResumeMovie = lMovieX.GetStatus
lStatus = sPlaying
RaiseEvent AGMovieResumed
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function PauseMovie() As String
On Local Error GoTo ErrHandler
'If lUseLess = True Then Exit Function
lMovieX.PauseMovie
lMovieX.TimeOut 1
PauseMovie = lMovieX.GetStatus
RaiseEvent AGMoviePaused
lStatus = sPaused
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function SetVolume(lVolume As Long) As String
On Local Error GoTo ErrHandler
'If lUseLess = True Then Exit Function
lMovieX.SetVolume lVolume
SetVolume = lMovieX.GetStatus
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function PlayMovie() As String
On Local Error GoTo ErrHandler
'If lUseLess = True Then Exit Function
lMovieX.PlayMovie
lMovieX.TimeOut 0.5
PlayMovie = lMovieX.GetStatus
lStatus = sPlaying
RaiseEvent AGMoviePlay
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function StopMovie() As String
On Local Error GoTo ErrHandler
'If lUseLess = True Then Exit Function
lMovieX.StopMovie
StopMovie = lMovieX.GetStatus
RaiseEvent AGMovieStopped
lStatus = sStopped
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function ChangeMoviePosition(lSecond As Long) As String
On Local Error Resume Next
'If lUseLess = True Then Exit Function
lMovieX.SetPositionTo lSecond
ChangeMoviePosition = lMovieX.GetStatus
End Function

Public Function ReturnMovieFrames() As Long
On Local Error Resume Next
'If lUseLess = True Then Exit Function
ReturnMovieFrames = lMovieX.GetLengthInFrames
End Function

Public Function ReturnTotalSeconds() As Long
On Local Error Resume Next
'If lUseLess = True Then Exit Function
ReturnTotalSeconds = lMovieX.GetLengthInSec
End Function

Public Function ReturnCurrentPosition() As Long
On Local Error Resume Next
'If lUseLess = True Then Exit Function
ReturnCurrentPosition = lMovieX.GetPositionInSec
End Function

Public Function ChangePlayRate(lValue As Long)
On Local Error Resume Next
'If lUseLess = True Then Exit Function
lMovieX.SetSpeed lValue
ChangePlayRate = lMovieX.GetStatus
End Function

Private Sub mnuCloseCDDoor_Click()
On Local Error Resume Next
'If lUseLess = True Then Exit sub
CloseDoor
End Sub

Private Sub mnuForward_Click()
Dim lSec As Long
'If lUseLess = True Then Exit sub
lSec = InputBox("Forward How Many Seconds?", "Forward")
ForwardSeconds lSec
End Sub

Private Sub mnuMute_Click()
On Local Error Resume Next
'If lUseLess = True Then Exit sub
Mute True
End Sub

Private Sub mnuOpenCDDoor_Click()
On Local Error Resume Next
'If lUseLess = True Then Exit sub
OpenCDDoor
End Sub

Private Sub mnuPause_Click()
On Local Error Resume Next
'If lUseLess = True Then Exit sub
PauseMovie
End Sub

Private Sub mnuPlay_Click()
On Local Error Resume Next
'If lUseLess = True Then Exit sub
PlayMovie
End Sub

Private Sub mnuPosition_Click()
Dim lPos As Long
'If lUseLess = True Then Exit sub
lPos = InputBox("Enter Position in Seconds")
ChangeMoviePosition lPos
End Sub

Private Sub mnuResume_Click()
On Local Error Resume Next
'If lUseLess = True Then Exit sub
ResumeMovie
End Sub

Private Sub mnuRewind_Click()
Dim lSec As Long
'If lUseLess = True Then Exit sub
lSec = InputBox("Enter Number of Seconds to Rewind")
RewindSeconds lSec
End Sub

Private Sub mnuStop_Click()
On Local Error Resume Next
'If lUseLess = True Then Exit sub
StopMovie
End Sub

Private Sub mnuUnmute_Click()
On Local Error Resume Next
'If lUseLess = True Then Exit sub
Mute False
End Sub

Private Sub fraPlayback_DblClick()
On Local Error Resume Next
'If lUseLess = True Then Exit sub
RaiseEvent AGDoubleClick
End Sub

Public Sub fraPlayback_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
'If lUseLess = True Then Exit sub
If Button = 2 Then RaiseEvent AGRightClick
End Sub

Private Sub fraPlayback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
'If lUseLess = True Then Exit sub
RaiseEvent AGMouseMove(Int(X), Int(Y))
End Sub

Private Sub UserControl_DblClick()
On Local Error Resume Next
'If lUseLess = True Then Exit sub
RaiseEvent AGDoubleClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
'If lUseLess = True Then Exit sub
If Button = 2 Then RaiseEvent AGRightClick
End Sub

Private Sub UserControl_Resize()
On Local Error Resume Next
'If lUseLess = True Then Exit sub
If lBlackDropSide <> 0 Then
    fraPlayback.Left = lBlackDropSide
    fraPlayback.Width = UserControl.ScaleWidth - (fraPlayback.Left + lBlackDropSide2)
    If lBlackDropTop <> 0 Then
        fraPlayback.Top = lBlackDropTop
        If lBlackDropBottom <> 0 Then
            fraPlayback.Height = UserControl.ScaleHeight - (lBlackDropTop + lBlackDropBottom)
        Else
            fraPlayback.Height = UserControl.ScaleHeight - lBlackDropTop
        End If
    End If

Else
    fraPlayback.Width = UserControl.ScaleWidth
    fraPlayback.Height = UserControl.ScaleHeight
End If
End Sub

Private Sub UserControl_Terminate()
On Local Error Resume Next
'If lUseLess = True Then Exit sub
lMovieX.CloseMovie
End Sub
