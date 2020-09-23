Attribute VB_Name = "mdlMultimedia"
Option Explicit
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type gFile
    fFilename As String
End Type
Private Type gFiles
    fIndex As Integer
    fFile(10000) As gFile
    fCount As Long
End Type
Dim glo_from As Long
Dim glo_to As Long
Dim glo_AliasName As String
Dim glo_hWnd As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Global lFiles As gFiles

Public Function OpenMultimedia(hWnd As Long, AliasName As String, FileName As String, typeDevice As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim cmdToDo As String * 255
Dim dwReturn As Long
Dim ret As String * 128
Dim tmp As String * 255
Dim lenShort As Long
Dim ShortPathAndFile As String
Const WS_CHILD = &H40000000
lenShort = GetShortPathName(FileName, tmp, 255)
ShortPathAndFile = Left$(tmp, lenShort)
cmdToDo = "open " & ShortPathAndFile & " type " & typeDevice & " Alias " & AliasName & " parent " & hWnd & " Style " & WS_CHILD
dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    OpenMultimedia = ret: Exit Function
End If
OpenMultimedia = "Success"
End Function

Public Function PlayMultimedia(AliasName As String, from_where As String, to_where As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If from_where = vbNullString Then from_where = 0
If to_where = vbNullString Then to_where = GetTotalframes(AliasName)
If AliasName = glo_AliasName Then
    glo_from = from_where
    glo_to = to_where
End If
Dim cmdToDo As String * 128
Dim dwReturn As Long
Dim ret As String * 128
cmdToDo = "play " & AliasName & " from " & from_where & " to " & to_where
dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    PlayMultimedia = ret
    Exit Function
End If
PlayMultimedia = "Success"
End Function

Public Function CloseMultimedia(AliasName As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Close " & AliasName, 0&, 0&, 0&)
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    CloseMultimedia = ret
    Exit Function
End If
If AliasName = glo_AliasName Then
KillTimer glo_hWnd, 500
End If
CloseMultimedia = "Success"
End Function

Public Function PauseMultimedia(AliasName As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Pause " & AliasName, 0&, 0&, 0&)
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    PauseMultimedia = ret
    Exit Function
End If
PauseMultimedia = "Success"
End Function

Public Function StopMultimedia(AliasName As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Stop " & AliasName, 0&, 0&, 0&)
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    StopMultimedia = ret
    Exit Function
End If
StopMultimedia = "Success"
End Function

Public Function ResumeMultimedia(AliasName As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Resume " & AliasName, 0&, 0&, 0&)
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    ResumeMultimedia = ret
    Exit Function
End If
ResumeMultimedia = "Success"
End Function

Public Function GetStatusMultimedia(AliasName As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim dwReturn As Long
Dim Status As String * 128
Dim ret As String * 128
dwReturn = mciSendString("status " & AliasName & " mode", Status, 128, 0&)
If Not dwReturn = 0 Then
    GetStatusMultimedia = "ERROR"
    Exit Function
End If
Dim i As Integer
Dim CharA As String
Dim RChar As String
RChar = Right$(Status, 1)
For i = 1 To Len(Status)
    CharA = Mid(Status, i, 1)
    If CharA = RChar Then Exit For
    GetStatusMultimedia = GetStatusMultimedia + CharA
Next i
End Function

Public Function GetTotalframes(AliasName As String) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim dwReturn As Long
Dim Total As String * 128
dwReturn = mciSendString("set " & AliasName & " time format frames", Total, 128, 0&)
dwReturn = mciSendString("status " & AliasName & " length", Total, 128, 0&)
If Not dwReturn = 0 Then
    GetTotalframes = -1
    Exit Function
End If
GetTotalframes = Val(Total)
End Function

Public Function GetTotalTimeByMS(AliasName As String) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim dwReturn As Long
Dim TotalTime As String * 128
dwReturn = mciSendString("set " & AliasName & " time format ms", TotalTime, 128, 0&)
dwReturn = mciSendString("status " & AliasName & " length", TotalTime, 128, 0&)
mciSendString "set " & AliasName & " time format frames", 0&, 0&, 0&
If Not dwReturn = 0 Then
    GetTotalTimeByMS = -1
    Exit Function
End If
GetTotalTimeByMS = Val(TotalTime)
End Function

Public Function MoveMultimedia(AliasName As String, to_where As Long) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("seek " & AliasName & " to " & to_where, 0&, 0&, 0&)
mciSendString "Play " & AliasName, 0&, 0&, 0&
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    MoveMultimedia = ret
    Exit Function
End If
MoveMultimedia = "Success"
End Function

Public Function GetCurrentMultimediaPos(AliasName As String) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim dwReturn As Long
Dim POS As String * 128
dwReturn = mciSendString("status " & AliasName & " position", POS, 128, 0&)
If Not dwReturn = 0 Then
    GetCurrentMultimediaPos = -1
    Exit Function
End If
GetCurrentMultimediaPos = Val(POS)
End Function

Public Function PutMultimedia(hWnd As Long, AliasName As String, Left As Long, Top As Long, Width As Long, Height As Long) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim dwReturn As Long
Dim ret As String * 128
If Width = 0 Or Height = 0 Then
    Dim rec As RECT
    Call GetWindowRect(hWnd, rec)
    Width = rec.Right - rec.Left
    Height = rec.Bottom - rec.Top
End If
dwReturn = mciSendString("put " & AliasName & " window at " & Left & " " & Top & " " & Width & " " & Height, 0&, 0&, 0&)
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    PutMultimedia = ret
    Exit Function
End If
PutMultimedia = "Success"
End Function

Public Function GetPercent(AliasName As String) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim TotalFrames As Long
Dim currframe As Long
TotalFrames = GetTotalframes(AliasName)
currframe = GetCurrentMultimediaPos(AliasName)
If TotalFrames = -1 Or currframe = -1 Then
    GetPercent = -1
    Exit Function
End If
GetPercent = currframe * 100 / TotalFrames
End Function

Public Function GetFramesPerSecond(AliasName As String) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim TotalFrames As Long
Dim TotalTime As Long
TotalTime = GetTotalTimeByMS(AliasName)
TotalFrames = GetTotalframes(AliasName)
If TotalFrames = -1 Or TotalTime = -1 Then
    GetFramesPerSecond = -1
    Exit Function
End If
GetFramesPerSecond = TotalFrames / (TotalTime / 1000)
End Function

Public Function GetSize(AliasName As String, CxOrCy As String) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Not CxOrCy = "cx" And Not CxOrCy = "cy" Then GetSize = -1: Exit Function
Dim dwReturn As Long
Dim Size As String * 128
Dim s1, s2, s3, Width, Height As Long
dwReturn = mciSendString("Where " & AliasName & " destination", Size, 128, 0&)
If Not dwReturn = 0 Then
    GetSize = -1
    Exit Function
End If
s1 = InStr(1, Size, " "): s2 = InStr(s1 + 1, Size, " "): s1 = InStr(s2 + 1, Size, " ")
Width = Mid(Size, s2, s1 - s2): Height = Mid(Size, s1 + 1)
If CxOrCy = "cx" Then
GetSize = Width
ElseIf CxOrCy = "cy" Then
GetSize = Height
End If
End Function

Public Function CloseAll() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Close All", 0&, 0&, 0&)
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    CloseAll = ret
    Exit Function
End If
CloseAll = "Success"
End Function

Public Function ChannelsControl(AliasName As String, Channel As String, OnOrOFF As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim cmdToDo As String * 128
Dim dwReturn As Long
Dim ret As String * 128
cmdToDo = "set " & AliasName & " audio " & Channel & " " & OnOrOFF
dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    ChannelsControl = ret
    Exit Function
End If
ChannelsControl = "Success"
End Function

Public Function AreMultimediaAtEnd(AliasName As String, lastFrame As Long) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim currpos As Long
If lastFrame = 0 Then lastFrame = GetTotalframes(AliasName)
currpos = Val(GetCurrentMultimediaPos(AliasName))
If currpos = -1 Or lastFrame = -1 Then
    AreMultimediaAtEnd = False
    Exit Function
End If
If lastFrame = currpos Or (lastFrame - 1) < currpos Then
AreMultimediaAtEnd = True
Else
AreMultimediaAtEnd = False
End If
End Function

Public Function SetAutoRepeat(hWnd As Long, AliasName As String, first_frame As String, last_frame As String, autoTrueOrFalse As Boolean) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim result As String
If first_frame = vbNullString Then first_frame = 0
If last_frame = vbNullString Then last_frame = GetTotalframes(AliasName)
glo_from = first_frame
glo_to = last_frame
glo_hWnd = hWnd
If autoTrueOrFalse = True Then
    glo_AliasName = AliasName
    result = SetTimer(hWnd, 500, 100, AddressOf TimerFunction)
Else
    glo_AliasName = vbNullString
    result = KillTimer(hWnd, 500)
End If
If result = 0 Then
    SetAutoRepeat = False
Else
    SetAutoRepeat = True
End If
End Function

Sub TimerFunction()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim currpos As Long
Dim result As String
currpos = Val(GetCurrentMultimediaPos(glo_AliasName))
If currpos = -1 Then Exit Sub
If Val(glo_to) = currpos Or (Val(glo_to) - 1) < currpos Then
    result = PlayMultimedia(glo_AliasName, Str(glo_from), Str(glo_to))
    If Not result = "Success" Then KillTimer glo_hWnd, 500
End If
End Sub

Public Sub SetDefaultDevice(typeDevice As String, drvDefaultDevice As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String * 255, msg3 As String
msg = GetWindowsDirectory(msg2, 255)
msg3 = Left$(msg2, msg)
'WriteINI "MCI", typeDevice, drvDefaultDevice, msg3 & "\" & "system.ini"
End Sub

Public Function GetDefaultDevice(typeDevice As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String * 255, msg2 As String, msg3 As String
msg2 = GetWindowsDirectory(msg, 255)
msg3 = Left$(msg, msg2)
'msg2 = ReadINI("MCI", typeDevice, "None", msg, 255, msg3 & "\" & "system.ini")
GetDefaultDevice = Left$(msg, msg2)
End Function
