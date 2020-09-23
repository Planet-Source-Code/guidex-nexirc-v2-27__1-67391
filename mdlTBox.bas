Attribute VB_Name = "mdlTBox"
Option Explicit
Private Declare Function QueryPerformanceCounter Lib "kernel32.dll" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As gPointAPI) As Long
Private Type gPointAPI
    X As Long
    Y As Long
End Type
Private lSizeV As gPointAPI
Private lStartTime As Currency
Private lPerfFreq As Currency
Private Type gColor
    cRGB As String
End Type
Private Type gColors
    cColor(15) As gColor
End Type
Private lColors As gColors

Public Function IRCCodeToRGB(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case lIndex
Case 0
    IRCCodeToRGB = vbWhite
Case 1
    IRCCodeToRGB = vbBlack
Case 2
    IRCCodeToRGB = RGB(0, 0, 140)
Case 3
    IRCCodeToRGB = RGB(0, 140, 0)
Case 4
    IRCCodeToRGB = vbRed
Case 5
    IRCCodeToRGB = RGB(110, 65, 0)
Case 6
    IRCCodeToRGB = RGB(140, 0, 140)
Case 7
    IRCCodeToRGB = RGB(248, 146, 0)
Case 8
    IRCCodeToRGB = vbYellow
Case 9
    IRCCodeToRGB = vbGreen
Case 10
    IRCCodeToRGB = RGB(0, 140, 140)
Case 11
    IRCCodeToRGB = RGB(0, 255, 255)
Case 12
    IRCCodeToRGB = vbBlue
Case 13
    IRCCodeToRGB = vbMagenta
Case 14
    IRCCodeToRGB = RGB(140, 140, 140)
Case 15
    IRCCodeToRGB = RGB(200, 200, 200)
End Select
End Function

Public Function GetTextWidth2(ByVal tText As String, ByVal hw As Long) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
GetTextExtentPoint32 hw, tText, Len(tText), lSizeV
GetTextWidth2 = lSizeV.X
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function GetTextWidth2(ByVal tText As String, ByVal hw As Long) As Integer"
End Function

Public Function Start() As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
QueryPerformanceCounter lStartTime
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function Start() As Long"
End Function

Public Function Finish() As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim cCurrentTime As Currency
QueryPerformanceCounter cCurrentTime
Finish = 1000 * (cCurrentTime - lStartTime) / lPerfFreq
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function Finish() As Long"
End Function

Public Sub DoColorSep(lTextBox As ctlTBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lTextBox.NewLine "1.:"
End Sub

Public Sub DoColor(lTextBox As ctlTBox, lData As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, n As Integer, l As Integer, msg(30) As String, lCurrentColor As String, lBold As Boolean
lData = Replace(lData, vbCrLf, "")
lData = Replace(lData, Chr(13), "")
lData = Replace(lData, Chr(10), "")
PlayWav App.Path & "\data\sounds\tdraw" & GetRnd(9) & ".wav", SND_ASYNC
If lSettings.sTimeStamping = True Then lData = "15|" & "14" & Time$ & "15" & "| " & "" & Color.Normal & ":. " & lData
l = lTextBox.ReturnWidth()
l = l / Screen.TwipsPerPixelX * 2.7
For i = 0 To 30
    If Len(lData) > l Then
        msg(i) = Left(lData, l)
        If msg(i) = "" Then
            Stop
        End If
        lData = Right(lData, Len(lData) - Len(msg(i)))
        For n = 0 To Len(msg(i))
            If Right(msg(i), 1) = " " Then
                Exit For
            Else
                lData = Right(msg(i), 1) & lData
                msg(i) = Left(msg(i), Len(msg(i)) - 1)
            End If
        Next n
    Else
        msg(i) = lData
        Exit For
    End If
Next i
For i = 0 To UBound(msg)
    If Len(msg(i)) <> 0 Then
        lTextBox.NewLine Trim(msg(i))
    Else
        Exit For
    End If
Next i
End Sub

Public Sub DoColorLines(lTextBox As ctlTBox, lData As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String
lData = Replace(lData, vbCrLf, "")
lData = Replace(lData, Chr(13), "")
lData = Replace(lData, Chr(10), "")
msg = lData
If Len(lData) <> 0 Then
    Do Until Len(msg) = 0
        If InStr(msg, Chr(13)) Then
            msg2 = Trim(Left(msg, 1) & Parse(msg, Left(msg, 1), Chr(13)))
            msg = Trim(Right(msg, Len(msg) - Len(msg2) - 2))
        Else
            msg2 = Trim(msg)
            msg = ""
        End If
        If Len(msg2) <> 0 Then
            DoColor lTextBox, msg2
        End If
    Loop
End If
End Sub

Public Sub ActiveWindowDoColor(lText As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(lText) <> 0 Then Call DoColor(mdiNexIRC.ActiveForm.txtIncoming, lText)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ActiveWindowDoColor(lText As String)"
End Sub
