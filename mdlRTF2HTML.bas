Attribute VB_Name = "mdlRTF2HTML"
Option Explicit
Private sCodes() As String
Private sRTFLine As String
Private sRTFWord As String
Private sRTFLeft As String
Private sRTFRight As String
Private sHoldHtml As String
Private bBullet As Boolean
Private iBracketLocation As Integer
Private iFindBackSlash As Integer
Private iFindNextSlash As Integer
Private iFindRightBracket As Integer
Private iFindSpace As Integer
Private iFindEOL As Integer
Private iCodeCounter As Integer
Private i As Integer

Public Function RTFtoHTML(sRTFText As String, Optional sOptions As String) As String
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
Dim sRTFRemaining As String
sRTFRemaining = sRTFText
subClearScodes
iFindRightBracket = 1
While iFindRightBracket > 0
    iFindRightBracket = InStr(iFindRightBracket + 1, sRTFRemaining, "}}")
    If iFindRightBracket > 0 Then iBracketLocation = iFindRightBracket
Wend
sRTFRemaining = Mid(sRTFRemaining, iBracketLocation + 2)
While sRTFRemaining <> ""
    sRTFLine = funGetNextLine(sRTFRemaining)
    sRTFRemaining = Mid(sRTFRemaining, Len(sRTFLine) + 1)
    subProcessRTFLine
Wend
RTFtoHTML = sHoldHtml
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Function RTFtoHTML(sRTFText As String, Optional sOptions As String) As String"
    Err.Clear
End Function

Private Sub subProcessRTFLine()
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
iFindBackSlash = 1
While sRTFLine > ""
    While iFindBackSlash > 0
        iFindBackSlash = InStr(sRTFLine, "\")
        If iFindBackSlash > 1 Then
            sHoldHtml = sHoldHtml & Left(sRTFLine, iFindBackSlash - 1)
        End If
        Select Case iFindBackSlash
            Case Is > 0
                iFindNextSlash = InStr(iFindBackSlash + 1, sRTFLine, "\")
                Select Case iFindNextSlash
                    Case Is > 0
                        sRTFWord = Left(sRTFLine, iFindNextSlash - 1)
                        sRTFLine = Mid(sRTFLine, Len(sRTFWord) + 1)
                        iFindSpace = InStr(sRTFWord, " ")
                       If iFindSpace > 0 Then
                            sRTFRight = Mid(sRTFWord, iFindSpace + 1)
                            sRTFWord = Left(sRTFWord, iFindSpace - 1)
                       End If
                    Case 0
                        iFindNextSlash = InStr(iFindBackSlash + 1, sRTFLine, " ")
                        Select Case iFindNextSlash
                            Case Is > 0
                                sRTFWord = Left(sRTFLine, iFindNextSlash - 1)
                                sRTFLine = Mid(sRTFLine, Len(sRTFWord) + 1)
                            Case 0
                        End Select
                End Select
            Case 0
                sRTFWord = sRTFLine
                sRTFLine = ""
        End Select
        subProcessRTFWord
    Wend
Wend
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub subProcessRTFLine()"
    Err.Clear
End Sub

Private Sub subProcessRTFWord()
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
Select Case sRTFWord
    Case "\i"
        sHoldHtml = sHoldHtml & "<i>"
        subPushCode "</i>"
    Case "\i0"
        subPopCode
    Case "\b"
        sHoldHtml = sHoldHtml & "<b>"
        subPushCode "</b>"
    Case "\b0"
        subPopCode
    Case "\ul"
        sHoldHtml = sHoldHtml & "<u>"
        subPushCode "</u>"
    Case "\ulnone"
        subPopCode
    Case "\'b7"
        If Not bBullet Then
            bBullet = True
            subPushCode "</ul>"
            subPushCode "</li>"
            sHoldHtml = sHoldHtml & "<ul><li>"
        Else
            sHoldHtml = sHoldHtml & "</li><li>"
        End If
    Case "\par"
        If bBullet And (InStr(sRTFLine, "\'") = 0) Then
            bBullet = False
            iCodeCounter = 1
            subPopCode
        End If
    Case vbCrLf
        sHoldHtml = sHoldHtml & "<br>"
    Case Else
        sRTFLeft = Left(sRTFWord, 1)
        Select Case sRTFLeft
            Case "\"
            Case Else
                If InStr(sRTFWord, "}") Then
                    sRTFWord = ""
                ElseIf Right(sRTFWord, 2) = vbCrLf Then
                    sRTFWord = sRTFWord & "<BR>"
                End If
                sHoldHtml = sHoldHtml & sRTFWord
        End Select
End Select
If Len(sRTFRight) > 0 Then
    sHoldHtml = sHoldHtml & sRTFRight
    sRTFRight = ""
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub subProcessRTFWord()"
    Err.Clear
End Sub

Private Function funGetNextLine(sRTF As String) As String
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
Dim sHoldLine As String
iFindEOL = InStr(sRTF, vbCrLf)
If iFindEOL > 0 Then
    sHoldLine = Left(sRTF, iFindEOL + 1)
Else
    sHoldLine = sRTF
End If
funGetNextLine = sHoldLine
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Function funGetNextLine(sRTF As String) As String"
    Err.Clear
End Function

Private Sub subPushCode(sRTFString As String)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
Dim lUbound As Long
lUbound = UBound(sCodes)
ReDim Preserve sCodes(UBound(sCodes) + 1)
sCodes(UBound(sCodes)) = sRTFString
iCodeCounter = iCodeCounter + 1
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub subPushCode(sRTFString As String)"
    Err.Clear
End Sub

Private Sub subPopCode()
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
iCodeCounter = iCodeCounter - 1
If iCodeCounter = 0 Then
    For i = UBound(sCodes) To 1 Step -1
        sHoldHtml = sHoldHtml & sCodes(i)
    Next i
    subClearScodes
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub subPopCode()"
    Err.Clear
End Sub

Private Sub subClearScodes()
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
ReDim sCodes(0)
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub subClearScodes()"
    Err.Clear
End Sub
