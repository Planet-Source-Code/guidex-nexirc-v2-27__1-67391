Attribute VB_Name = "mdlColorRTF"
Option Explicit

Public Sub DoColorSep(lTextBox As TBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lTextBox.NewLine "1.:"
End Sub

Public Sub DoColor(lTextBox As TBox, lData As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lData = Replace(lData, vbCrLf, "")
lData = Replace(lData, Chr(13), "")
lData = Replace(lData, Chr(10), "")
PlayWav App.Path & "\data\sounds\tdraw" & GetRnd(9) & ".wav", SND_ASYNC
If lSettings.sTimeStamping = True Then lData = "15|" & "14" & Time$ & "15" & "| " & "" & Color.Normal & ":. " & lData
lTextBox.NewLine Trim(lData)
End Sub

Public Sub DoColorLines(lTextBox As TBox, lData As String)
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
