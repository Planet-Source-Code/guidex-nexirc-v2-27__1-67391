Attribute VB_Name = "mdlAlarm"
Option Explicit
Private Type gAlarm
    aEnabled As Boolean
    aAlarmTime As String
    aAlarmDate As String
    aTime As String
    aAudio As String
    aDate As String
End Type
Private lAlarm As gAlarm

Public Sub ToggleAlarm(lOn As Boolean, lTime As String, lDate As String, lAudio As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lOn = True Then
    frmAlarm.tmrAlarm.Enabled = True
    lAlarm.aAudio = lAudio
    lAlarm.aEnabled = True
    lAlarm.aAlarmTime = lTime
    lAlarm.aAlarmDate = lDate
ElseIf lOn = False Then
    lAlarm.aAudio = ""
    frmAlarm.tmrAlarm.Enabled = False
    lAlarm.aEnabled = False
    lAlarm.aAlarmTime = lTime
    lAlarm.aAlarmDate = lDate
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ToggleAlarm(lOn As Boolean, lTime As String, lDate As String)"
End Sub

Public Sub SetTime()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lAlarm.aDate = Date
lAlarm.aTime = Time
frmAlarm.Caption = "NexIRC - Alarm (" & Time & ")"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetTime()"
End Sub

Public Sub CheckAlarm()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lAlarm.aEnabled = True Then
    If Len(lAlarm.aAlarmTime) <> 0 And Len(lAlarm.aAlarmDate) <> 0 Then
        If LCase(Time) = LCase(lAlarm.aAlarmTime) And LCase(Date) = LCase(lAlarm.aAlarmDate) Then
            frmAlarm.tmrAlarm.Enabled = False
            PlayFile lFiles.fFile(FindFileIndexByFilename(lAlarm.aAudio)).fFilename
            Unload frmAlarm
        End If
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub CheckAlarm()"
End Sub
