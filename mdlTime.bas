Attribute VB_Name = "mdlTime"
Option Explicit
Private Type gSystemTime
    sYear As Integer
    sMonth As Integer
    sDayOfWeek As Integer
    sDay As Integer
    sHour As Integer
    sMinute As Integer
    sSecond As Integer
    sMilliseconds As Integer
End Type
Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(32) As Integer
    StandardDate As gSystemTime
    StandardBias As Long
    DaylightName(32) As Integer
    DaylightDate As gSystemTime
    DaylightBias As Long
End Type
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Public Function GetGMTBias() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lpTimeZoneInformation As TIME_ZONE_INFORMATION
GetTimeZoneInformation lpTimeZoneInformation
GetGMTBias = lpTimeZoneInformation.Bias
End Function

Public Function GetGMTBiasString() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim X As Long, Y As Long
X = -GetGMTBias
Y = X Mod 60
X = X \ 60
If Y < 0 Then
    Y = -Y
    GetGMTBiasString = "GMT-" & Format$(X, "00") & ":" & _
    Format$(Y, "00")
ElseIf X < 0 Then
    GetGMTBiasString = "GMT-" & _
    Format$(X, "00") & ":" & Format$(Y, "00")
Else
    GetGMTBiasString = "GMT+" & _
    Format$(X, "00") & ":" & Format$(Y, "00")
End If
End Function

Private Function CTime() As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CTime = toCTime(Now)
End Function

Private Function toCTime(d As Date) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
toCTime = DateDiff("s", CDate(#1/1/1970# - GetGMTBias / 60 / 24), d)
End Function

Private Function AscTime(CTime As Long) As Date
If lSettings.sHandleErrors = True Then On Local Error Resume Next
AscTime = CDate(#1/1/1970# - GetGMTBias / 60 / 24) + (CTime / 3600& / 24)
End Function

Public Function ReturnIRCTime(lTime As String) As String
If lSettings.sHandleErrors = True Then On Error GoTo ErrHandler
Dim t As TIME_ZONE_INFORMATION, d As Date, msg As String
d = "January 1 1970 00:00:00"
GetTimeZoneInformation t
If IsNumeric(lTime) Then
    ReturnIRCTime = Format(DateAdd("s", Val(lTime) - (t.Bias * 60), d), "ddd mmm dd yyyy hh:mm:ss")
Else
    ReturnIRCTime = (DateDiff("s", d, lTime) - (t.Bias * 60))
End If
Exit Function
ErrHandler:
    ReturnIRCTime = "Invalid date/time format"
    If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnIRCTime(lTime As String) As String"
End Function
