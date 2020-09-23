Attribute VB_Name = "TBOXModule"
Option Explicit
Declare Function QueryPerformanceCounter Lib "kernel32.dll" (lpPerformanceCount As Currency) As Long
Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Const vbSpace = " "
Private Type dispParam
    belongs As Integer
    nLines As Integer
    iLine As Byte
    TX As String
End Type
Dim SizeV As PointAPI
Public Type tBuffer
    sString As String
    mBelongs As Integer
End Type
Public Type stringTable
    Words As String
    wholeLen As Integer
End Type
Const ColorChr As String = ""
Const BoldChr As String = ""
Const PlainChr As String = ""
Const UnderlineChr As String = ""
Const ReverseChr As String = ""
Const MinAllowedWidth = 80
Dim cStartTime As Currency
Dim cPerfFreq As Currency
Dim SpW() As String

Public Function GetTextWidth2(ByVal tText As String, ByVal hw As Long) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
GetTextExtentPoint32 hw, tText, Len(tText), SizeV
GetTextWidth2 = SizeV.X
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function GetTextWidth2(ByVal tText As String, ByVal hw As Long) As Integer"
End Function

Public Function Start() As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
QueryPerformanceCounter cStartTime
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function Start() As Long"
End Function

Public Function Finish() As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim cCurrentTime As Currency
QueryPerformanceCounter cCurrentTime
Finish = 1000 * (cCurrentTime - cStartTime) / cPerfFreq
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function Finish() As Long"
End Function
