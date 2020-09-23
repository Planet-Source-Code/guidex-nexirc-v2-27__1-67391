Attribute VB_Name = "tempTImer"
Declare Function QueryPerformanceCounter Lib "kernel32.dll" (lpPerformanceCount As Currency) As Long
Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Dim cStartTime As Currency
Dim cPerfFreq As Currency

Public Function Start() As Long
  If QueryPerformanceFrequency(cPerfFreq) = False Then
    Debug.Print "High-perf counter not supported"
  End If
  QueryPerformanceCounter cStartTime
End Function

Public Function Finish() As Long
  Dim cCurrentTime As Currency
On Error Resume Next
  QueryPerformanceCounter cCurrentTime
  Finish = 1000 * (cCurrentTime - cStartTime) / cPerfFreq
End Function



