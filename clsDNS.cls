VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDNS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function WSAStartup Lib "wsock32" (ByVal VersionReq As Long, WSADataReturn As gWSAInfo) As Long
Private Declare Function WSACleanup Lib "wsock32" () As Long
Private Declare Function WSAGetLastError Lib "wsock32" () As Long
Private Declare Function GetHostByAddr Lib "wsock32" (addr As Long, addrLen As Long, addrType As Long) As Long
Private Declare Function gethostbyname Lib "wsock32" Alias "GetHostByName" (ByVal hostname As String) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private lStarted As Boolean
Private Type gWSAInfo
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To 256) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type
Private Type gHostInfo
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Private Sub Class_Initialize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lWSA As gWSAInfo
lStarted = (WSAStartup(257, lWSA) = 0)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Class_Initialize()"
End Sub

Private Sub Class_Terminate()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lStarted = True Then WSACleanup
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Class_Terminate()"
End Sub

Private Function ReturnIP(lIP As String) As Boolean
If lSettings.sHandleErrors = True Then On Error GoTo ErrHandler
Dim msg, l As Long, b As Byte
msg = Split(lIP, ".")
If UBound(msg) < 3 Then
    Exit Function
End If
For l = LBound(msg) To UBound(msg)
    b = msg(l)
Next l
ReturnIP = True
ErrHandler:
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Function ReturnIP(lIP As String) As Boolean"
End Function

Private Function CreateIP(lAddress As String) As Long
If lSettings.sHandleErrors = True Then On Error GoTo ErrHandler
Dim msg, l As Long, i As Long
msg = Split(lAddress, ".")
For i = 0 To (UBound(msg) - 1)
    l = l + (msg(i) * (256 ^ i))
Next i
If msg(UBound(msg)) < 128 Then
    l = l + (msg(UBound(msg)) * (256 ^ 3))
Else
    l = l + ((msg(UBound(msg)) - 256) * (256 ^ 3))
End If
CreateIP = l
ErrHandler:
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Function CreateIP(lAddress As String) As Long"
End Function

Public Function AddrToName(lAddress As String) As String
If lSettings.sHandleErrors = True Then On Error GoTo ErrHandler
Dim lHostInfo As gHostInfo, lHost As String * 255, msg As String, l As Long, i As Long
If ReturnIP(lAddress) Then
    i = CreateIP(lAddress)
    l = GetHostByAddr(i, 4, 4&)
    If l = 0 Then
        Exit Function
    End If
    RtlMoveMemory lHostInfo, l, Len(lHostInfo)
    RtlMoveMemory ByVal lHost, lHostInfo.hName, 255
    msg = TrimNull(lHost)
    AddrToName = msg
End If
ErrHandler:
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function AddrToName(lAddress As String) As String"
End Function

Public Function NameToAddr(ByVal lHost As String)
If lSettings.sHandleErrors = True Then On Error GoTo ErrHandler
Dim lIpArr() As Byte, lHostEntry As gHostInfo, msg As String, l As Long, i As Long, n As Integer
l = gethostbyname(lHost)
If l = 0 Then
    Exit Function
End If
RtlMoveMemory lHostEntry, l, LenB(lHostEntry)
RtlMoveMemory i, lHostEntry.hAddrList, 4
ReDim lIpArr(1 To lHostEntry.hLength)
RtlMoveMemory lIpArr(1), i, lHostEntry.hLength
For n = 1 To lHostEntry.hLength
    msg = msg & lIpArr(n) & "."
Next n
msg = Mid(msg, 1, Len(msg) - 1)
NameToAddr = msg
ErrHandler:
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function NameToAddr(ByVal lHost As String)"
End Function

Public Function AddressToName(strIP As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lStarted = True Then
    AddressToName = AddrToName(strIP)
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function AddressToName(strIP As String) As String"
End Function

Public Function NameToAddress(strName As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lStarted = True Then
    NameToAddress = NameToAddr(strName)
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function NameToAddress(strName As String) As String"
End Function

Private Function TrimNull(lTrim As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim l As Long
l = InStr(1, lTrim, Chr(0))
If l > 0 Then
    TrimNull = Left(lTrim, l - 1)
Else
    TrimNull = lTrim
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Function TrimNull(lTrim As String) As String"
End Function
