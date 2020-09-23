Attribute VB_Name = "mdlDNS"
Option Explicit
Const WSADescription_Len = 256
Const WSASYS_Status_Len = 128
Private Type gHostInfo
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type
Private Type WSAdata
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type
Private Declare Function WSAStartup Lib "wsock32" (ByVal VersionReq As Long, WSADataReturn As WSAdata) As Long
Private Declare Function WSACleanup Lib "wsock32" () As Long
Private Declare Function WSAGetLastError Lib "wsock32" () As Long
Private Declare Function GetHostByAddr Lib "wsock32" (addr As Long, addrLen As Long, addrType As Long) As Long
Private Declare Function gethostbyname Lib "wsock32" (ByVal hostname As String) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Private Function IsIP(ByVal lIpAddress As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim t As String, s As String, i As Integer
s = lIpAddress
While InStr(s, ".") <> 0
    t = Left(s, InStr(s, ".") - 1)
    If IsNumeric(t) And Val(t) >= 0 And Val(t) <= 255 Then s = Mid(s, InStr(s, ".") + 1) _
        Else Exit Function
    i = i + 1
Wend
t = s
If IsNumeric(t) And InStr(t, ".") = 0 And Len(t) = Len(Trim(Str(Val(t)))) And Val(t) >= 0 And Val(t) <= 255 And lIpAddress <> "255.255.255.255" And i = 3 Then IsIP = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Function IsIP(ByVal lIpAddress As String) As Boolean"
End Function

Private Function CreateIP(lAddress As String) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As Long
msg = Left(lAddress, InStr(lAddress, ".") - 1)
lAddress = Mid(lAddress, InStr(lAddress, ".") + 1)
msg = msg + Left(lAddress, InStr(lAddress, ".") - 1) * 256
lAddress = Mid(lAddress, InStr(lAddress, ".") + 1)
msg = msg + Left(lAddress, InStr(lAddress, ".") - 1) * 256 * 256
lAddress = Mid(lAddress, InStr(lAddress, ".") + 1)
If lAddress < 128 Then
    msg = msg + lAddress * 256 * 256 * 256
Else
    msg = msg + (lAddress - 256) * 256 * 256 * 256
End If
CreateIP = msg
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Function CreateIP(lAddress As String) As Long"
End Function

Private Function NameByAddr(lAddress As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim l As Long, lIP As Long, lHost As String * 255, msg As String, lHostInfo As gHostInfo
If IsIP(lAddress) Then
    lIP = CreateIP(lAddress)
    l = GetHostByAddr(lIP, 4, 2)
    If l <> 0 Then
        RtlMoveMemory lHostInfo, l, Len(lHostInfo)
        RtlMoveMemory ByVal lHost, lHostInfo.hName, 255
        msg = lHost
        If InStr(msg, Chr(10)) <> 0 Then msg = Left(msg, InStr(msg, Chr(0)) - 1)
        msg = Trim(msg)
        NameByAddr = msg
    Else
        Exit Function
    End If
Else
    Exit Function
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Function NameByAddr(lAddress As String) As String"
End Function

Private Function AddrByName(ByVal lHost As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim hostent_addr As Long, hst As gHostInfo, hostip_addr As Long, temp_ip_address() As Byte, i As Integer, ip_address As String
If IsIP(lHost) Then
    AddrByName = lHost
    Exit Function
End If
hostent_addr = gethostbyname(lHost)
If hostent_addr = 0 Then
    Exit Function
End If
RtlMoveMemory hst, hostent_addr, LenB(hst)
RtlMoveMemory hostip_addr, hst.hAddrList, 4
ReDim temp_ip_address(1 To hst.hLength)
RtlMoveMemory temp_ip_address(1), hostip_addr, hst.hLength
For i = 1 To hst.hLength
    ip_address = ip_address & temp_ip_address(i) & "."
Next
ip_address = Mid(ip_address, 1, Len(ip_address) - 1)
AddrByName = ip_address
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Function AddrByName(ByVal lHost As String)"
End Function

Public Function AddressToName(strIP As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
AddressToName = StripTerminator(NameByAddr(strIP))
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function AddressToName(strIP As String)"
End Function

Public Function NameToAddress(strName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
NameToAddress = StripTerminator(AddrByName(strName))
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function NameToAddress(strName As String)"
End Function
