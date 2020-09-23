Attribute VB_Name = "mdlLongIP"
Option Explicit
Declare Function ntohl Lib "wsock32.dll" (ByVal netlong As Long) As Long
Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Src As Any, ByVal cb&)
Declare Function inet_addr Lib "wsock32.dll" (ByVal CP As String) As Long
Declare Function htonl Lib "wsock32.dll" (ByVal hostlong As Long) As Long

Function IrcGetIP(ByVal IPL$) As String
If lSettings.sHandleErrors = True Then On Error GoTo IrcGetAscIPError
Dim lpStr&
Dim nStr&
Dim retString$
Dim inn&
If Val(IPL) > 2147483647 Then
    inn = Val(IPL) - 4294967296#
Else
    inn = Val(IPL)
End If
inn = ntohl(inn)
retString = String(32, 0)
lpStr = inet_ntoa(inn)
If lpStr = 0 Then
    IrcGetIP = "0.0.0.0"
    Exit Function
End If
nStr = lstrlen(lpStr)
If nStr > 32 Then nStr = 32
MemCopy ByVal retString, ByVal lpStr, nStr
retString = Left(retString, nStr)
IrcGetIP = retString
Exit Function
IrcGetAscIPError:
    IrcGetIP = "0.0.0.0"
    Exit Function
    Resume
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Function IrcGetIP(ByVal IPL$) As String"
End Function

Function IrcGetLongIP(ByVal AscIp$) As String
If lSettings.sHandleErrors = True Then On Error GoTo IrcGetLongIpError
Dim inn&
inn = inet_addr(AscIp)
inn = htonl(inn)
If inn < 0 Then
    IrcGetLongIP = CVar(inn + 4294967296#)
    Exit Function
Else
    IrcGetLongIP = CVar(inn)
    Exit Function
End If
Exit Function
IrcGetLongIpError:
    IrcGetLongIP = "0"
    Exit Function
    Resume
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Function IrcGetLongIP(ByVal AscIp$) As String"
End Function
