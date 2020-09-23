VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIRCServer_User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Memos"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 1
Dim m As New clsIRCServer_User
Public Name As String, Nick As String, ID As String, Email As String
Public Onchannels As New Collection
Public Idle As Long, Index As Long
Public NewUser As Boolean, Identified As Boolean, IRCOp As Boolean
Public ident As String, SignOn As Long, DNS As String
Public NR As Boolean
Public OwnerOf As New Collection
Public SentQuit As Boolean
Public Modes As New Collection
Public Away As Boolean
Public AwayMsg As String
Public MsgsSent As Long
Public HasRegistered As Boolean
Public Killed As Boolean
Public IdentifiedAs As String
Public FloodProt As Integer
Public Server As String
Public ServerDescritption As String
Public LocalUser As Boolean
Public Hops As Integer
Public GID As String
Public Route As String
Public RealDNS As String
Public Ponged As Boolean

Public Function IsOnLink(sName As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If InStr(1, Route, sName) <> 0 Then IsOnLink = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function IsOnLink(sName As String) As Boolean"
End Function

Public Function GetOnChans() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long, Chan As clsIRCServer_Channel, NewChan As String
For i = 1 To Onchannels.Count
    Set Chan = ChanToObject(Onchannels(i))
    If Chan.IsOp(Nick) Then
        NewChan = "@" & Onchannels(i)
    ElseIf Chan.IsVoice(Nick) Then
        NewChan = "+" & Onchannels(i)
    Else
        NewChan = Onchannels(i)
    End If
    GetOnChans = GetOnChans & " " & NewChan
Next i
GetOnChans = Trim(GetOnChans)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function GetOnChans() As String"
End Function

Public Function GetOnChans2() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Onchannels.Count
    GetOnChans2 = GetOnChans2 & " " & Onchannels(i)
Next i
GetOnChans2 = Trim(GetOnChans2)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function GetOnChans() As String"
End Function

Public Function GetWhois(NickName As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
GetWhois = ":" & ServerName & " 311 " & NickName & " " & Nick & " " & ident & " " & DNS & " * :" & Name & vbCrLf
GetWhois = GetWhois & ":" & ServerName & " 307 " & NickName & " " & Nick & " DNS " & DNS & vbCrLf
If Modes.Count <> 0 Then GetWhois = GetWhois & ":" & ServerName & " 309 " & NickName & " " & Nick & " :is using modes: +" & GetModes & vbCrLf
If Onchannels.Count <> 0 Then GetWhois = GetWhois & ":" & ServerName & " 319 " & NickName & " " & Nick & " :" & GetOnChans & vbCrLf
If Away Then GetWhois = GetWhois & ":" & ServerName & " 301 " & NickName & " " & Nick & " :" & AwayMsg & vbCrLf
GetWhois = GetWhois & ":" & ServerName & " 312 " & NickName & " " & Nick & " " & Server & " :" & ServerDescritption & vbCrLf
If (IRCOp Or IsMode("o")) Then
    GetWhois = GetWhois & ":" & ServerName & " 313 " & NickName & " " & Nick & " :is an IRC Operator - Server Administrator" & vbCrLf
    GetWhois = GetWhois & ":" & ServerName & " 310 " & NickName & " " & Nick & " :is a network helper" & vbCrLf
End If
If Not Idle = 0 Then GetWhois = GetWhois & ":" & ServerName & " 317 " & NickName & " " & Nick & " " & (UnixTime - Idle) & " " & SignOn & " :seconds idle, signon time" & vbCrLf
GetWhois = GetWhois & ":" & ServerName & " 318 " & NickName & " " & Nick & " :End of /WHOIS list." & vbCrLf
Debug.Print GetWhois
End Function

Public Function IsOwner(Channel As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To OwnerOf.Count
    If OwnerOf(i) = Channel Then
        IsOwner = True
        Exit Function
    End If
Next i
End Function

Public Sub ClearOwnerShip()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To OwnerOf.Count
    OwnerOf.Remove i
Next i
End Sub

Public Property Get GetModes() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Modes.Count
    GetModes = GetModes + Modes(i)
Next i
End Property

Public Sub AddModes(NewModes As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Len(NewModes)
    Modes.Add Mid(NewModes, i, 1), Mid(NewModes, i, 1)
Next i
End Sub

Public Function IsMode(Mode As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Modes.Count
    If Modes(i) = Mode Then
        IsMode = True
        Exit Function
    End If
Next i
End Function

Public Function GetChanModes(lChanName As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim CurChan As clsIRCServer_Channel
Set CurChan = ChanToObject(lChanName)
If Not CurChan.IsOnChan(Nick) Then Exit Function
If CurChan.IsVoice(Nick) Then GetChanModes = GetChanModes & "+"
If CurChan.IsOp(Nick) Then GetChanModes = GetChanModes & "@"
End Function

Public Function GetMask() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
GetMask = Mid(DNS, 1, InStr(1, DNS, ".") - 1)
GetMask = Replace(DNS, GetMask, "*")
End Function

Public Function IsOnChan(Chan As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Onchannels.Count
    If UCase(Onchannels(i)) = UCase(Chan) Then
        IsOnChan = True
        Exit Function
    End If
Next i
End Function

Private Sub Class_Initialize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lUserCount = lUserCount + 1
CurGlobalUsers = CurGlobalUsers + 1
If (CurGlobalUsers > MaxGlobalUsers) Then MaxGlobalUsers = MaxGlobalUsers + 1
End Sub

Private Sub Class_Terminate()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If IRCOp Then Operators = Operators - 1
If LocalUser Then lUserCount = lUserCount - 1
CurGlobalUsers = CurGlobalUsers - 1
End Sub
