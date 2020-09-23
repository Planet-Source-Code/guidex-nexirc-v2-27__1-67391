VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIRCServer_Channel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 1
Public Name As String
Public Topic As String
Public Modes As New Collection
Public TopicSetOn As Long
Public TopicSetBy As String
Public Index As Long
Public Password As String
Public UserLevels As New Collection
Public Invited As New Collection
Public Key As String
Public Limit As Integer
Public All As New Collection
Public Ops As New Collection
Public Voices As New Collection
Public NormUsers As New Collection
Public Bans As New Collection
Public Exceptions As New Collection
Public Invites As New Collection
 
Public Property Get lUserCount() As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lUserCount = NormUsers.Count + Ops.Count + Voices.Count
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Property Get lUserCount() As Long"
End Property

Public Property Get GetAll() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To All.Count
    GetAll = GetAll & " " & All.Item(i)
Next i
GetAll = Trim(GetAll)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Property Get GetAll() As String"
End Property

Public Property Get GetOps() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Ops.Count
    GetOps = GetOps & " @" & Ops.Item(i)
Next i
GetOps = Trim(GetOps)
If GetOps <> "" Then GetOps = GetOps & " "
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Property Get GetOps() As String"
End Property

Public Property Get GetVoices() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Voices.Count
    GetVoices = GetVoices & " +" & Voices.Item(i)
Next i
GetVoices = Trim(GetVoices)
If GetVoices <> "" Then GetVoices = GetVoices & " "
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Property Get GetVoices() As String"
End Property

Public Property Get GetNorms() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To NormUsers.Count
    GetNorms = GetNorms & " " & NormUsers.Item(i)
Next i
GetNorms = Trim(GetNorms)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Property Get GetNorms() As String"
End Property

Public Function IsOp(Nick As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Ops.Count
    If Ops(i) = Nick Then
        IsOp = True
        Exit Function
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function IsOp(Nick As String) As Boolean"
End Function

Public Function IsVoice(Nick As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Voices.Count
    If Voices(i) = Nick Then
        IsVoice = True
        Exit Function
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function IsVoice(Nick As String) As Boolean"
End Function

Public Function IsNorm(Nick As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To NormUsers.Count
    If NormUsers(i) = Nick Then
        IsNorm = True
        Exit Function
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function IsNorm(Nick As String) As Boolean"
End Function

Public Function IsMode(Mode As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Modes.Count
    If Modes(i) = Mode Then
        IsMode = True
        Exit Function
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function IsMode(Mode As String) As Boolean"
End Function

Public Function IsOnChan(NickName As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To All.Count
    If All(i) = NickName Then
        IsOnChan = True
        Exit Function
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function IsOnChan(NickName As String) As Boolean"
End Function

Public Function IsBanned(User As clsIRCServer_User) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Bans.Count
    If (User.Nick & "!" & User.ident & "@" & User.DNS) Like Bans(i) Then
        IsBanned = True
        Exit Function
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function IsBanned(User As clsIRCServer_User) As Boolean"
End Function

Public Function IsBanned2(ID As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Bans.Count
    If (ID Like Bans(i)) Then
        IsBanned2 = True
        Exit Function
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function IsBanned2(ID As String) As Boolean"
End Function

Public Property Get GetModesForFile()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Modes.Count
    GetModesForFile = GetModesForFile + Modes(i)
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Property Get GetModesForFile()"
End Property

Public Property Get GetModes() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Modes.Count
    GetModes = GetModes + Modes(i)
Next i
If Key <> "" And Limit <> 0 Then
    GetModes = GetModes & "lk " & Limit & " " & Key
ElseIf Key <> "" And Limit = 0 Then
    GetModes = GetModes & "k " & Key
ElseIf Key = "" And Limit <> 0 Then
    GetModes = GetModes & "l " & Limit
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Property Get GetModes() As String"
End Property

Public Sub AddToUserList(NickName As String, level As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UserLevels.Add NickName & " " & level, NickName
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub AddToUserList(NickName As String, level As Long)"
End Sub

Public Sub RemoveFromUserList(NickName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UserLevels.Remove NickName
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub RemoveFromUserList(NickName As String)"
End Sub

Public Sub GetUserListItem(Index, ByRef NickName As String, ByRef level As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
NickName = UserLevels(Index)
level = CLng(Mid(NickName, InStr(1, NickName, " ")))
NickName = Replace(NickName, " " & level, "")
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub GetUserListItem(Index, ByRef NickName As String, ByRef level As Long)"
End Sub

Public Sub AddModes(NewModes As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Len(NewModes)
    Modes.Add Mid(NewModes, i, 1), Mid(NewModes, i, 1)
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub AddModes(NewModes As String)"
End Sub

Public Function ULOp(Nick As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim User As clsIRCServer_User, level As Long
Set User = NickToObject(Nick)
Dim i As Long
For i = 1 To UserLevels.Count
    GetUserListItem i, Nick, level
    If (Nick = User.Nick Or Nick = User.IdentifiedAs) And (level >= 100 And User.Identified) Then
        ULOp = True
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ULOp(Nick As String) As Boolean"
End Function

Public Function ULVoice(Nick As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim User As clsIRCServer_User, level As Long
Set User = NickToObject(Nick)
Dim i As Long
For i = 1 To UserLevels.Count
    GetUserListItem i, Nick, level
    If (Nick = User.Nick Or Nick = User.IdentifiedAs) And (level = 50 And User.Identified) Then
        ULVoice = True
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ULVoice(Nick As String) As Boolean"
End Function

Public Function IsInvited(Nick As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Invited.Count
    If Nick = Invited(i) Then
        IsInvited = True
        Invited.Remove i
        Exit Function
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function IsInvited(Nick As String) As Boolean"
End Function

Public Function IsInvited2(User As clsIRCServer_User) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Invites.Count
    If (User.Nick & "!" & User.ident & "@" & User.DNS) Like Invites(i) Then
        IsInvited2 = True
        Exit Function
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function IsInvited2(User As clsIRCServer_User) As Boolean"
End Function

Public Function IsException(User As clsIRCServer_User) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Exceptions.Count
    If (User.Nick & "!" & User.ident & "@" & User.DNS) Like Exceptions(i) Then
        IsException = True
        Exit Function
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function IsException(User As clsIRCServer_User) As Boolean"
End Function

Public Function IsException2(ID As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Exceptions.Count
    If (ID Like Exceptions(i)) Then
        IsException2 = True
        Exit Function
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function IsException2(ID As String) As Boolean"
End Function

Public Function IsInvited3(ID As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
For i = 1 To Invites.Count
    If (ID Like Invites(i)) Then
        IsInvited3 = True
        Exit Function
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function IsInvited3(ID As String) As Boolean"
End Function

Private Sub Class_Initialize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lChanCount = lChanCount + 1
lChannelsUBound = lChannelsUBound + 1
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Class_Initialize()"
End Sub

Private Sub Class_Terminate()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lChanCount = lChanCount - 1
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Class_Terminate()"
End Sub
