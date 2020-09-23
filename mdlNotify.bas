Attribute VB_Name = "mdlNotify"
Option Explicit
Private Type gContact
    cNickname As String
End Type
Private Type gNotifyList
    nEnabled As Boolean
    nNotify(150) As gContact
    nCount As Integer
End Type
Private lNotify As gNotifyList
Private lNotifyList As String

Public Function ReturnNotifyList() As String
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
ReturnNotifyList = lNotifyList
Exit Function
ErrHandler:
    Err.Clear
    If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnNotifyList() As String"
End Function

Public Sub SetNotifyList(lData As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lNotifyList = lData
Exit Sub
ErrHandler:
    Err.Clear
    If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetNotifyList(lData As String)"
End Sub

Public Function IsUserInNotifyList(lNickName As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lNickName) <> 0 Then
    i = FindNotifyIndex(lNickName)
    If Len(lNotify.nNotify(i).cNickname) <> 0 Then IsUserInNotifyList = True
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function IsUserInNotifyList(lNickname As String) As Boolean"
End Function

Public Function ReturnNotifyCount() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, c As Integer
For i = 0 To 100
    If Len(lNotify.nNotify(i).cNickname) <> 0 Then c = c + 1
Next i
ReturnNotifyCount = c
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnNotifyCount() As Integer"
End Function

Public Function ReturnNotifyNickname(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnNotifyNickname = lNotify.nNotify(lIndex).cNickname
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnNotifyNickname(lIndex As Integer) As String"
End Function

Public Sub ClearNotify()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To 150
    lNotify.nNotify(i).cNickname = ""
Next i
lNotifyList = ""
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ClearNotify()"
End Sub

Public Function AddNotify(lNickName As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If IsUserInNotifyList(lNickName) = False Then
    i = lNotify.nCount + 1
    lNotify.nCount = i
    With lNotify.nNotify(i)
        .cNickname = lNickName
    End With
    ProcessReplaceString sAddToNotify, lSettings.sActiveServerForm.txtIncoming, lNickName
    SaveNotify
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function AddNotify(lNickname As String) As Integer"
End Function

Public Sub SetNotifyEnabled(lEnabled As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lNotify.nEnabled = lEnabled
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetNotifyEnabled(lEnabled As Boolean)"
End Sub

Public Function ReturnNotifyEnabled() As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnNotifyEnabled = lNotify.nEnabled
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnNotifyEnabled() As Boolean"
End Function

Public Sub FillListBoxWithNotify(lListBox As ListBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To lNotify.nCount
    If Len(lNotify.nNotify(i).cNickname) <> 0 Then lListBox.AddItem lNotify.nNotify(i).cNickname
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub FillListBoxWithNotify(lListBox As ListBox)"
End Sub

Public Sub SetNotifyToListBox(lListBox As ListBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, F As Integer
ClearNotify
For i = 0 To lListBox.ListCount + 1
    If Len(lListBox.List(i)) <> 0 Then
        F = F + 1
        lNotify.nNotify(F).cNickname = lListBox.List(i)
        lNotifyList = lNotifyList & lListBox.List(i) & " "
    End If
Next i
lNotify.nCount = F
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetNotify()"
End Sub

Public Function FindNotifyIndex(lNickName As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To lNotify.nCount
    If LCase(lNotify.nNotify(i).cNickname) = LCase(lNickName) Then
        FindNotifyIndex = i
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindNotifyIndex(lNickname As String) As Integer"
End Function

Public Sub DefragNotify()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg(150) As String, i As Integer, c As Integer
For i = 0 To 150
    If Len(lNotify.nNotify(i).cNickname) <> 0 Then
        msg(i) = lNotify.nNotify(i).cNickname
        c = c + 1
    End If
Next i
For i = 0 To c
    If Len(msg(i)) <> 0 Then lNotify.nNotify(i).cNickname = msg(i)
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub DefragNotify()"
End Sub

Public Sub LoadNotify()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim cnumber, msg As String, i As Integer
ClearNotify
lNotify.nCount = 0
With lNotify
    If DoesFileExist(GetINIFile(iNotify)) = True Then
        lNotify.nCount = Int(ReadINI(GetINIFile(iNotify), "Settings", "Count", 0))
        If lNotify.nCount <> 0 Then
            For i = 1 To lNotify.nCount
                lNotify.nNotify(i).cNickname = ReadINI(GetINIFile(iNotify), Str(i), "Nickname", "")
            Next i
        End If
    End If
End With
For i = 0 To lNotify.nCount
    If Len(lNotify.nNotify(i).cNickname) <> 0 Then
        lNotifyList = lNotifyList & lNotify.nNotify(i).cNickname & " "
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub LoadNotify()"
End Sub

Public Sub RemoveFromNotify(lNickName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
i = FindNotifyIndex(lNickName)
If i <> 0 Then
    With lNotify.nNotify(i)
        .cNickname = ""
    End With
    DefragNotify
    SaveNotify
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub RemoveFromNotify(lNickname As String)"
End Sub

Public Sub SetNotifyNickname(lIndex As Integer, lNickName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lNotify.nNotify(lIndex).cNickname = lNickName
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub RemoveFromNotify(lNickname As String)"
End Sub

Public Sub SaveNotify()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, j As Integer
With lNotify
    If .nCount <> 0 Then
        For i = 1 To 150
            If Len(.nNotify(i).cNickname) <> 0 Then
                j = j + 1
                WriteINI GetINIFile(iNotify), Str(j), "Nickname", .nNotify(i).cNickname
            End If
        Next i
        WriteINI GetINIFile(iNotify), "Settings", "Count", Str(j)
    End If
End With
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SaveNotify()"
End Sub
