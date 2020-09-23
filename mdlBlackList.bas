Attribute VB_Name = "mdlBlackList"
Option Explicit
Private Type gBlacklist
    bNickname As String
    bAddress As String
End Type
Private Type gBlacklists
    bBlacklist(150) As gBlacklist
    bBlacklistCount As Integer
    bEnabled As Boolean
End Type
Private lBlacklist As gBlacklists

Public Function IsInBlacklist(lNickName As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
i = FindBlacklistIndex(lNickName)
If i <> 0 Then IsInBlacklist = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function IsInBlacklist(lNickname)"
End Function

Public Function RemoveFromBlacklist(lNickName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
i = FindBlacklistIndex(lNickName)
If i <> 0 Then
    lBlacklist.bBlacklist(i).bAddress = ""
    lBlacklist.bBlacklist(i).bNickname = ""
    WriteINI GetINIFile(iBlacklist), Trim(Str(i)), vbNullString, vbNullString
    If lSettings.sCustomizeVisible = True Then frmCustomize.lstBlacklist.RemoveItem FindListBoxIndex(lNickName, frmCustomize.lstBlacklist)
Else
    Exit Function
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function RemoveFromBlacklist(lNickname As String)"
End Function

Public Sub ClearBlacklist()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To 150
    lBlacklist.bBlacklist(i).bAddress = ""
    lBlacklist.bBlacklist(i).bNickname = ""
Next i
lBlacklist.bBlacklistCount = 0
If lSettings.sCustomizeVisible = True Then frmCustomize.lstBlacklist.Clear
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ClearBlacklist()"
End Sub

Public Function FindBlacklistIndex(lNickName As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lNickName) <> 0 And lBlacklist.bBlacklistCount <> 0 Then
    For i = 1 To lBlacklist.bBlacklistCount
        If LCase(lBlacklist.bBlacklist(i).bNickname) = LCase(lNickName) Then
            FindBlacklistIndex = i
            Exit For
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindBlacklistIndex(lNickname As String) As Integer"
End Function

Public Sub AddToBlacklist(lNickName As String, lAddress As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(lNickName) <> 0 Then
    lBlacklist.bBlacklistCount = lBlacklist.bBlacklistCount + 1
    With lBlacklist.bBlacklist(lBlacklist.bBlacklistCount)
        .bNickname = lNickName
        .bAddress = lAddress
    End With
    If lSettings.sCustomizeVisible = True Then frmCustomize.lstBlacklist.AddItem lNickName
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub AddToBlacklist(lNickname As String)"
End Sub

Public Sub SaveBlacklist()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
WriteINI GetINIFile(iBlacklist), "Settings", "Count", Trim(Str(lBlacklist.bBlacklistCount))
For i = 1 To lBlacklist.bBlacklistCount
    WriteINI GetINIFile(iBlacklist), Trim(Str(i)), "Nickname", lBlacklist.bBlacklist(i).bNickname
    WriteINI GetINIFile(iBlacklist), Trim(Str(i)), "Address", lBlacklist.bBlacklist(i).bAddress
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SaveBlacklist()"
End Sub

Public Sub FillListBoxWithBlacklist(lListBox As ListBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
lListBox.Clear
For i = 0 To 150
    If Len(lBlacklist.bBlacklist(i).bNickname) <> 0 Then
        lListBox.AddItem lBlacklist.bBlacklist(i).bNickname
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub FillListBoxWithBlacklist(lListBox As ListBox)"
End Sub

Public Sub LoadBlacklist()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, c As Integer, msg() As String, msg2() As String
lBlacklist.bBlacklistCount = ReadINI(GetINIFile(iBlacklist), "Settings", "Count", 0)
If lBlacklist.bBlacklistCount <> 0 Then
    For i = 1 To lBlacklist.bBlacklistCount
        msg(i) = ReadINI(GetINIFile(iBlacklist), Trim(Str(i)), "Nickname", "")
        msg2(i) = lBlacklist.bBlacklist(i).bAddress = ReadINI(GetINIFile(iBlacklist), Trim(Str(i)), "Address", "")
        If Len(msg(i)) <> 0 And Len(msg2(i)) <> 0 Then
            c = c + 1
            lBlacklist.bBlacklist(c).bNickname = msg(i)
            lBlacklist.bBlacklist(c).bAddress = msg2(i)
        End If
    Next i
    lBlacklist.bBlacklistCount = c
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub LoadBlacklist()"
End Sub
