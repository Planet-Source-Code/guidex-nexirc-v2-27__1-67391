Attribute VB_Name = "mdlIgnore"
Option Explicit
Private Type gIgnore
    iNickname As String
End Type
Private Type gIgnores
    iEnabled As Boolean
    iCount As Integer
    iIgnore(150) As gIgnore
End Type
Private lIgnore As gIgnores

Public Function ReturnIgnoreNickname(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnIgnoreNickname = lIgnore.iIgnore(lIndex).iNickname
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnIgnoreNickname(lIndex As Integer) As String"
End Function

Public Function SetIgnoreNickname(lNickName As String, lIndex As Integer) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(lNickName) <> 0 Then
    lIgnore.iIgnore(lIndex).iNickname = lNickName
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function SetIgnoreNickname(lNickName As String, lIndex As Integer) As Boolean"
End Function

Public Sub SaveListBoxToIgnore(lListBox As ListBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, l As Integer
If lListBox.ListCount <> 0 Then
    For i = 0 To lListBox.ListCount
        If Len(Trim(lListBox.List(i))) <> 0 Then
            l = l + 1
            WriteINI GetINIFile(iIRC), "Ignore", Str(l), lListBox.List(i)
            lIgnore.iIgnore(l).iNickname = lListBox.List(i)
        End If
    Next i
    lIgnore.iCount = l
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SaveListBoxToIgnore(lListBox As ListBox)"
End Sub

Public Sub ClearIgnore()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To 150
    lIgnore.iIgnore(i).iNickname = ""
Next i
lIgnore.iCount = 0
lIgnore.iEnabled = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ClearIgnore()"
End Sub

Public Sub SetIgnoreEnabled(lEnabled As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lIgnore.iEnabled = lEnabled
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetIgnoreEnabled(lEnabled As Boolean)"
End Sub

Public Sub SetIgnoreCount(lCount As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lIgnore.iCount = lCount
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetIgnoreCount(lCount As Integer)"
End Sub

Public Sub LoadIgnore()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
SetIgnoreEnabled ReadINI(GetINIFile(iIRC), "Ignore", "Enabled", False)
SetIgnoreCount ReadINI(GetINIFile(iIRC), "Ignore", "Count", 0)
If ReturnIgnoreCount <> 0 Then
    For i = 0 To ReturnIgnoreCount
        SetIgnoreNickname ReadINI(GetINIFile(iIRC), "Ignore", Str(i), ""), i
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub LoadIgnore()"
End Sub

Public Function ReturnIgnoreEnabled() As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnIgnoreEnabled = lIgnore.iEnabled
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnIgnoreEnabled() As Boolean"
End Function

Public Function ReturnIgnoreCount() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnIgnoreCount = lIgnore.iCount
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnIgnoreCount() As Integer"
End Function

Public Function CheckIgnoreList(lNickName As String, lForm As Form, Optional lShow As Boolean) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If lIgnore.iEnabled = True Then
    If Len(lNickName) <> 0 Then
        For i = 0 To lIgnore.iCount
            If LCase(Trim(lNickName)) = LCase(Trim(lIgnore.iIgnore(i).iNickname)) Then
                If lShow = True Then ProcessReplaceString sIgnoreMessage, lForm.txtIncoming, lIgnore.iIgnore(i).iNickname
                CheckIgnoreList = True
                Exit Function
            End If
        Next i
    End If
End If
CheckIgnoreList = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function CheckIgnoreList(lNickname As String) As Boolean"
End Function

Public Sub FillListBoxWithIgnore(lListBox As ListBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To ReturnIgnoreCount
    If Len(ReturnIgnoreNickname(i)) <> 0 Then
        lListBox.AddItem ReturnIgnoreNickname(i)
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub FillListBoxWithIgnore(lListBox As ListBox)"
End Sub

Public Sub AddToIgnore(lNickName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, m As Integer
If Len(lNickName) <> 0 Then
    For m = 0 To lIgnore.iCount
        If LCase(lNickName) = LCase(lIgnore.iIgnore(m).iNickname) Then Exit Sub
    Next m
    i = lIgnore.iCount + 1
    lIgnore.iCount = i
    lIgnore.iIgnore(i).iNickname = lNickName
    lIgnore.iEnabled = True
    WriteINI GetINIFile(iIRC), "Ignore", "Enabled", "True"
    WriteINI GetINIFile(iIRC), "Ignore", "Count", Trim(Str(lIgnore.iCount))
    WriteINI GetINIFile(iIRC), "Ignore", Trim(Str(lIgnore.iCount)), lIgnore.iIgnore(i).iNickname
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub AddToIgnore(lNickname As String)"
End Sub
