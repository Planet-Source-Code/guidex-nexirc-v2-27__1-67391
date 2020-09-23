Attribute VB_Name = "mdlAlternates"
Option Explicit
Private Type gAlternate
    aNickname As String
End Type
Private Type gAlternates
    aCount As Integer
    aAlternate(150) As gAlternate
End Type
Private lAlternates As gAlternates

Public Function SelectRandomAlternate(lWSK As Winsock) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, n As Integer, t As Long
t = lAlternates.aCount
RetHandler:
If lAlternates.aCount <> 0 Then
    n = n + 1
    If n = 150 Then Exit Function
    i = GetRnd(t)
    If Len(lAlternates.aAlternate(i).aNickname) <> 0 Then
        lSettings.sNickname = lAlternates.aAlternate(i).aNickname
        lWSK.SendData "NICK " & lAlternates.aAlternate(i).aNickname & vbCrLf
    Else
        GoTo RetHandler
    End If
End If
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Function SelectRandomAlternate(lWSK As Winsock) As Integer"
    Err.Clear
End Function

Public Sub LoadAlternates()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
lAlternates.aCount = Int(Trim(ReadINI(GetINIFile(iAlternates), "Settings", "Count", 0)))
If lAlternates.aCount <> 0 Then
    For i = 1 To lAlternates.aCount
        lAlternates.aAlternate(i).aNickname = Trim(ReadINI(GetINIFile(iAlternates), Trim(Str(i)), "Nickname", ""))
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub LoadAlternates()"
End Sub

Public Sub ClearAlternates()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To 150
    lAlternates.aAlternate(i).aNickname = ""
Next i
lAlternates.aCount = 0
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ClearAlternates()"
End Sub

Public Function AddAlternate(lNickName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lNickName) <> 0 Then
    If FindAlternateIndex(lNickName) <> 0 Then
        If lSettings.sGeneralPrompts = True Then
            MsgBox "Entry already exists!", vbInformation
        End If
    Else
        lAlternates.aCount = lAlternates.aCount + 1
        lAlternates.aAlternate(lAlternates.aCount).aNickname = lNickName
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function AddAlternate(lNickname As String, lSave As Boolean)"
End Function

Public Function FindAlternateIndex(lNickName As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lNickName) <> 0 Then
    For i = 1 To lAlternates.aCount
        If LCase(Trim(lNickName)) = LCase(Trim(lAlternates.aAlternate(i).aNickname)) Then
            FindAlternateIndex = i
            Exit For
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindAlternateIndex(lNickname As String) As Integer"
End Function

Public Sub SaveAlternates()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
WriteINI GetINIFile(iAlternates), "Settings", "Count", Trim(Str(lAlternates.aCount))
For i = 1 To 150
    If Len(lAlternates.aAlternate(i).aNickname) <> 0 Then WriteINI GetINIFile(iAlternates), Trim(Str(i)), "Nickname", lAlternates.aAlternate(i).aNickname
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SaveAlternates()"
End Sub

Public Sub FillComboWithAlternates(lCombo As ComboBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To 150
    If Len(lAlternates.aAlternate(i).aNickname) <> 0 Then lCombo.AddItem lAlternates.aAlternate(i).aNickname
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub FillComboWithAlternates(lCombo As ComboBox)"
End Sub

Public Function RemoveAlternate(lNickName As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
i = FindAlternateIndex(lNickName)
If i <> 0 Then
    lAlternates.aAlternate(i).aNickname = ""
    RemoveAlternate = True
End If
DefragAlternates
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function RemoveAlternate(lNickname As String) As Boolean"
End Function

Public Sub FillListBoxWithAlternates(lListBox As ListBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To 150
    If Len(lAlternates.aAlternate(i).aNickname) <> 0 Then lListBox.AddItem lAlternates.aAlternate(i).aNickname
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub FillListBoxWithAlternates(lListBox As ListBox)"
End Sub

Public Sub DefragAlternates()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg(150) As String, i As Integer, c As Integer
For i = 1 To lAlternates.aCount
    If Len(lAlternates.aAlternate(i).aNickname) <> 0 Then
        c = c + 1
        msg(c) = lAlternates.aAlternate(i).aNickname
    End If
Next i
For i = 0 To c
    If Len(msg(i)) <> 0 Then lAlternates.aAlternate(i).aNickname = msg(i)
Next i
lAlternates.aCount = c
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub DefragAlternates()"
End Sub

Public Function ReturnAlternateCount() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, c As Integer
For i = 0 To 150
    If Len(lAlternates.aAlternate(i).aNickname) <> 0 Then
        c = c + 1
    End If
Next i
ReturnAlternateCount = c
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnAlternateCount() As Integer"
End Function
