Attribute VB_Name = "mdlAutoPreform"
Option Explicit
Private Type gCommand
    cString As String
End Type
Private Type gAutoPerform
    cCommand(150) As gCommand
    cCount As Integer
End Type
Private lAutoPerform As gAutoPerform

Public Sub AddAutoPerform(lCommand As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(lCommand) <> 0 Then
    lAutoPerform.cCount = lAutoPerform.cCount + 1
    With lAutoPerform.cCommand(lAutoPerform.cCount)
        .cString = lCommand
    End With
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub AddAutoPerform(lCommand As String)"
End Sub

Public Sub LoadAutoPerform()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer, F As Integer
lAutoPerform.cCount = ReadINI(GetINIFile(iAutoPerform), "Settings", "Count", 0)
For i = 0 To lAutoPerform.cCount
    msg = ""
    msg = ReadINI(GetINIFile(iAutoPerform), Trim(Str(i)), "Command", "")
    If Len(msg) <> 0 Then
        If FindAutoPerformIndex(msg) = 0 Then
            F = F + 1
            lAutoPerform.cCommand(F).cString = msg
        End If
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub LoadAutoPerform()"
End Sub

Public Function ClearAutoPerform()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To 150
    lAutoPerform.cCommand(i).cString = ""
Next i
lAutoPerform.cCount = 0
If DoesFileExist(GetINIFile(iAutoPerform)) = True Then Kill GetINIFile(iAutoPerform)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ClearAutoPerform()"
End Function

Public Function FindAutoPerformIndex(lCommand As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lCommand) <> 0 Then
    For i = 0 To 150
        If LCase(lCommand) = LCase(lAutoPerform.cCommand(i).cString) Then
            FindAutoPerformIndex = i
            Exit For
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindAutoPerformIndex(lCommand As String) As Integer"
End Function

Public Sub SaveAutoPerform(lWriteToDisc As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, F As Integer
For i = 1 To 150
    If Len(lAutoPerform.cCommand(i).cString) <> 0 Then
        F = F + 1
        lAutoPerform.cCommand(F).cString = lAutoPerform.cCommand(i).cString
        If lWriteToDisc = True Then
            WriteINI GetINIFile(iAutoPerform), Trim(Str(F)), "Command", lAutoPerform.cCommand(i).cString
        End If
    End If
Next i
If F <> 0 Then
    If lWriteToDisc = True Then
        WriteINI GetINIFile(iAutoPerform), "Settings", "Count", Trim(Str(F))
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SaveAutoPerform()"
End Sub

Public Sub DeleteAutoPerform(lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(lAutoPerform.cCommand(lIndex).cString) <> 0 Then
    lAutoPerform.cCommand(lIndex).cString = ""
    SaveAutoPerform False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub DeleteAutoPerform(lIndex As Integer)"
End Sub

Public Sub FillListBoxWithAutoPerform(lListBox As ListBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
lListBox.Clear
For i = 0 To 150
    If Len(lAutoPerform.cCommand(i).cString) <> 0 Then lListBox.AddItem lAutoPerform.cCommand(i).cString
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub FillListBoxWithAutoPerform(lListbox As ListBox)"
End Sub

Public Sub FillComboWithAutoPerform(lCombo As ComboBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
lCombo.Clear
For i = 0 To 150
    If Len(lAutoPerform.cCommand(i).cString) <> 0 Then lCombo.AddItem lAutoPerform.cCommand(i).cString
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub FillComboWithAutoPerform(lCombo As ComboBox)"
End Sub

Public Sub RunAutoPerform(lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To 150
    If Len(lAutoPerform.cCommand(i).cString) <> 0 Then
        lForm.tcp.SendData lAutoPerform.cCommand(i).cString & vbCrLf
        ProcessReplaceString sRunAutoCommand, lForm.txtIncoming, lAutoPerform.cCommand(i).cString
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub RunAutoPerform(lForm As Form)"
End Sub
