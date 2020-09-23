Attribute VB_Name = "mdlAutoJoin"
Option Explicit
Private Type gAutoJoin
    sNetwork As String
    aChannelName As String
End Type
Private Type gAutoJoins
    aAutoJoin(150) As gAutoJoin
    aCount As Integer
End Type
Private lAutoJoin As gAutoJoins

Public Function FindAutoJoinIndex(lChannel As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lChannel) <> 0 And lAutoJoin.aCount <> 0 Then
    For i = 0 To lAutoJoin.aCount
        If LCase(lAutoJoin.aAutoJoin(i).aChannelName) = LCase(lChannel) Then
            FindAutoJoinIndex = i
            Exit Function
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindAutoJoinIndex(lChannel As String) As Integer"
End Function

Public Function FindEmptyAutoJoin() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To lAutoJoin.aCount
    If Len(lAutoJoin.aAutoJoin(i).aChannelName) = 0 And Len(lAutoJoin.aAutoJoin(i).sNetwork) = 0 Then
        FindEmptyAutoJoin = i
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindEmptyAutoJoin() As Integer"
End Function

Public Sub ActivateAutoJoin(lPrompt As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mbox As VbMsgBoxResult, i As Integer, s As Boolean
If lSettings.sAutoJoinEnabled = False Then Exit Sub
If lPrompt = True Then
    If lSettings.sGeneralPrompts = True Then
        mbox = MsgBox("You are about to auto join all channels in your auto join list, proceed?", vbYesNoCancel + vbQuestion)
    Else
        mbox = vbYes
    End If
    If mbox = vbYes Then
        s = True
    ElseIf mbox = vbNo Then
        s = False
    ElseIf mbox = vbCancel Then
        s = False
        Exit Sub
    End If
Else
    s = True
End If
If s = True Then
    For i = 0 To lAutoJoin.aCount
        If LCase(lAutoJoin.aAutoJoin(i).sNetwork) = LCase(lSettings.sNetwork) Then
            ProcessReplaceString sAutoJoin, lSettings.sActiveServerForm.txtIncoming, lAutoJoin.aAutoJoin(i).aChannelName
            lSettings.sActiveServerForm.tcp.SendData "JOIN " & lAutoJoin.aAutoJoin(i).aChannelName & vbCrLf
            If DoesChannelFolderEntryExist(lAutoJoin.aAutoJoin(i).aChannelName) = False Then
                AddtoChanFolder lAutoJoin.aAutoJoin(i).aChannelName
                SaveChanFolders
            End If
            DoEvents
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ActivateAutoJoin(lPrompt As Boolean)"
End Sub

Public Function AddAutoJoin(lChannelName As String, lNetwork As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lChannelName) <> 0 And Len(lNetwork) <> 0 Then
    For i = 0 To lAutoJoin.aCount
        If LCase(lAutoJoin.aAutoJoin(i).aChannelName) = LCase(lChannelName) And LCase(lNetwork) = LCase(lAutoJoin.aAutoJoin(i).sNetwork) Then
            If lSettings.sGeneralPrompts = True Then
                MsgBox "Unable to add this channel to your autojoin list, entry already exists!", vbExclamation
                Exit Function
            Else
                Exit Function
            End If
        End If
    Next i
    i = FindEmptyAutoJoin
    If i = 0 Then
        i = lAutoJoin.aCount + 1
        lAutoJoin.aCount = i
    End If
    lAutoJoin.aAutoJoin(i).aChannelName = lChannelName
    lAutoJoin.aAutoJoin(i).sNetwork = lNetwork
    WriteINI GetINIFile(iAutoJoin), Str(i), "Channel", lChannelName
    WriteINI GetINIFile(iAutoJoin), Str(i), "Network", lNetwork
    WriteINI GetINIFile(iAutoJoin), "Settings", "Count", Str(i)
    AddAutoJoin = i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function AddAutoJoin(lChannelName As String, lNetwork As String) As Integer"
End Function

Public Function ReturnAutoJoinCount() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, d As Integer
For i = 0 To 150
    If Len(lAutoJoin.aAutoJoin(i).aChannelName) <> 0 And Len(lAutoJoin.aAutoJoin(i).sNetwork) <> 0 Then
        d = d + 1
    End If
Next i
ReturnAutoJoinCount = d
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnAutoJoinCount() As Integer"
End Function

Public Sub FillListBoxWithAutoJoin(lListBox As ListBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
lListBox.Clear
For i = 0 To 150
    If Len(lAutoJoin.aAutoJoin(i).aChannelName) <> 0 Then lListBox.AddItem lAutoJoin.aAutoJoin(i).aChannelName
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub FillListBoxWithAutoJoin(lListBox As ListBox)"
End Sub

Public Function ReturnAutoJoinNetwork(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnAutoJoinNetwork = lAutoJoin.aAutoJoin(lIndex).sNetwork
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnAutoJoinNetwork(lIndex As Integer) As String"
End Function

Public Function ReturnAutoJoinChannel(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnAutoJoinChannel = lAutoJoin.aAutoJoin(lIndex).aChannelName
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnAutoJoinChannel(lIndex As Integer) As String"
End Function

Public Function CheckAutoJoin(lIndex As Integer) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(lAutoJoin.aAutoJoin(lIndex).aChannelName) <> 0 And Len(lAutoJoin.aAutoJoin(lIndex).sNetwork) <> 0 Then CheckAutoJoin = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function CheckAutoJoin(lIndex As Integer) As Boolean"
End Function

Public Sub DeleteAutoJoin(lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
WriteINI GetINIFile(iAutoJoin), Str(lIndex), "Channel", ""
WriteINI GetINIFile(iAutoJoin), Str(lIndex), "Network", ""
lAutoJoin.aAutoJoin(lIndex).aChannelName = ""
lAutoJoin.aAutoJoin(lIndex).sNetwork = ""
If lIndex = lAutoJoin.aCount Then
    lIndex = lIndex - 1
    WriteINI GetINIFile(iAutoJoin), "Settings", "Count", Str(lIndex)
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub DeleteAutoJoin(lIndex As Integer)"
End Sub

Public Sub LoadAutoJoin()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
lAutoJoin.aCount = ReadINI(GetINIFile(iAutoJoin), "Settings", "Count", 0)
For i = 0 To lAutoJoin.aCount
    lAutoJoin.aAutoJoin(i).aChannelName = ReadINI(GetINIFile(iAutoJoin), Str(i), "Channel", "")
    If Len(lAutoJoin.aAutoJoin(i).aChannelName) <> 0 Then
        lAutoJoin.aAutoJoin(i).sNetwork = ReadINI(GetINIFile(iAutoJoin), Str(i), "Network", "")
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub LoadAutoJoin()"
End Sub
