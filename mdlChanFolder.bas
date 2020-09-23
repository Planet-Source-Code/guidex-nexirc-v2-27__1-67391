Attribute VB_Name = "mdlChanFolder"
Option Explicit
Private Type gChanFolder
    cName As String
End Type
Private Type gChanFolders
    cCount As Integer
    cChanFolder(150) As gChanFolder
End Type
Private lChanFolders As gChanFolders

Public Function ReturnChannelFolderChannel(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnChannelFolderChannel = lChanFolders.cChanFolder(lIndex).cName
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnChannelFolderChannel(lIndex As Integer) As String"
End Function

Public Sub ClearChanFolders()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
lChanFolders.cCount = 0
For i = 0 To 150
    lChanFolders.cChanFolder(i).cName = ""
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ClearChanFolders()"
End Sub

Public Sub LoadChanFolders()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, F As Integer
lChanFolders.cCount = Int(ReadINI(GetINIFile(iChanFolder), "Settings", "Count", 0))
If lChanFolders.cCount <> 0 Then
    For i = 1 To lChanFolders.cCount
        msg = ReadINI(GetINIFile(iChanFolder), Str(i), "Name", "")
        If Len(msg) <> 0 Then
            lChanFolders.cChanFolder(i).cName = msg
            F = F + 1
        End If
    Next i
End If
If F <> 0 Then
    lChanFolders.cCount = F
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub LoadChanFolders()"
End Sub

Public Sub SaveChanFolders()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, F As Integer, msg(150) As String
WriteINI GetINIFile(iIRC), "Settings", "ShowChannelFolder", lSettings.sOptions.oShowChannelFolder
If DoesFileExist(GetINIFile(iChanFolder)) = True Then Kill GetINIFile(iChanFolder)
For i = 1 To 150
    If Len(lChanFolders.cChanFolder(i).cName) <> 0 Then
        F = F + 1
        msg(F) = lChanFolders.cChanFolder(i).cName
    End If
Next i
ClearChanFolders
If F <> 0 Then
    For i = 1 To F
        lChanFolders.cChanFolder(i).cName = msg(i)
        WriteINI GetINIFile(iChanFolder), Str(i), "Name", lChanFolders.cChanFolder(i).cName
    Next i
    lChanFolders.cCount = F
    WriteINI GetINIFile(iChanFolder), "Settings", "Count", lChanFolders.cCount
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SaveChanFolders()"
End Sub

Public Function AddtoChanFolder(lChannelName As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(lChannelName) <> 0 Then
    If DoesChannelFolderEntryExist(lChannelName) = 0 Then
        If Left(lChannelName, 1) <> "#" Then
            lChannelName = "#" & lChannelName
        End If
        lChanFolders.cCount = lChanFolders.cCount + 1
        lChanFolders.cChanFolder(lChanFolders.cCount).cName = lChannelName
        AddtoChanFolder = True
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function AddtoChanFolder(lChannelName As String) As Boolean"
End Function

Public Sub RefreshChanFolders()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ClearChanFolders
LoadChanFolders
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub RefreshChanFolders()"
End Sub

Public Sub RemoveFromChanFolder(lChannel As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lChannel) <> 0 Then
    i = FindChanFolderIndex(lChannel)
    If i <> 0 Then
        lChanFolders.cChanFolder(i).cName = ""
        WriteINI GetINIFile(iChanFolder), Str(i), "Name", ""
        RefreshChanFolders
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub RemoveFromChanFolder(lChannel As String)"
End Sub

Public Function FindChanFolderIndex(lChannel As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To lChanFolders.cCount
    If LCase(lChanFolders.cChanFolder(i).cName) = LCase(lChannel) Then
        FindChanFolderIndex = i
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindChanFolderIndex(lChannel As String) As Integer"
End Function

Public Function DoesChannelFolderEntryExist(lChannel As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To lChanFolders.cCount
    If LCase(lChanFolders.cChanFolder(i).cName) = LCase(lChannel) Then
        DoesChannelFolderEntryExist = True
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function DoesChannelFolderEntryExist(lChannel As String) As Boolean"
End Function
