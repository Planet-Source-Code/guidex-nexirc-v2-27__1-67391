Attribute VB_Name = "mdlAutoConnect"
Option Explicit
Private Type gServer
    sServer As String
    sPort As Long
End Type
Private Type gAutoConnect
    sServer(30) As gServer
    sCount As Integer
End Type
Private lAutoConnect As gAutoConnect

Public Sub LoadAutoConnect()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
lAutoConnect.sCount = Int(ReadINI(GetINIFile(iAutoConnect), "Settings", "Count", 0))
If lAutoConnect.sCount <> 0 Then
    For i = 1 To lAutoConnect.sCount
        lAutoConnect.sServer(i).sPort = CLng(ReadINI(GetINIFile(iAutoConnect), Trim(Str(i)), "Port", 0))
        lAutoConnect.sServer(i).sServer = ReadINI(GetINIFile(iAutoConnect), Trim(Str(i)), "Server", "")
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "If lSettings.sHandleErrors = True Then On Local Error Resume Next"
End Sub

Public Function FindAutoConnectIndex(lServer As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To lAutoConnect.sCount
    If Trim(LCase(lServer)) = Trim(LCase(lAutoConnect.sServer(i).sServer)) Then
        FindAutoConnectIndex = i
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindAutoConnectIndex(lServer As String) As Integer"
End Function

Public Sub SaveAutoConnect()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If lAutoConnect.sCount <> 0 Then
    WriteINI GetINIFile(iAutoConnect), "Settings", "Count", Trim(Str(lAutoConnect.sCount))
    For i = 1 To lAutoConnect.sCount
        WriteINI GetINIFile(iAutoConnect), Trim(Str(i)), "Server", lAutoConnect.sServer(i).sServer
        WriteINI GetINIFile(iAutoConnect), Trim(Str(i)), "Port", Str(lAutoConnect.sServer(i).sPort)
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindAutoConnectIndex(lServer As String) As Integer"
End Sub

Public Function AddToAutoConnect(lServer As String, lPort As Long) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lAutoConnect.sCount = lAutoConnect.sCount + 1
With lAutoConnect.sServer(lAutoConnect.sCount)
    .sPort = lPort
    .sServer = lServer
End With
AddToAutoConnect = lAutoConnect.sCount
SaveAutoConnect
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function AddToAutoConnect(lServer As String, lPort As Long) As Integer"
End Function

Public Sub ClearAutoConnect()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To 30
    lAutoConnect.sServer(i).sServer = ""
    lAutoConnect.sServer(i).sPort = 0
Next i
Kill GetINIFile(iAutoConnect)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ClearAutoConnect()"
End Sub

Public Sub FillComboWithAutoConnect(lCombo As ComboBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To 150
    If Len(lAutoConnect.sServer(i).sServer) <> 0 Then
        lCombo.AddItem lAutoConnect.sServer(i).sServer & " (" & lAutoConnect.sServer(i).sPort & ")"
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub FillComboWithAutoConnect(lCombo As ComboBox)"
End Sub

Public Sub FillListBoxWithAutoConnect(lListBox As ListBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To 150
    If Len(lAutoConnect.sServer(i).sServer) <> 0 Then lListBox.AddItem lAutoConnect.sServer(i).sServer & " (" & lAutoConnect.sServer(i).sPort & ")"
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub FillListBoxWithAutoConnect(lListBox As ListBox)"
End Sub

Public Sub RemoveAutoConnect(lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lAutoConnect.sServer(lIndex).sPort = 0
lAutoConnect.sServer(lIndex).sServer = ""
SaveAutoConnect
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub RemoveAutoConnect()"
End Sub

Public Sub PerformAutoConnect()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, F As Integer
For i = 1 To lAutoConnect.sCount
    If Len(lAutoConnect.sServer(i).sServer) <> 0 And lAutoConnect.sServer(i).sPort <> 0 Then
        If lSettings.sActiveServerForm.tcp.State = sckClosed Then
            ConnectToIRC lAutoConnect.sServer(i).sServer, Str(lAutoConnect.sServer(i).sPort), lSettings.sActiveServerForm
        Else
            F = NewStatusWindow(lAutoConnect.sServer(i).sServer, Str(lAutoConnect.sServer(i).sPort), True)
        End If
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub PerformAutoConnect()"
End Sub
