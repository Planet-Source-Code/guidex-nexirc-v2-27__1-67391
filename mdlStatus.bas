Attribute VB_Name = "mdlStatus"
Option Explicit
Private Type gStatusWindow
    sForm As Form
    sServer As String
    sPort As String
End Type
Private Type gStatusWindows
    sCount As Integer
    sStatusWindow(32) As gStatusWindow
End Type
Private lStatusWindows As gStatusWindows
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Function ReturnStatusWindowTBox(lIndex As Integer) As ctlTBox
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Set ReturnStatusWindowTBox = lStatusWindows.sStatusWindow(lIndex).sForm.txtIncoming
End Function

Public Sub UnloadStatusWindow(lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload lStatusWindows.sStatusWindow(lIndex).sForm
End Sub

Public Sub ShowStatusWindow(lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim re
re = ShowWindow(lStatusWindows.sStatusWindow(frmConnectionManager.lstConnections.ListIndex + 1).sForm.hWnd, 9)
lStatusWindows.sStatusWindow(frmConnectionManager.lstConnections.ListIndex + 1).sForm.SetFocus
lStatusWindows.sStatusWindow(frmConnectionManager.lstConnections.ListIndex + 1).sForm.WindowState = vbNormal
End Sub

Public Sub DisconnectStatusWindow(lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lStatusWindows.sStatusWindow(lIndex).sForm.tcp.Close
End Sub

Public Sub SetFocusOnStatus(lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lStatusWindows.sStatusWindow(lIndex).sForm.SetFocus
End Sub

Public Sub SetStatusWindowIncomingBackColor(lIndex As Integer, lBackColor As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lStatusWindows.sStatusWindow(lIndex).sForm.txtIncoming.SetBackColor lBackColor
End Sub

Public Sub SetStatusWindowColors(lIndex As Integer, lBackColor As String, lForeColor As String)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
lStatusWindows.sStatusWindow(lIndex).sForm.txtIncoming.BackColor = lBackColor
'lStatusWindows.sStatusWindow(lIndex).sForm.txtIncoming.ForeColor = lForeColor
lStatusWindows.sStatusWindow(lIndex).sForm.txtOutgoing.BackColor = lBackColor
lStatusWindows.sStatusWindow(lIndex).sForm.txtOutgoing.ForeColor = lForeColor
lStatusWindows.sStatusWindow(lIndex).sForm.lstSent.BackColor = lBackColor
lStatusWindows.sStatusWindow(lIndex).sForm.lstSent.ForeColor = lForeColor
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetStatusWindowColors(lIndex As Integer, lBackColor As String, lForeColor As String)"
    Err.Clear
End Sub

Public Sub SetStatusWindowOutgoingForeColor(lIndex As Integer, lForeColor As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lStatusWindows.sStatusWindow(lIndex).sForm.txtIncoming.ForeColor = lForeColor
End Sub

Public Sub SetStatusWindowOutgoingBackColor(lIndex As Integer, lBackColor As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lStatusWindows.sStatusWindow(lIndex).sForm.txtIncoming.BackColor = lBackColor
End Sub

Public Function CloseConnections() As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To ReturnTCPUBound
    If Len(ReturnStatusWindowServer(i)) <> 0 Then
        If lStatusWindows.sStatusWindow(i).sForm.tcp.State = sckConnected Then
            lStatusWindows.sStatusWindow(i).sForm.tcp.Close: DoEvents
            CloseConnections = True
            Exit For
        End If
    End If
Next i
End Function

Public Function StatusWindowSendData(lIndex As Integer, lData As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
If Len(lData) <> 0 Then lStatusWindows.sStatusWindow(lIndex).sForm.tcp.SendData lData & vbCrLf
StatusWindowSendData = True
Exit Function
ErrHandler:
    Err.Clear
End Function

Public Function IsStatusTCPConnected(lIndex As Integer) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lStatusWindows.sStatusWindow(lIndex).sForm.tcp.State = sckConnected Then IsStatusTCPConnected = True
End Function

Public Function ReturnStatusWindowHwnd(lIndex As Integer) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnStatusWindowHwnd = lStatusWindows.sStatusWindow(lIndex).sForm.hWnd
End Function

Public Function ReturnStatusWindowTCPState(lIndex As Integer) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnStatusWindowTCPState = lStatusWindows.sStatusWindow(lIndex).sForm.tcp.State
End Function

Public Sub SetStatusWindowState(lIndex As Integer, lWindowState As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lStatusWindows.sStatusWindow(lIndex).sForm.WindowState = lWindowState
End Sub

Public Sub SetStatusWindowFocus(lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lStatusWindows.sStatusWindow(lIndex).sForm.SetFocus
End Sub

Public Function ReturnStatusWindowCaption(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnStatusWindowCaption = lStatusWindows.sStatusWindow(lIndex).sForm.Caption
End Function

Public Function ReturnStatusWindowCount() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnStatusWindowCount = lStatusWindows.sCount
End Function

Public Function ReturnStatusWindow(lIndex As Integer) As Form
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Set ReturnStatusWindow = lStatusWindows.sStatusWindow(lIndex).sForm
End Function

Public Function ReturnStatusWindowServer(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnStatusWindowServer = lStatusWindows.sStatusWindow(lIndex).sServer
End Function

Public Function ReturnStatusWindowPort(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnStatusWindowPort = lStatusWindows.sStatusWindow(lIndex).sPort
End Function

Public Function FindStatusWindowIndexByTag(lTag As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String
For i = 1 To lStatusWindows.sCount
    msg = Trim(LCase(lStatusWindows.sStatusWindow(i).sForm.Tag))
    If Err.Number <> 0 Then Err.Clear
    If msg = Trim(LCase(lTag)) Then
        FindStatusWindowIndexByTag = i
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindStatusWindowIndexByTag(lTag As String) As Integer"
End Function

Public Function NewStatusWindow(lServer As String, lPort As String, lConnect As Boolean) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If lPort = "" Then lPort = "6667"
If Len(lServer) <> 0 Then
    i = lStatusWindows.sCount + 1
    lStatusWindows.sCount = i
    With lStatusWindows.sStatusWindow(i)
        Set .sForm = New frmStatus
        .sForm.Show
        .sForm.Caption = "Status " & i
        .sForm.Tag = "Status " & i
        .sPort = lPort
        If lConnect = True Then ConnectToIRC lServer, lPort, .sForm
        .sServer = lServer
        NewStatusWindow = i
    End With
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function NewStatusWindow(lServer As String, lPort As String, lConnect As Boolean) As Integer"
End Function

Public Function FindStatusWindowIndexByServer(lServer As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To lStatusWindows.sCount
    If LCase(lStatusWindows.sStatusWindow(i).sServer) = LCase(lServer) Then
        FindStatusWindowIndexByServer = i
        Exit Function
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindStatusWindowIndexByServer(lServer As String) As Integer"
End Function

Public Function FindStatusWindowIndexByhWND(lhWnd As Long) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To lStatusWindows.sCount
    If LCase(lStatusWindows.sStatusWindow(i).sForm.hWnd) = LCase(lhWnd) Then
        FindStatusWindowIndexByhWND = i
        Exit Function
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindStatusWindowIndexByhWND(lHwnd As Long) As Integer"
End Function
