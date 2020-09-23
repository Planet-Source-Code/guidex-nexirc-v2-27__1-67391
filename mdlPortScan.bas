Attribute VB_Name = "mdlPortScan"
Option Explicit
Private Type gPortScan
    pServerIndex As Integer
    pPortStart As Long
    pPortEnd As Long
    pCurrentNetwork As Integer
End Type
Private Type gPortScans
    pPortScan(150) As gPortScan
    pCount As Integer
    pInProgress As Boolean
    pIndex As Integer
End Type
Private lPortScans As gPortScans

Public Function CheckPortScanTimerProc(lListView As ListView) As Boolean
Dim i As Integer
If ReturnPortScanInProgress = False Then
    If ReturnPortScanIndex <> lListView.ListItems.Count Then
        SetPortScanIndex ReturnPortScanIndex + 1
        lListView.ListItems(ReturnPortScanIndex).SubItems(2) = ""
        NewPortScan lListView.ListItems(ReturnPortScanIndex).SubItems(1), 6660, 6670
        NewPortScan lListView.ListItems(ReturnPortScanIndex).SubItems(1), 7000, 7000
        CheckPortScanTimerProc = True
    End If
End If
End Function

Public Sub SetPortScanIndex(lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lPortScans.pIndex = lIndex
End Sub

Public Sub SetPortScanInProgress(lValue As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lPortScans.pInProgress = lValue
End Sub

Public Sub ClearPortScanIndex()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lPortScans.pIndex = 0
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ClearPortScanIndex()"
End Sub

Public Function ReturnPortScanIndex() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnPortScanIndex = lPortScans.pIndex
End Function

Public Function ReturnPortScanCount() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnPortScanCount = lPortScans.pCount
End Function

Public Function ReturnPortScanInProgress() As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnPortScanInProgress = lPortScans.pInProgress
End Function

Public Sub LoadPortScanRange()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sTestConnectionsLoaded = False Then
    Load frmCustomize.wskTestConnection(6660)
    Load frmCustomize.wskTestConnection(6661)
    Load frmCustomize.wskTestConnection(6662)
    Load frmCustomize.wskTestConnection(6663)
    Load frmCustomize.wskTestConnection(6664)
    Load frmCustomize.wskTestConnection(6665)
    Load frmCustomize.wskTestConnection(6666)
    Load frmCustomize.wskTestConnection(6667)
    Load frmCustomize.wskTestConnection(6668)
    Load frmCustomize.wskTestConnection(6669)
    Load frmCustomize.wskTestConnection(6670)
    Load frmCustomize.wskTestConnection(7000)
End If
End Sub

Public Sub ClearPortScan(lCloseTestConnections As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
ClearPortScanIndex
frmCustomize.tmrPortScan.Enabled = False
frmCustomize.tmrPortScanTimeout.Enabled = False
For i = 0 To 150
    With lPortScans.pPortScan(i)
        .pCurrentNetwork = 0
        .pPortEnd = 0
        .pPortStart = 0
        .pServerIndex = 0
    End With
Next i
lPortScans.pCount = 0
lPortScans.pInProgress = False
If lCloseTestConnections = True Then
    If lSettings.sTestConnectionsLoaded = True Then
        Unload frmCustomize.wskTestConnection(6660)
        Unload frmCustomize.wskTestConnection(6661)
        Unload frmCustomize.wskTestConnection(6662)
        Unload frmCustomize.wskTestConnection(6663)
        Unload frmCustomize.wskTestConnection(6664)
        Unload frmCustomize.wskTestConnection(6665)
        Unload frmCustomize.wskTestConnection(6666)
        Unload frmCustomize.wskTestConnection(6667)
        Unload frmCustomize.wskTestConnection(6668)
        Unload frmCustomize.wskTestConnection(6669)
        Unload frmCustomize.wskTestConnection(6670)
        Unload frmCustomize.wskTestConnection(7000)
    End If
    lSettings.sTestConnectionsLoaded = False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ClearPortScan(lCloseTestConnections As Boolean)"
End Sub

Public Sub RecievePortScanResults(lServer As String, lOpenPort As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, F As Integer, c As Integer, msg As String, G As Long, lLow As Long, lHigh As Long, lSPLT() As String, m As Integer
i = FindPortScanIndex(lServer)
If lSettings.sCustomizeVisible = True Then
    If Len(lServer) <> 0 And lOpenPort <> 0 Then
        For F = 1 To frmCustomize.lvwServers.ListItems.Count
            If LCase(frmCustomize.lvwServers.ListItems(F).SubItems(1)) = LCase(lServer) Then
                If Len(frmCustomize.lvwServers.ListItems(F).SubItems(2)) <> 0 Then
                    msg = frmCustomize.lvwServers.ListItems(F).SubItems(2) & "," & Trim(Str(lOpenPort))
                Else
                    msg = Trim(Str(lOpenPort))
                End If
                frmCustomize.lvwServers.ListItems(F).SubItems(2) = msg
            End If
        Next F
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function NewPortScan(lStartPort As Long, lEndPort As Long) As Integer"
End Sub

Public Function NewPortScan(lServer As String, lStartPort As Long, lEndPort As Long) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
If lStartPort <> 0 And lEndPort <> 0 Then
    lPortScans.pInProgress = True
    lPortScans.pCount = lPortScans.pCount + 1
    lPortScans.pPortScan(lPortScans.pCount).pCurrentNetwork = FindNetworkIndex(frmCustomize.cmbNetwork)
    lPortScans.pPortScan(lPortScans.pCount).pServerIndex = FindServerIndex(lServer)
    lPortScans.pPortScan(lPortScans.pCount).pPortStart = CLng(lStartPort)
    lPortScans.pPortScan(lPortScans.pCount).pPortEnd = CLng(lStartPort)
    frmCustomize.tmrPortScanTimeout.Enabled = True
    For i = lStartPort To lEndPort
        frmCustomize.wskTestConnection(i).Close: DoEvents
        frmCustomize.wskTestConnection(i).Connect lServer, i
    Next i
    NewPortScan = lPortScans.pCount
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function NewPortScan(lStartPort As Long, lEndPort As Long) As Integer"
End Function

Public Function FindPortScanIndex(lServer As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To lPortScans.pCount
    If LCase(lServer) = LCase(lServers.sServer(lPortScans.pPortScan(i).pServerIndex).sServer) Then
        FindPortScanIndex = i
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindPortScanIndex(lServer) As Integer"
End Function
