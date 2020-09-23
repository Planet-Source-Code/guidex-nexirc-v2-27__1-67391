Attribute VB_Name = "mdlStats"
Option Explicit
Private Const ERROR_SUCCESS = 0
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_SZ = 1
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Public Enum HKEYs
    eHKEY_CLASSES_ROOT = &H80000000
    eHKEY_CURRENT_USER = &H80000001
    eHKEY_LOCAL_MACHINE = &H80000002
    eHKEY_USERS = &H80000003
    eHKEY_PERFORMANCE_DATA = &H80000004
    eHKEY_CURRENT_CONFIG = &H80000005
    eHKEY_DYN_DATA = &H80000006
End Enum
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As String, lpcbData As Long) As Long

Public Function GetValue(ByVal PredefinedKey As HKEYs, ByVal KeyName As String, ByVal ValueName As String, Optional ComputerName As String) As Variant
On Error GoTo ErrHand
Dim GetHandle As Long, hKey As Long, lpData As String, lpDataDWORD As Long, lpcbData As Long, lpType As Long, lReturnCode As Long, lhRemoteRegistry As Long
If Left$(KeyName, 1) = "\" Then KeyName = Right$(KeyName, Len(KeyName) - 1)
If ComputerName = "" Then
    GetHandle = RegOpenKeyEx(PredefinedKey, KeyName, 0, KEY_ALL_ACCESS, hKey)
Else
    lReturnCode = RegConnectRegistry(ComputerName, PredefinedKey, lhRemoteRegistry)
    GetHandle = RegOpenKeyEx(lhRemoteRegistry, KeyName, 0, KEY_ALL_ACCESS, hKey)
End If
If GetHandle = ERROR_SUCCESS Then
    lpcbData = 255
    lpData = String(lpcbData, Chr(0))
    GetHandle = RegQueryValueEx(hKey, ValueName, 0, lpType, ByVal lpData, lpcbData)
    If GetHandle = ERROR_SUCCESS Then
        Select Case lpType
            Case REG_SZ
                GetHandle = RegQueryValueExString(hKey, ValueName, 0, lpType, ByVal lpData, lpcbData)
                If GetHandle = 0 Then
                    GetValue = Left$(lpData, lpcbData - 1)
                Else
                    GetValue = ""
                End If
            Case REG_DWORD
                GetHandle = RegQueryValueEx(hKey, ValueName, 0, lpType, lpDataDWORD, lpcbData)
                If GetHandle = 0 Then
                    GetValue = CLng(lpDataDWORD)
                Else
                    GetValue = 0
                End If
            Case REG_BINARY
                GetHandle = RegQueryValueEx(hKey, ValueName, 0, lpType, lpDataDWORD, lpcbData)
                If GetHandle = 0 Then
                    GetValue = CByte(lpDataDWORD)
                Else
                    GetValue = 0
                End If
        End Select
    End If
    RegCloseKey hKey
End If
Exit Function
ErrHand:
    Err.Raise "11002", "clsRegistry", "GetValue"
End Function

Public Sub ShowSystemStats(lForm As Form, Optional lOS As Boolean, Optional lPC As Boolean, Optional lCPUCount As Boolean, Optional lProc1 As Boolean, Optional lProc2 As Boolean, Optional SCSI As Boolean, Optional lConsoleOnly As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lProcessor1(4) As String, lProcessor2(4) As String, lSCSI1(2) As String, lSCSI2(2) As String, lSCSI3(2) As String, lSCSI4(2) As String, lPCName As String, i As Integer, c As Integer, lWinInfo(5) As String, msg(6) As String, t As ctlTBox
If lConsoleOnly = True Then frmSystemStatsConsole.txtSystemStats.Text = ""
If lOS = True Then
    lWinInfo(1) = GetValue(eHKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName")
    lWinInfo(2) = GetValue(eHKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "BuildLab")
    lWinInfo(3) = GetValue(eHKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CSDVersion")
    lWinInfo(4) = GetValue(eHKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentBuildNumber")
    lWinInfo(5) = GetValue(eHKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentVersion")
End If
If lPC = True Then
    lPCName = GetValue(eHKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Control\ComputerName\ComputerName", "ComputerName")
End If
If lCPUCount = True Then
    c = Int(GetValue(eHKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Control\Session Manager\Environment", "NUMBER_OF_PROCESSORS"))
End If
If lProc1 = True Then
    lProcessor1(1) = GetValue(eHKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\CentralProcessor\0", "ProcessorNameString")
    lProcessor1(2) = GetValue(eHKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\CentralProcessor\0", "Identifier")
    lProcessor1(3) = GetValue(eHKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\CentralProcessor\0", "VendorIdentifier")
    lProcessor1(4) = GetValue(eHKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\CentralProcessor\0", "~MHz")
End If
If lProc2 = True Then
    lProcessor2(1) = GetValue(eHKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\CentralProcessor\1", "ProcessorNameString")
    lProcessor2(2) = GetValue(eHKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\CentralProcessor\1", "Identifier")
    lProcessor2(3) = GetValue(eHKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\CentralProcessor\1", "VendorIdentifier")
    lProcessor2(4) = GetValue(eHKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\CentralProcessor\1", "~MHz")
End If
If SCSI = True Then
    lSCSI1(1) = GetValue(eHKEY_LOCAL_MACHINE, "HARDWARE\DEVICEMAP\Scsi\Scsi Port 0\Scsi Bus 0\Target Id 0\Logical Unit Id 0", "Identifier")
    lSCSI1(2) = GetValue(eHKEY_LOCAL_MACHINE, "HARDWARE\DEVICEMAP\Scsi\Scsi Port 0\Scsi Bus 0\Target Id 0\Logical Unit Id 0", "Type")
    lSCSI2(1) = GetValue(eHKEY_LOCAL_MACHINE, "HARDWARE\DEVICEMAP\Scsi\Scsi Port 0\Scsi Bus 0\Target Id 1\Logical Unit Id 0", "Identifier")
    lSCSI2(2) = GetValue(eHKEY_LOCAL_MACHINE, "HARDWARE\DEVICEMAP\Scsi\Scsi Port 0\Scsi Bus 0\Target Id 1\Logical Unit Id 0", "Type")
    lSCSI3(1) = GetValue(eHKEY_LOCAL_MACHINE, "HARDWARE\DEVICEMAP\Scsi\Scsi Port 1\Scsi Bus 0\Target Id 0\Logical Unit Id 0", "Identifier")
    lSCSI3(2) = GetValue(eHKEY_LOCAL_MACHINE, "HARDWARE\DEVICEMAP\Scsi\Scsi Port 1\Scsi Bus 0\Target Id 0\Logical Unit Id 0", "Type")
    lSCSI4(1) = GetValue(eHKEY_LOCAL_MACHINE, "HARDWARE\DEVICEMAP\Scsi\Scsi Port 1\Scsi Bus 0\Target Id 1\Logical Unit Id 0", "Identifier")
    lSCSI4(2) = GetValue(eHKEY_LOCAL_MACHINE, "HARDWARE\DEVICEMAP\Scsi\Scsi Port 1\Scsi Bus 0\Target Id 1\Logical Unit Id 0", "Type")
End If
If Len(lPCName) <> 0 Then msg(1) = "PC Name: " & lPCName
If c <> 0 Then msg(2) = "Number of Processors: " & Trim(Str(c))
For i = 0 To c
    Select Case i
    Case 0
        If Len(lProcessor1(1)) <> 0 Then msg(3) = "Processor " & Trim(Str((i + 1))) & ": " & lProcessor1(1) & ", Identifier: " & lProcessor1(2) & ", Vendor: " & lProcessor1(3) & ", Megahurtz: " & lProcessor1(4)
    Case 1
        If Len(lProcessor2(1)) <> 0 Then msg(4) = "Processor " & Trim(Str((i + 1))) & ": " & lProcessor2(1) & ", Identifier: " & lProcessor2(2) & ", Vendor: " & lProcessor2(3) & ", Megahurtz: " & lProcessor2(4)
    End Select
Next i
For i = 1 To 4
    Select Case i
    Case 1
        If Len(lSCSI1(1)) <> 0 Then msg(5) = "Device " & Trim(Str(i)) & ": " & lSCSI1(1) & ", Type: " & lSCSI1(2)
    Case 2
        If Len(lSCSI2(1)) <> 0 Then msg(5) = msg(5) & ", Device " & Trim(Str(i)) & ": " & lSCSI2(1) & ", Type: " & lSCSI2(2)
    Case 3
        If Len(lSCSI3(1)) <> 0 Then msg(5) = msg(5) & ", Device " & Trim(Str(i)) & ": " & lSCSI3(1) & ", Type: " & lSCSI3(2)
    Case 4
        If Len(lSCSI4(1)) <> 0 Then msg(5) = msg(5) & ", Device " & Trim(Str(i)) & ": " & lSCSI4(1) & ", Type: " & lSCSI4(2)
    End Select
Next i
If Len(lWinInfo(1)) <> 0 Then msg(6) = "Operating System: " & lWinInfo(1) & ", Version " & lWinInfo(5) & ", Build: " & lWinInfo(4) & ", SP: " & lWinInfo(3) & " (" & lWinInfo(2) & ")"
If lSettings.sActiveServerForm.tcp.State = sckConnected Or lConsoleOnly = True Then
    If LCase(mdiNexIRC.ActiveForm.Name) = "frmchannel" Then
        Set t = mdiNexIRC.ActiveForm.txtIncoming
    ElseIf LCase(mdiNexIRC.ActiveForm.Name) = "frmstatus" Then
        Set t = mdiNexIRC.ActiveForm.txtIncoming
    End If
    If Len(msg(1)) <> 0 And lPC = True Then
        If lConsoleOnly = True Then
            frmSystemStatsConsole.txtSystemStats.Text = Trim(msg(1))
        Else
            ProcessReplaceString sPm, t, lSettings.sNickname, "", msg(1): Pause 2
            lSettings.sActiveServerForm.tcp.SendData "PRIVMSG " & mdiNexIRC.ActiveForm.Tag & " :" & msg(1) & vbCrLf
        End If
    End If
    If Len(msg(2)) <> 0 And lCPUCount = True Then
        If lConsoleOnly = True Then
            frmSystemStatsConsole.txtSystemStats.Text = Trim(frmSystemStatsConsole.txtSystemStats.Text & vbCrLf & msg(2))
        Else
            ProcessReplaceString sPm, t, lSettings.sNickname, "", msg(2): Pause 2
            lSettings.sActiveServerForm.tcp.SendData "PRIVMSG " & mdiNexIRC.ActiveForm.Tag & " :" & msg(2) & vbCrLf
        End If
    End If
    If Len(msg(3)) <> 0 And lProc1 = True Then
        If lConsoleOnly = True Then
            frmSystemStatsConsole.txtSystemStats.Text = Trim(frmSystemStatsConsole.txtSystemStats.Text & vbCrLf & msg(3))
        Else
            ProcessReplaceString sPm, t, lSettings.sNickname, "", msg(3): Pause 2
            lSettings.sActiveServerForm.tcp.SendData "PRIVMSG " & mdiNexIRC.ActiveForm.Tag & " :" & msg(3) & vbCrLf
        End If
    End If
    If Len(msg(4)) <> 0 And lProc2 = True Then
        If lConsoleOnly = True Then
            frmSystemStatsConsole.txtSystemStats.Text = Trim(frmSystemStatsConsole.txtSystemStats.Text & vbCrLf & msg(4))
        Else
            ProcessReplaceString sPm, t, lSettings.sNickname, "", msg(4): Pause 2
            lSettings.sActiveServerForm.tcp.SendData "PRIVMSG " & mdiNexIRC.ActiveForm.Tag & " :" & msg(4) & vbCrLf
        End If
    End If
    If Len(msg(5)) <> 0 And SCSI = True Then
        If lConsoleOnly = True Then
            frmSystemStatsConsole.txtSystemStats.Text = Trim(frmSystemStatsConsole.txtSystemStats.Text & vbCrLf & msg(5))
        Else
            ProcessReplaceString sPm, t, lSettings.sNickname, "", msg(5): Pause 2
            lSettings.sActiveServerForm.tcp.SendData "PRIVMSG " & mdiNexIRC.ActiveForm.Tag & " :" & msg(5) & vbCrLf
        End If
    End If
    If Len(msg(6)) <> 0 And lOS = True Then
        If lConsoleOnly = True Then
            frmSystemStatsConsole.txtSystemStats.Text = Trim(frmSystemStatsConsole.txtSystemStats.Text & vbCrLf & msg(6))
        Else
            ProcessReplaceString sPm, t, lSettings.sNickname, "", msg(6): Pause 2
            lSettings.sActiveServerForm.tcp.SendData "PRIVMSG " & mdiNexIRC.ActiveForm.Tag & " :" & msg(6) & vbCrLf
        End If
    End If
    If lConsoleOnly = True Then frmSystemStatsConsole.Show
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ShowSystemStats()"
End Sub

