VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
