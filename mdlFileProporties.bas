Attribute VB_Name = "mdlFileProporties"
Option Explicit
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400
Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Private Type SHELLEXECUTEINFO
     cbSize As Long
     fMask As Long
     hWnd As Long
     lpVerb As String
     lpFile As String
     lpParameters As String
     lpDirectory As String
     nShow As Long
     hInstApp As Long
     lpIDList As Long
     lpClass As String
     hkeyClass As Long
     dwHotKey As Long
     hIcon As Long
     hProcess As Long
End Type

Public Sub ShowFileProperties(FormHwnd As Long, sFilename As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim udtSEI As SHELLEXECUTEINFO
With udtSEI
       .cbSize = Len(udtSEI)
       .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
       .hWnd = FormHwnd
       .lpVerb = "properties"
       .lpFile = sFilename
       .lpParameters = vbNullChar
       .lpDirectory = vbNullChar
       .nShow = 0
       .hInstApp = 0
       .lpIDList = 0
End With
Call ShellExecuteEX(udtSEI)
If udtSEI.hInstApp <= 32 Then
    ProcessReplaceString sFileNotFound, lSettings.sActiveServerForm.txtIncoming, sFilename
    If lSettings.sGeneralPrompts = True Then
        MsgBox sFilename & "not found, There is an error", vbCritical, "Error"
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ShowFileProperties(FormHwnd As Long, sFileName As String)"
End Sub
