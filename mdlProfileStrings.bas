Attribute VB_Name = "mdlProfileStrings"
Option Explicit
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Type gINIFiles
    iNicklistMenu As String
    iText As String
    iQueryMenu As String
    iChannelMenu As String
    iBlacklist As String
    iStatusMenu As String
    iAutoConnect As String
    iAutoPerform As String
    iInitialValues As String
    iBots As String
    iScripts As String
    iBotCommands As String
    iAlternates As String
    iSpectrum As String
    iAutoJoin As String
    iNotify As String
    iPlaylist As String
    iMedia As String
    iIRC As String
    iServers As String
    iIRCServer As String
    iChanFolder As String
    iErrorLog As String
End Type
Enum eINIFiles
    iAlternates = 0
    iAutoConnect = 1
    iAutoJoin = 2
    iAutoPerform = 3
    iBlacklist = 4
    iBotCommands = 5
    iBots = 6
    iChanFolder = 7
    iChannelMenu = 8
    iErrorLog = 9
    iInitialValues = 10
    iIRC = 11
    iIRCServer = 12
    iMedia = 13
    iNicklistMenu = 14
    iNotify = 15
    iPlaylist = 16
    iQueryMenu = 17
    iScripts = 18
    iServers = 19
    iStatusMenu = 20
    iSpectrum = 21
    iText = 22
End Enum
Private lINIFiles As gINIFiles

Public Sub SetPlaylistINIFile(lFileName As String)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
lINIFiles.iPlaylist = lFileName
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetPlaylistINIFile(lFileName As String)"
    Err.Clear
End Sub

Public Function GetINIFile(lType As eINIFiles) As String
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
Select Case lType
Case iAlternates
    GetINIFile = lINIFiles.iAlternates
Case iAutoConnect
    GetINIFile = lINIFiles.iAutoConnect
Case iAutoJoin
    GetINIFile = lINIFiles.iAutoJoin
Case iAutoPerform
    GetINIFile = lINIFiles.iAutoPerform
Case iBlacklist
    GetINIFile = lINIFiles.iBlacklist
Case iBotCommands
    GetINIFile = lINIFiles.iBotCommands
Case iBots
    GetINIFile = lINIFiles.iBots
Case iChanFolder
    GetINIFile = lINIFiles.iChanFolder
Case iChannelMenu
    GetINIFile = lINIFiles.iChannelMenu
Case iErrorLog
    GetINIFile = lINIFiles.iErrorLog
Case iInitialValues
    GetINIFile = lINIFiles.iInitialValues
Case iIRC
    GetINIFile = lINIFiles.iIRC
Case iIRCServer
    GetINIFile = lINIFiles.iIRCServer
Case iMedia
    GetINIFile = lINIFiles.iMedia
Case iNicklistMenu
    GetINIFile = lINIFiles.iNicklistMenu
Case iNotify
    GetINIFile = lINIFiles.iNotify
Case iPlaylist
    GetINIFile = lINIFiles.iPlaylist
Case iQueryMenu
    GetINIFile = lINIFiles.iQueryMenu
Case iScripts
    GetINIFile = lINIFiles.iScripts
Case iServers
    GetINIFile = lINIFiles.iServers
Case iStatusMenu
    GetINIFile = lINIFiles.iStatusMenu
Case iSpectrum
    GetINIFile = lINIFiles.iSpectrum
Case iText
    GetINIFile = lINIFiles.iText
End Select
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Function GetINIFile(lType As eINIFiles) As String"
    Err.Clear
End Function

Public Sub SetINIFiles()
On Local Error Resume Next
'MsgBox "SetINIFiles"
lINIFiles.iQueryMenu = App.Path & "\data\config\menu\query.mnu"
lINIFiles.iStatusMenu = App.Path & "\data\config\menu\status.mnu"
lINIFiles.iNicklistMenu = App.Path & "\data\config\menu\nicklist.mnu"
lINIFiles.iChannelMenu = App.Path & "\data\config\menu\channel.mnu"
lINIFiles.iAutoConnect = App.Path & "\data\config\autoconnect.ini"
lINIFiles.iIRCServer = App.Path & "\data\config\server\server.ini"
lINIFiles.iSpectrum = App.Path & "\data\config\fixed\spectrum.ini"
lINIFiles.iServers = App.Path & "\data\config\fixed\servers.ini"
lINIFiles.iBotCommands = App.Path & "\data\config\fixed\botcommands.ini"
lINIFiles.iScripts = App.Path & "\data\config\fixed\scripts.ini"
'If Len(lINIFiles.iText) <> 0 Then lINIFiles.iText = App.Path & "\data\config\fixed\text.ini"
lINIFiles.iText = App.Path & "\data\config\fixed\text.ini"
lINIFiles.iIRC = App.Path & "\data\config\settings.ini"
lINIFiles.iAutoPerform = App.Path & "\data\config\autoPerform.ini"
lINIFiles.iInitialValues = App.Path & "\data\config\initialaudiovalues.ini"
lINIFiles.iBots = App.Path & "\data\config\bots.ini"
lINIFiles.iErrorLog = App.Path & "\data\config\errorlog.ini"
lINIFiles.iAlternates = App.Path & "\data\config\alternates.ini"
lINIFiles.iChanFolder = App.Path & "\data\config\channelfolder.ini"
lINIFiles.iAutoJoin = App.Path & "\data\config\autojoinchannels.ini"
lINIFiles.iPlaylist = App.Path & "\data\config\playlist.ini"
lINIFiles.iNotify = App.Path & "\data\config\notify.ini"
lINIFiles.iBlacklist = App.Path & "\data\config\blacklist.ini"
End Sub

Public Function ReadINI(ByVal lFile As String, ByVal Section As String, ByVal Key As String, Optional lDefault As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, RetVal As String, Worked As Integer
RetVal = String$(255, 0)
Worked = GetPrivateProfileString(Section, Key, "", RetVal, Len(RetVal), lFile)
If Worked = 0 Then
    ReadINI = lDefault
Else
    ReadINI = Left(RetVal, InStr(RetVal, Chr(0)) - 1)
End If
End Function

Public Sub WriteINI(ByVal lFile As String, ByVal Section As String, ByVal Key As String, ByVal Value As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
WritePrivateProfileString Section, Key, Value, lFile
End Sub

Public Sub SetTextIniFile(lData As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If DoesFileExist(lData) = True Then
    lINIFiles.iText = lData
Else
    MsgBox "Error"
End If
End Sub
