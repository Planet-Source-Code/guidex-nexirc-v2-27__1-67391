VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type gVariable
    vName As String
    vData As String
End Type
Private Type gVariables
    vCount As Integer
    vVariable(150) As gVariable
End Type
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByVal lprc As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Const LR_LOADFROMFILE = &H10
Private Const IMAGE_BITMAP = 0
Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400&
Private Const MF_CHECKED = &H8&
Private Const MF_DISABLED = &H2&
Private Const MF_GRAYED = &H1&
Private Const MF_MENUBARBREAK = &H20&
Private Const MF_MENUBREAK = &H40&
Private Const MF_POPUP = &H10&
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_RETURNCMD = &H100&
Private Type FMENUITEM
    lIcon As Long
    lId As Long
    lFlags As Long
    sCaption As String
    sCommand As String
    sMenuName As String
End Type
Private Type FMENUITEMICON
    lIcon As Long
    sFilename As String
End Type
Private Type FMENUITEMCMDSTR
    bGoteExecWinStyle As Boolean
    bGoteMBType As Boolean
    bGotlMBButtons As Boolean
    bGotsCmd As Boolean
    bGotsExec As Boolean
    bGotsMBCaption As Boolean
    bGotsMBText As Boolean
    eExecWinStyle As VbAppWinStyle
    eMBType As VbMsgBoxStyle
    lMBButtons As Long
    sCmd As String
    sExec As String
    sMBCaption As String
    sMBText As String
End Type
Private m_tMenuItems() As FMENUITEM
Private m_lIcons() As FMENUITEMICON
Private m_lPopupMenus() As Long
Private m_lNextAvailableItemID As Long
Private m_lParentHWND As Long
Private m_lMainHMENU As Long
Private m_lNumIcons As Long
Private lVariables As gVariables
Public Event ItemClicked(sKey As String)

Public Property Get OwnerHWND() As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
OwnerHWND = m_lParentHWND
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Property Get OwnerHWND() As Long"
End Property

Public Property Let OwnerHWND(ByVal lVal As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
m_lParentHWND = lVal
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Property Get OwnerHWND() As Long"
End Property

Private Sub Class_Initialize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReDim Preserve m_tMenuItems(0 To 0)
ReDim Preserve m_lIcons(0 To 0)
ReDim Preserve m_lPopupMenus(0 To 0)
m_lIcons(0).lIcon = -1
m_lMainHMENU = CreatePopupMenu()
m_lNextAvailableItemID = 1
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Class_Initialize()"
End Sub

Private Sub Class_Terminate()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Call DestroyAllMenus
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Class_Terminate()"
End Sub

Private Function CacheIcon(ByVal sFile As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim a As Long
sFile = UCase$(sFile)
For a = 0 To UBound(m_lIcons)
    If (m_lIcons(a).sFilename = UCase$(sFile)) Then Exit Function
Next a
ReDim Preserve m_lIcons(0 To IIf(m_lIcons(0).lIcon = -1, 0, UBound(m_lIcons) + 1))
With m_lIcons(UBound(m_lIcons))
    .lIcon = LoadImage(-1, sFile, IMAGE_BITMAP, 13, 13, LR_LOADFROMFILE)
    .sFilename = sFile
End With
m_lNumIcons = (m_lNumIcons + 1)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Function CacheIcon(ByVal sFile As String)"
End Function

Public Function LoadMenus(ByVal sFile As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim tTMI As FMENUITEM
Dim a As Long, b As Long, c As Long, sTmpStr As String, lRet As Long
Dim lNumSections As Long, lNumItems As Long, lNumIcons As Long
lNumSections = CLng(GetPrivateProfileInt("INDEX", "NumSections", 0, sFile))
lNumIcons = CLng(GetPrivateProfileInt("ICONS", "NumIcons", 0, sFile))
For a = 1 To lNumIcons
    sTmpStr = ReadINI(sFile, "ICONS", "Icon" & CStr(a))
    If Len(sTmpStr) <> 0 Then sTmpStr = App.Path & "\data\images\menu\" & sTmpStr
    If (sTmpStr <> "") Then Call CacheIcon(sTmpStr)
Next a
If (lNumSections > 0) Then
    For a = 0 To (lNumSections - 1)
        tTMI.sMenuName = UCase$(ReadINI(sFile, CStr(a), "MenuName"))
        If (tTMI.sMenuName <> "") Then
            lNumItems = GetPrivateProfileInt(CStr(a), "NumItems", 0, sFile)
            If (lNumItems > 0) Then
                For b = 1 To lNumItems
                    tTMI.sCaption = ReadINI(sFile, CStr(a), "Item" & CStr(b))
                    tTMI.sCommand = ReadINI(sFile, CStr(a), "Item" & CStr(b) & "Command")
                    tTMI.lFlags = MF_STRING
                    sTmpStr = UCase$(ReadINI(sFile, CStr(a), "Item" & CStr(b) & "State"))
                    If (InStr(1, sTmpStr, "CHECKED", vbTextCompare) > 0) Then tTMI.lFlags = (tTMI.lFlags Or MF_CHECKED)
                    If (InStr(1, sTmpStr, "DISABLED", vbTextCompare) > 0) Then tTMI.lFlags = (tTMI.lFlags Or MF_DISABLED)
                    If (InStr(1, sTmpStr, "GRAYED", vbTextCompare) > 0) Then tTMI.lFlags = (tTMI.lFlags Or MF_GRAYED)
                    tTMI.lIcon = GetPrivateProfileInt(CStr(a), "Item" & CStr(b) & "Icon", 0, sFile)
                    If (tTMI.lIcon > m_lNumIcons) Then
                        tTMI.lIcon = 0
                        If lSettings.sGeneralPrompts = True Then
                            Call MsgBox("The requested icon # for [" & tTMI.sMenuName & "]Item" & CStr(b) & " is larger than the amount of icons specified. Defaulting to icon #0 (no icon). Please fix this error in your menu.", vbExclamation, "NexIRC")
                        End If
                    End If
                    If (tTMI.sCaption = "-") Then
                        tTMI.lFlags = MF_SEPARATOR
                        tTMI.sCaption = ""
                    End If
                    If (tTMI.sCaption = "") Then tTMI.sCaption = " "
                    If (tTMI.sCommand = "") Then
                        tTMI.sCommand = "NX_NONE();"
                    End If
                    ReDim Preserve m_tMenuItems(0 To (UBound(m_tMenuItems) + 1))
                    m_tMenuItems(UBound(m_tMenuItems)) = tTMI
                    With tTMI
                        .lFlags = 0&
                        .lIcon = 0&
                        .lId = 0&
                        .sCaption = ""
                        .sCommand = ""
                    End With
                Next b
            End If
        Else
            If (a = 0) Then
                If lSettings.sGeneralPrompts = True Then
                    Call MsgBox("ROOT menu was not the first section encountered." & vbCrLf & "ROOT must -always- be the first section in the file." & vbCrLf & vbCrLf & "Menu creation terminated. Please fix this error in your menu.", vbCritical, "FMenu")
                End If
                ReDim m_tMenuItems(0 To 0) As FMENUITEM
                Exit Function
            End If
        End If
    Next a
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function LoadMenus(ByVal sFile As String)"
End Function

Private Function GetIndexesOfAllChildren(ByVal sMenuName As String, ByRef lIndexes() As Long) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim a As Long
ReDim lIndexes(0 To 0) As Long
lIndexes(0) = -1
sMenuName = UCase$(sMenuName)
For a = 0 To UBound(m_tMenuItems)
    If (m_tMenuItems(a).sMenuName = sMenuName) Then
        ReDim Preserve lIndexes(0 To IIf(lIndexes(0) = -1, 0, (UBound(lIndexes) + 1)))
        lIndexes(UBound(lIndexes)) = a
    End If
Next a
GetIndexesOfAllChildren = IIf(lIndexes(0) > -1, UBound(lIndexes) + 1, -1)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Function GetIndexesOfAllChildren(ByVal sMenuName As String, ByRef lIndexes() As Long) As Long"
End Function

Private Function BuildMenu(ByVal DestHMENU As Long, ByVal sMenuName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim a As Long, lItems() As Long, sTmp As String
Call GetIndexesOfAllChildren(sMenuName, lItems())
If (lItems(0) > -1) Then
    For a = 0 To UBound(lItems)
        With m_tMenuItems(lItems(a))
            .lId = m_lNextAvailableItemID
            Call AppendMenu(DestHMENU, ByVal .lFlags, ByVal .lId, .sCaption)
            If (.lIcon >= 1) Then
                Call SetMenuItemBitmaps(DestHMENU, ByVal .lId, MF_BYCOMMAND, ByVal m_lIcons(.lIcon - 1).lIcon, ByVal m_lIcons(.lIcon - 1).lIcon)
            End If
            If (UCase$(Mid$(.sCommand, 1, 14)) = "NX_INHERITMENU") Then
                sTmp = UCase$(.sCommand)
                If (Right$(sTmp, 2) = ");") Then
                    sTmp = Mid$(sTmp, 1, Len(sTmp) - 2)
                End If
                Dim lUB As Long
                lUB = (UBound(m_lPopupMenus) + 1)
                ReDim Preserve m_lPopupMenus(0 To lUB)
                m_lPopupMenus(lUB) = CreatePopupMenu()
                sTmp = Mid$(sTmp, 16)
                If (m_lPopupMenus(lUB) > 0) Then
                    Call ModifyMenu(m_lMainHMENU, .lId, .lFlags Or MF_POPUP, m_lPopupMenus(lUB), .sCaption)
                    .lId = m_lPopupMenus(lUB)
                    Call BuildMenu(m_lPopupMenus(lUB), sTmp)
                End If
                .sCommand = "NX_NONE();"
            End If
        End With
        m_lNextAvailableItemID = (m_lNextAvailableItemID + 1)
    Next a
Else
    Call AppendMenu(DestHMENU, MF_STRING Or MF_DISABLED, 0, "Error Building Menu")
    Call AppendMenu(DestHMENU, MF_SEPARATOR, 0, "-")
    Call AppendMenu(DestHMENU, MF_STRING Or MF_DISABLED, 0, "InheritMenu failed. The")
    Call AppendMenu(DestHMENU, MF_STRING Or MF_DISABLED, 0, "requested section cannot")
    Call AppendMenu(DestHMENU, MF_STRING Or MF_DISABLED, 0, "be located.")
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Function BuildMenu(ByVal DestHMENU As Long, ByVal sMenuName As String)"
End Function

Private Function FindItemByID(ByVal lId As Long) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim a As Long
For a = 0 To UBound(m_tMenuItems)
    With m_tMenuItems(a)
        If .lId = lId Then
            FindItemByID = a
            Exit Function
        End If
    End With
Next a
FindItemByID = -1
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Function FindItemByID(ByVal lID As Long) As Long"
End Function

Private Function DestroyAllMenus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim a As Long
Call DestroyMenu(m_lMainHMENU)
For a = 0 To UBound(m_lPopupMenus)
    If (m_lPopupMenus(a) > 0) Then Call DestroyMenu(m_lPopupMenus(a))
Next a
For a = 0 To UBound(m_tMenuItems)
    If (m_tMenuItems(a).lIcon > 0) Then Call DeleteObject(m_tMenuItems(a).lIcon)
Next a
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Function DestroyAllMenus()"
End Function

Public Sub RunScriptFile(lFileName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, msg3 As String
If DoesFileExist(lFileName) = True Then
    msg = ReadFile(lFileName)
    If Len(msg) <> 0 Then
        Do Until Len(msg) = 0
            If InStr(msg, Chr(13)) Then
                msg2 = Trim(Left(msg, 1) & Parse(msg, Left(msg, 1), Chr(13)))
                msg = Trim(Right(msg, Len(msg) - Len(msg2) - 2))
            Else
                msg2 = Trim(msg)
                msg = ""
            End If
            RunCommand msg2, mdiNexIRC
        Loop
    End If
Else
    ProcessReplaceString sFileNotFound, mdiNexIRC.ActiveForm.txtIncoming, lFileName
    DoColorSep lSettings.sActiveServerForm.txtIncoming
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub RunScriptFile(lFilename As String)"
End Sub

Private Function FindVariableIndex(lName As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lName) <> 0 Then
    For i = 1 To 150
        If LCase(lName) = LCase(lVariables.vVariable(i).vName) Then
            FindVariableIndex = i
            Exit Function
        End If
    Next i
    FindVariableIndex = 151
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Function FindVariableIndex(lName As String) As Integer"
End Function

Public Function RunCommand(ByVal sCmd As String, lForm As Form) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lParms As String, lSPLT() As String, i As Integer, lcmd As String, lSpltCount As Integer, msg As String, F As Integer, c As Integer, msg2 As String, lIfPositive As Boolean, var() As String
If Len(sCmd) = 0 Then Exit Function
lParms = Parse(sCmd, "(", ")")
lcmd = Left(sCmd, Len(sCmd) - Len(lParms) - 3)
If Left(lcmd, 2) = "//" Then Exit Function
If InStr(lParms, "OpenDialog[") Then
    msg = Parse(lParms, "OpenDialog[", "]")
    If Len(msg) <> 0 Then
        msg = Right(msg, Len(msg) - 10)
        msg2 = "OpenDialog[" & msg & "]"
        var = Split(msg, "|")
        msg = ""
        var(0) = Right(var(0), Len(var(0)) - 11)
        var(0) = Replace(var(0), "{", "(")
        var(0) = Replace(var(0), "}", ")")
        var(0) = Replace(var(0), "+", "|")
        Select Case LCase(var(0))
        Case "$supported"
            var(0) = "Audio (*.wav;*.mp3;*.wma;*.snd;*.au;*.ogg)|*.wav;*.mp3;*.wma;*.snd;*.au;*.ogg|All Files (*.*)|*.*|"
        Case "$mp3"
            var(0) = "MP3 Files (*.mp3)|*.mp3|All Files (*.*)|*.*|"
        End Select
        Select Case LCase(var(2))
        Case "$curdir"
            var(2) = CurDir & "\"
        Case "$apppath"
            var(2) = App.Path & "\"
        End Select
        msg = OpenDialog(lForm, var(0), var(1), var(2))
        If Len(msg) <> 0 Then
            If InStr(lParms, msg2) Then
                lParms = Replace(LCase(lParms), LCase(msg2), msg, , -1, vbTextCompare)
            End If
        End If
    End If
End If
If InStr(LCase(lParms), "returnvsvariable[") Then
    msg = Parse(lParms, "ariable[", "]")
    msg = Right(msg, Len(msg) - 7)
    msg2 = "ReturnVSVariable[" & msg & "]"
    'mdiNexIRC.ctlVBScript.ExecuteStatement "mdiNexIRC.Tag = " & msg
    If Len(mdiNexIRC.Tag) <> 0 Then
        lParms = Replace(lParms, msg2, mdiNexIRC.Tag)
        mdiNexIRC.Tag = ""
    End If
End If
If InStr(LCase(lParms), "returnvariable[") Then
RVAGAIN:
    msg = Parse(lParms, "ariable[", "]")
    msg = Right(msg, Len(msg) - 7)
    If Len(msg) <> 0 Then
        msg2 = "ReturnVariable[" & msg & "]"
        i = FindVariableIndex(msg)
        If i = 151 Then
            ProcessReplaceString sVariableNotDeclared, lSettings.sActiveServerForm.txtIncoming, msg
            Exit Function
        End If
        If i <> 0 And Len(lVariables.vVariable(i).vData) <> 0 Then
            lParms = Replace(lParms, msg2, lVariables.vVariable(i).vData)
        Else
            lParms = Replace(lParms, msg2, " ")
        End If
        If InStr(LCase(lParms), "returnvariable[") Then GoTo RVAGAIN
    End If
End If
If InStr(LCase(lParms), "copyvariable[") Then
    msg = Parse(lParms, "copyvariable[", "]")
    var = Split(msg, "|")
    F = 0
    c = 0
    For i = 1 To 150
        If InStr(LCase(lVariables.vVariable(i).vData), LCase(var(0))) Then
            F = i
            Exit For
        End If
        If InStr(LCase(lVariables.vVariable(i).vData), LCase(var(1))) Then
            c = i
            Exit For
        End If
    Next i
    If F <> 0 And c <> 0 Then lVariables.vVariable(F).vData = lVariables.vVariable(c).vData
End If
lParms = Replace(lParms, "$ircsettings", GetINIFile(iIRC), 1, -1, vbTextCompare)
If InStr(LCase(lParms), "readini[") Then
    Dim ini() As String
    ini = Split(Parse(lParms, "[", "]"), "|")
    msg = ReadINI(ini(0), ini(1), ini(2), ini(3))
    msg2 = "r" & Parse(LCase(lParms), "readini", "]") & "]"
    lParms = Replace(LCase(lParms), LCase(msg2), msg, 1, -1, vbTextCompare)
End If
If InStr(lParms, "$activechannel") Then lParms = Replace(lParms, "$activechannel", mdiNexIRC.ActiveForm.Tag, 1, -1, vbTextCompare)
If InStr(lParms, "$currentaudio") Then lParms = Replace(lParms, "$currentaudio", lPlayback.pCurrentFile, 1, -1, vbTextCompare)
If InStr(lParms, "$server") Then lParms = Replace(lParms, "$server", lSettings.sServer, 1, -1, vbTextCompare)
If InStr(lParms, "$port2") Then lParms = Replace(lParms, "$port2", Val(Mid(lSettings.sPort, 1)), 1, -1, vbTextCompare)
If InStr(lParms, "$port") Then lParms = Replace(lParms, "$port", lSettings.sPort, 1, -1, vbTextCompare)
If InStr(lParms, "$apppath") Then lParms = Replace(lParms, "$apppath", App.Path, 1, -1, vbTextCompare)
If InStr(lParms, "$time") Then lParms = Replace(lParms, "$time", Time, 1, -1, vbTextCompare)
If InStr(lParms, "$date") Then lParms = Replace(lParms, "$date", Date, 1, -1, vbTextCompare)
If InStr(lParms, "$mynick") Then lParms = Replace(lParms, "$mynick", lSettings.sNickname, 1, -1, vbTextCompare)
If InStr(lParms, "$myip") Then lParms = Replace(lParms, "$myip", lSettings.sActiveServerForm.tcp.LocalIP, 1, -1, vbTextCompare)
If InStr(lParms, "$querynick") Then lParms = Replace(lParms, "$querynick", mdiNexIRC.ActiveForm.Caption)
If InStr(lParms, "$password") Then lParms = Replace(lParms, "$password", lSettings.sPassword, 1, -1, vbTextCompare)
If InStr(lParms, "$activeserver") Then lParms = Replace(lParms, "$activeserver", ReturnStatusWindowServer(FindStatusWindowIndexByTag(lSettings.sActiveServerForm.Tag)), 1, -1, vbTextCompare)
If InStr(lParms, "$activeport") Then lParms = Replace(lParms, "$activeport", Val(Mid(ReturnStatusWindowPort(FindStatusWindowIndexByTag(lSettings.sActiveServerForm.Tag)), 1)), 1, -1, vbTextCompare)
If InStr(lParms, "$lChannel") Then lParms = Replace(lParms, "$lChannel", mdiNexIRC.ActiveForm.Tag, 1, -1, vbTextCompare)
If InStr(lParms, "NULL") Then lParms = Replace(lParms, "NULL", "", 1, -1, vbTextCompare)
If InStr(lParms, "$nicklist") Then
    lParms = Replace(lParms, "$nicklist", mdiNexIRC.ActiveForm.ReturnSelectedItem())
    lParms = Replace(lParms, "@", "")
    lParms = Replace(lParms, "+", "")
End If
lSPLT = Split(lParms, "::")
lSpltCount = UBound(lSPLT)
If InStr(LCase(lParms), "$input") Then
    If InStr(LCase(lSPLT(0)), "$input") Then
        msg = InputBox(lSPLT(1), lSPLT(2), lSPLT(3))
        msg2 = "$input::" & lSPLT(1) & "::" & lSPLT(2) & "::" & lSPLT(3)
        If Len(msg) <> 0 Then
            lParms = Replace(lParms, msg2, msg)
            lSPLT(1) = msg
        End If
    End If
    If InStr(LCase(lSPLT(1)), "$input") Then
        msg = InputBox(lSPLT(2), lSPLT(3), lSPLT(4))
        msg2 = "$input::" & lSPLT(2) & "::" & lSPLT(3) & "::" & lSPLT(4)
        If Len(msg) <> 0 Then
            lParms = Replace(lParms, msg2, msg)
            lSPLT(1) = msg
        End If
    End If
End If
lParms = Replace(lParms, "::", " ", 1, -1, vbTextCompare)
Select Case LCase(Trim(lcmd))
Case "writeini"
    WriteINI lSPLT(0), lSPLT(1), lSPLT(2), lSPLT(3)
    RunCommand = True
    Exit Function
Case "banaddress"
    lSettings.sRetrieveAddressFromWhoisForBan = True
    lSettings.sBanChannel = mdiNexIRC.ActiveForm.Tag
    lSettings.sActiveServerForm.tcp.SendData "WHOIS " & lSPLT(0) & vbCrLf
    RunCommand = True
    Exit Function
Case "showchannelproporties"
    ProcessReplaceString sRequestChannelInformation, ReturnChannelIncomingTBox(FindChannelIndex(mdiNexIRC.ActiveForm.Tag))
    lSettings.sActiveServerForm.tcp.SendData "MODE " & mdiNexIRC.ActiveForm.Tag & " +b" & vbCrLf
    RunCommand = True
    Exit Function
Case "addignore"
    AddToIgnore lSPLT(0)
    RunCommand = True
    Exit Function
Case "connect"
    If LCase(lSPLT(0)) <> "null" Then
        If Len(lSPLT(0)) <> 0 And Len(lSPLT(1)) <> 0 Then
            ConnectToIRC lSPLT(0), lSPLT(1), lForm
        Else
            ConnectToIRC lSettings.sServer, lSettings.sPort, lForm
        End If
    Else
        ConnectToIRC lSettings.sServer, lSettings.sPort, lForm
    End If
    RunCommand = True
    Exit Function
Case "addbot"
    frmAddBotCommand.Show
    frmAddBotCommand.txtNickname.Text = lSPLT(0)
    frmAddBotCommand.cboNicknameType.ListIndex = 1
    RunCommand = True
    Exit Function
Case "docolor"
    If Len(lSPLT(0)) <> 0 And Len(lSPLT(1)) <> 0 Then
        If LCase(lSPLT(0)) = "$activeserver" Then
            Call DoColor(lSettings.sActiveServerForm.txtIncoming, lSPLT(1))
        Else
            If lForm.Name = "mdiNexIRC" Then
                Call DoColor(mdiNexIRC.ActiveForm.txtIncoming, lSPLT(1))
            Else
                Call DoColor(lForm.txtIncoming, lSPLT(1))
            End If
        End If
    Else
        'DoColor lForm.txtIncoming, "" & Color.Notify & " Could not find the window: " & lSPLT(0)
    End If
    RunCommand = True
    Exit Function
Case "addtonotify"
    AddNotify lSPLT(0)
    RunCommand = True
    Exit Function
Case "playmp3"
    PlayFile lParms
    'PlayMP3 lParms
    RunCommand = True
    Exit Function
Case "playaudio"
    PlayFile lParms
    RunCommand = True
    Exit Function
Case "closeactivewindow"
    Unload mdiNexIRC.ActiveForm
    RunCommand = True
    Exit Function
Case "declarevariable"
    For i = 1 To 150
        If LCase(lVariables.vVariable(i).vName) = LCase(lParms) Then
            F = i
            Exit For
        End If
    Next i
    If F = 0 Then
        lVariables.vCount = lVariables.vCount + 1
        lVariables.vVariable(lVariables.vCount).vName = lParms
    End If
    RunCommand = True
    Exit Function
Case "addfiletoplaylist"
    AddToPlaylist lParms
    RunCommand = True
    Exit Function
Case "editmenu"
    Select Case LCase(lSPLT(0))
    Case "channel"
        frmMenuEditor.Show
        frmMenuEditor.cboMenu.ListIndex = 1
    Case "status"
        frmMenuEditor.Show
        frmMenuEditor.cboMenu.ListIndex = 0
    Case "lQuery"
        frmMenuEditor.Show
        frmMenuEditor.cboMenu.ListIndex = 2
    Case "nicklist"
        frmMenuEditor.Show
        frmMenuEditor.cboMenu.ListIndex = 3
    End Select
    RunCommand = True
    Exit Function
Case "queryuser"
    NewQuery lSPLT(0)
    RunCommand = True
    Exit Function

Case "getfiletitle"
    For i = 1 To 150
        If LCase(lVariables.vVariable(i).vName) = LCase(lSPLT(0)) Then
            lVariables.vVariable(i).vData = GetFileTitle(lVariables.vVariable(i).vData)
            Exit For
        End If
    Next i
    RunCommand = True
    Exit Function
    
Case "addautoconnect"
    AddToAutoConnect lSPLT(0), CLng(lSPLT(1))
    RunCommand = True
    Exit Function
Case "if"
    If LCase(lParms) = "true" Then lIfPositive = True
    RunCommand = True
    Exit Function
Case "togglemixer"
    If lSettings.sShowQuickmix = True Then
        ToggleMixer False
    Else
        ToggleMixer True
    End If
    RunCommand = True
    Exit Function
Case "savelog"
    msg = "log-" & Time$ & "-" & Date$ & ".log"
    SaveFile App.Path & "\data\logs\" & msg, lSettings.sActiveServerForm.txtIncoming
    ProcessReplaceString sSaveLog, mdiNexIRC.ActiveForm.txtIncoming, msg
    RunCommand = True
    Exit Function
Case "endif"
    lIfPositive = False
    RunCommand = True
    Exit Function
Case "clearactiveserverwindowincoming"
    lSettings.sActiveServerForm.txtIncoming.Text = ""
    RunCommand = True
    Exit Function
Case "setvariabledata"
    For i = 1 To 150
        If LCase(lVariables.vVariable(i).vName) = LCase(lSPLT(0)) Then
            lVariables.vVariable(i).vData = lSPLT(1)
            Exit For
        End If
    Next i
    RunCommand = True
    Exit Function
Case "clearvariable"
    lParms = Left(lParms, Len(lParms) - 1)
    lParms = Right(lParms, Len(lParms) - 1)
    i = FindVariableIndex(lParms)
    If i <> 0 Then
        If i = lVariables.vCount Then lVariables.vCount = lVariables.vCount - 1
        lVariables.vVariable(FindVariableIndex(lParms)).vData = ""
        lVariables.vVariable(FindVariableIndex(lParms)).vName = ""
    End If
    RunCommand = True
    Exit Function
Case "msgbox"
    MsgBox "Script: " & lParms
    RunCommand = True
    Exit Function
Case "disconnect"
    lForm.tcp.Close
    RunCommand = True
    Exit Function
Case "showquickconnect"
    frmQuickConnect.Show 0, mdiNexIRC
    RunCommand = True
    Exit Function
Case "showconnectionmanager"
    frmConnectionManager.Show
    RunCommand = True
    Exit Function
Case "editautojoin"
    frmAutoJoin.Show
    RunCommand = True
    Exit Function
Case "joinchannel"
    lForm.tcp.SendData "JOIN " & lParms & vbCrLf
    RunCommand = True
    Exit Function
Case "addautojoin"
    If LCase(lSPLT(0)) = "$activechannel" Then lSPLT(0) = mdiNexIRC.ActiveForm.Tag
    If LCase(lSPLT(1)) = "$activenetwork" Then lSPLT(1) = lSettings.sNetwork
    If Len(lSPLT(0)) <> 0 And Len(lSPLT(1)) <> 0 Then
        AddAutoJoin lSPLT(0), lSPLT(1)
    End If
    RunCommand = True
    Exit Function
Case "addchannelfolder"
    If LCase(lSPLT(0)) = "$activechannel" Then lSPLT(0) = mdiNexIRC.ActiveForm.Tag
    If Len(lSPLT(0)) <> 0 Then AddtoChanFolder lSPLT(0)
    RunCommand = True
    Exit Function
Case "resumeplayback"
    MenuPlay
    RunCommand = True
    Exit Function
Case "activateautojoin"
    ActivateAutoJoin False
    RunCommand = True
    Exit Function
Case "vbscript"
    'mdiNexIRC.ctlVBScript.ExecuteStatement lParms
    RunCommand = True
    Exit Function
Case "dccchat"
    frmDCC_Chat.Show 0, mdiNexIRC
    RunCommand = True
    Exit Function
Case "showadvancedsystemstats"
    frmStats.Show
Case "showsystemstats"
    Select Case lSpltCount
    Case 6
        ShowSystemStats mdiNexIRC.ActiveForm, CBool(lSPLT(0)), CBool(lSPLT(1)), CBool(lSPLT(2)), CBool(lSPLT(3)), CBool(lSPLT(4)), CBool(lSPLT(5)), CBool(lSPLT(6))
    Case 5
        ShowSystemStats mdiNexIRC.ActiveForm, CBool(lSPLT(0)), CBool(lSPLT(1)), CBool(lSPLT(2)), CBool(lSPLT(3)), CBool(lSPLT(4)), CBool(lSPLT(5)), False
    End Select
    RunCommand = True
    Exit Function
Case "runvbscriptfile"
    If DoesFileExist(lParms) = True Then
        LoadScript lParms, True
    Else
        LoadScript lParms, False
    End If
    RunCommand = True
    Exit Function
Case "web"
    Surf lParms, mdiNexIRC.hWnd
Case "runscriptfile"
    RunScriptFile App.Path & "\data\scripts\" & lParms
    RunCommand = True
    Exit Function
Case "noticeops"
    lSettings.sActiveServerForm.tcp.SendData "NOTICE " & lSPLT(0) & " :" & lSPLT(1) & vbCrLf
    RunCommand = True
    Exit Function
Case "notice"
    lSettings.sActiveServerForm.tcp.SendData "NOTICE " & lSPLT(0) & " :" & lSPLT(1) & vbCrLf
    RunCommand = True
    Exit Function
Case "sendactiveserver"
    lSettings.sActiveServerForm.tcp.SendData lParms & vbCrLf
    RunCommand = True
    Exit Function
Case "activateautoconnect"
    PerformAutoConnect
    RunCommand = True
    Exit Function
Case "sendquitmessage"
    If Len(lParms) <> 0 Then
        If LCase(lParms) <> "null" Then
            lSettings.sActiveServerForm.tcp.SendData "QUIT :" & lParms & vbCrLf
        Else
            SendQuitMessage lForm
        End If
    Else
        SendQuitMessage lForm
    End If
    RunCommand = True
    Exit Function
End Select
'DoColor lForm.txtIncoming, "" & Color.Mode & "• Unknown Command [" & "" & Color.CTCP & sCmd & "" & Color.Mode & "]"
ProcessReplaceString sUnknownCommand, lForm.txtIncoming, sCmd
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function RunCommand(ByVal sCmd As String, lForm As Form) As Boolean"
End Function

Public Function ShowMenu(ByVal X As Long, ByVal Y As Long, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lRet As Long, sCmdSplit() As String
Call BuildMenu(m_lMainHMENU, "ROOT")
lRet = TrackPopupMenu(m_lMainHMENU, TPM_LEFTALIGN Or TPM_RETURNCMD, ByVal X, ByVal Y, 0, m_lParentHWND, ByVal 0&)
lRet = FindItemByID(lRet)
If (lRet > 0) Then
    With m_tMenuItems(lRet)
        .sCommand = Replace(.sCommand, "$LF$", vbCrLf, 1, -1, vbTextCompare)
        .sCommand = Replace(.sCommand, "$DATE$", Date$, 1, -1, vbTextCompare)
        .sCommand = Replace(.sCommand, "$TIME$", Time$, 1, -1, vbTextCompare)
        Call RunCommand(.sCommand, lForm)
    End With
Else
    Exit Function
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ShowMenu(ByVal X As Long, ByVal Y As Long, lForm As Form)"
End Function
