Attribute VB_Name = "mdlTextStrings"
Option Explicit
Enum eStringTypes
    sNone = 0
    sNickCompletor1 = 1
    sNickCompletor2 = 2
    sQuitReason = 3
    sJoin = 4
    sPart = 5
    sKick = 6
    sQuit = 7
    sPm = 8
    sProgrammingError = 9
    sIdentConnection = 10
    sEnterDCCChat = 11
    sAttemptingToConnect = 12
    sIgnoreMessage = 13
    sFileOffer = 14
    sAutoJoin = 15
    sSaveLog = 16
    sUserInBlacklist = 17
    sScriptCleared = 18
    sVersion = 19
    sNowConnected = 20
    sConnectionTerminated = 21
    sFileNotFound = 22
    sIdentListening = 23
    sIdentListenFailed = 24
    sConnectionEstablished = 25
    sConnectionError = 26
    sIdentClosed = 27
    sIdentConnect = 28
    sIdentRequestDenied = 29
    sRequestChannelInformation = 30
    sUnknownCommand = 31
    sUserOnline = 32
    sRunAutoCommand = 33
    sPreformBotCommand = 34
    sNowTalkingIn = 35
    sAlreadyRegistered = 36
    sRaw = 37
    sOwnMessage = 38
    sInitiateDCCChat = 39
    sQuitMessage = 40
    sNoNicknameGiven = 41
    sNicknameInUse = 42
    sUserVoiced = 43
    sUserDevoiced = 44
    sUserOped = 45
    sUserDeoped = 46
    sConnectionClosed = 47
'    sSavedLogAs = 48
    sIdentDDisabled = 49
    sUndernetLogin = 50
    sAddToNotify = 51
    sVariableNotDeclared = 52
    sErrorFindingWindow = 53
    sAction = 54
End Enum
Private Type gString
    sData As String
    sType As eStringTypes
    sFind(5) As String
    sDescription As String
    sGroup As Integer
End Type
Private Type gStrings
    sCount As Integer
    sPresetIndex As Integer
    sString(150) As gString
End Type
Private lStrings As gStrings

Public Function ReturnTextPresetFileByDescription(lDescription As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, c As Integer, msg As String, msg2 As String
msg = App.Path & "\data\config\fixed\text.ini"
If DoesFileExist(msg) = True Then
    c = Int(Trim(ReadINI(msg, "Settings", "Count", 0)))
    If c <> 0 Then
        For i = 1 To c
            msg2 = ReadINI(msg, Trim(Str(i)), "Description", "")
            If Len(msg2) <> 0 Then
                If LCase(msg2) = LCase(lDescription) Then
                    ReturnTextPresetFileByDescription = ReadINI(msg, Trim(Str(i)), "File", "")
                    Exit For
                End If
            End If
        Next i
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnTextPresetByDescription(lDescription As String)"
End Function

Public Function ReturnTextPresetIndexByDescription(lDescription As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, c As Integer, msg As String, msg2 As String
msg = App.Path & "\data\config\fixed\text.ini"
If DoesFileExist(msg) = True Then
    c = Int(Trim(ReadINI(msg, "Settings", "Count", 0)))
    If c <> 0 Then
        For i = 1 To c
            msg2 = ReadINI(msg, Trim(Str(i)), "Description", "")
            If Len(msg2) <> 0 Then
                If LCase(msg2) = LCase(lDescription) Then
                    ReturnTextPresetIndexByDescription = i
                    Exit For
                End If
            End If
        Next i
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnTextPresetIndexByDescription(lDescription As String) As String"
End Function

Public Sub FillListBoxWithTextDescriptionsGroup(lListBox As ListBox, lGroup As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, b As Boolean
If lGroup = 1 Then b = True
lListBox.Clear
For i = 1 To 150
    If Len(lStrings.sString(i).sDescription) <> 0 Then
        If b = False Then
            If lGroup = lStrings.sString(i).sGroup Then
                lListBox.AddItem lStrings.sString(i).sDescription
            End If
        Else
            lListBox.AddItem lStrings.sString(i).sDescription
        End If
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub FillListBoxWithTextDescriptionsGroup(lListBox As ListBox, lGroup As Integer)"
End Sub

Public Sub FillComboWithStringGroups(lCombo As ComboBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, c As Integer, msg As String
lCombo.Clear
c = CInt(Trim(ReadINI(GetINIFile(iText), "Groups", "Count", 0)))
For i = 0 To c
    msg = ReadINI(GetINIFile(iText), "Groups", CInt(Trim(i)), "")
    If Len(msg) <> 0 Then lCombo.AddItem msg
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub FillComboWithStringGroups(lCombo As ComboBox)"
End Sub

Public Sub FillListboxWithTextDescriptions(lListBox As ListBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
lListBox.Clear
For i = 1 To 150
    If Len(lStrings.sString(i).sDescription) <> 0 Then lListBox.AddItem lStrings.sString(i).sDescription
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub FillListboxWithTextDescriptions(lListbox As ListBox)"
End Sub

Public Function FindStringIndexByDescription(lDescription As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lDescription) <> 0 Then
    For i = 1 To 150
        If LCase(lDescription) = LCase(lStrings.sString(i).sDescription) Then
            FindStringIndexByDescription = i
            Exit For
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindStringIndexByDescription(lDescription As String) As Integer"
End Function

Public Function ReturnStringTypeByDescription(lDescription As String) As eStringTypes
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lDescription) <> 0 Then
    For i = 1 To 150
        If LCase(lDescription) = LCase(lStrings.sString(i).sDescription) Then
            ReturnStringTypeByDescription = lStrings.sString(i).sType
            Exit For
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnStringTypeByDescription(lDescription As String) As eStringTypes"
End Function

Private Function FindStringIndex(lType As eStringTypes) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To 150
    If lType = lStrings.sString(i).sType Then
        FindStringIndex = i
        Exit Function
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Function FindStringIndex(lType As eStringTypes) As Integer"
End Function

Public Sub ProcessReplaceString(lType As eStringTypes, lTextBox As ctlTBox, Optional r1 As String, Optional r2 As String, Optional r3 As String, Optional r4 As String, Optional r5 As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = ReturnReplacedString(lType, r1, r2, r3, r4, r5)
If InStr(msg, "$color") Then msg = ReturnReplacedColors(msg)
If InStr(msg, "$bold") Or InStr(msg, "$bold_end") Then msg = ReturnReplacedBold(msg)
If Len(msg) <> 0 Then DoColor lTextBox, msg
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ProcessReplaceString(lType As eStringTypes, lTextBox As ctlTBox, Optional r1 As String, Optional r2 As String, Optional r3 As String, Optional r4 As String, Optional r5 As String)"
End Sub

Private Function ReturnReplacedBold(lData As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(lData) <> 0 And InStr(lData, "$bold") Or InStr(lData, "$bold_end") Then
    'hfuewofheowu
    lData = Replace(lData, "$bold_end", "")
    lData = Replace(lData, "$bold", "")
    ReturnReplacedBold = lData
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Function ReturnReplacedBold(lData As String) As String"
End Function

Private Function ReturnReplacedColors(lData As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(lData) <> 0 And InStr(lData, "$color") Then
    lData = Replace(lData, "$color_end", "")
    lData = Replace(lData, "$color_normal", "" & Color.Normal)
    lData = Replace(lData, "$color_action", "" & Color.Action)
    lData = Replace(lData, "$color_bgtext", "" & Color.BGText)
    lData = Replace(lData, "$color_ctcp", "" & Color.CTCP)
    lData = Replace(lData, "$color_invite", "" & Color.Invite)
    lData = Replace(lData, "$color_join", "" & Color.Join)
    lData = Replace(lData, "$color_kick", "" & Color.Kick)
    lData = Replace(lData, "$color_mode", "" & Color.Mode)
    lData = Replace(lData, "$color_nick", "" & Color.Nick)
    lData = Replace(lData, "$color_normal", "" & Color.Normal)
    lData = Replace(lData, "$color_notice", "" & Color.Notice)
    lData = Replace(lData, "$color_notify", "" & Color.Notify)
    lData = Replace(lData, "$color_part", "" & Color.Part)
    lData = Replace(lData, "$color_quit", "" & Color.Quit)
    lData = Replace(lData, "$color_server", "" & Color.Server)
    lData = Replace(lData, "$color_topic", "" & Color.Topic)
    lData = Replace(lData, "$color_whois", "" & Color.Whois)
    ReturnReplacedColors = lData
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Function ReturnReplacedColors(lData As String) As String"
End Function

Public Function ReturnReplacedString(lType As eStringTypes, Optional r1 As String, Optional r2 As String, Optional r3 As String, Optional r4 As String, Optional r5 As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer
i = FindStringIndex(lType)
msg = lStrings.sString(i).sData
With lStrings.sString(i)
    If Len(r1) <> 0 Then msg = Replace(msg, .sFind(1), r1, 1, -1, vbTextCompare)
    If Len(r2) <> 0 Then msg = Replace(msg, .sFind(2), r2, 1, -1, vbTextCompare)
    If Len(r3) <> 0 Then msg = Replace(msg, .sFind(3), r3, 1, -1, vbTextCompare)
    If Len(r4) <> 0 Then msg = Replace(msg, .sFind(4), r4, 1, -1, vbTextCompare)
    If Len(r5) <> 0 Then msg = Replace(msg, .sFind(5), r5, 1, -1, vbTextCompare)
End With
ReturnReplacedString = msg
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnReplacedString(lType As eStringTypes, Optional r1 As String, Optional r2 As String, Optional r3 As String, Optional r4 As String, Optional r5 As String) As String"
End Function

Public Sub SaveTextStrings()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To 150
    If Len(lStrings.sString(i).sData) <> 0 Then
        WriteINI GetINIFile(iText), Trim(Str(i)), "Type", Trim(Str(lStrings.sString(i).sType))
        WriteINI GetINIFile(iText), Trim(Str(i)), "Data", Trim(lStrings.sString(i).sData)
        If Len(lStrings.sString(i).sFind(1)) <> 0 Then WriteINI GetINIFile(iText), Trim(Str(i)), "Find1", lStrings.sString(i).sFind(1)
        If Len(lStrings.sString(i).sFind(2)) <> 0 Then WriteINI GetINIFile(iText), Trim(Str(i)), "Find2", lStrings.sString(i).sFind(2)
        If Len(lStrings.sString(i).sFind(3)) <> 0 Then WriteINI GetINIFile(iText), Trim(Str(i)), "Find3", lStrings.sString(i).sFind(3)
        If Len(lStrings.sString(i).sFind(4)) <> 0 Then WriteINI GetINIFile(iText), Trim(Str(i)), "Find4", lStrings.sString(i).sFind(4)
        If Len(lStrings.sString(i).sFind(5)) <> 0 Then WriteINI GetINIFile(iText), Trim(Str(i)), "Find5", lStrings.sString(i).sFind(5)
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SaveTextStrings()"
End Sub

Public Function ReturnStringDataByType(lType As eStringTypes) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To 150
    If lStrings.sString(i).sType = lType Then
        ReturnStringDataByType = lStrings.sString(i).sData
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnStringData(lIndex As Integer) As String"
End Function

Public Sub SetStringData(lType As eStringTypes, lData As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
i = FindStringIndex(lType)
If i <> 0 Then
    lStrings.sString(i).sData = lData
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetStringData(lType As eStringTypes)"
End Sub

Public Sub SelectCurrentPreset()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim c As Integer, i As Integer, msg As String, msg2 As String
msg = App.Path & "\data\config\fixed\text.ini"
c = Int(Trim(ReadINI(msg, "Settings", "Count", 0)))
i = Int(Trim(ReadINI(msg, "Settings", "Index", 0)))
If i = 0 Then
    If c <> 0 Then
        For i = 1 To c
            msg2 = App.Path & "\data\config\fixed\text\" & ReadINI(msg, Trim(Str(i)), "File", "")
            If Len(msg2) <> 0 Then
                If DoesFileExist(msg2) Then SetTextIniFile msg2
                Exit For
            End If
        Next i
    End If
Else
    msg2 = ReadINI(msg, Trim(Str(i)), "File", "")
    If Len(msg2) <> 0 Then
        If Len(msg2) <> 0 Then SetTextIniFile App.Path & "\data\config\fixed\text\" & msg2
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SelectCurrentPreset()"
End Sub

Public Sub ClearStrings()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, c As Integer
lStrings.sCount = 0
lStrings.sPresetIndex = 0
For i = 0 To 150
    lStrings.sString(i).sData = ""
    lStrings.sString(i).sDescription = ""
    For c = 0 To 5
        lStrings.sString(i).sFind(c) = ""
    Next c
    lStrings.sString(i).sGroup = 0
    lStrings.sString(i).sType = 0
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ClearStrings()"
End Sub

Public Sub LoadStrings(Optional lShowProgress As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim c As Integer, i As Integer, F As Integer, t As Integer, lForm As New frmImportAndExportProgress, msg As String
ClearStrings
msg = GetINIFile(iText)
SelectCurrentPreset
c = ReadINI(GetINIFile(iText), "Settings", "Count", 0)
If lShowProgress = True Then
    Set lForm = New frmImportAndExportProgress
    lForm.Show
    lForm.Caption = "NexIRC - Loading Strings"
    lForm.lblProgress.Caption = "Loading Strings, please wait..."
    lForm.XP_ProgressBar1.Max = c
    DoEvents
End If
If lShowProgress = True Then lForm.lblProgress.Caption = "Loading Settings..."

If c <> 0 Then
    For i = 1 To c
        With lStrings.sString(i)
            .sData = ReadINI(GetINIFile(iText), Trim(Str(i)), "Data", "")
            If Len(.sData) <> 0 Then
                If lShowProgress = True Then lForm.XP_ProgressBar1.Value = i
                .sDescription = ReadINI(GetINIFile(iText), Trim(Str(i)), "Description", "")
                t = t + 1
                .sType = ReadINI(GetINIFile(iText), Trim(Str(i)), "Type", 0)
                For F = 1 To 5
                    .sFind(F) = ReadINI(GetINIFile(iText), Trim(Str(i)), "Find" & Trim(Str(F)), "")
                Next F
                .sGroup = CInt(Trim(ReadINI(GetINIFile(iText), Trim(Str(i)), "Group", 0)))
            End If
        End With
    Next i
    If lShowProgress = True Then Unload lForm
    lStrings.sCount = t
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub LoadStrings()"
End Sub
