Attribute VB_Name = "mdlFunctions"
Option Explicit
'Private Const LF_FACESIZE = 32
Private Type gLogFont
    lHeight As Long
    lWidth As Long
    lEscapement As Long
    lOrientation As Long
    lWeight As Long
    lItalic As Byte
    lUnderline As Byte
    lStrikeOut As Byte
    lCharSet As Byte
    lOutPrecision As Byte
    lClipPrecision As Byte
    lQuality As Byte
    lPitchAndFamily As Byte
    lFaceName(32) As Byte
End Type
Private Type gNewTextMetric
    nHeight As Long
    nAscent As Long
    nDescent As Long
    nInternalLeading As Long
    nExternalLeading As Long
    nAveCharWidth As Long
    nMaxCharWidth As Long
    nWeight As Long
    nOverhang As Long
    nDigitizedAspectX As Long
    nDigitizedAspectY As Long
    nFirstChar As Byte
    nLastChar As Byte
    nDefaultChar As Byte
    nBreakChar As Byte
    nItalic As Byte
    nUnderlined As Byte
    nStruckOut As Byte
    nPitchAndFamily As Byte
    nCharSet As Byte
    nFlags As Long
    nSizeEM As Long
    nCellHeight As Long
    nAveWidth As Long
End Type
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As SpecialFolderIDs, ByRef pIdl As Long) As Long
Private Declare Function SHGetPathFromIDListA Lib "shell32" (ByVal pIdl As Long, ByVal pszPath As String) As Long
Private Enum SpecialFolderIDs
    sfidPROGRAMS = &H2
End Enum

Public Function GetMyDocumentsDir() As String
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
Dim msg As String, l As Long, o As Long
If SHGetSpecialFolderLocation(0, &H2, l) = 0 Then
    msg = String$(255, 0)
    SHGetPathFromIDListA l, msg
    o = InStr(msg, Chr(0))
    If o > 0 Then msg = Left$(msg, o - 1)
End If
GetMyDocumentsDir = Left(msg, Len(msg) - 19)
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Function GetMyDocumentsDir() As String"
    Err.Clear
End Function

Public Function CheckListbox(strListBox As ListBox, CheckName As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
CheckListbox = False
For i = 0 To strListBox.ListCount - 1
    If strListBox.List(i) = CheckName Then
        CheckListbox = True
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function CheckListbox(strListBox As ListBox, CheckName As String) As Boolean"
End Function

Public Function KeyGen(kName As String, kPass As String, kType As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim cTable(512) As Integer, nKeys(16) As Integer, s0(512) As Integer, nArray(16) As Integer, pArray(16) As Integer, n As Integer, nPtr As Integer, cPtr As Integer, cFlip As Boolean, sIni As Integer, temp As Integer, rtn As Integer, gKey As String, nLen As Integer, pLen As Integer, kPtr As Integer, sPtr As Integer, nOffset As Integer, pOffset As Integer, tOffset As Integer
Const nXor As Integer = 18
Const pXor As Integer = 25
Const cLw As Integer = 65
Const nLw As Integer = 48
Const sOffset As Integer = 0
nLen = Len(kName)
pLen = Len(kPass)
nKeys(1) = 52
nKeys(2) = 69
nKeys(3) = 149
nKeys(4) = 37
nKeys(5) = 403
nKeys(6) = 20
nKeys(7) = 58
nKeys(8) = 29
nKeys(9) = 123
nKeys(10) = 84
nKeys(11) = 201
nKeys(12) = 202
nKeys(13) = 34
nKeys(14) = 38
nKeys(15) = 73
nKeys(16) = 30
sIni = 0
For n = 0 To 512
    s0(n) = n
Next n
For n = 0 To 512
    sIni = (sOffset + sIni + n) Mod 256
    temp = s0(n)
    s0(n) = s0(sIni)
    s0(sIni) = temp
Next n
If kType = 1 Then
    nPtr = 0
    For n = 0 To 512
        cTable(s0(n)) = (nLw + (nPtr))
        nPtr = nPtr + 1
        If nPtr = 10 Then nPtr = 0
    Next n
    gKey = String(16, " ")
ElseIf kType = 2 Then
    nPtr = 0
    cPtr = 0
    cFlip = False
    For n = 0 To 512
        If cFlip Then
            cTable(s0(n)) = (nLw + nPtr)
            nPtr = nPtr + 1
            If nPtr = 10 Then nPtr = 0
            cFlip = False
        Else
            cTable(s0(n)) = (cLw + cPtr)
            cPtr = cPtr + 1
            If cPtr = 26 Then cPtr = 0
            cFlip = True
        End If
    Next n
    gKey = String(16, " ")
Else
    gKey = String(19, " ")
End If
kPtr = 1
For n = 1 To nLen
    nArray(kPtr) = nArray(kPtr) + Asc(Mid(kName, n, 1)) Xor nXor
    nOffset = nOffset + nArray(kPtr)
    kPtr = kPtr + 1
    If kPtr = 9 Then kPtr = 1
Next n
For n = 1 To pLen
    pArray(kPtr) = pArray(kPtr) + Asc(Mid(kPass, n, 1)) Xor pXor
    pOffset = pOffset + pArray(kPtr)
    kPtr = kPtr + 1
    If kPtr = 9 Then kPtr = 1
Next n
tOffset = (nOffset + pOffset) Mod 512
kPtr = 1
sPtr = 1
For n = 1 To 16
    pArray(n) = pArray(n) Xor nKeys(n)
    rtn = Abs(((nArray(n) Xor pArray(n)) Mod 512) - tOffset)
    If kType = 3 Then
        If rtn < 16 Then
            Mid(gKey, kPtr, 2) = "0" & Hex(rtn)
        Else
            Mid(gKey, kPtr, 2) = Hex(rtn)
        End If
        If sPtr = 2 And kPtr < 18 Then
            kPtr = kPtr + 1
            Mid(gKey, kPtr + 1, 1) = "-"
        End If
        kPtr = kPtr + 2
        sPtr = sPtr + 1
        If sPtr = 3 Then sPtr = 1
    Else
        Mid(gKey, n, 1) = Chr(cTable(rtn))
    End If
Next n
KeyGen = gKey
End Function

Public Function EnumFontFamTypeProc(lpNLF As gLogFont, lpNTM As gNewTextMetric, ByVal FontType As Long, lParam As ListBox) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim FaceName As String
If ShowFontType = FontType Then
    FaceName = StrConv(lpNLF.lFaceName, vbUnicode)
    lParam.AddItem Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
End If
EnumFontFamTypeProc = 1
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function EnumFontFamTypeProc(lpNLF As gLogFont, lpNTM As gNewTextMetric, ByVal FontType As Long, lParam As ListBox) As Long"
End Function

Public Function FindImageComboIndex(lImageCombo As ImageCombo, lText As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lText) <> 0 Then
    For i = 0 To lImageCombo.ComboItems.Count
        If LCase(lImageCombo.ComboItems(i).Text) = LCase(lText) Then
            FindImageComboIndex = i
            Exit For
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindImageComboIndex(lImageCombo As ImageCombo, lText As String) As Integer"
End Function

Public Function FindListViewIndex(lListView As ListView, lText As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lText) <> 0 Then
    If lListView.ListItems.Count <> 0 Then
        For i = 1 To lListView.ListItems.Count
            If Trim(LCase(lListView.ListItems(i).Text)) = Trim(LCase(lText)) Then
                FindListViewIndex = i
                Exit For
            End If
            If Trim(LCase(lListView.ListItems(i).Text)) = "@" & Trim(LCase(lText)) Then
                FindListViewIndex = i
                Exit For
            End If
            If Trim(LCase("@" & lListView.ListItems(i).Text)) = Trim(LCase(lText)) Then
                FindListViewIndex = i
                Exit For
            End If
            If Trim(LCase(lListView.ListItems(i).Text)) = "+" & Trim(LCase(lText)) Then
                FindListViewIndex = i
                Exit For
            End If
            If Trim(LCase("+" & lListView.ListItems(i).Text)) = Trim(LCase(lText)) Then
                FindListViewIndex = i
                Exit For
            End If
        Next i
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindListViewIndex(lListView As ListView, lText As String) As Integer"
End Function

Public Function ProcessKeyDown(lKeyCode As Integer, lShift As Integer, lText As String, lListBox As ListBox, lOutgoing As TextBox) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case lKeyCode
Case 13
    lListBox.AddItem lText
    lListBox.Visible = False
    ProcessKeyDown = True
Case 27
    lListBox.Visible = False
    ProcessKeyDown = True
Case 9
    lListBox.Visible = True
    If lListBox.ListCount <> 0 Then
        If lListBox.ListIndex <> 0 And lListBox.ListIndex <> -1 Then
            lListBox.ListIndex = lListBox.ListIndex - 1
        Else
            lListBox.ListIndex = lListBox.ListCount - 1
        End If
        If Len(lListBox.Text) <> 0 Then
            lOutgoing.Text = lListBox.Text
        End If
    End If
    lOutgoing.SelStart = 0
    lOutgoing.SelLength = Len(lOutgoing.Text)
    ProcessKeyDown = True
Case 38
    lListBox.Visible = False
    If lListBox.ListCount <> 0 Then
        If lListBox.ListIndex <> 0 And lListBox.ListIndex <> -1 Then
            lListBox.ListIndex = lListBox.ListIndex - 1
        Else
            lListBox.ListIndex = lListBox.ListCount - 1
        End If
        If Len(lListBox.Text) <> 0 Then
            lOutgoing.Text = lListBox.Text
        End If
    End If
    lOutgoing.SelStart = 0
    lOutgoing.SelLength = Len(lOutgoing.Text)
    ProcessKeyDown = True
Case 40
    lListBox.Visible = False
    If lListBox.ListCount <> 0 Then
        If (lListBox.ListIndex + 1) <> lListBox.ListCount Then
            lListBox.ListIndex = (lListBox.ListIndex + 1)
        Else
            lListBox.ListIndex = 0
        End If
        If Len(lListBox.Text) <> 0 Then
            lOutgoing.Text = lListBox.Text
        End If
    End If
    lOutgoing.SelStart = 0
    lOutgoing.SelLength = Len(lOutgoing.Text)
    ProcessKeyDown = True
Case Else
    If lListBox.Visible = True Then lListBox.Visible = False
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ProcessKeyDown(lKeyCode As Integer, lShift As Integer, lText As String, lListBox As ListBox, lOutgoing As TextBox) As Boolean"
End Function

Public Function GetPathFromMedia(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String
Start:
i = GetRnd(lFiles.fCount)
If Len(lFiles.fFile(i).fFilename) <> 0 And i <> 0 And DoesFileExist(lFiles.fFile(i).fFilename) = True Then
    msg = lFiles.fFile(i).fFilename
    msg = GetFileTitle(msg)
    If Len(msg) <> 0 Then
        msg = Left(lFiles.fFile(i).fFilename, Len(lFiles.fFile(i).fFilename) - Len(msg))
        GetPathFromMedia = msg
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function GetDirFromFilePath(lFile As String) As String"
End Function

Public Function FindComboBoxIndex(lCombo As ComboBox, lText As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lText) <> 0 Then
    For i = 0 To lCombo.ListCount
        If Trim(LCase(lCombo.List(i))) = Trim(LCase(lText)) Then
            FindComboBoxIndex = i
            Exit For
            Exit Function
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindComboBoxIndex(lCombo As ComboBox, lText As String) As Integer"
End Function

Public Function GetFileTitle(lFileName As String) As String
On Local Error Resume Next
Dim msg() As String
If Len(lFileName) <> 0 Then
    msg = Split(lFileName, "\", -1, vbTextCompare)
    GetFileTitle = msg(UBound(msg))
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function GetFileTitle(lFilename As String) As String"
End Function

Public Function IsMP3(lFileName As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If LCase(Right(lFileName, 4)) = ".mp3" Then
    IsMP3 = True
Else
    IsMP3 = False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function IsMP3(lFilename As String) As Boolean"
End Function

Public Function AddSingleFileToPlaylist(lFileName As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lFile As String, lPath As String, i As Integer
If Len(lFileName) <> 0 Then
    If LCase(Right(lFileName, 4)) <> ".mp3" Then
        If lSettings.sExlusiveToMp3InPlaylist = True Then
            Exit Function
        End If
    End If
    lFile = lFileName
    lFile = GetFileTitle(lFileName)
    lPath = Left(lFileName, Len(lFileName) - Len(lFile))
    i = FindFileIndexByFilename(lFile)
    If i = 0 Then
        lFiles.fCount = lFiles.fCount + 1
        lFiles.fFile(lFiles.fCount).fFilename = lFileName
        If lSettings.sPlaylistVisible = True Then frmPlaylist.lstPlaylist.AddItem lFile
        SavePlaylist
        AddSingleFileToPlaylist = lFiles.fCount
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function AddSingleFileToPlaylist(lFilename As String) As Integer"
End Function

Public Function OpenScript(lFileName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim F As Form, msg As String
Set F = New frmTextEditor
lSettings.sTextCount = lSettings.sTextCount + 1
F.Caption = lFileName
msg = ReadFile(lFileName)
F.txtIncoming.Text = msg
F.Show
OpenScript = lSettings.sTextCount
AddTaskPanel lFileName, 2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function OpenScript(lFilename As String)"
End Function

Public Function PromptOpenScriptFile(lForm As Form) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = OpenDialog(lForm, "Text Files (*.txt)|*.txt", "NexIRC", App.Path & "\data\scripts\nexirc")
If Len(msg) <> 0 Then
    If DoesFileExist(msg) = True Then
        PromptOpenScriptFile = msg
    Else
        If lSettings.sGeneralPrompts = True Then
            MsgBox "Unable to open " & msg & ".", vbExclamation
        End If
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function PromptOpenScriptFile(lForm As Form) As String"
End Function

Public Function ConnectToIRC(lServer As String, lPort As String, lForm As Form) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim APort As String, i As Integer, mbox As VbMsgBoxResult
If Len(lSettings.sNickname) = 0 Or Len(lSettings.sEMail) = 0 Then
    Beep
    frmCustomize.Show
    DoEvents
    frmCustomize.optCheck(1).Value = True
    frmCustomize.fraSettings(0).Visible = False
    frmCustomize.fraSettings(1).Visible = True
    If Len(lSettings.sNickname) = 0 Then
        frmCustomize.txtNickname.SetFocus
        'frmCustomize.CreateBalloon "No nickname given", "You must type a Nickname below before you can connect to IRC", frmCustomize.txtNickname
        Exit Function
    End If
    If Len(lSettings.sEMail) = 0 Then
        frmCustomize.txtEmail.SetFocus
        frmCustomize.CreateBalloon "E-Mail Address not given", "The IRC server must know what your email address is to connect. Enter it below", frmCustomize.txtEmail
        Exit Function
    End If
    If Len(lSettings.sRealName) = 0 Then
        frmCustomize.txtRealName.SetFocus
        'frmCustomize.CreateBalloon "Real name", "The IRC server must know what your real name is to connect. " & vbCrLf & " Please enter it below", frmCustomize.txtRealName
        Exit Function
    End If
End If
If Len(lServer) = 0 Or Len(lPort) = 0 Or lPort = "0" Then
    If Len(lServer) = 0 Then
        lServer = InputBox("Enter server address:", "NexIRC", "")
        If Len(lServer) = 0 Then
            If lSettings.sGeneralPrompts = True Then
                MsgBox "Unable to connect to server", vbExclamation
                Exit Function
            End If
        End If
    End If
    If Len(lPort) = 0 Or lPort = "0" Then
        lPort = InputBox("Enter port:", "NexIRC", "6667")
        If Len(lPort) = 0 Then
            If lSettings.sGeneralPrompts = True Then
                MsgBox "Unable to connect to server", vbExclamation
                Exit Function
            End If
        End If
    End If
End If
If lSettings.sIdent.iEnabled = True Then EnableIdent lForm
If lForm.tcp.State <> 0 Then
    lForm.tcp.Close
End If
If Err.Number = 91 Then
    Err.Clear
    If lSettings.sActiveServerForm.tcp.State <> 0 Then lSettings.sActiveServerForm.tcp.Close
    Set lForm = lSettings.sActiveServerForm
    If Err.Number = 91 Then
        'If lSettings.sGeneralPrompts = True Then
        '    mbox = MsgBox("You have no status windows open. Would you like to open a new status window now?", vbYesNo + vbQuestion)
        '    If mbox = vbYes Then
        '        NewStatusWindow lServer, lPort, True
        '        Exit Function
        '    Else
        '        mdiNexIRC.picConnect.Picture = frmGraphics.picConnect1.Picture
        '        Exit Function
        '    End If
        'Else
        '    NewStatusWindow lServer, lPort, True
        '    Exit Function
        'End If
    End If
End If
APort = Val(Mid(lPort, i + 1))
ProcessReplaceString sAttemptingToConnect, lForm.txtIncoming, lServer, Trim(Str(APort))
If Left(lServer, 1) = ":" Then
    lServer = Right(lServer, Len(lServer) - 1)
    lSettings.sServer = lServer
End If
lForm.tcp.Connect lServer, APort
If Err.Number <> 0 Then
    ProcessRuntimeError Err.Description, Err.Number, "Public Function ConnectToIRC(lServer As String, lPort As String, lForm As Form) As Boolean"
End If
End Function

Public Function FindListBoxIndex(lText As String, lListBox As ListBox) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To lListBox.ListCount
    If LCase(lText) = LCase(lListBox.List(i)) Then
        FindListBoxIndex = i
        Exit Function
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindListBoxIndex(lText As String, lListBox As ListBox) As Integer"
End Function

Public Function NewScriptFile() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim F As Form
Set F = New frmTextEditor
lSettings.sTextCount = lSettings.sTextCount + 1
F.Caption = "Untitled " & lSettings.sTextCount & ".txt"
F.Show
NewScriptFile = lSettings.sTextCount
AddTaskPanel F.Caption, 2
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function NewScriptFile() As Integer"
End Function

Public Function OpenTextFile(lFileName As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim F As Form, fTitle As String
Set F = New frmTextEditor
lSettings.sTextCount = lSettings.sTextCount + 1
fTitle = lFileName
fTitle = GetFileTitle(fTitle)
F.Caption = fTitle
F.Show
OpenTextFile = lSettings.sTextCount
AddTaskPanel F.Caption, 2
F.txtIncoming.Text = ReadFile(lFileName): DoEvents
F.Tag = ""
F.txtIncoming.Tag = lFileName
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function OpenTextFile(lFilename As String) As Integer"
End Function

Public Function ApplyImageToPictureBox(lImageLocation As String, lPic As StdPicture) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String
If Len(lImageLocation) <> 0 Then
    lImageLocation = Trim(lImageLocation)
    msg = lImageLocation
    msg = GetFileTitle(msg)
    If Len(msg) <> 0 Then
        If DoesFileExist(lImageLocation) = True Then
            Set lPic = LoadPicture(lImageLocation)
            ApplyImageToPictureBox = True
        End If
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ApplyImageToPictureBox(lImageLocation As String, lPictureBox As PictureBox) As Boolean"
End Function

Public Function GetRnd(Num As Long) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Randomize Timer
GetRnd = Int((Num * Rnd) + 1)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function GetRnd(Num As Long) As Long"
End Function

Public Sub FormDrag(lFormname As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReleaseCapture
Call SendMessage(lFormname.hWnd, &HA1, 2, 0&)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub FormDrag(lFormname As Form)"
End Sub

Public Function SaveFile(lFileName As String, lText As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(lFileName) <> 0 And Len(lText) <> 0 Then
    Open lFileName For Output As #13
    Print #13, lText
    Close #13
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function SaveFile(lFilename As String, lText As String) As Boolean"
End Function

Public Function FindPanelIndex(lPanel As String, lStatusBar As StatusBar) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lPanel) <> 0 Then
    For i = 1 To lStatusBar.Panels.Count
        If Trim(LCase(lStatusBar.Panels(i).Text)) = Trim(LCase(lPanel)) Then
            FindPanelIndex = i
            Exit For
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindPanelIndex(lPanel As String, lStatusBar As StatusBar) As Integer"
End Function

Public Function DoesPanelExistInStatusBar(lPanel As String, lStatusBar As StatusBar) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lPanel) <> 0 Then
    For i = 1 To lStatusBar.Panels.Count
        If Trim(LCase(lStatusBar.Panels(i).Text)) = Trim(LCase(lPanel)) Then
            DoesPanelExistInStatusBar = True
            Exit For
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function DoesPanelExistInStatusBar(lPanel As String, lStatusBar As StatusBar) As Boolean"
End Function

Public Function ReadFile(lFile As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim n As Integer, msg As String
n = FreeFile
If DoesFileExist(lFile) = True Then
    Open lFile For Input As #n
        msg = StrConv(InputB(LOF(n), n), vbUnicode)
        If Len(msg) <> 0 Then
            ReadFile = Left(msg, Len(msg) - 2)
        End If
    Close #n
Else
    ProcessReplaceString sProgrammingError, lSettings.sActiveServerForm.txtIncoming, "File: " & lFile & " not found!", "404"
    If lSettings.sGeneralPrompts = True Then
        MsgBox lFile & " does not exist!", vbExclamation
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReadFile(lFile As String) As String"
End Function

Public Function DoesFileExist(lFileName As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = Dir(lFileName)
If msg <> "" Then
    DoesFileExist = True
Else
    DoesFileExist = False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function DoesFileExist(lFilename As String) As Boolean"
End Function

Public Function Parse(lWhole As String, lStart As String, lEnd As String)
On Local Error GoTo ErrHandler
Dim len1 As Integer, len2 As Integer, Str1 As String, Str2 As String
len1 = InStr(lWhole, lStart)
len2 = InStr(lWhole, lEnd)
Str1 = Right(lWhole, Len(lWhole) - len1)
Str2 = Right(lWhole, Len(lWhole) - len2)
Parse = Left(Str1, Len(Str1) - Len(Str2) - 1)
ErrHandler:
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function Parse(lWhole As String, lStart As String, lEnd As String)"
End Function

Public Function FindNetworkIndex(lNetwork As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To lServers.sNetworkCount
    If lServers.sNetwork(i).nDescription = lNetwork Then
        FindNetworkIndex = i
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindNetworkIndex(lNetwork As String)"
End Function

Public Function FindServerIndex(lServerName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To lServers.sServerCount
    If LCase(lServers.sServer(i).sServer) = LCase(lServerName) Then
        FindServerIndex = i
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindNetworkIndex(lNetwork As String)"
End Function

Public Function SetCheckBoxValueInt(lCheckbox As CheckBox, lValue As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lValue = 0 Then
    lCheckbox.Value = 1
Else
    lCheckbox.Value = 0
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function SetCheckBoxValueInt(lCheckbox As CheckBox, lValue As Integer)"
End Function

Public Function SetXPButtonValue(lButton As ctlXPButton, lValue As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lValue = True Then
    lButton.Value = True
Else
    lButton.Value = False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function SetXPButtonValue(lButton As ctlXPButton, lValue As Boolean)"
End Function

Public Function SetCheckBoxValue(lCheckbox As CheckBox, lValue As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lValue = True Then
    lCheckbox.Value = 1
Else
    lCheckbox.Value = 0
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function SetCheckBoxValue(lCheckbox As CheckBox, lValue As Boolean)"
End Function

Public Function GetCheckboxValue(lCheckbox As CheckBox) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lCheckbox.Value = 1 Then
    GetCheckboxValue = True
Else
    GetCheckboxValue = False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function GetCheckboxValue(lCheckbox As CheckBox) As Boolean"
End Function
