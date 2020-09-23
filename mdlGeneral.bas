Attribute VB_Name = "mdlGeneral"
Option Explicit
Public Const lChannelUBound = 50
Public lChannel(1 To lChannelUBound) As New frmChannel
Public lChannelName(1 To lChannelUBound) As String
Public lChannelTopic(1 To lChannelUBound) As String
Public lChannelModes(1 To lChannelUBound) As String
Public lChannelLimit(1 To lChannelUBound) As String
Public lQuery(1 To 150) As New frmQuery
Public lQueryName(1 To 150) As String
Public ACTION_lChannel As String
Public Const maxtcp = 10
Public connected As Boolean
Public CHAT_Index As Long
Public lChatWindow(1 To maxtcp) As New frmChat
Public lChatWindowName(1 To maxtcp) As String
Public lChatWindowx(1 To maxtcp) As New frmChat
Public lChatWindowNamex(1 To maxtcp) As String
Public Notify(1 To 150) As String
Public NOTIFYLIST As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public PingReply As Long
Public Type numericcode
    Num As String
    Server As String
    NickName As String
    parms As String
    ServerText As String
End Type
Public Type commandtrigger
    parms As String
    UserName As String
    Target As String
    Command As String
    JoinChannel As String
    ChanPart As String
    NickJoin As String
    NickPart As String
End Type
Public Type Channelstats
    Topic As String
    Name As String
End Type
Public chanstats(1 To lChannelUBound) As Channelstats
Public lEvents As commandtrigger
Public lRaw As numericcode
Public FileIndex As Integer
Public FileListenPort As Integer

Public Sub AddTaskbar(lCaption As String, lPicType As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If (mdiMain.StatusBar.Panels.Count + 1) = 17 Then Exit Sub
If FindPanelIndex(lCaption, mdiMain.StatusBar) = 0 Then
    mdiMain.StatusBar.Panels.Add (mdiMain.StatusBar.Panels.Count + 1), lCaption, lCaption
End If
If Err.Number = 35602 Then Exit Sub
If lSettings.sAutosizeStatusbarItems = True Then
    mdiMain.StatusBar.Panels.Item((mdiMain.StatusBar.Panels.Count)).AutoSize = sbrSpring
Else
    mdiMain.StatusBar.Panels.Item((mdiMain.StatusBar.Panels.Count)).AutoSize = sbrNoAutoSize
End If
mdiMain.StatusBar.Panels.Item((mdiMain.StatusBar.Panels.Count)).Style = sbrText
mdiMain.StatusBar.Panels.Item((mdiMain.StatusBar.Panels.Count)).Bevel = sbrRaised
Select Case lPicType
Case 1
    mdiMain.StatusBar.Panels.Item((mdiMain.StatusBar.Panels.Count)).Picture = mdiMain.imgTaskbar.ListImages(1).Picture
Case 2
    mdiMain.StatusBar.Panels.Item((mdiMain.StatusBar.Panels.Count)).Picture = mdiMain.imgTaskbar.ListImages(2).Picture
End Select
End Sub

Public Sub RemoveTaskbar(lCaption As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To mdiMain.StatusBar.Panels.Count
    If LCase(mdiMain.StatusBar.Panels.Item(i).Key) = LCase(lCaption) Then
        mdiMain.StatusBar.Panels.Remove i
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub RemoveTaskbar(lCaption As String)"
End Sub

Public Sub ShowStats(lTBox As TBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lOperators As Integer, lVoiced As Integer, lUsers As Integer, i As Integer, j As Integer
For i = 1 To lChannelUBound
    If LCase(ACTION_lChannel) = LCase(lChannelName(i)) Then
        For j = 1 To lChannel(i).lstNames.ListItems.Count - 1
            Select Case Left(lChannel(i).lstNames.ListItems(j).ForeColor, 1)
            Case lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nOpColor
                lOperators = lOperators + 1
            Case lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nVoiceColor
                lVoiced = lVoiced + 1
            Case Else
                lUsers = lUsers + 1
            End Select
        Next j
        DoColor lChannel(i).txtIncoming, "" & Color.Join & "• lOperators: " & lOperators & " lVoiced: " & lVoiced & " lUsers: " & lUsers & " - Total: " & Trim(Str(lOperators + lVoiced + lUsers))
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ShowStats(RTF as tBox)"
End Sub

Public Sub ClearVariables()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lRaw.NickName = ""
lRaw.Num = ""
lRaw.parms = ""
lRaw.Server = ""
lRaw.ServerText = ""
lEvents.UserName = ""
lEvents.Target = ""
lEvents.parms = ""
lEvents.NickPart = ""
lEvents.NickJoin = ""
lEvents.Command = ""
lEvents.ChanPart = ""
lEvents.JoinChannel = ""
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ClearVariables()"
End Sub

Public Sub CheckWord(lWord As String, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, j As Integer, y As Integer, lWordArr() As String, lErr As String, lParams As String
lRaw.ServerText = lWord
If lWord = "" Then Exit Sub
lErr = lWord
lWordArr = Split(lWord, Chr(32))
For i = 3 To UBound(lWordArr)
    If Trim(lWordArr(i)) <> "" Then
        lParams = lParams & " " & lWordArr(i)
    End If
Next i
If IsNumeric(lWordArr(1)) = True Then
    lSettings.sNickname = lWordArr(2)
    Call numeric(lWordArr(0), lWordArr(1), lWordArr(2), Mid(lParams, 2), lForm)
    Exit Sub
End If
If lWordArr(0) = "PING" Then
    lForm.tcp.SendData "PONG " & Mid(lWordArr(1), 2) & vbCrLf
    lForm.Caption = lForm.Tag & " " & lSettings.sNickname & " on " & lSettings.sServer & ":" & lForm.tcp.RemotePort
    Exit Sub
End If
Call Command(lWordArr(0), lWordArr(1), lWordArr(2), Mid(lParams, 2), lForm)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub CheckWord(lWord As String, lForm As Form)"
End Sub

Public Sub IsUserOnline(lTBox As TBox, lNickname As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, j As Integer
For i = 1 To lChannelUBound
    If lChannelName(i) <> "" Then
        For j = 1 To lChannel(i).lstNames.ListItems.Count - 1
            If LCase(lChannel(i).lstNames.ListItems(j).Text) = LCase(lNickname) Or LCase(lChannel(i).lstNames.ListItems(j).Text) = LCase("@" & lNickname) Or LCase(lChannel(i).lstNames.ListItems(j).Text) = LCase("+" & lNickname) Then
                DoColor lTBox, "" & Color.Notice & "* " & lNickname & " is on " & lChannelName(i)
                Exit For
            End If
        Next j
    End If
Next i
End Sub

Public Sub UpdateCaption(lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lChannelLimit(lIndex) = "" Then
    lChannel(lIndex).Caption = lChannelName(lIndex) & " [" & lChannel(lIndex).lstNames.ListItems.Count & "] [+" & lChannelModes(lIndex) & "] :" & lChannelTopic(lIndex)
Else
    lChannel(lIndex).Caption = lChannelName(lIndex) & " [" & lChannel(lIndex).lstNames.ListItems.Count & "] [+" & lChannelModes(lIndex) & " " & lChannelLimit(lIndex) & "] :" & lChannelTopic(lIndex)
End If
End Sub
