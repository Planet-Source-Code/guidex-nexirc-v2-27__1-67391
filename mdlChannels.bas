Attribute VB_Name = "mdlChannels"
Option Explicit
Private Const lChannelUBound = 64
Private Type gStats
    cTopic As String
    cName As String
'    cActionChannel As String
End Type
Private Type gChannels
    cChannel(1 To lChannelUBound) As New frmChannel
    cName(1 To lChannelUBound) As String
    cTopic(1 To lChannelUBound) As String
    cModes(1 To lChannelUBound) As String
    cLimit(1 To lChannelUBound) As String
    cStats(1 To lChannelUBound) As gStats
End Type
Private lChannels As gChannels
Private lActChannel As String

Public Sub DoColorChannel(lChannelIndex As Integer, lData As String)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
DoColor lChannels.cChannel(lChannelIndex).txtIncoming, lData
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Sub DoColorChannel(lChannelIndex As Integer, lData As String)"
    Err.Clear
End Sub

Public Sub SetActChannel(lData As String)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
lActChannel = lData
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnActChannel() As String"
    Err.Clear
End Sub

Public Function ReturnActChannel() As String
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
ReturnActChannel = lActChannel
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnActChannel() As String"
    Err.Clear
End Function

Public Function FindChannelIndexByNickName(lNickName As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
Dim i As Integer, n As Integer
If Len(lNickName) <> 0 Then
    For i = 0 To ReturnChannelUBound
        If Len(lChannels.cName(i)) <> 0 Then
            For n = 0 To ReturnChannelNamesCount(i)
                If Trim(LCase(lNickName)) = lChannels.cName(i) Then
                    FindChannelIndexByNickName = i
                    Exit For
                End If
            Next n
        End If
    Next i
End If
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Function FindChannelIndexByNickName(lNickName As String) As Integer"
    Err.Clear
End Function

Public Function ReturnChannelVisible(lIndex As Integer, lNamesIndex As Integer) As Boolean
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
ReturnChannelVisible = lChannels.cChannel(lIndex).Visible
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnChannelVisible(lIndex As Integer) As Boolean"
    Err.Clear
End Function

Public Sub SetChannelStatsTopic(lIndex As Integer, lData As String)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
lChannels.cTopic(lIndex) = lData
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetChannelTopic(lIndex As Integer, lData As String)"
    Err.Clear
End Sub

Public Sub SetChannelLimit(lIndex As Integer, lData As String)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
lChannels.cLimit(lIndex) = lData
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetChannelLimit(lIndex As Integer, lData As String)"
    Err.Clear
End Sub

Public Function ReturnChannelListItemName(lIndex As Integer, lNameIndex) As Integer
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
ReturnChannelListItemName = lChannels.cChannel(lIndex).lvwNames.ItemText(lNameIndex)
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnChannelName(lIndex As Integer, lNameIndex) As Integer"
    Err.Clear
End Function

Public Function ReturnChannelNamesSelected(lIndex As Integer) As Integer
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
ReturnChannelNamesSelected = lChannels.cChannel(lIndex).ReturnSelectedItem
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnChannelNamesSelected(lIndex As Integer)"
    Err.Clear
End Function

Public Sub SetChannelTopic(lIndex As Integer, lData As String)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
lChannels.cTopic(lIndex) = lData
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetChannelTopic(lIndex As Integer, lData As String)"
    Err.Clear
End Sub

Public Function ReturnChannelTopicTextBox(lIndex As Integer) As TextBox
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
ReturnChannelTopicTextBox = lChannels.cChannel(lIndex).txtTopic
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnChannelTopicTextBox(lIndex As Integer) As TextBox"
    Err.Clear
End Function

Public Sub SetChannelTopicToolTip(lIndex As Integer, lData As String)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
lChannels.cChannel(lIndex).txtTopic.ToolTipText = lData
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetChannelTopicToolTip(lIndex As Integer, lData As String)"
End Sub

Public Sub SetChannelTag(lIndex As Integer, lData As String)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
lChannels.cChannel(lIndex).Tag = lData
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetChannelTag(lIndex As Integer, lData As String)"
End Sub

Public Sub SetChannelCaption(lIndex As Integer, lData As String)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
lChannels.cChannel(lIndex).Caption = lData
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetChannelCaption(lIndex As Integer, lData As String)"
End Sub

Public Sub LoadChannel(lIndex As Integer, lName As String)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
Load lChannels.cChannel(lIndex)
lChannels.cChannel(lIndex).Show
lChannels.cChannel(lIndex).Tag = LCase(lName)
lChannels.cName(lIndex) = lName
lChannels.cModes(lIndex) = ""
UpdateChannelCaption lIndex
lChannels.cStats(lIndex).cName = lName
ProcessReplaceString sNowTalkingIn, lChannels.cChannel(lIndex).txtIncoming, lName
Call AddTaskPanel(lName, 2)
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Sub LoadChannel(lIndex As Integer, lName As String)"
    Err.Clear
End Sub

Public Sub UpdateChannelCaption(lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If ReturnChannelLimit(lIndex) = "" Then
    SetChannelCaption lIndex, ReturnChannelName(lIndex) & " [" & ReturnChannelNamesCount(lIndex) & "] [+" & ReturnChannelModes(lIndex) & "] :" & ReturnChannelTopic(lIndex)
Else
    SetChannelCaption lIndex, ReturnChannelName(lIndex) & " [" & ReturnChannelNamesCount(lIndex) & "] [+" & ReturnChannelModes(lIndex) & " " & ReturnChannelLimit(lIndex) & "] :" & ReturnChannelTopic(lIndex)
End If
End Sub

Public Function ReturnChannelLimit(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
ReturnChannelLimit = lChannels.cLimit(lIndex)
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnChannelLimit(lIndex As Integer) As Integer"
End Function

'Public Function ReturnChannelWindowNamesColor(lIndex As Integer, lNamesIndex As Integer) As String
'ReturnChannelWindowNamesColor = lChannels.cChannel(lIndex).lvwNames.ListItems(lNamesIndex).ForeColor
'ReturnChannelWindowNamesColor = lchannels.cChannel(lindex).lvwNames.item
'ErrHandler:
'    ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnChannelWindowNamesColor(lIndex As Integer, lNamesIndex As Integer) As String"
'    Err.Clear
'End Function

Public Sub SetChannelTopicTextBox(lIndex As Integer, lData As String)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
lChannels.cChannel(lIndex).txtTopic.Text = lData
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Function SetChannelWindowColors(lIndex As Integer, lBackColor As String, lForeColor As String)"
    Err.Clear
End Sub

Public Sub SetChannelWindowColors(lIndex As Integer, lBackColor As String, lForeColor As String)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
With lChannels.cChannel(lIndex)
    .txtIncoming.SetBackColor lBackColor
    .txtOutgoing.BackColor = lBackColor
    .txtOutgoing.ForeColor = lForeColor
    .lvwNames.BackColor = lBackColor
    .lvwNames.ForeColor = lForeColor
    .lstSent.BackColor = lBackColor
    .lstSent.ForeColor = lForeColor
End With
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Function SetChannelWindowColors(lIndex As Integer, lBackColor As String, lForeColor As String)"
    Err.Clear
End Sub

Public Function ReturnChannelNames(lIndex As Integer, lNamesIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'ReturnChannelNames = lChannels.cChannel(lIndex).lvwNames.ListItems(lNamesIndex).Text
ReturnChannelNames = lChannels.cChannel(lIndex).lvwNames.ItemText(lNamesIndex)
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnChannelNames(lIndex As Integer, lNamesIndex As Integer) As String"
    Err.Clear
End Function

Public Function ReturnChannelIncomingTBox(lIndex As Integer) As ctlTBox
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Set ReturnChannelIncomingTBox = lChannels.cChannel(lIndex).txtIncoming
End Function

Public Function ReturnChannelNamesListView(lIndex As Integer) As ctlListView
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Set ReturnChannelNamesListView = lChannels.cChannel(lIndex).lvwNames
End Function

Public Sub SetFocusOnChannel(lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lChannels.cChannel(lIndex).SetFocus
End Sub

Public Sub SetChannelWindowState(lIndex As Integer, lState As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lChannels.cChannel(lIndex).WindowState = lState
End Sub

Public Function ReturnChannelHwnd(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnChannelHwnd = lChannels.cChannel(lIndex).hWnd
End Function

Public Function ReturnChannelCaption(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnChannelCaption = lChannels.cChannel(lIndex).Caption
End Function

Public Sub RemoveChannelName(lIndex As Integer, lNameIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
'lChannels.cChannel(lIndex).lvwNames.ListItems.Remove lNameIndex
lChannels.cChannel(lIndex).lvwNames.ItemRemove lNameIndex
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Sub RemoveChannelName(lIndex As Integer, lNameIndex As Integer)"
    Err.Clear
End Sub

Public Sub SetChannelModes(lIndex As Integer, lData As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lChannels.cModes(lIndex) = lData
End Sub

Public Sub SetChannelWindowNamesForeColor(lIndex As Integer, lNameIndex As Integer, lForeColor As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'lChannels.cChannel(lIndex).lvwNames.ListItems(lNameIndex).ForeColor = lForeColor
'FIX!!!!!!!!!!!!!!!!!!1
End Sub

Public Function ReturnChannelName(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnChannelName = lChannels.cName(lIndex)
End Function

Public Function FindChannelIndex(lName As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lName) <> 0 Then
    For i = 1 To 150
        If Len(lChannels.cName(i)) <> 0 And Trim(LCase(lChannels.cName(i))) = Trim(LCase(lName)) Then
            FindChannelIndex = i
            Exit Function
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindChannelIndex(lName As String) As Integer"
End Function

Public Function ReturnChannelModes(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnChannelModes = lChannels.cModes(lIndex)
End Function

Public Function ReturnChannelTopic(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnChannelTopic = lChannels.cTopic(lIndex)
End Function

Public Sub SetChannelName(lIndex As Integer, lName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lChannels.cName(lIndex) = lName
End Sub

Public Sub UnloadChannel(lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload lChannels.cChannel(lIndex)
End Sub

Public Function ReturnChannelUBound() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnChannelUBound = lChannelUBound
End Function

Public Function ReturnChannelNamesCount(lIndex As Integer) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnChannelNamesCount = lChannels.cChannel(lIndex).lvwNames.Count
End Function

Public Sub AddUserToNicklist(lNickName As String, lListView As ctlListView)
'Public Sub AddUserToNicklist(lNickname As String, lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lListView.ItemAdd 0, lNickName, 0, 0
'Dim i As Integer, c As Integer
'Dim mColor(0 To 15) As Long
'mColor(0) = vbWhite
'mColor(1) = vbBlack
'mColor(2) = RGB(42, 42, 87)
'mColor(3) = RGB(33, 112, 33)
'mColor(4) = vbRed
'mColor(5) = RGB(109, 50, 50)
'mColor(6) = RGB(119, 33, 119)
'mColor(7) = RGB(252, 127, 0)
'mColor(8) = RGB(195, 195, 56)
'mColor(9) = RGB(0, 252, 0)
'mColor(10) = RGB(89, 167, 179)
'mColor(11) = RGB(0, 255, 255)
'mColor(12) = vbBlue
'mColor(13) = RGB(255, 0, 255)
'mColor(14) = RGB(127, 127, 127)
'mColor(15) = RGB(210, 210, 210)
'If Len(lNickName) <> 0 Then
'    If lSettings.sColoredNicklist = True Then
'        lListView.Sorted = False
'        'lockwindowupdate mdiNexIRC.hwnd
'        If Left(lNickName, 1) = "@" Then
'            Select Case lListView.BackColor
'            Case 0
'                lListView.ListItems.Add , , Right(lNickName, Len(lNickName) - 1), 3, 3
'            Case Else
'                lListView.ListItems.Add , , Right(lNickName, Len(lNickName) - 1), 1, 1
'            End Select
'            i = FindListViewIndex(lListView, lNickName)
'            c = Int(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nOpColor) - 1
'            lListView.ListItems(i).ForeColor = mColor(c)
'            lListView.ListItems(i).Tag = "o"
'        ElseIf Left(lNickName, 1) = "+" Then
'            Select Case lListView.BackColor
'            Case 0
'                lListView.ListItems.Add 1, , Right(lNickName, Len(lNickName) - 1), 4, 4
'            Case Else
'                lListView.ListItems.Add 1, , Right(lNickName, Len(lNickName) - 1), 2, 2
'            End Select
'            i = FindListViewIndex(lListView, lNickName)
'            c = Int(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nVoiceColor) - 1
'            lListView.ListItems(i).ForeColor = mColor(c)
'            lListView.ListItems(i).Tag = "v"
'        Else
'            lListView.ListItems.Add , , lNickName
'            i = FindListViewIndex(lListView, lNickName)
'            lListView.ListItems(i).Tag = "n"
'            If lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nNormalColor = "0" And lListView.BackColor = "0" Then
'                lListView.ListItems(i).ForeColor = vbWhite
'            Else
'                c = Int(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nNormalColor) - 1
'                lListView.ListItems(i).ForeColor = mColor(c)
'            End If
'        End If
'        'lockwindowupdate 0
'    Else
'        lListView.ListItems.Add , , lNickName
'        lListView.Sorted = True
'        Select Case Left(lNickName, 1)
'        Case "@"
'            lListView.ListItems(i).Tag = "o"
'        Case "+"
'            lListView.ListItems(i).Tag = "v"
'        Case Else
'            lListView.ListItems(i).Tag = "n"
'        End Select
'    End If
'End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub AddUserToNicklist(lNickname As String, lListView As ListView)"
End Sub
