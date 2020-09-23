Attribute VB_Name = "mdlHelp"
Option Explicit

Public Sub DisplayHelpInformationByName(lName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case lName
Case "frmAddMedia"
    DisplayHelpInformation 2
Case "frmChannel"
    DisplayHelpInformation 8
Case "frmChannels"
    DisplayHelpInformation 11
Case "frmChat"
    DisplayHelpInformation 12
Case "frmConnectionManager"
    DisplayHelpInformation 14
Case "frmIRCServer"
    DisplayHelpInformation 20
Case "frmMOTD"
    DisplayHelpInformation 23
Case "frmNotify"
    DisplayHelpInformation 24
Case "frmQuery"
    DisplayHelpInformation 27
Case "frmPlaylist"
    DisplayHelpInformation 26
Case "frmStatus"
    DisplayHelpInformation 33
Case "frmTextEditor"
    DisplayHelpInformation 34
End Select
If Err.Number = 91 Then
    frmCustomize.Show 0, mdiNexIRC
    frmCustomize.Visible = True
    frmCustomize.fraSettings(8).Visible = True
    frmCustomize.cboHelpTopic.ListIndex = 12
    frmCustomize.optCheck(8).Value = True
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub DisplayHelpInformationByName(lName As String)"
End Sub

Public Sub DisplayHelpInformation(lComboIndex As Integer, Optional lFocus As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To 9
    frmCustomize.fraSettings(i).Visible = False
    frmCustomize.optCheck(i).Value = False
Next i
frmCustomize.Show 0, mdiNexIRC
frmCustomize.chkShowMe.Visible = False
frmCustomize.fraSettings(8).Visible = True
frmCustomize.cmdConnect.Visible = False
frmCustomize.cboHelpTopic.ListIndex = lComboIndex
frmCustomize.optCheck(8).Value = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub DisplayHelpInformation(lComboIndex As Integer)"
End Sub
