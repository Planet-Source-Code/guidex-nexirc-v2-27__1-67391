Attribute VB_Name = "mdlQuery"
Option Explicit
Private Const lQueryUBound = 64
Private Type gQuerys
    qQuery(lQueryUBound) As New frmQuery
    qName(lQueryUBound) As String
End Type
Private lQuerys As gQuerys

Public Sub SetQueryWindowColors(lIndex As Integer, lBackColor As String, lForeColor As String)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
lQuerys.qQuery(lIndex).BackColor = lBackColor
lQuerys.qQuery(lIndex).txtIncoming.SetBackColor lBackColor
lQuerys.qQuery(lIndex).txtOutgoing.BackColor = lBackColor
lQuerys.qQuery(lIndex).txtOutgoing.ForeColor = lForeColor
ErrHandler:
End Sub

Public Sub SetQueryCaption(lIndex As Integer, lData As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lQuerys.qQuery(lIndex).Caption = lData
End Sub

Public Function ReturnQueryIncomingTBox(lIndex As Integer) As ctlTBox
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Set ReturnQueryIncomingTBox = lQuerys.qQuery(lIndex).txtIncoming
End Function

Public Sub SetFocusOnQueryWindow(lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lQuerys.qQuery(lIndex).SetFocus
End Sub

Public Sub SetQueryWindowState(lIndex As Integer, lState As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lQuerys.qQuery(lIndex).WindowState = lState
End Sub

Public Function ReturnQueryHwnd(lIndex As Integer) As Long
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnQueryHwnd = lQuerys.qQuery(lIndex).hWnd
End Function

Public Function ReturnQueryCaption(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnQueryCaption = lQuerys.qQuery(lIndex).Caption
End Function

Public Sub LoadQueryWindow(lIndex As Integer, lNickName As String, lEmail As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Load lQuerys.qQuery(lIndex)
lQuerys.qQuery(lIndex).SetQueryNickname lNickName
lQuerys.qQuery(lIndex).txtOutgoing.BackColor = IRCCodeToRGB(Color.BGText)
lQuerys.qQuery(lIndex).txtOutgoing.ForeColor = IRCCodeToRGB(Color.Normal)
lQuerys.qQuery(lIndex).txtIncoming.SetBackColor IRCCodeToRGB(Color.BGText)
lQuerys.qQuery(lIndex).Caption = lNickName & " [" & lEmail & "]"
lQuerys.qName(lIndex) = lNickName
Call AddTaskPanel(lNickName, 1)
IsUserOnline lQuerys.qQuery(lIndex).txtIncoming, lNickName
End Sub

Public Sub SetQueryName(lIndex As Integer, lName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lQuerys.qName(lIndex) = lName
End Sub

Public Function ReturnQueryName(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lIndex < 64 Then ReturnQueryName = lQuerys.qName(lIndex)
End Function

Public Function FindQueryIndex(lNickName As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To 64
    If LCase(lNickName) = LCase(lQuerys.qQuery(i).ReturnQueryNickname) Then
        FindQueryIndex = i
        Exit Function
    End If
Next i
End Function
