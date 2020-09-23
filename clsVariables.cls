VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Public Function JoinChannel() As String
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
'JoinChannel = lEvents.JoinChannel
'End Function

'Public Function NickJoin() As String
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
'NickJoin = lEvents.NickJoin
'End Function

'Public Function ChanPart() As String
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
'ChanPart = lEvents.ChanPart
'End Function

'Public Function NickPart() As String
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
'NickPart = lEvents.NickPart
'End Function

'Public Function Names(lName As String) As String
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
'Dim x As Integer, i As Integer
'For i = 1 To lChannelUBound
'    If LCase(lChannelName(i)) = LCase(lName) Then
'        Dim Name As String
'        For x = 1 To lChannel(i).lstNames.ListItems.Count - 1
'            Name = Name & " " & lChannel(i).lstNames.ListItems(x).Text
'            Exit For
'        Next x
'        Names = Name
'    End If
'Next i
'End Function

Public Function ServerName() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ServerName = lRaw.Server
End Function

Public Function ServerText() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ServerText = lRaw.ServerText
End Function

Public Function Num() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Num = lRaw.Num
End Function

Public Function Chan(namex As String, var As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To lChannelUBound
    If chanstats(i).Name = LCase(namex) Then
        Select Case LCase(var)
            Case "t"
                Chan = chanstats(i).Topic
        End Select
    End If
Next i
End Function

Public Function EchoChannel(namex As String, strData As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To lChannelUBound
    If LCase(lChannelName(i)) = LCase(namex) Then
        Dim RTF As TBox
        Set RTF = lChannel(i).txtIncoming
        Exit For
    End If
Next i
End Function
