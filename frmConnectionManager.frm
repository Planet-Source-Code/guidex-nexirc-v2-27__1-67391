VERSION 5.00
Begin VB.Form frmConnectionManager 
   Caption         =   "NexIRC - Connections"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConnectionManager.frx":0000
   LinkTopic       =   "frmConnectionManager"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4125
   Begin VB.Timer tmrConnections 
      Interval        =   2000
      Left            =   3600
      Top             =   120
   End
   Begin VB.ListBox lstConnections 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2700
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmConnectionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ActivateResize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Form_Resize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ActivateResize()"
End Sub

Public Sub CheckOpenConnections()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, c As Integer
If Len(Me.Tag) = 0 Then Me.Tag = "0"
If Int(Me.Tag) <> ReturnStatusWindowCount Then
    lstConnections.Clear
    For i = 0 To ReturnStatusWindowCount
        If Len(ReturnStatusWindowServer(i)) <> 0 Then
            If Len(ReturnStatusWindowCaption(i)) <> 0 Then
                If Left(LCase(ReturnStatusWindowCaption(i)), 6) <> "nexirc" Then
                    lstConnections.AddItem ReturnStatusWindowCaption(i) & " (" & ReturnStatusWindowServer(i) & ":" & ReturnStatusWindowPort(i) & ")"
                    c = c + 1
                    SetConnectionManagerCaption c
                End If
            End If
        End If
    Next i
Else
    Dim msg As String
    c = ReturnStatusWindowCount
    For i = 0 To lstConnections.ListCount
        msg = Left(lstConnections.List(i), 9)
        If Len(msg) <> 0 Then
            If Left(LCase(msg), 6) = "status" Then
                If Right(msg, 1) = ":" Then msg = Left(msg, Len(msg) - 1)
                msg = LCase(msg)
                If FindStatusWindowIndexByTag(msg) = 0 Then
                    lstConnections.RemoveItem i
                    c = lstConnections.ListCount
                    SetConnectionManagerCaption c
                End If
            ElseIf Left(LCase(msg), 6) = "nexirc" Then
                lstConnections.RemoveItem i
                c = lstConnections.ListCount
                SetConnectionManagerCaption c
            End If
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub CheckOpenConnections()"
End Sub

Public Function SetConnectionManagerCaption(lStatusCount As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Tag = lStatusCount
'lStatusWindows.sCount = lStatusCount
Me.Caption = "NexIRC - Connection manager (" & lStatusCount & ") connection(s)"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function SetConnectionManagerCaption(lStatusCount As Integer)"
End Function

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = mdiNexIRC.Icon
lSettings.sConnectionManagerVisible = True
Call AddTaskPanel("Manager", 1)
CheckOpenConnections
lstConnections.BackColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor
lstConnections.ForeColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sTextColor
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Call RemoveTaskbar("Manager")
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)"
End Sub

Private Sub Form_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lstConnections.Width = Me.ScaleWidth
lstConnections.Height = Me.ScaleHeight
If lstConnections.Left <> 0 Then lstConnections.Left = 0
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Resize()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sConnectionManagerVisible = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub lstConnections_DblClick()
If lSettings.sHandleErrors = True Then On Local Error Resume Next




If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lstConnections_DblClick()"
End Sub

Private Sub lstConnections_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 2 Then
    PopupMenu frmMenus.mnuConnectionManager
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lstConnections_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub tmrConnections_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckOpenConnections
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrConnections_Timer()"
End Sub
