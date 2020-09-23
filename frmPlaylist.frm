VERSION 5.00
Begin VB.Form frmPlaylist 
   Caption         =   "Playlist"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlaylist.frx":0000
   LinkTopic       =   "frmPlaylist"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   6360
   Begin VB.Timer tmrCheckBackColor 
      Interval        =   2000
      Left            =   120
      Top             =   2520
   End
   Begin VB.ListBox lstPlaylist 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2460
      IntegralHeight  =   0   'False
      Left            =   0
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmPlaylist"
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

Public Sub RefreshPlaylist()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String
lstPlaylist.Clear
For i = 0 To lFiles.fCount
    If Len(lFiles.fFile(i).fFilename) <> 0 Then
        msg = lFiles.fFile(i).fFilename
        msg = GetFileTitle(msg)
        lstPlaylist.AddItem msg
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub RefreshPlaylist()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String
Me.Icon = mdiNexIRC.Icon
lstPlaylist.Width = Me.ScaleWidth
lstPlaylist.Height = Me.ScaleHeight
If DoesPanelExistInStatusBar("Playlist", mdiNexIRC.StatusBar) = False Then
    Call AddTaskPanel("Playlist", 1)
End If
lSettings.sPlaylistVisible = True
For i = 0 To lFiles.fCount
    If Len(lFiles.fFile(i).fFilename) <> 0 Then
        msg = lFiles.fFile(i).fFilename
        msg = GetFileTitle(msg)
        lstPlaylist.AddItem msg
    End If
Next i
lstPlaylist.BackColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor
lstPlaylist.ForeColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sTextColor
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lstPlaylist.Left <> 0 Then lstPlaylist.Left = 0
If Me.ScaleWidth <> 0 Then
    lstPlaylist.Width = Me.ScaleWidth
    lstPlaylist.Height = Me.ScaleHeight
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Resize()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Call RemoveTaskbar("Playlist")
lSettings.sPlaylistVisible = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub lstPlaylist_DblClick()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
i = FindFileIndexByFilename(lstPlaylist.Text)
PlayFile lFiles.fFile(i).fFilename
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lstPlaylist_DblClick()"
End Sub

Private Sub lstPlaylist_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
RaiseAllStatusbarPanels mdiNexIRC.StatusBar
i = FindPanelIndex("Playlist", mdiNexIRC.StatusBar)
If i <> 0 Then mdiNexIRC.StatusBar.Panels(i).Bevel = sbrInset
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lstPlaylist_GotFocus()"
End Sub

Private Sub lstPlaylist_KeyDown(KeyCode As Integer, Shift As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If KeyCode = 46 Then
    'lockwindowupdate frmPlaylist.hwnd
    For i = 0 To frmPlaylist.lstPlaylist.ListCount
        If frmPlaylist.lstPlaylist.ListCount <> i Then
            If frmPlaylist.lstPlaylist.Selected(i) = True Then
                RemoveFromPlaylist frmPlaylist.lstPlaylist.List(i)
                frmPlaylist.lstPlaylist.RemoveItem i
                i = i - 1
            End If
        Else
            Exit For
        End If
    Next i
    SavePlaylist
    'lockwindowupdate 0
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lstPlaylist_KeyDown(KeyCode As Integer, Shift As Integer)"
End Sub

Private Sub lstPlaylist_KeyPress(KeyAscii As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If KeyAscii = 27 Then Me.WindowState = vbMinimized
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lstPlaylist_KeyPress(KeyAscii As Integer)"
End Sub

Private Sub lstPlaylist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 2 Then
    PopupMenu frmMenus.mnuPlaylist
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lstPlaylist_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub tmrCheckBackColor_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lstPlaylist.BackColor <> lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor Then
    lstPlaylist.BackColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrCheckBackColor_Timer()"
End Sub
