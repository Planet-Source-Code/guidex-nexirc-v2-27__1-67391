VERSION 5.00
Begin VB.Form frmAddMedia 
   Caption         =   "Browse for Media"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddToPlaylistSearch.frx":0000
   LinkTopic       =   "frmAddMedia"
   MDIChild        =   -1  'True
   ScaleHeight     =   2580
   ScaleWidth      =   4665
   Begin VB.DriveListBox ctlDrive 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   2160
      Width           =   4575
   End
   Begin VB.FileListBox ctlFiles 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   2280
      Pattern         =   "*.wav;*.mp3;*.wma;*.wmv;*.snd;*.au;*.ogg)|*.wav;*.mp3;*.wma;*.wmv;*.snd;*.au;*.ogg"
      TabIndex        =   1
      Top             =   -120
      Width           =   2295
   End
   Begin VB.DirListBox ctlDir 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmAddMedia"
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

Private Sub ctlDir_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ctlFiles.Path = ctlDir.Path
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub ctlDir_Change()"
End Sub

Private Sub ctlDrive_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ctlDir.Path = ctlDrive.Drive
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub ctlDrive_Change()"
End Sub

Private Sub ctlFiles_DblClick()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PlayFile ctlDir.Path & "\" & ctlFiles.List(ctlFiles.ListIndex)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub ctlFiles_DblClick()"
End Sub

Private Sub ctlFiles_PathChange()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, a As Integer, msg As String, msg2 As String, F As Integer
If ctlFiles.ListCount = 0 Then Exit Sub
For i = 0 To ctlFiles.ListCount
    If Len(ctlFiles.List(i)) <> 0 Then
        a = FindFileIndexByFilename(ctlFiles.List(i))
        If a = 0 Then
            msg2 = ctlDir.Path & "\" & ctlFiles.List(i)
            If lSettings.sExlusiveToMp3InPlaylist = True And LCase(Right(msg2, 4)) <> ".mp3" Then
                GoTo NextI
            End If
            F = F + 1
            lFiles.fCount = lFiles.fCount + 1
            lFiles.fFile(lFiles.fCount).fFilename = msg2
            If lSettings.sPlaylistVisible = True Then
                msg = lFiles.fFile(lFiles.fCount).fFilename
                msg = GetFileTitle(msg)
                frmPlaylist.lstPlaylist.AddItem msg
            End If
        End If
    End If
NextI:
Next i
If F <> 0 Then SavePlaylist
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub ctlFiles_PathChange()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = mdiNexIRC.Icon
lSettings.sAddMediaVisible = True
If lSettings.sExlusiveToMp3InPlaylist = True Then
    ctlFiles.Pattern = "*.mp3"
Else
    ctlFiles.Pattern = "*.wav;*.mp3;*.wma;*.wmv;*.snd;*.au;*.ogg"
End If
With lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex)
    ctlDir.BackColor = .sBackColor
    ctlDir.ForeColor = .sTextColor
    ctlDrive.BackColor = .sBackColor
    ctlDrive.ForeColor = .sTextColor
    ctlFiles.BackColor = .sBackColor
    ctlFiles.ForeColor = .sTextColor
    BackColor = .sBackColor
End With
ActivateResize
ctlDir.Path = CurDir
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If ctlDir.Left <> 0 Then ctlDir.Left = 0
ctlDir.Width = (Me.ScaleWidth / 2)
If Me.ScaleHeight <> 0 Then ctlDir.Height = (Me.ScaleHeight - ctlDrive.Height)
ctlFiles.Left = ctlDir.Width
ctlFiles.Width = Me.ScaleWidth / 2
If ctlFiles.Top <> 0 Then ctlFiles.Top = 0
If Me.ScaleHeight <> 0 Then ctlFiles.Height = (Me.ScaleHeight - ctlDrive.Height)
ctlDrive.Top = (Me.ScaleHeight - ctlDrive.Height)
ctlDrive.Width = Me.ScaleWidth
If ctlDrive.Left <> 0 Then ctlDrive.Left = 0
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Resize()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sAddMediaVisible = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub
