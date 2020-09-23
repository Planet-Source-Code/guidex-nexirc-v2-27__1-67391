VERSION 5.00
Begin VB.Form frmChannelFolder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Channel Folder"
   ClientHeight    =   3675
   ClientLeft      =   6225
   ClientTop       =   3405
   ClientWidth     =   3690
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChannelFolder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   3690
   StartUpPosition =   1  'CenterOwner
   Begin nexIRC.ctlXPButton cmdJoin 
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Join"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmChannelFolder.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chkShowChanFolder 
      Appearance      =   0  'Flat
      Caption         =   "Popup on connect"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   3390
      Width           =   2175
   End
   Begin VB.ListBox lstChannels 
      Height          =   2580
      IntegralHeight  =   0   'False
      ItemData        =   "frmChannelFolder.frx":0028
      Left            =   120
      List            =   "frmChannelFolder.frx":002A
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtChannel 
      Height          =   285
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin nexIRC.ctlXPButton cmdAdd 
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      ToolTipText     =   "Add Bot to List"
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Add"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmChannelFolder.frx":002C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdAutoJoin 
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      ToolTipText     =   "Add Bot to List"
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Auto Join"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmChannelFolder.frx":0048
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdHelp 
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      ToolTipText     =   "Add Bot to List"
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Help"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmChannelFolder.frx":0064
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdClose 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      ToolTipText     =   "Add Bot to List"
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmChannelFolder.frx":0080
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblChannelFolder 
      Caption         =   "Enter name of channel to join:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmChannelFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(txtChannel.Text) <> 0 Then
    If Left(txtChannel.Text, 1) <> "#" Then
        txtChannel.Text = "#" & Trim(txtChannel.Text)
    End If
    If AddtoChanFolder(txtChannel.Text) = True Then
        lstChannels.AddItem txtChannel.Text
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAdd_Click()"
End Sub

Private Sub cmdAutoJoin_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
i = AddAutoJoin(lstChannels.Text, lSettings.sNetwork)
frmAutoJoin.Show
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAutoJoin_Click()"
End Sub

Private Sub cmdClose_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sOptions.oShowChannelFolder = GetCheckboxValue(chkShowChanFolder)
SaveChanFolders
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdOK_Click()"
End Sub

Private Sub cmdHelp_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DisplayHelpInformation 9
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub OsenXPButton1_Click()"
End Sub

Private Sub cmdJoin_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.SendData "JOIN " & lstChannels.List(lstChannels.ListIndex) & vbCrLf
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdJoin_Click()"
End Sub

Private Sub cmdOK_Click()
End Sub

Private Sub cmdRemove_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
RemoveFromChanFolder lstChannels.Text
lstChannels.RemoveItem lstChannels.ListIndex
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdRemove_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
Me.Icon = mdiNexIRC.Icon
SetButtonType cmdAdd

SetButtonType cmdJoin
SetButtonType cmdAutoJoin
SetButtonType cmdClose
SetButtonType cmdHelp
'SetButtonType OsenXPButton1
'SetButtonType cmdOK
For i = 0 To 150
    If Len(ReturnChannelFolderChannel(i)) <> 0 Then
        lstChannels.AddItem ReturnChannelFolderChannel(i)
    End If
Next i
SetCheckBoxValue chkShowChanFolder, lSettings.sOptions.oShowChannelFolder
lSettings.sChannelFolderVisible = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sChannelFolderVisible = False
lSettings.sOptions.oShowChannelFolder = GetCheckboxValue(chkShowChanFolder)
SaveChanFolders
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub lstChannels_DblClick()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Left(lstChannels.List(lstChannels.ListIndex), 1) = "#" Then
    If lSettings.sActiveServerForm.tcp.State = sckConnected Then lSettings.sActiveServerForm.tcp.SendData "JOIN " & lstChannels.List(lstChannels.ListIndex) & vbCrLf
Else
    lSettings.sActiveServerForm.tcp.SendData "JOIN #" & lstChannels.List(lstChannels.ListIndex) & vbCrLf
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lstChannels_DblClick()"
End Sub

Private Sub OsenXPButton1_Click()

End Sub

Private Sub txtChannel_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtChannel.SelStart = 0
txtChannel.SelLength = Len(txtChannel.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtChannel_GotFocus()"
End Sub
