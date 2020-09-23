VERSION 5.00
Begin VB.Form frmScriptManager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Scripts"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   4080
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmScriptManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4080
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstScripts 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2295
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin nexIRC.ctlXPButton cmdCancel 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2040
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Cancel"
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
      MICON           =   "frmScriptManager.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdOK 
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "OK"
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
      MICON           =   "frmScriptManager.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdDelete 
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Delete"
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
      MICON           =   "frmScriptManager.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdNew 
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "New"
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
      MICON           =   "frmScriptManager.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdEdit 
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   "Edit"
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
      MICON           =   "frmScriptManager.frx":007C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuRun 
         Caption         =   "&Run"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowToUseThisWindow 
         Caption         =   "H&ow to use this Window"
      End
   End
End
Attribute VB_Name = "frmScriptManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdCancel_Click()"
End Sub

Private Sub cmdDelete_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, mbox As VbMsgBoxResult
msg = App.Path & "\data\scripts\" & lstScripts.Text
If DoesFileExist(msg) = True Then
    If lSettings.sGeneralPrompts = True Then
        mbox = MsgBox("Are you sure you wish to delete " & lstScripts.Text & "?", vbYesNo + vbQuestion)
    Else
        mbox = vbYes
    End If
    If mbox = vbYes Then
        lstScripts.RemoveItem lstScripts.ListIndex
        Kill msg
    End If
Else
    If lSettings.sGeneralPrompts = True Then
        MsgBox "Couldn't find " & msg
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdDelete_Click()"
End Sub

Private Sub cmdEdit_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
OpenTextFile App.Path & "\data\scripts\nexirc\" & lstScripts.Text
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdEdit_Click()"
End Sub

Private Sub cmdNew_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
NewScriptFile
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdNew_Click()"
End Sub

Private Sub cmdOK_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg2 As String
If DoesFileExist(App.Path & "\data\scripts\" & lstScripts.Text) = True Then
    LoadScript lstScripts.Text
Else
    If lSettings.sGeneralPrompts = True Then
        MsgBox "Couldn't find " & App.Path & "\data\scripts\" & lstScripts.Text
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdOK_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
Me.Icon = mdiNexIRC.Icon
SetButtonType cmdEdit
SetButtonType cmdNew
SetButtonType cmdOK
SetButtonType cmdDelete
SetButtonType cmdCancel

File1.Path = App.Path & "\data\scripts\nexirc\"

For i = 0 To File1.ListCount
    If Len(File1.List(i)) Then lstScripts.AddItem File1.List(i)
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub mnuClose_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuClose_Click()"
End Sub

Private Sub mnuDelete_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdDelete_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuDelete_Click()"
End Sub

Private Sub mnuExit_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuDelete_Click()"
End Sub

Private Sub mnuHowToUseThisWindow_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
MsgBox "No help is available for this Topic", vbExclamation
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuHowToUseThisWindow_Click()"
End Sub

Private Sub mnuRun_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdOK_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub mnuRun_Click()"
End Sub
