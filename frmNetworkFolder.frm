VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmServerFolder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Server Folder"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
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
   Icon            =   "frmNetworkFolder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   3690
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   7
      Top             =   360
      Width           =   2415
   End
   Begin VB.ListBox lstServers 
      Height          =   2985
      ItemData        =   "frmNetworkFolder.frx":0CCA
      Left            =   120
      List            =   "frmNetworkFolder.frx":0CD1
      TabIndex        =   6
      Top             =   720
      Width           =   2415
   End
   Begin VB.CheckBox chkShowOnStartup 
      Appearance      =   0  'Flat
      Caption         =   "Show on Startup"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3780
      Width           =   1935
   End
   Begin OsenXPCntrl.OsenXPButton cmdHelp 
      Height          =   405
      Left            =   2640
      TabIndex        =   0
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Help"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmNetworkFolder.frx":0CE1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdOK 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   2640
      TabIndex        =   1
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmNetworkFolder.frx":0E43
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdConnect 
      Height          =   405
      Left            =   2640
      TabIndex        =   2
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Connect"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   285
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmNetworkFolder.frx":0FA5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdRemove 
      Height          =   405
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Remove"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmNetworkFolder.frx":1107
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdAdd 
      Height          =   405
      Left            =   2640
      TabIndex        =   4
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Add"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmNetworkFolder.frx":1269
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblHelp 
      Caption         =   "Connect to Network:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmServerFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
For i = 0 To lServerFolder.sCount
    If Len(lServerFolder.sServer(i).sDescription) <> 0 Then
        lstServers.AddItem lServerFolder.sServer(i).sDescription
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub
