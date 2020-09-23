VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmSearchThroughPlaylist 
   Caption         =   "NexIRC - Search Playlists"
   ClientHeight    =   945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearchThroughPlaylist.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   945
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin OsenXPCntrl.OsenXPButton cmdCancel 
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Cancel"
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
      MICON           =   "frmSearchThroughPlaylist.frx":0CCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdSearch 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Search"
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
      MICON           =   "frmSearchThroughPlaylist.frx":0E2C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Search for:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmSearchThroughPlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSearch_Click()
'On Local Error Resume Next
Dim i As Integer, msg As String, f As New frmPlaylist, r As Integer, lFileCol() As Integer
For i = 0 To lFiles.fCount
    If InStr(LCase(lFiles.fFile(i).fFilename), LCase(txtSearch.text)) Then
        r = r + 1
        lFileCol(r) = i
    End If
Next i
f.Caption = "NexIRC - Search Results (" & r & ")"
f.Show
For i = 0 To r
    If r <> 0 And Len(lFiles.fFile(r).fFilename) <> 0 Then
        
    End If
Next i
End Sub

Private Sub Form_Load()
'On Local Error Resume Next

End Sub
