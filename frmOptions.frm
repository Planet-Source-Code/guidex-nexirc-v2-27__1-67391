VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Customize"
   ClientHeight    =   4920
   ClientLeft      =   5670
   ClientTop       =   3600
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin nexIRC.XP_ProgressBar XP_ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   251
      Top             =   3840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   6956042
   End
   Begin nexIRC.ctlXPButton optCheck 
      Height          =   360
      Index           =   6
      Left            =   120
      TabIndex        =   222
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      BTYPE           =   2
      TX              =   "About"
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
      MICON           =   "frmOptions.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton optCheck 
      Height          =   360
      Index           =   10
      Left            =   120
      TabIndex        =   221
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      BTYPE           =   2
      TX              =   "Bots"
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
      MICON           =   "frmOptions.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton optCheck 
      Height          =   360
      Index           =   9
      Left            =   120
      TabIndex        =   220
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      BTYPE           =   2
      TX              =   "Audio"
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
      MICON           =   "frmOptions.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton optCheck 
      Height          =   360
      Index           =   7
      Left            =   120
      TabIndex        =   219
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      BTYPE           =   2
      TX              =   "Ignore"
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
      MICON           =   "frmOptions.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton optCheck 
      Height          =   360
      Index           =   5
      Left            =   120
      TabIndex        =   218
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      BTYPE           =   2
      TX              =   "Text"
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
      MICON           =   "frmOptions.frx":007C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton optCheck 
      Height          =   360
      Index           =   4
      Left            =   120
      TabIndex        =   217
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      BTYPE           =   2
      TX              =   "Themes"
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
      MICON           =   "frmOptions.frx":0098
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton optCheck 
      Height          =   360
      Index           =   3
      Left            =   120
      TabIndex        =   216
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      BTYPE           =   2
      TX              =   "Notify"
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
      MICON           =   "frmOptions.frx":00B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton optCheck 
      Height          =   360
      Index           =   2
      Left            =   120
      TabIndex        =   215
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      BTYPE           =   2
      TX              =   "Options"
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
      MICON           =   "frmOptions.frx":00D0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton optCheck 
      Height          =   360
      Index           =   1
      Left            =   120
      TabIndex        =   214
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      BTYPE           =   2
      TX              =   "User"
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
      MICON           =   "frmOptions.frx":00EC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton optCheck 
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   213
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      BTYPE           =   2
      TX              =   "Network"
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
      MICON           =   "frmOptions.frx":0108
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   10
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":0124
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":0538
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":094C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":0D60
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":1174
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":1588
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":199C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":1DB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":21C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":25D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":29EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":2E00
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":3214
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":3628
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":3A3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":3E50
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wskTestConnection 
      Index           =   0
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrPortScan 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   120
      Top             =   0
   End
   Begin VB.Timer tmrPortScanTimeout 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   120
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   223
      Top             =   4320
      Width           =   7575
      Begin nexIRC.ctlXPButton cmdConnect 
         Height          =   375
         Left            =   2760
         TabIndex        =   246
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Connect"
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
         MICON           =   "frmOptions.frx":4264
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox chkThisSessionOnly 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Apply Only"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1320
         TabIndex        =   244
         Top             =   240
         Width           =   1815
      End
      Begin nexIRC.ctlXPButton cmdOK 
         Default         =   -1  'True
         Height          =   375
         Left            =   3960
         TabIndex        =   248
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "frmOptions.frx":4280
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdApply 
         Height          =   375
         Left            =   5160
         TabIndex        =   250
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Apply"
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
         MICON           =   "frmOptions.frx":429C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   6360
         TabIndex        =   252
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "frmOptions.frx":42B8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton optCheck 
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   242
         ToolTipText     =   "Show Help Topics"
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Check"
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
         MICON           =   "frmOptions.frx":42D4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   7560
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame fraSettings 
      BorderStyle     =   0  'None
      Caption         =   "Network"
      Height          =   4215
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.ComboBox cboServerMinimum 
         Height          =   315
         ItemData        =   "frmOptions.frx":42F0
         Left            =   5520
         List            =   "frmOptions.frx":4333
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox chkNewServerWindow 
         Appearance      =   0  'Flat
         Caption         =   "&New status window"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   960
         TabIndex        =   15
         Top             =   3840
         Width           =   1815
      End
      Begin VB.ComboBox cmbNetwork 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   0
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "cmbNetwork"
         Top             =   120
         Width           =   2175
      End
      Begin VB.TextBox txtServer 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   250
         Left            =   960
         MousePointer    =   1  'Arrow
         TabIndex        =   12
         Top             =   3000
         Width           =   5140
      End
      Begin VB.TextBox txtPort 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   250
         Left            =   960
         MousePointer    =   1  'Arrow
         TabIndex        =   13
         Top             =   3255
         Width           =   5140
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   250
         IMEMode         =   3  'DISABLE
         Left            =   960
         MousePointer    =   1  'Arrow
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   3490
         Width           =   5140
      End
      Begin MSComctlLib.ListView lvwServers 
         Height          =   1905
         Left            =   15
         TabIndex        =   6
         Top             =   495
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         TextBackground  =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   0
         MousePointer    =   1
         NumItems        =   0
      End
      Begin nexIRC.ctlXPButton cmdNetworkAdd 
         Height          =   315
         Left            =   2280
         TabIndex        =   2
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
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
         MICON           =   "frmOptions.frx":4381
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdNetworkDelete 
         Height          =   315
         Left            =   3360
         TabIndex        =   3
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
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
         MICON           =   "frmOptions.frx":439D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdSmallNetworks 
         Height          =   315
         Left            =   4440
         TabIndex        =   4
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BTYPE           =   2
         TX              =   "Small"
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
         MICON           =   "frmOptions.frx":43B9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdServerAdd 
         Height          =   315
         Left            =   0
         TabIndex        =   7
         Top             =   2520
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
         MICON           =   "frmOptions.frx":43D5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdServerEdit 
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Top             =   2520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
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
         MICON           =   "frmOptions.frx":43F1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdClearServers 
         Height          =   315
         Left            =   3120
         TabIndex        =   10
         Top             =   2520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BTYPE           =   2
         TX              =   "Clear"
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
         MICON           =   "frmOptions.frx":440D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdServerDelete 
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Top             =   2520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
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
         MICON           =   "frmOptions.frx":4429
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdScan 
         Height          =   315
         Left            =   4200
         TabIndex        =   11
         Top             =   2520
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         BTYPE           =   2
         TX              =   "Scan"
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
         MICON           =   "frmOptions.frx":4445
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00808080&
         Height          =   1935
         Left            =   0
         Top             =   480
         Width           =   6135
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   810
         Left            =   930
         Top             =   2985
         Width           =   5205
      End
      Begin VB.Label lblServer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Server:"
         Height          =   195
         Left            =   0
         TabIndex        =   16
         Top             =   3000
         Width           =   540
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Port:"
         Height          =   195
         Left            =   0
         TabIndex        =   17
         Top             =   3240
         Width           =   360
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P&assword:"
         Height          =   195
         Left            =   0
         TabIndex        =   18
         Top             =   3480
         Width           =   750
      End
   End
   Begin VB.Frame fraSettings 
      BorderStyle     =   0  'None
      Height          =   4215
      Index           =   10
      Left            =   1320
      TabIndex        =   183
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
      Begin VB.ComboBox cboBotTypes 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   185
         Top             =   120
         Width           =   4935
      End
      Begin VB.ComboBox cboAutoPerform 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   191
         Top             =   1440
         Width           =   4935
      End
      Begin VB.ComboBox cboCommands 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   189
         Top             =   840
         Width           =   2775
      End
      Begin VB.ComboBox cboBots 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   187
         Top             =   480
         Width           =   2775
      End
      Begin nexIRC.ctlXPButton cmdAddBot 
         Height          =   300
         Left            =   4080
         TabIndex        =   234
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
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
         MICON           =   "frmOptions.frx":4461
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdAddBotCommand 
         Height          =   300
         Left            =   4080
         TabIndex        =   235
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
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
         MICON           =   "frmOptions.frx":447D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdDeleteBot 
         Height          =   300
         Left            =   5160
         TabIndex        =   236
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
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
         MICON           =   "frmOptions.frx":4499
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdDeleteCommands 
         Height          =   300
         Left            =   5160
         TabIndex        =   237
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
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
         MICON           =   "frmOptions.frx":44B5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdAddAutoPerform 
         Height          =   300
         Left            =   1200
         TabIndex        =   238
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
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
         MICON           =   "frmOptions.frx":44D1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdDeleteAutoPerform 
         Height          =   300
         Left            =   2280
         TabIndex        =   239
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
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
         MICON           =   "frmOptions.frx":44ED
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdRun 
         Height          =   300
         Left            =   3360
         TabIndex        =   240
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         BTYPE           =   2
         TX              =   "Run"
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
         MICON           =   "frmOptions.frx":4509
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
         Height          =   300
         Left            =   4440
         TabIndex        =   241
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
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
         MICON           =   "frmOptions.frx":4525
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         X1              =   0
         X2              =   6240
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   6240
         Y1              =   1290
         Y2              =   1290
      End
      Begin VB.Label lblBotTypes 
         Caption         =   "Bot &Types:"
         Height          =   255
         Left            =   0
         TabIndex        =   184
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblAutoPreform 
         Caption         =   "&Auto Preform:"
         Height          =   255
         Left            =   0
         TabIndex        =   190
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblCommands 
         Caption         =   "&Commands:"
         Height          =   255
         Left            =   0
         TabIndex        =   188
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblBots 
         Caption         =   "&Bots:"
         Height          =   255
         Left            =   0
         TabIndex        =   186
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame fraSettings 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4215
      Index           =   9
      Left            =   1320
      TabIndex        =   181
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CheckBox chkPlaySounds 
         Appearance      =   0  'Flat
         Caption         =   "&Play System Sounds"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   3120
         TabIndex        =   173
         Top             =   840
         Width           =   2895
      End
      Begin VB.CheckBox chkEnableFind 
         Appearance      =   0  'Flat
         Caption         =   "&Enable Search"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   3120
         TabIndex        =   180
         Top             =   2520
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chkEnableListmedia 
         Appearance      =   0  'Flat
         Caption         =   "&Enable !List"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   3120
         TabIndex        =   178
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox chkOfferWhenPlayed 
         Appearance      =   0  'Flat
         Caption         =   "&Offer when Played"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   3120
         TabIndex        =   177
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CheckBox chkLogAudioDownloads 
         Appearance      =   0  'Flat
         Caption         =   "&Log Audio Downloads"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   3120
         TabIndex        =   179
         Top             =   2280
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox chkFileOfferInChannel 
         Appearance      =   0  'Flat
         Caption         =   "&File Offer in Channel"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   3120
         TabIndex        =   176
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CheckBox chkAudioServer 
         Appearance      =   0  'Flat
         Caption         =   "&Audio Server Enabled"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   3120
         TabIndex        =   175
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox chkInitialBass 
         Appearance      =   0  'Flat
         Caption         =   "Initial &Bass:"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   0
         TabIndex        =   155
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CheckBox chkInitialTreble 
         Appearance      =   0  'Flat
         Caption         =   "Initial &Treble:"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   0
         TabIndex        =   157
         Top             =   1800
         Width           =   1455
      End
      Begin MSComctlLib.Slider sldInitialCDAudio 
         Height          =   195
         Left            =   1440
         TabIndex        =   160
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   344
         _Version        =   393216
         Max             =   65535
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider sldInitialLineIN 
         Height          =   195
         Left            =   1440
         TabIndex        =   162
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   344
         _Version        =   393216
         Max             =   65535
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider sldInitialMic 
         Height          =   195
         Left            =   1440
         TabIndex        =   164
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   344
         _Version        =   393216
         Max             =   65535
         TickStyle       =   3
      End
      Begin VB.CheckBox chkInitialMIC 
         Appearance      =   0  'Flat
         Caption         =   "Initial &Mic:"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   0
         TabIndex        =   163
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CheckBox chkInitialLineIN 
         Appearance      =   0  'Flat
         Caption         =   "Initial Line &in:"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   0
         TabIndex        =   161
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CheckBox chkInitialCDAudio 
         Appearance      =   0  'Flat
         Caption         =   "Initial &CDAudio:"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   0
         TabIndex        =   159
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CheckBox chkInitialWave 
         Appearance      =   0  'Flat
         Caption         =   "Initial &Wave:"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   0
         TabIndex        =   153
         Top             =   1320
         Width           =   1335
      End
      Begin MSComctlLib.Slider sldInitialWave 
         Height          =   195
         Left            =   1440
         TabIndex        =   154
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   344
         _Version        =   393216
         Max             =   65535
         TickStyle       =   3
      End
      Begin VB.CheckBox chkLogoTwitchOnPeaks 
         Appearance      =   0  'Flat
         Caption         =   "&Logo Twitch on Peaks"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   0
         TabIndex        =   151
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox chkShuffle 
         Appearance      =   0  'Flat
         Caption         =   "&Shuffle"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   0
         TabIndex        =   169
         Top             =   3720
         Width           =   975
      End
      Begin VB.CheckBox chkContinuousPlay 
         Appearance      =   0  'Flat
         Caption         =   "&Enable"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   0
         TabIndex        =   168
         Top             =   3480
         Width           =   855
      End
      Begin VB.CheckBox chkShowQuickMix 
         Appearance      =   0  'Flat
         Caption         =   "&Mixer"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   0
         TabIndex        =   166
         Top             =   3000
         Width           =   855
      End
      Begin VB.CheckBox chkSearchForMedia 
         Appearance      =   0  'Flat
         Caption         =   "&Search for media"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   3120
         TabIndex        =   172
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkExclusiveToMp3InPlaylist 
         Appearance      =   0  'Flat
         Caption         =   "&Playlist exclusive to MP3's"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   3120
         TabIndex        =   171
         Top             =   360
         Width           =   2295
      End
      Begin VB.OptionButton optEngine 
         Appearance      =   0  'Flat
         Caption         =   "&Spectrum Analizer"
         ForeColor       =   &H80000008&
         Height          =   200
         Index           =   1
         Left            =   0
         TabIndex        =   149
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optEngine 
         Appearance      =   0  'Flat
         Caption         =   "&Nothing"
         ForeColor       =   &H80000008&
         Height          =   200
         Index           =   2
         Left            =   0
         TabIndex        =   150
         Top             =   600
         Width           =   2295
      End
      Begin MSComctlLib.Slider sldInitialTreble 
         Height          =   195
         Left            =   1440
         TabIndex        =   158
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   344
         _Version        =   393216
         Max             =   65535
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider sldInitialBass 
         Height          =   195
         Left            =   1440
         TabIndex        =   156
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   344
         _Version        =   393216
         Max             =   65535
         TickStyle       =   3
      End
      Begin VB.Label Label38 
         Caption         =   "&Audio Server"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   174
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label37 
         Caption         =   "&Other"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   170
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label36 
         Caption         =   "&Continuous Playback Mode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   167
         Top             =   3240
         Width           =   2775
      End
      Begin VB.Label lblAlwaysShow 
         Caption         =   "&Always Show"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   165
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label34 
         Caption         =   "&Initial Audio Values:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   152
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblVisuals 
         Caption         =   "&Visuals"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   148
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Frame fraSettings 
      BorderStyle     =   0  'None
      Height          =   4215
      Index           =   8
      Left            =   1320
      TabIndex        =   147
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
      Begin RichTextLib.RichTextBox txtHelpInformation 
         Height          =   3570
         Left            =   30
         TabIndex        =   146
         Top             =   495
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   6297
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         MousePointer    =   1
         Appearance      =   0
         TextRTF         =   $"frmOptions.frx":4541
      End
      Begin VB.ComboBox cboHelpTopic 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmOptions.frx":45BC
         Left            =   600
         List            =   "frmOptions.frx":45BE
         MousePointer    =   1  'Arrow
         Style           =   2  'Dropdown List
         TabIndex        =   145
         Top             =   120
         Width           =   5535
      End
      Begin VB.Label lblTopic 
         BackStyle       =   0  'Transparent
         Caption         =   "&Topic:"
         Height          =   255
         Left            =   0
         TabIndex        =   144
         Top             =   120
         Width           =   615
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00808080&
         Height          =   3615
         Left            =   15
         Top             =   480
         Width           =   6135
      End
   End
   Begin VB.Frame fraSettings 
      BorderStyle     =   0  'None
      Caption         =   "Ignore/Blacklist"
      Height          =   4215
      Index           =   7
      Left            =   1320
      TabIndex        =   143
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
      Begin nexIRC.ctlXPButton cmdAddToIgnore 
         Height          =   375
         Left            =   0
         TabIndex        =   224
         Top             =   3360
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "frmOptions.frx":45C0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ListBox lstBlacklist 
         Height          =   2595
         ItemData        =   "frmOptions.frx":45DC
         Left            =   3120
         List            =   "frmOptions.frx":45DE
         TabIndex        =   142
         Top             =   360
         Width           =   2895
      End
      Begin VB.CheckBox chkEnableIgnore 
         Appearance      =   0  'Flat
         Caption         =   "&Enable Ignore"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   0
         TabIndex        =   140
         Top             =   3000
         Width           =   1815
      End
      Begin VB.ListBox lstIgnore 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2595
         ItemData        =   "frmOptions.frx":45E0
         Left            =   0
         List            =   "frmOptions.frx":45E2
         TabIndex        =   139
         Top             =   360
         Width           =   2895
      End
      Begin nexIRC.ctlXPButton cmdRemoveIgnore 
         Height          =   375
         Left            =   0
         TabIndex        =   225
         Top             =   3840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Remove"
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
         MICON           =   "frmOptions.frx":45E4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdClearIgnore 
         Height          =   375
         Left            =   1200
         TabIndex        =   226
         Top             =   3840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Clear"
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
         MICON           =   "frmOptions.frx":4600
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdAddBlacklist 
         Height          =   375
         Left            =   3120
         TabIndex        =   227
         Top             =   3360
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "frmOptions.frx":461C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdRemoveBlacklist 
         Height          =   375
         Left            =   3120
         TabIndex        =   228
         Top             =   3840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Remove"
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
         MICON           =   "frmOptions.frx":4638
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdClearBlackList 
         Height          =   375
         Left            =   4320
         TabIndex        =   229
         Top             =   3840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Clear"
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
         MICON           =   "frmOptions.frx":4654
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblBlacklist 
         Caption         =   "&Blacklist:"
         Height          =   255
         Left            =   3120
         TabIndex        =   141
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label lblIgnoreList 
         BackStyle       =   0  'Transparent
         Caption         =   "&Ignore List:"
         Height          =   255
         Left            =   0
         TabIndex        =   138
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.Frame fraSettings 
      BorderStyle     =   0  'None
      Caption         =   "About"
      Height          =   4215
      Index           =   6
      Left            =   1320
      TabIndex        =   137
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label Label7 
         Caption         =   "http://www.team-nexgen.org"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   720
         MouseIcon       =   "frmOptions.frx":4670
         MousePointer    =   99  'Custom
         TabIndex        =   136
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Website:"
         Height          =   255
         Left            =   0
         TabIndex        =   135
         Top             =   3120
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   1095
         Left            =   0
         Picture         =   "frmOptions.frx":47C2
         Top             =   120
         Width           =   3525
      End
      Begin VB.Label Label2 
         Caption         =   "©1996-2012 Team Nexgen"
         Height          =   255
         Left            =   0
         TabIndex        =   130
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label lblAbout 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmOptions.frx":6067
         Height          =   855
         Left            =   0
         TabIndex        =   132
         Top             =   1920
         Width           =   6135
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "Version: 2.0"
         Height          =   255
         Left            =   0
         TabIndex        =   131
         Top             =   1560
         Width           =   6015
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Support:"
         Height          =   255
         Left            =   0
         TabIndex        =   133
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label lblEMail 
         BackStyle       =   0  'Transparent
         Caption         =   "guide_X@live.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   720
         MouseIcon       =   "frmOptions.frx":6187
         MousePointer    =   99  'Custom
         TabIndex        =   134
         Top             =   2880
         Width           =   2055
      End
   End
   Begin VB.Frame fraSettings 
      BorderStyle     =   0  'None
      Caption         =   "Text"
      Height          =   4215
      Index           =   5
      Left            =   1320
      TabIndex        =   182
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CheckBox chkTimeStamping 
         Appearance      =   0  'Flat
         Caption         =   "&Time Stamping"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   2400
         TabIndex        =   257
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CheckBox chkSaveChanges 
         Appearance      =   0  'Flat
         Caption         =   "Save Text Changes"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   249
         Top             =   3000
         Width           =   2055
      End
      Begin VB.ComboBox cboGroupBy 
         Height          =   315
         ItemData        =   "frmOptions.frx":62D9
         Left            =   2400
         List            =   "frmOptions.frx":62DB
         Style           =   2  'Dropdown List
         TabIndex        =   245
         Top             =   1080
         Width           =   3615
      End
      Begin VB.ComboBox cboPreset 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   128
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox txtString 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         MousePointer    =   1  'Arrow
         TabIndex        =   129
         Top             =   3840
         Width           =   6135
      End
      Begin VB.ListBox lstString 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   3660
         IntegralHeight  =   0   'False
         ItemData        =   "frmOptions.frx":62DD
         Left            =   0
         List            =   "frmOptions.frx":62DF
         Sorted          =   -1  'True
         TabIndex        =   126
         Top             =   120
         Width           =   2295
      End
      Begin nexIRC.ctlXPButton cmdPreviewString 
         Height          =   375
         Left            =   2400
         TabIndex        =   230
         Top             =   3360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Preview"
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
         MICON           =   "frmOptions.frx":62E1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         Caption         =   "Group:"
         Height          =   255
         Left            =   2400
         TabIndex        =   243
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label Label18 
         Caption         =   "&Preset:"
         Height          =   255
         Left            =   2400
         TabIndex        =   127
         Top             =   120
         Width           =   3495
      End
   End
   Begin VB.Frame fraSettings 
      BorderStyle     =   0  'None
      Caption         =   "Themes"
      Height          =   4215
      Index           =   4
      Left            =   1320
      TabIndex        =   125
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CheckBox chkColoredNicklist 
         Appearance      =   0  'Flat
         Caption         =   "&Colored Nicklist"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   259
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CheckBox chkShowExtraProgress 
         Appearance      =   0  'Flat
         Caption         =   "Show Extra Progress"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   258
         Top             =   3450
         Width           =   1815
      End
      Begin MSComctlLib.ImageCombo cboProgressBarColor 
         Height          =   330
         Left            =   4200
         TabIndex        =   256
         Top             =   2400
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         ImageList       =   "ImageList1"
      End
      Begin VB.ComboBox cboProgressBar 
         Height          =   315
         ItemData        =   "frmOptions.frx":62FD
         Left            =   4200
         List            =   "frmOptions.frx":62FF
         Style           =   2  'Dropdown List
         TabIndex        =   253
         Top             =   2040
         Width           =   1935
      End
      Begin VB.PictureBox picBGColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   2160
         Left            =   120
         ScaleHeight     =   2160
         ScaleWidth      =   2880
         TabIndex        =   78
         Top             =   600
         Width           =   2880
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Server"
            Height          =   195
            Index           =   15
            Left            =   1440
            MouseIcon       =   "frmOptions.frx":6301
            MousePointer    =   99  'Custom
            TabIndex        =   87
            Top             =   120
            Width           =   480
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Whois"
            Height          =   195
            Index           =   14
            Left            =   1440
            MouseIcon       =   "frmOptions.frx":6453
            MousePointer    =   99  'Custom
            TabIndex        =   88
            Top             =   360
            Width           =   435
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Topics"
            Height          =   195
            Index           =   13
            Left            =   1440
            MouseIcon       =   "frmOptions.frx":65A5
            MousePointer    =   99  'Custom
            TabIndex        =   89
            Top             =   600
            Width           =   450
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Quit Server"
            Height          =   195
            Index           =   12
            Left            =   1440
            MouseIcon       =   "frmOptions.frx":66F7
            MousePointer    =   99  'Custom
            TabIndex        =   90
            Top             =   840
            Width           =   825
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Channel Part"
            Height          =   195
            Index           =   11
            Left            =   1440
            MouseIcon       =   "frmOptions.frx":6849
            MousePointer    =   99  'Custom
            TabIndex        =   91
            Top             =   1080
            Width           =   930
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Notify"
            Height          =   195
            Index           =   10
            Left            =   1440
            MouseIcon       =   "frmOptions.frx":699B
            MousePointer    =   99  'Custom
            TabIndex        =   92
            Top             =   1320
            Width           =   435
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Nickname Changes"
            Height          =   195
            Index           =   9
            Left            =   1440
            MouseIcon       =   "frmOptions.frx":6AED
            MousePointer    =   99  'Custom
            TabIndex        =   93
            Top             =   1560
            Width           =   1350
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Modes Changes"
            Height          =   195
            Index           =   8
            Left            =   1440
            MouseIcon       =   "frmOptions.frx":6C3F
            MousePointer    =   99  'Custom
            TabIndex        =   94
            Top             =   1800
            Width           =   1140
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Channel Kick"
            Height          =   195
            Index           =   7
            Left            =   120
            MouseIcon       =   "frmOptions.frx":6D91
            MousePointer    =   99  'Custom
            TabIndex        =   86
            Top             =   1800
            Width           =   900
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Join Channel"
            Height          =   195
            Index           =   6
            Left            =   120
            MouseIcon       =   "frmOptions.frx":6EE3
            MousePointer    =   99  'Custom
            TabIndex        =   85
            Top             =   1560
            Width           =   915
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Invite to Channel"
            Height          =   195
            Index           =   5
            Left            =   120
            MouseIcon       =   "frmOptions.frx":7035
            MousePointer    =   99  'Custom
            TabIndex        =   84
            Top             =   1320
            Width           =   1245
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Action"
            Height          =   195
            Index           =   4
            Left            =   120
            MouseIcon       =   "frmOptions.frx":7187
            MousePointer    =   99  'Custom
            TabIndex        =   83
            Top             =   1080
            Width           =   450
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Notice"
            Height          =   195
            Index           =   3
            Left            =   120
            MouseIcon       =   "frmOptions.frx":72D9
            MousePointer    =   99  'Custom
            TabIndex        =   82
            Top             =   840
            Width           =   450
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Ctcp"
            Height          =   195
            Index           =   2
            Left            =   120
            MouseIcon       =   "frmOptions.frx":742B
            MousePointer    =   99  'Custom
            TabIndex        =   81
            Top             =   600
            Width           =   330
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Background"
            Height          =   195
            Index           =   0
            Left            =   120
            MouseIcon       =   "frmOptions.frx":757D
            MousePointer    =   99  'Custom
            TabIndex        =   79
            Top             =   120
            Width           =   840
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Normal Text"
            Height          =   195
            Index           =   1
            Left            =   120
            MouseIcon       =   "frmOptions.frx":76CF
            MousePointer    =   99  'Custom
            TabIndex        =   80
            Top             =   360
            Width           =   870
         End
      End
      Begin VB.CheckBox chkBorderlessObjects 
         Appearance      =   0  'Flat
         Caption         =   "&Borderless Objects"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3120
         TabIndex        =   105
         Top             =   3690
         Width           =   2295
      End
      Begin VB.CheckBox chkAlwaysShowAudioSettings 
         Appearance      =   0  'Flat
         Caption         =   "&Always show Theme Settings"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   3120
         TabIndex        =   104
         Top             =   3480
         Width           =   2535
      End
      Begin VB.ComboBox cboButtonType 
         Height          =   315
         ItemData        =   "frmOptions.frx":7821
         Left            =   4200
         List            =   "frmOptions.frx":7823
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CheckBox chkSaveColorsToTheme 
         Appearance      =   0  'Flat
         Caption         =   "&Save IRC Colors to Theme"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         TabIndex        =   123
         Top             =   3720
         Width           =   2295
      End
      Begin VB.CheckBox chkApplyThemeToIRCColors 
         Appearance      =   0  'Flat
         Caption         =   "&Apply Theme to IRC Colors"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         TabIndex        =   124
         Top             =   3960
         Width           =   2295
      End
      Begin VB.CheckBox chkRefreshPictureColors 
         Appearance      =   0  'Flat
         Caption         =   "&Apply Theme to Toolbar colors"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   3120
         TabIndex        =   103
         Top             =   3960
         Width           =   2535
      End
      Begin VB.ComboBox cboColorTheme 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   120
         Width           =   5415
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   5520
         MouseIcon       =   "frmOptions.frx":7825
         MousePointer    =   99  'Custom
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   121
         Top             =   2880
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   5160
         MouseIcon       =   "frmOptions.frx":7977
         MousePointer    =   99  'Custom
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   120
         Top             =   2880
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   4800
         MouseIcon       =   "frmOptions.frx":7AC9
         MousePointer    =   99  'Custom
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   119
         Top             =   2880
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   4440
         MouseIcon       =   "frmOptions.frx":7C1B
         MousePointer    =   99  'Custom
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   118
         Top             =   2880
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   4080
         MouseIcon       =   "frmOptions.frx":7D6D
         MousePointer    =   99  'Custom
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   117
         Top             =   2880
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   3720
         MouseIcon       =   "frmOptions.frx":7EBF
         MousePointer    =   99  'Custom
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   116
         Top             =   2880
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   3360
         MouseIcon       =   "frmOptions.frx":8011
         MousePointer    =   99  'Custom
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   115
         Top             =   2880
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   3000
         MouseIcon       =   "frmOptions.frx":8163
         MousePointer    =   99  'Custom
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   114
         Top             =   2880
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   2640
         MouseIcon       =   "frmOptions.frx":82B5
         MousePointer    =   99  'Custom
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   113
         Top             =   2880
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   2280
         MouseIcon       =   "frmOptions.frx":8407
         MousePointer    =   99  'Custom
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   112
         Top             =   2880
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   1920
         MouseIcon       =   "frmOptions.frx":8559
         MousePointer    =   99  'Custom
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   111
         Top             =   2880
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1560
         MouseIcon       =   "frmOptions.frx":86AB
         MousePointer    =   99  'Custom
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   110
         Top             =   2880
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1200
         MouseIcon       =   "frmOptions.frx":87FD
         MousePointer    =   99  'Custom
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   109
         Top             =   2880
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   840
         MouseIcon       =   "frmOptions.frx":894F
         MousePointer    =   99  'Custom
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   108
         Top             =   2880
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   480
         MouseIcon       =   "frmOptions.frx":8AA1
         MousePointer    =   99  'Custom
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   107
         Top             =   2880
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         MouseIcon       =   "frmOptions.frx":8BF3
         MousePointer    =   99  'Custom
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   106
         Top             =   2880
         Width           =   255
      End
      Begin MSComctlLib.ImageCombo cboOpColor 
         Height          =   330
         Left            =   4200
         TabIndex        =   96
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         MousePointer    =   1
         ImageList       =   "ImageList1"
      End
      Begin MSComctlLib.ImageCombo cboVoiceColor 
         Height          =   330
         Left            =   4200
         TabIndex        =   98
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         MousePointer    =   1
         ImageList       =   "ImageList1"
      End
      Begin MSComctlLib.ImageCombo cboNormalColor 
         Height          =   330
         Left            =   4200
         TabIndex        =   100
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         MousePointer    =   1
         ImageList       =   "ImageList1"
      End
      Begin VB.Label Label5 
         Caption         =   "Scroll Color"
         Height          =   255
         Left            =   3120
         TabIndex        =   255
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lblProgressBarStyle 
         Caption         =   "Scroll Style:"
         Height          =   255
         Left            =   3120
         TabIndex        =   254
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblOpsColor 
         Caption         =   "&Ops Color:"
         Height          =   255
         Left            =   3120
         TabIndex        =   95
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "&Voice Color:"
         Height          =   255
         Left            =   3120
         TabIndex        =   97
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label41 
         Caption         =   "&Normal Color:"
         Height          =   255
         Left            =   3120
         TabIndex        =   99
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "&Button Type:"
         Height          =   255
         Left            =   3120
         TabIndex        =   101
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00808080&
         Height          =   2205
         Left            =   105
         Top             =   570
         Width           =   2910
      End
      Begin VB.Label Label8 
         Caption         =   "Theme:"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   160
         Width           =   1455
      End
      Begin VB.Label lblExample 
         Height          =   255
         Left            =   3120
         TabIndex        =   122
         Top             =   3120
         Visible         =   0   'False
         Width           =   3015
      End
   End
   Begin VB.Frame fraSettings 
      BorderStyle     =   0  'None
      Caption         =   "Nortify"
      Height          =   4215
      Index           =   3
      Left            =   1320
      TabIndex        =   75
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
      Begin VB.ListBox lstNotify 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   3960
         ItemData        =   "frmOptions.frx":8D45
         Left            =   0
         List            =   "frmOptions.frx":8D47
         TabIndex        =   67
         Top             =   120
         Width           =   1815
      End
      Begin VB.CheckBox chkShowQuickNotify 
         Appearance      =   0  'Flat
         Caption         =   "Show &quick notify"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   1920
         TabIndex        =   74
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CheckBox chkShowNotifyWindow 
         Appearance      =   0  'Flat
         Caption         =   "Show &Notify window"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   1920
         TabIndex        =   73
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CheckBox chkWhois 
         Appearance      =   0  'Flat
         Caption         =   "Show /WHOIS on Query"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   1920
         TabIndex        =   72
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CheckBox chkNotifyOnActive 
         Appearance      =   0  'Flat
         Caption         =   "&Show Notify in Active Window"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   1920
         TabIndex        =   71
         Top             =   1440
         Width           =   3975
      End
      Begin VB.TextBox txtNotifyNickName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1920
         MousePointer    =   1  'Arrow
         TabIndex        =   68
         Top             =   120
         Width           =   4140
      End
      Begin VB.CheckBox chkEnable 
         Appearance      =   0  'Flat
         Caption         =   "&Enable Notify"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   1920
         TabIndex        =   69
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox chkWhoisNotify 
         Appearance      =   0  'Flat
         Caption         =   "&Preform /WHOIS on Notify"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   1920
         TabIndex        =   70
         Top             =   1200
         Width           =   3975
      End
      Begin nexIRC.ctlXPButton cmdAdd 
         Height          =   375
         Left            =   1920
         TabIndex        =   231
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "frmOptions.frx":8D49
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
         Left            =   3120
         TabIndex        =   232
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "frmOptions.frx":8D65
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdClear 
         Height          =   375
         Left            =   4320
         TabIndex        =   233
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Clear"
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
         MICON           =   "frmOptions.frx":8D81
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame fraSettings 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Options"
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   2
      Left            =   1320
      TabIndex        =   66
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CheckBox chkBalloons 
         Appearance      =   0  'Flat
         Caption         =   "&Balloons"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4200
         TabIndex        =   247
         Top             =   2520
         Width           =   975
      End
      Begin VB.CheckBox chkShowModes 
         Appearance      =   0  'Flat
         Caption         =   "&Show Modes"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   33
         Top             =   3480
         Width           =   1935
      End
      Begin VB.CheckBox chkShowMe 
         Appearance      =   0  'Flat
         Caption         =   "&Show Options"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   0
         TabIndex        =   27
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CheckBox chkShowWhoisInChannel 
         Appearance      =   0  'Flat
         Caption         =   "&Show Whois"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   35
         Top             =   3960
         Width           =   1815
      End
      Begin VB.CheckBox chkAutoPortScanner 
         Appearance      =   0  'Flat
         Caption         =   "&Auto Port Scanner"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4200
         TabIndex        =   60
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox chkByPassStartupScreen 
         Appearance      =   0  'Flat
         Caption         =   "&Auto Close Splash"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   0
         TabIndex        =   22
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox chkUpdateCheck 
         Appearance      =   0  'Flat
         Caption         =   "&Update Check"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   0
         TabIndex        =   21
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkDownloadManager 
         Appearance      =   0  'Flat
         Caption         =   "&Download Manager"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   4200
         TabIndex        =   59
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CheckBox chkShowTips 
         Appearance      =   0  'Flat
         Caption         =   "&Show Tips"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   23
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox chkAutoSizeStatusbarItems 
         Appearance      =   0  'Flat
         Caption         =   "&Autosize Taskbar"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4200
         TabIndex        =   54
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chkNickCompletor 
         Appearance      =   0  'Flat
         Caption         =   "&Nick Completor"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2040
         TabIndex        =   50
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CheckBox chkSecureQuery 
         Appearance      =   0  'Flat
         Caption         =   "&Secure Query"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4200
         TabIndex        =   55
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox chkWallops 
         Appearance      =   0  'Flat
         Caption         =   "&Op"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2040
         TabIndex        =   39
         Top             =   840
         Width           =   615
      End
      Begin VB.CheckBox chkServerMSG 
         Appearance      =   0  'Flat
         Caption         =   "&Server"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2040
         TabIndex        =   38
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox chkInvisible 
         Appearance      =   0  'Flat
         Caption         =   "&Invisible Mode"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2040
         TabIndex        =   37
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox chkShowMOTD 
         Appearance      =   0  'Flat
         Caption         =   "&Own Window"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2040
         TabIndex        =   42
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CheckBox chkShowAddress 
         Appearance      =   0  'Flat
         Caption         =   "&Show Address"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   29
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CheckBox chkSkipMOTD 
         Appearance      =   0  'Flat
         Caption         =   "&Skip"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2040
         TabIndex        =   41
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox chkRejoin 
         Appearance      =   0  'Flat
         Caption         =   "&Rejoin when Kicked"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2040
         TabIndex        =   52
         Top             =   3720
         Width           =   1815
      End
      Begin VB.CheckBox chkAutoJoin 
         Appearance      =   0  'Flat
         Caption         =   "&Join on Invite"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2040
         TabIndex        =   49
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CheckBox chkShowQuits 
         Appearance      =   0  'Flat
         Caption         =   "&Show Quits"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   30
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CheckBox chkShowTopics 
         Appearance      =   0  'Flat
         Caption         =   "&Show Topic Changes"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   34
         Top             =   3720
         Width           =   1935
      End
      Begin VB.CheckBox chkShowKicks 
         Appearance      =   0  'Flat
         Caption         =   "&Show Kicks"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   31
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CheckBox chkConnectOnStartup 
         Appearance      =   0  'Flat
         Caption         =   "&Connect"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   26
         Top             =   1800
         Width           =   975
      End
      Begin VB.CheckBox chkShowSplashScreenOnStartup 
         Appearance      =   0  'Flat
         Caption         =   "&Show Splash"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox chkAddJoinedChannelsToChannelFolder 
         Appearance      =   0  'Flat
         Caption         =   "&Use Channel Folder"
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2040
         TabIndex        =   48
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CheckBox chkAutoJoinEnabled 
         Appearance      =   0  'Flat
         Caption         =   "&Autojoin"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2040
         TabIndex        =   51
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CheckBox chkDCCPrompts 
         Appearance      =   0  'Flat
         Caption         =   "&DCC"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4200
         TabIndex        =   63
         Top             =   2760
         Width           =   735
      End
      Begin VB.CheckBox chkGeneralPrompts 
         Appearance      =   0  'Flat
         Caption         =   "&General"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4200
         TabIndex        =   62
         Top             =   2280
         Width           =   975
      End
      Begin VB.CheckBox chkDCCEnabled 
         Appearance      =   0  'Flat
         Caption         =   "&DCC Enabled"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4200
         TabIndex        =   57
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox chkNavigateOnStartup 
         Appearance      =   0  'Flat
         Caption         =   "&Show Homepage"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   25
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CheckBox chkShowServerOnStartup 
         Appearance      =   0  'Flat
         Caption         =   "&Show Server"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   24
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox chkAutoSelectAlternateNickname 
         Appearance      =   0  'Flat
         Caption         =   "&Select Alternates"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4200
         TabIndex        =   56
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox chkShowJoinPart 
         Appearance      =   0  'Flat
         Caption         =   "&Show Join's/Part's"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   32
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CheckBox chkReconnectOnDisconnect 
         Appearance      =   0  'Flat
         Caption         =   "&Auto Reconnect"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4200
         TabIndex        =   58
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtQuitMessage 
         Height          =   250
         Left            =   4680
         MousePointer    =   1  'Arrow
         TabIndex        =   65
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CheckBox chkBackgroundWebpage 
         Appearance      =   0  'Flat
         Caption         =   "&In Background"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2040
         TabIndex        =   44
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtHomepage 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   250
         Left            =   3000
         MousePointer    =   1  'Arrow
         TabIndex        =   46
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblChannels 
         Caption         =   "&Channels"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   47
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblStartup 
         Caption         =   "&Startup"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblShowInChannel 
         Caption         =   "S&how in Channel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label lblMessages 
         Caption         =   "&Messages"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   36
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblPrompts 
         Caption         =   "&Prompts"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   61
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblGeneral 
         Caption         =   "&General"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   53
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblMOTD 
         Caption         =   "M&OTD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   40
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label39 
         Caption         =   "&Quit:"
         Height          =   255
         Left            =   4200
         TabIndex        =   64
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label lblHomepage 
         Caption         =   "&Homepage:"
         Height          =   255
         Left            =   2040
         TabIndex        =   45
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblBrowser 
         Caption         =   "&Browser"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   43
         Top             =   1800
         Width           =   975
      End
   End
   Begin VB.Frame fraSettings 
      BorderStyle     =   0  'None
      Caption         =   "User"
      Height          =   4215
      Index           =   1
      Left            =   1320
      TabIndex        =   192
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox txtNickname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   250
         Left            =   1220
         MousePointer    =   1  'Arrow
         TabIndex        =   193
         Top             =   120
         Width           =   4900
      End
      Begin VB.CheckBox chkIdentShow 
         Appearance      =   0  'Flat
         Caption         =   "&Show Identd Requests"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   1080
         TabIndex        =   205
         Top             =   3960
         Width           =   2055
      End
      Begin VB.TextBox txtIdentUserID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   250
         Left            =   1215
         MousePointer    =   1  'Arrow
         TabIndex        =   196
         Top             =   960
         Width           =   4900
      End
      Begin VB.TextBox txtIdentSystem 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   250
         Left            =   1215
         MousePointer    =   1  'Arrow
         TabIndex        =   197
         Top             =   1215
         Width           =   4900
      End
      Begin VB.TextBox txtIdentPort 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   250
         Left            =   1215
         MousePointer    =   1  'Arrow
         TabIndex        =   198
         Top             =   1470
         Width           =   4900
      End
      Begin VB.CheckBox chkIdent 
         Appearance      =   0  'Flat
         Caption         =   "&Enable Identd Server"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   1080
         TabIndex        =   203
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox txtRealName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   250
         Left            =   1215
         MousePointer    =   1  'Arrow
         TabIndex        =   195
         Top             =   630
         Width           =   4900
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   250
         Left            =   1215
         MousePointer    =   1  'Arrow
         TabIndex        =   194
         Top             =   375
         Width           =   4900
      End
      Begin VB.ComboBox cboAlternates 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   199
         Top             =   1845
         Width           =   4935
      End
      Begin nexIRC.ctlXPButton cmdAddAlternate 
         Height          =   375
         Left            =   2640
         TabIndex        =   200
         Top             =   2280
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "frmOptions.frx":8D9D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdRemoveAlternate 
         Height          =   375
         Left            =   3840
         TabIndex        =   201
         Top             =   2280
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Remove"
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
         MICON           =   "frmOptions.frx":8DB9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdClearAlternates 
         Height          =   375
         Left            =   5040
         TabIndex        =   202
         Top             =   2280
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Clear"
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
         MICON           =   "frmOptions.frx":8DD5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin nexIRC.ctlXPButton cmdDefaultIdent 
         Height          =   375
         Left            =   5040
         TabIndex        =   206
         Top             =   3720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Default"
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
         MICON           =   "frmOptions.frx":8DF1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nickname:"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   212
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "&E-Mail:"
         Height          =   255
         Left            =   0
         TabIndex        =   211
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "&Real Name:"
         Height          =   255
         Left            =   0
         TabIndex        =   210
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "&UserID:"
         Height          =   255
         Left            =   0
         TabIndex        =   209
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "&System:"
         Height          =   255
         Left            =   0
         TabIndex        =   208
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "&Port:"
         Height          =   255
         Left            =   0
         TabIndex        =   207
         Top             =   1440
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00404040&
         BorderColor     =   &H00808080&
         Height          =   810
         Left            =   1200
         Top             =   945
         Width           =   4950
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00404040&
         BorderColor     =   &H00808080&
         Height          =   825
         Left            =   1200
         Top             =   90
         Width           =   4950
      End
      Begin VB.Label lblAlternates 
         Caption         =   "&Nicknames:"
         Height          =   255
         Left            =   0
         TabIndex        =   204
         Top             =   1845
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCustomize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lToolTip As clsToolTip
Private lApply As Boolean
Private lDoneLoading As Boolean

Public Sub CreateBalloon(lTitle As String, lText As String, lObject As Object)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sBalloons = True Then
    lToolTip.TipText = lText
    Set lToolTip.ParentControl = lObject
    lToolTip.Title = lTitle
    lToolTip.Create
End If
End Sub

Public Sub ActiveateSaveSettings()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, F As Integer, srColor As String, X As Integer, sTheme As Integer, m As Integer
mdiNexIRC.tmrNotify.Enabled = GetCheckboxValue(chkEnable)
SaveBlacklist
lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nOpColor = cboOpColor.SelectedItem.Index
lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nVoiceColor = cboVoiceColor.SelectedItem.Index
lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nNormalColor = cboNormalColor.SelectedItem.Index
lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sProgressBarStyle = cboProgressBar.ListIndex
lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sProgressBarColor = cboProgressBarColor.SelectedItem.Index - 1
For i = 1 To ReturnChannelUBound
    If Len(ReturnChannelName(i)) <> 0 Then
        'SortNicklist ReturnChannelNamesListView(i)
    End If
Next i
SaveNicklistOptions lSpectrumThemes.sIndex
SetStringData sQuitReason, txtQuitMessage.Text
SetIgnoreEnabled GetCheckboxValue(chkEnableIgnore)
lSettings.sAlwaysShowAudioSettings = GetCheckboxValue(chkAlwaysShowAudioSettings)
lSettings.sLogoTwitchOnPeaks = GetCheckboxValue(chkLogoTwitchOnPeaks)
sTheme = FindSpectrumThemeByName(cboColorTheme.Text)
If sTheme <> 0 Then
    If sTheme <> lSpectrumThemes.sIndex Then
        lSpectrumThemes.sIndex = sTheme
        ApplySpectrumTheme lSpectrumThemes.sSpectrumTheme(sTheme).sName
    End If
End If
Select Case lPlayback.pCurrentEngine
Case pMp3
    If lSettings.sAlwaysShowAudioSettings = True Then
        mdiNexIRC.picMP3OCX.Visible = True
    Else
        mdiNexIRC.picMP3OCX.Visible = False
    End If
End Select
lSettings.sDownloadManager = GetCheckboxValue(chkDownloadManager)
lSettings.sShowTips = GetCheckboxValue(chkShowTips)
lSettings.sButtonType = cboButtonType.ListIndex
If cmdSmallNetworks.Value = True Then
    lSettings.sShowSmallNetworks = False
Else
    lSettings.sShowSmallNetworks = True
End If
lSettings.sBalloons = GetCheckboxValue(chkBalloons)
lSettings.sSearchForMedia = GetCheckboxValue(chkSearchForMedia)
lSettings.sUpdateCheck = GetCheckboxValue(chkUpdateCheck)
lSettings.sPlaySounds = GetCheckboxValue(chkPlaySounds)
lSettings.sReconnectOnDisconnect = GetCheckboxValue(chkReconnectOnDisconnect)
lSettings.sRefreshPictureColors = GetCheckboxValue(chkRefreshPictureColors)
lSettings.sShowServerOnStartup = GetCheckboxValue(chkShowServerOnStartup)
lSettings.sDCCEnabled = GetCheckboxValue(chkDCCEnabled)
lSettings.sExlusiveToMp3InPlaylist = GetCheckboxValue(chkExclusiveToMp3InPlaylist)
lSettings.sContinuousPlay = GetCheckboxValue(chkContinuousPlay)
lSettings.sShuffle = GetCheckboxValue(chkShuffle)
lSettings.sServerMinimum = cboServerMinimum.Text
frmMobileMixer.chkContinuous.Value = chkContinuousPlay.Value
frmMobileMixer.chkShuffle.Value = chkShuffle.Value
If lSettings.sBackgroundWebpage = False And chkBackgroundWebpage.Value = 1 Then
    frmWeb.Visible = True
ElseIf lSettings.sBackgroundWebpage = True And chkBackgroundWebpage.Value = 0 Then
    frmWeb.Visible = False
End If
'If optEngine(1).Value = True Then
'    SwitchPlaybackEngine pMp3
'ElseIf optEngine(2).Value = True Then
'    SwitchPlaybackEngine pMediaPlayer
'End If
lSettings.sShowNotifyWindow = GetCheckboxValue(chkShowNotifyWindow)
lSettings.sShowQuickNotify = GetCheckboxValue(chkShowQuickNotify)
If lSettings.sShowQuickNotify = True Then
    mdiNexIRC.picNotify.Visible = True
Else
    mdiNexIRC.picNotify.Visible = False
End If
With lInitialAudioValues
    .iBass = sldInitialBass.Value
    .iWave = sldInitialWave.Value
    .iTreble = sldInitialTreble.Value
    .iLineIN = sldInitialLineIN.Value
    .iMic = sldInitialMic.Value
    .iCDAudio = sldInitialCDAudio.Value
    If chkInitialBass.Value = 1 Then .iInitialBassEnabled = True
    If chkInitialTreble.Value = 1 Then .iInitialTrebleEnabled = True
    If chkInitialCDAudio.Value = 1 Then .iInitialCDAudioEnabled = True
    If chkInitialLineIN.Value = 1 Then .iInitialLineInEnabled = True
    If chkInitialMIC.Value = 1 Then .iInitialMicEnabled = True
    If chkInitialWave.Value = 1 Then .iInitialWaveEnabled = True
End With
lSettings.sAutoPortScanner = GetCheckboxValue(chkAutoPortScanner)
lSettings.sColoredNicklist = GetCheckboxValue(chkColoredNicklist)
lSettings.sShowWhoisInChannel = GetCheckboxValue(chkShowWhoisInChannel)
lSettings.sSecureQuery = GetCheckboxValue(chkSecureQuery)
lSettings.sUseNickCompletor = GetCheckboxValue(chkNickCompletor)
lSettings.sAutosizeStatusbarItems = GetCheckboxValue(chkAutoSizeStatusbarItems)
lSettings.sAudioServer = GetCheckboxValue(chkAudioServer)
lSettings.sFileOfferInChannel = GetCheckboxValue(chkFileOfferInChannel)
lSettings.sOfferWhenPlayed = GetCheckboxValue(chkOfferWhenPlayed)
lSettings.sBorderlessObjects = GetCheckboxValue(chkBorderlessObjects)
lSettings.sEnableSearch = GetCheckboxValue(chkEnableFind)
lSettings.sEnableList = GetCheckboxValue(chkEnableListmedia)
lSettings.sTimeStamping = GetCheckboxValue(chkTimeStamping)
lSettings.sAutoSelectAlternateNickname = GetCheckboxValue(chkAutoSelectAlternateNickname)
lSettings.sBackgroundWebpage = GetCheckboxValue(chkBackgroundWebpage)
lSettings.sNavigateOnStartup = GetCheckboxValue(chkNavigateOnStartup)
lSettings.sAutoJoinEnabled = GetCheckboxValue(chkAutoJoinEnabled)
lSettings.sShowQuickmix = GetCheckboxValue(chkShowQuickMix)
If lSettings.sShowQuickmix = True Then
    ToggleMixer True
Else
    ToggleMixer False
End If
lSettings.sAddJoinedChannelsToChannelFolder = GetCheckboxValue(chkAddJoinedChannelsToChannelFolder)
lSettings.sDCCPrompts = GetCheckboxValue(chkDCCPrompts)
lSettings.sGeneralPrompts = GetCheckboxValue(chkGeneralPrompts)
lSettings.sSaveIRCColorsToTheme = GetCheckboxValue(chkSaveColorsToTheme)
lSettings.sNetwork = cmbNetwork.Text
lSettings.sServer = txtServer.Text
lSettings.sPort = txtPort.Text
lSettings.sPassword = txtPassword.Text
lSettings.sNickname = txtNickname.Text
lSettings.sEMail = txtEmail.Text
lSettings.sRealName = txtRealName.Text
lSettings.sIdent.iUserID = txtIdentUserID.Text
lSettings.sIdent.iSystem = txtIdentSystem.Text
lSettings.sIdent.iPort = txtIdentPort.Text
lSettings.sIdent.iEnabled = GetCheckboxValue(chkIdent)
lSettings.sIdent.iShow = GetCheckboxValue(chkIdentShow)
lSettings.sOptions.oWhois = GetCheckboxValue(chkWhois)
lSettings.sAutoJoinOnInvite = GetCheckboxValue(chkAutoJoin)
lSettings.sOptions.oReJoin = GetCheckboxValue(chkRejoin)
lSettings.sOptions.oSkipMOTD = GetCheckboxValue(chkSkipMOTD)
lSettings.sOptions.oShowMOTD = GetCheckboxValue(chkShowMOTD)
lSettings.sOptions.oShowAddress = GetCheckboxValue(chkShowAddress)
lSettings.sOptions.oInvisable = GetCheckboxValue(chkInvisible)
lSettings.sOptions.oServerMessages = GetCheckboxValue(chkServerMSG)
lSettings.sOptions.oOpMessages = GetCheckboxValue(chkWallops)
lSettings.sOptions.oShowQuit = GetCheckboxValue(chkShowQuits)
lSettings.sOptions.oShowJoinPart = GetCheckboxValue(chkShowJoinPart)
lSettings.sOptions.oShowModes = GetCheckboxValue(chkShowModes)
lSettings.sOptions.oShowTopics = GetCheckboxValue(chkShowTopics)
lSettings.sOptions.oShowKicks = GetCheckboxValue(chkShowKicks)
lSettings.sApplyThemeToIRCColors = GetCheckboxValue(chkApplyThemeToIRCColors)
lSettings.sShowOptionsOnStartup = GetCheckboxValue(chkShowMe)
lSettings.sConnectOnStartup = GetCheckboxValue(chkConnectOnStartup)
lSettings.sShowSplashOnStartup = GetCheckboxValue(chkShowSplashScreenOnStartup)
lSettings.sByPassStartupScreen = GetCheckboxValue(chkByPassStartupScreen)
lSettings.sHomepage = txtHomepage.Text
SetNotifyToListBox lstNotify
lSettings.sOptions.oShowNotifyInActiveWindow = GetCheckboxValue(chkNotifyOnActive)
SetNotifyEnabled GetCheckboxValue(chkEnable)
lSettings.sOptions.oWhoisNotify = GetCheckboxValue(chkWhoisNotify)
For X = 0 To 15
    If Len(lblcolor(X).Tag) = 1 Then lblcolor(X).Tag = "0" & lblcolor(X).Tag
    srColor = srColor & X & ":" & lblcolor(X).Tag & " "
Next X
ApplyIRCColors Trim(srColor), False, lSettings.sSaveIRCColorsToTheme
mdiNexIRC.UpdateMainButtonTypes
ClearIgnore
SaveListBoxToIgnore lstIgnore
mdiNexIRC.ActivateResize
If lApply = True Then Me.SetFocus
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ActiveateSaveSettings()"
End Sub

Public Sub ShowAbout()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To 8
    fraSettings(i).Visible = False
    optCheck(i).Value = False
Next i
fraSettings(6).Visible = True
optCheck(6).Value = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ShowAbout()"
End Sub

Private Sub cboBotTypes_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
cboBots.Clear
cboCommands.Clear
For i = 0 To ReturnBotCount
    If Len(ReturnBotNickname(i)) <> 0 Then
        If ReturnBotType(i) = cboBotTypes.ListIndex Then
            cboBots.AddItem ReturnBotNickname(i)
        End If
    End If
Next i
On Local Error Resume Next
If ReturnBotCommandCount <> 0 Then cboBots.ListIndex = 0
For i = 0 To ReturnBotCommandCount
    If Len(ReturnBotCommand(i)) <> 0 Then
        If ReturnBotCommandType(i) = cboBotTypes.ListIndex Then
            cboCommands.AddItem ReturnBotCommand(i)
        End If
    End If
Next i
If cboCommands.ListCount <> 0 Then
    cboCommands.ListIndex = 0
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cboBotTypes_Click()"
End Sub

Private Sub cboColorTheme_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim m As Integer, i As Integer, Color(0 To 15) As Long, getColors() As String, TagColor() As String, o As Integer, srColor As String
If lSettings.sApplyThemeToIRCColors = False Then Exit Sub
m = FindSpectrumThemeByName(cboColorTheme.Text)
If Len(lSpectrumThemes.sSpectrumTheme(m).sIRCColors) <> 0 Then
    cboOpColor.ComboItems(Int(lSpectrumThemes.sSpectrumTheme(m).sNicklistOptions.nOpColor)).Selected = True
    cboVoiceColor.ComboItems(Int(lSpectrumThemes.sSpectrumTheme(m).sNicklistOptions.nVoiceColor)).Selected = True
    cboNormalColor.ComboItems(Int(lSpectrumThemes.sSpectrumTheme(m).sNicklistOptions.nNormalColor)).Selected = True
    Color(0) = vbWhite
    Color(1) = vbBlack
    Color(2) = RGB(42, 42, 87)
    Color(3) = RGB(33, 112, 33)
    Color(4) = vbRed
    Color(5) = RGB(109, 50, 50)
    Color(6) = RGB(119, 33, 119)
    Color(7) = RGB(252, 127, 0)
    Color(8) = RGB(195, 195, 56)
    Color(9) = RGB(0, 252, 0)
    Color(10) = RGB(89, 167, 179)
    Color(11) = RGB(0, 255, 255)
    Color(12) = vbBlue
    Color(13) = RGB(255, 0, 255)
    Color(14) = RGB(127, 127, 127)
    Color(15) = RGB(210, 210, 210)
    For i = o To 15
        picColor(i).BackColor = Color(i)
    Next i
    srColor = Trim(lSpectrumThemes.sSpectrumTheme(m).sIRCColors)
    getColors = Split(srColor, " ")
    For i = 0 To UBound(getColors)
        TagColor = Split(getColors(i), ":")
        lblcolor(TagColor(0)).Tag = TagColor(1)
        If Len(TagColor(1)) <> 0 Then lblcolor(TagColor(0)).ForeColor = Color(TagColor(1))
        If i = 0 Then
            picBGColor.BackColor = Color(TagColor(1))
        End If
    Next i
    lblcolor(0).ForeColor = lblcolor(1).ForeColor
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cboColorTheme_Click()"
End Sub

Private Sub cboGroupBy_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'Stop
FillListBoxWithTextDescriptionsGroup lstString, Int(cboGroupBy.ListIndex + 1)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cboGroupBy_Change()"
End Sub

Private Sub cboHelpTopic_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtHelpInformation.LoadFile App.Path & "\data\help\" & Trim(Str(cboHelpTopic.ListIndex)) & ".rtf"
txtHelpInformation.Refresh
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cboHelpTopic_Click()"
End Sub

Private Sub cboPreset_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String
If lDoneLoading = True Then
    msg = App.Path & "\data\config\fixed\text.ini"
    If DoesFileExist(msg) = True Then
        i = ReturnTextPresetIndexByDescription(cboPreset.Text)
        If i <> 0 Then WriteINI msg, "Settings", "Index", Trim(Str(i))
        ClearStrings
        LoadStrings True
        lstString.Clear
        txtString.Text = ""
        FillListboxWithTextDescriptions lstString
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cboPreset_Click()"
End Sub

Private Sub cboServerMinimum_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sServerMinimum = cboServerMinimum.Text
Dim i As Integer, c As Integer, F As Integer, msg As String, s As Integer
msg = cmbNetwork.Text
s = lvwServers.SelectedItem.Index
If Err.Number <> 0 Then Err.Clear
cmbNetwork.Clear
lvwServers.ListItems.Clear
For i = 0 To 1000
    If Len(lServers.sNetwork(i).nDescription) <> 0 Then
        If lSettings.sShowSmallNetworks = False Then
            For F = 0 To lServers.sServerCount
                If lServers.sServer(F).sNetwork = i Then
                    c = c + 1
                End If
            Next F
            If c > CInt(Trim(cboServerMinimum.Text)) - 1 Then
            'If c > lSettings.sServerMinimum Then
                cmbNetwork.AddItem lServers.sNetwork(i).nDescription
            End If
            c = 0
        Else
            cmbNetwork.AddItem lServers.sNetwork(i).nDescription
        End If
    End If
Next i
If cmbNetwork.ListCount <> 0 And Len(msg) <> 0 Then cmbNetwork.ListIndex = FindComboBoxIndex(cmbNetwork, msg)
lvwServers.ListItems(s).Selected = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cboServerMinimum_Click()"
End Sub

Private Sub chkInitialBass_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If chkInitialBass.Value = 1 Then
    sldInitialBass.Enabled = True
Else
    sldInitialBass.Enabled = False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkInitialBass_Click()"
End Sub

Private Sub chkInitialCDAudio_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If chkInitialCDAudio.Value = 1 Then
    sldInitialCDAudio.Enabled = True
Else
    sldInitialCDAudio.Enabled = False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkInitialCDAudio_Click()"
End Sub

Private Sub chkInitialLineIN_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If chkInitialLineIN.Value = 1 Then
    sldInitialLineIN.Enabled = True
Else
    sldInitialLineIN.Enabled = False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkInitialLineIN_Click()"
End Sub

Private Sub chkInitialMIC_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If chkInitialMIC.Value = 1 Then
    sldInitialMic.Enabled = True
Else
    sldInitialMic.Enabled = False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkInitialLineIN_Click()"
End Sub

Private Sub chkInitialTreble_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If chkInitialTreble.Value = 1 Then
    sldInitialTreble.Enabled = True
Else
    sldInitialTreble.Enabled = False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkInitialTreble_Click()"
End Sub

Private Sub chkInitialWave_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If chkInitialWave.Value = 1 Then
    sldInitialWave.Enabled = True
Else
    sldInitialWave.Enabled = False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub chkInitialWave_Click()"
End Sub

Private Sub chkPortScaner_Click()

End Sub

Private Sub cmbNetwork_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, j As Integer, mItem As ListItem, word() As String
lvwServers.ListItems.Clear
i = FindNetworkIndex(cmbNetwork.Text)
If i <> 0 Then
    For j = 1 To lServers.sServerCount
        If lServers.sServer(j).sNetwork = i Then
            Set mItem = lvwServers.ListItems.Add(, , lServers.sServer(j).sDescription)
            mItem.SubItems(1) = lServers.sServer(j).sServer
            mItem.SubItems(2) = lServers.sServer(j).sPortRange
        End If
    Next j
End If
If lSettings.sAutoPortScanner = True Then
    If cmdScan.Value = True Then
        ClearPortScan False
        tmrPortScan.Enabled = True
    Else
        cmdScan.Value = True
    End If
Else
    ClearPortScan False
    tmrPortScan.Enabled = False
    tmrPortScanTimeout.Enabled = False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmbNetwork_Click()"
End Sub

Private Sub cmdAdd_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(Trim(txtNickname.Text)) <> 0 Then
    i = CheckListbox(lstNotify, Trim(txtNotifyNickName.Text))
    If i = 0 Then
        lstNotify.AddItem Trim(txtNotifyNickName.Text)
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAdd_Click()"
End Sub

Private Sub cmdAddAlternate_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer
msg = InputBox("New Alternate Nickname:", "Choose Alternate Nickname")
If Len(msg) <> 0 Then
    cboAlternates.AddItem msg
    i = AddAlternate(msg)
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAddAlternate_Click()"
End Sub

Private Sub cmdAddAutoPerform_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = Trim(InputBox("Enter Command:", "Add Auto Perform"))
If Len(msg) <> 0 Then
    AddAutoPerform msg
    cboAutoPerform.AddItem msg
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAddAutoPerform_Click()"
End Sub

Private Sub cmdAddBlackList_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
AddToBlacklist InputBox("Enter Nickname:", "NexIRC"), InputBox("Enter Address:", "NexIRC")
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAddBlackList_Click()"
End Sub

Private Sub cmdAddBot_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer
msg = Trim(InputBox("Enter Bot's Nickname:", "Add Bot"))
If Len(msg) <> 0 Then
    AddBot msg, cboBotTypes.ListIndex: DoEvents
    i = FindBotIndex(msg)
    If i <> 0 Then cboBots.AddItem msg
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAddBot_Click()"
End Sub

Private Sub cmdAddBotCommand_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer
msg = Trim(InputBox("Enter Command:", "Add Command"))
If Len(msg) <> 0 Then
    i = AddBotCommand(msg, cboBotTypes.ListIndex)
    If i <> 0 Then
        cboCommands.AddItem msg
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAddBotCommand_Click()"
End Sub

Private Sub cmdAddtoIgnore_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = InputBox("Enter Nickname:", App.Title, "")
If Len(msg) <> 0 Then
    lstIgnore.AddItem msg
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAddtoIgnore_Click()"
End Sub

Private Sub cmdApply_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ActiveateSaveSettings
End Sub

Private Sub cmdApplyPreset_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdApplyPreset_Click()"
End Sub

Private Sub cmdCancel_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lDoneLoading = False
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdCancel_Click()"
End Sub

Private Sub cmdClear_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lstNotify.Clear
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdClearAlternates_Click()"
End Sub

Private Sub cmdClearAlternates_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ClearAlternates
cboAlternates.Clear
Kill GetINIFile(iAlternates)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdClearAlternates_Click()"
End Sub

Private Sub cmdClearBlacklist_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ClearBlacklist
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdClearBlacklist_Click()"
End Sub

Private Sub cmdClearIgnore_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
lstIgnore.Clear
ClearIgnore
chkEnableIgnore.Value = 0
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdClearIgnore_Click()"
End Sub

Private Sub cmdClearServers_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lvwServers.ListItems.Clear
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdClearServers_Click()"
End Sub

Private Sub cmdConnect_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim xServer As String, xPort As String, i As Integer, PortFound As Boolean
If Len(txtNickname.Text) = 0 Then
    Beep
    ResetCheck 1
    CreateBalloon "Nickname", "Your nickname is the name by which people will know you on IRC. " & vbCrLf & "Remember that there are many hundreds of thousands of people " & vbCrLf & " on IRC, so it's possible that someone might already be using the " & vbCrLf & " nickname you've chosen. If that's the case, you should try to pick " & vbCrLf & " a different, more unique, nickname. You can enter an alternative " & vbCrLf & " nickname as well in case someone is using your first nickname.", txtNickname
    txtNickname.SetFocus
    Exit Sub
End If
If Len(txtEmail.Text) = 0 Then
    Beep
    ResetCheck 1
    CreateBalloon "E-Mail", "You must enter a full email address like user@host.com.", txtEmail
    txtEmail.SetFocus
    Exit Sub
End If
If Len(txtRealName.Text) = 0 Then
    Beep
    ResetCheck 1
    CreateBalloon "Real Name", "You can enter your real name here, however note that whatever " & vbCrLf & " you enter can be seen by other people on IRC. Most people usually " & vbCrLf & " enter a witty one-liner or comment.", txtRealName
    txtRealName.SetFocus
    Exit Sub
End If
tmrPortScan.Enabled = False
tmrPortScanTimeout.Enabled = False
If chkNewServerWindow.Value = 1 Then
    NewStatusWindow txtServer.Text, txtPort.Text, True
End If
lModes.mI = GetCheckboxValue(chkInvisible)
lModes.mS = GetCheckboxValue(chkServerMSG)
lModes.mW = GetCheckboxValue(chkWallops)
xServer = Trim(txtServer)
xPort = Val(Trim(txtPort))
If xServer = "" Then xServer = "127.0.0.1"
If xPort = "" Then xPort = "6667"
lSettings.sEMail = txtEmail.Text
lSettings.sRealName = txtRealName.Text
lSettings.sNickname = txtNickname.Text
With lSettings.sActiveServerForm.tcp
    If .State <> 0 Then
        .Close
    End If
    If Left(xServer, 1) = ":" Then
        xServer = Right(xServer, Len(xServer) - 1)
        lSettings.sServer = xServer
    End If
    ConnectToIRC xServer, xPort, lSettings.sActiveServerForm
End With
ActiveateSaveSettings
SaveSettings
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdConnect_Click()"
End Sub

Private Sub cmdDefaultIdent_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtIdentPort = "113"
txtIdentSystem.Text = "UNIX"
txtIdentUserID.Text = "nexirc"
txtNickname.Text = "nexirc"
txtEmail.Text = "nexirc@tnexgen.com"
txtRealName.Text = "John Doe"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdDefaultIdent_Click()"
End Sub

Private Sub cmdDelete_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To lstNotify.ListCount - 1
    If lstNotify.List(lstNotify.ListIndex) = lstNotify.List(i) Then
        lstNotify.RemoveItem i
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdDelete_Click()"
End Sub

Private Sub cmdDeleteAutoPerform_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DeleteAutoPerform FindAutoPerformIndex(cboAutoPerform.Text)
cboAutoPerform.RemoveItem cboAutoPerform.ListIndex
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdDeleteAutoPerform_Click()"
End Sub

Private Sub cmdDeleteBot_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = cboBots.Text
If Len(msg) <> 0 Then
    RemoveBot cboBots.Text
    cboBots.RemoveItem cboBots.ListIndex
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdDeleteBot_Click()"
End Sub

Private Sub cmdDeleteCommands_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
RemoveBotCommand cboCommands.Text
cboCommands.RemoveItem cboCommands.ListIndex
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdDeleteCommands_Click()"
End Sub

Private Sub cmdEdit_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer
If Len(cboAutoPerform.Text) <> 0 Then
    msg = InputBox("Enter new text", "Edit", cboAutoPerform.Text)
    If Len(msg) <> 0 Then
        i = cboAutoPerform.ListIndex
        cboAutoPerform.RemoveItem cboAutoPerform.ListIndex
        cboAutoPerform.AddItem msg, i
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdEdit_Click()"
End Sub

Private Sub cmdNetworkAdd_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmAddNetwork.Show , frmCustomize
frmCustomize.SetFocus
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdNetworkAdd_Click()"
End Sub

Private Sub cmdNetworkDelete_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim delNet As String
Dim iNet As String
iNet = cmbNetwork.Text
If LCase(iNet) = "allgroups" Then Exit Sub
WriteINI GetINIFile(iServers), cmbNetwork.Text, vbNullString, vbNullString
Dim i As Integer
For i = 0 To cmbNetwork.ListCount - 1
    If LCase(cmbNetwork.List(i)) = LCase(cmbNetwork.Text) Then
        cmbNetwork.RemoveItem i
        Exit For
    End If
Next i
Dim varX As String
For i = 0 To 1000
    varX = i
    delNet = ReadINI(GetINIFile(iServers), "AllGroups", varX, "")
    If LCase(delNet) = LCase(iNet) Then
        WriteINI GetINIFile(iServers), "AllGroups", varX, vbNullString
        WriteINI GetINIFile(iServers), iNet, vbNullString, vbNullString
        Exit For
    End If
Next i
cmbNetwork.Text = "AllGroups"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdNetworkDelete_Click()"
End Sub

Private Sub cmdOK_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim X As Integer, m As Integer, i As String
Me.Visible = False
If chkNewServerWindow.Value = 1 Then
    NewStatusWindow txtServer.Text, txtPort.Text, True
End If
ActiveateSaveSettings
If chkThisSessionOnly.Value = 0 Then SaveSettings
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdOK_Click()"
End Sub

Private Sub cmdPreviewString_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As eStringTypes
i = ReturnStringTypeByDescription(lstString.Text)
ProcessReplaceString i, lSettings.sActiveServerForm.txtIncoming, "Param 1", "Param 2", "Param 3", "Param 4", "Param 5"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdPreviewString_Click()"
End Sub

Private Sub cmdRemoveAlternate_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mBool As Boolean
mBool = RemoveAlternate(cboAlternates.Text)
cboAlternates.RemoveItem cboAlternates.ListIndex
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdRemoveAlternate_Click()"
End Sub

Private Sub cmdRemoveBlacklist_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(lstBlacklist.Text) <> 0 Then RemoveFromBlacklist lstBlacklist.Text
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdRemoveBlacklist_Click()"
End Sub

Private Sub cmdRemoveIgnore_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lstIgnore.ListIndex <> -1 Then lstIgnore.RemoveItem lstIgnore.ListIndex
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdRemoveIgnore_Click()"
End Sub

Private Sub cmdRun_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = cboAutoPerform.Text
If Len(msg) <> 0 Then RunAutoPerform lSettings.sActiveServerForm
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdRun_Click()"
End Sub

Private Sub cmdScan_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If cmdScan.Value = True Then
    frmCustomize.tmrPortScan.Enabled = True
    lSettings.sTestConnectionsLoaded = True
Else
    ClearPortScan False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdScan_Click()"
End Sub

Private Sub cmdServerAdd_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmAddServer.Show , Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdServerAdd_Click()"
End Sub

Private Sub cmdServerDelete_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, d As Integer, m As Integer, mbox As VbMsgBoxResult
If lSettings.sGeneralPrompts = True Then
    mbox = MsgBox("Are you sure you wish to delete " & lvwServers.SelectedItem.SubItems(1) & "?", vbQuestion + vbYesNo)
    If mbox = vbNo Then Exit Sub
End If
d = FindServerIndex(lvwServers.SelectedItem.SubItems(1))
If d <> 0 Then
    For m = 0 To 150
        msg = ReadINI(GetINIFile(iServers), cmbNetwork.Text, Trim(Str(m)), "")
        If InStr(LCase(msg), LCase(lvwServers.SelectedItem.SubItems(1))) Then
            If InStr(LCase(msg), LCase(lvwServers.SelectedItem.Text)) Then
                WriteINI GetINIFile(iServers), cmbNetwork.Text, Trim(Str(m)), vbNullString
                Exit For
            End If
        End If
    Next m
    lServers.sServer(d).sDescription = ""
    lServers.sServer(d).sNetwork = 0
    lServers.sServer(d).sPassword = ""
    lServers.sServer(d).sPortRange = 0
    lServers.sServer(d).sServer = ""
    lvwServers.ListItems.Remove lvwServers.SelectedItem.Index
End If
If lvwServers.ListItems.Count = 0 Then Call cmdNetworkDelete_Click
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdServerDelete_Click()"
End Sub

Private Sub cmdServerEdit_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmEditServer.Show 1, Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdServerEdit_Click()"
End Sub

Private Sub cmdSmallNetworks_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, c As Integer, F As Integer, msg As String, s As Integer
msg = cmbNetwork.Text
s = lvwServers.SelectedItem.Index
lSettings.sServerMinimum = cboServerMinimum.Text
lSettings.sShowSmallNetworks = False
If cmdSmallNetworks.Value = True Then
    lSettings.sShowSmallNetworks = False
Else
    lSettings.sShowSmallNetworks = True
End If
cmbNetwork.Clear
lvwServers.ListItems.Clear
For i = 0 To 1000
    If Len(lServers.sNetwork(i).nDescription) <> 0 Then
        If lSettings.sShowSmallNetworks = False Then
            For F = 0 To lServers.sServerCount
                If lServers.sServer(F).sNetwork = i Then
                    c = c + 1
                End If
            Next F
            If c > lSettings.sServerMinimum - 3 Then
                cmbNetwork.AddItem lServers.sNetwork(i).nDescription
            End If
            c = 0
        Else
            cmbNetwork.AddItem lServers.sNetwork(i).nDescription
        End If
    End If
Next i
If cmbNetwork.ListCount <> 0 And Len(msg) <> 0 Then cmbNetwork.ListIndex = FindComboBoxIndex(cmbNetwork, msg)
lvwServers.ListItems(s).Selected = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdSmallNetworks_Click()"
End Sub

Private Sub cmdUpdateList_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Surf "http://www.team-nexgen.org/servers.shtml", Me.hWnd
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdUpdateList_Click()"
End Sub

Private Sub ctlXPButton1_Click()

End Sub

Private Sub SetProgress(lValue As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sShowExtraProgress = True Then
    XP_ProgressBar1.Value = lValue
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub SetProgress(lValue As Integer)"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim m As Integer, i As Integer, Color(0 To 15) As Long, getColors() As String, TagColor() As String, o As Integer, srColor As String, F As Integer, c As Integer, msg As String, msg2 As String, l As Integer
'If lSettings.sShowExtraProgress = True Then
'    Me.Visible = True
'    Me.SetFocus
'Else
'    Me.Visible = False
'End If
SetProgressBarColor XP_ProgressBar1
XP_ProgressBar1.Scrolling = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sProgressBarStyle
DoEvents
SetProgress 1
cboProgressBar.AddItem "Standard"
cboProgressBar.AddItem "Smooth"
cboProgressBar.AddItem "Office XP"
cboProgressBar.AddItem "Pastel"
cboProgressBar.AddItem "Java"
cboProgressBar.AddItem "Media Player"
cboProgressBar.AddItem "Custom"
cboProgressBar.AddItem "Picture"
cboProgressBar.AddItem "Metallic"
cboProgressBar.ListIndex = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sProgressBarStyle
Set lToolTip = New clsToolTip
'Dim i As Integer, c As Integer, msg As String, msg2 As String
msg = App.Path & "\data\config\fixed\text.ini"
SetProgress 2
c = Int(ReadINI(msg, "Settings", "Count", 0))
l = Int(ReadINI(msg, "Settings", "Index", 0))
If c <> 0 Then
    For i = 1 To c
        msg2 = ReadINI(msg, Trim(Str(i)), "Description", "")
        If Len(msg2) <> 0 Then
            cboPreset.AddItem msg2
        End If
    Next i
End If
If l <> 0 Then
    cboPreset.ListIndex = l - 1
End If
SetProgress 5

cboProgressBarColor.ComboItems.Add , , "White", 1, 1
cboProgressBarColor.ComboItems.Add , , "Black", 2, 2
cboProgressBarColor.ComboItems.Add , , "Dark Blue", 3, 3
cboProgressBarColor.ComboItems.Add , , "Dark Green", 4, 4
cboProgressBarColor.ComboItems.Add , , "Red", 5, 5
cboProgressBarColor.ComboItems.Add , , "Brown", 6, 6
cboProgressBarColor.ComboItems.Add , , "Purple", 7, 7
cboProgressBarColor.ComboItems.Add , , "Orange", 8, 8
cboProgressBarColor.ComboItems.Add , , "Yellow", 9, 9
cboProgressBarColor.ComboItems.Add , , "Light Green", 10, 10
cboProgressBarColor.ComboItems.Add , , "Dark Blue Green", 11, 11
cboProgressBarColor.ComboItems.Add , , "Light Blue Green", 12, 12
cboProgressBarColor.ComboItems.Add , , "Light Blue", 13, 13
cboProgressBarColor.ComboItems.Add , , "Magenta", 14, 14
cboProgressBarColor.ComboItems.Add , , "Gray", 15, 15
cboProgressBarColor.ComboItems.Add , , "Light Gray", 16, 16
cboProgressBarColor.ComboItems(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sProgressBarColor + 1).Selected = True
cboOpColor.ComboItems.Add , , "White", 1, 1
cboOpColor.ComboItems.Add , , "Black", 2, 2
cboOpColor.ComboItems.Add , , "Dark Blue", 3, 3
cboOpColor.ComboItems.Add , , "Dark Green", 4, 4
cboOpColor.ComboItems.Add , , "Red", 5, 5
cboOpColor.ComboItems.Add , , "Brown", 6, 6
cboOpColor.ComboItems.Add , , "Purple", 7, 7
cboOpColor.ComboItems.Add , , "Orange", 8, 8
cboOpColor.ComboItems.Add , , "Yellow", 9, 9
cboOpColor.ComboItems.Add , , "Light Green", 10, 10
cboOpColor.ComboItems.Add , , "Dark Blue Green", 11, 11
cboOpColor.ComboItems.Add , , "Light Blue Green", 12, 12
cboOpColor.ComboItems.Add , , "Light Blue", 13, 13
cboOpColor.ComboItems.Add , , "Magenta", 14, 14
cboOpColor.ComboItems.Add , , "Gray", 15, 15
cboOpColor.ComboItems.Add , , "Light Gray", 16, 16
SetProgress 10
cboVoiceColor.ComboItems.Add , , "White", 1, 1
cboVoiceColor.ComboItems.Add , , "Black", 2, 2
cboVoiceColor.ComboItems.Add , , "Dark Blue", 3, 3
cboVoiceColor.ComboItems.Add , , "Dark Green", 4, 4
cboVoiceColor.ComboItems.Add , , "Red", 5, 5
cboVoiceColor.ComboItems.Add , , "Brown", 6, 6
cboVoiceColor.ComboItems.Add , , "Purple", 7, 7
cboVoiceColor.ComboItems.Add , , "Orange", 8, 8
cboVoiceColor.ComboItems.Add , , "Yellow", 9, 9
cboVoiceColor.ComboItems.Add , , "Light Green", 10, 10
cboVoiceColor.ComboItems.Add , , "Dark Blue Green", 11, 11
cboVoiceColor.ComboItems.Add , , "Light Blue Green", 12, 12
cboVoiceColor.ComboItems.Add , , "Light Blue", 13, 13
cboVoiceColor.ComboItems.Add , , "Magenta", 14, 14
cboVoiceColor.ComboItems.Add , , "Gray", 15, 15
cboVoiceColor.ComboItems.Add , , "Light Gray", 16, 16
SetProgress 15
cboNormalColor.ComboItems.Add , , "White", 1, 1
cboNormalColor.ComboItems.Add , , "Black", 2, 2
cboNormalColor.ComboItems.Add , , "Dark Blue", 3, 3
cboNormalColor.ComboItems.Add , , "Dark Green", 4, 4
cboNormalColor.ComboItems.Add , , "Red", 5, 5
cboNormalColor.ComboItems.Add , , "Brown", 6, 6
cboNormalColor.ComboItems.Add , , "Purple", 7, 7
cboNormalColor.ComboItems.Add , , "Orange", 8, 8
cboNormalColor.ComboItems.Add , , "Yellow", 9, 9
cboNormalColor.ComboItems.Add , , "Light Green", 10, 10
cboNormalColor.ComboItems.Add , , "Dark Blue Green", 11, 11
cboNormalColor.ComboItems.Add , , "Light Blue Green", 12, 12
cboNormalColor.ComboItems.Add , , "Light Blue", 13, 13
cboNormalColor.ComboItems.Add , , "Magenta", 14, 14
cboNormalColor.ComboItems.Add , , "Gray", 15, 15
cboNormalColor.ComboItems.Add , , "Light Gray", 16, 16
SetProgress 20
If Len(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nOpColor) <> 0 Then
    i = Int(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nOpColor)
    cboOpColor.ComboItems(i).Selected = True
End If
If Len(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nVoiceColor) <> 0 Then
    i = Int(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nVoiceColor)
    cboVoiceColor.ComboItems(i).Selected = True
End If
If Len(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nNormalColor) <> 0 Then
    i = Int(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sNicklistOptions.nNormalColor)
    cboNormalColor.ComboItems(i).Selected = True
End If
SetProgress 25
LoadPortScanRange
SetProgress 30
FillComboWithStringGroups cboGroupBy
FillListBoxWithBlacklist lstBlacklist
SetProgress 35
CutRegion cmbNetwork.hWnd, cmbNetwork, True
CutRegion cboAlternates.hWnd, cboAlternates, True
CutRegion cboAutoPerform.hWnd, cboAutoPerform, True
CutRegion cboBots.hWnd, cboBots, True
CutRegion cboBotTypes.hWnd, cboBotTypes, True
CutRegion cboButtonType.hWnd, cboButtonType, True
CutRegion cboColorTheme.hWnd, cboColorTheme, True
CutRegion cboCommands.hWnd, cboCommands, True
CutRegion cboHelpTopic.hWnd, cboHelpTopic, True
SetProgress 40
cboBotTypes.AddItem "0 - Unknown/Custom Bot"
cboBotTypes.AddItem "1 - Eggdrop"
cboBotTypes.AddItem "2 - Undernet X"
cboBotTypes.AddItem "3 - ChanServ"
cboBotTypes.AddItem "4 - MemoServ"
cboBotTypes.ListIndex = 0
SetProgress 45
FillComboWithAutoPerform cboAutoPerform
If cboAutoPerform.ListCount <> 0 Then cboAutoPerform.ListIndex = 0
For i = 0 To 10
    SetButtonType optCheck(i)
Next i
SetProgress 50
SetButtonType cmdApply
SetButtonType cmdAddBlacklist
SetButtonType cmdRemoveBlacklist
SetButtonType cmdPreviewString
SetButtonType cmdClearBlackList
SetButtonType cmdScan
SetButtonType cmdClearServers
SetButtonType cmdClear
SetButtonType cmdSmallNetworks
SetButtonType cmdAddBot
SetButtonType cmdDeleteBot
SetButtonType cmdAddBotCommand
SetButtonType cmdDeleteCommands
SetButtonType cmdAddAutoPerform
SetButtonType cmdDeleteAutoPerform
SetButtonType cmdRun
SetButtonType cmdEdit
SetButtonType cmdConnect
SetButtonType cmdOK
SetButtonType cmdCancel
SetButtonType cmdNetworkAdd
SetButtonType cmdNetworkDelete
SetButtonType cmdServerAdd
SetButtonType cmdClear
SetButtonType cmdServerDelete
SetButtonType cmdServerEdit
SetButtonType cmdAdd
SetButtonType cmdDelete
SetButtonType cmdAddToIgnore
SetButtonType cmdRemoveIgnore
SetButtonType cmdClearIgnore
SetButtonType cmdAddAlternate
SetButtonType cmdRemoveAlternate
SetButtonType cmdClearAlternates
SetButtonType cmdDefaultIdent
SetProgress 60
lSettings.sCustomizeVisible = True
FillComboWithAlternates cboAlternates
If cboAlternates.ListCount <> 0 Then cboAlternates.ListIndex = 0
For i = 1 To lSpectrumThemes.sCount
    If Len(lSpectrumThemes.sSpectrumTheme(i).sName) <> 0 Then
        cboColorTheme.AddItem lSpectrumThemes.sSpectrumTheme(i).sName
    End If
Next i
SetProgress 65
cboColorTheme.ListIndex = Trim(FindComboBoxIndex(cboColorTheme, lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sName))
'If lRegInfo.rRegistered = False Then
'    chkShowSplashScreenOnStartup.Enabled = False
'End If
If lSettings.sShowSmallNetworks = True Then
    cmdSmallNetworks.Value = False
Else
    cmdSmallNetworks.Value = True
End If
For i = 0 To 1000
    If Len(lServers.sNetwork(i).nDescription) <> 0 Then
        If lSettings.sShowSmallNetworks = False Then
            For F = 0 To lServers.sServerCount
                If lServers.sServer(F).sNetwork = i Then
                    c = c + 1
                End If
            Next F
            If c > lSettings.sServerMinimum Then
                cmbNetwork.AddItem lServers.sNetwork(i).nDescription
            End If
            c = 0
        Else
            cmbNetwork.AddItem lServers.sNetwork(i).nDescription
        End If
    End If
Next i
SetProgress 70
lvwServers.ColumnHeaders.Add , , "Description", 2000
lvwServers.ColumnHeaders.Add , , "Server", 2500
lvwServers.ColumnHeaders.Add , , "Ports", 1000
If cmbNetwork.ListCount <> 0 And Len(lSettings.sNetwork) <> 0 Then cmbNetwork.ListIndex = FindComboBoxIndex(cmbNetwork, lSettings.sNetwork)
i = 0
For i = 1 To lvwServers.ListItems.Count
    If LCase(lvwServers.ListItems(i).ListSubItems(1).Text) = LCase(lSettings.sServer) Then
        lvwServers.ListItems(i).Selected = True
    End If
Next i
txtServer.Text = lSettings.sServer
txtPort.Text = lSettings.sPort
txtPassword.Text = lSettings.sPassword
txtNickname = lSettings.sNickname
txtEmail.Text = lSettings.sEMail
txtRealName.Text = lSettings.sRealName
txtIdentUserID.Text = lSettings.sIdent.iUserID
txtIdentSystem.Text = lSettings.sIdent.iSystem
txtIdentPort.Text = lSettings.sIdent.iPort
txtHomepage.Text = lSettings.sHomepage
SetCheckBoxValue chkIdent, lSettings.sIdent.iEnabled
SetCheckBoxValue chkIdentShow, lSettings.sIdent.iShow
FillListBoxWithNotify lstNotify
lblVersion.Caption = "Version: " & App.Major & "." & App.Minor
cboServerMinimum.Text = lSettings.sServerMinimum
SetProgress 76
cboHelpTopic.AddItem "How to use the 'Add Bot/Command' window"
cboHelpTopic.AddItem "How to use the 'Add Folder to Playlist' window"
cboHelpTopic.AddItem "How to use the 'Add Media to Playlist' window"
cboHelpTopic.AddItem "How to use the 'Add Network' window"
cboHelpTopic.AddItem "How to use the 'Add Server' window"
cboHelpTopic.AddItem "How to use the 'Alarm' window"
cboHelpTopic.AddItem "How to use the 'Auto Join' window"
cboHelpTopic.AddItem "How to use the 'Bots' window"
cboHelpTopic.AddItem "How to use the 'Channel' Window"
cboHelpTopic.AddItem "How to use the 'Channel Folder' window"
cboHelpTopic.AddItem "How to use the 'Channel Listing' window"
cboHelpTopic.AddItem "How to use the 'Channels' window"
cboHelpTopic.AddItem "How to use the 'Chat' window"
cboHelpTopic.AddItem "How to use the 'Color Editor' window"
cboHelpTopic.AddItem "How to use the 'Connection Manager' window"
cboHelpTopic.AddItem "How to use the 'DCC Accept' window"
cboHelpTopic.AddItem "How to use the 'DCC Chat' window"
cboHelpTopic.AddItem "How to use the 'DCC Get' window"
cboHelpTopic.AddItem "How to use the 'Download Manager' window"
cboHelpTopic.AddItem "How to use the 'Edit server' window"
cboHelpTopic.AddItem "How to use the 'IRC Server' window"
cboHelpTopic.AddItem "How to use the 'Join Channel' window"
cboHelpTopic.AddItem "How to use the 'Message Server' window"
cboHelpTopic.AddItem "How to use the 'MOTD' window"
cboHelpTopic.AddItem "How to use the 'Notify' window"
cboHelpTopic.AddItem "How to use the 'Options' window"
cboHelpTopic.AddItem "How to use the 'Playlist' window"
cboHelpTopic.AddItem "How to use the 'Query' window"
cboHelpTopic.AddItem "How to use the 'Quick Connect' window"
'cboHelpTopic.AddItem "How to use the 'Register' window"
cboHelpTopic.AddItem "How to use the 'Script Range' window"
cboHelpTopic.AddItem "How to use the 'Search within Playlist' window"
cboHelpTopic.AddItem "How to use the 'Send File' window"
cboHelpTopic.AddItem "How to use the 'Status' window"
cboHelpTopic.AddItem "How to use the 'Text Editor' window"
cboHelpTopic.AddItem "How to use the 'Auto Connect' window"
SetProgress 86
cboButtonType.AddItem "1 - Windows 16 Bit"
cboButtonType.AddItem "2 - Windows 32 Bit"
cboButtonType.AddItem "3 - Windows XP"
cboButtonType.AddItem "4 - Mac"
cboButtonType.AddItem "5 - Java"
cboButtonType.AddItem "6 - Netscape"
cboButtonType.AddItem "7 - Simple Flat"
cboButtonType.AddItem "8 - Flat Highlight"
cboButtonType.AddItem "9 - Office XP"
cboButtonType.AddItem "10 - Transparent"
cboButtonType.AddItem "11 - 3D Hover"
cboButtonType.AddItem "12 - Oval Flat"
cboButtonType.AddItem "13 - KDE/2"
cboButtonType.ListIndex = lSettings.sButtonType
SetProgress 91
sldInitialWave.Value = lInitialAudioValues.iWave
sldInitialBass.Value = lInitialAudioValues.iBass
sldInitialCDAudio.Value = lInitialAudioValues.iCDAudio
sldInitialLineIN.Value = lInitialAudioValues.iLineIN
sldInitialMic.Value = lInitialAudioValues.iMic
sldInitialTreble.Value = lInitialAudioValues.iTreble
If lInitialAudioValues.iInitialBassEnabled = True Then
    chkInitialBass.Value = 1
    sldInitialBass.Enabled = True
Else
    chkInitialBass.Value = 0
    sldInitialBass.Enabled = False
End If
If lInitialAudioValues.iInitialTrebleEnabled = True Then
    chkInitialTreble.Value = 1
    sldInitialTreble.Enabled = True
Else
    chkInitialTreble.Value = 0
    sldInitialTreble.Enabled = False
End If
If lInitialAudioValues.iInitialCDAudioEnabled = True Then
    chkInitialCDAudio.Value = 1
    sldInitialCDAudio.Enabled = True
Else
    chkInitialCDAudio.Value = 0
    sldInitialCDAudio.Enabled = False
End If
If lInitialAudioValues.iInitialLineInEnabled = True Then
    chkInitialLineIN.Value = 1
    sldInitialLineIN.Enabled = True
Else
    chkInitialLineIN.Value = 0
    sldInitialLineIN.Enabled = False
End If
If lInitialAudioValues.iInitialMicEnabled = True Then
    chkInitialMIC.Value = 1
    sldInitialMic.Enabled = True
Else
    chkInitialLineIN.Value = 0
    sldInitialMic.Enabled = False
End If
If lInitialAudioValues.iInitialWaveEnabled = True Then
    chkInitialWave.Value = 1
    sldInitialWave.Enabled = True
Else
    chkInitialWave.Value = 0
    sldInitialWave.Enabled = False
End If
SetProgress 94
txtQuitMessage.Text = ReturnStringDataByType(sQuitReason)
SetCheckBoxValue chkByPassStartupScreen, lSettings.sByPassStartupScreen
SetCheckBoxValue chkShowWhoisInChannel, lSettings.sShowWhoisInChannel
SetCheckBoxValue chkFileOfferInChannel, lSettings.sFileOfferInChannel
SetCheckBoxValue chkAudioServer, lSettings.sAudioServer
SetCheckBoxValue chkAutoPortScanner, lSettings.sAutoPortScanner
SetCheckBoxValue chkDownloadManager, lSettings.sDownloadManager
SetCheckBoxValue chkColoredNicklist, lSettings.sColoredNicklist
SetCheckBoxValue chkUpdateCheck, lSettings.sUpdateCheck
SetCheckBoxValue chkSecureQuery, lSettings.sSecureQuery
SetCheckBoxValue chkNickCompletor, lSettings.sUseNickCompletor
SetCheckBoxValue chkAutoSizeStatusbarItems, lSettings.sAutosizeStatusbarItems
SetCheckBoxValue chkBorderlessObjects, lSettings.sBorderlessObjects
SetCheckBoxValue chkShowTips, lSettings.sShowTips
SetCheckBoxValue chkPlaySounds, lSettings.sPlaySounds
SetCheckBoxValue chkReconnectOnDisconnect, lSettings.sReconnectOnDisconnect
SetCheckBoxValue chkTimeStamping, lSettings.sTimeStamping
SetCheckBoxValue chkOfferWhenPlayed, lSettings.sOfferWhenPlayed
SetCheckBoxValue chkEnableFind, lSettings.sEnableSearch
SetCheckBoxValue chkEnableListmedia, lSettings.sEnableList
SetCheckBoxValue chkSaveColorsToTheme, lSettings.sSaveIRCColorsToTheme
SetCheckBoxValue chkApplyThemeToIRCColors, lSettings.sApplyThemeToIRCColors
SetCheckBoxValue chkAutoSelectAlternateNickname, lSettings.sAutoSelectAlternateNickname
SetCheckBoxValue chkRefreshPictureColors, lSettings.sRefreshPictureColors
SetCheckBoxValue chkShowServerOnStartup, lSettings.sShowServerOnStartup
SetCheckBoxValue chkShowQuickNotify, lSettings.sShowQuickNotify
SetCheckBoxValue chkShowNotifyWindow, lSettings.sShowNotifyWindow
SetCheckBoxValue chkLogoTwitchOnPeaks, lSettings.sLogoTwitchOnPeaks
SetCheckBoxValue chkAlwaysShowAudioSettings, lSettings.sAlwaysShowAudioSettings
SetCheckBoxValue chkDCCEnabled, lSettings.sDCCEnabled
SetCheckBoxValue chkExclusiveToMp3InPlaylist, lSettings.sExlusiveToMp3InPlaylist
SetCheckBoxValue chkSearchForMedia, lSettings.sSearchForMedia
SetCheckBoxValue chkNavigateOnStartup, lSettings.sNavigateOnStartup
SetCheckBoxValue chkAutoJoinEnabled, lSettings.sAutoJoinEnabled
SetCheckBoxValue chkShowQuickMix, lSettings.sShowQuickmix
SetCheckBoxValue chkAddJoinedChannelsToChannelFolder, lSettings.sAddJoinedChannelsToChannelFolder
SetCheckBoxValue chkShuffle, lSettings.sShuffle
SetCheckBoxValue chkContinuousPlay, lSettings.sContinuousPlay
SetCheckBoxValue chkShowSplashScreenOnStartup, lSettings.sShowSplashOnStartup
SetCheckBoxValue chkConnectOnStartup, lSettings.sConnectOnStartup
SetProgress 96
If lPlayback.pCurrentEngine = pMediaPlayer Then
    optEngine(2).Value = True
ElseIf lPlayback.pCurrentEngine = pMp3 Then
    optEngine(1).Value = True
End If
SetCheckBoxValue chkWhois, lSettings.sOptions.oWhois
SetCheckBoxValue chkAutoJoin, lSettings.sAutoJoinOnInvite
SetCheckBoxValue chkInvisible, lSettings.sOptions.oInvisable
SetCheckBoxValue chkServerMSG, lSettings.sOptions.oServerMessages
SetCheckBoxValue chkWallops, lSettings.sOptions.oOpMessages
SetCheckBoxValue chkRejoin, lSettings.sOptions.oReJoin
SetCheckBoxValue chkWhois, lSettings.sOptions.oWhois
SetCheckBoxValue chkSkipMOTD, lSettings.sOptions.oSkipMOTD
SetCheckBoxValue chkShowMOTD, lSettings.sOptions.oShowMOTD
SetCheckBoxValue chkShowAddress, lSettings.sOptions.oShowAddress
SetCheckBoxValue chkEnable, ReturnNotifyEnabled
SetCheckBoxValue chkWhoisNotify, lSettings.sOptions.oWhoisNotify
SetCheckBoxValue chkNotifyOnActive, lSettings.sOptions.oShowNotifyInActiveWindow
SetCheckBoxValue chkShowQuits, lSettings.sOptions.oShowQuit
SetCheckBoxValue chkShowJoinPart, lSettings.sOptions.oShowJoinPart
SetCheckBoxValue chkShowModes, lSettings.sOptions.oShowModes
SetCheckBoxValue chkShowTopics, lSettings.sOptions.oShowTopics
SetCheckBoxValue chkShowKicks, lSettings.sOptions.oShowKicks
SetCheckBoxValue chkEnableIgnore, ReturnIgnoreEnabled
SetCheckBoxValue chkShowMe, lSettings.sShowOptionsOnStartup
SetCheckBoxValue chkBackgroundWebpage, lSettings.sBackgroundWebpage
SetCheckBoxValue chkGeneralPrompts, lSettings.sGeneralPrompts
SetCheckBoxValue chkDCCPrompts, lSettings.sDCCPrompts
FillListboxWithTextDescriptions lstString
SetProgress 98
For m = 0 To 150
    If Len(ReturnIgnoreNickname(m)) <> 0 Then
        If FindListBoxIndex(ReturnIgnoreNickname(m), lstIgnore) = 0 Then
            lstIgnore.AddItem ReturnIgnoreNickname(m)
        End If
    End If
Next m
Color(0) = vbWhite
Color(1) = vbBlack
Color(2) = RGB(42, 42, 87)
Color(3) = RGB(33, 112, 33)
Color(4) = vbRed
Color(5) = RGB(109, 50, 50)
Color(6) = RGB(119, 33, 119)
Color(7) = RGB(252, 127, 0)
Color(8) = RGB(195, 195, 56)
Color(9) = RGB(0, 252, 0)
Color(10) = RGB(89, 167, 179)
Color(11) = RGB(0, 255, 255)
Color(12) = vbBlue
Color(13) = RGB(255, 0, 255)
Color(14) = RGB(127, 127, 127)
Color(15) = RGB(210, 210, 210)
SetProgress 99
For i = o To 15
    picColor(i).BackColor = Color(i)
Next i
srColor = Trim(lSettings.sColors)
getColors = Split(srColor, " ")
For i = 0 To UBound(getColors)
    TagColor = Split(getColors(i), ":")
'    lblColor(TagColor(0)).Tag = TagColor(1)
'    lblColor(TagColor(0)).ForeColor = color(TagColor(1))
    If Len(lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor) <> 0 Then picBGColor.BackColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor
    If i = 1 Then
        If Len(TagColor(1)) <> 0 Then picBGColor.BackColor = Color(TagColor(1))
    End If
Next i
lblcolor(0).ForeColor = lblcolor(1).ForeColor
optCheck_Click (0)
SetProgress 100
XP_ProgressBar1.Visible = False
'If lSettings.sShowExtraProgress = True Then Me.Visible = False
lDoneLoading = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lDoneLoading = False
Set lToolTip = Nothing
lSettings.sCustomizeVisible = False
If lSettings.sMainVisisble = True Then
    mdiNexIRC.SetFocus
End If
If lSettings.sTestConnectionsLoaded = True Then ClearPortScan True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub Label12_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtServer.SetFocus
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Label12_Click()"
End Sub

Private Sub Label13_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtPort.SetFocus
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Label13_Click()"
End Sub

Private Sub Label14_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtPassword.SetFocus
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Label14_Click()"
End Sub

Private Sub Image2_Click()
Surf "http://www.team-nexgen.org/scripts.shtml", Me.hWnd
'fhioewfhiop
'<input type="image" src="https://www.paypal.com/en_US/i/btn/x-click-but04.gif" border="0" name="submit" alt="Make payments with PayPal - it's fast, free and secure!">
End Sub

Private Sub Label7_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Surf "http://www.team-nexgen.org", Me.hWnd
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Label7_Click()"
End Sub

Private Sub lblcolor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblcolor(Index).Font.Bold = True
End Sub

Private Sub lblcolor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To 15
    lblcolor(i).Font.Underline = False
Next i
lblcolor(Index).Font.Underline = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lblEMail_Click()"
End Sub

Private Sub lblcolor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblcolor(Index).Font.Bold = False
End Sub

Private Sub lblEMail_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Surf "mailto:guide_X@live.com", Me.hWnd
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lblEMail_Click()"
End Sub

Private Sub lstString_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next

txtString.Text = ReturnStringDataByType(ReturnStringTypeByDescription(lstString.Text))
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lstString_Click()"
End Sub

Private Sub lvwServers_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.txtServer = lvwServers.SelectedItem.SubItems(1)
Me.txtPort = lvwServers.SelectedItem.SubItems(2)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lvwServers_Click()"
End Sub

Private Sub lvwServers_DblClick()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdOK_Click
ConnectToIRC txtServer.Text, txtPort.Text, lSettings.sActiveServerForm
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lvwServers_DblClick()"
End Sub

Private Sub ResetCheck(lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To optCheck.Count - 1
    If i <> lIndex Then
        optCheck(i).Value = False
    End If
    fraSettings(i).Visible = False
Next i
DoEvents
fraSettings(lIndex).Visible = True
Select Case lIndex
Case 0
    Me.Caption = "NexIRC - Network/Server"
    
Case 1
    Me.Caption = "NexIRC - User/Identd"
    txtNickname.SetFocus
Case 2
    Me.Caption = "NexIRC - Options"
    chkUpdateCheck.SetFocus
Case 3
    Me.Caption = "NexIRC - Notify"
    lstNotify.SetFocus
Case 4
    Me.Caption = "NexIRC - Themes/Colors"
    cboColorTheme.SetFocus
Case 5
    Me.Caption = "NexIRC - Text"
    lstString.SetFocus
Case 6
    Me.Caption = "NexIRC - About"
    
Case 7
    Me.Caption = "NexIRC - Ignore"
    lstIgnore.SetFocus
Case 8
    Me.Caption = "NexIRC - Help"
    DisplayHelpInformation 25
    cboHelpTopic.SetFocus
Case 9
    Me.Caption = "NexIRC - Audio"
    optEngine(1).SetFocus
Case 10
    Me.Caption = "NexIRC - Bots"
    cboBotTypes.SetFocus
End Select
'Me.SetFocus
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub ResetCheck(lIndex As Integer)"
End Sub

Private Sub optCheck_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
chkShowMe.Visible = True
cmdConnect.Visible = True
ResetCheck Index
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub optCheck_Click(Index As Integer)"
End Sub

Private Sub picBGColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To 15
    lblcolor(i).Font.Underline = False
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picBGColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub picColor_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim Color(0 To 15) As Long, i As Integer
Color(0) = vbWhite
Color(1) = vbBlack
Color(2) = RGB(42, 42, 87)
Color(3) = RGB(33, 112, 33)
Color(4) = vbRed
Color(5) = RGB(109, 50, 50)
Color(6) = RGB(119, 33, 119)
Color(7) = RGB(252, 127, 0)
Color(8) = RGB(195, 195, 56)
Color(9) = RGB(0, 252, 0)
Color(10) = RGB(89, 167, 179)
Color(11) = RGB(0, 255, 255)
Color(12) = vbBlue
Color(13) = RGB(255, 0, 255)
Color(14) = RGB(127, 127, 127)
Color(15) = RGB(210, 210, 210)
For i = 0 To 15
    If LCase(lblExample.Caption) = LCase(lblcolor(i).Caption) Then
        lblcolor(i).ForeColor = Color(Index)
        lblExample.ForeColor = Color(Index)
        If LCase(lblExample.Caption) = "background color" Then
            lblExample.BackColor = Color(Index)
            lblExample.ForeColor = lblcolor(1).ForeColor
            picBGColor.BackColor = Color(Index)
        End If
        lblcolor(i).Tag = Index
        If i = 0 Then
            picBGColor.BackColor = lblcolor(0).ForeColor
        End If
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picColor_Click(Index As Integer)"
End Sub

Private Sub lblColor_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next

lblExample.Caption = lblcolor(Index).Caption
lblExample.ForeColor = lblcolor(Index).ForeColor
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lblColor_Click(Index As Integer)"
End Sub

Private Sub picBGColor_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lblExample.Caption = "Background Color"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub picBGColor_Click()"
End Sub

Private Sub tmrPortScan_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If CheckPortScanTimerProc(lvwServers) = False Then
    tmrPortScan.Enabled = False
    tmrPortScanTimeout.Enabled = False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrPortScan_Timer()"
End Sub

Private Sub tmrPortScanTimeout_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
SetPortScanInProgress False
tmrPortScanTimeout.Enabled = False
End Sub

Private Sub txtEMail_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CreateBalloon "E-Mail", "You must enter a full email address like user@host.com.", txtEmail
txtEmail.SelStart = 0
txtEmail.SelLength = Len(txtEmail.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtEMail_GotFocus()"
End Sub

Private Sub txtEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'CreateBalloon "E-Mail", "Input your real e-mail address in this feild, you can not connect without it", txtEMail
End Sub

Private Sub txtHomepage_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtHomepage.SelStart = 0
txtHomepage.SelLength = Len(txtHomepage.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtHomepage_GotFocus()"
End Sub

Private Sub txtIdentPort_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtIdentPort.SelStart = 0
txtIdentPort.SelLength = Len(txtIdentPort.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtIdentPort_GotFocus()"
End Sub

Private Sub txtIdentSystem_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtIdentSystem.SelStart = 0
txtIdentSystem.SelLength = Len(txtIdentSystem.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtIdentSystem_GotFocus()"
End Sub

Private Sub txtIdentUserID_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtIdentUserID.SelStart = 0
txtIdentUserID.SelLength = Len(txtIdentUserID.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtIdentUserID_GotFocus()"
End Sub

Private Sub txtNickname_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CreateBalloon "Nickname", "Your nickname is the name by which people will know you on IRC. " & vbCrLf & "Remember that there are many hundreds of thousands of people " & vbCrLf & " on IRC, so it's possible that someone might already be using the " & vbCrLf & " nickname you've chosen. If that's the case, you should try to pick " & vbCrLf & " a different, more unique, nickname. You can enter an alternative " & vbCrLf & " nickname as well in case someone is using your first nickname.", txtNickname
txtNickname.SelLength = Len(txtNickname.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtNickname_GotFocus()"
End Sub

Private Sub txtNickname_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtNickname_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub txtNotifyNickName_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtNotifyNickName.SelStart = 0
txtNotifyNickName.SelLength = Len(txtNotifyNickName.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtNotifyNickName_GotFocus()"
End Sub

Private Sub txtPassword_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPort.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtPassword_GotFocus()"
End Sub

Private Sub txtQuitMessage_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtQuitMessage.SelStart = 0
txtQuitMessage.SelLength = Len(txtQuitMessage.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtQuitMessage_GotFocus()"
End Sub

Private Sub txtRealname_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CreateBalloon "Real Name", "You can enter your real name here, however note that whatever " & vbCrLf & " you enter can be seen by other people on IRC. Most people usually " & vbCrLf & " enter a witty one-liner or comment.", txtRealName
txtRealName.SelStart = 0
txtRealName.SelLength = Len(txtRealName.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtRealname_GotFocus()"
End Sub

Private Sub txtServer_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtServer.SelStart = 0
txtServer.SelLength = Len(txtServer.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtServer_GotFocus()"
End Sub

Private Sub txtport_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtPort.SelStart = 0
txtPort.SelLength = Len(txtPort.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtport_GotFocus()"
End Sub

Private Sub txtString_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If chkSaveChanges.Value = 1 Then
    SetStringData ReturnStringTypeByDescription(lstString.Text), txtString.Text
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtString_Change()"
End Sub

Private Sub txtString_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtString.SelStart = 0
txtString.SelLength = Len(txtString.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtString_GotFocus()"
End Sub

Private Sub wskTestConnection_Connect(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
RecievePortScanResults wskTestConnection(Index).RemoteHost, wskTestConnection(Index).RemotePort
wskTestConnection(Index).Close
SetPortScanInProgress False
End Sub

Private Sub wskTestConnection_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
wskTestConnection(Index).Close
End Sub
