VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmWeb 
   BorderStyle     =   0  'None
   Caption         =   "Webpage"
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWeb.frx":0000
   LinkTopic       =   "frmWeb"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   2805
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      ExtentX         =   3413
      ExtentY         =   2355
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmWeb"
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

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sWebVisible = True
web.Silent = True
If Len(lSettings.sHomepage) <> 0 Then web.Navigate lSettings.sHomepage
DoEvents
ActivateResize
mdiNexIRC.ActivateResize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If web.Left <> 0 Then web.Left = 0
web.Width = Me.ScaleWidth
web.Height = Me.ScaleHeight
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Resize()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sWebVisible = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub web_NavigateError(ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
web.Navigate App.Path & "\data\help\404.html"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub web_NavigateError(ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)"
End Sub
