VERSION 5.00
Begin VB.Form frmChannels 
   AutoRedraw      =   -1  'True
   Caption         =   "NexIRC - Channel List"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChannels.frx":0000
   LinkTopic       =   "frmChannels"
   MDIChild        =   -1  'True
   ScaleHeight     =   2565
   ScaleWidth      =   4980
   Begin nexIRC.ctlListView lvwChannels 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3413
   End
End
Attribute VB_Name = "frmChannels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lListViewIndex As Integer

Public Sub ActivateResize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Form_Resize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub ActivateResize()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = mdiNexIRC.Icon
lSettings.sChannelListVisible = True
lvwChannels.Initialize
lvwChannels.BorderStyle = bsThin
lvwChannels.ViewMode = vmDetails
lvwChannels.ColumnAdd 0, "Items", 447, [caLeft]
lvwChannels.Font.Name = "Tahoma"
lvwChannels.FullRowSelect = True
lvwChannels.HeaderFlat = False
lvwChannels.HeaderHide = False
lvwChannels.Sort 0, soAscending, stString = True
'lvwChannels.ColumnHeaders.Add , , "Channel", 3000
'lvwChannels.ColumnHeaders.Add , , "Users"
'lvwChannels.ColumnHeaders.Add , , "Topic", 3500
'lvwChannels.BackColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor
'lvwChannels.ForeColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sTextColor
Call AddTaskPanel("List", 1)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub Form_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lvwChannels.Width = Me.ScaleWidth
lvwChannels.Height = Me.ScaleHeight
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Resize()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sChannelListVisible = False
Call RemoveTaskbar("List")
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub lvwChannels_ColumnClick(Column As Integer)
lvwChannels.Sort Column, soAscending, stString
End Sub

'Private Sub lvwChannels_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'lvwChannels.AllowColumnReorder = True
'If ColumnHeader = "Channel" Then
    'lvwChannels.SortKey = 0
'End If
'If ColumnHeader = "Users" Then
'    lvwChannels.SortKey = 1
'End If
'If ColumnHeader = "Topic" Then
'    lvwChannels.SortKey = 2
'End If
'If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lvwChannels_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)"
'End Sub

Private Sub lvwChannels_DblClick()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.SendData "JOIN " & lvwChannels.ItemText(lListViewIndex) & vbCrLf
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lvwChannels_DblClick()"
End Sub

Private Sub lvwChannels_ItemClick(Item As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lListViewIndex = Item
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lvwChannels_ItemClick(Item As Integer)"
End Sub

Private Sub lvwChannels_KeyPress(KeyAscii As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If KeyAscii = 27 Then
    Me.WindowState = vbMinimized
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lvwChannels_KeyPress(KeyAscii As Integer)"
End Sub

'Private Sub lvwChannels_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Set lvwChannels.SelectedItem = lvwChannels.HitTest(X, Y)
'If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lvwChannels_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
'End Sub

Private Sub lvwChannels_ItemDblClick(Item As ListItem)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sActiveServerForm.tcp.SendData "JOIN " & Item.Text & vbCrLf
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lvwChannels_ItemDblClick(Item As ListItem)"
End Sub
