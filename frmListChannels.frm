VERSION 5.00
Begin VB.Form frmListChannels 
   Caption         =   "NexIRC - Channel List"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListChannels.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3375
   ScaleWidth      =   6810
   Begin VB.ListBox lstChannels 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3150
      IntegralHeight  =   0   'False
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frmListChannels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const LB_SETTABSTOPS = &H192

Public Sub SetListTabStops(ListHandle As Long, ParamArray ParmList() As Variant)
On Local Error Resume Next
Dim i As Long
Dim ListTabs() As Long
Dim NumColumns As Long
ReDim ListTabs(UBound(ParmList))
For i = 0 To UBound(ParmList)
    ListTabs(i) = ParmList(i)
Next i
NumColumns = UBound(ParmList) + 1
Call SendMessage(ListHandle, LB_SETTABSTOPS, NumColumns, ListTabs(0))
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Call SetListTabStops(lstChannels.hwnd, 0, 74, 100)
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
lstChannels.Move lstChannels.left, lstChannels.top, Me.Width - 100, Me.Height - 400
End Sub

Private Sub lstChannels_DblClick()
On Local Error Resume Next
Dim CCC() As String
CCC = Split(lstChannels.List(lstChannels.ListIndex), vbTab)
mdiMain.tcp.SendData "JOIN " & CCC(0) & vbCrLf
End Sub
