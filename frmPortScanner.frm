VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPortScanner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Scan"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPortScanner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   2295
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   2640
   End
   Begin VB.TextBox txtBeginPort 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Text            =   "6660"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtEndPort 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Text            =   "7000"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtIP 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblCurrent 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Remote Host:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Current Port:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Start:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   150
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "End:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "frmPortScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SOCKET_ERROR = 134
Dim TotalPorts As Long
Dim PortDone As Integer
Dim OnPort As Long
Dim LocalHost As Integer
Dim PortOpen As Long
Dim Host As String
Dim IP As String
Dim iReturn As Long, sLowByte As String, sHighByte As String
Dim sMsg As String, HostLen As Long
Dim HOSTENT As HOSTENT
Dim PointerToPointer As Long, ListAddress As Long
Dim WSAdata As WSAdata, DotA As Long, DotAddr As String, ListAddr As Long
Dim MaxUDP As Long, MaxSockets As Long, i As Integer
Dim Description As String, Status As String
' Ping Variables
Dim bReturn As Boolean, hIP As Long
Dim szBuffer As String
Dim addr As Long
Dim RCode As String
Dim RespondingHost As String
' TRACERT Variables
Dim TraceRT As Boolean
Dim Ttl As Integer
' WSock32 Constants
Const WS_VERSION_MAJOR = &H101 \ &H100 And &HFF&
Const WS_VERSION_MINOR = &H101 And &HFF&
Const MIN_SOCKETS_REQD = 0
Private Type HOSTENT
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Private Type WSAdata
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
Private Type Inet_address
    Byte4 As String * 1
    Byte3 As String * 1
    Byte2 As String * 1
    Byte1 As String * 1
End Type
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, HostLen&) As Long
Private Declare Function gethostbyname& Lib "WSOCK32.DLL" (ByVal hostname$)
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private IPLong5 As Inet_address

'Public Sub vbGetHostName()
'    Host = String(64, &H0)
'    If gethostname(Host, HostLen) = SOCKET_ERROR Then
'        sMsg = "WSock32 Error" & str$(WSAGetLastError())
'        'MsgBox sMsg, 0, ""
'    Else
'        Host = Left$(Trim$(Host), Len(Trim$(Host)) - 1)
'        Host.Text = Host
'    End If
'End Sub

Public Sub AccessScanner()
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
lPortScans.pInProgress = True
cmdStart_Click
End Sub

Function ScanPort(thePort As Long, ws1 As Winsock) As Boolean
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
ScanPort = False
On Error GoTo gotport
ws1.Close
ws1.LocalPort = thePort
ws1.Listen
Pause 0.1
ws1.Close
Exit Function
gotport:
If Err.Number = 10048 Then
    ScanPort = True
End If
End Function

Private Sub cmdClose_Click()
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
End Sub

Private Sub cmdStart_Click()
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
LocalHost = 2
'txtBeginPort.Enabled = False
'txtEndPort.Enabled = False
'cmdStart.Enabled = False
'cmdStop.Enabled = True
OnPort = CLng(txtBeginPort.Text)
PortDone = 0
Call Scanner(CLng(txtBeginPort.Text), CLng(txtEndPort.Text))
End Sub

Public Sub Scanner(Begin As Long, ending As Long)
On Error GoTo errd
TotalPorts = 0
PortOpen = 0
Do Until OnPort = CLng(txtEndPort.Text)
Pause 0.05
If PortDone = 1 Then lblCurrent.Caption = lblCurrent.Caption - 1: Exit Sub
DoEvents
lblCurrent.Caption = str(OnPort)
If LocalHost = 1 Then
    If ScanPort(OnPort, Winsock1) = True Then
        TotalPorts = TotalPorts + 1
        PortOpen = PortOpen + 1
        If txtStatus = "" Then txtStatus = "Port " & OnPort & " is currently open.": GoTo thisPart
        RecievePortScanResults txtIP.Text, Trim(str(OnPort))
        txtStatus = txtStatus & vbCrLf & "Port " & OnPort & " is currently open."
        txtStatus.SelStart = Len(txtStatus)
    End If
ElseIf Len(txtIP.Text) > 1 Then
    Host = txtIP.Text
    vbGetHostByName
    Winsock1.Connect IP, OnPort
    Pause 0.2
    Winsock1.Close
End If
thisPart:
OnPort = OnPort + 1
Loop
'lblCurrent = "Done"
'txtStatus = txtStatus & vbCrLf & OnPort - 1 & " port(s) sucessfulley scanned." & vbCrLf & PortOpen & " Port(s) Open."
lPortScans.pInProgress = False
Unload Me
Exit Sub
errd:
'Stop
'MsgBox Err.Description
End Sub

Private Sub cmdStop_Click()
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
cmdStop.Enabled = False
txtBeginPort.Enabled = True
txtEndPort.Enabled = True
cmdStart.Enabled = True
PortDone = 1
txtStatus = txtStatus & vbCrLf & OnPort - 1 & " port(s) sucessfulley scanned." & vbCrLf & PortOpen & " Port(s) Open."
txtStatus.SelStart = Len(txtStatus)
cmdStart.SetFocus
Timer1.Enabled = False
End Sub

Private Sub Form_Load()
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
'optRemote = True
lblCurrent.Caption = "0"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
Call Clean_Up
Winsock1.Close
End Sub

Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
End Sub

Private Sub lblCurrent_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
End Sub

Private Sub optLocal_Click()
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtIP.Enabled = False
LocalHost = 1
End Sub

Private Sub optRemote_Click()
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtIP.Enabled = True
LocalHost = 2
End Sub

Private Sub Winsock1_Connect()
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
    txtStatus.Text = txtStatus.Text & vbCrLf & "Port " & OnPort & " is currently open."
    RecievePortScanResults txtIP.Text, Trim(str(OnPort))
    txtStatus.SelStart = Len(txtStatus)
    OnPort = OnPort + 1
    PortOpen = PortOpen + 1
End Sub

Public Sub vbGetHostByName()
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
    Dim szString As String
    Host = Trim$(Host)
    szString = String(64, &H0)
    Host = Host + Right$(szString, 64 - Len(Host))

    If gethostbyname(Host) = SOCKET_ERROR Then
        sMsg = "Winsock Error" & str$(WSAGetLastError())
'        MsgBox sMsg, 0, ""
    Else
        PointerToPointer = gethostbyname(Host) ' Get the pointer to the address of the winsock hostent structure
        CopyMemory HOSTENT.h_name, ByVal _
        PointerToPointer, Len(HOSTENT) ' Copy Winsock structure to the VisualBasic structure
        ListAddress = HOSTENT.h_addr_list ' Get the ListAddress of the Address List
        CopyMemory ListAddr, ByVal ListAddress, 4 ' Copy Winsock structure To the VisualBasic structure
        CopyMemory IPLong5, ByVal ListAddr, 4 ' Get the first list entry from the Address List
        CopyMemory addr, ByVal ListAddr, 4
        IP = Trim$(CStr(Asc(IPLong5.Byte4)) + "." + CStr(Asc(IPLong5.Byte3)) _
        + "." + CStr(Asc(IPLong5.Byte2)) + "." + CStr(Asc(IPLong5.Byte1)))
    End If
End Sub

Sub Clean_Up()
'If lSettings.sHandleErrors = True Then On Local Error Resume Next
lblCurrent = 1
PortDone = 1
txtStatus = ""
cmdStop.Enabled = False
txtBeginPort.Enabled = True
txtEndPort.Enabled = True
cmdStart.Enabled = True
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
txtStatus.Text = txtStatus.Text & vbCrLf & "Error: " & Description
End Sub
