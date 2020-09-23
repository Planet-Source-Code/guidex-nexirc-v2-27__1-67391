VERSION 5.00
Begin VB.Form frmMenuEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Menu Editor"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenuEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   5910
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstSubMenus 
      Height          =   1380
      IntegralHeight  =   0   'False
      Left            =   3120
      TabIndex        =   11
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtImageIndex 
      Height          =   285
      Left            =   960
      MousePointer    =   1  'Arrow
      TabIndex        =   7
      Top             =   1200
      Width           =   615
   End
   Begin VB.ComboBox cboMenu 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.ListBox lstMenu 
      Height          =   1380
      IntegralHeight  =   0   'False
      Left            =   960
      TabIndex        =   9
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtCommand 
      Height          =   285
      Left            =   960
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox txtCaption 
      Height          =   285
      Left            =   960
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin nexIRC.ctlXPButton cmdOK 
      Height          =   375
      Left            =   4680
      TabIndex        =   12
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
      MICON           =   "frmMenuEditor.frx":000C
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
      Height          =   375
      Left            =   4680
      TabIndex        =   13
      Top             =   600
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
      MICON           =   "frmMenuEditor.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdAddMenu 
      Height          =   375
      Left            =   960
      TabIndex        =   14
      ToolTipText     =   "Add Bot to List"
      Top             =   1560
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
      MICON           =   "frmMenuEditor.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdDeleteMenu 
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   1560
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
      MICON           =   "frmMenuEditor.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdAddSubMenu 
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   16
      ToolTipText     =   "Add Bot to List"
      Top             =   1560
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
      MICON           =   "frmMenuEditor.frx":007C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdDeleteSubMenu 
      Height          =   375
      Left            =   4680
      TabIndex        =   17
      Top             =   1560
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
      MICON           =   "frmMenuEditor.frx":0098
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdUpSubMenu 
      Height          =   615
      Left            =   5280
      TabIndex        =   18
      Top             =   2400
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   "Up"
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
      MICON           =   "frmMenuEditor.frx":00B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin nexIRC.ctlXPButton cmdDownSubMenu 
      Height          =   615
      Left            =   5280
      TabIndex        =   19
      Top             =   3120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   "Down"
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
      MICON           =   "frmMenuEditor.frx":00D0
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
      X1              =   3360
      X2              =   3360
      Y1              =   1560
      Y2              =   1920
   End
   Begin VB.Label lblMySubMenus 
      Caption         =   "Sub M&enus:"
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblMyMenu 
      Caption         =   "&Menus:"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblImage 
      Caption         =   "&Image:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   5760
      Y1              =   2055
      Y2              =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   120
      X2              =   5760
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label lblMenu 
      Caption         =   "&Menu:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblCommand 
      Caption         =   "Command:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblCaption 
      Caption         =   "&Caption:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frmMenuEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lSections As Integer
Dim lMenu As String

Private Sub CheckFields()
If lSettings.sHandleErrors Then On Local Error Resume Next
If Len(txtCaption.Text) <> 0 And Len(txtCommand.Text) <> 0 Then
    cmdAddSubMenu.Enabled = True
    cmdDeleteSubMenu.Enabled = True
    If lstSubMenus.Text <> "" Then
        If FindMenuIndex(lMenu, lstMenu.Text, lstSubMenus.Text) <> 0 Then
            cmdAddSubMenu.Caption = "Save"
            cmdAddSubMenu.Enabled = False
        End If
    Else
        cmdAddSubMenu.Caption = "Add"
    End If
Else
    cmdAddSubMenu.Caption = "Add"
    cmdAddSubMenu.Enabled = False
    cmdDeleteSubMenu.Enabled = False
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub CheckFields()"
End Sub

Private Sub cboMenu_Click()
If lSettings.sHandleErrors Then On Local Error Resume Next
Dim i As Integer, msg As String
lstMenu.Clear
lstSubMenus.Clear
Select Case cboMenu.ListIndex
Case 0
    lMenu = GetINIFile(iStatusMenu)
    lSections = ReadINI(lMenu, "Index", "NumSections", 0)
    If lSections <> 0 Then
        For i = 0 To lSections
            msg = ReadINI(lMenu, Trim(Str(i)), "MenuName", "")
            If Len(msg) <> 0 Then
                lstMenu.AddItem msg
            End If
        Next i
    End If
Case 1
    lMenu = GetINIFile(iChannelMenu)
    lSections = ReadINI(lMenu, "Index", "NumSections", 0)
    If lSections <> 0 Then
        For i = 0 To lSections
            msg = ReadINI(lMenu, Trim(Str(i)), "MenuName", "")
            If Len(msg) <> 0 Then
                lstMenu.AddItem msg
            End If
        Next i
    End If
Case 2
    lMenu = GetINIFile(iQueryMenu)
    lSections = ReadINI(lMenu, "Index", "NumSections", 0)
    If lSections <> 0 Then
        For i = 0 To lSections
            msg = ReadINI(lMenu, Trim(Str(i)), "MenuName", "")
            If Len(msg) <> 0 Then
                lstMenu.AddItem msg
            End If
        Next i
    End If
Case 3
    lMenu = GetINIFile(iNicklistMenu)
    lSections = ReadINI(lMenu, "Index", "NumSections", 0)
    If lSections <> 0 Then
        For i = 0 To lSections
            msg = ReadINI(lMenu, Trim(Str(i)), "MenuName", "")
            If Len(msg) <> 0 Then
                lstMenu.AddItem msg
            End If
        Next i
    End If
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cboMenu_Click()"
End Sub

Private Sub cmdAddMenu_Click()
If lSettings.sHandleErrors Then On Local Error Resume Next
Dim msg As String, i As Integer, c As Integer
msg = InputBox("Enter Menu Name: ", "Add Menu", "")
If Len(msg) <> 0 Then
    lSections = lSections + 1
    WriteINI lMenu, "Index", "NumSections", Trim(Str(lSections))
    WriteINI lMenu, Trim(Str(lSections - 1)), "MenuName", msg
    lstMenu.AddItem msg
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdAddMenu_Click()"
End Sub

Private Sub cmdAddSubMenu_Click()
If lSettings.sHandleErrors Then On Local Error Resume Next
Dim i As Integer, c As Integer, j As Integer, msg As String, msg2 As String, msg3 As String
Select Case LCase(cmdAddSubMenu.Caption)
Case "add"
    For i = 0 To lSections
        msg = ReadINI(lMenu, Trim(Str(i)), "MenuName", "")
        If LCase(msg) = LCase(lstMenu.Text) Then
            c = Int(ReadINI(lMenu, Trim(Str(i)), "NumItems", 0))
            c = c + 1
            If c <> 0 Then
                If Len(txtCaption.Text) <> 0 And Len(txtCommand.Text) <> 0 Then
                    WriteINI lMenu, Trim(Str(i)), "NumItems", Trim(Str(c))
                    WriteINI lMenu, Trim(Str(i)), "Item" & Trim(Str(c)), txtCaption.Text
                    WriteINI lMenu, Trim(Str(i)), "Item" & Trim(Str(c)) & "Command", txtCommand.Text
                    If IsNumeric(txtImageIndex.Text) = True Then WriteINI lMenu, Trim(Str(i)), "Item" & Trim(Str(c)) & "Icon", txtImageIndex.Text
                    lstSubMenus.AddItem txtCaption.Text
                End If
                Exit For
            End If
        End If
    Next i
Case "save"
    For i = 0 To lSections
        msg = ReadINI(lMenu, Trim(Str(i)), "MenuName", "")
        If LCase(msg) = LCase(lstMenu.Text) Then
            c = Int(ReadINI(lMenu, Trim(Str(i)), "NumItems", 0))
            For j = 1 To c
                msg2 = ReadINI(lMenu, Trim(Str(i)), "Item" & Trim(Str(j)), "")
                If LCase(msg2) = LCase(txtCaption.Text) Then
                    WriteINI lMenu, Trim(Str(i)), "Item" & Trim(Str(c)), txtCaption.Text
                    WriteINI lMenu, Trim(Str(i)), "Item" & Trim(Str(c)) & "Command", txtCommand.Text
                    If IsNumeric(txtImageIndex.Text) = True Then WriteINI lMenu, Trim(Str(i)), "Item" & Trim(Str(c)) & "Icon", txtImageIndex.Text
                End If
            Next j
        End If
    Next i
End Select
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lstSubMenus_Click()"
End Sub

Private Sub cmdCancel_Click()
If lSettings.sHandleErrors Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdCancel_Click()"
End Sub

Private Sub cmdDeleteMenu_Click()
If lSettings.sHandleErrors Then On Local Error Resume Next
Dim i As Integer, msg As String, c As Integer, j As Integer, b As VbMsgBoxResult
txtCaption.Text = ""
txtCommand.Text = ""
lstSubMenus.Clear
For i = 0 To lSections
    msg = ReadINI(lMenu, Trim(Str(i)), "MenuName", "")
    If LCase(msg) = LCase(lstMenu.Text) Then
        If lSettings.sGeneralPrompts = True Then
            b = MsgBox("Are you sure you wish to remove " & msg, vbYesNo + vbQuestion, "NexIRC")
            If b = vbNo Then
                Exit For
            End If
        End If
        lstMenu.RemoveItem lstMenu.ListIndex
        WriteINI lMenu, Trim(Str(i)), vbNullString, vbNullString
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdDeleteMenu_Click()"
End Sub

Private Sub cmdDeleteSubmenu_Click()
If lSettings.sHandleErrors Then On Local Error Resume Next
Dim i As Integer, c As Integer, j As Integer, msg As String, msg2 As String, msg3 As String, F As Integer
For i = 0 To lSections
    msg = ReadINI(lMenu, Trim(Str(i)), "MenuName", "")
    If LCase(msg) = LCase(lstMenu.Text) Then
        c = Int(ReadINI(lMenu, Trim(Str(i)), "NumItems", 0))
        If c <> 0 Then
            For F = 1 To c
                msg2 = ReadINI(lMenu, Trim(Str(i)), "Item" & Trim(Str(F)), "")
                If Len(msg2) <> 0 Then
                    WriteINI lMenu, Trim(Str(i)), "Item" & Trim(Str(F)), vbNullString
                    WriteINI lMenu, Trim(Str(i)), "Item" & Trim(Str(F)) & "Command", vbNullString
                    WriteINI lMenu, Trim(Str(i)), "Item" & Trim(Str(F)) & "Icon", vbNullString
                    lstSubMenus.RemoveItem F
                    Exit For
                End If
            Next F
        End If
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdDeleteSubmenu_Click()"
End Sub

Private Sub cmdDownSubMenu_Click()
If lSettings.sHandleErrors Then On Local Error Resume Next
Dim lMenuUpName As String, lMenuUpCommand As String, lMenuUpImage As Integer
Dim lMenuDownName As String, lMenuDownCommand As String, lMenuDownImage As Integer
Dim i As Integer, c As Integer, F As Integer, msg As String, l As Integer
If lstSubMenus.ListIndex = lstSubMenus.ListCount Then
    Exit Sub
ElseIf lstSubMenus.ListIndex = -1 Then
    Exit Sub
ElseIf Len(lstSubMenus.Text) <> 0 Then
    For i = 0 To lSections
        msg = ReadINI(lMenu, Trim(Str(i)), "MenuName", "")
        If LCase(msg) = LCase(lstMenu.Text) Then
            c = Int(ReadINI(lMenu, Trim(Str(i)), "NumItems", ""))
            For F = 1 To c
                lMenuDownName = ReadINI(lMenu, Trim(Str(i)), "Item" & Trim(Str(F - 1)), "")
                If LCase(lMenuDownName) = LCase(lstSubMenus.Text) Then
                    lMenuUpName = ReadINI(lMenu, Trim(Str(i)), "Item" & Trim(Str(F)), "")
                    lMenuUpCommand = ReadINI(lMenu, Trim(Str(i)), "Item" & Trim(Str(F)) & "Command", "")
                    lMenuUpImage = ReadINI(lMenu, Trim(Str(i)), "Item" & Trim(Str(F)) & "Icon", "")
                    lMenuDownCommand = ReadINI(lMenu, Trim(Str(i)), "Item" & Trim(Str(F - 1)) & "Command", "")
                    lMenuDownImage = ReadINI(lMenu, Trim(Str(i)), "Item" & Trim(Str(F - 1)) & "Icon", "")
                    WriteINI lMenu, Trim(Str(i)), "Item" & (F - 1), lMenuUpName
                    WriteINI lMenu, Trim(Str(i)), "Item" & (F - 1) & "Command", lMenuUpCommand
                    WriteINI lMenu, Trim(Str(i)), "Item" & (F - 1) & "Icon", Str(lMenuUpImage)
                    WriteINI lMenu, Trim(Str(i)), "Item" & (F), lMenuDownName
                    WriteINI lMenu, Trim(Str(i)), "Item" & (F) & "Command", lMenuDownCommand
                    WriteINI lMenu, Trim(Str(i)), "Item" & (F) & "Icon", Str(lMenuDownImage)
                    l = lstSubMenus.ListIndex + 1
                    lstSubMenus.Clear
                    DoEvents
                    lstMenu_Click
                    lstSubMenus.ListIndex = l
                    Exit For
                End If
            Next F
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdUPMenu_Click()"
End Sub

Private Sub cmdOK_Click()
If lSettings.sHandleErrors Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdOK_Click()"
End Sub

Private Sub cmdUPMenu_Click()

End Sub

Private Sub cmdUpSubMenu_Click()
If lSettings.sHandleErrors Then On Local Error Resume Next
Dim lMenuUpName As String, lMenuUpCommand As String, lMenuUpImage As Integer
Dim lMenuDownName As String, lMenuDownCommand As String, lMenuDownImage As Integer
Dim i As Integer, c As Integer, F As Integer, msg As String, l As Integer
If lstSubMenus.ListIndex = 0 Then
    Exit Sub
ElseIf lstSubMenus.ListIndex = -1 Then
    Exit Sub
ElseIf Len(lstSubMenus.Text) <> 0 Then
    For i = 0 To lSections
        msg = ReadINI(lMenu, Trim(Str(i)), "MenuName", "")
        If LCase(msg) = LCase(lstMenu.Text) Then
            c = Int(ReadINI(lMenu, Trim(Str(i)), "NumItems", ""))
            For F = 1 To c
                lMenuUpName = ReadINI(lMenu, Trim(Str(i)), "Item" & Trim(Str(F)), "")
                If LCase(lMenuUpName) = LCase(lstSubMenus.Text) Then
                    lMenuUpCommand = ReadINI(lMenu, Trim(Str(i)), "Item" & Trim(Str(F)) & "Command", "")
                    lMenuUpImage = Int(ReadINI(lMenu, Trim(Str(i)), "Item" & Trim(Str(F)) & "Icon", 0))
                    lMenuDownName = ReadINI(lMenu, Trim(Str(i)), "Item" & Trim(Str(F - 1)), "")
                    lMenuDownCommand = ReadINI(lMenu, Trim(Str(i)), "Item" & Trim(Str(F - 1)) & "Command", "")
                    lMenuDownImage = Int(ReadINI(lMenu, Trim(Str(i)), "Item" & Trim(Str(F - 1)) & "Icon", 0))
                    WriteINI lMenu, Trim(Str(i)), "Item" & (F - 1), lMenuUpName
                    WriteINI lMenu, Trim(Str(i)), "Item" & (F - 1) & "Command", lMenuUpCommand
                    WriteINI lMenu, Trim(Str(i)), "Item" & (F - 1) & "Icon", Str(lMenuUpImage)
                    WriteINI lMenu, Trim(Str(i)), "Item" & (F), lMenuDownName
                    WriteINI lMenu, Trim(Str(i)), "Item" & (F) & "Command", lMenuDownCommand
                    WriteINI lMenu, Trim(Str(i)), "Item" & (F) & "Icon", Str(lMenuDownImage)
                    l = lstSubMenus.ListIndex - 1
                    lstSubMenus.Clear
                    DoEvents
                    lstMenu_Click
                    lstSubMenus.ListIndex = l
                    Exit For
                End If
            Next F
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub cmdUPMenu_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors Then On Local Error Resume Next
SetButtonType cmdUpSubMenu
SetButtonType cmdDownSubMenu
SetButtonType cmdAddSubMenu
SetButtonType cmdDeleteSubMenu
SetButtonType cmdOK
SetButtonType cmdCancel
SetButtonType cmdAddMenu
SetButtonType cmdDeleteMenu
Me.Icon = mdiNexIRC.Icon
cboMenu.Clear
cboMenu.AddItem "Status Menu"
cboMenu.AddItem "lChannel Menu"
cboMenu.AddItem "lQuery Menu"
cboMenu.AddItem "Nicklist Menu"
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub Form_Load()"
End Sub

Private Sub lstMenu_Click()
If lSettings.sHandleErrors Then On Local Error Resume Next
Dim i As Integer, msg As String, c As Integer, j As Integer
txtCaption.Text = ""
txtCommand.Text = ""
txtImageIndex.Text = ""
lstSubMenus.Clear
For i = 0 To lSections
    msg = ReadINI(lMenu, Trim(Str(i)), "MenuName", "")
    If LCase(msg) = LCase(lstMenu.Text) Then
        c = Int(ReadINI(lMenu, Trim(Str(i)), "NumItems", 0))
        If c <> 0 Then
            For j = 1 To c
                msg = ReadINI(lMenu, Trim(Str(i)), "Item" & Trim(Str(j)), "")
                If Len(msg) <> 0 Then
                    lstSubMenus.AddItem msg
                End If
            Next j
        End If
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lstMenu_Click()"
End Sub

Private Sub lstSubMenus_Click()
If lSettings.sHandleErrors Then On Local Error Resume Next
Dim i As Integer, c As Integer, j As Integer, msg As String
For i = 0 To lSections
    msg = ReadINI(lMenu, Trim(Str(i)), "MenuName", "")
    If LCase(msg) = LCase(lstMenu.Text) Then
        c = Int(ReadINI(lMenu, Trim(Str(i)), "NumItems", 0))
        If c <> 0 Then
            For j = 1 To c
                msg = ReadINI(lMenu, Trim(Str(i)), "Item" & Trim(Str(j)), "")
                If LCase(msg) = LCase(lstSubMenus.Text) Then
                    txtCaption.Text = msg
                    txtCommand.Text = ReadINI(lMenu, Trim(Str(i)), "Item" & Trim(Str(j)) & "Command", "")
                    txtImageIndex.Text = ReadINI(lMenu, Trim(Str(i)), "Item" & Trim(Str(j)) & "Icon", 0)
                End If
            Next j
        End If
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub lstSubMenus_Click()"
End Sub

Private Sub txtCaption_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckFields
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtCaption_Change()"
End Sub

Private Sub txtCaption_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtCaption.SelStart = 0
txtCaption.SelLength = Len(txtCaption.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtCaption_GotFocus()"
End Sub

Private Sub txtCommand_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckFields
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtCommand_Change()"
End Sub

Private Sub txtCommand_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtCommand.SelStart = 0
txtCommand.SelLength = Len(txtCommand.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtCommand_GotFocus()"
End Sub

Private Sub txtImageIndex_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtImageIndex.SelStart = 0
txtImageIndex.SelLength = Len(txtImageIndex.Text)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub txtImageIndex_GotFocus()"
End Sub
