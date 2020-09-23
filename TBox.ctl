VERSION 5.00
Begin VB.UserControl ctlTBox 
   BackColor       =   &H00000000&
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7035
   DrawWidth       =   51
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   469
   Begin VB.PictureBox ctlDisplay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      Height          =   5490
      Left            =   0
      ScaleHeight     =   366
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   445
      TabIndex        =   0
      Top             =   0
      Width           =   6675
      Begin VB.Timer tmrScrollBar2 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   1080
         Top             =   120
      End
      Begin VB.Timer tmrScrollBar 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   600
         Top             =   120
      End
      Begin VB.Timer tmrClear 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   120
         Top             =   120
      End
   End
   Begin VB.VScrollBar ctlScrollBar 
      Enabled         =   0   'False
      Height          =   5520
      LargeChange     =   20
      Left            =   6780
      Max             =   10
      Min             =   1
      TabIndex        =   1
      Top             =   0
      Value           =   1
      Width           =   240
   End
End
Attribute VB_Name = "ctlTBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const lSizesCount = 132
Private Const lColorCode As String = ""
Private Const lBoldCode As String = ""
Private Const lPlainCode As String = ""
Private Const lUnderlineCode As String = ""
Private Const lReverseChr As String = ""
Private Const lSpaceCode As String = " "
Private Type gStringTable
    sData As String
    sLength As Integer
End Type
Private Type gPointAPI
    X As Long
    Y As Long
End Type
Private Type gRECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type gStyle
    sBold As Boolean
    sUnderline As Boolean
    sTextColor As Long
    sRect As Boolean
    sRectColor As Long
End Type
Private Type gDisplayParam
    sBelongs As Integer
    sLineCount As Integer
    sText As String
End Type
Private lFontStyle As gStyle
Private lSizeV As gPointAPI
Private lDisplayParams(8000) As gDisplayParam
Private lWinRect As gRECT
Private lMenu As clsFMenu
Private lTable() As gStringTable
Private lFontSizes(lSizesCount) As Integer
Private lFontSizes2(lSizesCount) As Integer
Private lLineSize As Integer
Private lLinesUBound As Integer
Private lLinesR As Integer
Private lStartY As Integer
Private lData() As Integer
Private lValue As Integer
Private lMax As Integer
Private lBackBuffer As Long
Private lDispWidth As Long
Private lColorCodes(99) As Long
Private lLineCount As Long
Private H(5) As Long
Private lBitmap As Long
Private mBackColor As Long
Private lBackColor As Long
Private lUrl As String
Private lStringProc As String
Private lResizeControl As Boolean
Private lDoClear As Boolean
Private lInit As Boolean
Private lLineForceRef As Boolean
Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ArrPtr& Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any)
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As gRECT, ByVal hBrush As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As gPointAPI) As Long

Public Function ReturnWidth() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnWidth = ctlDisplay.Width
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function ReturnWidth() As Integer"
End Function

Public Function DefineColorChr(ByVal sText As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, n As Integer, msg As String, msg2 As String, msg3 As String, l As Byte, lParts() As String
lParts = Split(sText, lColorCode)
msg = lParts(0)
For i = 1 To UBound(lParts)
    Select Case True
    Case lParts(i) Like "##,##*" Or lParts(i) Like "##,#*"
        If lParts(i) Like "##,##*" Then l = 2 Else l = 1
        msg2 = Mid(lParts(i), 1, 2)
        msg3 = Mid(lParts(i), 4, l)
        lParts(i) = Replace(lParts(i), msg2 & "," & msg3, vbNullString, , 1)
        msg = msg & lColorCode & LZ(msg2) & LZ(msg3) & lParts(i)
    Case lParts(i) Like "#,##*" Or lParts(i) Like "#,#*"
        If lParts(i) Like "#,##*" Then l = 2 Else l = 1
        msg2 = Mid(lParts(i), 1, 1)
        msg3 = Mid(lParts(i), 3, l)
        lParts(i) = Replace(lParts(i), msg2 & "," & msg3, vbNullString, , 1)
        msg = msg & lColorCode & LZ(msg2) & LZ(msg3) & lParts(i)
    Case lParts(i) Like "#*" Or lParts(i) Like "##*"
        If lParts(i) Like "##*" Then l = 2 Else l = 1
        msg2 = Mid(lParts(i), 1, l)
        lParts(i) = Replace(lParts(i), msg2, vbNullString, , 1)
        msg = msg & lColorCode & LZ(msg2) & "99" & lParts(i)
    Case Else
        msg = msg & lColorCode & "01" & "00" & lParts(i)
    End Select
Next i
DefineColorChr = msg
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function DefineColorChr(ByVal sText As String) As String"
End Function

Private Function LZ(ByVal lData As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
LZ = String(2 - Len(lData), "0") & lData
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Function LZ(ByVal lData As String) As String"
End Function

Public Sub SetFontScript(lScript As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ctlDisplay.Font.Charset = lScript
ctlDisplay.Font.Charset = lScript
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetFontScript(lScript As Integer)"
End Sub

Public Function GetctlScrollBarValue() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
GetctlScrollBarValue = ctlScrollBar.Value
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function GetctlScrollBarValue() As Integer"
End Function

Public Sub InitializeAgain()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UserControl_Initialize
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub InitializeAgain()"
End Sub

Public Sub SetTag(lTag As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UserControl.Tag = lTag
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetTag(lTag As String)"
End Sub

Public Sub SetBorderStyle(lFlat As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lFlat = True Then
    ctlDisplay.BorderStyle = 0
Else
    ctlDisplay.BorderStyle = 1
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetBorderStyle(lFlat As Boolean)"
End Sub

Public Sub SetBackColor(lBackColor As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ctlDisplay.BackColor = lBackColor
mBackColor = CLng(lBackColor)
UserControl.BackColor = CLng(lBackColor)
If lBackColor = 0 Then
    ctlDisplay.Picture = LoadPicture(App.Path & "\data\images\status2.gif")
Else
    ctlDisplay.Picture = LoadPicture(App.Path & "\data\images\status.gif")
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SetBackColor(lBackColor As String)"
End Sub

Public Sub NewLine(ByVal lData As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim n As Integer, i As Integer, c As Integer
lLinesR = lLinesR + 1
If InStr(lData, lColorCode) Then lData = DefineColorChr(lData)
ReDim Preserve lTable(lLinesR)
n = SetTextMetrics(lLinesR, lData)
For i = lLineCount + 1 To lLineCount + n
    lDisplayParams(i).sBelongs = lLinesR
    lDisplayParams(i).sLineCount = n
Next i
FormatString lLineCount + 1, lLineCount + n, ctlDisplay.ScaleWidth, ctlDisplay.hDC
lLineCount = lLineCount + n
lLineForceRef = False
If ctlScrollBar.Value = ctlScrollBar.Max Then
    ctlScrollBar.Max = lLineCount
    ctlScrollBar.Value = lLineCount
    ctlScrollBar.Enabled = True
    lValue = ctlScrollBar.Value
Else
    ctlScrollBar.Max = lLineCount
End If
lMax = lLineCount
lInit = True
lLineForceRef = True
DisplayText
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub NewLine(ByVal lData As String)"
End Sub

Function SetTextMetrics(ByVal tLine As Integer, ByVal strtext As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, iseek As Integer, Res As Single, i2 As Integer, parts() As String
lTable(tLine).sData = strtext
strtext = Replace(strtext, lBoldCode, vbNullString)
strtext = Replace(strtext, lPlainCode, vbNullString)
strtext = Replace(strtext, lUnderlineCode, vbNullString)
strtext = Replace(strtext, lReverseChr, vbNullString)
If InStr(strtext, lColorCode) > 0 Then
    Do
        iseek = InStr(iseek + 1, strtext, lColorCode)
        If iseek > 0 Then
            strtext = Mid(strtext, 1, iseek - 1) & Mid(strtext, iseek + 5, Len(strtext) - iseek + 4)
            iseek = iseek - 1
            If iseek = 0 Then iseek = 1
        End If
    Loop Until iseek = 0
End If
lTable(tLine).sLength = GetTextWidth(strtext)
SetTextMetrics = Int(lTable(tLine).sLength / ctlDisplay.ScaleWidth) + IIf((lTable(tLine).sLength Mod ctlDisplay.ScaleWidth) = 0, 0, 1)
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Function SetTextMetrics(ByVal tLine As Integer, ByVal strtext As String) As Integer"
End Function

Public Function GetTextWidth(ByVal tText As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
GetTextExtentPoint32 ctlDisplay.hDC, tText, Len(tText), lSizeV
GetTextWidth = lSizeV.X
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function GetTextWidth(ByVal tText As String) As Integer"
End Function

Private Sub ctlDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lTopCorrection As Integer, lLeftCorrection As Integer, mbox As VbMsgBoxResult
lDoClear = False
If mdiNexIRC.picTopToolbar.Visible = True Then
    lTopCorrection = (mdiNexIRC.picTopToolbar.Height) / Screen.TwipsPerPixelY + 80
Else
    lTopCorrection = (mdiNexIRC.picTopToolbar.Height) / Screen.TwipsPerPixelY + 50
End If
If mdiNexIRC.mnuFile.Visible = False Then
    lTopCorrection = lTopCorrection - 20
End If
If mdiNexIRC.picMobileMixer.Visible = True Then
    lLeftCorrection = lLeftCorrection + mdiNexIRC.picMobileMixer.Width / Screen.TwipsPerPixelX + 10
Else
    lLeftCorrection = 10
End If
If Button = 2 Then
'    MsgBox UserControl.Tag
    Select Case LCase(UserControl.Tag)
    Case "status"
        If DoesFileExist(GetINIFile(iStatusMenu)) = True Then
            Set lMenu = New clsFMenu
                With lMenu
                    .OwnerHWND = ctlDisplay.hWnd
                    Call .LoadMenus(GetINIFile(iStatusMenu))
                    Call .ShowMenu(X + (mdiNexIRC.Left / Screen.TwipsPerPixelX) + (mdiNexIRC.ActiveForm.Left / Screen.TwipsPerPixelX) + lLeftCorrection, Y + (mdiNexIRC.Top / Screen.TwipsPerPixelY) + (mdiNexIRC.ActiveForm.Top / Screen.TwipsPerPixelY) + lTopCorrection, mdiNexIRC.ActiveForm)
                End With
            Set lMenu = Nothing
        End If
    Case "channel"
'        Stop
        If DoesFileExist(GetINIFile(iChannelMenu)) = True Then
            Set lMenu = New clsFMenu
                With lMenu
                    .OwnerHWND = ctlDisplay.hWnd
                    Call .LoadMenus(GetINIFile(iChannelMenu))
                    Call .ShowMenu(X + (mdiNexIRC.Left / Screen.TwipsPerPixelX) + (mdiNexIRC.ActiveForm.Left / Screen.TwipsPerPixelX) + lLeftCorrection, Y + (mdiNexIRC.Top / Screen.TwipsPerPixelY) + (mdiNexIRC.ActiveForm.Top / Screen.TwipsPerPixelY) + lTopCorrection, mdiNexIRC.ActiveForm)
                End With
            Set lMenu = Nothing
        End If
    Case "query"
        If DoesFileExist(GetINIFile(iQueryMenu)) = True Then
            Set lMenu = New clsFMenu
                With lMenu
                    .OwnerHWND = ctlDisplay.hWnd
                    Call .LoadMenus(GetINIFile(iQueryMenu))
                    Call .ShowMenu(X + (mdiNexIRC.Left / Screen.TwipsPerPixelX) + (mdiNexIRC.ActiveForm.Left / Screen.TwipsPerPixelX) + lLeftCorrection, Y + (mdiNexIRC.Top / Screen.TwipsPerPixelY) + (mdiNexIRC.ActiveForm.Top / Screen.TwipsPerPixelY) + lTopCorrection, mdiNexIRC.ActiveForm)
                End With
            Set lMenu = Nothing
        End If
    End Select
End If
If lUrl <> "" Then
    Surf lUrl, mdiNexIRC.hWnd
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub ctlDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub ctlDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim ln As Integer, strL As String, strR As String, sf As Boolean, strtext As String, iseek As Integer, i2 As Integer
ln = Int((ctlDisplay.ScaleHeight - Y) / lLineSize)
If ln >= lLinesUBound Then
    If ctlScrollBar.Value > 1 Then
        ctlScrollBar.Enabled = True
        ctlScrollBar.Value = ctlScrollBar.Value - 1
        tmrScrollBar = False
        tmrScrollBar2 = True
    End If
ElseIf ln < 0 Then
    If ctlScrollBar.Value < ctlScrollBar.Max Then
        ctlScrollBar.Enabled = True
        ctlScrollBar.Value = ctlScrollBar.Value + 1
        tmrScrollBar = True
        tmrScrollBar2 = False
    End If
Else
    tmrScrollBar2 = False
    tmrScrollBar = False
End If
If ln >= 0 And ln <= ctlDisplay.ScaleHeight \ lLineSize And Sgn(ctlScrollBar.Value - ln) >= 0 And X > 0 And Button = 0 Then
    strtext = lDisplayParams(ctlScrollBar.Value - ln).sText
    RemoveSpecial strtext
    H(3) = StrPtr(strtext)
    Dim lng As Integer, i As Integer
    For i = 0 To Len(strtext) - 1
        On Error Resume Next
        lng = lng + lFontSizes(lData(i))
        If lng >= X Then
            Exit For
        End If
    Next i
    If lng >= X Then
        Dim Q As Integer
        Q = ctlScrollBar.Value - ln
        Do Until sf = True And i >= 0
            For i2 = i To 0 Step -1
                If lData(i2) = 32 Then sf = True: Exit For: sf = True
                strL = strL & ChrW$(lData(i2))
            Next i2
            If sf = False Then
                If lDisplayParams(Q - 1).sBelongs = lDisplayParams(Q).sBelongs Then
                    strtext = lDisplayParams(Q - 1).sText
                    Q = Q - 1
                    RemoveSpecial strtext
                    H(3) = StrPtr(strtext)
                    i = Len(strtext)
                Else
                    sf = True
                End If
            End If
        Loop
    End If
    lng = 0
    sf = False
    strtext = lDisplayParams(ctlScrollBar.Value - ln).sText
    RemoveSpecial strtext
    H(3) = StrPtr(strtext)
    For i = 0 To Len(strtext) - 1
        lng = lng + lFontSizes(lData(i))
        If lng >= X Then
            Exit For
        End If
    Next i
    If lng >= X Then
        Q = ctlScrollBar.Value - ln
        Do Until sf = True
            For i2 = i + 1 To Len(strtext) - 1
                If lData(i2) = 32 Then sf = True: Exit For: sf = True
                strR = strR & ChrW$(lData(i2))
            Next i2
            If sf = False Then
                If lDisplayParams(Q + 1).sBelongs = lDisplayParams(Q).sBelongs Then
                    strtext = lDisplayParams(Q + 1).sText
                    Q = Q + 1
                    RemoveSpecial strtext
                    H(3) = StrPtr(strtext)
                    i = -1
                Else
                    sf = True
                End If
            End If
        Loop
    End If
End If
lUrl = Replace(StrReverse(strL) & strR, Chr(0), "")
If lUrl Like "*www.*.*" Or lUrl Like "*http://*" Or lUrl Like "*@*.*" Then
    ctlDisplay.MousePointer = 99
Else
    lUrl = ""
    ctlDisplay.MousePointer = 0
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub ctlDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub ctlDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lDoClear = True
tmrScrollBar = False
tmrScrollBar2 = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub ctlDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub ctlDisplay_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lResizeControl = True
Dim Res As Single
Dim tkLines As Integer, eLoop As Boolean, th As Single, i As Integer, i2 As Integer, Dispwidth As Integer
Dispwidth = ctlDisplay.ScaleWidth
If Dispwidth <> lDispWidth And lInit Then
    lLineCount = 0
    For i = 1 To lLinesR
        tkLines = Int(lTable(i).sLength / Dispwidth) + IIf((lTable(i).sLength Mod Dispwidth) = 0, 0, 1)
        For i2 = lLineCount + 1 To lLineCount + tkLines
            With lDisplayParams(i2)
                .sBelongs = i
                .sLineCount = tkLines
            End With
        Next i2
        lLineCount = lLineCount + tkLines
    Next i
    lLineForceRef = False
    If lLineCount > 0 Then
        If ctlScrollBar.Value = ctlScrollBar.Max Then
            ctlScrollBar.Max = lLineCount
            lValue = lLineCount
            ctlScrollBar.Enabled = True
            ctlScrollBar.Value = lLineCount
        Else
            ctlScrollBar.Max = lLineCount
        End If
    End If
    lMax = lLineCount
End If
lLineForceRef = True
lDispWidth = Dispwidth
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub ctlDisplay_Resize()"
End Sub

Public Function LoadTextSizes()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim ts As gPointAPI, i As Long
For i = 0 To lSizesCount
    GetTextExtentPoint32 ctlDisplay.hDC, ChrW(i), 1, ts
    lFontSizes(i) = ts.X
Next i
ctlDisplay.FontBold = True
For i = 0 To lSizesCount
    GetTextExtentPoint32 ctlDisplay.hDC, ChrW(i), 1, ts
    lFontSizes2(i) = ts.X
Next i
ctlDisplay.FontBold = False
lLinesUBound = Int(ctlDisplay.ScaleHeight / ts.Y) + 1
lLineSize = ts.Y
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function LoadTextSizes()"
End Function

Private Sub ctlScrollBar_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lLineForceRef Then
    DisplayText
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub ctlScrollBar_Change()"
End Sub

Private Sub ctlScrollBar_Scroll()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lLineForceRef Then
    DisplayText
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub ctlScrollBar_Scroll()"
End Sub

Private Sub tmrClear_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If lDoClear Then
    For i = 0 To lLineCount
        lDisplayParams(i).sText = ""
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrClear_Timer()"
End Sub

Private Sub tmrScrollBar_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If ctlScrollBar.Value < ctlScrollBar.Max Then
    ctlScrollBar.Enabled = True
    ctlScrollBar.Value = ctlScrollBar.Value + 1
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrScrollBar_Timer()"
End Sub

Private Sub tmrScrollBar2_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If ctlScrollBar.Value > 1 Then
    ctlScrollBar.Enabled = True
    ctlScrollBar.Value = ctlScrollBar.Value - 1
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub tmrScrollBar2_Timer()"
End Sub

Private Sub UserControl_Initialize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor
Case "0"
    mBackColor = 0
    ctlDisplay.Picture = LoadPicture(App.Path & "\data\images\status2.gif")
Case Else
    mBackColor = RGB(255, 255, 255)
    ctlDisplay.Picture = LoadPicture(App.Path & "\data\images\status.gif")
End Select
DeleteDC lBackBuffer
DeleteObject lBitmap
lBackColor = CreateSolidBrush(mBackColor)
H(0) = 1
H(1) = 2
H(3) = StrPtr(lStringProc)
H(4) = &H7FFFFFFF
RtlMoveMemory ByVal ArrPtr(lData), VarPtr(H(0)), 4
With lWinRect
    .Left = 0
    .Top = 0
    .Bottom = Screen.Height / Screen.TwipsPerPixelY
    .Right = Screen.Width / Screen.TwipsPerPixelX
End With
lColorCodes(0) = RGB(255, 255, 255)
lColorCodes(1) = RGB(0, 0, 0)
lColorCodes(2) = RGB(0, 0, 127)
lColorCodes(3) = RGB(0, 147, 0)
lColorCodes(4) = RGB(255, 0, 0)
lColorCodes(5) = RGB(127, 0, 0)
lColorCodes(6) = RGB(156, 0, 156)
lColorCodes(7) = RGB(252, 127, 0)
lColorCodes(8) = RGB(255, 255, 0)
lColorCodes(9) = RGB(0, 252, 0)
lColorCodes(10) = RGB(0, 147, 147)
lColorCodes(11) = RGB(0, 255, 255)
lColorCodes(12) = RGB(0, 0, 252)
lColorCodes(13) = RGB(255, 0, 255)
lColorCodes(14) = RGB(127, 127, 127)
lColorCodes(15) = RGB(210, 210, 210)
lColorCodes(99) = RGB(255, 255, 255)
LoadTextSizes
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub UserControl_Initialize()"
End Sub

Private Sub UserControl_Resize()
If lSettings.sHandleErrors Then On Local Error Resume Next
With ctlDisplay
    .Width = UserControl.ScaleWidth - ctlScrollBar.Width
    .Height = UserControl.ScaleHeight
End With
With ctlScrollBar
    .Height = UserControl.ScaleHeight
    .Left = ctlDisplay.Width
End With
lLinesUBound = Int(ctlDisplay.ScaleHeight / lLineSize) + 1
lLineForceRef = True
DisplayText
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub UserControl_Initialize()"
End Sub

Public Sub DisplayText()
If lSettings.sHandleErrors Then On Local Error Resume Next
Dim l As Integer, i As Integer, i2 As Integer, dx As Integer, dx2 As Integer, li As Integer, lBelong As Integer, lp As Integer, isB As Boolean, ib As Integer
If lInit = True And lLineForceRef = True Then
    lWinRect.Bottom = ctlDisplay.ScaleHeight
    lWinRect.Right = ctlDisplay.ScaleWidth
    FillRect ctlDisplay.hDC, lWinRect, lBackColor
    For i = ctlScrollBar.Value To ctlScrollBar.Value - lLinesUBound Step -1
        If Sgn(i) = 1 Then
            l = l + 1
        End If
    Next i
    lStartY = ctlDisplay.ScaleHeight - lLineSize * l
    If lResizeControl = True Then
        lResizeControl = False
        FormatString ctlScrollBar.Value - l + 1, ctlScrollBar.Value, ctlDisplay.ScaleWidth, ctlDisplay.hDC
    Else
    End If
    Start
    For i = ctlScrollBar.Value - l + 1 To ctlScrollBar.Value
        dx = 0
        dx2 = 0
        li = 0
        lp = -1
        If lBelong <> lDisplayParams(i).sBelongs Then
            With lFontStyle
                .sTextColor = lColorCodes(1)
                .sRect = False
                .sBold = False
                .sUnderline = False
            End With
            isB = False
            ctlDisplay.FontBold = False
            ctlDisplay.FontUnderline = False
            Select Case lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sBackColor
            Case "0"
                SetBkColor ctlDisplay.hDC, 0
            Case Else
                SetBkColor ctlDisplay.hDC, RGB(255, 255, 255)
            End Select
            SetTextColor ctlDisplay.hDC, lColorCodes(1)
            lBelong = lDisplayParams(i).sBelongs
        End If
        If lDisplayParams(i).sText <> vbNullString Then
            lStringProc = lDisplayParams(i).sText
            H(3) = StrPtr(lStringProc)
            For i2 = 0 To Len(lStringProc) - 1
                If lData(i2) = 3 Then
                    TextOut ctlDisplay.hDC, dx2, lStartY, Mid$(lDisplayParams(i).sText, li + 1, i2 - li), Len(Mid$(lDisplayParams(i).sText, li + 1, i2 - li))
                    If Not Val(Mid$(lDisplayParams(i).sText, i2 + 4, 2)) = 99 Then
                        SetBkColor ctlDisplay.hDC, lColorCodes(Val(Mid$(lDisplayParams(i).sText, i2 + 4, 2)))
                    End If
                    If Val(Mid$(lDisplayParams(i).sText, i2 + 2, 2)) = 99 Then
                        SetTextColor ctlDisplay.hDC, lColorCodes(1)
                    Else
                        SetTextColor ctlDisplay.hDC, lColorCodes(Val(Mid$(lDisplayParams(i).sText, i2 + 2, 2)))
                    End If
                    dx2 = dx
                    i2 = i2 + 4
                    lp = i2
                    li = i2 + 1
                ElseIf lData(i2) = 2 Then
                    TextOut ctlDisplay.hDC, dx2, lStartY, Mid$(lDisplayParams(i).sText, li + 1, i2 - li), Len(Mid$(lDisplayParams(i).sText, li + 1, i2 - li))
                    ctlDisplay.FontBold = Not ctlDisplay.FontBold
                    isB = Not isB
                    lp = i2
                    li = i2 + 1
                    dx2 = dx
                ElseIf lData(i2) = 31 Then
                    TextOut ctlDisplay.hDC, dx2, lStartY, Mid$(lDisplayParams(i).sText, li + 1, i2 - li), Len(Mid$(lDisplayParams(i).sText, li + 1, i2 - li))
                    ctlDisplay.FontUnderline = Not ctlDisplay.FontUnderline
                    lp = i2
                    li = i2 + 1
                    dx2 = dx
                Else
                    If isB Then
                        dx = dx + lFontSizes2(lData(i2))
                    Else
                        dx = dx + lFontSizes(lData(i2))
                    End If
                End If
            Next i2
        End If
        TextOut ctlDisplay.hDC, dx2, lStartY, Mid$(lDisplayParams(i).sText, lp + 2, i2 - lp), Len(Mid$(lDisplayParams(i).sText, lp + 2, i2 - lp))
        lStartY = lStartY + lLineSize
    Next i
    ctlDisplay.Refresh
End If
End Sub

Public Sub FormatString(ByVal lStart As Integer, ByVal lEnd As Integer, ByVal lDispWidth As Integer, ByVal lhdc As Long)
If lSettings.sHandleErrors Then On Local Error Resume Next
Dim b As Boolean, i As Integer, n As Integer, t As Integer, lLen As Integer, lTotalLength As Integer, lw As Integer, lWidth As Integer, R As Integer, lWS As Integer, lClosed As Boolean, mBool As Boolean
n = lDisplayParams(lStart).sBelongs
For i = lStart To 0 Step -1
    If n <> lDisplayParams(lStart).sBelongs Then
        lStart = i + 1
        i = -1
    End If
Next i
n = 0
For i = lStart To lEnd
    If Not n = lDisplayParams(i).sBelongs Then
        mBool = False
        If lTable(lDisplayParams(i).sBelongs).sLength < lDispWidth Then
            lDisplayParams(i).sText = lTable(lDisplayParams(i).sBelongs).sData
        Else
            b = IIf(lTable(lDisplayParams(i).sBelongs).sData Like "<*> *", False, False)
            lStringProc = lTable(lDisplayParams(i).sBelongs).sData
            H(3) = StrPtr(lStringProc)
            t = 0
            lw = 1
            lWidth = 0
            lWS = 0
            lTotalLength = 0
            R = lDisplayParams(i).sLineCount * lDispWidth - lTable(lDisplayParams(i).sBelongs).sLength
            For t = 0 To Len(lStringProc) - 1
                If lData(t) = 3 Then
                     t = t + 4
                     lWS = lWS + 4
                ElseIf lData(t) = 2 Then
                mBool = Not mBool
                ElseIf lData(t) = 15 Or lData(t) = 31 Or lData(t) = 22 Then
                Else
                    If lData(t) = 32 Then lWidth = lFontSizes(32): lWS = 0
                    lWS = lWS + 1
                    If lData(t) >= 0 Then
                    lTotalLength = lTotalLength + IIf(mBool, lFontSizes2(lData(t)), lFontSizes(lData(t)))
                    lWidth = lWidth + lFontSizes(lData(t))
                    End If
                    If lTotalLength > lDispWidth And t < lDisplayParams(i).sLineCount Then
                      If R - lWidth >= 10000 Then
                        Debug.Print b
                        If b Then
                            lDisplayParams(i + t).sText = "   " & Mid$(lStringProc, lw, t - lw - lWS + 3)
                        Else
                            lDisplayParams(i + t).sText = Mid$(lStringProc, lw, t - lw - lWS + 3)
                        End If
                        R = R - lWidth
                        t = t + 1
                        lw = t + 2 - lWS
                        t = t - lWS + 2
                        lWS = 0
                        lWidth = 0
                        lTotalLength = lFontSizes(32)
                        lw = t + 1
                      Else
                            If t + 1 = lDisplayParams(i).sLineCount Then
                                Dim muvo As Integer, isOne As Boolean
                                
                                For muvo = i To lLineCount
                                If lDisplayParams(muvo).sBelongs = n Then
                                    lDisplayParams(muvo).sLineCount = lDisplayParams(muvo).sLineCount + 1
                                
                                End If
                                Next muvo
                            Else
                        If b And t > 0 Then
                            lDisplayParams(i + t).sText = "   " & Mid$(lStringProc, lw, t - lw + 1)
                        Else
                            lDisplayParams(i + t).sText = Mid$(lStringProc, lw, t - lw + 1)
                        End If
                        lWS = 0
                        lWidth = 0
                        lw = t + 1
                        lTotalLength = IIf(mBool, lFontSizes2(lData(t)), lFontSizes(lData(t)))
                        t = t + 1
                        End If
                      End If
                    End If
                End If
            Next t
            If b And t > 0 Then
                lDisplayParams(i + t).sText = "   " & Mid$(lStringProc, lw, t - lw + 1)
            Else
                lDisplayParams(i + t).sText = Mid$(lStringProc, lw, t - lw + 1)
            End If
        End If
        n = lDisplayParams(i).sBelongs
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub FormatString(ByVal lStart As Integer, ByVal lEnd As Integer, ByVal lDispWidth As Integer, ByVal lHDC As Long)"
End Sub

Private Sub UserControl_Terminate()
If lSettings.sHandleErrors Then On Local Error Resume Next
Dim l As Long
RtlMoveMemory ByVal ArrPtr(lData), 0&, 4
DeleteObject l
DeleteObject lFontStyle.sRectColor
DeleteDC lBackBuffer
DeleteObject lBitmap
DeleteObject lBackColor
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Private Sub UserControl_Terminate()"
End Sub

Public Sub UpdateBar()
If lSettings.sHandleErrors Then On Local Error Resume Next
lLineForceRef = True
ctlScrollBar.Enabled = True
ctlScrollBar.Value = lValue
lLineForceRef = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub UpdateBar()"
End Sub

Public Sub RemoveSpecial(lData As String)
If lSettings.sHandleErrors Then On Local Error Resume Next
Dim i As Integer
lData = Replace(lData, lBoldCode, vbNullString)
lData = Replace(lData, lPlainCode, vbNullString)
lData = Replace(lData, lUnderlineCode, vbNullString)
lData = Replace(lData, lReverseChr, vbNullString)
If InStr(lData, lColorCode) > 0 Then
    Do
        i = InStr(i + 1, lData, lColorCode)
        If i > 0 Then
            lData = Mid(lData, 1, i - 1) & Mid(lData, i + 5, Len(lData) - i + 4)
            i = i - 1
            If i = 0 Then i = 1
        End If
    Loop Until i = 0
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub RemoveSpecial(lData As String)"
End Sub
