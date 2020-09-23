VERSION 5.00
Begin VB.UserControl ctlFormDragger 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "ctlFormDragger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const lDefRepositionForm = True
Private Const lDefCaption = ""
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private lRepositionForm As Boolean
Private lFormCaption As String
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Any) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event FormDropped(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)
Event FormMoved(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
Dim l As Long, pt As POINTAPI, o As Long
UserControl_Paint
o = UserControl.Extender.Parent.hWnd
If Button = vbLeftButton And X >= 0 And X <= UserControl.ScaleWidth And Y >= 0 And Y <= UserControl.ScaleHeight Then
    ReleaseCapture
    DragObject o
End If
RaiseEvent MouseDown(Button, Shift, X, Y)
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
    Err.Clear
End Sub

Private Sub DragObject(ByVal hWnd As Long)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
Dim pt As POINTAPI, ptPrev As POINTAPI, objRect As RECT, DragRect As RECT, l As Long, lBorderWidth As Long, lObjWidth As Long, lObjHeight As Long, lXOffset As Long, lYOffset As Long, b As Boolean
ReleaseCapture
GetWindowRect hWnd, objRect
lObjWidth = objRect.Right - objRect.Left
lObjHeight = objRect.Bottom - objRect.Top
GetCursorPos pt
ptPrev.X = pt.X
ptPrev.Y = pt.Y
lXOffset = pt.X - objRect.Left
lYOffset = pt.Y - objRect.Top
With DragRect
    .Left = pt.X - lXOffset
    .Top = pt.Y - lYOffset
    .Right = .Left + lObjWidth
    .Bottom = .Top + lObjHeight
End With
lBorderWidth = 3
DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
Do While GetKeyState(&H1) < 0
    GetCursorPos pt
    If pt.X <> ptPrev.X Or pt.Y <> ptPrev.Y Then
        ptPrev.X = pt.X
        ptPrev.Y = pt.Y
        DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
        RaiseEvent FormMoved(pt.X - lXOffset, pt.Y - lYOffset, lObjWidth, lObjHeight)
        With DragRect
            .Left = pt.X - lXOffset
            .Top = pt.Y - lYOffset
            .Right = .Left + lObjWidth
            .Bottom = .Top + lObjHeight
        End With
        DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
        b = True
    End If
    DoEvents
Loop
DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
If b Then
    If lRepositionForm Then
        MoveWindow hWnd, DragRect.Left, DragRect.Top, DragRect.Right - DragRect.Left, DragRect.Bottom - DragRect.Top, True
    End If
    RaiseEvent FormDropped(DragRect.Left, DragRect.Top, DragRect.Right - DragRect.Left, DragRect.Bottom - DragRect.Top)
End If
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub DragObject(ByVal hWnd As Long)"
    Err.Clear
End Sub

Private Sub DrawDragRectangle(ByVal X As Long, ByVal Y As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal lWidth As Long)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
Dim l As Long, o As Long
o = CreatePen(0, lWidth, &HE0E0E0)
l = GetDC(0)
Call SelectObject(l, o)
Call SetROP2(l, 10)
Call Rectangle(l, X, Y, x1, y1)
Call SelectObject(l, GetStockObject(7))
Call DeleteObject(o)
Call SelectObject(l, o)
Call ReleaseDC(0, l)
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub DrawDragRectangle(ByVal x As Long, ByVal y As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal lWidth As Long)"
    Err.Clear
End Sub

Private Sub UserControl_InitProperties()
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
lFormCaption = lDefCaption
lFormCaption = lDefCaption
lRepositionForm = lDefRepositionForm
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub UserControl_InitProperties()"
    Err.Clear
End Sub

Private Sub UserControl_Paint()
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
Dim lBackColor As Long, lCaption As String
With UserControl
    .Cls
    .Extender.Align = vbAlignTop
    .Extender.Top = 0
    .Height = GetSystemMetrics(4) * Screen.TwipsPerPixelY
    If GetActiveWindow = UserControl.Extender.Parent.hWnd Then
        .ForeColor = vbTitleBarText
        lBackColor = vbActiveTitleBar
    Else
        .ForeColor = vbInactiveTitleBarText
        lBackColor = vbInactiveTitleBar
    End If
    UserControl.Line (Screen.TwipsPerPixelX, Screen.TwipsPerPixelY)-(UserControl.ScaleWidth - (2 * Screen.TwipsPerPixelX), UserControl.ScaleHeight - Screen.TwipsPerPixelY), lBackColor, BF
    .CurrentX = 4 * Screen.TwipsPerPixelX
    .CurrentY = 3 * Screen.TwipsPerPixelY
    .Font.Name = "Tahoma"
    .Font.Bold = True
    lCaption = lFormCaption
    If UserControl.TextWidth(lCaption) > (UserControl.ScaleWidth - (4 * Screen.TwipsPerPixelX)) Then
        Do While UserControl.TextWidth(lCaption & "...") > (UserControl.ScaleWidth - (4 * Screen.TwipsPerPixelX)) And Len(lCaption) > 0
            lCaption = Trim$(Left$(lCaption, Len(lCaption) - 1))
        Loop
        lCaption = lCaption & "..."
    End If
    UserControl.Print lCaption;
End With
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub UserControl_Paint()"
    Err.Clear
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
lFormCaption = PropBag.ReadProperty("Caption", lDefCaption)
lRepositionForm = PropBag.ReadProperty("RepositionForm", lDefRepositionForm)
UserControl_Paint
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub UserControl_ReadProperties(PropBag As PropertyBag)"
    Err.Clear
End Sub

Private Sub UserControl_Resize()
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
UserControl_Paint
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub UserControl_Resize()"
    Err.Clear
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
Call PropBag.WriteProperty("Caption", lFormCaption, lDefCaption)
Call PropBag.WriteProperty("RepositionForm", lRepositionForm, lDefRepositionForm)
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub UserControl_WriteProperties(PropBag As PropertyBag)"
    Err.Clear
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Sets/Returns the caption of the control."
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
Caption = lFormCaption
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Property Get Caption() As String"
    Err.Clear
End Property

Public Property Let Caption(ByVal New_Caption As String)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
lFormCaption = New_Caption
PropertyChanged "Caption"
UserControl_Paint
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Property Let Caption(ByVal New_Caption As String)"
    Err.Clear
End Property

Public Property Get RepositionForm() As Boolean
Attribute RepositionForm.VB_Description = "Specifies whether the control should move the form to it's new location."
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
RepositionForm = lRepositionForm
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Property Get RepositionForm() As Boolean"
    Err.Clear
End Property

Public Property Let RepositionForm(ByVal New_RepositionForm As Boolean)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
lRepositionForm = New_RepositionForm
PropertyChanged "RepositionForm"
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Property Let RepositionForm(ByVal New_RepositionForm As Boolean)"
    Err.Clear
End Property

Private Sub UserControl_Click()
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
RaiseEvent Click
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub UserControl_Click()"
    Err.Clear
End Sub

Private Sub UserControl_DblClick()
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
RaiseEvent DblClick
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub UserControl_DblClick()"
    Err.Clear
End Sub

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
hWnd = UserControl.hWnd
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Property Get hWnd() As Long"
    Err.Clear
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
RaiseEvent MouseMove(Button, Shift, X, Y)
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)"
    Err.Clear
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
RaiseEvent MouseUp(Button, Shift, X, Y)
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)"
    Err.Clear
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
If lSettings.sHandleErrors = True Then On Local Error GoTo ErrHandler
UserControl.Refresh
ErrHandler:
    ProcessRuntimeError Err.Description, Err.Number, "Public Sub Refresh()"
    Err.Clear
End Sub
