VERSION 5.00
Begin VB.UserControl FormDragger 
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
Attribute VB_Name = "FormDragger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type PointAPI
    x As Long
    y As Long
End Type
Private Const BDR_SUNKENINNER = &H8
Private Const BF_LEFT As Long = &H1
Private Const BF_TOP As Long = &H2
Private Const BF_RIGHT As Long = &H4
Private Const BF_BOTTOM As Long = &H8
Private Const BF_RECT As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BDR_RAISED = &H5
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TOOLWINDOW = &H80
Private Const VK_LBUTTON = &H1
Private Const PS_SOLID = 0
Private Const R2_NOTXORPEN = 10
Private Const BLACK_PEN = 7
Private Const SM_CYCAPTION = 4
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "GDI32" (ByVal nIndex As Long) As Long
Private Declare Function Rectangle Lib "GDI32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetROP2 Lib "GDI32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreatePen Lib "GDI32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Any) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function SetParent Lib "user32" (ByVal HwndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event FormDropped(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)
Event FormMoved(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)
Const m_def_RepositionForm = True
Const m_def_Caption = ""
Dim m_RepositionForm As Boolean
Dim m_Caption As String

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim na As Long
    Dim pt As PointAPI
    Dim frmHwnd As Long
    UserControl_Paint
    frmHwnd = UserControl.Extender.Parent.hWnd
    If Button = vbLeftButton And x >= 0 And x <= UserControl.ScaleWidth And y >= 0 And y <= UserControl.ScaleHeight Then
        ReleaseCapture
        DragObject frmHwnd
    End If
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub DragObject(ByVal hWnd As Long)
    Dim pt As PointAPI
    Dim ptPrev As PointAPI
    Dim objRect As RECT
    Dim DragRect As RECT
    Dim na As Long
    Dim lBorderWidth As Long
    Dim lObjWidth As Long
    Dim lObjHeight As Long
    Dim lXOffset As Long
    Dim lYOffset As Long
    Dim bMoved As Boolean
    ReleaseCapture
    GetWindowRect hWnd, objRect
    lObjWidth = objRect.Right - objRect.Left
    lObjHeight = objRect.Bottom - objRect.Top
    GetCursorPos pt
    ptPrev.x = pt.x
    ptPrev.y = pt.y
    lXOffset = pt.x - objRect.Left
    lYOffset = pt.y - objRect.Top
    With DragRect
        .Left = pt.x - lXOffset
        .Top = pt.y - lYOffset
        .Right = .Left + lObjWidth
        .Bottom = .Top + lObjHeight
    End With
    lBorderWidth = 3
    DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
    Do While GetKeyState(VK_LBUTTON) < 0
        GetCursorPos pt
        If pt.x <> ptPrev.x Or pt.y <> ptPrev.y Then
            ptPrev.x = pt.x
            ptPrev.y = pt.y
            DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
            RaiseEvent FormMoved(pt.x - lXOffset, pt.y - lYOffset, lObjWidth, lObjHeight)
            With DragRect
                .Left = pt.x - lXOffset
                .Top = pt.y - lYOffset
                .Right = .Left + lObjWidth
                .Bottom = .Top + lObjHeight
            End With
            DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
            bMoved = True
        End If
        DoEvents
    Loop
    DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
    If bMoved Then
        If m_RepositionForm Then
            MoveWindow hWnd, DragRect.Left, DragRect.Top, DragRect.Right - DragRect.Left, DragRect.Bottom - DragRect.Top, True
        End If
        RaiseEvent FormDropped(DragRect.Left, DragRect.Top, DragRect.Right - DragRect.Left, DragRect.Bottom - DragRect.Top)
    End If
End Sub

Private Sub DrawDragRectangle(ByVal x As Long, ByVal y As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal lWidth As Long)
    Dim hdc As Long
    Dim hPen As Long
    hPen = CreatePen(PS_SOLID, lWidth, &HE0E0E0)
    hdc = GetDC(0)
    Call SelectObject(hdc, hPen)
    Call SetROP2(hdc, R2_NOTXORPEN)
    Call Rectangle(hdc, x, y, X1, Y1)
    Call SelectObject(hdc, GetStockObject(BLACK_PEN))
    Call DeleteObject(hPen)
    Call SelectObject(hdc, hPen)
    Call ReleaseDC(0, hdc)
End Sub

Private Sub UserControl_InitProperties()
    m_Caption = m_def_Caption
    m_Caption = m_def_Caption
    m_RepositionForm = m_def_RepositionForm
End Sub

Private Sub UserControl_Paint()
    Dim lBackColor As Long
    Dim sCaption As String
    With UserControl
        .Cls
        .Extender.Align = vbAlignTop
        .Extender.Top = 0
        .Height = GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY
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
        .Font.Name = "MS Sans Serif"
        .Font.bold = True
        sCaption = m_Caption
        If UserControl.TextWidth(sCaption) > (UserControl.ScaleWidth - (4 * Screen.TwipsPerPixelX)) Then
            Do While UserControl.TextWidth(sCaption & "...") > (UserControl.ScaleWidth - (4 * Screen.TwipsPerPixelX)) And Len(sCaption) > 0
                sCaption = Trim$(Left$(sCaption, Len(sCaption) - 1))
            Loop
            sCaption = sCaption & "..."
        End If
        UserControl.Print sCaption;
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_RepositionForm = PropBag.ReadProperty("RepositionForm", m_def_RepositionForm)
    UserControl_Paint
End Sub

Private Sub UserControl_Resize()
    UserControl_Paint
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("RepositionForm", m_RepositionForm, m_def_RepositionForm)
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Sets/Returns the caption of the control."
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    UserControl_Paint
End Property

Public Property Get RepositionForm() As Boolean
Attribute RepositionForm.VB_Description = "Specifies whether the control should move the form to it's new location."
    RepositionForm = m_RepositionForm
End Property

Public Property Let RepositionForm(ByVal New_RepositionForm As Boolean)
    m_RepositionForm = New_RepositionForm
    PropertyChanged "RepositionForm"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub
