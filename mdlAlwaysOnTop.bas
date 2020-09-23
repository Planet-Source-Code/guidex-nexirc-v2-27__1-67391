Attribute VB_Name = "mdlAlwaysOnTop"
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lFlag As Integer
If SetOnTop Then
    lFlag = HWND_TOPMOST
Else
    lFlag = HWND_NOTOPMOST
End If
SetWindowPos myfrm.hWnd, lFlag, myfrm.Left / Screen.TwipsPerPixelX, myfrm.Top / Screen.TwipsPerPixelY, myfrm.Width / Screen.TwipsPerPixelX, myfrm.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)"
End Sub

