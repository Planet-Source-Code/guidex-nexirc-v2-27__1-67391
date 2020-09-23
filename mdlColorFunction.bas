Attribute VB_Name = "mdlColorFunction"
Option Explicit
Private Const lColorCode As String = "", lBoldCode As String = "", lPlainCode As String = "", lUnderlineCode As String = "", lReverseChr As String = "", lSpaceCode As String = " "

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
