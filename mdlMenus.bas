Attribute VB_Name = "mdlMenus"
Option Explicit

Public Function FindRootMenuIndex(lMenuFile As String, lRootMenuName As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer, c As Integer
If Len(lRootMenuName) <> 0 Then
    c = Int(ReadINI(lMenuFile, "Index", "NumSections", ""))
    For i = 0 To c
        msg = ReadINI(lMenuFile, Trim(Str(i)), "MenuName", "")
        If Len(msg) <> 0 Then
            If LCase(msg) = LCase(lRootMenuName) Then
                FindRootMenuIndex = i
                Exit For
            End If
        End If
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindRootMenuIndex(lMenuFile As String, lRootMenuName As String) As Integer"
End Function

Public Function FindMenuIndex(lMenuFile As String, lRootMenuName As String, lMenuCaption As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, m As Integer, c As Integer
If Len(lRootMenuName) <> 0 And Len(lMenuCaption) <> 0 Then
    m = FindRootMenuIndex(lMenuFile, lRootMenuName)
    If m <> 0 Then
        c = Int(ReadINI(lMenuFile, Trim(Str(m)), "NumItems", ""))
        If c <> 0 Then
            For i = 1 To c
                msg = ReadINI(lMenuFile, Trim(Str(m)), "Item" & Trim(Str(c)), "")
                If LCase(msg) = LCase(lMenuCaption) Then
                    FindMenuIndex = i
                    Exit For
                End If
            Next i
        End If
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindMenuIndex(lMenuCaption As String) As Integer"
End Function
