Attribute VB_Name = "mdlModes"
Option Explicit
Global MyModes As String

Public Sub OP(strValue As String, UserName As String, Target As String, strModeName As Variant, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, f As Integer, msg As String, l As Integer
UserName = Replace(UserName, ":", "")
For i = 1 To lChannelUBound
    If LCase(lChannelName(i)) = LCase(Target) Then
        For f = 1 To lChannel(i).lstNames.ListItems.Count - 1
            If Trim(strValue) = "+" Then
                If Len(lChannel(i).lstNames.ListItems(f).Text) <> 0 Then
                    msg = lChannel(i).lstNames.ListItems(f).Text
                    If Left(msg, 1) = "+" Or Left(msg, 1) = "@" Then msg = Right(msg, Len(msg) - 1)
                    If LCase(msg) = LCase(strModeName) Then
                        lChannel(i).lstNames.ListItems.Remove f
                        AddUserToNicklist "@" & msg, lChannel(i).lstNames
                        If lSettings.sOptions.oShowModes = True Then
                            DoColor lChannel(i).txtIncoming, "" & Color.Mode & "• " & UserName & " ops " & strModeName
                        Else
                            DoColor lForm.txtIncoming, "" & Color.Mode & "• " & UserName & " ops " & strModeName & " in " & Target
                            DoColorSep lForm.txtIncoming
                        End If
                    End If
                End If
            ElseIf Trim(strValue) = "-" Then
                If Len(lChannel(i).lstNames.ListItems(f).Text) <> 0 Then
                    msg = lChannel(i).lstNames.ListItems(f).Text
                    If Left(msg, 1) = "+" Or Left(msg, 1) = "@" Then msg = Right(msg, Len(msg) - 1)
                    If LCase(msg) = LCase(strModeName) Then
                        lChannel(i).lstNames.ListItems.Remove f
                        AddUserToNicklist msg, lChannel(i).lstNames
                        If lSettings.sOptions.oShowModes = True Then
                            Call DoColor(lChannel(i).txtIncoming, "" & Color.Mode & "• " & UserName & " deops " & strModeName)
                        Else
                            DoColor lForm.txtIncoming, "" & Color.Mode & "• " & UserName & " deops " & strModeName & " in " & Target
                            DoColorSep lForm.txtIncoming
                        End If
                    End If
                End If
            End If
        Next f
    End If
Next i
End Sub

Public Sub VOICE(strValue As String, UserName As String, Target As String, strModeName As Variant, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
Dim x As Integer
Dim f As Integer
Dim msg As String
For i = 1 To lChannelUBound
    If LCase(lChannelName(i)) = LCase(Target) Then
        For x = 1 To lChannel(i).lstNames.ListItems.Count - 1
            If strValue = "+" Then
                If LCase(lChannel(i).lstNames.ListItems(x).Text) = LCase(strModeName) Then
                    msg = lChannel(i).lstNames.ListItems(x).Text
                    lChannel(i).lstNames.ListItems.Remove x
                    AddUserToNicklist "+" & msg, lChannel(i).lstNames
                    If lSettings.sOptions.oShowModes = True Then
                        DoColor lChannel(i).txtIncoming, "" & Color.Mode & "• " & UserName & " adds a voice to " & strModeName
                    Else
                        DoColor lForm.txtIncoming, "" & Color.Mode & "• " & UserName & " adds a voice to " & strModeName & " in " & Target
                        DoColorSep lForm.txtIncoming
                    End If
                End If
            Else
                If LCase(lChannel(i).lstNames.ListItems(x).Text) = "+" & LCase(strModeName) Then
                    lChannel(i).lstNames.ListItems.Remove x
                    AddUserToNicklist str(strModeName), lChannel(i).lstNames
                    lChannel(i).lstNames.ListItems(f).ForeColor = lSpectrumThemes.sSpectrumTheme(lSpectrumThemes.sIndex).sTextColor
                    If lSettings.sOptions.oShowModes = True Then
                        Call DoColor(lChannel(i).txtIncoming, "" & Color.Mode & "• " & UserName & " devoiced " & strModeName)
                    Else
                        DoColor lForm.txtIncoming, "" & Color.Mode & "• " & UserName & " devoiced " & strModeName & " in " & Target
                        DoColorSep lForm.txtIncoming
                    End If
                End If
            End If
        Next x
    End If
Next i
End Sub

Public Sub INVISIBLE(strValue As String, UserName As String, Target As String, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If strValue = "+" Then
    MyModes = MyModes & "i"
Else
    MyModes = Replace(MyModes, "i", "")
End If
lForm.Caption = lForm.Tag & ": [" & MyModes & "] " & lSettings.sNickname & " on " & lSettings.sServer & ":" & lForm.tcp.RemotePort
DoColor lForm.txtIncoming, "" & Color.Mode & "• " & lSettings.sNickname & " sets mode " & strValue & "i"
DoColorSep lForm.txtIncoming
End Sub

Public Sub BAN(strValue As String, UserName As String, Target As String, strModeName As Variant, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To lChannelUBound
    If LCase(lChannelName(i)) = LCase(Target) Then
        If strValue = "+" Then
            If lSettings.sOptions.oShowModes = True Then
                Call DoColor(lChannel(i).txtIncoming, "" & Color.Mode & "• " & UserName & " bans " & strModeName)
            Else
                DoColor lForm.txtIncoming, "" & Color.Mode & "• " & UserName & " bans " & strModeName & " in " & Target
                DoColorSep lForm.txtIncoming
            End If
        Else
            If lSettings.sOptions.oShowModes = True Then
                Call DoColor(lChannel(i).txtIncoming, "" & Color.Mode & "• " & UserName & " unbans " & strModeName)
            Else
                DoColor lForm.txtIncoming, "" & Color.Mode & "• " & UserName & " unbans " & strModeName & " in " & Target
                DoColorSep lForm.txtIncoming
            End If
        End If
    End If
Next i
End Sub

Public Sub Limit(strValue As String, UserName As String, Target As String, strModeName As Variant)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim x As Integer
Dim i As Integer
Select Case strValue
Case "+"
    For i = 1 To lChannelUBound
        If LCase(lChannelName(i)) = LCase(Target) Then
            lChannelLimit(i) = strModeName
            lChannelModes(i) = Replace(lChannelModes(i), "l", "")
            lChannelModes(i) = lChannelModes(i) & "l"
            UpdateCaption i
            Exit For
        End If
    Next i
Case "-"
End Select
End Sub

Public Sub ChangeTopic(xUsername As String, xChannel As String, xTopic As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 1 To lChannelUBound
    If LCase(lChannelName(i)) = LCase(xChannel) Then
        lChannel(i).txtTopic = xTopic
        chanstats(i).Topic = xTopic
        Call DoColor(lChannel(i).txtIncoming, "" & Color.Topic & "• " & xUsername & " changes topic to '" & xTopic & "'")
        lChannel(i).Caption = lChannelName(i) & " [+" & lChannelModes(i) & "] :" & xTopic
    End If
Next i
End Sub

Sub REGISTER(strValue As String, UserName As String, Target As String, lForm As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If strValue = "+" Then
    MyModes = MyModes & "r"
Else
    MyModes = Replace(MyModes, "r", "")
End If
lForm.Caption = lForm.Tag & ": [" & MyModes & "] " & lSettings.sNickname & " on " & lSettings.sServer
DoColor lForm.txtIncoming, "" & Color.Mode & "• " & lSettings.sNickname & " sets mode " & strValue & "r"
DoColorSep lForm.txtIncoming
End Sub
