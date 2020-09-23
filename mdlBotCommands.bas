Attribute VB_Name = "mdlBots"
Option Explicit
Enum eBotType
    bUnknownBot = 0
    bEggdrop = 1
    bX = 2
    bChanServ = 3
    bMemoServ = 4
End Enum
Private Type gBotCommand
    cBotCommand As String
    cBotType As eBotType
End Type
Private Type gBot
    bNickname As String
    bPassword As String
    bType As eBotType
End Type
Private Type gBots
    bCount As Integer
    bBot(150) As gBot
End Type
Private Type gBotCommands
    bCount As Integer
    bBotCommand(150) As gBotCommand
End Type
Private lBots As gBots, lBotCommands As gBotCommands

Public Sub FillComboWithBotCommands(lComboBox As ComboBox, lBotType As eBotType)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To lBotCommands.bCount
    If lBotType = lBotCommands.bBotCommand(i).cBotType And Len(lBotCommands.bBotCommand(i).cBotCommand) <> 0 Then lComboBox.AddItem lBotCommands.bBotCommand(i).cBotCommand
Next i
End Sub

Public Sub FillComboWithBots(lComboBox As ComboBox)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To lBots.bCount
    If Len(lBots.bBot(i).bNickname) <> 0 Then lComboBox.AddItem lBotCommands.bBotCommand(i).cBotCommand
Next i
End Sub

Public Sub SetBotCommandType(lIndex As Integer, lBotType As eBotType)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lBotCommands.bBotCommand(lIndex).cBotType = lBotType
End Sub

Public Function ReturnBotCommandType(lIndex As Integer) As eBotType
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnBotCommandType = lBotCommands.bBotCommand(lIndex).cBotType
End Function

Public Function ReturnBotCommand(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnBotCommand = lBotCommands.bBotCommand(lIndex).cBotCommand
End Function

Public Function ReturnBotType(lIndex As Integer) As eBotType
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnBotType = lBots.bBot(lIndex).bType
End Function

Public Function ReturnBotNickname(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnBotNickname = lBots.bBot(lIndex).bNickname
End Function

Public Function ReturnBotCount() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnBotCount = lBots.bCount
End Function

Public Function ReturnBotCommandCount() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReturnBotCommandCount = lBotCommands.bCount
End Function

Public Sub LoadBots()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, F As Integer
lBots.bCount = ReadINI(GetINIFile(iBots), "Settings", "BotCount", 0)
lBotCommands.bCount = ReadINI(GetINIFile(iBotCommands), "Settings", "CommandCount", 0)
If lBots.bCount <> 0 Then
    For i = 1 To lBots.bCount
        msg = ""
        msg = ReadINI(GetINIFile(iBots), "Bot " & Trim(Str(i)), "Nickname", "")
        If Len(msg) <> 0 Then
            lBots.bBot(i).bNickname = msg
            lBots.bBot(i).bPassword = ReadINI(GetINIFile(iBots), "Bot " & Trim(Str(i)), "Password", "")
            lBots.bBot(i).bType = ReadINI(GetINIFile(iBots), "Bot " & Trim(Str(i)), "Type", "")
            F = F + 1
        End If
    Next i
    lBots.bCount = F
End If
If lBotCommands.bCount <> 0 Then
    For i = 1 To lBotCommands.bCount
        msg = ""
        msg = ReadINI(GetINIFile(iBotCommands), "Command " & Trim(Str(i)), "BotCommand", "")
        If Len(msg) <> 0 Then
            lBotCommands.bBotCommand(i).cBotCommand = msg
            lBotCommands.bBotCommand(i).cBotType = ReadINI(GetINIFile(iBotCommands), "Command " & Trim(Str(i)), "Type", "")
            F = F + 1
        End If
    Next i
    lBotCommands.bCount = F
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub LoadBots()"
End Sub

Public Function FindBotIndex(lNickName As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To lBots.bCount
    If LCase(lBots.bBot(i).bNickname) = LCase(lNickName) Then
        FindBotIndex = i
        Exit For
    End If
Next i
End Function

Public Function FindCommandIndex(lCommand As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To lBotCommands.bCount
    If LCase(lBotCommands.bBotCommand(i).cBotCommand) = LCase(lCommand) Then
        FindCommandIndex = i
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindCommandIndex(lCommand As String) As Integer"
End Function

Public Sub RemoveBotCommand(lCommand As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
i = FindCommandIndex(lCommand)
If i <> 0 Then
    lBotCommands.bBotCommand(i).cBotCommand = ""
    lBotCommands.bBotCommand(i).cBotType = 0
    SaveBots
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub RemoveBot(lBot As String)"
End Sub

Public Sub RemoveBot(lBot As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
i = FindBotIndex(lBot)
If i <> 0 Then
    lBots.bBot(i).bNickname = ""
    lBots.bBot(i).bType = 0
    SaveBots
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub RemoveBot(lBot As String)"
End Sub

Public Function AddBotCommand(lBotCommand As String, lBotType As eBotType) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lBotCommand) <> 0 Then
    i = lBotCommands.bCount + 1
    lBotCommands.bCount = i
    lBotCommands.bBotCommand(i).cBotType = lBotType
    lBotCommands.bBotCommand(i).cBotCommand = lBotCommand
    SaveBotCommands
    AddBotCommand = i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function AddBotCommand(lBotCommand As String, lBotType As eBotType) As Integer"
End Function

Public Function FindOpenBotIndex() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To 150
    If Len(lBots.bBot(i).bNickname) = 0 Then
        FindOpenBotIndex = i
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function FindOpenBotIndex() As Integer"
End Function

Public Function GetBotPassFromRegistry(lBotname As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
If Len(lBotname) <> 0 Then GetBotPassFromRegistry = GetSetting(App.ProductName, lBotname, "Password", "")
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function GetBotPassFromRegistry(lBotname As String) As String"
End Function

Public Sub SaveBotPassToRegistry(lBotname As String, lBotPassword As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(lBotname) <> 0 And Len(lBotPassword) <> 0 Then SaveSetting App.ProductName, lBotname, "Password", lBotPassword
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SaveBotPassToRegistry(lBotname As String, lBotPassword As String)"
End Sub

Public Sub AddBot(lBotname As String, lBotType As eBotType, Optional lPassword As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lBotname) <> 0 Then
    i = lBots.bCount + 1
    If Len(lPassword) = 0 Then
        lPassword = GetBotPassFromRegistry(lBotname)
        If Len(lPassword) = 0 Then lPassword = InputBox("Enter the password for " & lBotname & ".", "Add Bot", "")
        SaveBotPassToRegistry lBotname, lPassword
    End If
    lBots.bCount = i
    lBots.bBot(i).bNickname = lBotname
    lBots.bBot(i).bType = lBotType
    lBots.bBot(i).bPassword = lPassword
    SaveBots
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function AddBot(lBot As String, lBotType As eBotType) As Integer"
End Sub

Public Sub SaveBotCommands()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
WriteINI GetINIFile(iBotCommands), "Settings", "CommandCount", Trim(Str(lBotCommands.bCount))
For i = 1 To lBotCommands.bCount
    WriteINI GetINIFile(iBotCommands), "Command " & Trim(Str(i)), "BotCommand ", lBotCommands.bBotCommand(i).cBotCommand
    WriteINI GetINIFile(iBotCommands), "Command " & Trim(Str(i)), "Type ", lBotCommands.bBotCommand(i).cBotType
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SaveBotCommands()"
End Sub

Public Sub SaveBots()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
WriteINI GetINIFile(iBots), "Settings", "BotCount", Trim(Str(lBots.bCount))
For i = 1 To lBots.bCount
    WriteINI GetINIFile(iBots), "Bot " & Trim(Str(i)), "Nickname", lBots.bBot(i).bNickname
    WriteINI GetINIFile(iBots), "Bot " & Trim(Str(i)), "Type", lBots.bBot(i).bType
    WriteINI GetINIFile(iBots), "Bot " & Trim(Str(i)), "Password", lBots.bBot(i).bPassword
Next i
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub SaveBots()"
End Sub

Public Sub PerformBotCommand(lForm As Form, lBotIndex As Integer, lCommand As String, lValue1 As String, Optional lValue2 As String, Optional lSaveToAutoPerform As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
If Len(lCommand) <> 0 And Len(lValue1) <> 0 Then
    msg = "PRIVMSG " & lBots.bBot(lBotIndex).bNickname & " : " & lCommand
    lForm.tcp.SendData msg & vbCrLf
    ProcessReplaceString sPreformBotCommand, lForm.txtIncoming, lBots.bBot(lBotIndex).bNickname, lCommand
    If lSaveToAutoPerform = True Then
        AddAutoPerform msg
        SaveAutoPerform True
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub PerformBotCommand(lBotIndex As Integer, lCommandIndex As Integer)"
End Sub
