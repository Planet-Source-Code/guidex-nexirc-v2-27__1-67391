Attribute VB_Name = "mdlServerFolder"
Option Explicit
Private Type gServer
    sDescription As String
    sNetwork As Integer
End Type
Private Type gServerFolder
    sServer(150) As gServer
    sCount As Integer
End Type
Global lServerFolder As gServerFolder

Public Function AddToServerFolder(lDescription As String, lNetwork As Integer) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(lDescription) <> 0 And lNetwork <> 0 Then
    lServerFolder.sCount = lServerFolder.sCount + 1
    With lServerFolder.sServer(lServerFolder.sCount)
        .sDescription = lDescription
        .sNetwork = lNetwork
    End With
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Function AddToServerFolder(lDescription As String, lNetwork As Integer) As Integer"
End Function

Public Sub LoadServerFolder()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, c As Integer
lServerFolder.sCount = Int(ReadINI(lINIFiles.iServerFolder, "Settings", "Count", 0))
If lServerFolder.sCount <> 0 Then
    For i = 1 To lServerFolder.sCount
        c = c + 1
        With lServerFolder.sServer(c)
            .sDescription = ReadINI(lINIFiles.iServerFolder, Trim(str(i)), "Description", "")
            If Len(.sDescription) <> 0 Then
                .sNetwork = Int(ReadINI(lINIFiles.iServerFolder, Trim(str(i)), "Network", 0))
            Else
                c = c - 1
            End If
        End With
    Next i
End If
If Err.Number <> 0 Then ProcessRuntimeError Err.Description, Err.Number, "Public Sub LoadServerFolder()"
End Sub
