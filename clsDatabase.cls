VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private lDatabaseFileName As String

Public Property Let FileName(ByVal FileName As String)
Attribute FileName.VB_Description = "Specify a FileName to be used when reading/writing"
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lDatabaseFileName = FileName
End Property

Public Property Get FileName() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FileName = lDatabaseFileName
End Property

Public Sub WriteINIEntry(lSection As String, lKey As String, lItems As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
WriteINI lSection, lKey, lItems, lDatabaseFileName
End Sub

Public Function ReadEntry(lSection As String, lKey As String, lDefault As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = Space$(255)
msg = ReadINI(lDatabaseFileName, lSection, lKey, lDefault)
ReadEntry = StripTerminator(msg)
End Function
