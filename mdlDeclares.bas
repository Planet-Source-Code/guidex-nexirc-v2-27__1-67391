Attribute VB_Name = "mdlDeclares"
Option Explicit
Private Type SHELLEXECUTEINFO
     cbSize As Long
     fMask As Long
     hwnd As Long
     lpVerb As String
     lpFile As String
     lpParameters As String
     lpDirectory As String
     nShow As Long
     hInstApp As Long
     lpIDList As Long
     lpClass As String
     hkeyClass As Long
     dwHotKey As Long
     hIcon As Long
     hProcess As Long
End Type
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400
Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long

