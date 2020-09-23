VERSION 5.00
Begin VB.Form frmNicklistOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Nicklist Options"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNicklistOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmNicklistOptions.frx":000C
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmNicklistOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Icon = mdiMain.Icon
End Sub
