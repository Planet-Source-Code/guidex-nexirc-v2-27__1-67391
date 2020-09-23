VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmSysInf 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Information"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSysInf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   8670
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   4800
      Width           =   915
   End
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "Start"
      Height          =   315
      Left            =   7440
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   915
   End
   Begin ComctlLib.ListView lstMultItems 
      Height          =   915
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1614
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Processor"
         Object.Width           =   12347
      EndProperty
   End
   Begin ComctlLib.ListView lstClasses 
      Height          =   2775
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Classes"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ListView lstInfo 
      Height          =   2775
      Left            =   1680
      TabIndex        =   1
      Top             =   1800
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Value"
         Object.Width           =   6586
      EndProperty
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8281
      MultiRow        =   -1  'True
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   16
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Processor"
            Key             =   "Processor"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "BIOS"
            Key             =   "BIOS"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Boot Config"
            Key             =   "Boot Config"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sytem Devices"
            Key             =   "Sytem Devices"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Logical Memory Config"
            Key             =   "Logical Memory Config"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Operating System"
            Key             =   "Operating System"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Processes"
            Key             =   "Processes"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Environment"
            Key             =   "Environment"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "DMA Channel"
            Key             =   "DMA Channel"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "IRQ"
            Key             =   "IRQ"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Device Memory Address"
            Key             =   "Device Memory Address"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab12 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Port Resources"
            Key             =   "Port Resources"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab13 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Services"
            Key             =   "Services"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab14 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "System Drivers"
            Key             =   "System Drivers"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab15 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Time Zone"
            Key             =   "Time Zone"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab16 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "User Accounts"
            Key             =   "User Accounts"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSysInf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***********************************************************************
'This application and its components were explicitly developed for
'PSC(Planet Source Code) Users as Open Source Projects.
'This code and the code of its components are property of their author.
'
'If you compile this application you may not redistribute it.
'However, you may use any of this code in you're own application(s).
'
'Alex Smoljanovic, Salex Software (c) 2001-2003
'salex_software@shaw.ca
'***********************************************************************


Dim objShell As Object
'Dimensionalize or declare variable objShell as an object who's methods and properties are inherited by the objects class which is returned by the GetObject or CreateObject function

Private Sub cmdStartStop_Click()
Dim itmX As ListItem 'dimensionalize itmX as ListItem type structure
 Set itmX = lstInfo.ListItems.Item("name")
 'initialize itmX with the list item who's key evaluates to "name"
    If LCase(cmdStartStop.Caption) = "start" Then
    'Determine which action should be performed, starting or stopping a service
     If MsgBox("Are you sure you want to start the service """ & itmX.SubItems(1) & """?", vbQuestion + vbYesNo) = vbYes Then
     'Request action confirmation
      Set objShell = CreateObject("Shell.application")
      'Initialize objShell with the return of the CreateObject function
       If objShell.ServiceStart(itmX.SubItems(1), True) = True Then
       'If the function return evaluates to true the method called was successful(the sequences which start a service have started)
        MsgBox "The service was started.", vbInformation, "ProcessXP"
        'Inform user of the services status change, note that they may be a an intermediate status as the service is being prepared
         lstMultItems_ItemClick lstMultItems.SelectedItem
         'See sub routine lstMultiItems_ItemClick for more info...
       Else
        MsgBox "The service couldn't be started.", vbExclamation, "ProcessXP"
        'The function return evaluated to false, so was the success of the action
       End If
        Set objShell = Nothing
        'Terminate object objShell...
     End If
    Else
     If MsgBox("Are you sure you want to stop the service """ & itmX.SubItems(1) & """?", vbQuestion + vbYesNo) = vbYes Then
     'Request confirmation...
      Set objShell = CreateObject("Shell.application")
      'initialize object...
       If objShell.ServiceStop(itmX.SubItems(1), True) = True Then
       'Function was successful(the sequences which stop a service have started)
        MsgBox "The service was stopped.", vbInformation, "ProcessXP"
        'Inform user
         lstMultItems_ItemClick lstMultItems.SelectedItem
         'See this sub routine for more info...
       Else
        MsgBox "The service couldn't be stopped.", vbExclamation, "ProcessXP"
        'Method was unsucessful...
       End If
        Set objShell = Nothing 'terminate object(no longer needed)
     End If
    End If
End Sub

Private Sub Command1_Click()
 Unload Me 'unload this form instance
End Sub

Private Sub Form_Activate()
' TransPrep Me.hWnd 'see function TransPrep for more info...
End Sub

Private Sub Form_Load()
On Error Resume Next
'On the event of an error resume execution of this procedure on the next line
 TabStrip_Click 'see TabStrip_Click sub routine for more info...
  If HostOS.OperatingSystem = fullCompatibleOS Then
  'If the variable HostOS's OperatingSystem member evaluates to the value of the constant fullCompatibleOS as it will when the host operating system is windowsxp or greater (windows 2003 .net server family[compatibility not tested])
   'VGrad picBot.hdc, picBot.ScaleHeight, picBot.ScaleWidth, COLOR_BTNFACE, Black2White
   'See vgrad function (mdlGUI), this is the intermediate function which formats the arguments to call the appropriate function in the ProcXPGUI DLL
  Else
   'VGrad picBot.hdc, picBot.ScaleHeight, picBot.ScaleWidth, COLOR_BTNFACE, white2black
   'See VGrad(mdlGUI) function for more info...
  End If
End Sub

Private Sub lstClasses_ItemClick(ByVal Item As ComctlLib.ListItem)
 QueryInfo Item.Key
 'See QueryInfo function for more info...
End Sub

Private Sub lstMultItems_ItemClick(ByVal Item As ComctlLib.ListItem)
 QueryInfo Trim(lstClasses.SelectedItem.Key), Mid$(Item.Key, 2)
 'See QueryInfo function for more info...
End Sub

Public Sub TabStrip_Click()
 Select Case TabStrip.SelectedItem.Key
  Case "Processor":
  'If the TabStrip's currently selected tab's key prop. evaluates to "Processor" then...
   QueryInfo "Win32_Processor"
   'See queryinfo for more info...
  Case "BIOS":
   QueryInfo "Win32_BIOS"
  Case "Boot Config":
   QueryInfo "Win32_BootConfiguration"
  Case "Sytem Devices":
   QueryInfo "Win32_DiskPartition"
  Case "Logical Memory Config":
   QueryInfo "Win32_LogicalMemoryConfiguration"
  Case "Operating System":
   QueryInfo "Win32_OperatingSystem"
  Case "Processes":
   QueryInfo "Win32_Process"
  Case "Environment":
   QueryInfo "Win32_Environment"
  Case "DMA Channel":
   QueryInfo "Win32_DMAChannel"
  Case "IRQ":
   QueryInfo "Win32_IRQResource"
  Case "Device Memory Address":
   QueryInfo "Win32_DeviceMemoryAddress"
  Case "Port Resources":
   QueryInfo "Win32_PortResource"
  Case "Services":
   QueryInfo "Win32_Service"
  Case "System Drivers":
   QueryInfo "Win32_SystemDriver"
  Case "Time Zone":
   QueryInfo "Win32_TimeZone"
  Case "User Accounts":
   QueryInfo "Win32_UserAccount"
 End Select
End Sub

'NOTE, for user friendliness reasons every listitem populated
'in the list is 'hard coded' this resulted in creating extremely large procedure,
'this was done because the items property names are unrecognizable to some users and consist of
'no spaces and or are grammatically incorrect.

'Additionally some values returned are as arrays of string or numbers or are
'numbers or strings, even though the use of VarType and TypeName would resolve such
'problems I have still decided to 'do it the long way'

'So instead of something similar to:
' for each item in item
'  List.Add item
' next item

'The following practice was used
' List.Add "Item1 Name"
'  List.Add "Item2 Name"
'   List.Add "Item3 Name"
'    List.Add "Item4 Name"
'ect...
'A very un-intersting and repetative process, but to the end user its appreciated(hopefully)

Sub GetBIOS(Optional MultiItem$)
On Error Resume Next
'On the even of an error resume execution of this procedure on the next line
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
'Dimensionalize objWMI as an object, objItem as an object, itmX as ListItem type structure, tmpCnt as long data type, tmpBuffer as string data type
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
 'MultiItem's argument value is only initialized when a specified item is being requested
   lstInfo.ListItems.Clear 'Remove all ListItems from the list item control
     Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
     'Initialize objWMI to the return of the GetObject function
     'Format "winmgmts:\\COMPUTERNAME\root\cimv2"
     'If you wish to ommit the computer name from this query than replace it with a dot "." character to specify the local machine
       Set objItem = objWMI.execquery("Select * from Win32_BIOS", , 48)
       'initialize objItem with the return of the object objWMI's execquery method
       'The query is composed of a standard query;
       'Select *(All Elements) from COLUMN
       'To select a specific element based on a value;
       'Select * from COLUMN where PROPERTY = "VALUE"
       'Select * from Clowns where name = "Bozo"
        For Each Item In objItem
        'For Each loop; loops through every item in objItem
         With Item 'With Block statement(where Object Identifier is omitted[.Method])
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Basic Input Output Systems": lstMultItems.ListItems.Add , "d" & .SoftwareElementID, .SoftwareElementID
          'increment tmpCnt by one, update listviews column headers text properties
           If Trim(MultiItem$) = Trim(CStr(Item.SoftwareElementID)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Description")
           itmX.SubItems(1) = .Description
            Set itmX = lstInfo.ListItems.Add(, , "Build Number")
            itmX.SubItems(1) = CStr(.BuildNumber)
             Set itmX = lstInfo.ListItems.Add(, , "Code Set")
             itmX.SubItems(1) = CStr(.CodeSet)
              Set itmX = lstInfo.ListItems.Add(, , "Current Language")
              itmX.SubItems(1) = CStr(.CurrentLanguage)
               Set itmX = lstInfo.ListItems.Add(, , "Identification Code")
               itmX.SubItems(1) = CStr(.IdentificationCode)
                Set itmX = lstInfo.ListItems.Add(, , "Installable Languages")
                itmX.SubItems(1) = CStr(.InstallableLanguages)
                 Set itmX = lstInfo.ListItems.Add(, , "Language Edition")
                 itmX.SubItems(1) = CStr(.LanguageEdition)
                  Set itmX = lstInfo.ListItems.Add(, , "Manufacturer")
                  itmX.SubItems(1) = CStr(.Manufacturer)
                   Set itmX = lstInfo.ListItems.Add(, , "Name")
                   itmX.SubItems(1) = CStr(.Name)
                    Set itmX = lstInfo.ListItems.Add(, , "Other Target Operating System")
                    itmX.SubItems(1) = .OtherTargetOS
                     Set itmX = lstInfo.ListItems.Add(, , "Primary BIOS?")
                     itmX.SubItems(1) = CBoolStr(.PrimaryBIOS)
                     'Function CBoolStr converts a number to a boolean value as a string: "True" or "False" not an actual boolean value
                      Set itmX = lstInfo.ListItems.Add(, , "Release Date")
                      itmX.SubItems(1) = CStr(Day(.ReleaseDate)) & "\" & CStr(Month(.ReleaseDate)) & "\" & CStr(Year(.ReleaseDate)) & " " & CStr(Hour(.ReleaseDate)) & ":" & CStr(Minute(.ReleaseDate))
                       Set itmX = lstInfo.ListItems.Add(, , "Serial Number")
                       itmX.SubItems(1) = CStr(.SerialNumber)
                        Set itmX = lstInfo.ListItems.Add(, , "SMBIOS BIOS Version")
                        itmX.SubItems(1) = CStr(.SMBIOSBIOSVersion)
                         Set itmX = lstInfo.ListItems.Add(, , "SMBIOS Major Version")
                         itmX.SubItems(1) = CStr(.SMBIOSMajorVersion)
                          Set itmX = lstInfo.ListItems.Add(, , "SMBIOS Minor Version")
                          itmX.SubItems(1) = CStr(.SMBIOSMinorVersion)
                           Set itmX = lstInfo.ListItems.Add(, , "SMBIOS Present")
                           itmX.SubItems(1) = CBoolStr(.SMBIOSPresent)
                            Set itmX = lstInfo.ListItems.Add(, , "Software Element ID")
                            itmX.SubItems(1) = .SoftwareElementID
                             Set itmX = lstInfo.ListItems.Add(, , "Software Element State")
                             itmX.SubItems(1) = CStr(.SoftwareElementState)
                              Set itmX = lstInfo.ListItems.Add(, , "Status")
                              itmX.SubItems(1) = CStr(.Status)
                               Set itmX = lstInfo.ListItems.Add(, , "TargetOperatingSystem")
                               itmX.SubItems(1) = CStr(.TargetOperatingSystem)
                                Set itmX = lstInfo.ListItems.Add(, , "Version")
                                itmX.SubItems(1) = .Version
           End If
         End With
        Next Item 'Select next item in the collection
         lstClasses.ListItems.Clear 'Clear the lstClasses listview control(Remove all pre-existing items)
          lstClasses.ListItems.Add , "Win32_BIOS", "System BIOS" 'Add an item who's Key will identify from which column to query
           Set objWMI = Nothing 'Terminate object objWMI
            Set objItem = Nothing 'Terminate object...
End Sub

Sub GetProcessorInfo(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
     Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
       Set objItem = objWMI.execquery("Select * from Win32_Processor", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Processors": lstMultItems.ListItems.Add , "d" & .DeviceID, .DeviceID
           If Trim(MultiItem$) = Trim(CStr(Item.DeviceID)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Description")
           itmX.SubItems(1) = .Description
            Set itmX = lstInfo.ListItems.Add(, , "Architecture")
            itmX.SubItems(1) = CStr(.Architecture)
             Set itmX = lstInfo.ListItems.Add(, , "Availability")
             itmX.SubItems(1) = CStr(.Availability)
              Set itmX = lstInfo.ListItems.Add(, , "Address Width")
              itmX.SubItems(1) = CStr(.AddressWidth)
               Set itmX = lstInfo.ListItems.Add(, , "CPU Status")
               itmX.SubItems(1) = CStr(.CpuStatus)
                Set itmX = lstInfo.ListItems.Add(, , "Current Clock Speed")
                itmX.SubItems(1) = CStr(.CurrentClockSpeed) ' & "MHz" *** I am unfamiliar with voltage measurement units
                 Set itmX = lstInfo.ListItems.Add(, , "Current Voltage")
                 itmX.SubItems(1) = CStr(.CurrentVoltage)
                  Set itmX = lstInfo.ListItems.Add(, , "Data Width")
                  itmX.SubItems(1) = CStr(.DataWidth)
                   Set itmX = lstInfo.ListItems.Add(, , "Device ID")
                   itmX.SubItems(1) = CStr(.DeviceID)
                    Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                    itmX.SubItems(1) = .ErrorDescription
                     Set itmX = lstInfo.ListItems.Add(, , "Ext Clock")
                     itmX.SubItems(1) = CStr(.ExtClock)
                      Set itmX = lstInfo.ListItems.Add(, , "Family")
                      itmX.SubItems(1) = CStr(.Family)
                       Set itmX = lstInfo.ListItems.Add(, , "L2 Cache Size")
                       itmX.SubItems(1) = FormatByteSize(.L2CacheSize)
                        Set itmX = lstInfo.ListItems.Add(, , "L2 Cache Speed")
                        itmX.SubItems(1) = FormatByteSize(.L2CacheSpeed)
                         Set itmX = lstInfo.ListItems.Add(, , "Level")
                         itmX.SubItems(1) = CStr(.level)
                          Set itmX = lstInfo.ListItems.Add(, , "Load Percentage")
                          itmX.SubItems(1) = CStr(.LoadPercentage)
                           Set itmX = lstInfo.ListItems.Add(, , "Manufacturer")
                           itmX.SubItems(1) = .Manufacturer
                            Set itmX = lstInfo.ListItems.Add(, , "Max Clock Speed")
                            itmX.SubItems(1) = CStr(.MaxClockSpeed)
                             Set itmX = lstInfo.ListItems.Add(, , "Name")
                             itmX.SubItems(1) = .Name
                              Set itmX = lstInfo.ListItems.Add(, , "Other Family Description")
                              itmX.SubItems(1) = .OtherFamilyDescription
                               Set itmX = lstInfo.ListItems.Add(, , "Plug-N-Play DeviceID")
                               itmX.SubItems(1) = CStr(.PNPDeviceID)
                                Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported")
                                itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                                 Set itmX = lstInfo.ListItems.Add(, , "Processor ID")
                                 itmX.SubItems(1) = .ProcessorId
                                  Set itmX = lstInfo.ListItems.Add(, , "Processor Type")
                                  itmX.SubItems(1) = CStr(.ProcessorType)
                                   Set itmX = lstInfo.ListItems.Add(, , "Revision")
                                   itmX.SubItems(1) = CStr(.Revision)
                                    Set itmX = lstInfo.ListItems.Add(, , "Role")
                                    itmX.SubItems(1) = .Role
                                     Set itmX = lstInfo.ListItems.Add(, , "Socket Designation")
                                     itmX.SubItems(1) = .SocketDesignation
                                      Set itmX = lstInfo.ListItems.Add(, , "Status")
                                      itmX.SubItems(1) = .Status
                                       Set itmX = lstInfo.ListItems.Add(, , "Status Info")
                                       itmX.SubItems(1) = CStr(.StatusInfo)
                                        Set itmX = lstInfo.ListItems.Add(, , "Stepping")
                                        itmX.SubItems(1) = CStr(.Stepping)
                                         Set itmX = lstInfo.ListItems.Add(, , "System Name")
                                         itmX.SubItems(1) = .SystemName
                                          Set itmX = lstInfo.ListItems.Add(, , "Unique ID")
                                          itmX.SubItems(1) = .UniqueId
                                           Set itmX = lstInfo.ListItems.Add(, , "Upgrade Method")
                                           itmX.SubItems(1) = CStr(.UpgradeMethod)
                                            Set itmX = lstInfo.ListItems.Add(, , "Version")
                                            itmX.SubItems(1) = CStr(.Version)
                                             Set itmX = lstInfo.ListItems.Add(, , "Voltage Capabilities")
                                             itmX.SubItems(1) = CStr(.VoltageCaps)
                                          
            
           End If
         End With
        Next Item
         lstClasses.ListItems.Clear
          lstClasses.ListItems.Add , "Win32_Processor", "Processors"
           Set objWMI = Nothing
            Set objItem = Nothing
End Sub

Sub GetBootConfig(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
     Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
       Set objItem = objWMI.execquery("Select * from Win32_BootConfiguration", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Boot Configurations": lstMultItems.ListItems.Add , "d" & .Name, .Name
           If Trim(MultiItem$) = Trim(CStr(Item.Name)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Boot Directory")
           itmX.SubItems(1) = .BootDirectory
            Set itmX = lstInfo.ListItems.Add(, , "Configuration Path")
            itmX.SubItems(1) = CStr(.ConfigurationPath)
             Set itmX = lstInfo.ListItems.Add(, , "Description")
             itmX.SubItems(1) = CStr(.Description)
              Set itmX = lstInfo.ListItems.Add(, , "Last Drive")
              itmX.SubItems(1) = CStr(.LastDrive)
               Set itmX = lstInfo.ListItems.Add(, , "Name")
               itmX.SubItems(1) = CStr(.Name)
                Set itmX = lstInfo.ListItems.Add(, , "Scratch Directory")
                itmX.SubItems(1) = CStr(.ScratchDirectory)
                 Set itmX = lstInfo.ListItems.Add(, , "Setting ID")
                 itmX.SubItems(1) = CStr(.SettingID)
                  Set itmX = lstInfo.ListItems.Add(, , "Temp Directory")
                  itmX.SubItems(1) = CStr(.TempDirectory)
           End If
         End With
        Next Item
         lstClasses.ListItems.Clear
          lstClasses.ListItems.Add , "Win32_BootConfiguration", "Boot Config"
           Set objWMI = Nothing
            Set objItem = Nothing
End Sub

Sub GetDiskPartition(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
     Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
       Set objItem = objWMI.execquery("Select * from Win32_DiskPartition", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Disk Partitions": lstMultItems.ListItems.Add , "d" & CStr(.DeviceID), .DeviceID
          If Trim(MultiItem$) = Trim(CStr(Item.DeviceID)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Access")
           itmX.SubItems(1) = CStr(.Access)
            Set itmX = lstInfo.ListItems.Add(, , "Availability")
            itmX.SubItems(1) = CStr(.Availability)
             Set itmX = lstInfo.ListItems.Add(, , "Block Size")
             itmX.SubItems(1) = FormatByteSize(.BlockSize)
              Set itmX = lstInfo.ListItems.Add(, , "Bootable")
              itmX.SubItems(1) = CBoolStr(.Bootable)
               Set itmX = lstInfo.ListItems.Add(, , "Boot Partition")
               itmX.SubItems(1) = CBoolStr(.BootPartition)
                Set itmX = lstInfo.ListItems.Add(, , "Config Manager Error Code")
                itmX.SubItems(1) = CStr(.ConfigManagerErrorCode)
                 Set itmX = lstInfo.ListItems.Add(, , "Config Manager User Config")
                 itmX.SubItems(1) = CStr(.ConfigManagerUserConfig)
                  Set itmX = lstInfo.ListItems.Add(, , "Description")
                  itmX.SubItems(1) = CStr(.Description)
                   Set itmX = lstInfo.ListItems.Add(, , "Device ID")
                   itmX.SubItems(1) = CStr(.DeviceID)
                    Set itmX = lstInfo.ListItems.Add(, , "Disk Index")
                    itmX.SubItems(1) = CStr(.DiskIndex)
                     Set itmX = lstInfo.ListItems.Add(, , "Error Cleared")
                     itmX.SubItems(1) = CBoolStr(.ErrorCleared)
                      Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                      itmX.SubItems(1) = CStr(.ErrorDescription)
                       Set itmX = lstInfo.ListItems.Add(, , "Error Methodology")
                       itmX.SubItems(1) = CStr(.ErrorMethodology)
                        Set itmX = lstInfo.ListItems.Add(, , "Hidden Sectors")
                        itmX.SubItems(1) = GroupDigits(.HiddenSectors)
                         Set itmX = lstInfo.ListItems.Add(, , "Index")
                         itmX.SubItems(1) = CStr(.Index)
                          Set itmX = lstInfo.ListItems.Add(, , "Last Error Code")
                          itmX.SubItems(1) = CStr(.LastErrorCode)
                           Set itmX = lstInfo.ListItems.Add(, , "Name")
                           itmX.SubItems(1) = .Name
                            Set itmX = lstInfo.ListItems.Add(, , "Number Of Blocks")
                            itmX.SubItems(1) = GroupDigits(.NumberOfBlocks)
                             Set itmX = lstInfo.ListItems.Add(, , "PlugNPlay Device ID")
                             itmX.SubItems(1) = CStr(.PNPDeviceID)
                              Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported")
                              itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                               Set itmX = lstInfo.ListItems.Add(, , "PrimaryPartition")
                               itmX.SubItems(1) = CBoolStr(.PrimaryPartition)
                                Set itmX = lstInfo.ListItems.Add(, , "Purpose")
                                itmX.SubItems(1) = CStr(.Purpose)
                                 Set itmX = lstInfo.ListItems.Add(, , "Rewrite Partition")
                                 itmX.SubItems(1) = CBoolStr(.RewritePartition)
                                  Set itmX = lstInfo.ListItems.Add(, , "Size")
                                  itmX.SubItems(1) = FormatByteSize(.Size)
                                   Set itmX = lstInfo.ListItems.Add(, , "Starting Offset")
                                   itmX.SubItems(1) = CStr(.StartingOffset)
                                    Set itmX = lstInfo.ListItems.Add(, , "Status")
                                    itmX.SubItems(1) = .Status
                                     Set itmX = lstInfo.ListItems.Add(, , "StatusInfo")
                                     itmX.SubItems(1) = CStr(.StatusInfo)
                                      Set itmX = lstInfo.ListItems.Add(, , "Status")
                                      itmX.SubItems(1) = .Status
                                       Set itmX = lstInfo.ListItems.Add(, , "SystemName")
                                       itmX.SubItems(1) = CStr(.SystemName)
                                        Set itmX = lstInfo.ListItems.Add(, , "Type")
                                        itmX.SubItems(1) = CStr(.Type)

           End If
         End With
        Next Item
         lstClasses.ListItems.Clear
          lstClasses.ListItems.Add , "Win32_DiskPartition", "Disk Partitions"
           lstClasses.ListItems.Add , "Win32_CDROMDrive", "CDROM Drives"
            lstClasses.ListItems.Add , "Win32_DiskDrive", "Disk Drives"
             lstClasses.ListItems.Add , "Win32_FloppyController", "Floppy Controllers"
              lstClasses.ListItems.Add , "Win32_IDEController", "IDE Controllers"
               lstClasses.ListItems.Add , "Win32_Keyboard", "Keyboards"
                lstClasses.ListItems.Add , "Win32_MotherBoardDevice", "Mother Board Device"
                 lstClasses.ListItems.Add , "Win32_NetworkAdapter", "Network Adapters"
                  lstClasses.ListItems.Add , "Win32_PnPEntity", "Plug-N-Play Entities"
                   lstClasses.ListItems.Add , "Win32_PotsModem", "POTS Modems"
                    lstClasses.ListItems.Add , "Win32_Printer", "Printers"
                     lstClasses.ListItems.Add , "Win32_SoundDevice", "Sound Devices"
                      lstClasses.ListItems.Add , "Win32_LogicalDisk", "Logical Disks"
                       lstClasses.ListItems.Add , "Win32_ParallelPort", "Parallel Ports"
                        lstClasses.ListItems.Add , "Win32_SerialPort", "Serial Ports"
                         lstClasses.ListItems.Add , "Win32_USBController", "USB Controllers"
                          lstClasses.ListItems.Add , "Win32_USBHub", "USB Hubs"
                           lstClasses.ListItems.Add , "Win32_DesktopMonitor", "Desktop Monitors"
                            lstClasses.ListItems.Add , "Win32_VideoController", "Video Controllers"
                             Set objWMI = Nothing
                              Set objItem = Nothing
End Sub

Sub GetCDROM(MultiItem$)
Stop
'On Error Resume Next
 Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
  If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
    lstInfo.ListItems.Clear
      Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
       Set objItem = objWMI.execquery("Select * from Win32_CDROMDrive", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "CDROM Drives": lstMultItems.ListItems.Add , "d" & CStr(.DeviceID), "Drive: " & .ID
          If Trim(MultiItem$) = Trim(CStr(Item.DeviceID)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Availability")
           itmX.SubItems(1) = CStr(.Availability)
            Set itmX = lstInfo.ListItems.Add(, , "Compression Method")
            itmX.SubItems(1) = CStr(.CompressionMethod)
             Set itmX = lstInfo.ListItems.Add(, , "Config Manager Error Code")
             itmX.SubItems(1) = CStr(.ConfigManagerErrorCode)
              Set itmX = lstInfo.ListItems.Add(, , "Config Manager User Config")
              itmX.SubItems(1) = CBoolStr(.ConfigManagerUserConfig)
               Set itmX = lstInfo.ListItems.Add(, , "Default Block Size")
               itmX.SubItems(1) = FormatByteSize(.DefaultBlockSize)
                Set itmX = lstInfo.ListItems.Add(, , "Description")
                itmX.SubItems(1) = CStr(.Description)
                 Set itmX = lstInfo.ListItems.Add(, , "Device ID")
                 itmX.SubItems(1) = CStr(.DeviceID)
                  Set itmX = lstInfo.ListItems.Add(, , "Drive")
                  itmX.SubItems(1) = CStr(.Drive)
                   Set itmX = lstInfo.ListItems.Add(, , "Drive Integrity")
                   itmX.SubItems(1) = CBoolStr(.DriveIntegrity)
                    Set itmX = lstInfo.ListItems.Add(, , "Error Cleared")
                    itmX.SubItems(1) = CBoolStr(.ErrorCleared)
                     Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                     itmX.SubItems(1) = CStr(.ErrorDescription)
                      Set itmX = lstInfo.ListItems.Add(, , "ErrorMethodology")
                      itmX.SubItems(1) = CStr(.ErrorMethodology)
                       Set itmX = lstInfo.ListItems.Add(, , "File System Flags")
                       itmX.SubItems(1) = CStr(.FileSystemFlags)
                        Set itmX = lstInfo.ListItems.Add(, , "File System Flags Extended")
                        itmX.SubItems(1) = CStr(.FileSystemFlagsEx)
                         Set itmX = lstInfo.ListItems.Add(, , "ID")
                         itmX.SubItems(1) = CStr(.ID)
                          Set itmX = lstInfo.ListItems.Add(, , "Last Error Code")
                          itmX.SubItems(1) = CStr(.LastErrorCode)
                           Set itmX = lstInfo.ListItems.Add(, , "Manufacturer")
                           itmX.SubItems(1) = CStr(.Manufacturer)
                            Set itmX = lstInfo.ListItems.Add(, , "Number Of Blocks")
                            itmX.SubItems(1) = GroupDigits(.NumberOfBlocks)
                             Set itmX = lstInfo.ListItems.Add(, , "Max Block Size")
                             itmX.SubItems(1) = FormatByteSize(.MaxBlockSize)
                              Set itmX = lstInfo.ListItems.Add(, , "Maximum Component Length")
                              itmX.SubItems(1) = CStr(.MaximumComponentLength)
                               Set itmX = lstInfo.ListItems.Add(, , "Max Media Size")
                               itmX.SubItems(1) = FormatByteSize(.MaxMediaSize)
                                Set itmX = lstInfo.ListItems.Add(, , "Media Loaded")
                                itmX.SubItems(1) = CBoolStr(.MediaLoaded)
                                 Set itmX = lstInfo.ListItems.Add(, , "MediaType")
                                 itmX.SubItems(1) = CStr(.MediaType)
                                  Set itmX = lstInfo.ListItems.Add(, , "Mfr Assigned Revision Level")
                                  itmX.SubItems(1) = CStr(.MfrAssignedRevisionLevel)
                                   Set itmX = lstInfo.ListItems.Add(, , "Min Block Size")
                                   itmX.SubItems(1) = FormatByteSize(.MinBlockSize)
                                    Set itmX = lstInfo.ListItems.Add(, , "Name")
                                    itmX.SubItems(1) = .Name
                                     Set itmX = lstInfo.ListItems.Add(, , "Needs Cleaning?")
                                     itmX.SubItems(1) = CBoolStr(.NeedsCleaning)
                                      Set itmX = lstInfo.ListItems.Add(, , "Number Of Media Supported")
                                      itmX.SubItems(1) = GroupDigits(.NumberOfMediaSupported)
                                       Set itmX = lstInfo.ListItems.Add(, , "Plug-N-Play Device ID")
                                       itmX.SubItems(1) = CStr(.PNPDeviceID)
                                        Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported")
                                        itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                                         Set itmX = lstInfo.ListItems.Add(, , "Revision Level")
                                         itmX.SubItems(1) = CStr(.RevisionLevel)
                                          Set itmX = lstInfo.ListItems.Add(, , "SCSI Bus")
                                          itmX.SubItems(1) = CStr(.SCSIBus)
                                           Set itmX = lstInfo.ListItems.Add(, , "SCSI Logical Unit")
                                           itmX.SubItems(1) = CStr(.SCSILogicalUnit)
                                            Set itmX = lstInfo.ListItems.Add(, , "SCSI Port")
                                            itmX.SubItems(1) = CStr(.SCSIPort)
                                             Set itmX = lstInfo.ListItems.Add(, , "SCSI Target Id")
                                             itmX.SubItems(1) = CStr(.SCSITargetId)
                                              Set itmX = lstInfo.ListItems.Add(, , "Size")
                                              itmX.SubItems(1) = FormatByteSize(.Size)
                                               Set itmX = lstInfo.ListItems.Add(, , "Status")
                                               itmX.SubItems(1) = CStr(.Status)
                                                Set itmX = lstInfo.ListItems.Add(, , "Status Info")
                                                itmX.SubItems(1) = CStr(.StatusInfo)
                                                 Set itmX = lstInfo.ListItems.Add(, , "System Name")
                                                 itmX.SubItems(1) = CStr(.SystemName)
                                                  Set itmX = lstInfo.ListItems.Add(, , "Transfer Rate")
                                                  itmX.SubItems(1) = CStr(.TransferRate)
                                                   Set itmX = lstInfo.ListItems.Add(, , "Volume Name")
                                                   itmX.SubItems(1) = CStr(.VolumeName)
                                                    Set itmX = lstInfo.ListItems.Add(, , "Volume Serial Number")
                                                    itmX.SubItems(1) = CStr(.VolumeSerialNumber)
           End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
End Sub

Sub GetDiskDrive(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
       Set objItem = objWMI.execquery("Select * from Win32_DiskDrive", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Disk Drives": lstMultItems.ListItems.Add , "d" & CStr(.DeviceID), CStr(.DeviceID)
          If Trim(MultiItem$) = Trim(CStr(Item.DeviceID)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Availability")
           itmX.SubItems(1) = CStr(.Availability)
            Set itmX = lstInfo.ListItems.Add(, , "Bytes Per Sector")
            itmX.SubItems(1) = GroupDigits(.BytesPerSector)
             Set itmX = lstInfo.ListItems.Add(, , "Compression Method")
             itmX.SubItems(1) = CStr(.CompressionMethod)
              Set itmX = lstInfo.ListItems.Add(, , "Config Manager Error Code")
              itmX.SubItems(1) = CStr(.ConfigManagerErrorCode)
               Set itmX = lstInfo.ListItems.Add(, , "Config Manager User Config")
               itmX.SubItems(1) = CBoolStr(.ConfigManagerUserConfig)
                Set itmX = lstInfo.ListItems.Add(, , "Default Block Size")
                itmX.SubItems(1) = FormatByteSize(.DefaultBlockSize)
                 Set itmX = lstInfo.ListItems.Add(, , "Description")
                 itmX.SubItems(1) = CStr(.Description)
                  Set itmX = lstInfo.ListItems.Add(, , "Device ID")
                  itmX.SubItems(1) = CStr(.DeviceID)
                   Set itmX = lstInfo.ListItems.Add(, , "Error Cleared")
                   itmX.SubItems(1) = CBoolStr(.ErrorCleared)
                    Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                    itmX.SubItems(1) = CStr(.ErrorDescription)
                     Set itmX = lstInfo.ListItems.Add(, , "Error Methodology")
                     itmX.SubItems(1) = CStr(.ErrorMethodology)
                      Set itmX = lstInfo.ListItems.Add(, , "Index")
                      itmX.SubItems(1) = CStr(.Index)
                       Set itmX = lstInfo.ListItems.Add(, , "Interface Type")
                       itmX.SubItems(1) = CStr(.InterfaceType)
                        Set itmX = lstInfo.ListItems.Add(, , "Last Error Code")
                        itmX.SubItems(1) = CStr(.LastErrorCode)
                         Set itmX = lstInfo.ListItems.Add(, , "Manufacturer")
                         itmX.SubItems(1) = CStr(.Manufacturer)
                          Set itmX = lstInfo.ListItems.Add(, , "Max Block Size")
                          itmX.SubItems(1) = FormatByteSize(.MaxBlockSize)
                           Set itmX = lstInfo.ListItems.Add(, , "Max Media Size")
                           itmX.SubItems(1) = FormatByteSize(.MaxMediaSize)
                            Set itmX = lstInfo.ListItems.Add(, , "Media Loaded")
                            itmX.SubItems(1) = CBoolStr(.MediaLoaded)
                             Set itmX = lstInfo.ListItems.Add(, , "Media Type")
                             itmX.SubItems(1) = CStr(.MediaType)
                              Set itmX = lstInfo.ListItems.Add(, , "Min Block Size")
                              itmX.SubItems(1) = FormatByteSize(.MinBlockSize)
                               Set itmX = lstInfo.ListItems.Add(, , "Model")
                               itmX.SubItems(1) = CStr(.Model)
                                Set itmX = lstInfo.ListItems.Add(, , "Name")
                                itmX.SubItems(1) = CStr(.Name)
                                 Set itmX = lstInfo.ListItems.Add(, , "Needs Cleaning?")
                                 itmX.SubItems(1) = CBoolStr(.NeedsCleaning)
                                  Set itmX = lstInfo.ListItems.Add(, , "Number Of Media Supported")
                                  itmX.SubItems(1) = GroupDigits(.NumberOfMediaSupported)
                                   Set itmX = lstInfo.ListItems.Add(, , "Partitions")
                                   itmX.SubItems(1) = GroupDigits(.Partitions)
                                    Set itmX = lstInfo.ListItems.Add(, , "Plug-N-Play Device ID")
                                    itmX.SubItems(1) = CStr(.PNPDeviceID)
                                     Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported?")
                                     itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                                      Set itmX = lstInfo.ListItems.Add(, , "SCSI Bus")
                                      itmX.SubItems(1) = CStr(.SCSIBus)
                                       Set itmX = lstInfo.ListItems.Add(, , "SCSI Logical Unit")
                                       itmX.SubItems(1) = CStr(.SCSILogicalUnit)
                                        Set itmX = lstInfo.ListItems.Add(, , "SCSI Port")
                                        itmX.SubItems(1) = CStr(.SCSIPort)
                                         Set itmX = lstInfo.ListItems.Add(, , "SCSI Target Id")
                                         itmX.SubItems(1) = CStr(.SCSITargetId)
                                          Set itmX = lstInfo.ListItems.Add(, , "SCSI Bus")
                                          itmX.SubItems(1) = CStr(.SCSIBus)
                                           Set itmX = lstInfo.ListItems.Add(, , "SCSI Logical Unit")
                                           itmX.SubItems(1) = CStr(.SCSILogicalUnit)
                                            Set itmX = lstInfo.ListItems.Add(, , "SCSI Port")
                                            itmX.SubItems(1) = CStr(.SCSIPort)
                                             Set itmX = lstInfo.ListItems.Add(, , "SCSI Target Id")
                                             itmX.SubItems(1) = CStr(.SCSITargetId)
                                              Set itmX = lstInfo.ListItems.Add(, , "Sectors Per Track")
                                              itmX.SubItems(1) = GroupDigits(.SectorsPerTrack)
                                               Set itmX = lstInfo.ListItems.Add(, , "Signature")
                                               itmX.SubItems(1) = CStr(.Signature)
                                                Set itmX = lstInfo.ListItems.Add(, , "Size")
                                                itmX.SubItems(1) = FormatByteSize(.Size)
                                                 Set itmX = lstInfo.ListItems.Add(, , "Status")
                                                 itmX.SubItems(1) = CStr(.Status)
                                                  Set itmX = lstInfo.ListItems.Add(, , "Status Info")
                                                  itmX.SubItems(1) = CStr(.StatusInfo)
                                                   Set itmX = lstInfo.ListItems.Add(, , "System Name")
                                                   itmX.SubItems(1) = CStr(.SystemName)
                                                    Set itmX = lstInfo.ListItems.Add(, , "Total Cylinders")
                                                    itmX.SubItems(1) = GroupDigits(.TotalCylinders)
                                                     Set itmX = lstInfo.ListItems.Add(, , "Total Heads")
                                                     itmX.SubItems(1) = GroupDigits(.TotalHeads)
                                                      Set itmX = lstInfo.ListItems.Add(, , "Total Sectors")
                                                      itmX.SubItems(1) = GroupDigits(.TotalSectors)
                                                       Set itmX = lstInfo.ListItems.Add(, , "Tracks Per Cylinder")
                                                       itmX.SubItems(1) = GroupDigits(.TracksPerCylinder)
           End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
End Sub

Sub GetFloppyController(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
      Set objItem = objWMI.execquery("Select * from Win32_FloppyController", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Floppy Controllers": lstMultItems.ListItems.Add , "d" & CStr(.DeviceID), CStr(.Description)
          If Trim(MultiItem$) = Trim(CStr(Item.DeviceID)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Availability")
           itmX.SubItems(1) = CStr(.Availability)
            Set itmX = lstInfo.ListItems.Add(, , "Config Manager Error Code")
            itmX.SubItems(1) = CStr(.ConfigManagerErrorCode)
             Set itmX = lstInfo.ListItems.Add(, , "Config Manager User Config")
             itmX.SubItems(1) = CBoolStr(.ConfigManagerUserConfig)
              Set itmX = lstInfo.ListItems.Add(, , "Description")
              itmX.SubItems(1) = CStr(.Description)
               Set itmX = lstInfo.ListItems.Add(, , "Device ID")
               itmX.SubItems(1) = CStr(.DeviceID)
                Set itmX = lstInfo.ListItems.Add(, , "Error Cleared")
                itmX.SubItems(1) = CBoolStr(.ErrorCleared)
                 Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                 itmX.SubItems(1) = CStr(.ErrorDescription)
                  Set itmX = lstInfo.ListItems.Add(, , "Last Error Code")
                  itmX.SubItems(1) = CStr(.LastErrorCode)
                   Set itmX = lstInfo.ListItems.Add(, , "Manufacturer")
                   itmX.SubItems(1) = CStr(.Manufacturer)
                    Set itmX = lstInfo.ListItems.Add(, , "Max Number Controlled")
                    itmX.SubItems(1) = GroupDigits(.MaxNumberControlled)
                     Set itmX = lstInfo.ListItems.Add(, , "Name")
                     itmX.SubItems(1) = CStr(.Name)
                      Set itmX = lstInfo.ListItems.Add(, , "Plug-N-Play Device ID")
                      itmX.SubItems(1) = CStr(.PNPDeviceID)
                       Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported")
                       itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                        Set itmX = lstInfo.ListItems.Add(, , "Protocol Supported")
                        itmX.SubItems(1) = CStr(.ProtocolSupported)
                         Set itmX = lstInfo.ListItems.Add(, , "Status")
                         itmX.SubItems(1) = CStr(.Status)
                          Set itmX = lstInfo.ListItems.Add(, , "Status Info")
                          itmX.SubItems(1) = CStr(.StatusInfo)
                           Set itmX = lstInfo.ListItems.Add(, , "System Name")
                           itmX.SubItems(1) = CStr(.SystemName)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
End Sub

Sub GetIDEController(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
       Set objItem = objWMI.execquery("Select * from Win32_IDEController", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "IDE Controllers": lstMultItems.ListItems.Add , "d" & CStr(.Description), CStr(.Description)
          If Trim(MultiItem$) = Trim(CStr(Item.Description)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Availability")
           itmX.SubItems(1) = CStr(.Availability)
            Set itmX = lstInfo.ListItems.Add(, , "Config Manager Error Code")
            itmX.SubItems(1) = CStr(.ConfigManagerErrorCode)
             Set itmX = lstInfo.ListItems.Add(, , "Config Manager User Config")
             itmX.SubItems(1) = CBoolStr(.ConfigManagerUserConfig)
              Set itmX = lstInfo.ListItems.Add(, , "Description")
              itmX.SubItems(1) = CStr(.Description)
               Set itmX = lstInfo.ListItems.Add(, , "Device ID")
               itmX.SubItems(1) = CStr(.DeviceID)
                Set itmX = lstInfo.ListItems.Add(, , "Error Cleared")
                itmX.SubItems(1) = CBoolStr(.ErrorCleared)
                 Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                 itmX.SubItems(1) = CStr(.ErrorDescription)
                  Set itmX = lstInfo.ListItems.Add(, , "Last Error Code")
                  itmX.SubItems(1) = CStr(.LastErrorCode)
                   Set itmX = lstInfo.ListItems.Add(, , "Manufacturer")
                   itmX.SubItems(1) = CStr(.Manufacturer)
                    Set itmX = lstInfo.ListItems.Add(, , "Max Number Controlled")
                    itmX.SubItems(1) = GroupDigits(.MaxNumberControlled)
                     Set itmX = lstInfo.ListItems.Add(, , "Name")
                     itmX.SubItems(1) = CStr(.Name)
                      Set itmX = lstInfo.ListItems.Add(, , "Plug-N-Play Device ID")
                      itmX.SubItems(1) = CStr(.PNPDeviceID)
                       Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported")
                       itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                        Set itmX = lstInfo.ListItems.Add(, , "Protocol Supported")
                        itmX.SubItems(1) = CStr(.ProtocolSupported)
                         Set itmX = lstInfo.ListItems.Add(, , "Status")
                         itmX.SubItems(1) = CStr(.Status)
                          Set itmX = lstInfo.ListItems.Add(, , "Status Info")
                          itmX.SubItems(1) = CStr(.StatusInfo)
                           Set itmX = lstInfo.ListItems.Add(, , "System Name")
                           itmX.SubItems(1) = CStr(.SystemName)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
End Sub

Sub GetKeyBoard(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
     Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
       Set objItem = objWMI.execquery("Select * from Win32_Keyboard", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Keyboards": lstMultItems.ListItems.Add , "d" & CStr(.DeviceID), CStr(.Description)
          If Trim(MultiItem$) = Trim(CStr(Item.DeviceID)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Availability")
           itmX.SubItems(1) = CStr(.Availability)
            Set itmX = lstInfo.ListItems.Add(, , "Config Manager Error Code")
            itmX.SubItems(1) = CStr(.ConfigManagerErrorCode)
             Set itmX = lstInfo.ListItems.Add(, , "Config Manager User Config")
             itmX.SubItems(1) = CBoolStr(.ConfigManagerUserConfig)
              Set itmX = lstInfo.ListItems.Add(, , "Description")
              itmX.SubItems(1) = CStr(.Description)
               Set itmX = lstInfo.ListItems.Add(, , "Device ID")
               itmX.SubItems(1) = CStr(.DeviceID)
                Set itmX = lstInfo.ListItems.Add(, , "Error Cleared")
                itmX.SubItems(1) = CBoolStr(.ErrorCleared)
                 Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                 itmX.SubItems(1) = CStr(.ErrorDescription)
                  Set itmX = lstInfo.ListItems.Add(, , "Last Error Code")
                  itmX.SubItems(1) = CStr(.LastErrorCode)
                   Set itmX = lstInfo.ListItems.Add(, , "Is Locked?")
                   itmX.SubItems(1) = CBoolStr(.IsLocked)
                    Set itmX = lstInfo.ListItems.Add(, , "Layout")
                    itmX.SubItems(1) = CStr(.Layout)
                     Set itmX = lstInfo.ListItems.Add(, , "Name")
                     itmX.SubItems(1) = CStr(.Name)
                      Set itmX = lstInfo.ListItems.Add(, , "Plug-N-Play Device ID")
                      itmX.SubItems(1) = CStr(.PNPDeviceID)
                       Set itmX = lstInfo.ListItems.Add(, , "Number Of Function Keys")
                       itmX.SubItems(1) = GroupDigits(.NumberOfFunctionKeys)
                        Set itmX = lstInfo.ListItems.Add(, , "Password")
                        itmX.SubItems(1) = CStr(.Password)
                         Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported?")
                         itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                          Set itmX = lstInfo.ListItems.Add(, , "Status")
                          itmX.SubItems(1) = CStr(.Status)
                           Set itmX = lstInfo.ListItems.Add(, , "Status Info")
                           itmX.SubItems(1) = CStr(.StatusInfo)
                            Set itmX = lstInfo.ListItems.Add(, , "System Name")
                            itmX.SubItems(1) = CStr(.SystemName)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
End Sub

Sub GetMotherBoard(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
     Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
       Set objItem = objWMI.execquery("Select * from Win32_MotherBoardDevice", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Mother Board Device(s)": lstMultItems.ListItems.Add , "d" & CStr(.DeviceID), CStr(.Description)
          If Trim(MultiItem$) = Trim(CStr(Item.DeviceID)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Availability")
           itmX.SubItems(1) = CStr(.Availability)
            Set itmX = lstInfo.ListItems.Add(, , "Config Manager Error Code")
            itmX.SubItems(1) = CStr(.ConfigManagerErrorCode)
             Set itmX = lstInfo.ListItems.Add(, , "Config Manager User Config")
             itmX.SubItems(1) = CBoolStr(.ConfigManagerUserConfig)
              Set itmX = lstInfo.ListItems.Add(, , "Description")
              itmX.SubItems(1) = CStr(.Description)
               Set itmX = lstInfo.ListItems.Add(, , "Device ID")
               itmX.SubItems(1) = CStr(.DeviceID)
                Set itmX = lstInfo.ListItems.Add(, , "Error Cleared")
                itmX.SubItems(1) = CBoolStr(.ErrorCleared)
                 Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                 itmX.SubItems(1) = CStr(.ErrorDescription)
                  Set itmX = lstInfo.ListItems.Add(, , "Last Error Code")
                  itmX.SubItems(1) = CStr(.LastErrorCode)
                   Set itmX = lstInfo.ListItems.Add(, , "Layout")
                   itmX.SubItems(1) = CStr(.Layout)
                    Set itmX = lstInfo.ListItems.Add(, , "Name")
                    itmX.SubItems(1) = CStr(.Name)
                     Set itmX = lstInfo.ListItems.Add(, , "Plug-N-Play Device ID")
                     itmX.SubItems(1) = CStr(.PNPDeviceID)
                      Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported?")
                      itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                       Set itmX = lstInfo.ListItems.Add(, , "Primary Bus Type")
                       itmX.SubItems(1) = CStr(.PrimaryBusType)
                        Set itmX = lstInfo.ListItems.Add(, , "Revision Number")
                        itmX.SubItems(1) = CStr(.RevisionNumber)
                         Set itmX = lstInfo.ListItems.Add(, , "Secondary Bus Type")
                         itmX.SubItems(1) = CStr(.SecondaryBusType)
                          Set itmX = lstInfo.ListItems.Add(, , "Status")
                          itmX.SubItems(1) = CStr(.Status)
                           Set itmX = lstInfo.ListItems.Add(, , "Status Info")
                           itmX.SubItems(1) = CStr(.StatusInfo)
                            Set itmX = lstInfo.ListItems.Add(, , "System Name")
                            itmX.SubItems(1) = CStr(.SystemName)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
End Sub

Sub GetNetworkAdapter(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object, i&
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
     Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
       Set objItem = objWMI.execquery("Select * from Win32_NetworkAdapter", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Network Adapters": lstMultItems.ListItems.Add , "d" & CStr(.DeviceID), CStr(.Description)
          If Trim(MultiItem$) = Trim(CStr(Item.DeviceID)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Adapter Type")
           itmX.SubItems(1) = CStr(.AdapterType)
            Set itmX = lstInfo.ListItems.Add(, , "Adapter Type ID")
            itmX.SubItems(1) = CStr(.AdapterTypeId)
             Set itmX = lstInfo.ListItems.Add(, , "Auto Sense?")
             itmX.SubItems(1) = CBoolStr(.AutoSense)
              Set itmX = lstInfo.ListItems.Add(, , "Availability")
              itmX.SubItems(1) = CStr(.Availability)
               Set itmX = lstInfo.ListItems.Add(, , "Config Manager Error Code")
               itmX.SubItems(1) = CStr(.ConfigManagerErrorCode)
                Set itmX = lstInfo.ListItems.Add(, , "ConfigManagerUserConfig")
                itmX.SubItems(1) = CBoolStr(.ConfigManagerUserConfig)
                 Set itmX = lstInfo.ListItems.Add(, , "Description")
                 itmX.SubItems(1) = CStr(.Description)
                  Set itmX = lstInfo.ListItems.Add(, , "Device ID")
                  itmX.SubItems(1) = CStr(.DeviceID)
                   Set itmX = lstInfo.ListItems.Add(, , "Error Cleared")
                   itmX.SubItems(1) = CBoolStr(.ErrorCleared)
                    Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                    itmX.SubItems(1) = CStr(.ErrorDescription)
                     Set itmX = lstInfo.ListItems.Add(, , "Index")
                     itmX.SubItems(1) = CStr(.Index)
                      Set itmX = lstInfo.ListItems.Add(, , "Installed?")
                      itmX.SubItems(1) = CBoolStr(.Installed)
                       Set itmX = lstInfo.ListItems.Add(, , "Last Error Code")
                       itmX.SubItems(1) = CStr(.LastErrorCode)
                        Set itmX = lstInfo.ListItems.Add(, , "MAC Address")
                        itmX.SubItems(1) = CStr(.MACAddress)
                         Set itmX = lstInfo.ListItems.Add(, , "Manufacturer")
                         itmX.SubItems(1) = CStr(.Manufacturer)
                          Set itmX = lstInfo.ListItems.Add(, , "Max Number Controlled")
                          itmX.SubItems(1) = GroupDigits(.MaxNumberControlled)
                           Set itmX = lstInfo.ListItems.Add(, , "Max Speed")
                           itmX.SubItems(1) = CStr(.MaxSpeed)
                            Set itmX = lstInfo.ListItems.Add(, , "Name")
                            itmX.SubItems(1) = CStr(.Name)
                             Set itmX = lstInfo.ListItems.Add(, , "Net Connection ID")
                             itmX.SubItems(1) = CStr(.NetConnectionID)
                              Set itmX = lstInfo.ListItems.Add(, , "Net Connection Status")
                              itmX.SubItems(1) = CStr(.NetConnectionStatus)
                               Set itmX = lstInfo.ListItems.Add(, , "Network Addresses")
                                Dim tmpBuff$ 'Dimensionalize tmpBuff as string data type
                                 For i = LBound(.NetworkAddresses) To UBound(.NetworkAddresses)
                                 'For Next loop; enumerate through each element in the array
                                 'LBound function returns the lowest element index in an array
                                 'UBound function returns the highest element index in an array
                                  tmpBuff = tmpBuff & ", " & .NetworkAddresses(i)
                                  'appent the element: i to the temp buffer
                                 Next i
                                  itmX.SubItems(1) = Mid$(tmpBuff, 2) 'update the subitem(column) and remove the last comma and space
                                Set itmX = lstInfo.ListItems.Add(, , "Permanent Address")
                                itmX.SubItems(1) = CStr(.PermanentAddress)
                                 Set itmX = lstInfo.ListItems.Add(, , "Plug-N-Play Device ID")
                                 itmX.SubItems(1) = CStr(.PNPDeviceID)
                                  Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported")
                                  itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                                   Set itmX = lstInfo.ListItems.Add(, , "Partitions")
                                   itmX.SubItems(1) = CStr(.Partitions)
                                    Set itmX = lstInfo.ListItems.Add(, , "Product Name")
                                    itmX.SubItems(1) = CStr(.ProductName)
                                     Set itmX = lstInfo.ListItems.Add(, , "Service Name")
                                     itmX.SubItems(1) = CStr(.ServiceName)
                                      Set itmX = lstInfo.ListItems.Add(, , "Speed")
                                      itmX.SubItems(1) = CStr(.Speed)
                                       Set itmX = lstInfo.ListItems.Add(, , "Status")
                                       itmX.SubItems(1) = CStr(.Status)
                                        Set itmX = lstInfo.ListItems.Add(, , "Status Info")
                                        itmX.SubItems(1) = CStr(.StatusInfo)
                                         Set itmX = lstInfo.ListItems.Add(, , "System Name")
                                         itmX.SubItems(1) = CStr(.SystemName)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
End Sub

Sub GetPnPEntity(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
     Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
       Set objItem = objWMI.execquery("Select * from Win32_PnPEntity", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Plug-N-Play Entities": lstMultItems.ListItems.Add , "d" & CStr(.DeviceID), CStr(.Description)
          If Trim(MultiItem$) = Trim(CStr(Item.DeviceID)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Availability")
           itmX.SubItems(1) = CStr(.Availability)
            Set itmX = lstInfo.ListItems.Add(, , "Class Guid")
            itmX.SubItems(1) = CStr(.ClassGuid)
             Set itmX = lstInfo.ListItems.Add(, , "Config Manager Error Code")
             itmX.SubItems(1) = CStr(.ConfigManagerErrorCode)
              Set itmX = lstInfo.ListItems.Add(, , "Config Manager User Config")
              itmX.SubItems(1) = CBoolStr(.ConfigManagerUserConfig)
               Set itmX = lstInfo.ListItems.Add(, , "Device ID")
               itmX.SubItems(1) = CStr(.DeviceID)
                Set itmX = lstInfo.ListItems.Add(, , "Description")
                itmX.SubItems(1) = CStr(.Description)
                 Set itmX = lstInfo.ListItems.Add(, , "Error Cleared")
                 itmX.SubItems(1) = CBoolStr(.ErrorCleared)
                  Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                  itmX.SubItems(1) = CStr(.ErrorDescription)
                   Set itmX = lstInfo.ListItems.Add(, , "Last Error Code")
                   itmX.SubItems(1) = CStr(.LastErrorCode)
                    Set itmX = lstInfo.ListItems.Add(, , "Manufacturer")
                    itmX.SubItems(1) = CStr(.Manufacturer)
                     Set itmX = lstInfo.ListItems.Add(, , "Plug-N-Play Device ID")
                     itmX.SubItems(1) = CStr(.PNPDeviceID)
                      Set itmX = lstInfo.ListItems.Add(, , "Name")
                      itmX.SubItems(1) = CStr(.Name)
                       Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported?")
                       itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                        Set itmX = lstInfo.ListItems.Add(, , "Service")
                        itmX.SubItems(1) = CStr(.Service)
                         Set itmX = lstInfo.ListItems.Add(, , "Status")
                         itmX.SubItems(1) = CStr(.Status)
                          Set itmX = lstInfo.ListItems.Add(, , "Status Info")
                          itmX.SubItems(1) = CStr(.StatusInfo)
                           Set itmX = lstInfo.ListItems.Add(, , "System Name")
                           itmX.SubItems(1) = CStr(.SystemName)
                            If MultiItem$ <> "" Then GoTo KillObj
          End If
         End With
        Next Item
KillObj:
         Set objWMI = Nothing
          Set objItem = Nothing
End Sub

Sub GetPotsModem(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object, i&
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
       Set objItem = objWMI.execquery("Select * from Win32_PotsModem", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "POTS Modems": lstMultItems.ListItems.Add , "d" & CStr(.DeviceID), CStr(.Description)
          If Trim(MultiItem$) = Trim(CStr(Item.DeviceID)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
            Set itmX = lstInfo.ListItems.Add(, , "Answer Mode")
            itmX.SubItems(1) = CStr(.AnswerMode)
             Set itmX = lstInfo.ListItems.Add(, , "Attached To")
             itmX.SubItems(1) = CStr(.AttachedTo)
              Set itmX = lstInfo.ListItems.Add(, , "Availability")
              itmX.SubItems(1) = CStr(.Availability)
               Set itmX = lstInfo.ListItems.Add(, , "Blind Off")
               itmX.SubItems(1) = CStr(.BlindOff)
                Set itmX = lstInfo.ListItems.Add(, , "Blind On")
                itmX.SubItems(1) = CStr(.BlindOn)
                 Set itmX = lstInfo.ListItems.Add(, , "Compatibility Flags")
                 itmX.SubItems(1) = CStr(.CompatibilityFlags)
                  Set itmX = lstInfo.ListItems.Add(, , "Compression Info")
                  itmX.SubItems(1) = CStr(.CompressionInfo)
                   Set itmX = lstInfo.ListItems.Add(, , "Compression Off")
                   itmX.SubItems(1) = CStr(.CompressionOff)
                    Set itmX = lstInfo.ListItems.Add(, , "Compression On")
                    itmX.SubItems(1) = CStr(.CompressionOn)
                     Set itmX = lstInfo.ListItems.Add(, , "Config Manager Error Code")
                     itmX.SubItems(1) = CStr(.ConfigManagerErrorCode)
                      Set itmX = lstInfo.ListItems.Add(, , "Config Manager User Config")
                      itmX.SubItems(1) = CBoolStr(.ConfigManagerUserConfig)
                       Set itmX = lstInfo.ListItems.Add(, , "Configuration Dialog")
                       itmX.SubItems(1) = CStr(.ConfigurationDialog)
                        Set itmX = lstInfo.ListItems.Add(, , "Countries Supported")
                         tmpBuffer = ""
                         For i = LBound(.CountriesSupported) To UBound(.CountriesSupported)
                          tmpBuffer = tmpBuffer & ", " & .CountriesSupported(i)
                         Next i
                          itmX.SubItems(1) = Mid$(tmpBuffer, 3)
                         Set itmX = lstInfo.ListItems.Add(, , "Country Selected")
                         itmX.SubItems(1) = CStr(.CountrySelected)
                          Set itmX = lstInfo.ListItems.Add(, , "Current Passwords")
                           tmpBuffer = ""
                           For i = LBound(.CurrentPasswords) To UBound(.CurrentPasswords)
                            tmpBuffer = tmpBuffer & ", " & .CurrentPasswords(i)
                           Next i
                            itmX.SubItems(1) = Mid$(tmpBuffer, 3)
                           Set itmX = lstInfo.ListItems.Add(, , "DCB")
                           tmpBuffer = ""
                            For i = LBound(.DCB) To UBound(.DCB)
                             tmpBuffer = tmpBuffer & ", " & CStr(.DCB(i))
                            Next i
                             itmX.SubItems(1) = Mid$(tmpBuffer, 3)
                            Set itmX = lstInfo.ListItems.Add(, , "Default")
                            tmpBuffer = ""
                             For i = LBound(.Default) To UBound(.Default)
                              tmpBuffer = tmpBuffer & ", " & CStr(.Default(i))
                             Next i
                              itmX.SubItems(1) = Mid$(tmpBuffer, 3)
                             Set itmX = lstInfo.ListItems.Add(, , "Description")
                             itmX.SubItems(1) = .Description
                              Set itmX = lstInfo.ListItems.Add(, , "Device ID")
                              itmX.SubItems(1) = CStr(.DeviceID)
                               Set itmX = lstInfo.ListItems.Add(, , "Device Loader")
                               itmX.SubItems(1) = CStr(.DeviceLoader)
                                Set itmX = lstInfo.ListItems.Add(, , "Device Type")
                                itmX.SubItems(1) = CStr(.DeviceType)
                                 Set itmX = lstInfo.ListItems.Add(, , "Dial Type")
                                 itmX.SubItems(1) = .DialType
                                  Set itmX = lstInfo.ListItems.Add(, , "Error Cleared?")
                                  itmX.SubItems(1) = CBoolStr(.ErrorCleared)
                                   Set itmX = lstInfo.ListItems.Add(, , "Error Control Forced")
                                   itmX.SubItems(1) = CStr(.ErrorControlForced)
                                    Set itmX = lstInfo.ListItems.Add(, , "Error Control Info")
                                    itmX.SubItems(1) = CStr(.ErrorControlInfo)
                                     Set itmX = lstInfo.ListItems.Add(, , "Error Control Off")
                                     itmX.SubItems(1) = CStr(.ErrorControlOff)
                                      Set itmX = lstInfo.ListItems.Add(, , "Error Control On")
                                      itmX.SubItems(1) = CStr(.ErrorControlOn)
                                       Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                                       itmX.SubItems(1) = CStr(.ErrorDescription)
                                        Set itmX = lstInfo.ListItems.Add(, , "Flow Control Hard")
                                        itmX.SubItems(1) = CStr(.FlowControlHard)
                                         Set itmX = lstInfo.ListItems.Add(, , "Flow Control Off")
                                         itmX.SubItems(1) = CStr(.FlowControlOff)
                                          Set itmX = lstInfo.ListItems.Add(, , "Flow Control Soft")
                                          itmX.SubItems(1) = CStr(.FlowControlSoft)
                                           Set itmX = lstInfo.ListItems.Add(, , "Inactivity Scale")
                                           itmX.SubItems(1) = CStr(.InactivityScale)
                                            Set itmX = lstInfo.ListItems.Add(, , "Inactivity Timeout")
                                            itmX.SubItems(1) = CStr(.InactivityTimeout)
                                             Set itmX = lstInfo.ListItems.Add(, , "Index")
                                             itmX.SubItems(1) = CStr(.Index)
                                              Set itmX = lstInfo.ListItems.Add(, , "Last Error Code")
                                              itmX.SubItems(1) = CStr(.LastErrorCode)
                                               Set itmX = lstInfo.ListItems.Add(, , "Max Baud Rate To Phone")
                                               itmX.SubItems(1) = CStr(.MaxBaudRateToPhone)
                                                Set itmX = lstInfo.ListItems.Add(, , "Max Baud Rate To Serial Port")
                                                itmX.SubItems(1) = CStr(.MaxBaudRateToSerialPort)
                                                 Set itmX = lstInfo.ListItems.Add(, , "Max Number Of Passwords")
                                                 itmX.SubItems(1) = GroupDigits(.MaxNumberOfPasswords)
                                                  Set itmX = lstInfo.ListItems.Add(, , "Model")
                                                  itmX.SubItems(1) = CStr(.Model)
                                                   Set itmX = lstInfo.ListItems.Add(, , "Modem INF File Path")
                                                   itmX.SubItems(1) = CStr(.ModemInfPath)
                                                    Set itmX = lstInfo.ListItems.Add(, , "Modem INF File Section")
                                                    itmX.SubItems(1) = CStr(.ModemInfSection)
                                                     Set itmX = lstInfo.ListItems.Add(, , "Modulation Bell")
                                                     itmX.SubItems(1) = CStr(.ModulationBell)
                                                      Set itmX = lstInfo.ListItems.Add(, , "Modulation CCITT")
                                                      itmX.SubItems(1) = CStr(.ModulationCCITT)
                                                       Set itmX = lstInfo.ListItems.Add(, , "Modulation Scheme")
                                                       itmX.SubItems(1) = CStr(.ModulationScheme)
                                                        Set itmX = lstInfo.ListItems.Add(, , "Name")
                                                        itmX.SubItems(1) = CStr(.Name)
                                                         Set itmX = lstInfo.ListItems.Add(, , "Plug-N-Play Device ID")
                                                         itmX.SubItems(1) = CStr(.PNPDeviceID)
                                                          Set itmX = lstInfo.ListItems.Add(, , "Port Sub-Class")
                                                          itmX.SubItems(1) = CStr(.PortSubClass)
                                                           Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported?")
                                                           itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                                                            Set itmX = lstInfo.ListItems.Add(, , "Prefix")
                                                            itmX.SubItems(1) = CStr(.Prefix)
                                                             Set itmX = lstInfo.ListItems.Add(, , "Properties")
                                                              tmpBuffer = ""
                                                              For i = LBound(.Properties) To UBound(.Properties)
                                                               tmpBuffer = tmpBuffer & ", " & CStr(.Properties(i))
                                                              Next i
                                                               itmX.SubItems(1) = Mid$(tmpBuffer, 3)
                                                              Set itmX = lstInfo.ListItems.Add(, , "Provider Name")
                                                              itmX.SubItems(1) = CStr(.ProviderName)
                                                               Set itmX = lstInfo.ListItems.Add(, , "Pulse")
                                                               itmX.SubItems(1) = CStr(.Pulse)
                                                                Set itmX = lstInfo.ListItems.Add(, , "Reset")
                                                                itmX.SubItems(1) = CStr(.Reset)
                                                                 Set itmX = lstInfo.ListItems.Add(, , "Responses Key Name")
                                                                 itmX.SubItems(1) = CStr(.ResponsesKeyName)
                                                                  Set itmX = lstInfo.ListItems.Add(, , "Rings Before Answer")
                                                                  itmX.SubItems(1) = CStr(.RingsBeforeAnswer)
                                                                   Set itmX = lstInfo.ListItems.Add(, , "Speaker Mode Dial")
                                                                   itmX.SubItems(1) = CStr(.SpeakerModeDial)
                                                                    Set itmX = lstInfo.ListItems.Add(, , "Speaker Mode Off")
                                                                    itmX.SubItems(1) = CStr(.SpeakerModeOff)
                                                                     Set itmX = lstInfo.ListItems.Add(, , "Speaker Mode On")
                                                                     itmX.SubItems(1) = CStr(.SpeakerModeOn)
                                                                      Set itmX = lstInfo.ListItems.Add(, , "Speaker Mode Setup")
                                                                      itmX.SubItems(1) = CStr(.SpeakerModeSetup)
                                                                       Set itmX = lstInfo.ListItems.Add(, , "Speaker Volume High")
                                                                       itmX.SubItems(1) = CStr(.SpeakerVolumeHigh)
                                                                        Set itmX = lstInfo.ListItems.Add(, , "Speaker Volume Info")
                                                                        itmX.SubItems(1) = CStr(.SpeakerVolumeInfo)
                                                                         Set itmX = lstInfo.ListItems.Add(, , "Speaker Volume Low")
                                                                         itmX.SubItems(1) = CStr(.SpeakerVolumeLow)
                                                                          Set itmX = lstInfo.ListItems.Add(, , "Speaker Volume Medium")
                                                                          itmX.SubItems(1) = CStr(.SpeakerVolumeMed)
                                                                           Set itmX = lstInfo.ListItems.Add(, , "Status")
                                                                           itmX.SubItems(1) = CStr(.Status)
                                                                            Set itmX = lstInfo.ListItems.Add(, , "Status Info")
                                                                            itmX.SubItems(1) = CStr(.StatusInfo)
                                                                             Set itmX = lstInfo.ListItems.Add(, , "String Format")
                                                                             itmX.SubItems(1) = CStr(.StringFormat)
                                                                              Set itmX = lstInfo.ListItems.Add(, , "Supports Callback?")
                                                                              itmX.SubItems(1) = CBoolStr(.SupportsCallback)
                                                                               Set itmX = lstInfo.ListItems.Add(, , "Supports Synchronous Connect?")
                                                                               itmX.SubItems(1) = CBoolStr(.SupportsSynchronousConnect)
                                                                                Set itmX = lstInfo.ListItems.Add(, , "Terminator")
                                                                                itmX.SubItems(1) = CStr(.Terminator)
                                                                                 Set itmX = lstInfo.ListItems.Add(, , "Tone")
                                                                                 itmX.SubItems(1) = CStr(.Tone)
                                                                                  Set itmX = lstInfo.ListItems.Add(, , "Voice Switch Feature")
                                                                                  itmX.SubItems(1) = CStr(.VoiceSwitchFeature)
                                                                                   Set itmX = lstInfo.ListItems.Add(, , "System Name")
                                                                                   itmX.SubItems(1) = CStr(.SystemName)

          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
End Sub

Sub GetPrinter(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object, i&
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
      If Trim(MultiItem) <> "" Then
       Set objItem = objWMI.execquery("Select * from Win32_Printer where DeviceID = """ & MultiItem & """", , 48)
      Else
       Set objItem = objWMI.execquery("Select * from Win32_Printer", , 48)
      End If
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Printers": lstMultItems.ListItems.Add , "d" & CStr(.DeviceID), CStr(.Caption)
          If tmpCnt& = 1 Then
            Set itmX = lstInfo.ListItems.Add(, , "Attributes")
            itmX.SubItems(1) = CStr(.Attributes)
             Set itmX = lstInfo.ListItems.Add(, , "Availability")
             itmX.SubItems(1) = CStr(.Availability)
              Set itmX = lstInfo.ListItems.Add(, , "Available Job Sheets")
               tmpBuffer = ""
                For i = LBound(.AvailableJobSheets) To UBound(.AvailableJobSheets)
                 tmpBuffer = tmpBuffer & ", " & .AvailableJobSheets(i)
                Next i
                 itmX.SubItems(1) = Mid$(tmpBuffer, 3)
               Set itmX = lstInfo.ListItems.Add(, , "Average Pages Per Minute")
               itmX.SubItems(1) = CStr(.AveragePagesPerMinute)
                Set itmX = lstInfo.ListItems.Add(, , "Capabilities")
                 tmpBuffer = ""
                  For i = LBound(.Capabilities) To UBound(.Capabilities)
                   tmpBuffer = tmpBuffer & ", " & .Capabilities(i)
                  Next i
                   itmX.SubItems(1) = Mid$(tmpBuffer, 3)
                 Set itmX = lstInfo.ListItems.Add(, , "Capability Descriptions")
                 tmpBuffer = ""
                  For i = LBound(.CapabilityDescriptions) To UBound(.CapabilityDescriptions)
                   tmpBuffer = tmpBuffer & ", " & .CapabilityDescriptions(i)
                  Next i
                   itmX.SubItems(1) = Mid$(tmpBuffer, 3)
                  Set itmX = lstInfo.ListItems.Add(, , "Caption")
                  itmX.SubItems(1) = CStr(.Caption)
                   Set itmX = lstInfo.ListItems.Add(, , "Character Sets Supported")
                    tmpBuffer = ""
                    For i = LBound(.CharSetsSupported) To UBound(.CharSetsSupported)
                     tmpBuffer = tmpBuffer & ", " & .CharSetsSupported(i)
                    Next i
                     itmX.SubItems(1) = Mid$(tmpBuffer, 3)
                    Set itmX = lstInfo.ListItems.Add(, , "Comment")
                    itmX.SubItems(1) = CStr(.Comment)
                     Set itmX = lstInfo.ListItems.Add(, , "Config Manager Error Code")
                     itmX.SubItems(1) = CStr(.ConfigManagerErrorCode)
                      Set itmX = lstInfo.ListItems.Add(, , "Config Manager User Config")
                      itmX.SubItems(1) = CBoolStr(.ConfigManagerUserConfig)
                       Set itmX = lstInfo.ListItems.Add(, , "Current Capabilities")
                        tmpBuffer = ""
                        For i = LBound(.CurrentCapabilities) To UBound(.CurrentCapabilities)
                         tmpBuffer = tmpBuffer & ", " & .CurrentCapabilities(i)
                        Next i
                         itmX.SubItems(1) = Mid$(tmpBuffer, 3)
                        Set itmX = lstInfo.ListItems.Add(, , "Current Character Set")
                        itmX.SubItems(1) = CStr(.CurrentCharSet)
                         Set itmX = lstInfo.ListItems.Add(, , "Current Language")
                         itmX.SubItems(1) = CStr(.CurrentLanguage)
                          Set itmX = lstInfo.ListItems.Add(, , "Current Mime Type")
                          itmX.SubItems(1) = CStr(.CurrentMimeType)
                           Set itmX = lstInfo.ListItems.Add(, , "Current Natural Language")
                           itmX.SubItems(1) = CStr(.CurrentNaturalLanguage)
                            Set itmX = lstInfo.ListItems.Add(, , "CurrentPaperType")
                            itmX.SubItems(1) = CStr(.CurrentPaperType)
                             Set itmX = lstInfo.ListItems.Add(, , "Default Printer")
                             itmX.SubItems(1) = CBoolStr(.Default)
                              Set itmX = lstInfo.ListItems.Add(, , "Default Capabilities")
                              tmpBuffer = ""
                               For i = LBound(.DefaultCapabilities) To UBound(.DefaultCapabilities)
                                tmpBuffer = tmpBuffer & ", " & .DefaultCapabilities(i)
                               Next i
                                itmX.SubItems(1) = Mid$(tmpBuffer, 3)
                               Set itmX = lstInfo.ListItems.Add(, , "Default Copies")
                               itmX.SubItems(1) = GroupDigits(.DefaultCopies)
                                Set itmX = lstInfo.ListItems.Add(, , "Default Language")
                                itmX.SubItems(1) = CStr(.DefaultLanguage)
                                 Set itmX = lstInfo.ListItems.Add(, , "Default Mime Type")
                                 itmX.SubItems(1) = CStr(.DefaultMimeType)
                                  Set itmX = lstInfo.ListItems.Add(, , "Default Number Up")
                                  itmX.SubItems(1) = GroupDigits(.DefaultNumberUp)
                                   Set itmX = lstInfo.ListItems.Add(, , "Default Paper Type")
                                   itmX.SubItems(1) = CStr(.DefaultPaperType)
                                    Set itmX = lstInfo.ListItems.Add(, , "Default Priority")
                                    itmX.SubItems(1) = CStr(.DefaultPriority)
                                     Set itmX = lstInfo.ListItems.Add(, , "Description")
                                     itmX.SubItems(1) = CStr(.Description)
                                      Set itmX = lstInfo.ListItems.Add(, , "Detected Error State")
                                      itmX.SubItems(1) = CStr(.DetectedErrorState)
                                       Set itmX = lstInfo.ListItems.Add(, , "Device ID")
                                       itmX.SubItems(1) = CStr(.DeviceID)
                                        Set itmX = lstInfo.ListItems.Add(, , "Direct")
                                        itmX.SubItems(1) = CBoolStr(.Direct)
                                         Set itmX = lstInfo.ListItems.Add(, , "Do Complete First")
                                         itmX.SubItems(1) = CBoolStr(.DoCompleteFirst)
                                          Set itmX = lstInfo.ListItems.Add(, , "Driver Name")
                                          itmX.SubItems(1) = CStr(.DriverName)
                                           Set itmX = lstInfo.ListItems.Add(, , "Enable BIDI")
                                           itmX.SubItems(1) = CBoolStr(.EnableBIDI)
                                            Set itmX = lstInfo.ListItems.Add(, , "Enable Dev Query Print")
                                            itmX.SubItems(1) = CBoolStr(.EnableDevQueryPrint)
                                             Set itmX = lstInfo.ListItems.Add(, , "Error Cleared")
                                             itmX.SubItems(1) = CBoolStr(.ErrorCleared)
                                              Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                                              itmX.SubItems(1) = CStr(.ErrorDescription)
                                               Set itmX = lstInfo.ListItems.Add(, , "ErrorInformation")
                                               tmpBuffer = ""
                                                For i = LBound(.ErrorInformation) To UBound(.ErrorInformation)
                                                 tmpBuffer = tmpBuffer & ", " & .ErrorInformation(i)
                                                Next i
                                                 itmX.SubItems(1) = Mid$(tmpBuffer, 3)
                                                Set itmX = lstInfo.ListItems.Add(, , "Extended Detected Error State")
                                                itmX.SubItems(1) = CStr(.ExtendedDetectedErrorState)
                                                 Set itmX = lstInfo.ListItems.Add(, , "Extended Printer Status")
                                                 itmX.SubItems(1) = CStr(.ExtendedPrinterStatus)
                                                  Set itmX = lstInfo.ListItems.Add(, , "Hidden")
                                                  itmX.SubItems(1) = CBoolStr(.Hidden)
                                                   Set itmX = lstInfo.ListItems.Add(, , "Horizontal Resolution")
                                                   itmX.SubItems(1) = CStr(.HorizontalResolution)
                                                    Set itmX = lstInfo.ListItems.Add(, , "Job Count Since Last Reset")
                                                    itmX.SubItems(1) = CStr(.JobCountSinceLastReset)
                                                     Set itmX = lstInfo.ListItems.Add(, , "Keep Printed Jobs")
                                                     itmX.SubItems(1) = CBoolStr(.KeepPrintedJobs)
                                                      Set itmX = lstInfo.ListItems.Add(, , "Languages Supported")
                                                       tmpBuffer = ""
                                                       For i = LBound(.LanguagesSupported) To UBound(.LanguagesSupported)
                                                        tmpBuffer = tmpBuffer & ", " & .LanguagesSupported(i)
                                                       Next i
                                                        itmX.SubItems(1) = Mid$(tmpBuffer, 3)
                                                       Set itmX = lstInfo.ListItems.Add(, , "Last Error Code")
                                                       itmX.SubItems(1) = CStr(.LastErrorCode)
                                                        Set itmX = lstInfo.ListItems.Add(, , "Local")
                                                        itmX.SubItems(1) = CStr(.Local)
                                                         Set itmX = lstInfo.ListItems.Add(, , "Location")
                                                         itmX.SubItems(1) = CStr(.Location)
                                                          Set itmX = lstInfo.ListItems.Add(, , "Port Sub-Class")
                                                          itmX.SubItems(1) = CStr(.PortSubClass)
                                                           Set itmX = lstInfo.ListItems.Add(, , "Marking Technology")
                                                           itmX.SubItems(1) = CStr(.MarkingTechnology)
                                                            Set itmX = lstInfo.ListItems.Add(, , "Max Copies")
                                                            itmX.SubItems(1) = CStr(.MaxCopies)
                                                             Set itmX = lstInfo.ListItems.Add(, , "Max Number Up")
                                                             itmX.SubItems(1) = .MaxNumberUp
                                                              Set itmX = lstInfo.ListItems.Add(, , "Max Size Supported")
                                                              itmX.SubItems(1) = FormatByteSize(.MaxSizeSupported)
                                                               Set itmX = lstInfo.ListItems.Add(, , "Mime Types Supported")
                                                                tmpBuffer = ""
                                                                For i = LBound(.MimeTypesSupported) To UBound(.MimeTypesSupported)
                                                                 tmpBuffer = tmpBuffer & ", " & .MimeTypesSupported(i)
                                                                Next i
                                                                 itmX.SubItems(1) = Mid$(tmpBuffer, 3)
                                                                Set itmX = lstInfo.ListItems.Add(, , "Name")
                                                                itmX.SubItems(1) = CStr(.Name)
                                                                 Set itmX = lstInfo.ListItems.Add(, , "Natural Languages Supported")
                                                                  tmpBuffer = ""
                                                                  For i = LBound(.NaturalLanguagesSupported) To UBound(.NaturalLanguagesSupported)
                                                                   tmpBuffer = tmpBuffer & ", " & .NaturalLanguagesSupported(i)
                                                                  Next i
                                                                   itmX.SubItems(1) = Mid$(tmpBuffer, 3)
                                                                  Set itmX = lstInfo.ListItems.Add(, , "Network")
                                                                  itmX.SubItems(1) = CBoolStr(.Network)
                                                                   Set itmX = lstInfo.ListItems.Add(, , "Paper Sizes Supported")
                                                                    tmpBuffer = ""
                                                                    For i = LBound(.PaperSizesSupported) To UBound(.PaperSizesSupported)
                                                                     tmpBuffer = tmpBuffer & ", " & .PaperSizesSupported(i)
                                                                    Next i
                                                                     itmX.SubItems(1) = Mid$(tmpBuffer, 3)
                                                                    Set itmX = lstInfo.ListItems.Add(, , "Paper Types Available")
                                                                     tmpBuffer = ""
                                                                     For i = LBound(.PaperTypesAvailable) To UBound(.PaperTypesAvailable)
                                                                      tmpBuffer = tmpBuffer & ", " & .PaperTypesAvailable(i)
                                                                     Next i
                                                                      itmX.SubItems(1) = Mid$(tmpBuffer, 3)
                                                                     Set itmX = lstInfo.ListItems.Add(, , "Parameters")
                                                                     itmX.SubItems(1) = CStr(.Parameters)
                                                                      Set itmX = lstInfo.ListItems.Add(, , "Plug-N-Play Device ID")
                                                                      itmX.SubItems(1) = CStr(.PNPDeviceID)
                                                                       Set itmX = lstInfo.ListItems.Add(, , "Port Name")
                                                                       itmX.SubItems(1) = CStr(.PortName)
                                                                        Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported")
                                                                        itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                                                                         Set itmX = lstInfo.ListItems.Add(, , "Printer Paper Names")
                                                                          tmpBuffer = ""
                                                                          For i = LBound(.PrinterPaperNames) To UBound(.PrinterPaperNames)
                                                                            tmpBuffer = tmpBuffer & ", " & .PrinterPaperNames(i)
                                                                          Next i
                                                                           itmX.SubItems(1) = Mid$(tmpBuffer, 3)
                                                                          Set itmX = lstInfo.ListItems.Add(, , "Printer State")
                                                                          itmX.SubItems(1) = CStr(.PrinterState)
                                                                           Set itmX = lstInfo.ListItems.Add(, , "Printer Status")
                                                                           itmX.SubItems(1) = CStr(.PrinterStatus)
                                                                            Set itmX = lstInfo.ListItems.Add(, , "Print Job Data Type")
                                                                            itmX.SubItems(1) = CStr(.PrintJobDataType)
                                                                             Set itmX = lstInfo.ListItems.Add(, , "Print Processor")
                                                                             itmX.SubItems(1) = CStr(.PrintProcessor)
                                                                              Set itmX = lstInfo.ListItems.Add(, , "Priority")
                                                                              itmX.SubItems(1) = CStr(.Priority)
                                                                               Set itmX = lstInfo.ListItems.Add(, , "Published")
                                                                               itmX.SubItems(1) = CBoolStr(.Published)
                                                                                Set itmX = lstInfo.ListItems.Add(, , "Queued")
                                                                                itmX.SubItems(1) = CBoolStr(.Terminator)
                                                                                 Set itmX = lstInfo.ListItems.Add(, , "Raw Only")
                                                                                 itmX.SubItems(1) = CBoolStr(.RawOnly)
                                                                                  Set itmX = lstInfo.ListItems.Add(, , "Separator File")
                                                                                  itmX.SubItems(1) = CStr(.SeparatorFile)
                                                                                   Set itmX = lstInfo.ListItems.Add(, , "Server Name")
                                                                                   itmX.SubItems(1) = CStr(.ServerName)
                                                                                    Set itmX = lstInfo.ListItems.Add(, , "Shared")
                                                                                    itmX.SubItems(1) = CBoolStr(.Shared)
                                                                                     Set itmX = lstInfo.ListItems.Add(, , "Share Name")
                                                                                     itmX.SubItems(1) = CStr(.ShareName)
                                                                                      Set itmX = lstInfo.ListItems.Add(, , "Spool Enabled")
                                                                                      itmX.SubItems(1) = CBoolStr(.SpoolEnabled)
                                                                                       Set itmX = lstInfo.ListItems.Add(, , "Status")
                                                                                       itmX.SubItems(1) = CStr(.Status)
                                                                                        Set itmX = lstInfo.ListItems.Add(, , "Status Info")
                                                                                        itmX.SubItems(1) = CStr(.StatusInfo)
                                                                                         Set itmX = lstInfo.ListItems.Add(, , "Vertical Resolution")
                                                                                         itmX.SubItems(1) = CStr(.VerticalResolution)
                                                                                          Set itmX = lstInfo.ListItems.Add(, , "Work Offline")
                                                                                          itmX.SubItems(1) = CBoolStr(.WorkOffline)

          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
End Sub

Sub GetSoundDevice(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
     Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
       Set objItem = objWMI.execquery("Select * from Win32_SoundDevice", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Sound Devices": lstMultItems.ListItems.Add , "d" & CStr(.DeviceID), CStr(.Description)
          If Trim(MultiItem$) = Trim(CStr(Item.DeviceID)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Availability")
           itmX.SubItems(1) = CStr(.Availability)
            Set itmX = lstInfo.ListItems.Add(, , "Caption")
            itmX.SubItems(1) = CStr(.Caption)
             Set itmX = lstInfo.ListItems.Add(, , "Config Manager Error Code")
             itmX.SubItems(1) = CStr(.ConfigManagerErrorCode)
              Set itmX = lstInfo.ListItems.Add(, , "Config Manager User Config")
              itmX.SubItems(1) = CBoolStr(.ConfigManagerUserConfig)
               Set itmX = lstInfo.ListItems.Add(, , "Device ID")
               itmX.SubItems(1) = CStr(.DeviceID)
                Set itmX = lstInfo.ListItems.Add(, , "Description")
                itmX.SubItems(1) = CStr(.Description)
                 Set itmX = lstInfo.ListItems.Add(, , "Error Cleared")
                 itmX.SubItems(1) = CBoolStr(.ErrorCleared)
                  Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                  itmX.SubItems(1) = CStr(.ErrorDescription)
                   Set itmX = lstInfo.ListItems.Add(, , "Last Error Code")
                   itmX.SubItems(1) = CStr(.LastErrorCode)
                    Set itmX = lstInfo.ListItems.Add(, , "Manufacturer")
                    itmX.SubItems(1) = CStr(.Manufacturer)
                     Set itmX = lstInfo.ListItems.Add(, , "DMA Buffer Size")
                     itmX.SubItems(1) = FormatByteSize(.DMABufferSize)
                      Set itmX = lstInfo.ListItems.Add(, , "MPU401 Address")
                      itmX.SubItems(1) = CStr(.MPU401Address)
                       Set itmX = lstInfo.ListItems.Add(, , "Name")
                       itmX.SubItems(1) = CStr(.Name)
                        Set itmX = lstInfo.ListItems.Add(, , "Plug-N-Play Device ID")
                        itmX.SubItems(1) = CStr(.PNPDeviceID)
                         Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported")
                         itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                          Set itmX = lstInfo.ListItems.Add(, , "Product Name")
                          itmX.SubItems(1) = CStr(.ProductName)
                           Set itmX = lstInfo.ListItems.Add(, , "Status")
                           itmX.SubItems(1) = CStr(.Status)
                            Set itmX = lstInfo.ListItems.Add(, , "Status Info")
                            itmX.SubItems(1) = CStr(.StatusInfo)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
End Sub

Sub GetLogicalDisk(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
      If Trim(MultiItem) <> "" Then
       Set objItem = objWMI.execquery("Select * from Win32_LogicalDisk where DeviceID = """ & MultiItem & """", , 48)
      Else
       Set objItem = objWMI.execquery("Select * from Win32_LogicalDisk", , 48)
      End If
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Logical Disks": lstMultItems.ListItems.Add , "d" & CStr(.DeviceID), Trim(CStr(.Caption) & " " & CStr(.Description))
          If tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Access")
           itmX.SubItems(1) = CStr(.Access)
            Set itmX = lstInfo.ListItems.Add(, , "Availability")
            itmX.SubItems(1) = CStr(.Availability)
             Set itmX = lstInfo.ListItems.Add(, , "Block Size")
             itmX.SubItems(1) = FormatByteSize(.BlockSize)
              Set itmX = lstInfo.ListItems.Add(, , "Caption")
              itmX.SubItems(1) = CStr(.Caption)
               Set itmX = lstInfo.ListItems.Add(, , "Compressed")
               itmX.SubItems(1) = CBoolStr(.Compressed)
                Set itmX = lstInfo.ListItems.Add(, , "Config Manager Error Code")
                itmX.SubItems(1) = CStr(.ConfigManagerErrorCode)
                 Set itmX = lstInfo.ListItems.Add(, , "Config Manager User Config")
                 itmX.SubItems(1) = CBoolStr(.ConfigManagerUserConfig)
                  Set itmX = lstInfo.ListItems.Add(, , "Description")
                  itmX.SubItems(1) = CStr(.Description)
                   Set itmX = lstInfo.ListItems.Add(, , "Device ID")
                   itmX.SubItems(1) = CStr(.DeviceID)
                    Set itmX = lstInfo.ListItems.Add(, , "Drive Type")
                    itmX.SubItems(1) = CStr(.DriveType)
                     Set itmX = lstInfo.ListItems.Add(, , "Error Cleared")
                     itmX.SubItems(1) = CBoolStr(.ErrorCleared)
                      Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                      itmX.SubItems(1) = CStr(.ErrorDescription)
                       Set itmX = lstInfo.ListItems.Add(, , "Error Methodology")
                       itmX.SubItems(1) = CStr(.ErrorMethodology)
                        Set itmX = lstInfo.ListItems.Add(, , "File System")
                        itmX.SubItems(1) = CStr(.FileSystem)
                         Set itmX = lstInfo.ListItems.Add(, , "Free Space")
                         itmX.SubItems(1) = FormatByteSize(.FreeSpace, "Free")
                          Set itmX = lstInfo.ListItems.Add(, , "Last Error Code")
                          itmX.SubItems(1) = CStr(.LastErrorCode)
                           Set itmX = lstInfo.ListItems.Add(, , "Maximum Component Length")
                           itmX.SubItems(1) = CStr(.MaximumComponentLength)
                            Set itmX = lstInfo.ListItems.Add(, , "Media Type")
                            itmX.SubItems(1) = CStr(.MediaType)
                             Set itmX = lstInfo.ListItems.Add(, , "Name")
                             itmX.SubItems(1) = CStr(.Name)
                              Set itmX = lstInfo.ListItems.Add(, , "Number Of Blocks")
                              itmX.SubItems(1) = GroupDigits(.NumberOfBlocks)
                               Set itmX = lstInfo.ListItems.Add(, , "Plug-N-Play Device ID")
                               itmX.SubItems(1) = CStr(.PNPDeviceID)
                                 Set itmX = lstInfo.ListItems.Add(, , "Plug-N-Play Device ID")
                                 itmX.SubItems(1) = CStr(.PNPDeviceID)
                                  Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported")
                                  itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                                   Set itmX = lstInfo.ListItems.Add(, , "Provider Name")
                                   itmX.SubItems(1) = CStr(.ProviderName)
                                    Set itmX = lstInfo.ListItems.Add(, , "Purpose")
                                    itmX.SubItems(1) = CStr(.Purpose)
                                     Set itmX = lstInfo.ListItems.Add(, , "Quotas Disabled")
                                     itmX.SubItems(1) = CBoolStr(.QuotasDisabled)
                                      Set itmX = lstInfo.ListItems.Add(, , "Quotas Incomplete")
                                      itmX.SubItems(1) = CBoolStr(.QuotasIncomplete)
                                       Set itmX = lstInfo.ListItems.Add(, , "Quotas Rebuilding")
                                       itmX.SubItems(1) = CBoolStr(.QuotasRebuilding)
                                        Set itmX = lstInfo.ListItems.Add(, , "Size")
                                        itmX.SubItems(1) = FormatByteSize(.Size)
                                         Set itmX = lstInfo.ListItems.Add(, , "Status")
                                         itmX.SubItems(1) = CStr(.Status)
                                          Set itmX = lstInfo.ListItems.Add(, , "Status Info")
                                          itmX.SubItems(1) = CStr(.StatusInfo)
                                           Set itmX = lstInfo.ListItems.Add(, , "Supports Disk Quotas")
                                           itmX.SubItems(1) = CBoolStr(.SupportsDiskQuotas)
                                            Set itmX = lstInfo.ListItems.Add(, , "Supports File Based Compression")
                                            itmX.SubItems(1) = CBoolStr(.SupportsFileBasedCompression)
                                             Set itmX = lstInfo.ListItems.Add(, , "Volume Dirty")
                                             itmX.SubItems(1) = CBoolStr(.VolumeDirty)
                                              Set itmX = lstInfo.ListItems.Add(, , "Volume Name")
                                              itmX.SubItems(1) = CStr(.VolumeName)
                                               Set itmX = lstInfo.ListItems.Add(, , "Volume Serial Number")
                                               itmX.SubItems(1) = CStr(.VolumeSerialNumber)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
End Sub

Sub GetParPort(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object, i&
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
     Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
       Set objItem = objWMI.execquery("Select * from Win32_ParallelPort", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Parallel Ports": lstMultItems.ListItems.Add , "d" & CStr(.DeviceID), CStr(.Caption)
          If Trim(MultiItem$) = Trim(CStr(Item.DeviceID)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Availability")
           itmX.SubItems(1) = CStr(.Availability)
            Set itmX = lstInfo.ListItems.Add(, , "Capabilities")
             tmpBuffer = ""
             For i = LBound(.Capabilities) To UBound(.Capabilities)
              tmpBuffer = tmpBuffer & ", " & .Capabilities(i)
             Next i
              itmX.SubItems(1) = Mid$(tmpBuffer, 3)
             Set itmX = lstInfo.ListItems.Add(, , "Capability Descriptions")
              tmpBuffer = ""
              For i = LBound(.CapabilityDescriptions) To UBound(.CapabilityDescriptions)
               tmpBuffer = tmpBuffer & ", " & .CapabilityDescriptions(i)
              Next i
               itmX.SubItems(1) = Mid$(tmpBuffer, 3)
              Set itmX = lstInfo.ListItems.Add(, , "Caption")
              itmX.SubItems(1) = CBoolStr(.Caption)
               Set itmX = lstInfo.ListItems.Add(, , "Config Manager Error Code")
               itmX.SubItems(1) = CStr(.ConfigManagerErrorCode)
                Set itmX = lstInfo.ListItems.Add(, , "Config Manager User Config")
                itmX.SubItems(1) = CBoolStr(.ConfigManagerUserConfig)
                 Set itmX = lstInfo.ListItems.Add(, , "Description")
                 itmX.SubItems(1) = CStr(.Description)
                  Set itmX = lstInfo.ListItems.Add(, , "Device ID")
                  itmX.SubItems(1) = CStr(.DeviceID)
                   Set itmX = lstInfo.ListItems.Add(, , "DMA Support")
                   itmX.SubItems(1) = CBoolStr(.DMASupport)
                    Set itmX = lstInfo.ListItems.Add(, , "Error Cleared")
                    itmX.SubItems(1) = CBoolStr(.ErrorCleared)
                     Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                     itmX.SubItems(1) = CStr(.ErrorDescription)
                      Set itmX = lstInfo.ListItems.Add(, , "Last Error Code")
                      itmX.SubItems(1) = CStr(.LastErrorCode)
                       Set itmX = lstInfo.ListItems.Add(, , "Max Number Controlled")
                       itmX.SubItems(1) = GroupDigits(.MaxNumberControlled)
                        Set itmX = lstInfo.ListItems.Add(, , "Name")
                        itmX.SubItems(1) = CStr(.Name)
                         Set itmX = lstInfo.ListItems.Add(, , "OS Auto Discovered")
                         itmX.SubItems(1) = CBoolStr(.OSAutoDiscovered)
                          Set itmX = lstInfo.ListItems.Add(, , "Plug-N-Play Device ID")
                          itmX.SubItems(1) = CStr(.PNPDeviceID)
                           Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported")
                           itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                            Set itmX = lstInfo.ListItems.Add(, , "Protocol Supported")
                            itmX.SubItems(1) = CStr(.ProtocolSupported)
                             Set itmX = lstInfo.ListItems.Add(, , "Status")
                             itmX.SubItems(1) = CStr(.Status)
                              Set itmX = lstInfo.ListItems.Add(, , "Status Info")
                              itmX.SubItems(1) = CStr(.StatusInfo)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
End Sub

Sub GetSerPort(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object, i&
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
     Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
       Set objItem = objWMI.execquery("Select * from Win32_SerialPort", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Serial Ports": lstMultItems.ListItems.Add , "d" & CStr(.DeviceID), CStr(.Caption)
          If Trim(MultiItem$) = Trim(CStr(Item.DeviceID)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Availability")
           itmX.SubItems(1) = CStr(.Availability)
            Set itmX = lstInfo.ListItems.Add(, , "Binary")
            itmX.SubItems(1) = CBoolStr(.Binary)
             Set itmX = lstInfo.ListItems.Add(, , "Capabilities")
             tmpBuffer = ""
              For i = LBound(.Capabilities) To UBound(.Capabilities)
               tmpBuffer = tmpBuffer & ", " & .Capabilities(i)
              Next i
               itmX.SubItems(1) = Mid$(tmpBuffer, 3)
              Set itmX = lstInfo.ListItems.Add(, , "Capability Descriptions")
              tmpBuffer = ""
               For i = LBound(.CapabilityDescriptions) To UBound(.CapabilityDescriptions)
                tmpBuffer = tmpBuffer & ", " & .CapabilityDescriptions(i)
               Next i
                itmX.SubItems(1) = Mid$(tmpBuffer, 3)
               Set itmX = lstInfo.ListItems.Add(, , "Caption")
               itmX.SubItems(1) = CStr(.Caption)
                Set itmX = lstInfo.ListItems.Add(, , "Config Manager Error Code")
                itmX.SubItems(1) = CStr(.ConfigManagerErrorCode)
                 Set itmX = lstInfo.ListItems.Add(, , "Config Manager User Config")
                 itmX.SubItems(1) = CBoolStr(.ConfigManagerUserConfig)
                  Set itmX = lstInfo.ListItems.Add(, , "Description")
                  itmX.SubItems(1) = CStr(.Description)
                   Set itmX = lstInfo.ListItems.Add(, , "Device ID")
                   itmX.SubItems(1) = CStr(.DeviceID)
                    Set itmX = lstInfo.ListItems.Add(, , "Error Cleared")
                    itmX.SubItems(1) = CBoolStr(.ErrorCleared)
                     Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                     itmX.SubItems(1) = CStr(.ErrorDescription)
                      Set itmX = lstInfo.ListItems.Add(, , "Last Error Code")
                      itmX.SubItems(1) = CStr(.LastErrorCode)
                       Set itmX = lstInfo.ListItems.Add(, , "Max Baud Rate")
                       itmX.SubItems(1) = CStr(.MaxBaudRate)
                        Set itmX = lstInfo.ListItems.Add(, , "Maximum Input Buffer Size")
                        itmX.SubItems(1) = FormatByteSize(.MaximumInputBufferSize)
                         Set itmX = lstInfo.ListItems.Add(, , "Maximum Output Buffer Size")
                         itmX.SubItems(1) = FormatByteSize(.MaximumOutputBufferSize)
                          Set itmX = lstInfo.ListItems.Add(, , "Max Number Controlled")
                          itmX.SubItems(1) = GroupDigits(.MaxNumberControlled)
                           Set itmX = lstInfo.ListItems.Add(, , "Name")
                           itmX.SubItems(1) = CStr(.Name)
                            Set itmX = lstInfo.ListItems.Add(, , "OS Auto Discovered")
                            itmX.SubItems(1) = CBoolStr(.OSAutoDiscovered)
                             Set itmX = lstInfo.ListItems.Add(, , "PNP Device ID")
                             itmX.SubItems(1) = CStr(.PNPDeviceID)
                              Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported")
                              itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                               Set itmX = lstInfo.ListItems.Add(, , "Protocol Supported")
                               itmX.SubItems(1) = CStr(.ProtocolSupported)
                                Set itmX = lstInfo.ListItems.Add(, , "Provider Type")
                                itmX.SubItems(1) = CStr(.ProviderType)
                                 Set itmX = lstInfo.ListItems.Add(, , "Settable BaudRate")
                                 itmX.SubItems(1) = CBoolStr(.SettableBaudRate)
                                  Set itmX = lstInfo.ListItems.Add(, , "Settable Data Bits")
                                  itmX.SubItems(1) = CBoolStr(.SettableDataBits)
                                   Set itmX = lstInfo.ListItems.Add(, , "Settable Flow Control")
                                   itmX.SubItems(1) = CBoolStr(.SettableFlowControl)
                                    Set itmX = lstInfo.ListItems.Add(, , "Settable Parity")
                                    itmX.SubItems(1) = CBoolStr(.SettableParity)
                                     Set itmX = lstInfo.ListItems.Add(, , "Settable Parity Check")
                                     itmX.SubItems(1) = CBoolStr(.SettableParityCheck)
                                      Set itmX = lstInfo.ListItems.Add(, , "Settable RLSD")
                                      itmX.SubItems(1) = CBoolStr(.SettableRLSD)
                                       Set itmX = lstInfo.ListItems.Add(, , "Settable Stop Bits")
                                       itmX.SubItems(1) = CBoolStr(.SettableStopBits)
                                        Set itmX = lstInfo.ListItems.Add(, , "Status")
                                        itmX.SubItems(1) = CStr(.Status)
                                         Set itmX = lstInfo.ListItems.Add(, , "Status Info")
                                         itmX.SubItems(1) = CStr(.StatusInfo)
                                          Set itmX = lstInfo.ListItems.Add(, , "Supports 16-Bit Mode")
                                          itmX.SubItems(1) = CBoolStr(.Supports16BitMode)
                                           Set itmX = lstInfo.ListItems.Add(, , "Supports DTRDSR")
                                           itmX.SubItems(1) = CBoolStr(.SupportsDTRDSR)
                                            Set itmX = lstInfo.ListItems.Add(, , "Supports Elapsed Timeouts")
                                            itmX.SubItems(1) = CBoolStr(.SupportsElapsedTimeouts)
                                             Set itmX = lstInfo.ListItems.Add(, , "Supports Int Timeouts")
                                             itmX.SubItems(1) = CBoolStr(.SupportsIntTimeouts)
                                              Set itmX = lstInfo.ListItems.Add(, , "Supports Parity Check")
                                              itmX.SubItems(1) = CBoolStr(.SupportsParityCheck)
                                               Set itmX = lstInfo.ListItems.Add(, , "Supports RLSD")
                                               itmX.SubItems(1) = CBoolStr(.SupportsRLSD)
                                                Set itmX = lstInfo.ListItems.Add(, , "Supports RTSCTS")
                                                itmX.SubItems(1) = CBoolStr(.SupportsRTSCTS)
                                                 Set itmX = lstInfo.ListItems.Add(, , "Supports Special Characters")
                                                 itmX.SubItems(1) = CBoolStr(.SupportsSpecialCharacters)
                                                  Set itmX = lstInfo.ListItems.Add(, , "Supports XOn XOff")
                                                  itmX.SubItems(1) = CBoolStr(.SupportsXOnXOff)
                                                   Set itmX = lstInfo.ListItems.Add(, , "Supports XOn XOff Set")
                                                   itmX.SubItems(1) = CBoolStr(.SupportsXOnXOffSet)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
End Sub

Sub GetUSBController(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
     Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
       Set objItem = objWMI.execquery("Select * from Win32_USBController", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "USB Controllers": lstMultItems.ListItems.Add , "d" & CStr(.DeviceID), CStr(.Caption)
          If Trim(MultiItem$) = Trim(CStr(Item.DeviceID)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Availability")
           itmX.SubItems(1) = CStr(.Availability)
            Set itmX = lstInfo.ListItems.Add(, , "Caption")
            itmX.SubItems(1) = CStr(.Caption)
             Set itmX = lstInfo.ListItems.Add(, , "Config Manager Error Code")
             itmX.SubItems(1) = CStr(.ConfigManagerErrorCode)
              Set itmX = lstInfo.ListItems.Add(, , "Config Manager User Config")
              itmX.SubItems(1) = CBoolStr(.ConfigManagerUserConfig)
               Set itmX = lstInfo.ListItems.Add(, , "Description")
               itmX.SubItems(1) = CStr(.Description)
                Set itmX = lstInfo.ListItems.Add(, , "Device ID")
                itmX.SubItems(1) = CStr(.DeviceID)
                 Set itmX = lstInfo.ListItems.Add(, , "Error Cleared")
                 itmX.SubItems(1) = CBoolStr(.ErrorCleared)
                  Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                  itmX.SubItems(1) = CStr(.ErrorDescription)
                   Set itmX = lstInfo.ListItems.Add(, , "Last Error Code")
                   itmX.SubItems(1) = CStr(.LastErrorCode)
                    Set itmX = lstInfo.ListItems.Add(, , "Manufacturer")
                    itmX.SubItems(1) = CStr(.Manufacturer)
                     Set itmX = lstInfo.ListItems.Add(, , "Max Number Controlled")
                     itmX.SubItems(1) = GroupDigits(.MaxNumberControlled)
                      Set itmX = lstInfo.ListItems.Add(, , "Name")
                      itmX.SubItems(1) = CStr(.Name)
                       Set itmX = lstInfo.ListItems.Add(, , "Plug-N-Play Device ID")
                       itmX.SubItems(1) = CStr(.PNPDeviceID)
                        Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported")
                        itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                         Set itmX = lstInfo.ListItems.Add(, , "Protocol Supported")
                         itmX.SubItems(1) = CStr(.ProtocolSupported)
                          Set itmX = lstInfo.ListItems.Add(, , "Status")
                          itmX.SubItems(1) = CStr(.Status)
                           Set itmX = lstInfo.ListItems.Add(, , "Status Info")
                           itmX.SubItems(1) = CStr(.StatusInfo)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
End Sub

Sub GetUSBHub(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
      Set objItem = objWMI.execquery("Select * from Win32_USBHub", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "USB Hubs": lstMultItems.ListItems.Add , "d" & CStr(.DeviceID), CStr(.Caption)
          If Trim(MultiItem$) = Trim(CStr(Item.DeviceID)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Availability")
           itmX.SubItems(1) = CStr(.Availability)
            Set itmX = lstInfo.ListItems.Add(, , "Caption")
            itmX.SubItems(1) = CStr(.Caption)
             Set itmX = lstInfo.ListItems.Add(, , "Class Code")
             itmX.SubItems(1) = CStr(.ClassCode)
              Set itmX = lstInfo.ListItems.Add(, , "Config Manager Error Code")
              itmX.SubItems(1) = CStr(.ConfigManagerErrorCode)
               Set itmX = lstInfo.ListItems.Add(, , "Config Manager User Config")
               itmX.SubItems(1) = CBoolStr(.ConfigManagerUserConfig)
                Set itmX = lstInfo.ListItems.Add(, , "Current Config Value")
                itmX.SubItems(1) = CStr(.CurrentConfigValue)
                 Set itmX = lstInfo.ListItems.Add(, , "Description")
                 itmX.SubItems(1) = CStr(.Description)
                  Set itmX = lstInfo.ListItems.Add(, , "Device ID")
                  itmX.SubItems(1) = CStr(.DeviceID)
                   Set itmX = lstInfo.ListItems.Add(, , "Error Cleared")
                   itmX.SubItems(1) = CBoolStr(.ErrorCleared)
                    Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                    itmX.SubItems(1) = CStr(.ErrorDescription)
                     Set itmX = lstInfo.ListItems.Add(, , "Gang Switched")
                     itmX.SubItems(1) = CBoolStr(.GangSwitched)
                      Set itmX = lstInfo.ListItems.Add(, , "Last Error Code")
                      itmX.SubItems(1) = CStr(.LastErrorCode)
                       Set itmX = lstInfo.ListItems.Add(, , "Name")
                       itmX.SubItems(1) = CStr(.Name)
                        Set itmX = lstInfo.ListItems.Add(, , "Number Of Configs")
                        itmX.SubItems(1) = GroupDigits(.NumberOfConfigs)
                         Set itmX = lstInfo.ListItems.Add(, , "Number Of Ports")
                         itmX.SubItems(1) = GroupDigits(.NumberOfPorts)
                          Set itmX = lstInfo.ListItems.Add(, , "Plug-N-Play Device ID")
                          itmX.SubItems(1) = CStr(.PNPDeviceID)
                           Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported")
                           itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                            Set itmX = lstInfo.ListItems.Add(, , "Protocol Code")
                            itmX.SubItems(1) = CStr(.ProtocolCode)
                             Set itmX = lstInfo.ListItems.Add(, , "Status")
                             itmX.SubItems(1) = CStr(.Status)
                              Set itmX = lstInfo.ListItems.Add(, , "Status Info")
                              itmX.SubItems(1) = CStr(.StatusInfo)
                               Set itmX = lstInfo.ListItems.Add(, , "Subclass Code")
                               itmX.SubItems(1) = CStr(.SubclassCode)
                                Set itmX = lstInfo.ListItems.Add(, , "USB Version")
                                itmX.SubItems(1) = CStr(.USBVersion)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
End Sub

Sub GetDesktopMonitor(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
      Set objItem = objWMI.execquery("Select * from Win32_DesktopMonitor", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Desktop Monitors": lstMultItems.ListItems.Add , "d" & CStr(.DeviceID), CStr(.Caption)
          If Trim(MultiItem$) = Trim(CStr(.DeviceID)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Availability")
           itmX.SubItems(1) = CStr(.Availability)
            Set itmX = lstInfo.ListItems.Add(, , "Caption")
            itmX.SubItems(1) = CStr(.Caption)
             Set itmX = lstInfo.ListItems.Add(, , "Bandwidth")
             itmX.SubItems(1) = CStr(.Bandwidth)
              Set itmX = lstInfo.ListItems.Add(, , "Config Manager Error Code")
              itmX.SubItems(1) = CStr(.ConfigManagerErrorCode)
               Set itmX = lstInfo.ListItems.Add(, , "Config Manager User Config")
               itmX.SubItems(1) = CBoolStr(.ConfigManagerUserConfig)
                Set itmX = lstInfo.ListItems.Add(, , "Description")
                itmX.SubItems(1) = CStr(.Description)
                 Set itmX = lstInfo.ListItems.Add(, , "Device ID")
                 itmX.SubItems(1) = CStr(.DeviceID)
                  Set itmX = lstInfo.ListItems.Add(, , "Display Type")
                  itmX.SubItems(1) = CStr(.DisplayType)
                   Set itmX = lstInfo.ListItems.Add(, , "Error Cleared")
                   itmX.SubItems(1) = CBoolStr(.ErrorCleared)
                    Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                    itmX.SubItems(1) = CStr(.ErrorDescription)
                     Set itmX = lstInfo.ListItems.Add(, , "Is Locked")
                     itmX.SubItems(1) = CBoolStr(.IsLocked)
                      Set itmX = lstInfo.ListItems.Add(, , "Last Error Code")
                      itmX.SubItems(1) = CStr(.LastErrorCode)
                       Set itmX = lstInfo.ListItems.Add(, , "Monitor Manufacturer")
                       itmX.SubItems(1) = CStr(.MonitorManufacturer)
                        Set itmX = lstInfo.ListItems.Add(, , "Monitor Type")
                        itmX.SubItems(1) = CStr(.MonitorType)
                         Set itmX = lstInfo.ListItems.Add(, , "Name")
                         itmX.SubItems(1) = CStr(.Name)
                          Set itmX = lstInfo.ListItems.Add(, , "Pixels Per X Logical Inch")
                          itmX.SubItems(1) = CStr(.PixelsPerXLogicalInch)
                           Set itmX = lstInfo.ListItems.Add(, , "Pixels Per Y Logical Inch")
                           itmX.SubItems(1) = CStr(.PixelsPerYLogicalInch)
                            Set itmX = lstInfo.ListItems.Add(, , "Plug-N-Play Device ID")
                            itmX.SubItems(1) = CStr(.PNPDeviceID)
                             Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported")
                             itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                              Set itmX = lstInfo.ListItems.Add(, , "Screen Height")
                              itmX.SubItems(1) = CStr(.ScreenHeight)
                               Set itmX = lstInfo.ListItems.Add(, , "Screen Width")
                               itmX.SubItems(1) = CStr(.ScreenWidth)
                                Set itmX = lstInfo.ListItems.Add(, , "Status")
                                itmX.SubItems(1) = CStr(.Status)
                                 Set itmX = lstInfo.ListItems.Add(, , "Status Info")
                                 itmX.SubItems(1) = CStr(.StatusInfo)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
End Sub

Sub GetVidController(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object, i&
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
     Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
       Set objItem = objWMI.execquery("Select * from Win32_VideoController", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Video Controllers": lstMultItems.ListItems.Add , "d" & CStr(.DeviceID), CStr(.Description)
          If Trim(MultiItem$) = Trim(CStr(.DeviceID)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Accelerator Capabilities")
            tmpBuffer = ""
            For i = LBound(.AcceleratorCapabilities) To UBound(.AcceleratorCapabilities)
             tmpBuffer = ", " & .AcceleratorCapabilities(i)
            Next i
             tmpBuffer = Mid$(tmpBuffer, 3): itmX.SubItems(1) = tmpBuffer
            Set itmX = lstInfo.ListItems.Add(, , "Adapter Compatibility")
            itmX.SubItems(1) = CStr(.AdapterCompatibility)
             Set itmX = lstInfo.ListItems.Add(, , "Adapter DAC Type")
             itmX.SubItems(1) = CStr(.AdapterDACType)
              Set itmX = lstInfo.ListItems.Add(, , "Adapter RAM")
              itmX.SubItems(1) = FormatByteSize(.AdapterRAM)
               Set itmX = lstInfo.ListItems.Add(, , "Availability")
               itmX.SubItems(1) = CStr(.Availability)
                Set itmX = lstInfo.ListItems.Add(, , "Capability Descriptions")
                tmpBuffer = ""
                 For i = LBound(.CapabilityDescriptions) To UBound(.CapabilityDescriptions)
                  tmpBuffer = ", " & .CapabilityDescriptions(i)
                 Next i
                  itmX.SubItems(1) = Mid$(tmpBuffer, 3)
                 Set itmX = lstInfo.ListItems.Add(, , "Caption")
                 itmX.SubItems(1) = CStr(.Caption)
                  Set itmX = lstInfo.ListItems.Add(, , "Color Table Entries")
                  itmX.SubItems(1) = CStr(.ColorTableEntries)
                   Set itmX = lstInfo.ListItems.Add(, , "Config Manager Error Code")
                   itmX.SubItems(1) = CStr(.ConfigManagerErrorCode)
                    Set itmX = lstInfo.ListItems.Add(, , "Config Manager User Config")
                    itmX.SubItems(1) = CBoolStr(.ConfigManagerUserConfig)
                     Set itmX = lstInfo.ListItems.Add(, , "Current Bits Per Pixel")
                     itmX.SubItems(1) = CStr(.CurrentBitsPerPixel)
                      Set itmX = lstInfo.ListItems.Add(, , "Current Horizontal Resolution")
                      itmX.SubItems(1) = GroupDigits(.CurrentHorizontalResolution)
                       Set itmX = lstInfo.ListItems.Add(, , "Current Number Of Colors")
                       itmX.SubItems(1) = GroupDigits(.CurrentNumberOfColors)
                        Set itmX = lstInfo.ListItems.Add(, , "Current Number Of Columns")
                        itmX.SubItems(1) = GroupDigits(.CurrentNumberOfColumns)
                         Set itmX = lstInfo.ListItems.Add(, , "Manufacturer")
                         itmX.SubItems(1) = CStr(.Manufacturer)
                          Set itmX = lstInfo.ListItems.Add(, , "Current Number Of Rows")
                          itmX.SubItems(1) = GroupDigits(.CurrentNumberOfRows)
                           Set itmX = lstInfo.ListItems.Add(, , "Current Refresh Rate")
                           itmX.SubItems(1) = CStr(.CurrentRefreshRate)
                            Set itmX = lstInfo.ListItems.Add(, , "Current Scan Mode")
                            itmX.SubItems(1) = CStr(.CurrentScanMode)
                             Set itmX = lstInfo.ListItems.Add(, , "Current Vertical Resolution")
                             itmX.SubItems(1) = GroupDigits(.CurrentVerticalResolution)
                              Set itmX = lstInfo.ListItems.Add(, , "Description")
                              itmX.SubItems(1) = CStr(.Description)
                               Set itmX = lstInfo.ListItems.Add(, , "Device ID")
                               itmX.SubItems(1) = CStr(.DeviceID)
                                Set itmX = lstInfo.ListItems.Add(, , "Device Specific Pens")
                                itmX.SubItems(1) = CStr(.DeviceSpecificPens)
                                 Set itmX = lstInfo.ListItems.Add(, , "Dither Type")
                                 itmX.SubItems(1) = CStr(.DitherType)
                                  Set itmX = lstInfo.ListItems.Add(, , "Driver Version")
                                  itmX.SubItems(1) = CStr(.DriverVersion)
                                   Set itmX = lstInfo.ListItems.Add(, , "Error Cleared")
                                   itmX.SubItems(1) = CBoolStr(.ErrorCleared)
                                    Set itmX = lstInfo.ListItems.Add(, , "Error Description")
                                    itmX.SubItems(1) = CStr(.ErrorDescription)
                                     Set itmX = lstInfo.ListItems.Add(, , "ICM Intent")
                                     itmX.SubItems(1) = CStr(.ICMIntent)
                                      Set itmX = lstInfo.ListItems.Add(, , "ICM Method")
                                      itmX.SubItems(1) = CStr(.ICMMethod)
                                       Set itmX = lstInfo.ListItems.Add(, , "INF File Filename")
                                       itmX.SubItems(1) = CStr(.InfFilename)
                                        Set itmX = lstInfo.ListItems.Add(, , "INF File Section")
                                        itmX.SubItems(1) = CStr(.InfSection)
                                         Set itmX = lstInfo.ListItems.Add(, , "Installed Display Drivers")
                                         itmX.SubItems(1) = CStr(.InstalledDisplayDrivers)
                                          Set itmX = lstInfo.ListItems.Add(, , "Last Error Code")
                                          itmX.SubItems(1) = CStr(.LastErrorCode)
                                           Set itmX = lstInfo.ListItems.Add(, , "Max Memory Supported")
                                           itmX.SubItems(1) = FormatByteSize(.MaxMemorySupported)
                                            Set itmX = lstInfo.ListItems.Add(, , "Max Number Controlled")
                                            itmX.SubItems(1) = GroupDigits(.MaxNumberControlled)
                                             Set itmX = lstInfo.ListItems.Add(, , "Max Refresh Rate")
                                             itmX.SubItems(1) = CStr(.MaxRefreshRate)
                                              Set itmX = lstInfo.ListItems.Add(, , "Monochrome")
                                              itmX.SubItems(1) = CBoolStr(.Monochrome)
                                               Set itmX = lstInfo.ListItems.Add(, , "Min Refresh Rate")
                                               itmX.SubItems(1) = CStr(.MinRefreshRate)
                                                Set itmX = lstInfo.ListItems.Add(, , "Name")
                                                itmX.SubItems(1) = CStr(.Name)
                                                 Set itmX = lstInfo.ListItems.Add(, , "Number Of Color Planes")
                                                 itmX.SubItems(1) = GroupDigits(.NumberOfColorPlanes)
                                                  Set itmX = lstInfo.ListItems.Add(, , "Number Of Video Pages")
                                                  itmX.SubItems(1) = GroupDigits(.NumberOfVideoPages)
                                                   Set itmX = lstInfo.ListItems.Add(, , "Plug-N-Play Device ID")
                                                   itmX.SubItems(1) = CStr(.PNPDeviceID)
                                                    Set itmX = lstInfo.ListItems.Add(, , "Power Management Supported")
                                                    itmX.SubItems(1) = CBoolStr(.PowerManagementSupported)
                                                     Set itmX = lstInfo.ListItems.Add(, , "Protocol Supported")
                                                     itmX.SubItems(1) = CStr(.ProtocolSupported)
                                                      Set itmX = lstInfo.ListItems.Add(, , "Reserved System Palette Entries")
                                                      itmX.SubItems(1) = CStr(.ReservedSystemPaletteEntries)
                                                       Set itmX = lstInfo.ListItems.Add(, , "Specification Version")
                                                       itmX.SubItems(1) = CStr(.SpecificationVersion)
                                                        Set itmX = lstInfo.ListItems.Add(, , "Status")
                                                        itmX.SubItems(1) = CStr(.Status)
                                                         Set itmX = lstInfo.ListItems.Add(, , "Status Info")
                                                         itmX.SubItems(1) = CStr(.StatusInfo)
                                                          Set itmX = lstInfo.ListItems.Add(, , "System Palette Entries")
                                                          itmX.SubItems(1) = CStr(.SystemPaletteEntries)
                                                           Set itmX = lstInfo.ListItems.Add(, , "Video Architecture")
                                                           itmX.SubItems(1) = CStr(.VideoArchitecture)
                                                            Set itmX = lstInfo.ListItems.Add(, , "Video Memory Type")
                                                            itmX.SubItems(1) = CStr(.VideoMemoryType)
                                                             Set itmX = lstInfo.ListItems.Add(, , "Video Mode")
                                                             itmX.SubItems(1) = CStr(.VideoMode)
                                                              Set itmX = lstInfo.ListItems.Add(, , "Video Mode Description")
                                                              itmX.SubItems(1) = CStr(.VideoModeDescription)
                                                               Set itmX = lstInfo.ListItems.Add(, , "Video Processor")
                                                               itmX.SubItems(1) = CStr(.VideoProcessor)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
End Sub

Sub GetLogMemConfig(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
      If Trim(MultiItem) <> "" Then
       Set objItem = objWMI.execquery("Select * from Win32_LogicalMemoryConfiguration where SettingID = """ & MultiItem & """", , 48)
      Else
       Set objItem = objWMI.execquery("Select * from Win32_LogicalMemoryConfiguration", , 48)
      End If
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Logical Memory Configurations": lstMultItems.ListItems.Add , "d" & CStr(.SettingID), CStr(.Caption)
          If tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Available Virtual Memory")
           itmX.SubItems(1) = FormatByteSize(.AvailableVirtualMemory)
            Set itmX = lstInfo.ListItems.Add(, , "Caption")
            itmX.SubItems(1) = CStr(.Caption)
             Set itmX = lstInfo.ListItems.Add(, , "Description")
             itmX.SubItems(1) = CStr(.Description)
              Set itmX = lstInfo.ListItems.Add(, , "Name")
              itmX.SubItems(1) = CStr(.Name)
               Set itmX = lstInfo.ListItems.Add(, , "Setting ID")
               itmX.SubItems(1) = CStr(.SettingID)
                Set itmX = lstInfo.ListItems.Add(, , "Total Page File Space")
                itmX.SubItems(1) = FormatByteSize(.TotalPageFileSpace)
                 Set itmX = lstInfo.ListItems.Add(, , "Total Physical Memory")
                 itmX.SubItems(1) = FormatByteSize(.TotalPhysicalMemory)
                  Set itmX = lstInfo.ListItems.Add(, , "Total Virtual Memory")
                  itmX.SubItems(1) = FormatByteSize(.TotalVirtualMemory)

          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
           lstClasses.ListItems.Clear
            lstClasses.ListItems.Add , "Win32_LogicalMemoryConfiguration", "Logical Memory Config"
End Sub

Sub GetOperatingSystem(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
      Set objItem = objWMI.execquery("Select * from Win32_OperatingSystem", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Operating Systems": lstMultItems.ListItems.Add , "d" & CStr(.Name), CStr(.Caption)
          If Trim(MultiItem$) = Trim(CStr(.Name)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Boot Device")
           itmX.SubItems(1) = CStr(.BootDevice)
            Set itmX = lstInfo.ListItems.Add(, , "Build Number")
            itmX.SubItems(1) = CStr(.BuildNumber)
             Set itmX = lstInfo.ListItems.Add(, , "Build Type")
             itmX.SubItems(1) = CStr(.BuildType)
              Set itmX = lstInfo.ListItems.Add(, , "Caption")
              itmX.SubItems(1) = CStr(.Caption)
               Set itmX = lstInfo.ListItems.Add(, , "Code Set")
               itmX.SubItems(1) = CStr(.CodeSet)
                Set itmX = lstInfo.ListItems.Add(, , "Country Code")
                itmX.SubItems(1) = CStr(.CountryCode)
                 Set itmX = lstInfo.ListItems.Add(, , "CSD Version")
                 itmX.SubItems(1) = CStr(.CSDVersion)
                  Set itmX = lstInfo.ListItems.Add(, , "Computer System Name")
                  itmX.SubItems(1) = CStr(.CSName)
                   Set itmX = lstInfo.ListItems.Add(, , "Current Time Zone")
                   itmX.SubItems(1) = CStr(.CurrentTimeZone)
                    Set itmX = lstInfo.ListItems.Add(, , "Debug")
                    itmX.SubItems(1) = CBoolStr(.Debug)
                     Set itmX = lstInfo.ListItems.Add(, , "Description")
                     itmX.SubItems(1) = CStr(.Description)
                      Set itmX = lstInfo.ListItems.Add(, , "Distributed")
                      itmX.SubItems(1) = CBoolStr(.Distributed)
                       Set itmX = lstInfo.ListItems.Add(, , "Encryption Level")
                       itmX.SubItems(1) = CStr(.EncryptionLevel)
                        Set itmX = lstInfo.ListItems.Add(, , "Foreground Application Boost")
                        itmX.SubItems(1) = CStr(.ForegroundApplicationBoost)
                         Set itmX = lstInfo.ListItems.Add(, , "Free Physical Memory")
                         itmX.SubItems(1) = FormatByteSize(.FreePhysicalMemory)
                          Set itmX = lstInfo.ListItems.Add(, , "Free Space In Paging Files")
                          itmX.SubItems(1) = FormatByteSize(.FreeSpaceInPagingFiles)
                           Set itmX = lstInfo.ListItems.Add(, , "Free Virtual Memory")
                           itmX.SubItems(1) = FormatByteSize(.FreeVirtualMemory)
                            Set itmX = lstInfo.ListItems.Add(, , "Large System Cache")
                            itmX.SubItems(1) = FormatByteSize(.LargeSystemCache)
                             Set itmX = lstInfo.ListItems.Add(, , "Locale")
                             itmX.SubItems(1) = CStr(.Locale)
                              Set itmX = lstInfo.ListItems.Add(, , "Manufacturer")
                              itmX.SubItems(1) = CStr(.Manufacturer)
                               Set itmX = lstInfo.ListItems.Add(, , "Max Number Of Processes")
                               itmX.SubItems(1) = GroupDigits(.MaxNumberOfProcesses)
                                Set itmX = lstInfo.ListItems.Add(, , "Max Process Memory Size")
                                itmX.SubItems(1) = FormatByteSize(.MaxProcessMemorySize)
                                 Set itmX = lstInfo.ListItems.Add(, , "Name")
                                 itmX.SubItems(1) = CStr(.Name)
                                  Set itmX = lstInfo.ListItems.Add(, , "Number Of Licensed Users")
                                  itmX.SubItems(1) = GroupDigits(.NumberOfLicensedUsers)
                                   Set itmX = lstInfo.ListItems.Add(, , "Number Of Processes")
                                   itmX.SubItems(1) = GroupDigits(.NumberOfProcesses)
                                    Set itmX = lstInfo.ListItems.Add(, , "Number Of Users")
                                    itmX.SubItems(1) = GroupDigits(.NumberOfUsers)
                                     Set itmX = lstInfo.ListItems.Add(, , "Organization")
                                     itmX.SubItems(1) = CStr(.Organization)
                                      Set itmX = lstInfo.ListItems.Add(, , "Operating System Language")
                                      itmX.SubItems(1) = CStr(.OSLanguage)
                                       Set itmX = lstInfo.ListItems.Add(, , "Operating System Product Suite")
                                       itmX.SubItems(1) = CStr(.OSProductSuite)
                                        Set itmX = lstInfo.ListItems.Add(, , "Operating System Type")
                                        itmX.SubItems(1) = CStr(.OSType)
                                         Set itmX = lstInfo.ListItems.Add(, , "Other Type Description")
                                         itmX.SubItems(1) = CStr(.OtherTypeDescription)
                                          Set itmX = lstInfo.ListItems.Add(, , "Plus Product ID")
                                          itmX.SubItems(1) = CStr(.PlusProductID)
                                           Set itmX = lstInfo.ListItems.Add(, , "Plus Version Number")
                                           itmX.SubItems(1) = CStr(.PlusVersionNumber)
                                            Set itmX = lstInfo.ListItems.Add(, , "Primary")
                                            itmX.SubItems(1) = CBoolStr(.Primary)
                                             Set itmX = lstInfo.ListItems.Add(, , "Product Type")
                                             itmX.SubItems(1) = CStr(.ProductType)
                                              Set itmX = lstInfo.ListItems.Add(, , "Quantum Length")
                                              itmX.SubItems(1) = CStr(.QuantumLength)
                                               Set itmX = lstInfo.ListItems.Add(, , "Quantum Type")
                                               itmX.SubItems(1) = CStr(.QuantumType)
                                                Set itmX = lstInfo.ListItems.Add(, , "Registered User")
                                                itmX.SubItems(1) = CStr(.RegisteredUser)
                                                 Set itmX = lstInfo.ListItems.Add(, , "Status")
                                                 itmX.SubItems(1) = CStr(.Status)
                                                  Set itmX = lstInfo.ListItems.Add(, , "Serial Number")
                                                  itmX.SubItems(1) = CStr(.SerialNumber)
                                                   Set itmX = lstInfo.ListItems.Add(, , "Service Pack Major Version")
                                                   itmX.SubItems(1) = CStr(.ServicePackMajorVersion)
                                                    Set itmX = lstInfo.ListItems.Add(, , "Service Pack Minor Version")
                                                    itmX.SubItems(1) = CStr(.ServicePackMinorVersion)
                                                     Set itmX = lstInfo.ListItems.Add(, , "Size Stored In Paging Files")
                                                     itmX.SubItems(1) = FormatByteSize(.SizeStoredInPagingFiles)
                                                      Set itmX = lstInfo.ListItems.Add(, , "Status")
                                                      itmX.SubItems(1) = CStr(.Status)
                                                       Set itmX = lstInfo.ListItems.Add(, , "Suite Mask")
                                                       itmX.SubItems(1) = CStr(.SuiteMask)
                                                        Set itmX = lstInfo.ListItems.Add(, , "System Device")
                                                        itmX.SubItems(1) = CStr(.SystemDevice)
                                                         Set itmX = lstInfo.ListItems.Add(, , "System Directory")
                                                         itmX.SubItems(1) = CStr(.SystemDirectory)
                                                          Set itmX = lstInfo.ListItems.Add(, , "System Drive")
                                                          itmX.SubItems(1) = CStr(.SystemDrive)
                                                           Set itmX = lstInfo.ListItems.Add(, , "Total Swap Space Size")
                                                           itmX.SubItems(1) = FormatByteSize(.TotalSwapSpaceSize)
                                                            Set itmX = lstInfo.ListItems.Add(, , "Total Virtual Memory Size")
                                                            itmX.SubItems(1) = FormatByteSize(.TotalVirtualMemorySize)
                                                             Set itmX = lstInfo.ListItems.Add(, , "Total Visible Memory Size")
                                                             itmX.SubItems(1) = FormatByteSize(.TotalVisibleMemorySize)
                                                              Set itmX = lstInfo.ListItems.Add(, , "Version")
                                                              itmX.SubItems(1) = CStr(.Version)
                                                               Set itmX = lstInfo.ListItems.Add(, , "Windows Directory")
                                                               itmX.SubItems(1) = CStr(.WindowsDirectory)
           End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
           lstClasses.ListItems.Clear
            lstClasses.ListItems.Add , "Win32_OperatingSystem", "Operating Systems"
End Sub

Function ParseProcCmdLine(ByVal CommandLine$) As String
 If Left$(CommandLine$, 1) = """" And InStr(1, CommandLine$, """ ", 1) > 0 Then
  ParseProcCmdLine = Mid(CommandLine$, InStr(InStr(1, CommandLine$, """", 1) + 1, CommandLine$, """ ", 1) + 1)
 ElseIf Left(CommandLine$, 1) <> """" And InStr(1, CommandLine$, " ", 1) > 0 Then
  ParseProcCmdLine = Mid$(CommandLine$, InStr(1, CommandLine$, " ", 1) + 1)
 End If
End Function

Sub GetProcesses(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, tmpCmdLine$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
      If Trim(MultiItem) <> "" Then
       Set objItem = objWMI.execquery("Select * from Win32_Process where ProcessID = """ & MultiItem & """", , 48)
      Else
       Set objItem = objWMI.execquery("Select * from Win32_Process", , 48)
      End If
        For Each Item In objItem
         With Item
         Err.Clear
          tmpCmdLine$ = " " & ParseProcCmdLine(.CommandLine)
           If Err.Number Then tmpCmdLine$ = ""
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Processes": lstMultItems.ListItems.Add , "d" & CStr(.ProcessID), CStr(.Caption) & " " & tmpCmdLine$
          If tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, "Caption", "Caption")
            itmX.SubItems(1) = Item.Caption
             Set itmX = lstInfo.ListItems.Add(, , "Command Line")
             itmX.SubItems(1) = .CommandLine
              Set itmX = lstInfo.ListItems.Add(, , "Computer")
              itmX.SubItems(1) = Item.CSName
               Set itmX = lstInfo.ListItems.Add(, , "Description")
               itmX.SubItems(1) = Item.Description
                Set itmX = lstInfo.ListItems.Add(, , "Executable Path")
                itmX.SubItems(1) = Item.ExecutablePath
                 Set itmX = lstInfo.ListItems.Add(, , "Handle")
                 itmX.SubItems(1) = Item.Handle
                  Set itmX = lstInfo.ListItems.Add(, , "Handle Count")
                  itmX.SubItems(1) = GroupDigits(Item.HandleCount)
                   Set itmX = lstInfo.ListItems.Add(, , "Kernel Mode Time")
                   itmX.SubItems(1) = CStr(Item.KernelModeTime)
                    Set itmX = lstInfo.ListItems.Add(, , "Max Working Size")
                    itmX.SubItems(1) = FormatByteSize(Item.MaximumWorkingSetSize)
                     Set itmX = lstInfo.ListItems.Add(, , "Min Working Size")
                     itmX.SubItems(1) = FormatByteSize(Item.MinimumWorkingSetSize)
                      Set itmX = lstInfo.ListItems.Add(, , "Name")
                      itmX.SubItems(1) = Item.Name
                       Set itmX = lstInfo.ListItems.Add(, , "Other Operation Count")
                       itmX.SubItems(1) = GroupDigits(Item.OtherOperationCount)
                        Set itmX = lstInfo.ListItems.Add(, , "Other Transfer Count")
                        itmX.SubItems(1) = GroupDigits(Item.OtherTransferCount)
                         Set itmX = lstInfo.ListItems.Add(, , "Page Faults")
                         itmX.SubItems(1) = GroupDigits(Item.PageFaults)
                          Set itmX = lstInfo.ListItems.Add(, , "Page File Usage")
                          itmX.SubItems(1) = FormatByteSize(Item.PagefileUsage)
                           Set itmX = lstInfo.ListItems.Add(, , "Parent Process")
                           itmX.SubItems(1) = CStr(Item.ParentProcessId)
                            Set itmX = lstInfo.ListItems.Add(, , "Peak Page File Usage")
                            itmX.SubItems(1) = FormatByteSize(Item.PeakPagefileUsage)
                             Set itmX = lstInfo.ListItems.Add(, , "Peak Virtual Size")
                             itmX.SubItems(1) = FormatByteSize(Item.PeakVirtualSize)
                              Set itmX = lstInfo.ListItems.Add(, , "Peak Working Set")
                              itmX.SubItems(1) = FormatByteSize(Item.PeakWorkingSetSize)
                               Set itmX = lstInfo.ListItems.Add(, , "Priority")
                               itmX.SubItems(1) = CStr(Item.Priority)
                                Set itmX = lstInfo.ListItems.Add(, , "Private Page Count")
                                itmX.SubItems(1) = GroupDigits(Item.PrivatePageCount)
                                 Set itmX = lstInfo.ListItems.Add(, , "Process ID")
                                 itmX.SubItems(1) = CStr(Item.ProcessID)
                                  Set itmX = lstInfo.ListItems.Add(, , "Quota NonPaged Pool Usage")
                                  itmX.SubItems(1) = CStr(Item.QuotaNonPagedPoolUsage)
                                   Set itmX = lstInfo.ListItems.Add(, , "Quota Paged Pool Usage")
                                   itmX.SubItems(1) = CStr(Item.QuotaPagedPoolUsage)
                                    Set itmX = lstInfo.ListItems.Add(, , "Quota Peak NonPaged Pool Usage")
                                    itmX.SubItems(1) = CStr(Item.QuotaPeakNonPagedPoolUsage)
                                     Set itmX = lstInfo.ListItems.Add(, , "Quota Peak Paged Pool Usage")
                                     itmX.SubItems(1) = CStr(Item.QuotaPeakPagedPoolUsage)
                                      Set itmX = lstInfo.ListItems.Add(, , "Read Operation Count")
                                      itmX.SubItems(1) = CStr(Item.ReadOperationCount)
                                       Set itmX = lstInfo.ListItems.Add(, , "Read Transfer Count")
                                       itmX.SubItems(1) = CStr(Item.ReadTransferCount)
                                        Set itmX = lstInfo.ListItems.Add(, , "Session ID")
                                        itmX.SubItems(1) = CStr(Item.SessionId)
                                         Set itmX = lstInfo.ListItems.Add(, , "Status")
                                         itmX.SubItems(1) = Item.SessionId
                                          Set itmX = lstInfo.ListItems.Add(, , "Thread Count")
                                          itmX.SubItems(1) = GroupDigits(Item.ThreadCount)
                                           Set itmX = lstInfo.ListItems.Add(, , "User Mode Time")
                                           itmX.SubItems(1) = CStr(Item.UserModeTime)
                                            Set itmX = lstInfo.ListItems.Add(, , "Virtual Size")
                                            itmX.SubItems(1) = FormatByteSize(Item.VirtualSize)
                                             Set itmX = lstInfo.ListItems.Add(, , "Windows Version")
                                             itmX.SubItems(1) = Item.WindowsVersion
                                              Set itmX = lstInfo.ListItems.Add(, , "Working Set Size")
                                              itmX.SubItems(1) = FormatByteSize(Item.WorkingSetSize)
                                               Set itmX = lstInfo.ListItems.Add(, , "Write Operation Count")
                                               itmX.SubItems(1) = GroupDigits(Item.WriteOperationCount)
                                                Set itmX = lstInfo.ListItems.Add(, , "Write Transfer Count")
                                                itmX.SubItems(1) = GroupDigits(Item.WriteTransferCount)
           End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
           lstClasses.ListItems.Clear
            lstClasses.ListItems.Add , "Win32_Process", "Processes"
End Sub

Sub GetEnvirnStrings(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
      Set objItem = objWMI.execquery("Select * from Win32_Environment", , 48)
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Environment Variables": lstMultItems.ListItems.Add , "d" & CStr(.Name) & CStr(.Description), CStr(.Name)
          If Trim(MultiItem$) = Trim(CStr(.Name) & CStr(.Description)) Or Len(MultiItem$) = 0 And tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Caption")
            itmX.SubItems(1) = Item.Caption
             Set itmX = lstInfo.ListItems.Add(, , "Description")
             itmX.SubItems(1) = CStr(.Description)
              Set itmX = lstInfo.ListItems.Add(, , "Name")
              itmX.SubItems(1) = CStr(.Name)
               Set itmX = lstInfo.ListItems.Add(, , "Status")
               itmX.SubItems(1) = CStr(.Status)
                Set itmX = lstInfo.ListItems.Add(, , "System Variable")
                itmX.SubItems(1) = CBoolStr(.SystemVariable)
                 Set itmX = lstInfo.ListItems.Add(, , "User Name")
                 itmX.SubItems(1) = CStr(.UserName)
                  Set itmX = lstInfo.ListItems.Add(, , "Variable Value")
                  itmX.SubItems(1) = CStr(.VariableValue)
           End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
           lstClasses.ListItems.Clear
            lstClasses.ListItems.Add , "Win32_Environment", "Environment Variables"
End Sub

Sub GetDMAChannel(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object, i&
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
      If Trim(MultiItem) <> "" Then
       Set objItem = objWMI.execquery("Select * from Win32_DMAChannel where DMAChannel = """ & MultiItem & """", , 48)
      Else
       Set objItem = objWMI.execquery("Select * from Win32_DMAChannel", , 48)
      End If
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "DMA Channels": lstMultItems.ListItems.Add , "d" & CStr(.DMAChannel), CStr(.Caption)
          If tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Address Size")
           itmX.SubItems(1) = CStr(.AddressSize)
            Set itmX = lstInfo.ListItems.Add(, , "Availability")
            itmX.SubItems(1) = CStr(.Availability)
             Set itmX = lstInfo.ListItems.Add(, , "Burst Mode")
             itmX.SubItems(1) = CBoolStr(.BurstMode)
              Set itmX = lstInfo.ListItems.Add(, , "Byte Mode")
              itmX.SubItems(1) = CStr(.ByteMode)
               Set itmX = lstInfo.ListItems.Add(, , "Caption")
               itmX.SubItems(1) = CStr(.Caption)
                Set itmX = lstInfo.ListItems.Add(, , "Channel Timing")
                itmX.SubItems(1) = CStr(.ChannelTiming)
                 Set itmX = lstInfo.ListItems.Add(, , "Computer System Name")
                 itmX.SubItems(1) = CStr(.CSName)
                  Set itmX = lstInfo.ListItems.Add(, , "Description")
                  itmX.SubItems(1) = CStr(.Description)
                   Set itmX = lstInfo.ListItems.Add(, , "DMA Channel")
                   itmX.SubItems(1) = CStr(.DMAChannel)
                    Set itmX = lstInfo.ListItems.Add(, , "Max Transfer Size")
                    itmX.SubItems(1) = FormatByteSize(.MaxTransferSize)
                     Set itmX = lstInfo.ListItems.Add(, , "Name")
                     itmX.SubItems(1) = CStr(.Name)
                      Set itmX = lstInfo.ListItems.Add(, , "Port")
                      itmX.SubItems(1) = CStr(.Port)
                       Set itmX = lstInfo.ListItems.Add(, , "Status")
                       itmX.SubItems(1) = CStr(.Status)
                        Set itmX = lstInfo.ListItems.Add(, , "Transfer Widths")
                         tmpBuffer = ""
                          For i = LBound(.TransferWidths) To UBound(.TransferWidths)
                           tmpBuffer = ", " & .TransferWidths(i)
                          Next i
                           itmX.SubItems(1) = Mid$(tmpBuffer, 3)
                         Set itmX = lstInfo.ListItems.Add(, , "Type C Timing")
                         itmX.SubItems(1) = CStr(.TypeCTiming)
                          Set itmX = lstInfo.ListItems.Add(, , "Word Mode")
                          itmX.SubItems(1) = CStr(.WordMode)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
           lstClasses.ListItems.Clear
            lstClasses.ListItems.Add , "Win32_DMAChannel", "DMA Channels"
End Sub

Sub GetIRQResource(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
      If Trim(MultiItem) <> "" Then
       Set objItem = objWMI.execquery("Select * from Win32_IRQResource where IRQNumber = """ & MultiItem & """", , 48)
      Else
       Set objItem = objWMI.execquery("Select * from Win32_IRQResource", , 48)
      End If
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "IRQ Resources": lstMultItems.ListItems.Add , "d" & CStr(.IRQNumber), CStr(.Caption)
          If tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Availability")
           itmX.SubItems(1) = CStr(.Availability)
            Set itmX = lstInfo.ListItems.Add(, , "Caption")
            itmX.SubItems(1) = CStr(.Caption)
             Set itmX = lstInfo.ListItems.Add(, , "Computer System Name")
             itmX.SubItems(1) = CStr(.CSName)
              Set itmX = lstInfo.ListItems.Add(, , "Description")
              itmX.SubItems(1) = CStr(.Description)
               Set itmX = lstInfo.ListItems.Add(, , "Hardware")
               itmX.SubItems(1) = CBoolStr(.Hardware)
                Set itmX = lstInfo.ListItems.Add(, , "IRQ Number")
                itmX.SubItems(1) = CStr(.IRQNumber)
                 Set itmX = lstInfo.ListItems.Add(, , "Name")
                 itmX.SubItems(1) = CStr(.Name)
                  Set itmX = lstInfo.ListItems.Add(, , "Shareable")
                  itmX.SubItems(1) = CBoolStr(.Shareable)
                   Set itmX = lstInfo.ListItems.Add(, , "Status")
                   itmX.SubItems(1) = CStr(.Status)
                    Set itmX = lstInfo.ListItems.Add(, , "Trigger Level")
                    itmX.SubItems(1) = CStr(.TriggerLevel)
                     Set itmX = lstInfo.ListItems.Add(, , "Trigger Type")
                     itmX.SubItems(1) = CStr(.TriggerType)
                      Set itmX = lstInfo.ListItems.Add(, , "Vector")
                      itmX.SubItems(1) = CStr(.Vector)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
           lstClasses.ListItems.Clear
            lstClasses.ListItems.Add , "Win32_IRQResource", "IRQ Resources"
End Sub


Sub GetDeviceMemAdd(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
     If Trim(MultiItem) <> "" Then
      Set objItem = objWMI.execquery("Select * from Win32_DeviceMemoryAddress where  StartingAddress  = """ & MultiItem & """", , 48)
     Else
      Set objItem = objWMI.execquery("Select * from Win32_DeviceMemoryAddress", , 48)
     End If
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Device Memory Addresses": lstMultItems.ListItems.Add , "d" & CStr(.StartingAddress), CStr(.Caption)
          If tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Caption")
           itmX.SubItems(1) = CStr(.Caption)
            Set itmX = lstInfo.ListItems.Add(, , "Computer System Name")
            itmX.SubItems(1) = CStr(.CSName)
             Set itmX = lstInfo.ListItems.Add(, , "Description")
             itmX.SubItems(1) = CStr(.Description)
              Set itmX = lstInfo.ListItems.Add(, , "Ending Address")
              itmX.SubItems(1) = CStr(.EndingAddress)
               Set itmX = lstInfo.ListItems.Add(, , "Memory Type")
               itmX.SubItems(1) = CStr(.MemoryType)
                Set itmX = lstInfo.ListItems.Add(, , "Name")
                itmX.SubItems(1) = CStr(.Name)
                 Set itmX = lstInfo.ListItems.Add(, , "Starting Address")
                 itmX.SubItems(1) = CStr(.StartingAddress)
                  Set itmX = lstInfo.ListItems.Add(, , "Status")
                  itmX.SubItems(1) = CBoolStr(.Status)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
           lstClasses.ListItems.Clear
            lstClasses.ListItems.Add , "Win32_DeviceMemoryAddress", "Device Memory Addresses"
End Sub

Sub GetPortResource(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
     If Trim(MultiItem) <> "" Then
      Set objItem = objWMI.execquery("Select * from Win32_PortResource where Caption  = """ & MultiItem & """", , 48)
     Else
      Set objItem = objWMI.execquery("Select * from Win32_PortResource", , 48)
     End If
        For Each Item In objItem
         With Item
          Err.Clear
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Port Resources": lstMultItems.ListItems.Add , "d" & CStr(.Caption), CStr(.StartingAddress) & " - " & CStr(.EndingAddress) & " (" & CStr(.Caption) & ")"
           If Err.Number Then lstMultItems.ListItems.Add , "d" & CStr(.Caption), CStr(.Caption)
          If tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Alias")
           itmX.SubItems(1) = CBoolStr(.Alias)
            Set itmX = lstInfo.ListItems.Add(, , "Caption")
            itmX.SubItems(1) = CStr(.Caption)
             Set itmX = lstInfo.ListItems.Add(, , "Computer System Name")
             itmX.SubItems(1) = CStr(.CSName)
              Set itmX = lstInfo.ListItems.Add(, , "Description")
              itmX.SubItems(1) = CStr(.Description)
               Set itmX = lstInfo.ListItems.Add(, , "Ending Address")
               itmX.SubItems(1) = CStr(.EndingAddress)
                Set itmX = lstInfo.ListItems.Add(, , "Name")
                itmX.SubItems(1) = CStr(.Name)
                 Set itmX = lstInfo.ListItems.Add(, , "Starting Address")
                 itmX.SubItems(1) = CStr(.StartingAddress)
                  Set itmX = lstInfo.ListItems.Add(, , "Status")
                  itmX.SubItems(1) = CBoolStr(.Status)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
           lstClasses.ListItems.Clear
            lstClasses.ListItems.Add , "Win32_PortResource", "Port Resources"
End Sub


Sub GetServices(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
      If Trim(MultiItem) <> "" Then
       Set objItem = objWMI.execquery("Select * from Win32_Service where name = """ & MultiItem & """", , 48)
      Else
       Set objItem = objWMI.execquery("Select * from Win32_Service", , 48)
      End If
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Services": lstMultItems.ListItems.Add , "d" & CStr(.Name), CStr(.DisplayName)
          If tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Accept Pause")
           itmX.SubItems(1) = CBoolStr(.AcceptPause)
            Set itmX = lstInfo.ListItems.Add(, , "Accept Stop")
            itmX.SubItems(1) = CStr(.AcceptStop)
             Set itmX = lstInfo.ListItems.Add(, , "Caption")
             itmX.SubItems(1) = CStr(.Caption)
              Set itmX = lstInfo.ListItems.Add(, , "Description")
              itmX.SubItems(1) = CStr(.Description)
               Set itmX = lstInfo.ListItems.Add(, , "Desktop Interactaction")
               itmX.SubItems(1) = CBoolStr(.DesktopInteract)
                Set itmX = lstInfo.ListItems.Add(, , "Display Name")
                itmX.SubItems(1) = CStr(.DisplayName)
                 Set itmX = lstInfo.ListItems.Add(, , "Error Control")
                 itmX.SubItems(1) = CStr(.ErrorControl)
                  Set itmX = lstInfo.ListItems.Add(, , "Exit Code")
                  itmX.SubItems(1) = CStr(.ExitCode)
                   Set itmX = lstInfo.ListItems.Add(, "name", "Name")
                   itmX.SubItems(1) = CStr(.Name)
                    Set itmX = lstInfo.ListItems.Add(, , "Path Name")
                    itmX.SubItems(1) = CStr(.PathName)
                     Set itmX = lstInfo.ListItems.Add(, , "Process Id")
                     itmX.SubItems(1) = CStr(.ProcessID)
                      Set itmX = lstInfo.ListItems.Add(, , "Service Specific Exit Code")
                      itmX.SubItems(1) = CStr(.ServiceSpecificExitCode)
                       Set itmX = lstInfo.ListItems.Add(, , "Service Type")
                       itmX.SubItems(1) = CStr(.ServiceType)
                        Set itmX = lstInfo.ListItems.Add(, , "Started")
                        itmX.SubItems(1) = CBoolStr(.Started)
                         If .Started = False Then cmdStartStop.Caption = "Start" Else cmdStartStop.Caption = "Stop"
                         Set itmX = lstInfo.ListItems.Add(, , "Start Mode")
                         itmX.SubItems(1) = CStr(.StartMode)
                          Set itmX = lstInfo.ListItems.Add(, , "State")
                          itmX.SubItems(1) = CStr(.State)
                           Set itmX = lstInfo.ListItems.Add(, , "Status")
                           itmX.SubItems(1) = CStr(.Status)
                            Set itmX = lstInfo.ListItems.Add(, , "Tag Id")
                            itmX.SubItems(1) = CStr(.TagId)
                             Set itmX = lstInfo.ListItems.Add(, , "Wait Hint")
                             itmX.SubItems(1) = CStr(.WaitHint)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
           lstClasses.ListItems.Clear
            lstClasses.ListItems.Add , "Win32_Service", "Services"
             lstClasses.ListItems.Add , "Win32_TerminalService", "Terminal Services"
End Sub

Sub GetTermService(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
      If Trim(MultiItem) <> "" Then
       Set objItem = objWMI.execquery("Select * from Win32_TerminalService where name = """ & MultiItem & """", , 48)
      Else
       Set objItem = objWMI.execquery("Select * from Win32_TerminalService", , 48)
      End If
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Terminal Services": lstMultItems.ListItems.Add , "d" & CStr(.Name), CStr(.DisplayName)
          If tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Accept Pause")
           itmX.SubItems(1) = CBoolStr(.AcceptPause)
            Set itmX = lstInfo.ListItems.Add(, , "Accept Stop")
            itmX.SubItems(1) = CStr(.AcceptStop)
             Set itmX = lstInfo.ListItems.Add(, , "Caption")
             itmX.SubItems(1) = CStr(.Caption)
              Set itmX = lstInfo.ListItems.Add(, , "Check Point")
              itmX.SubItems(1) = CStr(.CheckPoint)
               Set itmX = lstInfo.ListItems.Add(, , "Description")
               itmX.SubItems(1) = CStr(.Description)
                Set itmX = lstInfo.ListItems.Add(, , "Desktop Interactaction")
                itmX.SubItems(1) = CBoolStr(.DesktopInteract)
                 Set itmX = lstInfo.ListItems.Add(, , "Disconnected Sessions")
                 itmX.SubItems(1) = CStr(.DisconnectedSessions)
                  Set itmX = lstInfo.ListItems.Add(, , "Display Name")
                  itmX.SubItems(1) = CStr(.DisplayName)
                   Set itmX = lstInfo.ListItems.Add(, , "Error Control")
                   itmX.SubItems(1) = CStr(.ErrorControl)
                    Set itmX = lstInfo.ListItems.Add(, , "Estimated Session Capacity")
                    itmX.SubItems(1) = CStr(.EstimatedSessionCapacity)
                    Set itmX = lstInfo.ListItems.Add(, , "Exit Code")
                    itmX.SubItems(1) = CStr(.ExitCode)
                     Set itmX = lstInfo.ListItems.Add(, "name", "Name")
                     itmX.SubItems(1) = CStr(.Name)
                      Set itmX = lstInfo.ListItems.Add(, , "Path Name")
                      itmX.SubItems(1) = CStr(.PathName)
                       Set itmX = lstInfo.ListItems.Add(, , "Process Id")
                       itmX.SubItems(1) = CStr(.ProcessID)
                        Set itmX = lstInfo.ListItems.Add(, , "Raw Session Capacity")
                        itmX.SubItems(1) = CStr(.RawSessionCapacity)
                         Set itmX = lstInfo.ListItems.Add(, , "Resource Constraint")
                         itmX.SubItems(1) = CStr(.ResourceConstraint)
                          Set itmX = lstInfo.ListItems.Add(, , "Service Specific Exit Code")
                          itmX.SubItems(1) = CStr(.ServiceSpecificExitCode)
                           Set itmX = lstInfo.ListItems.Add(, , "Service Type")
                           itmX.SubItems(1) = CStr(.ServiceType)
                            Set itmX = lstInfo.ListItems.Add(, , "Started")
                            itmX.SubItems(1) = CBoolStr(.Started)
                             Set itmX = lstInfo.ListItems.Add(, , "Start Mode")
                             itmX.SubItems(1) = CStr(.StartMode)
                              Set itmX = lstInfo.ListItems.Add(, , "Start Name")
                              itmX.SubItems(1) = CStr(.StartName)
                               Set itmX = lstInfo.ListItems.Add(, , "State")
                               itmX.SubItems(1) = CStr(.State)
                                Set itmX = lstInfo.ListItems.Add(, , "Status")
                                itmX.SubItems(1) = CStr(.Status)
                                 Set itmX = lstInfo.ListItems.Add(, , "Tag Id")
                                 itmX.SubItems(1) = CStr(.TagId)
                                  Set itmX = lstInfo.ListItems.Add(, , "Total Sessions")
                                  itmX.SubItems(1) = GroupDigits(.TotalSessions)
                                   Set itmX = lstInfo.ListItems.Add(, , "Wait Hint")
                                   itmX.SubItems(1) = CStr(.WaitHint)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
End Sub

Sub GetSystemDriver(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
      If Trim(MultiItem) <> "" Then
       Set objItem = objWMI.execquery("Select * from Win32_SystemDriver where name = """ & MultiItem & """", , 48)
      Else
       Set objItem = objWMI.execquery("Select * from Win32_SystemDriver", , 48)
      End If
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "System Drivers": lstMultItems.ListItems.Add , "d" & CStr(.Name), CStr(.Description)
          If tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Accept Pause")
           itmX.SubItems(1) = CBoolStr(.AcceptPause)
            Set itmX = lstInfo.ListItems.Add(, , "Accept Stop")
            itmX.SubItems(1) = CStr(.AcceptStop)
             Set itmX = lstInfo.ListItems.Add(, , "Caption")
             itmX.SubItems(1) = CStr(.Caption)
              Set itmX = lstInfo.ListItems.Add(, , "Description")
              itmX.SubItems(1) = CStr(.Description)
                Set itmX = lstInfo.ListItems.Add(, , "Desktop Interactaction")
                itmX.SubItems(1) = CBoolStr(.DesktopInteract)
                 Set itmX = lstInfo.ListItems.Add(, , "Display Name")
                 itmX.SubItems(1) = CStr(.DisplayName)
                   Set itmX = lstInfo.ListItems.Add(, , "Error Control")
                   itmX.SubItems(1) = CStr(.ErrorControl)
                    Set itmX = lstInfo.ListItems.Add(, , "Exit Code")
                    itmX.SubItems(1) = CStr(.ExitCode)
                     Set itmX = lstInfo.ListItems.Add(, "name", "Name")
                     itmX.SubItems(1) = CStr(.Name)
                      Set itmX = lstInfo.ListItems.Add(, , "Path Name")
                      itmX.SubItems(1) = CStr(.PathName)
                        Set itmX = lstInfo.ListItems.Add(, , "Service Specific Exit Code")
                        itmX.SubItems(1) = CStr(.ServiceSpecificExitCode)
                         Set itmX = lstInfo.ListItems.Add(, , "Service Type")
                         itmX.SubItems(1) = CStr(.ServiceType)
                          Set itmX = lstInfo.ListItems.Add(, , "Started")
                          itmX.SubItems(1) = CBoolStr(.Started)
                           Set itmX = lstInfo.ListItems.Add(, , "Start Mode")
                           itmX.SubItems(1) = CStr(.StartMode)
                            Set itmX = lstInfo.ListItems.Add(, , "Start Name")
                            itmX.SubItems(1) = CStr(.StartName)
                             Set itmX = lstInfo.ListItems.Add(, , "State")
                             itmX.SubItems(1) = CStr(.State)
                              Set itmX = lstInfo.ListItems.Add(, , "Status")
                              itmX.SubItems(1) = CStr(.Status)
                               Set itmX = lstInfo.ListItems.Add(, , "Tag Id")
                               itmX.SubItems(1) = CStr(.TagId)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
           lstClasses.ListItems.Clear
            lstClasses.ListItems.Add , "Win32_SystemDriver", "System Drivers"
End Sub

Sub GetTimeZone(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
      Set objItem = objWMI.execquery("Select * from Win32_TimeZone", , 48)
      For Each Item In objItem
       With Item
        tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "Time Zone Settings": lstMultItems.ListItems.Add , "d" & CStr(.Description), CStr(.Description)
        If tmpCnt& = 1 Then
          Set itmX = lstInfo.ListItems.Add(, , "Bias")
          itmX.SubItems(1) = CStr(.Bias)
           Set itmX = lstInfo.ListItems.Add(, , "Caption")
           itmX.SubItems(1) = CStr(.Caption)
            Set itmX = lstInfo.ListItems.Add(, , "Daylight Bias")
            itmX.SubItems(1) = CStr(.DaylightBias)
             Set itmX = lstInfo.ListItems.Add(, , "Daylight Day")
             itmX.SubItems(1) = CStr(.DaylightDay)
              Set itmX = lstInfo.ListItems.Add(, , "Daylight Day Of Week")
              itmX.SubItems(1) = CStr(.DaylightDayOfWeek)
               Set itmX = lstInfo.ListItems.Add(, , "Daylight Hour")
               itmX.SubItems(1) = CStr(.DaylightHour)
                Set itmX = lstInfo.ListItems.Add(, , "Daylight Millisecond")
                itmX.SubItems(1) = CStr(.DaylightMillisecond)
                 Set itmX = lstInfo.ListItems.Add(, , "Daylight Minute")
                 itmX.SubItems(1) = CStr(.DaylightMinute)
                  Set itmX = lstInfo.ListItems.Add(, , "Daylight Month")
                  itmX.SubItems(1) = CStr(.DaylightMonth)
                   Set itmX = lstInfo.ListItems.Add(, , "Daylight Name")
                   itmX.SubItems(1) = CStr(.DaylightName)
                   Set itmX = lstInfo.ListItems.Add(, , "Daylight Second")
                   itmX.SubItems(1) = CStr(.DaylightSecond)
                    Set itmX = lstInfo.ListItems.Add(, , "Daylight Year")
                    itmX.SubItems(1) = CStr(.DaylightYear)
                     Set itmX = lstInfo.ListItems.Add(, , "Description")
                     itmX.SubItems(1) = CStr(.Description)
                      Set itmX = lstInfo.ListItems.Add(, , "Setting ID")
                      itmX.SubItems(1) = CStr(.SettingID)
                       Set itmX = lstInfo.ListItems.Add(, , "Standard Bias")
                       itmX.SubItems(1) = CStr(.StandardBias)
                        Set itmX = lstInfo.ListItems.Add(, , "Standard Day")
                        itmX.SubItems(1) = CStr(.StandardDay)
                         Set itmX = lstInfo.ListItems.Add(, , "Standard Day Of Week")
                         itmX.SubItems(1) = CStr(.StandardDayOfWeek)
                          Set itmX = lstInfo.ListItems.Add(, , "Standard Hour")
                          itmX.SubItems(1) = CStr(.StandardHour)
                           Set itmX = lstInfo.ListItems.Add(, , "Standard Millisecond")
                           itmX.SubItems(1) = CStr(.StandardMillisecond)
                            Set itmX = lstInfo.ListItems.Add(, , "Standard Minute")
                            itmX.SubItems(1) = CStr(.StandardMinute)
                             Set itmX = lstInfo.ListItems.Add(, , "Standard Month")
                             itmX.SubItems(1) = CStr(.StandardMonth)
                              Set itmX = lstInfo.ListItems.Add(, , "Standard Name")
                              itmX.SubItems(1) = CStr(.StandardName)
                               Set itmX = lstInfo.ListItems.Add(, , "Standard Second")
                               itmX.SubItems(1) = CStr(.StandardSecond)
                                Set itmX = lstInfo.ListItems.Add(, , "Standard Year")
                                itmX.SubItems(1) = CStr(.StandardYear)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
           lstClasses.ListItems.Clear
            lstClasses.ListItems.Add , "Win32_TimeZone", "Time Zone Settings"
End Sub

Sub GetUserAccount(Optional MultiItem$)
On Error Resume Next
Dim objWMI As Object, objItem As Object, itmX As ListItem, tmpCnt&, tmpBuffer$, Item As Object
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
      If Trim(MultiItem) <> "" Then
       Set objItem = objWMI.execquery("Select * from Win32_UserAccount where name = """ & MultiItem & """", , 48)
      Else
       Set objItem = objWMI.execquery("Select * from Win32_UserAccount", , 48)
      End If
        For Each Item In objItem
         With Item
          tmpCnt& = tmpCnt& + 1: lstMultItems.ColumnHeaders(1).Text = "User Accounts": lstMultItems.ListItems.Add , "d" & CStr(.Name), CStr(.Name)
          If tmpCnt& = 1 Then
           Set itmX = lstInfo.ListItems.Add(, , "Account Type")
           itmX.SubItems(1) = CStr(.AccountType)
            Set itmX = lstInfo.ListItems.Add(, , "Caption")
            itmX.SubItems(1) = CStr(.Caption)
             Set itmX = lstInfo.ListItems.Add(, , "Description")
             itmX.SubItems(1) = CStr(.Description)
              Set itmX = lstInfo.ListItems.Add(, , "Disabled")
              itmX.SubItems(1) = CBoolStr(.Disabled)
                Set itmX = lstInfo.ListItems.Add(, , "Domain (System Name)")
                itmX.SubItems(1) = CStr(.Domain)
                 Set itmX = lstInfo.ListItems.Add(, , "Full Name")
                 itmX.SubItems(1) = CStr(.FullName)
                   Set itmX = lstInfo.ListItems.Add(, , "Local Account")
                   itmX.SubItems(1) = CBoolStr(.LocalAccount)
                    Set itmX = lstInfo.ListItems.Add(, , "Lockout")
                    itmX.SubItems(1) = CBoolStr(.Lockout)
                     Set itmX = lstInfo.ListItems.Add(, , "Name")
                     itmX.SubItems(1) = CStr(.Name)
                      Set itmX = lstInfo.ListItems.Add(, , "Password Changeable")
                      itmX.SubItems(1) = CBoolStr(.PasswordChangeable)
                        Set itmX = lstInfo.ListItems.Add(, , "Password Expires")
                        itmX.SubItems(1) = CBoolStr(.PasswordExpires)
                         Set itmX = lstInfo.ListItems.Add(, , "Password Required")
                         itmX.SubItems(1) = CBoolStr(.PasswordRequired)
                          Set itmX = lstInfo.ListItems.Add(, , "SID")
                          itmX.SubItems(1) = CStr(.SID)
                           Set itmX = lstInfo.ListItems.Add(, , "SID Type")
                           itmX.SubItems(1) = CStr(.SIDType)
                            Set itmX = lstInfo.ListItems.Add(, , "Status")
                            itmX.SubItems(1) = CStr(.Status)
          End If
         End With
        Next Item
         Set objWMI = Nothing
          Set objItem = Nothing
           lstClasses.ListItems.Clear
            lstClasses.ListItems.Add , "Win32_UserAccount", "User Accounts"
End Sub

Private Function QueryInfo(ClassID$, Optional MultiItem$) As Boolean
 'If the length of MultiItem evaluates to zero, then remove all items in the MultiItem listview control
 If Len(MultiItem$) = 0 Then lstMultItems.ListItems.Clear
   lstInfo.ListItems.Clear 'Remove all items in the lstInfo listview control
    Select Case ClassID 'select case statement (See MSDN help system)
     Case "Win32_Processor": 'If ClassID evaluates to Win32_Processor then...
      GetProcessorInfo MultiItem
      'See GetProcessorInfo function for more info,
      'and pass MultiItem as the MultiItem argument, if this variable doesn't equal to nothing then this item will be selected
     Case "Win32_BIOS": '..
      GetBIOS MultiItem '..
     Case "Win32_BootConfiguration":
      GetBootConfig MultiItem
     Case "Win32_DiskPartition":
      GetDiskPartition MultiItem
     Case "Win32_CDROMDrive":
      GetCDROM MultiItem
     Case "Win32_DiskDrive":
      GetDiskDrive MultiItem
     Case "Win32_FloppyController":
      GetFloppyController MultiItem
     Case "Win32_IDEController":
      GetIDEController MultiItem
     Case "Win32_Keyboard":
      GetKeyBoard MultiItem
     Case "Win32_MotherBoardDevice":
      GetMotherBoard MultiItem
     Case "Win32_NetworkAdapter":
      GetNetworkAdapter MultiItem
     Case "Win32_PnPEntity":
      GetPnPEntity MultiItem
     Case "Win32_PotsModem":
      GetPotsModem MultiItem
     Case "Win32_Printer":
      GetPrinter MultiItem
     Case "Win32_SoundDevice":
      GetSoundDevice MultiItem
     Case "Win32_LogicalDisk":
      GetLogicalDisk MultiItem
     Case "Win32_ParallelPort":
      GetParPort MultiItem
     Case "Win32_SerialPort":
      GetSerPort MultiItem
     Case "Win32_USBController":
      GetUSBController MultiItem
     Case "Win32_USBHub":
      GetUSBHub MultiItem
     Case "Win32_DesktopMonitor":
      GetDesktopMonitor MultiItem
     Case "Win32_VideoController":
      GetVidController MultiItem
     Case "Win32_LogicalMemoryConfiguration":
      GetLogMemConfig MultiItem
     Case "Win32_OperatingSystem":
      GetOperatingSystem MultiItem
     Case "Win32_Process":
      GetProcesses MultiItem
     Case "Win32_Environment":
      GetEnvirnStrings MultiItem
     Case "Win32_DMAChannel":
      GetDMAChannel MultiItem
     Case "Win32_IRQResource"
      GetIRQResource MultiItem
     Case "Win32_DeviceMemoryAddress":
      GetDeviceMemAdd MultiItem
     Case "Win32_PortResource":
      GetPortResource MultiItem
     Case "Win32_Service":
      GetServices MultiItem
'       lstInfo.Height = 4935
        'lstInfo.Top = 2520
         cmdStartStop.Visible = True
         'Move controls and make command button cmdStartStop visible
         'This button will allow users to Start or Stop a service
          Exit Function
          'discontinue execution of this procedure...
     Case "Win32_TerminalService"
      GetTermService MultiItem
     Case "Win32_SystemDriver"
      GetSystemDriver MultiItem
     Case "Win32_TimeZone":
      GetTimeZone MultiItem
     Case "Win32_UserAccount":
      GetUserAccount MultiItem
    End Select
     'lstInfo.Height = 5295: lstInfo.Top = 2160:  cmdStartStop.Visible = False
     'Services are not being displayed, so move controls back to their original coordinates and dimensions
End Function

Private Function CBoolStr(bVal) As String
On Error Resume Next
'on the event of an error resume execution on the next line
 If IsNull(bVal) = True Then CBoolStr = "<Undetermined>": Exit Function
 'If bVal evaluates to Null then return "<undetermined>", exit this procedure
  If bVal = True Then CBoolStr = "True" Else CBoolStr = "False"
  'return appropriate string based on bVal's value...
  'This function will convert the number 0 to the string "False" and 1 to the string "True"
End Function

Function FormatByteSize(ByVal bSize, Optional strAppend$ = "") As String
On Error GoTo errh 'On the event of an error jump to the label errh
Dim tmpBuffer$ 'Dimensionalize tmpBuffer as string type
If IsNull(bSize) = True Then FormatByteSize = "<Undetermined Size>": Exit Function
'If variable bSize is NULL (contains no valid data) then return "<undetermined size>"
 If bSize = 0 Then FormatByteSize = "0 Bytes " & strAppend$: Exit Function
 'if bsize evaluates to zero then return appropriate string with the optional strAppend arguments value appended
  If bSize < 1024 Then FormatByteSize = FormatNumber(bSize, 0, , , vbTrue) & " Bytes " & strAppend$: Exit Function
  'If bSize can be converted to a larger measurement unit then convert it, and format the number by grouping the digits, and append both the measurement unit used and the optional strAppend argument
  If bSize >= 1024 Then bSize = bSize / 1024: tmpBuffer = " KB's " & strAppend$
   'If bSize can be converted to a larger measurement unit then convert it, and format the number by grouping the digits, and append both the measurement unit used and the optional strAppend argument
   If bSize >= 1024 Then bSize = bSize / 1024: tmpBuffer = " MB's " & strAppend$
    'If bSize can be converted to a larger measurement unit then convert it, and format the number by grouping the digits, and append both the measurement unit used and the optional strAppend argument
    If bSize >= 1024 Then bSize = bSize / 1024: tmpBuffer = " GB's " & strAppend$
     'If bSize can be converted to a larger measurement unit then convert it, and format the number by grouping the digits, and append both the measurement unit used and the optional strAppend argument
     If bSize >= 1024 Then bSize = bSize / 1024: tmpBuffer = " TB's " & strAppend$
      FormatByteSize = CStr(FormatNumber(bSize, 0, , , vbTrue)) & tmpBuffer: tmpBuffer = ""
       'return no decimal, and group digits(##,###)
       Exit Function 'exit this procedure as no error has occured
errh: 'label errh
 FormatByteSize = IIf(IsNull(bSize), "", bSize & " bytes")
 'Use IIf operator to conditionally return a value
 'If bSize is Null then return "", else return the initial value of the variable and its initial measurement
End Function

Function GroupDigits(ByVal bNum) As String
'This function just groups the digits and removes decimals and returns the string value
 If IsNull(bNum) = True Then GroupDigits = "<Undetermined>": Exit Function
  GroupDigits = FormatNumber$(bNum, 0, , , vbTrue) ' (###,###)
End Function
