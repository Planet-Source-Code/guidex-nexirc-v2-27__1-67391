Attribute VB_Name = "mdlMovieX"
Option Explicit
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_EXPLORER = &H80000
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Const lKey = "W34NSD84D"

'Function SaveDialog(Form1 As Form, Filter As String, Title As String, InitDir As String) As String
'On Local Error Resume Next
'Dim o As OPENFILENAME, l As Long
'o.lStructSize = Len(o)
'o.hwndOwner = Form1.hWnd
'o.hInstance = App.hInstance
'If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
'For l = 1 To Len(Filter)
'    If Mid$(Filter, l, 1) = "|" Then Mid$(Filter, l, 1) = Chr$(0)
'Next
'o.lpstrFilter = Filter
'o.lpstrFile = Space$(254)
'o.nMaxFile = 255
'o.lpstrFileTitle = Space$(254)
'o.nMaxFileTitle = 255
'o.lpstrInitialDir = InitDir
'o.lpstrTitle = Title
'o.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
'l = GetSaveFileName(o)
'If (l) Then
'    SaveDialog = Trim$(o.lpstrFile)
'Else
'    SaveDialog = ""
'End If
'End Function

'Public Function DoesFileExist(lFileName As String) As Boolean
'On Local Error Resume Next
'Dim msg As String
'msg = Dir(lFileName)
'If msg <> "" Then
'    DoesFileExist = True
'Else
'    DoesFileExist = False
'End If
'End Function

'Function OpenDialog(lhWnd As Long, Filter As String, Title As String, InitDir As String) As String
'On Local Error Resume Next
'Dim o As OPENFILENAME, l As Long
'o.lStructSize = Len(o)
'o.hwndOwner = lhWnd
'o.hInstance = App.hInstance
'If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
''For l = 1 To Len(Filter)
'    If Mid$(Filter, l, 1) = "|" Then Mid$(Filter, l, 1) = Chr$(0)
'Next
'o.lpstrFilter = Filter
'o.lpstrFile = Space$(254)
'o.nMaxFile = 255
'o.lpstrFileTitle = Space$(254)
'o.nMaxFileTitle = 255
'o.lpstrInitialDir = InitDir
'o.lpstrTitle = Title
'o.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
'l = GetOpenFileName(o)
'If (l) Then
'    Dim msg As String
'    msg = Trim$(o.lpstrFile)
'    OpenDialog = Left(msg, Len(msg) - 1)
'Else
'    OpenDialog = ""
'End If
'End Function

'Public Function TestKeyValid(lUserName As String, lPassword As String) As Boolean
'On Local Error Resume Next
'Dim msg As String
'msg = KeyGen(lUserName, lKey, 1)
'If msg = lPassword Then
'    TestKeyValid = True
'End If
'End Function

'Function KeyGen(kName As String, kPass As String, kType As Integer) As String
'On Local Error GoTo ErrHandler
'Dim cTable(512) As Integer, nKeys(16) As Integer, s0(512) As Integer, nArray(16) As Integer, pArray(16) As Integer, n As Integer, nPtr As Integer, cPtr As Integer, cFlip As Boolean, sIni As Integer, temp As Integer, rtn As Integer, gKey As String, nLen As Integer, pLen As Integer, kPtr As Integer, sPtr As Integer, nOffset As Integer, pOffset As Integer, tOffset As Integer
'Const nXor As Integer = 18, pXor As Integer = 25, cLw As Integer = 65, nLw As Integer = 48, sOffset As Integer = 0
'nLen = Len(kName)
'pLen = Len(kPass)
'nKeys(1) = 52
'nKeys(2) = 69
'nKeys(3) = 149
'nKeys(4) = 37
'nKeys(5) = 403
'nKeys(6) = 20
'nKeys(7) = 58
'nKeys(8) = 29
'nKeys(9) = 123
'nKeys(10) = 84
'nKeys(11) = 201
'nKeys(12) = 202
'nKeys(13) = 34
'nKeys(14) = 38
'nKeys(15) = 73
'nKeys(16) = 30
'sIni = 0
'For n = 0 To 512
'    s0(n) = n
'Next n
'For n = 0 To 512
'    sIni = (sOffset + sIni + n) Mod 256
'    temp = s0(n)
'    s0(n) = s0(sIni)
'    s0(sIni) = temp
'Next n
'If kType = 1 Then
'    nPtr = 0
'    For n = 0 To 512
'        cTable(s0(n)) = (nLw + (nPtr))
'        nPtr = nPtr + 1
'        If nPtr = 10 Then nPtr = 0
'    Next n
'    gKey = String(16, " ")
'ElseIf kType = 2 Then
'    nPtr = 0
'    cPtr = 0
'    cFlip = False
'    For n = 0 To 512
'        If cFlip Then
'            cTable(s0(n)) = (nLw + nPtr)
'            nPtr = nPtr + 1
'            If nPtr = 10 Then nPtr = 0
'            cFlip = False
'        Else
'            cTable(s0(n)) = (cLw + cPtr)
'            cPtr = cPtr + 1
'            If cPtr = 26 Then cPtr = 0
'            cFlip = True
'        End If
'    Next n
'    gKey = String(16, " ")
'Else
'    gKey = String(19, " ")
'End If
'kPtr = 1
'For n = 1 To nLen
'  nArray(kPtr) = nArray(kPtr) + Asc(Mid(kName, n, 1)) Xor nXor
'  nOffset = nOffset + nArray(kPtr)
'  kPtr = kPtr + 1
'    If kPtr = 9 Then kPtr = 1
'Next n
'For n = 1 To pLen
'  pArray(kPtr) = pArray(kPtr) + Asc(Mid(kPass, n, 1)) Xor pXor
'  pOffset = pOffset + pArray(kPtr)
'  kPtr = kPtr + 1
'    If kPtr = 9 Then kPtr = 1
'Next n
'tOffset = (nOffset + pOffset) Mod 512
'kPtr = 1
'sPtr = 1
'For n = 1 To 16
'    pArray(n) = pArray(n) Xor nKeys(n)
'    rtn = Abs(((nArray(n) Xor pArray(n)) Mod 512) - tOffset)
'    If kType = 3 Then
'        If rtn < 16 Then
'            Mid(gKey, kPtr, 2) = "0" & Hex(rtn)
'        Else
'            Mid(gKey, kPtr, 2) = Hex(rtn)
'        End If
'        If sPtr = 2 And kPtr < 18 Then
'            kPtr = kPtr + 1
'            Mid(gKey, kPtr + 1, 1) = "-"
'        End If
'        kPtr = kPtr + 2
'        sPtr = sPtr + 1
'        If sPtr = 3 Then sPtr = 1
'    Else
'        Mid(gKey, n, 1) = Chr(cTable(rtn))
'    End If
'Next
'KeyGen = gKey
'Exit Function
'ErrHandler:
'    Err.Clear
'End Function

Sub Main()

End Sub
