Attribute VB_Name = "basMain"
Option Explicit
Global gCurrentDir As String
Global gsVBPath As String
Global gsAppName As String

Public Const appName = "AliasExe32"
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'DirParse
Public Enum enumMethod
  eFullPath = 1
  eFilename = 2
  sExtension = 3
  eBaseFileName = 4
  eBaseExtension = 5
End Enum

Public lnghWnd As Long
Public Const MAX_PATH = 260
Public Const WM_CLOSE = &H10
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'*******************************************

Private Const STILL_ACTIVE = &H103
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const SYNCHRONIZE = &H100000
Public Const WAIT_FAILED = -1&        'Error on call
Public Const WAIT_OBJECT_0 = 0        'Normal completion
Public Const WAIT_ABANDONED = &H80&   '
Public Const WAIT_TIMEOUT = &H102&    'Timeout period elapsed
Public Const IGNORE = 0               'Ignore signal
Public Const INFINITE = -1&           'Infinite timeout
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
'CommonDialog*************************
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHOWHELP = &H10
Public Const OFS_MAXPATHNAME = 128

Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_FILEMUSTEXIST _
             Or OFN_NODEREFERENCELINKS

Public Type OpenFileName
    nStructSize As Long
    hwndOwner As Long
    hInstance As Long
    sFilter As String
    sCustomFilter As String
    nCustFilterSize As Long
    nFilterIndex As Long
    sfile As String
    nFileSize As Long
    sFileTitle As String
    nTitleSize As Long
    sInitDir As String
    sDlgTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExt As Integer
    sDefFileExt As String
    nCustDataSize As Long
    fnHook As Long
    sTemplateName As String
 End Type

Public OFN As OpenFileName
Public Declare Function GetOpenFileName Lib "comdlg32.dll" _
    Alias "GetOpenFileNameA" _
   (pOpenfilename As OpenFileName) As Long
'******************
'Browse Dialog Contstants*************************************
Type SHITEMID   ' mkid
    cb As Long       ' Size of the ID (including cb itself)
    abID() As Byte  ' The item ID (variable length)
End Type
Type ITEMIDLIST   ' idl
    mkid As SHITEMID
End Type

Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" _
  (ByVal hwndOwner As Long, ByVal nFolder As Long, _
  pIdl As ITEMIDLIST) As Long

Public Const NOERROR = 0
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const CSIDL_DRIVES = &H11

Private Type BROWSEINFO
    hwndOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32" (ByRef pBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidRes As Long, ByVal pszFolder As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pVoid As Long)
'Create Directory*******************************************************************
Private Type SECURITY_ATTRIBUTES

nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Function fEnumWindows() As Boolean
Dim hwnd As Long
'
' The EnumWindows function enumerates all top-level windows
' on the screen by passing the handle of each window, in turn,
' to an application-defined callback function. EnumWindows
' continues until the last top-level window is enumerated or
' the callback function returns FALSE.
'
Call EnumWindows(AddressOf fEnumWindowsCallBack, hwnd)
End Function
'See if project is loaded
Public Function fEnumWindowsCallBack(ByVal hwnd As Long, ByVal lpData As Long) As Long
Dim lResult    As Long
Dim sWndName   As String
fEnumWindowsCallBack = 1
sWndName = Space$(MAX_PATH)
lResult = GetWindowText(hwnd, sWndName, MAX_PATH)
sWndName = Left$(sWndName, lResult)
If InStr(sWndName, "Microsoft Visual Basic") Then
   If InStr(sWndName, gsAppName) Then
     lnghWnd = hwnd
   End If
End If
End Function

Public Sub CreateNewDirectory(NewDirectory As String)
    Dim sDirTest As String
    Dim SecAttrib As SECURITY_ATTRIBUTES
    Dim bSuccess As Boolean
    Dim sPath As String
    Dim iCounter As Integer
    Dim sTempDir As String
    
    sPath = NewDirectory
    
    If Right(sPath, Len(sPath)) <> "\" Then
        sPath = sPath & "\"
    End If
    
    iCounter = 1
    
    Do Until InStr(iCounter, sPath, "\") = 0
        iCounter = InStr(iCounter, sPath, "\")
        sTempDir = Left(sPath, iCounter)
        sDirTest = Dir(sTempDir)
        iCounter = iCounter + 1
        'create directory
        SecAttrib.lpSecurityDescriptor = &O0
        SecAttrib.bInheritHandle = False
        SecAttrib.nLength = Len(SecAttrib)
        bSuccess = CreateDirectory(sTempDir, SecAttrib)
    Loop

End Sub

Public Function ParseString(ByVal sParse As String, ByVal Delimiter As String) As String
      Dim sChar As String
      Dim temp As String
      Dim ipos As Integer
      Dim i As Integer

    For i = 1 To Len(sParse)
        sChar$ = Mid(sParse, i, 1)
        temp$ = temp$ + sChar$
        ipos = InStr(1, temp$, Delimiter)
        If ipos > 0 Then
            temp$ = Left$(temp$, ipos - 1)
        End If
    Next
      ParseString = temp$

End Function

Public Function FileExists(sSpec As String) As Integer
    Dim fileFile As Integer
    'attempt to open file
    fileFile = FreeFile
    On Error Resume Next
    Open sSpec For Input As fileFile
    'check for error
    If Err Then
        FileExists = False
    Else
        'file exists
        'close file
        Close fileFile
        FileExists = True
    End If
End Function

Public Function OpenFileName(fForm As Form, iType As Integer) As String
  Dim result As Long
  Dim sp As Long
  Dim LongName As String, sReturn As String
  Dim shortName As String
  Dim ShortSize As Long
  Dim n As String
  Dim n2 As String
  Dim f As String
  n = Chr$(0)
  n2 = n & n
  OFN.nStructSize = Len(OFN)
  OFN.hwndOwner = fForm.hwnd
         Select Case iType
           Case 1 'Browse Project
             f = "VB Project Files" & n & "*.vbp" & n2
             OFN.sFilter = f
             OFN.nFilterIndex = 1
             OFN.sfile = "" & Space$(260) & n
             OFN.nFileSize = Len(OFN.sfile)
             OFN.sDefFileExt = ".vbp"
             OFN.sFileTitle = Space$(512)
             OFN.nTitleSize = Len(OFN.sFileTitle)
             OFN.sInitDir = gsVBPath
             OFN.sDlgTitle = "Choose a .VBP File to Open"
             OFN.flags = OFS_FILE_OPEN_FLAGS
             result = GetOpenFileName(OFN)
             If result Then
                sReturn = Left$(OFN.sfile, InStr(OFN.sfile, vbNullChar) - 1)
              End If
             If Len(sReturn) > 0 Then
               OpenFileName = Trim(sReturn)
             Else
              OpenFileName = vbNullString
             End If
           Case 2 'browse for vb6.exe
             f = "VB6 Executable" & n & "VB6.exe" & n2
             OFN.sFilter = f
             OFN.nFilterIndex = 1
             OFN.sfile = "" & Space$(260) & n
             OFN.nFileSize = Len(OFN.sfile)
             OFN.sDefFileExt = "VB6.exe"
             OFN.sFileTitle = Space$(512)
             OFN.nTitleSize = Len(OFN.sFileTitle)
             OFN.sInitDir = "c:\"
             OFN.sDlgTitle = "Select VB6 Executable"
             OFN.flags = OFS_FILE_OPEN_FLAGS
             result = GetOpenFileName(OFN)
             If result Then
                sReturn = Left$(OFN.sfile, InStr(OFN.sfile, vbNullChar) - 1) 'Trim(OFN.sfile)
              End If
             If Len(sReturn) > 0 Then
               OpenFileName = Trim(sReturn)
             Else
               OpenFileName = vbNullString
             End If
          Case Else
        End Select
End Function
Public Function BrowseForFolder(fForm As Form, sFolder As String, ByVal sTitle As String) As Boolean
    Dim bRes As Boolean
    bRes = False
    
    Dim BI As BROWSEINFO, pItemIDList As Long
    
    BI.hwndOwner = fForm.hwnd
    BI.pidlRoot = 0
    BI.pszDisplayName = sFolder
    BI.lpszTitle = sTitle
    BI.ulFlags = 0
    BI.lpfn = 0
    BI.lParam = 0
    BI.iImage = 0
    pItemIDList = SHBrowseForFolder(BI)
    
    If pItemIDList Then
        sFolder = Space(261)
        If SHGetPathFromIDList(pItemIDList, sFolder) Then
            sFolder = Left(sFolder, InStr(sFolder, vbNullChar) - 1)
            bRes = True
        End If
        
        CoTaskMemFree (pItemIDList)
    End If
    
    BrowseForFolder = bRes
End Function

Public Function dirParse(sPath As String, ByVal method As enumMethod) As String
    Dim strFileName As String
    Dim iSep As Integer
    Dim nDot As Integer
    
    strFileName = sPath
    Do
        iSep = InStr(strFileName, "\")
        If iSep = 0 Then iSep = InStr(strFileName, ":")
          If iSep = 0 Then
            Select Case method
             Case 1 ' path
              dirParse = Left(sPath, Len(sPath) - Len(strFileName))
             Case 2 ' filename w\extension
              dirParse = strFileName
             Case 3 ' extension
               nDot = InStr(strFileName, ".")
                If nDot Then
                 dirParse = Mid$(strFileName, nDot)
                End If
             Case 4 ' basefilename without "."
                nDot = InStr(strFileName, ".")
                 If nDot Then
                  dirParse = Left$(strFileName, nDot - 1)
                 Else
                  dirParse = strFileName
                 End If
             Case 5 ' baseExtension without "."
               nDot = InStr(strFileName, ".")
                If nDot Then
                 dirParse = Mid$(strFileName, nDot + 1)
                End If
             End Select
           Exit Function
        Else
            strFileName = Right(strFileName, Len(strFileName) - iSep)
        End If
    Loop
End Function

Public Sub Main()
If App.PrevInstance Then Exit Sub
gCurrentDir = App.Path
If Right(gCurrentDir, 1) <> "\" Then gCurrentDir = gCurrentDir & "\"
If Not DirExists(gCurrentDir & "Build") Then
 On Error Resume Next
 MkDir gCurrentDir & "Build"
End If
gCurrentDir = gCurrentDir & "Build" & "\"
frmMain.Show
End Sub
Public Function DirExists(ByVal sDirName As String) As Boolean
    Const sWILDCARD$ = "*.*"
    Dim strDummy As String
    On Error Resume Next
    If Right$(sDirName, 1) <> "\" Then
        sDirName = sDirName & "\"
    End If
    If Len(sDirName) < 3 Then
     DirExists = False
     Exit Function
    Else
    strDummy = Dir$(sDirName & sWILDCARD, vbDirectory)
    DirExists = Not (strDummy = "")
    Err = 0
    End If
End Function

Sub SaveFileAs(Filename As String, sContent As String)
    On Error Resume Next

    ' Open the file.
    Open gCurrentDir & Filename For Output As #1
    Screen.MousePointer = 11
    Print #1, sContent
    Close #1
    Screen.MousePointer = 0
    ' Set the form's caption.
    If Err Then
        MsgBox Error, 48, App.Title
    End If
End Sub
