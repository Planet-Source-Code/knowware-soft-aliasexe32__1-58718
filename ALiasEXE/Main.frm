VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AliasEXE32"
   ClientHeight    =   4020
   ClientLeft      =   6075
   ClientTop       =   3990
   ClientWidth     =   6945
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picSetup 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3255
      ScaleWidth      =   6735
      TabIndex        =   4
      Top             =   0
      Width           =   6735
      Begin VB.Frame fraSetup 
         Height          =   1455
         Index           =   4
         Left            =   2280
         TabIndex        =   20
         Top             =   0
         Width           =   2415
         Begin VB.TextBox txtSummary 
            BackColor       =   &H00800000&
            ForeColor       =   &H00FFFFFF&
            Height          =   1935
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   25
            Top             =   1080
            Width           =   4095
         End
         Begin VB.Label lblStatus 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1560
            TabIndex        =   26
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label lblCaption 
            BackStyle       =   0  'Transparent
            Caption         =   $"Main.frx":030A
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   7
            Left            =   120
            TabIndex        =   21
            Top             =   80
            Width           =   3975
         End
      End
      Begin VB.Frame fraSetup 
         Height          =   735
         Index           =   3
         Left            =   4080
         TabIndex        =   16
         Top             =   1920
         Width           =   975
         Begin VB.TextBox txtDestinationPath 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   1680
            Width           =   4095
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "Browse..."
            Height          =   375
            Index           =   2
            Left            =   3000
            TabIndex        =   17
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblCaption 
            Caption         =   "Destination .exe Path:"
            Height          =   225
            Index           =   4
            Left            =   120
            TabIndex        =   19
            Top             =   1440
            Width           =   2055
         End
         Begin VB.Label lblCaption 
            BackStyle       =   0  'Transparent
            Caption         =   " Please choose 'Browse' to locate the destination folder where the actual program executable will reside."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   5
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   3975
         End
      End
      Begin VB.Frame fraSetup 
         Height          =   1215
         Index           =   2
         Left            =   5280
         TabIndex        =   12
         Top             =   360
         Width           =   975
         Begin VB.TextBox txtProjectPath 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   1680
            Width           =   4095
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "Browse..."
            Height          =   375
            Index           =   1
            Left            =   3000
            TabIndex        =   13
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblCaption 
            Caption         =   ".VBP Project Path:"
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   15
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label lblCaption 
            BackStyle       =   0  'Transparent
            Caption         =   " Please choose 'Browse' to locate the project file (.vbp) that you wish to create the AliasEXE executable for."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   3
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   3975
         End
      End
      Begin VB.Frame fraSetup 
         Height          =   855
         Index           =   1
         Left            =   5280
         TabIndex        =   8
         Top             =   1800
         Width           =   1095
         Begin VB.TextBox txtVB6Path 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   1680
            Width           =   4095
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "Browse..."
            Height          =   375
            Index           =   0
            Left            =   3000
            TabIndex        =   9
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblCaption 
            BackStyle       =   0  'Transparent
            Caption         =   " It is necessary to know where your VB6.EXE file is located. Please choose 'Browse' to locate the executable."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   3975
         End
         Begin VB.Label lblCaption 
            Caption         =   "VB6.exe Path:"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   1440
            Width           =   1215
         End
      End
      Begin VB.Frame fraSetup 
         Height          =   855
         Index           =   0
         Left            =   2880
         TabIndex        =   7
         Top             =   1800
         Width           =   975
         Begin VB.TextBox txtIntro 
            BackColor       =   &H00800000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   2775
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   27
            Text            =   "Main.frx":0392
            Top             =   240
            Width           =   4215
         End
      End
      Begin VB.Frame fraImage 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3025
         Left            =   80
         TabIndex        =   5
         Top             =   80
         Width           =   2100
         Begin VB.Label lblCaption 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "AliasEXE32 Wizard"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C0C0&
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   6
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Image Image1 
            Height          =   1080
            Left            =   360
            Picture         =   "Main.frx":0398
            Stretch         =   -1  'True
            Top             =   480
            Width           =   1155
         End
      End
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Hidden          =   -1  'True
      Left            =   0
      Pattern         =   "*.frm"
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next ->"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<- &Back"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Left            =   -120
      TabIndex        =   28
      Top             =   3300
      Width           =   7095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private locTop As Integer
Private locLeft As Integer
Private locWidth As Integer
Private locHeight As Integer
Private iCount As Integer
Const MaxCount = 4

Private sVB6exe As String
Private sProjectName As String
Private sProjectSetupFolder As String
Private sDestinationPath As String


Private sExeName32 As String
Private sLocalExeName As String
Private sVersionCompanyName As String
Private sIconForm As String
Private sFrx As String
Private sProjectSetupPath As String

Private Declare Function WaitForSingleObject Lib "kernel32" _
   (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long
   
Private Declare Function OpenProcess Lib "kernel32" _
   (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long


Private Const INFINITE = -1&
Private Const SYNCHRONIZE = &H100000

Private Function CheckIfOpen(sVBP As String) As Boolean
Dim fNum As Integer
Dim KeyName As String
Dim LineToRead As String
Dim pos As Integer

    fNum = FreeFile()
    Open sVBP For Input As fNum
    
    Do While Not EOF(fNum)
        ' Read the next line from the file.
        Line Input #fNum, LineToRead
        LineToRead = Trim$(LineToRead)
            If InStr(LineToRead, "VersionCompanyName=") Then
            ElseIf InStr(LineToRead, "Name=") Then
                pos = InStr(LineToRead, "=")
                KeyName = Trim$(Mid$(LineToRead, pos + 1, Len(LineToRead) - pos))
                gsAppName = ParseString(KeyName, Chr$(34))
            End If
    Loop
    Close fNum
Call fEnumWindows
If lnghWnd > 0 Then
 CheckIfOpen = True
Else
 CheckIfOpen = False
End If
End Function

Private Sub CopyComplete()
Dim i As Integer
If Not DirExists(sProjectSetupFolder) Then
  Call CreateNewDirectory(sProjectSetupFolder)
End If
 If Right(sProjectSetupFolder, 1) <> "\" Then sProjectSetupFolder = sProjectSetupFolder & "\"
 If Right(gCurrentDir, 1) <> "\" Then gCurrentDir = gCurrentDir & "\"
 On Error Resume Next
 FileCopy gCurrentDir & sExeName32, sProjectSetupFolder & sExeName32
MsgBox "Operation Completed Successfully!" & vbNewLine & vbNewLine & sExeName32 & "  Has been successfully copied to:" & vbNewLine & sProjectSetupFolder, vbInformation
Call ResetProgram

End Sub
Private Function CreateForm() As String
Dim sForm1 As String
Screen.MousePointer = 11
sForm1 = "VERSION 5.00" & vbNewLine
sForm1 = sForm1 & "Begin VB.Form AliasEXEForm" & vbNewLine
sForm1 = sForm1 & "   Caption = """ & "AliasEXEForm" & """" & vbNewLine
sForm1 = sForm1 & "   ClientHeight = 8235" & vbNewLine
sForm1 = sForm1 & "   ClientLeft = 1650" & vbNewLine
sForm1 = sForm1 & "   ClientTop = 1545" & vbNewLine
sForm1 = sForm1 & "   ClientWidth = 6585" & vbNewLine
If Not sFrx = vbNullString Then
sForm1 = sForm1 & "   Icon  =  """ & sFrx & """" & ":0000" & vbNewLine
Else
sForm1 = sForm1 & "   Icon  =  """ & """" & vbNewLine
End If
sForm1 = sForm1 & "   LinkTopic = """ & "AliasEXEForm" & """" & vbNewLine
sForm1 = sForm1 & "   ScaleHeight = 8235" & vbNewLine
sForm1 = sForm1 & "   ScaleWidth = 6585" & vbNewLine
sForm1 = sForm1 & "End" & vbNewLine
sForm1 = sForm1 & "Attribute VB_Name = """ & "AliasEXEForm" & """" & vbNewLine
sForm1 = sForm1 & "Attribute VB_GlobalNameSpace = False" & vbNewLine
sForm1 = sForm1 & "Attribute VB_Creatable = False" & vbNewLine
sForm1 = sForm1 & "Attribute VB_PredeclaredId = True" & vbNewLine
sForm1 = sForm1 & "Attribute VB_Exposed = False" & vbNewLine
sForm1 = sForm1 & "Option Explicit"
CreateForm = sForm1

End Function

Private Sub CreateIntro()
Dim sLine As String
sLine = sLine & "  This wizard is designed to help automate " & vbNewLine
sLine = sLine & "the creation of the AliasEXE file which," & vbNewLine
sLine = sLine & "when executed, will compare the destination" & vbNewLine
sLine = sLine & "executable to the local executable file." & vbNewLine
sLine = sLine & "If a newer destination executable is found the " & vbNewLine
sLine = sLine & "local executable will then be replaced by the destination executable file." & vbNewLine & vbNewLine
sLine = sLine & " Information will need to be collected in order" & vbNewLine
sLine = sLine & "to complete the wizard. press the " & vbNewLine
sLine = sLine & "'Next ->' button to proceed."
txtIntro.Text = sLine
End Sub

Private Sub CreateProjectFiles()
Call SaveFileAs("AliasEXE32.vbp", CreateVBPFile)
Call SaveFileAs("AliasEXEForm.Frm", CreateForm)
Call SaveFileAs("AliasEXE32.Bas", CreateModule)
End Sub


Private Sub CreateSummary()
Dim sLine As String
sLine = vbNewLine
sLine = sLine & "VB6.exe Path:" & "  " & sVB6exe & vbNewLine & vbNewLine
sLine = sLine & "Project File:" & "  " & dirParse(sProjectName, eBaseFileName) & ".vbp" & vbNewLine & vbNewLine
sLine = sLine & "Destination Path:" & "  " & sDestinationPath & vbNewLine & vbNewLine
sLine = sLine & "AliasEXE exe will be copied to:" & vbNewLine
sLine = sLine & "  " & sProjectSetupFolder
txtSummary.Text = sLine
End Sub
Private Function GetIcon(ByVal sPath As String, frm As String) As String
Dim i As Integer
Dim fNum As Integer, bFound As Boolean
Dim sForm As String, sFile As String
Dim LineToRead As String, iCount As Integer
 bFound = False
 If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
 For i = 0 To File1.ListCount - 1
    fNum = FreeFile()
    sForm = sPath & File1.List(i)
    sFile = dirParse(sForm, eBaseFileName)
    If InStr(LCase(frm), LCase(sFile)) > 0 Then  '= LCase(sPath & frm & ".frm") Then
         bFound = True
         Exit For
     End If
   Close fNum
 Next
If bFound = True Then
 GetIcon = sFile & ".frx"
Else
 GetIcon = vbNullString
End If
End Function

Private Sub GetRegSettings()
'Dim sysVar As Variant, strDirSource As String
    locLeft = (Screen.Width - 6800) / 2
    locTop = (Screen.Height - 4400) / 3.5
    locWidth = 6800
    locHeight = 4400
    gsVBPath = GetSetting(appName, "Settings", "VBPath", "C:\")
    sVB6exe = GetSetting(appName, "Settings", "VBEXE", "")
    If Not DirExists(gsVBPath) Then gsVBPath = "C:\"
    If Not sVB6exe = vbNullString Then
     If Not FileExists(sVB6exe) Then sVB6exe = vbNullString
    End If
   If Len(sVB6exe) > 0 Then txtVB6Path.Text = sVB6exe
End Sub

Private Sub KillVbApp()
Dim ret As Long
Screen.MousePointer = vbHourglass
ret = SendMessage(lnghWnd, WM_CLOSE, 0, 0)
Do
 DoEvents
Loop Until ret = 0
Screen.MousePointer = vbDefault

End Sub

Private Sub NextFrame(ByVal idx As Integer)
Dim i As Integer
 For i = 0 To fraSetup.Count - 1
  fraSetup(i).Visible = False
 Next
 fraSetup(idx).Visible = True
End Sub

Private Sub ReadVBPFile(sVBP As String)
Dim fNum As Integer, sFile As String
Dim KeyName As String
Dim LineToRead As String
Dim pos As Integer

    fNum = FreeFile()
    Open sVBP For Input As fNum
    
    Do While Not EOF(fNum)
        ' Read the next line from the file.
        Line Input #fNum, LineToRead
        LineToRead = Trim$(LineToRead)
            If InStr(LineToRead, "ExeName32=") Then
                pos = InStr(LineToRead, "=")
                KeyName = Trim$(Mid$(LineToRead, pos + 1, Len(LineToRead) - pos))
                sFile = ParseString(KeyName, Chr$(34))
                sExeName32 = dirParse(sFile, eBaseFileName) & "_Run.exe"
                sLocalExeName = sFile
            ElseIf InStr(LineToRead, "VersionCompanyName=") Then
                pos = InStr(LineToRead, "=")
                KeyName = Trim$(Mid$(LineToRead, pos + 1, Len(LineToRead) - pos))
                sVersionCompanyName = ParseString(KeyName, Chr$(34))
            ElseIf InStr(LineToRead, "Name=") Then
                pos = InStr(LineToRead, "=")
                KeyName = Trim$(Mid$(LineToRead, pos + 1, Len(LineToRead) - pos))
                gsAppName = ParseString(KeyName, Chr$(34))
            ElseIf InStr(LineToRead, "IconForm=") Then
                pos = InStr(LineToRead, "=")
                KeyName = Trim$(Mid$(LineToRead, pos + 1, Len(LineToRead) - pos))
                sIconForm = ParseString(KeyName, Chr$(34))
            End If
    Loop
    Close fNum

End Sub

Private Sub ResetProgram()
Dim i As Integer
For i = 0 To fraSetup.Count - 1
 fraSetup(i).Visible = False
Next
fraSetup(0).Visible = True
iCount = 0
cmdBack.Enabled = False
cmdNext.Enabled = True
cmdNext.Caption = "&Next ->"
gsAppName = vbNullString
sProjectName = vbNullString
txtProjectPath.Text = vbNullString
sProjectSetupFolder = vbNullString
txtSummary.Text = vbNullString
sDestinationPath = vbNullString
txtDestinationPath.Text = vbNullString
Caption = "AliasEXE32 Wizard"
lblStatus.Caption = vbNullString
lblStatus.Refresh
lnghWnd = 0
End Sub

 Private Sub SaveRegSettings()
        SaveSetting appName, "Settings", "VBPath", gsVBPath
        SaveSetting appName, "Settings", "VBEXE", sVB6exe

 End Sub

Private Sub setForm()
Dim i As Integer
Call CreateIntro
For i = 0 To fraSetup.Count - 1
 fraSetup(i).Move 2260, 0, 4335, 3135
 fraSetup(i).Visible = False
 fraSetup(i).BorderStyle = 0
Next
fraSetup(0).ZOrder
fraSetup(0).Visible = True
iCount = 0
cmdBack.Enabled = False
txtIntro.Move 0, 80, 4335, 3050
End Sub

Private Sub cmdBack_Click()
iCount = iCount - 1
Call NextFrame(iCount)
If iCount = 0 Then cmdBack.Enabled = False
cmdNext.Caption = "&Next ->"
cmdNext.Enabled = True
End Sub

Private Sub cmdBrowse_Click(index As Integer)

Select Case index
Case 0 'Browse for vb6.exe
    Dim sFile As String
    sFile = OpenFileName(Me, 2)
    If sFile <> vbNullString Then
      sVB6exe = sFile
      txtVB6Path.Text = sVB6exe
      gsVBPath = dirParse(sFile, eFullPath)
    End If
   
Case 1 'browse for vbp project file
    Dim bOpen As Boolean, bKillApp As Boolean
    sProjectName = OpenFileName(Me, 1)
    DoEvents
    lnghWnd = 0
    If Len(sProjectName) > 0 Then
     Caption = Caption & "  Project [" & dirParse(sProjectName, eFilename) & "]"
     bOpen = CheckIfOpen(sProjectName)
       If bOpen Then
        bKillApp = MsgBox(gsAppName & "  Project is currently open." & vbNewLine & "Project will have to be terminated before compiling." & vbNewLine & "Do you wish to terminate " & gsAppName & " now?", vbExclamation + vbDefaultButton1 + vbYesNoCancel, "Terminate App") = vbYes
         If bKillApp = True Then
            picSetup.Enabled = False
            cmdNext.Enabled = False
            cmdCancel.Enabled = False
            cmdBack.Enabled = False
           Call KillVbApp
            cmdNext.Enabled = True
            cmdCancel.Enabled = True
            cmdBack.Enabled = True
            picSetup.Enabled = True
         Else
          Exit Sub
         End If
       End If
     txtProjectPath.Text = sProjectName
     sProjectSetupFolder = dirParse(sProjectName, eFullPath) & "Setup\"
    End If
Case 2 'browse for destination network folder
    Dim PathName As String
    If BrowseForFolder(frmMain, PathName, "Please select the Folder where the project .exe file will reside.") Then
      If Right(PathName, 1) <> "\" Then PathName = PathName & "\"
      sDestinationPath = PathName
      txtDestinationPath.Text = sDestinationPath
    End If
Case 3 'browse for destination folder
    Dim sTarget As String
    If BrowseForFolder(frmMain, sTarget, "Please select the Folder where the AliasEXE32.exe file will be copied to.") Then
      If Right(sTarget, 1) <> "\" Then sTarget = sTarget & "\"
      sProjectSetupFolder = sTarget
    End If
Case Else
End Select

End Sub

Private Function CreateModule() As String
Dim sMod As String, TargetName As String, sExeName As String
Screen.MousePointer = 11
TargetName = sDestinationPath & sLocalExeName
'Globals
sMod = "Attribute VB_Name = """ & "AliasEXE32" & """" & vbNewLine
sMod = sMod & "Option Explicit" & vbNewLine
sMod = sMod & "Private gCurrentDir As String" & vbNewLine
sMod = sMod & "'FileCopy" & vbNewLine
sMod = sMod & "Private Declare Function SHFileOperation Lib """ & "Shell32.dll" & """ Alias """ & "SHFileOperationA" & """" & "(lpFileOp As SHFILEOPSTRUCT) As Long" & vbNewLine
sMod = sMod & "Private Type SHFILEOPSTRUCT" & vbNewLine
sMod = sMod & "    hwnd As Long" & vbNewLine
sMod = sMod & "    wFunc As Long" & vbNewLine
sMod = sMod & "    pFrom As String" & vbNewLine
sMod = sMod & "    pTo As String" & vbNewLine
sMod = sMod & "    fFlags As Long" & vbNewLine
sMod = sMod & "    fAnyOperationsAborted As Boolean" & vbNewLine
sMod = sMod & "    hNameMappings As Long" & vbNewLine
sMod = sMod & "    lpszProgressTitle As String 'Used only if FOF_SIMPLEPROGRESS specified" & vbNewLine
sMod = sMod & "End Type" & vbNewLine
sMod = sMod & "Private Const FO_COPY = 2" & vbNewLine
sMod = sMod & "Private Const FOF_SILENT = &H4              'Don't display progress dialog" & vbNewLine
sMod = sMod & "Private Const FOF_NOCONFIRMATION = &H10     'Don't display Confirmation" & vbNewLine
sMod = sMod & "'Check for Previous Instance***********************************************************************" & vbNewLine
sMod = sMod & "Private Const MAX_PATH = 260" & vbNewLine
sMod = sMod & "Private Type PROCESSENTRY32" & vbNewLine
sMod = sMod & "    dwSize As Long" & vbNewLine
sMod = sMod & "    cntUsage As Long" & vbNewLine
sMod = sMod & "    th32ProcessID As Long" & vbNewLine
sMod = sMod & "    th32DefaultHeapID As Long" & vbNewLine
sMod = sMod & "    th32ModuleID As Long" & vbNewLine
sMod = sMod & "    cntThreads As Long" & vbNewLine
sMod = sMod & "    th32ParentProcessID As Long" & vbNewLine
sMod = sMod & "    pcPriClassBase As Long" & vbNewLine
sMod = sMod & "    dwFlags As Long" & vbNewLine
sMod = sMod & "    szExeFile As String * MAX_PATH" & vbNewLine
sMod = sMod & "End Type" & vbNewLine
sMod = sMod & "Private Type OSVERSIONINFO" & vbNewLine
sMod = sMod & "   dwOSVersionInfoSize As Long" & vbNewLine
sMod = sMod & "   dwMajorVersion As Long" & vbNewLine
sMod = sMod & "   dwMinorVersion As Long" & vbNewLine
sMod = sMod & "   dwBuildNumber As Long" & vbNewLine
sMod = sMod & "   dwPlatformId As Long" & vbNewLine
sMod = sMod & "   szCSDVersion As String * 128" & vbNewLine
sMod = sMod & "End Type" & vbNewLine
sMod = sMod & vbNewLine
sMod = sMod & "Private Declare Function CreateToolhelp32Snapshot Lib """ & "kernel32" & """" & " (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long" & vbNewLine
sMod = sMod & "Private Declare Function Process32First Lib """ & "kernel32" & """" & " (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long" & vbNewLine
sMod = sMod & "Private Declare Function Process32Next Lib """ & "kernel32" & """" & " (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long" & vbNewLine
sMod = sMod & "Private Declare Function CloseHandle Lib """ & "kernel32" & """" & " (ByVal hObject As Long) As Long" & vbNewLine
sMod = sMod & "Private Declare Function OpenProcess Lib """ & "kernel32" & """" & " (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long" & vbNewLine
sMod = sMod & "Private Declare Function EnumProcesses Lib """ & "psapi.dll" & """" & " (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long" & vbNewLine
sMod = sMod & "Private Declare Function GetModuleFileNameExA Lib """ & "psapi.dll" & """" & " (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long" & vbNewLine
sMod = sMod & "Private Declare Function EnumProcessModules Lib """ & "psapi.dll" & """" & " (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long" & vbNewLine
sMod = sMod & "Private Declare Function GetVersionExA Lib """ & "kernel32" & """" & " (lpVersionInformation As OSVERSIONINFO) As Integer" & vbNewLine
sMod = sMod & "Private Const PROCESS_QUERY_INFORMATION = 1024" & vbNewLine
sMod = sMod & "Private Const PROCESS_VM_READ = 16" & vbNewLine
sMod = sMod & "Private Const STANDARD_RIGHTS_REQUIRED = &HF0000" & vbNewLine
sMod = sMod & "Private Const SYNCHRONIZE = &H100000" & vbNewLine
sMod = sMod & "Private Const PROCESS_ALL_ACCESS = &H1F0FFF" & vbNewLine
sMod = sMod & "Private Const TH32CS_SNAPPROCESS = &H2" & vbNewLine
sMod = sMod & "Private Const hNull = 0" & vbNewLine
'IsPreviousInstance
sMod = sMod & "Function IsPreviousInstance(ByVal sFileName As String) As Boolean" & vbNewLine
sMod = sMod & "  Dim bInstance As Boolean" & vbNewLine
sMod = sMod & "      bInstance = False" & vbNewLine
sMod = sMod & "      On Error Resume Next" & vbNewLine
sMod = sMod & " Select Case getVersion()" & vbNewLine
sMod = sMod & "   Case 1 'Windows 95/98" & vbNewLine
sMod = sMod & "        Dim hSnapshot As Long, ret As Long, sFile As String, sAppName As String, P As PROCESSENTRY32" & vbNewLine
sMod = sMod & "        P.dwSize = Len(P)" & vbNewLine
sMod = sMod & "        hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, ByVal 0)" & vbNewLine
sMod = sMod & "        If hSnapshot Then" & vbNewLine
sMod = sMod & "            ret = Process32First(hSnapshot, P)" & vbNewLine
sMod = sMod & "            Do While ret" & vbNewLine
sMod = sMod & "                sFile = dirParse(Left$(P.szExeFile, InStr(P.szExeFile, Chr$(0)) - 1))" & vbNewLine
sMod = sMod & "                If UCase(sFile) = UCase(sFileName) Then" & vbNewLine
sMod = sMod & "                 bInstance = True" & vbNewLine
sMod = sMod & "                 Exit Do" & vbNewLine
sMod = sMod & "                End If" & vbNewLine
sMod = sMod & "                ret = Process32Next(hSnapshot, P)" & vbNewLine
sMod = sMod & "            Loop" & vbNewLine
sMod = sMod & "            ret = CloseHandle(hSnapshot)" & vbNewLine
sMod = sMod & "        End If" & vbNewLine
sMod = sMod & "   Case 2 'Windows NT" & vbNewLine
sMod = sMod & "         Dim cb As Long, cbNeeded As Long, NumElements As Long, ProcessIDs() As Long, cbNeeded2 As Long" & vbNewLine
sMod = sMod & "         Dim NumElements2 As Long, Modules(1 To 200) As Long, lRet As Long, ModuleName As String, nSize As Long" & vbNewLine
sMod = sMod & "         Dim hProcess As Long, i As Long, sReturn As String" & vbNewLine
sMod = sMod & "         cb = 8" & vbNewLine
sMod = sMod & "         cbNeeded = 96" & vbNewLine
sMod = sMod & "         Do While cb <= cbNeeded" & vbNewLine
sMod = sMod & "            cb = cb * 2" & vbNewLine
sMod = sMod & "            ReDim ProcessIDs(cb / 4) As Long" & vbNewLine
sMod = sMod & "            lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)" & vbNewLine
sMod = sMod & "         Loop" & vbNewLine
sMod = sMod & "         NumElements = cbNeeded / 4" & vbNewLine
sMod = sMod & "         For i = 1 To NumElements" & vbNewLine
sMod = sMod & "            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessIDs(i))" & vbNewLine
sMod = sMod & "            If hProcess <> 0 Then" & vbNewLine
sMod = sMod & "                lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)" & vbNewLine
sMod = sMod & "               If lRet <> 0 Then" & vbNewLine
sMod = sMod & "                   ModuleName = Space(MAX_PATH)" & vbNewLine
sMod = sMod & "                   nSize = 500" & vbNewLine
sMod = sMod & "                   lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)" & vbNewLine
sMod = sMod & "                   sReturn = dirParse(Left(ModuleName, lRet))" & vbNewLine
sMod = sMod & "                   If UCase(sReturn) = UCase(sFileName) Then" & vbNewLine
sMod = sMod & "                    bInstance = True" & vbNewLine
sMod = sMod & "                    Exit For" & vbNewLine
sMod = sMod & "                   End If" & vbNewLine
sMod = sMod & "                End If" & vbNewLine
sMod = sMod & "            End If" & vbNewLine
sMod = sMod & "         lRet = CloseHandle(hProcess)" & vbNewLine
sMod = sMod & "         Next" & vbNewLine
sMod = sMod & " End Select" & vbNewLine
sMod = sMod & "IsPreviousInstance = bInstance" & vbNewLine
sMod = sMod & "End Function" & vbNewLine
'StrZToStr
sMod = sMod & "Function StrZToStr(s As String) As String" & vbNewLine
sMod = sMod & "   StrZToStr = Left$(s, Len(s) - 1)" & vbNewLine
sMod = sMod & "End Function" & vbNewLine
'Get Version
sMod = sMod & "Private Function getVersion() As Long" & vbNewLine
sMod = sMod & "   Dim osinfo As OSVERSIONINFO" & vbNewLine
sMod = sMod & "   Dim retvalue As Integer" & vbNewLine
sMod = sMod & "   osinfo.dwOSVersionInfoSize = 148" & vbNewLine
sMod = sMod & "   osinfo.szCSDVersion = Space$(128)" & vbNewLine
sMod = sMod & "   retvalue = GetVersionExA(osinfo)" & vbNewLine
sMod = sMod & "   getVersion = osinfo.dwPlatformId" & vbNewLine
sMod = sMod & "End Function" & vbNewLine
'FileExists
sMod = sMod & "Function FileExists(sSpec As String) As Integer" & vbNewLine
sMod = sMod & "    Dim fileFile As Integer" & vbNewLine
sMod = sMod & "    'attempt to open file" & vbNewLine
sMod = sMod & "    fileFile = FreeFile" & vbNewLine
sMod = sMod & "    On Error Resume Next" & vbNewLine
sMod = sMod & "    Open sSpec For Input As fileFile" & vbNewLine
sMod = sMod & "    'check for error" & vbNewLine
sMod = sMod & "    If Err Then" & vbNewLine
sMod = sMod & "        FileExists = False" & vbNewLine
sMod = sMod & "    Else" & vbNewLine
sMod = sMod & "        'file exists" & vbNewLine
sMod = sMod & "        'close file" & vbNewLine
sMod = sMod & "        Close fileFile" & vbNewLine
sMod = sMod & "        FileExists = True" & vbNewLine
sMod = sMod & "    End If" & vbNewLine
sMod = sMod & "End Function" & vbNewLine
'Copyfiles
sMod = sMod & "Private Sub CopyFiles(ByVal sTarget As String, ByVal sLocal As String)" & vbNewLine
sMod = sMod & "Dim FileOp As SHFILEOPSTRUCT, ret As Long" & vbNewLine
sMod = sMod & "     Screen.Mousepointer = vbDefault" & vbNewLine
sMod = sMod & "     FileOp.fFlags = FOF_NOCONFIRMATION Or FOF_SILENT" & vbNewLine
sMod = sMod & "     FileOp.wFunc = FO_COPY" & vbNewLine
sMod = sMod & "     FileOp.pFrom = sTarget & Chr$(0)" & vbNewLine
sMod = sMod & "     FileOp.pTo = sLocal & Chr$(0)" & vbNewLine
sMod = sMod & "     ret = SHFileOperation(FileOp) <> 0" & vbNewLine
sMod = sMod & "     FileOp.fFlags = 0" & vbNewLine
sMod = sMod & "End Sub" & vbNewLine
'DirParse
sMod = sMod & "Private Function dirParse(sPath As String) As String" & vbNewLine
sMod = sMod & "    Dim strFileName As String" & vbNewLine
sMod = sMod & "    Dim iSep As Integer" & vbNewLine
sMod = sMod & vbNewLine
sMod = sMod & "    strFileName = sPath" & vbNewLine
sMod = sMod & "    Do" & vbNewLine
sMod = sMod & "        iSep = InStr(strFileName, """ & "\""" & ")" & vbNewLine
sMod = sMod & "        If iSep = 0 Then iSep = InStr(strFileName, """ & ":""" & ")" & vbNewLine
sMod = sMod & "          If iSep = 0 Then" & vbNewLine
sMod = sMod & "              dirParse = strFileName" & vbNewLine
sMod = sMod & "           Exit Function" & vbNewLine
sMod = sMod & "        Else" & vbNewLine
sMod = sMod & "            strFileName = Right(strFileName, Len(strFileName) - iSep)" & vbNewLine
sMod = sMod & "        End If" & vbNewLine
sMod = sMod & "    Loop" & vbNewLine
sMod = sMod & "End Function" & vbNewLine
'LoadExe
sMod = sMod & "Private Sub LoadExe(sPath As String)" & vbNewLine
sMod = sMod & "    Dim ret As Long" & vbNewLine
sMod = sMod & "    Screen.Mousepointer = vbDefault" & vbNewLine
sMod = sMod & "    ret = Shell(sPath, vbNormalFocus)" & vbNewLine
sMod = sMod & "    End" & vbNewLine
sMod = sMod & "End Sub" & vbNewLine
'SubMain
sMod = sMod & "Public Sub Main()" & vbNewLine
sMod = sMod & "Dim sAppName As String, bIsRunning As Boolean" & vbNewLine
sMod = sMod & "sAppName = """ & sLocalExeName & """" & vbNewLine
sMod = sMod & "bIsRunning = IsPreviousInstance(sAppName)" & vbNewLine
sMod = sMod & "If bIsRunning = True Then" & vbNewLine
sMod = sMod & "  MsgBox sAppName & """ & "  Is Already Running...""" & ",""" & "48" & vbNewLine
sMod = sMod & "  End" & vbNewLine
sMod = sMod & "  Exit Sub" & vbNewLine
sMod = sMod & "End If" & vbNewLine
sMod = sMod & vbNewLine
'        AppActivate SaveTitle$
'        SendKeys "% R", True
sMod = sMod & "Dim LocalExe As String" & vbNewLine
sMod = sMod & "Dim TargetExe As String" & vbNewLine
sMod = sMod & "Dim SourceFileVar As Double" & vbNewLine
sMod = sMod & "Dim TargetFileVar As Double" & vbNewLine
sMod = sMod & "Screen.Mousepointer = vbDefault" & vbNewLine
sMod = sMod & "gCurrentDir = App.Path" & vbNewLine
sMod = sMod & "If Right(gCurrentDir, 1) <> """ & "\" & """" & " Then gCurrentDir = gCurrentDir & """ & "\" & """" & vbNewLine
sMod = sMod & "Localexe = gCurrentDir & """ & sLocalExeName & """" & vbNewLine
sMod = sMod & "TargetExe = """ & TargetName & """" & vbNewLine
sMod = sMod & "If FileExists(LocalExe) Then" & vbNewLine
sMod = sMod & "  If FileExists(TargetExe) Then" & vbNewLine
sMod = sMod & "    SourceFileVar = Format(FileDateTime(LocalExe), """ & "yyyymmddhhmmss" & """" & ")" & vbNewLine
sMod = sMod & "    TargetFileVar = Format(FileDateTime(TargetExe), """ & "yyyymmddhhmmss" & """" & ")" & vbNewLine
sMod = sMod & "    If SourceFileVar >= TargetFileVar Then" & vbNewLine
sMod = sMod & "     Call LoadExe(LocalExe)" & vbNewLine
sMod = sMod & "    Else" & vbNewLine
sMod = sMod & "     Call CopyFiles(TargetExe, LocalExe)" & vbNewLine
sMod = sMod & "     Call LoadExe(LocalExe)" & vbNewLine
sMod = sMod & "    End If" & vbNewLine
sMod = sMod & "  Else" & vbNewLine
sMod = sMod & "   Call LoadExe(LocalExe)" & vbNewLine
sMod = sMod & "  End If" & vbNewLine
sMod = sMod & "Else" & vbNewLine
sMod = sMod & "  Call CopyFiles(TargetExe, LocalExe)" & vbNewLine
sMod = sMod & "  Call LoadExe(LocalExe)" & vbNewLine
sMod = sMod & "End If" & vbNewLine
sMod = sMod & "End Sub" & vbNewLine

CreateModule = sMod
End Function

Private Function CreateVBPFile() As String
Dim sVBP As String
Screen.MousePointer = 11
sVBP = "Type=Exe" & vbNewLine
sVBP = sVBP & "Form =AliasEXEForm.frm" & vbNewLine
'sVBP = sVBP & "Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\WINDOWS\SYSTEM\stdole2.tlb#OLE Automation" & vbNewLine
sVBP = sVBP & "Module=AliasEXE32; AliasEXE32.bas" & vbNewLine
sVBP = sVBP & "IconForm=""" & sIconForm & """" & vbNewLine
sVBP = sVBP & "Startup=""" & "Sub Main" & """" & vbNewLine
sVBP = sVBP & "HelpFile=""" & """" & vbNewLine
sVBP = sVBP & "ExeName32 =""" & sExeName32 & """" & vbNewLine
sVBP = sVBP & "Path32=""" & """" & vbNewLine
sVBP = sVBP & "Command32=""" & """" & vbNewLine
sVBP = sVBP & "Name=""" & gsAppName & """" & vbNewLine
sVBP = sVBP & "HelpContextID= """ & "0" & """" & vbNewLine
sVBP = sVBP & "CompatibleMode=""" & "0" & """" & vbNewLine
sVBP = sVBP & "MajorVer=1" & vbNewLine
sVBP = sVBP & "MinorVer=0" & vbNewLine
sVBP = sVBP & "RevisionVer=0" & vbNewLine
sVBP = sVBP & "AutoIncrementVer=0" & vbNewLine
sVBP = sVBP & "ServerSupportFiles=0" & vbNewLine
sVBP = sVBP & "VersionCompanyName=""" & sVersionCompanyName & """" & vbNewLine
sVBP = sVBP & "CompilationType=0" & vbNewLine
sVBP = sVBP & "OptimizationType=0" & vbNewLine
sVBP = sVBP & "FavorPentiumPro(tm)=0" & vbNewLine
sVBP = sVBP & "CodeViewDebugInfo=0" & vbNewLine
sVBP = sVBP & "NoAliasing=0" & vbNewLine
sVBP = sVBP & "BoundsCheck=0" & vbNewLine
sVBP = sVBP & "OverflowCheck=0" & vbNewLine
sVBP = sVBP & "FlPointCheck=0" & vbNewLine
sVBP = sVBP & "FDIVCheck=0" & vbNewLine
sVBP = sVBP & "UnroundedFP=0" & vbNewLine
sVBP = sVBP & "StartMode=0" & vbNewLine
sVBP = sVBP & "Unattended=0" & vbNewLine
sVBP = sVBP & "Retained=0" & vbNewLine
sVBP = sVBP & "ThreadPerObject=0" & vbNewLine
sVBP = sVBP & "MaxNumberOfThreads=1"
CreateVBPFile = sVBP
End Function


Private Sub LoadProject()
Dim sFile As String, sPath As String
Screen.MousePointer = 11
If Len(sProjectName) > 0 Then
   Call ReadVBPFile(sProjectName)
    If Not sIconForm = vbNullString Then
        File1.Path = dirParse(sProjectName, eFullPath)
        File1.Pattern = "*.frm"
        sPath = File1.Path
        sFrx = GetIcon(sPath, sIconForm)
        If Not Right(sPath, 1) = "\" Then sPath = sPath & "\"
        sFrx = GetIcon(sPath, sIconForm)
        On Error Resume Next
        If Not sFrx = vbNullString Then
          FileCopy sPath & sFrx, gCurrentDir & sFrx
        End If
    End If
End If
End Sub
Private Sub CompileProject()
Dim sCompile As String, makePath As String
Dim iTask As Long, ret As Long, pHandle As Long
    cmdNext.Enabled = False
    cmdCancel.Enabled = False
    cmdBack.Enabled = False
    lblStatus.Caption = "Compiling Project...Please Wait."
    lblStatus.Refresh
    Screen.MousePointer = 11
    Call LoadProject
    Call CreateProjectFiles
On Error GoTo compileErr
    makePath = gCurrentDir & "AliasEXE32.vbp"
    sCompile = sVB6exe & " /make """ & makePath & """"
        iTask = Shell(sCompile, vbHide)
        pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
        ret = WaitForSingleObject(pHandle, INFINITE)
        ret = CloseHandle(pHandle)
    Call CopyComplete
    cmdNext.Enabled = True
    cmdCancel.Enabled = True
    Screen.MousePointer = 0
Exit Sub
compileErr:
 MsgBox CStr(Err), 16
    lblStatus.Caption = "Error..Press Finish to Retry."
    lblStatus.Refresh
    cmdNext.Enabled = True
    cmdCancel.Enabled = True
End Sub
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
End Sub

Private Sub cmdNext_Click()
If cmdNext.Caption = "&Next ->" Then
    If iCount = 0 Then
            cmdBack.Enabled = True
            'If sVB6exe = vbNullString Then
            iCount = IIf(sVB6exe = vbNullString, iCount + 1, iCount + 2)
            Call NextFrame(iCount)
            Exit Sub
    ElseIf iCount = 1 Then
         If Len(txtVB6Path.Text) = 0 Then
          MsgBox "Please Select Your VB6.EXE Location before continuing...", 16
          Exit Sub
         Else
            cmdBack.Enabled = True
            iCount = iCount + 1
            Call NextFrame(iCount)
            Exit Sub
         End If
    ElseIf iCount = 2 Then
         If Len(txtProjectPath.Text) = 0 Then
          MsgBox "Please Select Your .VBP Project File before continuing...", 16
          Exit Sub
         Else
            cmdBack.Enabled = True
            iCount = iCount + 1
            Call NextFrame(iCount)
            Exit Sub
         End If
    ElseIf iCount = 3 Then
         If Len(txtDestinationPath.Text) = 0 Then
          MsgBox "Please Select Your Network Folder before continuing...", 16
          Exit Sub
         Else
            cmdBack.Enabled = True
            iCount = iCount + 1
            Call CreateSummary
            Call NextFrame(iCount)
            cmdNext.Caption = "&Finish"
         End If
    End If
Else
 Call CompileProject
End If
End Sub

Private Sub File1_Click()

End Sub

Private Sub Form_Load()
GetRegSettings
frmMain.Move locLeft, locTop, locWidth, locHeight
Call setForm
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call SaveRegSettings
Close
End Sub


Private Sub Form_Resize()
picSetup.Move 0, 0, ScaleWidth, 3260

End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub Image1_Click()
'Call CreateNewDirectory(sProjectSetupFolder)
End Sub

Private Sub txtDestinationPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtDestinationPath.ToolTipText = Trim(txtDestinationPath.Text)
End Sub


Private Sub txtProjectPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtProjectPath.ToolTipText = Trim(txtProjectPath.Text)
End Sub


Private Sub txtVB6Path_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtVB6Path.ToolTipText = Trim(txtVB6Path.Text)
End Sub


