Attribute VB_Name = "Module1"
Option Explicit
'Declarations for "FileExists" and registry functions
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const ERROR_SUCCESS = 0&
Private Function FileExists(sSource As String) As Boolean
    If Right(sSource, 2) = ":\" Then
        Dim allDrives As String
        allDrives = Space$(64)
        Call GetLogicalDriveStrings(Len(allDrives), allDrives)
        FileExists = InStr(1, allDrives, Left(sSource, 1), 1) > 0
        Exit Function
    Else
        If Not sSource = "" Then
            Dim WFD As WIN32_FIND_DATA
            Dim hFile As Long
            hFile = FindFirstFile(sSource, WFD)
            FileExists = hFile <> INVALID_HANDLE_VALUE
            Call FindClose(hFile)
        Else
            FileExists = False
        End If
    End If
End Function
Private Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim hCurKey As Long
    Dim lRegResult As Long
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    If lRegResult <> ERROR_SUCCESS Then
    End If
    lRegResult = RegCloseKey(hCurKey)
End Sub
Private Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
    Dim hCurKey As Long
    Dim lValueType As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    Dim intZeroPos As Integer
    Dim lRegResult As Long
    If Not IsEmpty(Default) Then
      GetSettingString = Default
    Else
      GetSettingString = ""
    End If
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)
    If lRegResult = ERROR_SUCCESS Then
      If lValueType = REG_SZ Or REG_EXPAND_SZ Then
        strBuffer = String(lDataBufferSize, " ")
        lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
         intZeroPos = InStr(strBuffer, Chr$(0))
        If intZeroPos > 0 Then
          GetSettingString = Left$(strBuffer, intZeroPos - 1)
        Else
          GetSettingString = strBuffer
        End If
        If lValueType = REG_EXPAND_SZ Then GetSettingString = StripTerminator(ExpandEnvStr(GetSettingString))
      End If
    Else
    End If
    lRegResult = RegCloseKey(hCurKey)
End Function
'String functions
Private Function ExpandEnvStr(sData As String) As String
    Dim C As Long, s As String
    s = ""
    C = ExpandEnvironmentStrings(sData, s, C)
    s = String$(C - 1, 0)
    C = ExpandEnvironmentStrings(sData, s, C)
    ExpandEnvStr = s
End Function
Private Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
Private Function strUnQuoteString(ByVal strQuotedString As String)
    strQuotedString = Trim$(strQuotedString)
    If Mid$(strQuotedString, 1, 1) = Chr(34) Then
        If Right$(strQuotedString, 1) = Chr(34) Then
            strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
        End If
    End If
    strUnQuoteString = strQuotedString
End Function
Private Function PathOnly(ByVal filepath As String) As String
    Dim temp As String
    temp = Mid$(filepath, 1, InStrRev(filepath, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    PathOnly = temp
End Function
Private Function ChangeExt(ByVal filepath As String, Optional newext As String) As String
    Dim temp As String
    If InStr(1, filepath, ".") = 0 Then
        temp = filepath
    Else
        temp = Mid$(filepath, 1, InStrRev(filepath, "."))
        temp = Left(temp, Len(temp) - 1)
    End If
    If newext <> "" Then newext = "." + newext
    ChangeExt = temp + newext
End Function
Private Function ExtOnly(ByVal filepath As String, Optional dot As Boolean) As String
    ExtOnly = Mid$(filepath, InStrRev(filepath, ".") + 1)
    If dot = True Then ExtOnly = "." + ExtOnly
End Function
Private Function FileOnly(ByVal filepath As String) As String
    FileOnly = Mid$(filepath, InStrRev(filepath, "\") + 1)
End Function
Private Function AddBackSlash(mPath As String) As String
    AddBackSlash = IIf(Right(mPath, 1) = "\", mPath, mPath & "\")
End Function
Private Function OneGulp(Src As String) As String
    On Error Resume Next
    Dim f As Integer, temp As String
    f = FreeFile
    DoEvents
    Open Src For Binary As #f
    temp = String(LOF(f), Chr$(0))
    Get #f, , temp
    Close #f
    OneGulp = temp
End Function
'Respond to the Shell
Public Sub Main()
    Dim MyCommand As String
    Dim RootDir As String
    Dim TargetVbw As String
    Dim TargetVbwNew As String
    Dim VBPath As String
    Dim VBSwitch As String
    Dim temp As String
    Dim tmp() As String
    Dim z As Long
    Dim mCurDir As String
    MyCommand = Command()
    'Always a good idea to remove surrounding quotes
    'from a command line - the "Send To" command adds
    'quotes for example
    MyCommand = strUnQuoteString(MyCommand)
    If Len(MyCommand) = 0 Then
        'No command line so the exe was clicked
        Form1.Show
    Else
        If ScannerEnabled = 0 Then
            'This should never get called, but in case something goes
            'wrong then just pass the command line to VB
            Shell GetVBPath & " " & MyCommand, vbNormalFocus
        Else
            If VBWScanEnabled = 0 Then GoTo Done
            VBPath = GetVBPath 'VB6.EXE
            If Len(VBPath) = 0 Then GoTo Done
            'Interpret other command line switches other than "Open"
            Select Case LCase(Right(MyCommand, 5))
                Case " /run"
                    MyCommand = Left(MyCommand, Len(MyCommand) - 5)
                    VBSwitch = " /run"
                Case "/make"
                    MyCommand = Left(MyCommand, Len(MyCommand) - 6)
                    VBSwitch = " /make"
            End Select
            Select Case LCase(ExtOnly(MyCommand))
            Case "vbp"
                RootDir = PathOnly(MyCommand) 'Project folder
                TargetVbw = ChangeExt(MyCommand, "vbw") 'vbw file
                TargetVbwNew = ChangeExt(MyCommand, "bbw") 'name used for renaming
                If FileExists(TargetVbwNew) Then Kill TargetVbwNew 'remove it if it exists
                'We could just kill the file but renaming is neater
                'as you can choose to use it if you prefer
                If FileExists(TargetVbw) Then Name TargetVbw As TargetVbwNew
            Case "vbg" 'Project Groups are a bit trickier
                'We need to parse the .vbg file to find all the Projects
                mCurDir = CurDir
                temp = OneGulp(MyCommand)
                tmp = Split(temp, vbCrLf)
                For z = 0 To UBound(tmp)
                    If InStr(1, tmp(z), "=") Then
                        temp = Split(tmp(z), "=")(1)
                        If InStr(1, temp, "\") <> 0 Then
                            'Project path may be in a different folder
                            'and appear something like
                            '..\..\..\Project1.vbp
                            If InStr(1, PathOnly(temp), ".") <> 0 Then
                                ChDir PathOnly(temp)
                                temp = AddBackSlash(CurDir) & FileOnly(temp)
                            Else
                                'a pure relative path like Settings\Project1.vbp
                                If Not FileExists(temp) Then
                                    temp = AddBackSlash(PathOnly(MyCommand)) & temp
                                    If Not FileExists(temp) Then GoTo Done
                                End If
                            End If
                        End If
                        RootDir = PathOnly(temp) 'Project folder
                        TargetVbw = ChangeExt(temp, "vbw") 'vbw file
                        TargetVbwNew = ChangeExt(temp, "bbw") 'name used for renaming
                        If FileExists(TargetVbwNew) Then Kill TargetVbwNew 'remove it if it exists
                        'We could just kill the file but renaming is neater
                        'as you can choose to use it if you prefer
                        If FileExists(TargetVbw) Then Name TargetVbw As TargetVbwNew
                    End If
                Next
                ChDir mCurDir
            End Select
            
            'OK - now we're safe to run VB
            Shell VBPath & " " & Chr(34) & MyCommand & Chr(34) & VBSwitch, vbNormalFocus
        End If
        GoTo Done
    End If
    Exit Sub
Done:
    End
End Sub
'Are we intercepting Shell commands to VB?
Public Function ScannerEnabled() As Long
    Dim temp As String
    temp = GetSettingString(HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\open\command", "")
    ScannerEnabled = IIf(temp = Chr(34) & AddBackSlash(App.Path) & "VBPrjScr.exe" & Chr(34) & " " & Chr(34) & "%1" & Chr(34), 1, 0)
End Function
Public Function VBWScanEnabled() As Long
'Are we disabling the .vbw files?
    VBWScanEnabled = Val(GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "DisableVBW", 1))
End Function
Public Sub EnableScanning(mEnabled As Long)
    Dim AppRunPath As String
    If mEnabled = 0 Then
        'Return Shell commands to VB
        If ScannerEnabled = 0 Then Exit Sub
        SaveSettingString HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\open\command", "", Chr(34) & GetVBPath & Chr(34) & " " & Chr(34) & "%1" & Chr(34)
        SaveSettingString HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\Make\command", "", GetVBPath & Chr(34) & "%1" & Chr(34) & " /make"
        SaveSettingString HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\Run Project\command", "", GetVBPath & Chr(34) & "%1" & Chr(34) & " /run"
        SaveSettingString HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\open\command", "", Chr(34) & GetVBPath & Chr(34) & " " & Chr(34) & "%1" & Chr(34)
        SaveSettingString HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\Make\command", "", GetVBPath & Chr(34) & "%1" & Chr(34) & " /make"
        SaveSettingString HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\Run Project\command", "", GetVBPath & Chr(34) & "%1" & Chr(34) & " /run"
    Else
        'Intercepting Shell commands to VB
        If ScannerEnabled = 1 Then Exit Sub
        AppRunPath = AddBackSlash(App.Path) & "VBPrjScr.exe"
        SaveSettingString HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\open\command", "", Chr(34) & AppRunPath & Chr(34) & " " & Chr(34) & "%1" & Chr(34)
        SaveSettingString HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\Make\command", "", AppRunPath & Chr(34) & "%1" & Chr(34) & " /make"
        SaveSettingString HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\Run Project\command", "", AppRunPath & Chr(34) & "%1" & Chr(34) & " /run"
        SaveSettingString HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\open\command", "", Chr(34) & AppRunPath & Chr(34) & " " & Chr(34) & "%1" & Chr(34)
        SaveSettingString HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\Make\command", "", AppRunPath & Chr(34) & "%1" & Chr(34) & " /make"
        SaveSettingString HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\Run Project\command", "", AppRunPath & Chr(34) & "%1" & Chr(34) & " /run"
    End If
End Sub
Public Function GetVBPath() As String
    'Ascertain VB path from the "Default Icon" registry setting
    Dim VBPath As String, z As Long
    VBPath = GetSettingString(HKEY_CLASSES_ROOT, "VisualBasic.Project\DefaultIcon", "")
    If Len(VBPath) <> 0 Then
        z = InStr(1, VBPath, ",")
        If z < 1 Then
            If FileExists(VBPath) Then GetVBPath = VBPath
        Else
            VBPath = Left(VBPath, z - 1)
            If FileExists(VBPath) Then GetVBPath = VBPath
        End If
    End If
End Function
