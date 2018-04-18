Attribute VB_Name = "apiMisc"
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


' elevate privilege
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long

Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8

Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Const ANYSIZE_ARRAY = 1
Private Type LUID_AND_ATTRIBUTES
        pLuid As LUID
        Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Private Const SE_DEBUG_PRIVILEGE = "SeDebugPrivilege"
Private Const SE_PRIVILEGE_ENABLED = &H2


Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function NtSuspendProcess Lib "ntdll.dll" (ByVal hProc As Long) As Long
Public Declare Function NtResumeProcess Lib "ntdll.dll" (ByVal hProc As Long) As Long

Public Declare Sub Sleep Lib "kernel32 " (ByVal dwMilliseconds As Long)

Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)

Public Const UP_BUTTON = &H40004
Public Const DOWN_BUTTON = &HE000E

Public Const TH32CS_SNAPPROCESS = &H2
Public Const PROCESS_SUSPEND_RESUME = &H800

Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202

Public Const WS_EX_TOPMOST = &H8

Public Const TOOLTIPS_CLASS = "tooltips_class32"

Public Const CW_USEDEFAULT = &H80000000

Public Const MAX_PATH = 260


Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Public Type LASTINPUTINFO
    cbSize As Long
    dwTime As Long
End Type

Private Const MAX_LOG = 128 ' maximum logger entities

Public Const CFG_FILE = "SETI@REST.ini"
Public Const CFG_APP_NAME = "SETI@HOME"

Public errcode As Long
Public appPath As String

Sub Main()

    If App.PrevInstance = True Then
        MsgBox App.EXEName & " is already running!", vbCritical
    End If
    
    'vars
    If Len(App.Path) = 3 Then
        appPath = App.Path
    Else
        appPath = App.Path & "\"
    End If
    
    'objs
    
    InitCommonControls
    
    ElevatePrivilege
    
    initProc
    
    Load frmMain
    frmMain.Show
    
End Sub

Private Sub ElevatePrivilege()
    Dim hProcess As Long
    Dim hToken As Long
    
    
    Dim lUniqueID As LUID
    
    Dim lResult As Long
    
    
    Dim tp As TOKEN_PRIVILEGES
    Dim tpPrev As TOKEN_PRIVILEGES
    
    Dim tpPrevLen As Long
    
    hProcess = GetCurrentProcess
    lResult = OpenProcessToken(hProcess, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken)
    lResult = LookupPrivilegeValue(vbNullString, SE_DEBUG_PRIVILEGE, lUniqueID)
    
'    tp.PrivilegeCount = 1
'    tp.Privileges(0).pLuid = lUniqueID
'    tp.Privileges(0).Attributes = 0
'    lResult = AdjustTokenPrivileges(hToken, -1, tp, Len(tp), tpNew, Len(tpNew))
    
    ' set the new privilege
    tp.PrivilegeCount = 1
    tp.Privileges(0).pLuid = lUniqueID
    tp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    
    errcode = Err.LastDllError
    
    lResult = AdjustTokenPrivileges(hToken, False, tp, Len(tp), tpPrev, tpPrevLen)
    
    errcode = Err.LastDllError
    
    CloseHandle (hToken)
End Sub

Public Function GetINI(Key As String, Optional Default = vbNullString) As String
    Dim strTemp As String * 128
    GetPrivateProfileString CFG_APP_NAME, Key, Default, strTemp, Len(strTemp), appPath & CFG_FILE
    GetINI = Left(strTemp, InStr(strTemp, Chr(0)) - 1)
End Function

Public Function SaveINI(Key As String, Setting As String)
    WritePrivateProfileString CFG_APP_NAME, Key, Setting, appPath & CFG_FILE
End Function

Public Sub logger(txt As String, Optional noTime As Boolean)
    If frmMain.lstLog.ListCount > MAX_LOG Then
        frmMain.lstLog.Clear
    End If
    
    If noTime = False Then
        frmMain.lstLog.AddItem "[" & Time & "] " & txt, 0
    Else
        frmMain.lstLog.AddItem txt, 0
    End If
End Sub

