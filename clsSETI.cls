VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSETI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FREE_TIME = 1000 ' micro second
Private Const SETI_WORK_TIME = 10 ' micro second
Private Const SETI_SLEEP_TIME = 15 ' micro secondF

Dim hProcessSETA As Long    ' handle of process

Private alreadySuspend As Boolean
Private ProcessKeywords() As String

'Public isAstroPulse As Boolean ' * astropulse consume more GPU and cause significant GDI lag

Public Sub SetKeywordSETA(Text As String)
    On Error GoTo errhandler
    ProcessKeywords = Split(Text, "|")
    Exit Sub
errhandler:
    MsgBox Err.Description, vbCritical
End Sub

Public Function FindSETA() As String
    Dim setaExeName() As String
    Dim procFind As Long, procName As String, procID
    Dim hSnap As Long
    Dim proc As PROCESSENTRY32
    
    ' Clean up first
    CloseHandle hProcessSETA
    hProcessSETA = 0
    
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    'If hSnap = hNull Then
    '    MsgBox "err"
    '    End
    'End If
    
    proc.dwSize = Len(proc)
    procFind = Process32First(hSnap, proc)
    
    Dim iProcessKeywords As Integer
    
    Do While procFind
        procName = Left(proc.szExeFile, InStr(1, proc.szExeFile, Chr(0)) - 1)
        For iProcessKeywords = 0 To UBound(ProcessKeywords)
            If InStr(1, procName, ProcessKeywords(iProcessKeywords), vbTextCompare) > 0 Then
                hProcessSETA = OpenProcess(PROCESS_SUSPEND_RESUME, False, proc.th32ProcessID)
                'hProcessSETA = OpenProcess(CLng(PROCESS_SUSPEND_RESUME), False, CLng(7468))
                'MsgBox hProcessSETA
                'End
                errcode = Err.LastDllError
                FindSETA = procName
                Exit Do
            End If
        Next iProcessKeywords
    
        procFind = Process32Next(hSnap, proc)
    Loop
    
    CloseHandle hSnap
End Function

Public Function SuspendSETA() As Boolean
    Debug.Print "SuspendSETA", Timer

    If alreadySuspend = False Then
        alreadySuspend = True
    Else
        Exit Function
    End If

    Dim hResult As Long
    SuspendSETA = False
    hResult = NtSuspendProcess(hProcessSETA)
    If hResult = 0 Then ' success
        SuspendSETA = True
    Else
        ' SETI process is not find, we'd try to find it
        ' Causes:
        ' 1. user manually checked SETI
        ' 2. GPU process changed
        FindSETA
    End If
End Function

Public Sub ResumeSETA()
    Debug.Print "ResumeSETA", Timer
    
    alreadySuspend = False
    
    NtResumeProcess hProcessSETA
End Sub
