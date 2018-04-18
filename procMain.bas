Attribute VB_Name = "procMain"
Option Explicit

Private Declare Function GetLastInputInfo Lib "user32" (plii As LASTINPUTINFO) As Boolean
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Public Const TICK_RATE = 10

Private Const MAX_INTENSITY = 50

Private Type Config
    
    isSETIEnable As Boolean
    
    szKeyword As String
    idleTime As Integer ' ms
    intensity As Integer
    
    isSpeedFanEnable As Boolean
    
    speed As Integer
    isHideSpeedFan As Boolean

    isExperimentalFeatureEnable As Boolean
End Type


Public ctrlSETI As New clsSETI
Public ctrlSpeedFan As New clsSpeedFan

Public cfg As Config

Dim setiSchedulerTick As Integer
Dim speedfanSchedulerTick As Integer

Dim isIdle As Boolean

Dim freeTime As Long
Dim lii As LASTINPUTINFO

Dim hTimer As Long
Dim hThread As Long, hThreadID As Long

Dim CONST_IDLE_TIME(0 To 4) As Integer
Dim CONST_SPEED(0 To 9) As Integer

Public Sub initProc()
    CONST_IDLE_TIME(0) = 1000
    CONST_IDLE_TIME(1) = 2000
    CONST_IDLE_TIME(2) = 5000
    CONST_IDLE_TIME(3) = 10000
    CONST_IDLE_TIME(4) = 30000
    
    CONST_SPEED(0) = 10
    CONST_SPEED(1) = 9
    CONST_SPEED(2) = 8
    CONST_SPEED(3) = 7
    CONST_SPEED(4) = 6
    CONST_SPEED(5) = 5
    CONST_SPEED(6) = 4
    CONST_SPEED(7) = 3
    CONST_SPEED(8) = 2
    CONST_SPEED(9) = 1
    
    lii.cbSize = Len(lii)
    
    hTimer = SetTimer(0, 0, TICK_RATE, AddressOf controlProc) ' install the timer
End Sub

Public Sub controlProc() 'scheduler
    ' ---IDLE INFO---
    GetLastInputInfo lii
    freeTime = GetTickCount - lii.dwTime
    
    If freeTime > CONST_IDLE_TIME(cfg.idleTime) Then
        ' now we're idle
        If isIdle = False Then
            logger "User Idle" 'Str(freeTime)"
            isIdle = True
            
            ' ---SETI---
            If cfg.isSETIEnable Then
                ctrlSETI.ResumeSETA
            End If
        End If
    Else
        ' now we're busy
        If isIdle = True Then
            logger "User Active"
            isIdle = False
        End If
        
        ' ---SETI---
        If cfg.isSETIEnable Then
            Select Case setiSchedulerTick
            Case cfg.intensity
                ctrlSETI.SuspendSETA
            Case Is >= MAX_INTENSITY
                setiSchedulerTick = -1 ' clear the tick
                ctrlSETI.ResumeSETA
            End Select
            setiSchedulerTick = setiSchedulerTick + 1
        End If
    End If

    ' ---SPEEDFAN---
    If cfg.isSpeedFanEnable Then
        Select Case speedfanSchedulerTick
        Case CONST_SPEED(cfg.speed)
            ctrlSpeedFan.RefreshSpeedFan
        Case Is > CONST_SPEED(cfg.speed)
            speedfanSchedulerTick = -1  ' clear the tick
            'ctrlSpeedFan.ResumeSETA
        End Select
        speedfanSchedulerTick = speedfanSchedulerTick + 1
    End If
End Sub

Public Sub recycleProc()
    ' restore seti
    ctrlSETI.ResumeSETA
    
    ' restore speedfan
    ctrlSpeedFan.ShowSpeedFan
    
    ' destroy timer
    hTimer = KillTimer(0, hTimer)
End Sub
