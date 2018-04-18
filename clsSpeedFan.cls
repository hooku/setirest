VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpeedFan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SW_HIDE = 0
Private Const SW_SHOW = 5

Dim cycleSpeedFan As Integer
Dim fanPercent As Integer

Dim hSpeedFanWin As Long    ' handle of speedfan main window
Dim hSpeedFan As Long       ' handle of speedfan edit textbox
Dim hSpeedFanTray As Long   ' handle of speedfan tray window

Public Function FindSpeedFan() As Boolean
    Dim hWndTmp As Long
    Dim hWndSpinEdt As Long, hWndWinCtl As Long
    hSpeedFanWin = FindWindowEx(0, 0, "TForm1", "SpeedFan 4.47")
    If hSpeedFanWin = 0 Then
        FindSpeedFan = False
        Exit Function
    End If
    
    hWndTmp = hSpeedFanWin
    hWndTmp = FindWindowEx(hWndTmp, 0, "TPageControl", vbNullString)
    
    
    hWndTmp = FindWindowEx(hWndTmp, 0, "TTabSheet", "Readings")
    hSpeedFanTray = FindWindowEx(0, 0, "TApplication", "SpeedFan")
    
    hWndSpinEdt = FindWindowEx(hWndTmp, 0, "TRxSpinEdit", 0) ' first SpinEdt
    Do ' the last WinCtl is we want
        hWndWinCtl = FindWindowEx(hWndSpinEdt, 0, "TWinControl", vbNullString)
        If hWndWinCtl <> 0 Then
            hSpeedFan = hWndWinCtl
            Exit Do
        End If
        hWndSpinEdt = FindWindowEx(hWndTmp, hWndSpinEdt, "TRxSpinEdit", 0) ' next SpinEdt
    Loop Until hWndSpinEdt = 0
    
    FindSpeedFan = True
End Function

Public Sub HideSpeedFan()
    Debug.Print "HideSpeedFan"

    ShowWindow hSpeedFanWin, SW_HIDE
    ShowWindow hSpeedFanTray, SW_HIDE
End Sub

Public Sub ShowSpeedFan()
    Debug.Print "ShowSpeedFan"
    
    ShowWindow hSpeedFanWin, SW_SHOW
    ShowWindow hSpeedFanTray, SW_SHOW
End Sub

Public Sub SetSpeedFan(Percent As Integer)
    fanPercent = Percent
End Sub

Public Sub RefreshSpeedFan()
    'Debug.Print "RefreshSpeedFan", Timer
    
    'SetWindowText Percent
    PostMessage hSpeedFan, WM_LBUTTONDOWN, 0, DOWN_BUTTON
    PostMessage hSpeedFan, WM_LBUTTONUP, 0, DOWN_BUTTON
    PostMessage hSpeedFan, WM_LBUTTONDOWN, 0, UP_BUTTON
    PostMessage hSpeedFan, WM_LBUTTONUP, 0, UP_BUTTON
End Sub
