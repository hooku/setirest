VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SETI@REST Alpha"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmMain.frx":0442
   ScaleHeight     =   7515
   ScaleWidth      =   5910
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAdvanced 
      Caption         =   "&Advanced"
      Height          =   435
      Left            =   120
      TabIndex        =   19
      Top             =   6960
      Width           =   1635
   End
   Begin SETIREST.ctlTray trySETI 
      Left            =   180
      Top             =   1860
      _ExtentX        =   635
      _ExtentY        =   635
   End
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   5655
      TabIndex        =   18
      Top             =   2340
      Visible         =   0   'False
      Width           =   5655
   End
   Begin SETIREST.ctlBalloontip balSETI 
      Left            =   540
      Top             =   1860
      _ExtentX        =   635
      _ExtentY        =   635
   End
   Begin VB.CommandButton cmdMinimize 
      Caption         =   "Minimize to &Tray"
      Default         =   -1  'True
      Height          =   435
      Left            =   4140
      TabIndex        =   13
      Top             =   6960
      Width           =   1635
   End
   Begin VB.CommandButton cmdManual 
      Caption         =   "&Manual"
      Height          =   435
      Left            =   2400
      TabIndex        =   12
      Top             =   6960
      Width           =   1635
   End
   Begin VB.ListBox lstLog 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      ItemData        =   "frmMain.frx":058C
      Left            =   120
      List            =   "frmMain.frx":058E
      TabIndex        =   7
      Top             =   120
      Width           =   5655
   End
   Begin VB.Frame famSpeedFan 
      Caption         =   "Fan Control"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   5655
      Begin VB.PictureBox boxSpeedFan 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   60
         ScaleHeight     =   1335
         ScaleWidth      =   5535
         TabIndex        =   5
         Top             =   300
         Width           =   5535
         Begin VB.CheckBox chkSpeedFan 
            Caption         =   "Enable Speed&Fan Speed Toggle"
            ForeColor       =   &H80000002&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Tag             =   $"frmMain.frx":0590
            Top             =   60
            Width           =   3315
         End
         Begin VB.CheckBox chkHideSpeedfan 
            Caption         =   "Auto &Hide SpeedFan Main Window"
            Height          =   315
            Left            =   360
            TabIndex        =   8
            Tag             =   "Make SpeedFan main window hide from desktop & taskbar.#Hide SpeedFan Main Window#1"
            Top             =   960
            Width           =   4815
         End
         Begin VB.ComboBox cmbSpeed 
            Height          =   375
            HelpContextID   =   5
            ItemData        =   "frmMain.frx":06E4
            Left            =   2880
            List            =   "frmMain.frx":0706
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Tag             =   "!!!SET TO MAXIMUM SPEED MAY CAUSE UNDESIRED BEHAVIOR TO YOUR FAN!!!#Fan Speed#2"
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label labSpeed 
            Caption         =   "Force Fan Speed to"
            Height          =   255
            Left            =   360
            TabIndex        =   9
            Top             =   540
            Width           =   2415
         End
      End
   End
   Begin VB.Frame famSETI 
      Caption         =   "GPU Usage Toggle"
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   5655
      Begin VB.PictureBox boxSETI 
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   60
         ScaleHeight     =   1935
         ScaleWidth      =   5535
         TabIndex        =   2
         Top             =   300
         Width           =   5535
         Begin VB.CheckBox chkSETI 
            Caption         =   "Enable &SETI@HOME GPU Toggle"
            ForeColor       =   &H80000002&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Tag             =   $"frmMain.frx":0746
            Top             =   60
            Width           =   3315
         End
         Begin VB.ComboBox cmbIntensity 
            Height          =   375
            HelpContextID   =   4
            ItemData        =   "frmMain.frx":07C9
            Left            =   2820
            List            =   "frmMain.frx":0863
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Tag             =   $"frmMain.frx":092B
            Top             =   1440
            Width           =   2415
         End
         Begin VB.ComboBox cmbIdleTime 
            Height          =   375
            HelpContextID   =   1
            ItemData        =   "frmMain.frx":0A07
            Left            =   2820
            List            =   "frmMain.frx":0A1A
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Tag             =   "The amount of idle time before make GPU to 100%.#Idle Time Threshold#1"
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox txtKeyword 
            Height          =   405
            Left            =   2820
            TabIndex        =   3
            Tag             =   $"frmMain.frx":0A4C
            Text            =   "cuda|opencl"
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label labIntensity 
            Caption         =   "Toggle Intensity"
            Height          =   255
            Left            =   360
            TabIndex        =   14
            Top             =   1500
            Width           =   2415
         End
         Begin VB.Label labKeyword 
            Caption         =   "GPU Process Keywords"
            Height          =   255
            Left            =   360
            TabIndex        =   11
            Top             =   540
            Width           =   2415
         End
         Begin VB.Label labIdleTime 
            Caption         =   "Idle Time Threshold"
            Height          =   255
            Left            =   360
            TabIndex        =   10
            Top             =   1020
            Width           =   2415
         End
      End
   End
   Begin VB.Menu mAdvanced 
      Caption         =   "advanced"
      Visible         =   0   'False
      Begin VB.Menu mThrottleStop 
         Caption         =   "ThrottleStop Idle Toggle"
      End
   End
   Begin VB.Menu mLog 
      Caption         =   "log"
      Visible         =   0   'False
      Begin VB.Menu mClearLog 
         Caption         =   "&Clear Log"
      End
   End
   Begin VB.Menu mTray 
      Caption         =   "tray"
      Visible         =   0   'False
      Begin VB.Menu mAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mSeperator0 
         Caption         =   "-"
      End
      Begin VB.Menu mSETI 
         Caption         =   "Enable &SETI@HOME GPU Toggle"
      End
      Begin VB.Menu mSpeedFan 
         Caption         =   "Enable Speed&Fan Speed Toggle"
      End
      Begin VB.Menu mSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long

Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&

Private Const MAX_LOG = 128 ' maximum logger entities

Private Const hNull = 0

Private Sub chkHideSpeedfan_Click()
    procMain.cfg.isHideSpeedFan = Me.chkHideSpeedfan.Value

    If Me.chkHideSpeedfan.Value = 1 Then
        ctrlSpeedFan.HideSpeedFan
    Else
        ctrlSpeedFan.ShowSpeedFan
    End If
End Sub

Private Sub chkSETI_Click()
    procMain.cfg.isSETIEnable = Me.chkSETI.Value

    If Me.chkSETI.Value = 1 Then
        ctrlSETI.SetKeywordSETA Me.txtKeyword
        Dim szSETIProcessName As String
        szSETIProcessName = ctrlSETI.FindSETA
        If Len(szSETIProcessName) > 0 Then
            logger "Found " & szSETIProcessName
        Else
            MsgBox "Please check either your BOINC's GPU compution is enabled or the GPU process keyword is correct." & vbCrLf & vbCrLf & App.EXEName & " will continue to detect your BOINC's GPU process.", vbExclamation, "SETI GPU Process is Not Detected"
        End If
    Else
        logger "GPU Usage Toggle Disabled"
    End If

    Me.mSETI.Checked = Me.chkSETI.Value
    Me.labKeyword.Enabled = (Me.chkSETI.Value = 1)
    Me.txtKeyword.Enabled = (Me.chkSETI.Value = 1)
    Me.labIdleTime.Enabled = (Me.chkSETI.Value = 1)
    Me.cmbIdleTime.Enabled = (Me.chkSETI.Value = 1)
    Me.labIntensity.Enabled = (Me.chkSETI.Value = 1)
    Me.cmbIntensity.Enabled = (Me.chkSETI.Value = 1)
End Sub

Private Sub chkSpeedFan_Click()
    procMain.cfg.isSpeedFanEnable = Me.chkSpeedFan.Value
    
    If Me.chkSpeedFan.Value = 1 Then
        If ctrlSpeedFan.FindSpeedFan = True Then
            logger "Found Speedfan Window!"
        Else
            MsgBox "Please make sure that SpeedFan has been started, and the SpeedFan Main Window is not minimized to tray!", vbExclamation, "Failed to find SpeedFan Window"
            Me.chkSpeedFan.Value = 0
        End If
    Else
        ctrlSpeedFan.ShowSpeedFan ' restore speedfan
    End If
    
    Me.mSpeedFan.Checked = Me.chkSpeedFan.Value
    Me.labSpeed.Enabled = (Me.chkSpeedFan.Value = 1)
    Me.cmbSpeed.Enabled = (Me.chkSpeedFan.Value = 1)
    Me.chkHideSpeedfan.Enabled = (Me.chkSpeedFan.Value = 1)
End Sub

Private Sub cmbIdleTime_Click()
    procMain.cfg.idleTime = Me.cmbIdleTime.ListIndex
End Sub

Private Sub cmbIntensity_Click()
    procMain.cfg.intensity = Me.cmbIntensity.ListIndex
End Sub

Private Sub cmbSpeed_Click()
    procMain.cfg.speed = Me.cmbSpeed.ListIndex
End Sub

Private Sub cmdAdvanced_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.PopupMenu Me.mAdvanced, , Me.cmdAdvanced.Left, Me.cmdAdvanced.Top + Me.cmdAdvanced.Height
End Sub

Private Sub cmdManual_Click()
    Shell "explorer http://win2000.howeb.cn/setirest.html"
End Sub

Private Sub cmdMinimize_Click()
    Me.Hide
    Me.trySETI.AddIcon 0, Me.MouseIcon, App.EXEName
End Sub

Private Sub famSETI_DblClick()
    On Error Resume Next
    Shell "C:\Program Files\GPU-Z\GPU-Z.exe", vbNormalFocus
End Sub

Private Sub famSpeedFan_DblClick()
    On Error Resume Next
    Shell "C:\Program Files\SpeedFan\speedfan.exe", vbNormalFocus
End Sub

Private Sub Form_Load()
'vars
    

    
    'lii.cbSize = Len(lii)
    
    With cfg
        .isSETIEnable = GetINI("EnableSeti", Me.chkSETI.Value)
        .szKeyword = GetINI("Keyword", Me.txtKeyword.Text)
        .idleTime = GetINI("IdleTime", Me.cmbIdleTime.HelpContextID)
        .intensity = GetINI("Intensity", Me.cmbIntensity.HelpContextID)
        
        .isSpeedFanEnable = GetINI("EnableSpeedfan", Me.chkSpeedFan.Value)
        .speed = GetINI("Speed", Me.cmbSpeed.HelpContextID)
        .isHideSpeedFan = GetINI("HideSpeedfan", Me.chkHideSpeedfan.Value)
        .isExperimentalFeatureEnable = GetINI("EnableExperimentalFeature", CInt(Me.cmdAdvanced.Visible))
    End With
    
'objs
    RefreshUI
    
    Me.lstLog.Height = 2460
    
    If GetINI("DisableBalloontip") <> "1" Then
        With Me.balSETI
            .AddBalloon Me.chkSETI
            .AddBalloon Me.txtKeyword
            .AddBalloon Me.cmbIdleTime
            .AddBalloon Me.cmbIntensity
            
            .AddBalloon Me.chkSpeedFan
            .AddBalloon Me.cmbSpeed
            .AddBalloon Me.chkHideSpeedfan
        End With
    End If
    
    'chkSETI_Click
    'chkSpeedFan_Click
    
' append about menu
    Dim hSysMenu As Long
    hSysMenu = GetSystemMenu(Me.hwnd, False)

    AppendMenu hSysMenu, MF_SEPARATOR, 0, vbNullString
    AppendMenu hSysMenu, MF_STRING, 0, "&About..."

    logger App.title & "Alpha Ver " & App.Major & "." & App.Minor & " Build " & App.Revision, True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveINI "EnableSeti", Me.chkSETI.Value
    SaveINI "Keyword", Me.txtKeyword.Text
    SaveINI "IdleTime", Me.cmbIdleTime.ListIndex
    SaveINI "Intensity", Me.cmbIntensity.ListIndex
    
    SaveINI "EnableSpeedfan", Me.chkSpeedFan.Value
    SaveINI "Speed", Me.cmbSpeed.ListIndex
    SaveINI "HideSpeedfan", Me.chkHideSpeedfan.Value
    
    SaveINI "EnableExperimentalFeature", Abs(CInt(Me.cmdAdvanced.Visible))
    
    SaveINI "AppRevision", App.Revision
    
    Me.trySETI.DeleteIcon 0
    
    recycleProc
    
    End
End Sub

Private Sub lstLog_DblClick()
    Clipboard.SetText Me.lstLog.List(Me.lstLog.ListIndex)
End Sub

Private Sub lstLog_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu Me.mLog
    End If
End Sub

Private Sub mAbout_Click()
    MsgBox App.title & " " & App.Major & "." & App.Minor & " Build " & App.Revision _
    & vbCrLf & vbCrLf & _
    "Xiaojing" & vbCrLf & vbCrLf & "http://win2000.howeb.cn", _
    vbInformation, "About " & App.title
End Sub

Private Sub mClearLog_Click()
    Me.lstLog.Clear
End Sub

Private Sub RefreshUI()
    With Me
        .chkSETI.Value = Abs(CInt(procMain.cfg.isSETIEnable))
        .txtKeyword.Text = procMain.cfg.szKeyword
        .cmbIdleTime.ListIndex = procMain.cfg.idleTime
        .cmbIntensity.ListIndex = procMain.cfg.intensity
        
        .chkSpeedFan.Value = Abs(CInt(procMain.cfg.isSpeedFanEnable))
        .cmbSpeed.ListIndex = procMain.cfg.speed
        .chkHideSpeedfan.Value = Abs(CInt(procMain.cfg.isHideSpeedFan))
        
        .cmdAdvanced.Visible = Abs(CInt(procMain.cfg.isExperimentalFeatureEnable))
    End With
End Sub

Private Sub mExit_Click()
    Unload Me
End Sub

Private Sub mRestore_Click()
    Me.Show
    Me.trySETI.DeleteIcon 0
End Sub

Private Sub mSETI_Click()
    Me.chkSETI.Value = 1 - Me.chkSETI.Value
End Sub

Private Sub mSpeedFan_Click()
    Me.chkSpeedFan.Value = 1 - Me.chkSpeedFan.Value
End Sub

Private Sub mThrottleStop_Click()
    MsgBox "Not support"
End Sub

Private Sub trySETI_DblClick(Button As Integer, ID As Integer)
    mRestore_Click
End Sub

Private Sub trySETI_MouseDown(Button As Integer, ID As Integer)
    SetForegroundWindow Me.hwnd
    Me.PopupMenu Me.mTray
End Sub

Private Sub txtKeyword_LostFocus()
    ctrlSETI.SetKeywordSETA Me.txtKeyword
End Sub
