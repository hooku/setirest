VERSION 5.00
Begin VB.UserControl ctlBalloontip 
   BackColor       =   &H80000018&
   BackStyle       =   0  'Transparent
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   EditAtDesignTime=   -1  'True
   InvisibleAtRuntime=   -1  'True
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
   ToolboxBitmap   =   "ctlBalloontip.ctx":0000
   Begin VB.Image Image1 
      Height          =   330
      Left            =   0
      Picture         =   "ctlBalloontip.ctx":00FA
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "ctlBalloontip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const WS_POPUP As Long = &H80000000
Private Const WS_BORDER As Long = &H800000

Private Const TTF_IDISHWND                As Long = &H1
Private Const TTF_TRANSPARENT = &H100
Private Const TTF_CENTERTIP = &H2
Private Const TTM_ADDTOOLA = WM_USER + 4
Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA = WM_USER + 12
Private Const TTM_SETMAXTIPWIDTH = WM_USER + 24
Private Const TTM_SETTIPBKCOLOR = WM_USER + 19
Private Const TTM_SETTIPTEXTCOLOR = WM_USER + 20
Private Const TTM_SETTITLE = WM_USER + 32
Private Const TTM_SETDELAYTIME            As Long = WM_USER + 3
Private Const TTM_SETMARGIN               As Long = WM_USER + 26
Private Const TTM_UPDATE                  As Long = (WM_USER + 29)
Private Const TTS_NOPREFIX = &H2
Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
Private Const TTDT_AUTOPOP                As Long = 2
Private Const TTDT_AUTOMATIC              As Long = 0
Private Const TTDT_INITIAL                As Long = 3
Private Const TTDT_RESHOW                 As Long = 1
'Tool Tip Icons
Private Const TTI_ERROR                   As Long = 3
Private Const TTI_INFO                    As Long = 1
Private Const TTI_NONE                    As Long = 0
Private Const TTI_WARNING                 As Long = 2
'Tool Tip API Class
Private Const TOOLTIPS_CLASSA = "tooltips_class32"



Private Const TipDelayAuto As Long = 1000
Private Const TipDelayInit As Long = 1000
Private Const TipDelayResh As Long = 1000

Private Const MaxIndex = 8

Private Type RECT
    Left                              As Long
    Top                               As Long
    Right                             As Long
    Bottom                            As Long
End Type

Private Type TOOLINFO
    lSize                             As Long
    lFlags                            As Long
    hwnd                              As Long
    lId                               As Long
    lpRect                            As RECT
    hInstance                         As Long
    lpszText                          As String
    lParam                            As Long
End Type

Private hWndBalloon(MaxIndex) As Long
Private balloonCount As Integer

Private ti                                As TOOLINFO

Public Sub AddBalloon(Parent As Object)
    Dim Text As String, title As String, icon As Integer
    Dim strTemp() As String
    strTemp = Split(Parent.Tag, "#")
    
    AddTooltip Parent.hwnd, strTemp(0), strTemp(1), Val(strTemp(2))
End Sub

Private Function AddTooltip(hwnd As Long, Text As String, Optional title As String = vbNullString, Optional icon As Integer = TTI_NONE)
    hWndBalloon(balloonCount) = CreateWindowEx(0&, TOOLTIPS_CLASSA, vbNullString, _
    WS_POPUP Or WS_BORDER Or TTS_ALWAYSTIP Or TTS_NOPREFIX Or TTS_BALLOON, _
    CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, _
    hwnd, 0&, App.hInstance, 0&)

    'MsgBox hWndBalloon(balloonCOunt)
    'Make balloon topmost
    SetWindowPos hWndBalloon(balloonCount), HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
            
    'Get the rectangle of parent
    GetClientRect hwnd, ti.lpRect
    
    With ti
        .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
        .hwnd = hwnd
        .lId = 0
        .hInstance = App.hInstance
        .lpszText = Text
    End With

    ' Add the tooltip structure
    SendMessage hWndBalloon(balloonCount), TTM_ADDTOOLA, 0&, ti
    'SetLastError 0
    SendMessage hWndBalloon(balloonCount), TTM_SETMAXTIPWIDTH, 0, 180
    SendMessage hWndBalloon(balloonCount), TTM_SETTITLE, CLng(icon), ByVal title
    
    ' Set tooltip delay
    SendMessage hWndBalloon(balloonCount), TTM_SETDELAYTIME, TTDT_AUTOPOP, TipDelayAuto
    SendMessage hWndBalloon(balloonCount), TTM_SETDELAYTIME, TTDT_INITIAL, TipDelayInit
    SendMessage hWndBalloon(balloonCount), TTM_SETDELAYTIME, TTDT_RESHOW, TipDelayResh
    
    balloonCount = balloonCount + 1
End Function

Private Sub UserControl_Resize()
    UserControl.Width = 360
    UserControl.Height = 360
End Sub
