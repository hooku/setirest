VERSION 5.00
Begin VB.UserControl ctlTray 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1500
   ScaleWidth      =   3000
   ToolboxBitmap   =   "ctlTray.ctx":0000
   Begin VB.PictureBox picTray 
      Height          =   480
      Index           =   0
      Left            =   540
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   330
      Left            =   0
      Picture         =   "ctlTray.ctx":00FA
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "ctlTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Const NIM_ADD = 0
Private Const NIM_MODIFY = 1
Private Const NIM_DELETE = 2
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_MOUSEMOVE = &H200

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIF_MESSAGE = 1
Private Const NIF_ICON = 2
Private Const NIF_TIP = 4
Private Nid As NOTIFYICONDATA

Private Const MaxIndex = 4
Private icon() As Boolean

Event MouseDown(Button As Integer, ID As Integer)
Event DblClick(Button As Integer, ID As Integer)
Event MouseMove(ID As Integer)

Private Sub picTray_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lMsg As Long
Static bInHere As Boolean
lMsg = X / Screen.TwipsPerPixelX
Select Case lMsg
    Case WM_LBUTTONDOWN:
        RaiseEvent MouseDown(1, Index)
    Case WM_RBUTTONDOWN:
        RaiseEvent MouseDown(2, Index)
    Case WM_LBUTTONDBLCLK:
        RaiseEvent DblClick(1, Index)
    Case WM_MOUSEMOVE:
        RaiseEvent MouseMove(Index)
End Select
End Sub

Private Sub UserControl_Initialize()
    ReDim icon(MaxIndex) As Boolean
    Dim i As Integer
    For i = 1 To MaxIndex
        Load picTray(i)
    Next
End Sub

'MemberInfo=0
Public Function AddIcon(Index As Integer, Picture As Long, text As String) As Boolean
    '----------------------------------
    ' Add Icon
    If Index < 0 Or Index > MaxIndex Then
        MsgBox "Error Icon ID"
        Exit Function
    End If
    icon(Index) = True
    Nid.cbSize = Len(Nid)
    Nid.hwnd = picTray(Index).hwnd
    Nid.uID = Index
    Nid.uCallbackMessage = WM_MOUSEMOVE
    Nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    Nid.szTip = text & Chr(0)
    Nid.hIcon = Picture
    Shell_NotifyIcon NIM_ADD, Nid
    '-------------------------------------
End Function

'MemberInfo=0
Public Function DeleteIcon(Index As Integer) As Boolean
    Nid.cbSize = Len(Nid)
    Nid.hwnd = picTray(Index).hwnd
    Nid.uID = Index
    Shell_NotifyIcon NIM_DELETE, Nid
End Function

'MemberInfo=0
Public Function SetIconText(Index As Integer, text As String) As Boolean
    Nid.cbSize = Len(Nid)
    Nid.hwnd = picTray(Index).hwnd
    Nid.uID = Index
    'nid.UCallbackMessage = WM_MOUSEMOVE
    Nid.uFlags = NIF_TIP
    Nid.szTip = text & Chr(0)
    Shell_NotifyIcon NIM_MODIFY, Nid
End Function

'MemberInfo=0
Public Function SetIconPic(Index As Integer, Picture As Long) As Boolean
    Nid.cbSize = Len(Nid)
    Nid.hwnd = picTray(Index).hwnd
    Nid.uID = Index
    Nid.uCallbackMessage = WM_MOUSEMOVE
    Nid.uFlags = NIF_ICON
    Nid.hIcon = Picture
    'Nid.szTip = text & Chr(0)
    Shell_NotifyIcon NIM_MODIFY, Nid
End Function

Private Sub UserControl_Resize()
    UserControl.Width = 360
    UserControl.Height = 360
End Sub
