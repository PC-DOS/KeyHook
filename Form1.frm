VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Key Hook Tool - PC-DOS Workshop"
   ClientHeight    =   2685
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10635
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   10635
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "按鍵位碼封鎖"
      Height          =   2025
      Left            =   6540
      TabIndex        =   13
      Top             =   45
      Width           =   4065
      Begin VB.ListBox List1 
         Height          =   960
         ItemData        =   "Form1.frx":0ECA
         Left            =   780
         List            =   "Form1.frx":0ECC
         TabIndex        =   17
         Top             =   1020
         Width           =   3180
      End
      Begin VB.CommandButton Command3 
         Caption         =   "封鎖該按鍵(&L)"
         Enabled         =   0   'False
         Height          =   345
         Left            =   90
         TabIndex        =   16
         Top             =   615
         Width           =   3855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已封鎖:"
         Height          =   180
         Left            =   90
         TabIndex        =   18
         Top             =   1035
         Width           =   630
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1350
         TabIndex        =   15
         Top             =   180
         Width           =   2625
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "請按一個鍵:"
         Height          =   180
         Left            =   135
         TabIndex        =   14
         Top             =   300
         Width           =   990
      End
   End
   Begin 工程1.cSysTray cSysTray1 
      Left            =   5130
      Top             =   15
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "Form1.frx":0ECE
      TrayTip         =   "Hot Key Lock - 雙擊還原窗口"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "最小化到系統托盤(&M)"
      Height          =   375
      Left            =   8670
      TabIndex        =   11
      Top             =   2265
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "復位全部設定(&R)"
      Height          =   375
      Left            =   6855
      TabIndex        =   10
      Top             =   2265
      Width           =   1710
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "窗口總在最前(&T)"
      Height          =   375
      Left            =   75
      TabIndex        =   9
      Top             =   2265
      Value           =   1  'Checked
      Width           =   1680
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "功能鍵控制選項"
      Height          =   915
      Left            =   60
      TabIndex        =   2
      Top             =   1155
      Width           =   6420
      Begin VB.CheckBox Check7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "禁用Fx功能鍵(&F)"
         Height          =   330
         Left            =   2460
         TabIndex        =   12
         Top             =   540
         Width           =   1830
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "禁用整個鍵盤(&K)"
         Height          =   330
         Left            =   4440
         TabIndex        =   7
         Top             =   540
         Width           =   1815
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "禁用Alt鍵(&A)"
         Height          =   330
         Left            =   150
         TabIndex        =   6
         Top             =   540
         Width           =   1425
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "禁用Shift鍵(&S)"
         Height          =   330
         Left            =   4440
         TabIndex        =   5
         Top             =   210
         Width           =   1800
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "禁用Ctrl鍵(&C)"
         Height          =   330
         Left            =   2460
         TabIndex        =   4
         Top             =   210
         Width           =   1680
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "禁用Windows徽標鍵(&W)"
         Height          =   330
         Left            =   150
         TabIndex        =   3
         Top             =   210
         Width           =   2250
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      Height          =   630
      Left            =   -30
      Top             =   2160
      Width           =   10755
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Key Hook Tool"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   705
      TabIndex        =   8
      Top             =   45
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "您可以在下面勾選需要的功能並立即啟用."
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   855
      Width           =   4785
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "這個工具可以幫助您禁用Windows徽標鍵,Ctrl鍵,Alt鍵,Shift鍵,Fx功能鍵等系統功能鍵,也可以實現對鍵盤和任意按鍵的封鎖"
      Height          =   405
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   495
      Width           =   5790
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   -30
      Picture         =   "Form1.frx":1DA8
      Top             =   45
      Width           =   720
   End
   Begin VB.Menu TrayMenu 
      Caption         =   "TrayMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "顯示主窗口(&M)"
      End
      Begin VB.Menu b0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNoWin 
         Caption         =   "禁用Windows徽標鍵(&W)"
      End
      Begin VB.Menu mnuNoCtrl 
         Caption         =   "禁用Ctrl鍵(&C)"
      End
      Begin VB.Menu mnuNoShift 
         Caption         =   "禁用Shift鍵(&S)"
      End
      Begin VB.Menu mnuNoAlt 
         Caption         =   "禁用Alt鍵(&A)"
      End
      Begin VB.Menu mnuNoKeys 
         Caption         =   "禁用整個鍵盤(&K)"
      End
      Begin VB.Menu mnuNoFx 
         Caption         =   "禁用Fx功能鍵(&F)"
      End
      Begin VB.Menu b1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "復位所有設定(&R)"
      End
      Begin VB.Menu b3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&E)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bchk As Boolean
Private Const HWND_BOTTOM = 1
Private Const HWND_BROADCAST = &HFFFF&
Private Const HWND_DESKTOP = 0
Private Const HWND_NOTOPMOST = -2
Private Const WS_EX_TRANSPARENT = &H20&
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Type RECTL
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Dim cRect As RECT
Const LCR_UNLOCK = 0
Dim dwMouseFlag As Integer
Const ME_LBCLICK = 1
Const ME_LBDBLCLICK = 2
Const ME_RBCLICK = 3
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_ABSOLUTE = &H8000
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_MOVE = &H1
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10
Private Const MOUSETRAILS = 39
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Const SWP_NOACTIVATE = &H10
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Dim lpszCaptionNew As String
Private Const SC_MINIMIZE = &HF020&
Private Const WS_MAXIMIZEBOX = &H10000
Dim HKStateCtrl As Integer
Dim HKStateFn As Integer
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_ICONIC = WS_MINIMIZE
Const SC_ICON = SC_MINIMIZE
Const SC_TASKLIST = &HF130&
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Dim bCodeUse As Boolean
Private Const WS_CAPTION = &HC00000
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * 1024
End Type
Const SC_RESTORE = &HF120&
Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Dim lMeWinStyle As Long
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
Const SWP_NOOWNERZORDER = &H200
Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SC_MOVE = &HF010&
Private Const SC_SIZE = &HF000&
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Const WS_EX_APPWINDOW = &H40000
Private Type WINDOWINFORMATION
hWindow As Long
hWindowDC As Long
hThreadProcess As Long
hThreadProcessID As Long
lpszCaption As String
lpszClassName As String
lpszThreadProcessName As String * 1024
lpszThreadProcessPath As String
lpszExe As String
lpszPath As String
End Type
Private Type WINDOWPARAM
bEnabled As Boolean
bHide As Boolean
bTrans As Boolean
bClosable As Boolean
bSizable As Boolean
bMinisizable As Boolean
bTop As Boolean
lpTransValue As Integer
End Type
Dim lpWindow As WINDOWINFORMATION
Dim lpWindowParam() As WINDOWPARAM
Dim lpCur As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Dim lpRtn As Long
Dim hWindow As Long
Dim lpLength As Long
Dim lpArray() As Byte
Dim lpArray2() As Byte
Dim lpBuff As String
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const LWA_COLORKEY = &H1
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&
Private Const WS_SYSMENU = &H80000
Private Const GWL_STYLE = (-16)
Private Const MF_BYCOMMAND = &H0
Private Const SC_CLOSE = &HF060&
Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Private Const MF_INSERT = &H0&
Private Const SC_MAXIMIZE = &HF030&
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Type WINDOWINFOBOXDATA
lpszCaption As String
lpszClass As String
lpszThread As String
lpszHandle As String
lpszDC As String
End Type
Dim dwWinInfo As WINDOWINFOBOXDATA
Dim bError As Boolean
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Const WM_CLOSE = &H10
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOMOVE = &H2
Dim mov As Boolean
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Const ANYSIZE_ARRAY = 1
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Private Type LUID
UsedPart As Long
IgnoredForNowHigh32BitPart As Long
End Type
Private Type TOKEN_PRIVILEGES
PrivilegeCount As Long
TheLuid As LUID
Attributes As Long
End Type
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal _
ProcessHandle As Long, _
ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" _
Alias "LookupPrivilegeValueA" _
(ByVal lpSystemName As String, ByVal lpName As String, lpLuid _
As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
(ByVal TokenHandle As Long, _
ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES _
, ByVal BufferLength As Long, _
PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Type TestCounter
TimesLeft As Integer
ResetTime As Integer
End Type
Dim PassTest As TestCounter
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Const VK_ADD = &H6B
Private Const VK_ATTN = &HF6
Private Const VK_BACK = &H8
Private Const VK_CANCEL = &H3
Private Const VK_CAPITAL = &H14
Private Const VK_CLEAR = &HC
Private Const VK_CONTROL = &H11
Private Const VK_CRSEL = &HF7
Private Const VK_DECIMAL = &H6E
Private Const VK_DELETE = &H2E
Private Const VK_DIVIDE = &H6F
Private Const VK_DOWN = &H28
Private Const VK_END = &H23
Private Const VK_EREOF = &HF9
Private Const VK_ESCAPE = &H1B
Private Const VK_EXECUTE = &H2B
Private Const VK_EXSEL = &HF8
Private Const VK_F1 = &H70
Private Const VK_F10 = &H79
Private Const VK_F11 = &H7A
Private Const VK_F12 = &H7B
Private Const VK_F13 = &H7C
Private Const VK_F14 = &H7D
Private Const VK_F15 = &H7E
Private Const VK_F16 = &H7F
Private Const VK_F17 = &H80
Private Const VK_F18 = &H81
Private Const VK_F19 = &H82
Private Const VK_F2 = &H71
Private Const VK_F20 = &H83
Private Const VK_F21 = &H84
Private Const VK_F22 = &H85
Private Const VK_F23 = &H86
Private Const VK_F24 = &H87
Private Const VK_F3 = &H72
Private Const VK_F4 = &H73
Private Const VK_F5 = &H74
Private Const VK_F6 = &H75
Private Const VK_F7 = &H76
Private Const VK_F8 = &H77
Private Const VK_F9 = &H78
Private Const VK_HELP = &H2F
Private Const VK_HOME = &H24
Private Const VK_INSERT = &H2D
Private Const VK_LBUTTON = &H1
Private Const VK_LCONTROL = &HA2
Private Const VK_LEFT = &H25
Private Const VK_LMENU = &HA4
Private Const VK_LSHIFT = &HA0
Private Const VK_MBUTTON = &H4
Private Const VK_MENU = &H12
Private Const VK_MULTIPLY = &H6A
Private Const VK_NEXT = &H22
Private Const VK_NONAME = &HFC
Private Const VK_NUMLOCK = &H90
Private Const VK_NUMPAD0 = &H60
Private Const VK_NUMPAD1 = &H61
Private Const VK_NUMPAD2 = &H62
Private Const VK_NUMPAD3 = &H63
Private Const VK_NUMPAD4 = &H64
Private Const VK_NUMPAD5 = &H65
Private Const VK_NUMPAD6 = &H66
Private Const VK_NUMPAD7 = &H67
Private Const VK_NUMPAD8 = &H68
Private Const VK_NUMPAD9 = &H69
Private Const VK_OEM_CLEAR = &HFE
Private Const VK_PA1 = &HFD
Private Const VK_PAUSE = &H13
Private Const VK_PLAY = &HFA
Private Const VK_PRINT = &H2A
Private Const VK_PRIOR = &H21
Private Const VK_PROCESSKEY = &HE5
Private Const VK_RBUTTON = &H2
Private Const VK_RCONTROL = &HA3
Private Const VK_RETURN = &HD
Private Const VK_RIGHT = &H27
Private Const VK_RMENU = &HA5
Private Const VK_RSHIFT = &HA1
Private Const VK_SCROLL = &H91
Private Const VK_SELECT = &H29
Private Const VK_SEPARATOR = &H6C
Private Const VK_SHIFT = &H10
Private Const VK_SNAPSHOT = &H2C
Private Const VK_SPACE = &H20
Private Const VK_SUBTRACT = &H6D
Private Const VK_TAB = &H9
Private Const VK_UP = &H26
Private Const VK_ZOOM = &HFB
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Dim lpX As Long
Dim lpY As Long
Sub GetProcessName(ByVal processID As Long, szExeName As String, szPathName As String)
On Error Resume Next
Dim my As PROCESSENTRY32
Dim hProcessHandle As Long
Dim success As Long
Dim l As Long
l = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If l Then
my.dwSize = 1060
If (Process32First(l, my)) Then
Do
If my.th32ProcessID = processID Then
CloseHandle l
szExeName = Left$(my.szExeFile, InStr(1, my.szExeFile, Chr$(0)) - 1)
For l = Len(szExeName) To 1 Step -1
If Mid$(szExeName, l, 1) = "\" Then
Exit For
End If
Next l
szPathName = Left$(szExeName, l)
Exit Sub
End If
Loop Until (Process32Next(l, my) < 1)
End If
CloseHandle l
End If
End Sub
Private Sub DisableClose(hwnd As Long, Optional ByVal MDIChild As Boolean)
On Error Resume Next
Exit Sub
Dim hSysMenu As Long
Dim nCnt As Long
Dim cID As Long
hSysMenu = GetSystemMenu(hwnd, False)
If hSysMenu = 0 Then
Exit Sub
End If
nCnt = GetMenuItemCount(hSysMenu)
If MDIChild Then
cID = 3
Else
cID = 1
End If
If nCnt Then
RemoveMenu hSysMenu, nCnt - cID, MF_BYPOSITION Or MF_REMOVE
RemoveMenu hSysMenu, nCnt - cID - 1, MF_BYPOSITION Or MF_REMOVE
DrawMenuBar hwnd
End If
End Sub
Private Function GetPassword(hwnd As Long) As String
On Error Resume Next
lpLength = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)
If lpLength > 0 Then
ReDim lpArray(lpLength + 1) As Byte
ReDim lpArray2(lpLength - 1) As Byte
CopyMemory lpArray(0), lpLength, 2
SendMessage hwnd, WM_GETTEXT, lpLength + 1, lpArray(0)
CopyMemory lpArray2(0), lpArray(0), lpLength
GetPassword = StrConv(lpArray2, vbUnicode)
Else
GetPassword = ""
End If
End Function
Private Function GetWindowClassName(hwnd As Long) As String
On Error Resume Next
Dim lpszWindowClassName As String * 256
lpszWindowClassName = Space(256)
GetClassName hwnd, lpszWindowClassName, 256
lpszWindowClassName = Trim(lpszWindowClassName)
GetWindowClassName = lpszWindowClassName
End Function
Private Sub AdjustToken()
On Error Resume Next
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
End Sub
Private Sub Check1_Click()
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
End Sub
Private Sub Check1_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check1_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
With Me.Label3
.Caption = CStr(KeyCode)
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
On Error Resume Next
With Me.Label3
.Caption = CStr(KeyCode)
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
On Error Resume Next
With Me.Label3
.Caption = CStr(KeyCode)
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check1_KeyPress(KeyAscii As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check1_LostFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check1_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check1_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check1_Validate(Cancel As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check2_Click()
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
End Sub
Private Sub Check3_Click()
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
End Sub
Private Sub Check4_Click()
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
End Sub
Private Sub Check5_Click()
On Error Resume Next
Dim ans As Integer
InMsg = True
If Check5.Value = 1 Then
ans = MsgBox("禁用整個鍵盤會導致鍵盤無法響應任何人為動作，請確保您有能正常工作的鼠標、語音控制設備或觸摸輸入設備，繼續嗎？", vbExclamation + vbYesNo, "Info")
If ans = vbNo Then
Check5.Value = 0
Else
Check5.Value = 1
End If
End If
InMsg = False
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
End Sub
Private Sub Check6_Click()
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check7_Click()
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Check7_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Check7_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Check7_GotFocus()
On Error Resume Next
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Check7_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
With Me.Label3
.Caption = CStr(KeyCode)
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
On Error Resume Next
With Me.Label3
.Caption = CStr(KeyCode)
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Check7_KeyPress(KeyAscii As Integer)
On Error Resume Next
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Check7_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Check7_LostFocus()
On Error Resume Next
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Check7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Check7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Check7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Check7_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Check7_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Check7_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Check7_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Check7_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Check7_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Check7_Validate(Cancel As Boolean)
On Error Resume Next
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Command2_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command2_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command2_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
With Me.Label3
.Caption = CStr(KeyCode)
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
On Error Resume Next
With Me.Label3
.Caption = CStr(KeyCode)
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command2_KeyPress(KeyAscii As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command2_LostFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command2_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command2_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command2_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command2_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command2_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command2_Validate(Cancel As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command1_Click()
On Error Resume Next
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Dim nAns As Integer
nAns = MsgBox("確定復位所有設定嗎?", vbQuestion + vbYesNo, "Ask")
Select Case nAns
Case vbYes
Me.Label3.Caption = ""
Me.Command3.Enabled = False
List1.Clear
On Error Resume Next
On Error Resume Next
With Me.mnuExit
.Checked = False
.Enabled = True
End With
With Me.mnuNoAlt
.Checked = False
.Enabled = True
End With
With mnuNoShift
.Checked = False
.Enabled = True
End With
With mnuNoCtrl
.Checked = False
.Enabled = True
End With
With mnuNoWin
.Checked = False
.Enabled = True
End With
With mnuNoKeys
.Checked = False
.Enabled = True
End With
With Me.mnuReset
.Checked = False
.Enabled = True
End With
With Me.Check7
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.Enabled = True
.Value = 0
.Visible = True
End With
With Me.lblTitle
.ForeColor = RGB(0, 0, 255)
.AutoSize = True
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.BackStyle = 0
.BorderStyle = 0
End With
With Me.Label1(0)
.Alignment = 0
.AutoSize = False
.BackColor = RGB(255, 255, 255)
.BackStyle = 0
.BorderStyle = 0
.ForeColor = RGB(0, 0, 0)
End With
With Me.Label1(1)
.Alignment = 0
.AutoSize = False
.BackColor = RGB(255, 255, 255)
.BackStyle = 0
.BorderStyle = 0
.ForeColor = RGB(0, 0, 0)
End With
With Frame1
.BackColor = RGB(255, 255, 255)
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
With Me.Check1
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check2
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check3
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check4
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check5
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check6
.BackColor = &HC0C0C0
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 1
End With
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
Case vbNo
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
Case Else
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Select
End Sub
Private Sub Command2_Click()
On Error Resume Next
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
With Me.cSysTray1
.InTray = True
.TrayTip = "Key Hook Tool - 雙擊還原窗口"
End With
With Me
.Hide
End With
End Sub
Private Sub Command3_Click()
On Error Resume Next
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
If Trim(Label3.Caption) = "" Then
MsgBox "無有效鍵位碼供添加", vbCritical, "Error"
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
Exit Sub
Else
Me.List1.AddItem CStr(CLng(Label3.Caption))
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End If
End Sub
Private Sub Command3_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command3_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command3_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
With Me.Label3
.Caption = CStr(KeyCode)
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command3_KeyPress(KeyAscii As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command3_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command3_LostFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command3_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command3_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command3_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command3_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command3_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command3_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub cSysTray1_MouseDblClick(Button As Integer, Id As Long)
On Error Resume Next
If Button = 1 Then
On Error Resume Next
With Me.cSysTray1
.InTray = False
.TrayTip = "Key Hook Tool - 雙擊還原窗口"
End With
With Me
.Show
End With
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
Else
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
Exit Sub
End If
End Sub
Private Sub cSysTray1_MouseDown(Button As Integer, Id As Long)
On Error Resume Next
If Button = 2 Then
On Error Resume Next
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
PopupMenu Me.TrayMenu
Else
Exit Sub
End If
End Sub
Private Sub cSysTray1_MouseMove(Id As Long)
On Error Resume Next
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub cSysTray1_MouseUp(Button As Integer, Id As Long)
On Error Resume Next
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Form_Activate()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_Click()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_DblClick()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_Deactivate()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_Initialize()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
With Me.Label3
.Caption = CStr(KeyCode)
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
With Me.Label3
.Caption = CStr(KeyCode)
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_LinkClose()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_LinkError(LinkErr As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_LinkOpen(Cancel As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_Load()
On Error Resume Next
HooK
With Me.Check7
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.Enabled = True
.Value = 0
.Visible = True
End With
With Me.mnuExit
.Checked = False
.Enabled = True
End With
With Me.mnuNoAlt
.Checked = False
.Enabled = True
End With
With mnuNoShift
.Checked = False
.Enabled = True
End With
With mnuNoCtrl
.Checked = False
.Enabled = True
End With
With mnuNoWin
.Checked = False
.Enabled = True
End With
With mnuNoKeys
.Checked = False
.Enabled = True
End With
With Me.mnuReset
.Checked = False
.Enabled = True
End With
With Me.lblTitle
.ForeColor = RGB(0, 0, 255)
.AutoSize = True
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.BackStyle = 0
.BorderStyle = 0
End With
With Me.Label1(0)
.Alignment = 0
.AutoSize = False
.BackColor = RGB(255, 255, 255)
.BackStyle = 0
.BorderStyle = 0
.ForeColor = RGB(0, 0, 0)
End With
With Me.Label1(1)
.Alignment = 0
.AutoSize = False
.BackColor = RGB(255, 255, 255)
.BackStyle = 0
.BorderStyle = 0
.ForeColor = RGB(0, 0, 0)
End With
With Frame1
.BackColor = RGB(255, 255, 255)
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
With Me.Check1
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check2
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check3
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check4
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check5
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check6
.BackColor = &HC0C0C0
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 1
End With
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Sub
Private Sub Form_LostFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_Paint()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
With Me.cSysTray1
.InTray = False
.TrayTip = "Key Hook Tool - 雙擊還原窗口"
End With
UnHooK
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_Resize()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_Terminate()
On Error Resume Next
With Me.cSysTray1
.InTray = False
.TrayTip = "Key Hook Tool - 雙擊還原窗口"
End With
UnHooK
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
With Me.cSysTray1
.InTray = False
.TrayTip = "Key Hook Tool - 雙擊還原窗口"
End With
UnHooK
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Frame1_Click()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Frame1_DblClick()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Image1_Click()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Image1_DblClick()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Image1_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Image1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Image1_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Image1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Image1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Image1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Image1_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Image1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label1_Change(Index As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label1_Click(Index As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label1_DblClick(Index As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label1_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label1_LinkClose(Index As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label1_LinkError(Index As Integer, LinkErr As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label1_LinkNotify(Index As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label1_LinkOpen(Index As Integer, Cancel As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label1_OLECompleteDrag(Index As Integer, Effect As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label1_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label1_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label1_OLEGiveFeedback(Index As Integer, Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label1_OLESetData(Index As Integer, Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label1_OLEStartDrag(Index As Integer, Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label2_Change()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label2_Click()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label2_DblClick()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label2_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label2_LinkClose()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label2_LinkError(LinkErr As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label2_LinkNotify()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label2_LinkOpen(Cancel As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label2_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label2_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label2_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label2_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label2_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label3_Change()
On Error Resume Next
If Trim(Label3.Caption) = "" Then
Me.Command3.Enabled = False
Else
Command3.Enabled = True
End If
End Sub
Private Sub Label3_Click()
On Error Resume Next
If Label3.Caption = "" Then
Command3.Enabled = False
Exit Sub
End If
Command3.Enabled = True
With Me.List1
.AddItem CStr(CLng(Label3.Caption))
End With
End Sub
Private Sub Label3_DblClick()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label3_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label3_LinkClose()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label3_LinkError(LinkErr As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label3_LinkNotify()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label3_LinkOpen(Cancel As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label3_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label3_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label3_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label3_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label3_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label3_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label4_Change()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label4_Click()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label4_DblClick()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label4_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label4_LinkClose()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label4_LinkError(LinkErr As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label4_LinkNotify()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label4_LinkOpen(Cancel As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label4_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label4_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label4_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label4_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label4_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Label4_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub lblTitle_Change()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub lblTitle_Click()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub lblTitle_DblClick()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub lblTitle_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub lblTitle_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub lblTitle_LinkClose()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub lblTitle_LinkError(LinkErr As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub lblTitle_LinkNotify()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub lblTitle_LinkOpen(Cancel As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub lblTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub lblTitle_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub lblTitle_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub lblTitle_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub lblTitle_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub lblTitle_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub lblTitle_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_Click()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_DblClick()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_ItemCheck(Item As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
With Me.Label3
.Caption = CStr(KeyCode)
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_KeyPress(KeyAscii As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_LostFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_Scroll()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub List1_Validate(Cancel As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub mnuExit_Click()
On Error Resume Next
UnHooK
On Error Resume Next
With Me.cSysTray1
.InTray = False
.TrayTip = "Key Hook Tool - 雙擊還原窗口"
End With
UnHooK
Unload Me
End
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub mnuNoAlt_Click()
On Error Resume Next
bchk = mnuNoAlt.Checked
Select Case bchk
Case True
With Me.Check4
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
bchk = Not bchk
mnuNoAlt.Checked = bchk
Case False
With Me.Check4
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 1
End With
bchk = Not bchk
mnuNoAlt.Checked = bchk
End Select
If 1 = 2 Then
With Me.Check4
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = CInt(bchk)
End With
End If
End Sub
Private Sub mnuNoCtrl_Click()
On Error Resume Next
bchk = mnuNoCtrl.Checked
Select Case bchk
Case True
With Me.Check2
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
mnuNoCtrl.Checked = False
Case False
With Me.Check2
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 1
End With
mnuNoCtrl.Checked = True
End Select
If 1 = 2 Then
With Me.Check2
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = CInt(bchk)
End With
End If
End Sub
Private Sub mnuNoFx_Click()
On Error Resume Next
bchk = mnuNoFx.Checked
Select Case bchk
Case True
With Me.Check7
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
bchk = Not bchk
mnuNoFx.Checked = bchk
Case False
With Me.Check7
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 1
End With
bchk = Not bchk
mnuNoFx.Checked = bchk
End Select
If 1 = 2 Then
With Me.Check7
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = CInt(bchk)
End With
End If
End Sub
Private Sub mnuNoKeys_Click()
On Error Resume Next
Dim ans As Integer
bchk = mnuNoKeys.Checked
Select Case bchk
Case True
With Me.Check5
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
bchk = Not bchk
mnuNoKeys.Checked = bchk
Me.mnuNoKeys.Checked = CBool(Check5.Value)
Case False
With Me.Check5
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 1
End With
bchk = Not bchk
mnuNoKeys.Checked = bchk
Me.mnuNoKeys.Checked = CBool(Check5.Value)
End Select
If 1 = 2 Then
With Me.Check5
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = CInt(bchk)
End With
End If
End Sub
Private Sub mnuNoShift_Click()
On Error Resume Next
bchk = mnuNoShift.Checked
Select Case bchk
Case True
With Me.Check3
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
mnuNoShift.Checked = False
Case False
With Me.Check3
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 1
End With
mnuNoShift.Checked = True
End Select
If 1 = 2 Then
With Me.Check3
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = CInt(bchk)
End With
End If
End Sub
Private Sub mnuNoWin_Click()
On Error Resume Next
bchk = mnuNoWin.Checked
Select Case bchk
Case True
With Me.Check1
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
mnuNoWin.Checked = False
Case False
With Me.Check1
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 1
End With
mnuNoWin.Checked = True
End Select
If 1 = 2 Then
With Me.Check1
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = CInt(bchk)
End With
End If
End Sub
Private Sub mnuReset_Click()
If 1 = 245 Then
On Error Resume Next
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Dim nAns As Integer
nAns = MsgBox("確定復位所有設定嗎?", vbQuestion + vbYesNo, "Ask")
Select Case nAns
Case vbYes
Me.Label3.Caption = ""
Me.Command3.Enabled = False
List1.Clear
On Error Resume Next
On Error Resume Next
With Me.mnuExit
.Checked = False
.Enabled = True
End With
With Me.mnuNoAlt
.Checked = False
.Enabled = True
End With
With mnuNoShift
.Checked = False
.Enabled = True
End With
With mnuNoCtrl
.Checked = False
.Enabled = True
End With
With mnuNoWin
.Checked = False
.Enabled = True
End With
With mnuNoKeys
.Checked = False
.Enabled = True
End With
With Me.mnuReset
.Checked = False
.Enabled = True
End With
With Me.lblTitle
.ForeColor = RGB(0, 0, 255)
.AutoSize = True
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.BackStyle = 0
.BorderStyle = 0
End With
With Me.Label1(0)
.Alignment = 0
.AutoSize = False
.BackColor = RGB(255, 255, 255)
.BackStyle = 0
.BorderStyle = 0
.ForeColor = RGB(0, 0, 0)
End With
With Me.Label1(1)
.Alignment = 0
.AutoSize = False
.BackColor = RGB(255, 255, 255)
.BackStyle = 0
.BorderStyle = 0
.ForeColor = RGB(0, 0, 0)
End With
With Frame1
.BackColor = RGB(255, 255, 255)
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
With Me.Check1
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check2
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check3
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check4
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check5
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check6
.BackColor = &HC0C0C0
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 1
End With
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
Case vbNo
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
Case Else
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Select
End If
On Error Resume Next
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
nAns = MsgBox("確定復位所有設定嗎?", vbQuestion + vbYesNo, "Ask")
Select Case nAns
Case vbYes
On Error Resume Next
On Error Resume Next
With Me.mnuExit
.Checked = False
.Enabled = True
End With
With Me.mnuNoAlt
.Checked = False
.Enabled = True
End With
With mnuNoShift
.Checked = False
.Enabled = True
End With
With mnuNoCtrl
.Checked = False
.Enabled = True
End With
With mnuNoWin
.Checked = False
.Enabled = True
End With
With mnuNoKeys
.Checked = False
.Enabled = True
End With
With Me.mnuReset
.Checked = False
.Enabled = True
End With
With Me.lblTitle
.ForeColor = RGB(0, 0, 255)
.AutoSize = True
.Alignment = 0
.BackColor = RGB(255, 255, 255)
.BackStyle = 0
.BorderStyle = 0
End With
With Me.Label1(0)
.Alignment = 0
.AutoSize = False
.BackColor = RGB(255, 255, 255)
.BackStyle = 0
.BorderStyle = 0
.ForeColor = RGB(0, 0, 0)
End With
With Me.Label1(1)
.Alignment = 0
.AutoSize = False
.BackColor = RGB(255, 255, 255)
.BackStyle = 0
.BorderStyle = 0
.ForeColor = RGB(0, 0, 0)
End With
With Frame1
.BackColor = RGB(255, 255, 255)
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
With Me.Check1
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check2
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check3
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check4
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check5
.BackColor = RGB(255, 255, 255)
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 0
End With
With Me.Check6
.BackColor = &HC0C0C0
.Alignment = 0
.ForeColor = RGB(0, 0, 0)
.Visible = True
.Enabled = True
.Value = 1
End With
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
Case vbNo
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
Case Else
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check7.Value = 0 Then
Me.mnuNoFx.Checked = False
End If
If Check7.Value = 1 Then
Me.mnuNoFx.Checked = True
End If
Me.mnuNoFx.Checked = CBool(Check7.Value)
End Select
End Sub
Private Sub mnuShow_Click()
On Error Resume Next
Dim Button As Integer
Button = 1
If Button = 1 Then
On Error Resume Next
With Me.cSysTray1
.InTray = False
.TrayTip = "Key Hook Tool - 雙擊還原窗口"
End With
With Me
.Show
End With
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
Else
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
Exit Sub
End If
End Sub
Private Sub Check2_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check2_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check2_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
With Me.Label3
.Caption = CStr(KeyCode)
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check2_KeyPress(KeyAscii As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check2_LostFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check2_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check2_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check2_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check2_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check2_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check2_Validate(Cancel As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check3_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check3_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check3_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check3_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
With Me.Label3
.Caption = CStr(KeyCode)
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check3_KeyPress(KeyAscii As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check3_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check3_LostFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check3_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check3_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check3_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check3_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check3_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check3_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check3_Validate(Cancel As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check4_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check4_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check4_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check4_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
With Me.Label3
.Caption = CStr(KeyCode)
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check4_KeyPress(KeyAscii As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check4_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check4_LostFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check4_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check4_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check4_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check4_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check4_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check4_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check4_Validate(Cancel As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check5_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check5_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check5_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check5_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
With Me.Label3
.Caption = CStr(KeyCode)
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check5_KeyPress(KeyAscii As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check5_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check5_LostFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check5_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check5_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check5_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check5_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check5_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check5_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check5_Validate(Cancel As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check6_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check6_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check6_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check6_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
With Me.Label3
.Caption = CStr(KeyCode)
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check6_KeyPress(KeyAscii As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check6_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check6_LostFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check6_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check6_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check6_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check6_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check6_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check6_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Check6_Validate(Cancel As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command1_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command1_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
With Me.Label3
.Caption = CStr(KeyCode)
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command1_KeyPress(KeyAscii As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command1_LostFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command1_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command1_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Command1_Validate(Cancel As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Frame1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Frame1_GotFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Frame1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
With Me.Label3
.Caption = CStr(KeyCode)
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
On Error Resume Next
With Me.Label3
.Caption = CStr(KeyCode)
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BackStyle = 1
.BorderStyle = 1
.ForeColor = RGB(0, 0, 0)
End With
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Frame1_KeyPress(KeyAscii As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Frame1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Frame1_LostFocus()
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Frame1_OLECompleteDrag(Effect As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Frame1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Frame1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Frame1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Frame1_OLESetData(Data As DataObject, DataFormat As Integer)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Frame1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
Private Sub Frame1_Validate(Cancel As Boolean)
On Error Resume Next
Exit Sub
Exit Sub
On Error Resume Next
If Check1.Value = 0 Then
Me.mnuNoWin.Checked = False
End If
If Check1.Value = 1 Then
Me.mnuNoWin.Checked = True
End If
Me.mnuNoWin.Checked = CBool(Check1.Value)
On Error Resume Next
If Check2.Value = 0 Then
Me.mnuNoCtrl.Checked = False
End If
If Check2.Value = 1 Then
Me.mnuNoCtrl.Checked = True
End If
Me.mnuNoCtrl.Checked = CBool(Check2.Value)
On Error Resume Next
If Check3.Value = 0 Then
Me.mnuNoShift.Checked = False
End If
If Check3.Value = 1 Then
Me.mnuNoShift.Checked = True
End If
Me.mnuNoShift.Checked = CBool(Check3.Value)
On Error Resume Next
If Check4.Value = 0 Then
Me.mnuNoAlt.Checked = False
End If
If Check4.Value = 1 Then
Me.mnuNoAlt.Checked = True
End If
Me.mnuNoAlt.Checked = CBool(Check4.Value)
On Error Resume Next
If Check5.Value = 0 Then
Me.mnuNoKeys.Checked = False
End If
If Check5.Value = 1 Then
Me.mnuNoKeys.Checked = True
End If
Me.mnuNoKeys.Checked = CBool(Check5.Value)
On Error Resume Next
With Check6
If .Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If .Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End With
End Sub
