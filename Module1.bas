Attribute VB_Name = "Module1"
Option Explicit
Public InMsg As Boolean
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Type PKBDLLHOOKSTRUCT
vkCode As Long
scanCode As Long
flags As Long
time As Long
dwExtraInfo As Long
End Type
Private Const VK_LWIN = &H5B
Public Const VK_TAB = &H9
Private Const WM_KEYDOWN = &H100
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_KEYUP = &H101
Private Const WM_SYSKEYUP = &H105
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
Private Const VK_RWIN = &H5C
Private Const VK_UP = &H26
Private Const VK_ZOOM = &HFB
Private Const HC_ACTION = 0
Private Const WH_KEYBOARD_LL = 13
Private lngHook As Long
Public Function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If InMsg = True Then
Exit Function
End If
Dim blnHook As Boolean
Dim p As PKBDLLHOOKSTRUCT
If nCode = HC_ACTION Then
Select Case wParam
Case WM_KEYDOWN, WM_SYSKEYDOWN, WM_KEYUP, WM_SYSKEYUP
Call CopyMemory(p, ByVal lParam, Len(p))
If 1 = 245 Then
If p.vkCode = VK_LWIN Or p.vkCode = VK_TAB Or p.vkCode = VK_RWIN Or p.vkCode = VK_CONTROL Or p.vkCode = VK_MENU Or p.vkCode = VK_ESCAPE Or p.vkCode = VK_DELETE Then
blnHook = True
End If
Select Case p.vkCode
Case VK_LWIN
blnHook = True
Case VK_RWIN
blnHook = True
Case VK_CONTROL
blnHook = True
Case VK_MENU
blnHook = True
Case VK_F1
blnHook = True
Case VK_F2
blnHook = True
Case VK_F3
blnHook = True
Case VK_F4
blnHook = True
Case VK_F5
blnHook = True
Case VK_F6
blnHook = True
Case VK_F7
blnHook = True
Case VK_F8
blnHook = True
Case VK_F10
blnHook = True
Case VK_F11
blnHook = True
Case VK_F12
blnHook = True
Case VK_F9
blnHook = True
Case VK_LMENU
blnHook = True
Case VK_ESCAPE
blnHook = True
Case VK_DELETE
blnHook = True
End Select
End If
If Form1.Check1.Value = 1 Then
Select Case p.vkCode
Case VK_LWIN
blnHook = True
Case VK_RWIN
blnHook = True
End Select
End If
If Form1.Check2.Value = 1 Then
If p.vkCode = VK_LCONTROL Then
blnHook = True
End If
If p.vkCode = VK_RCONTROL Then
blnHook = True
End If
End If
If Form1.Check3.Value = 1 Then
Select Case p.vkCode
Case VK_LSHIFT
blnHook = True
Case VK_RSHIFT
blnHook = True
End Select
End If
If Form1.Check4.Value = 1 Then
If p.vkCode = VK_LMENU Then
blnHook = True
End If
If p.vkCode = VK_RMENU Then
blnHook = True
End If
End If
If Form1.Check7.Value = 1 Then
Select Case p.vkCode
Case VK_F1
blnHook = True
Case VK_F11
blnHook = True
Case VK_F12
blnHook = True
Case VK_F2
blnHook = True
Case VK_F3
blnHook = True
Case VK_F4
blnHook = True
Case VK_F5
blnHook = True
Case VK_F6
blnHook = True
Case VK_F7
blnHook = True
Case VK_F8
blnHook = True
Case VK_F9
blnHook = True
End Select
End If
If Form1.Check5.Value = 1 Then
blnHook = True
End If
Dim nIndexNum As Long
nIndexNum = Form1.List1.ListCount
If nIndexNum = 0 Then
'Nothing But An Empty If Body
End If
If nIndexNum = 1 Then
If p.vkCode = Form1.List1.List(0) Then
blnHook = True
End If
End If
If nIndexNum > 1 Then
Dim nLoop As Long
For nLoop = 0 To nIndexNum - 1
If p.vkCode = Form1.List1.List(nLoop) Then
blnHook = True
End If
Next
End If
Case Else
Call CopyMemory(p, ByVal lParam, Len(p))
If 1 = 245 Then
If p.vkCode = VK_LWIN Or p.vkCode = VK_TAB Or p.vkCode = VK_RWIN Or p.vkCode = VK_CONTROL Or p.vkCode = VK_MENU Or p.vkCode = VK_ESCAPE Or p.vkCode = VK_DELETE Then
blnHook = True
End If
Select Case p.vkCode
Case VK_LWIN
blnHook = True
Case VK_RWIN
blnHook = True
Case VK_CONTROL
blnHook = True
Case VK_MENU
blnHook = True
Case VK_F1
blnHook = True
Case VK_F2
blnHook = True
Case VK_F3
blnHook = True
Case VK_F4
blnHook = True
Case VK_F5
blnHook = True
Case VK_F6
blnHook = True
Case VK_F7
blnHook = True
Case VK_F8
blnHook = True
Case VK_F10
blnHook = True
Case VK_F11
blnHook = True
Case VK_F12
blnHook = True
Case VK_F9
blnHook = True
Case VK_LMENU
blnHook = True
Case VK_ESCAPE
blnHook = True
Case VK_DELETE
blnHook = True
End Select
End If
End Select
End If
If blnHook Then
LowLevelKeyboardProc = 1
Else
Call CallNextHookEx(WH_KEYBOARD_LL, nCode, wParam, lParam)
End If
End Function
Public Sub HooK()
lngHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)
End Sub
Public Sub UnHooK()
Call UnhookWindowsHookEx(lngHook)
End Sub

