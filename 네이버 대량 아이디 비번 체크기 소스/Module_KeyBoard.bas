Attribute VB_Name = "mdKeyBoard"




'** 운영 체제 정보를 알기 위한 API
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVerInfo) As Long

Type OSVerInfo
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type


'//한글 토글키 부분
Public Const IME_CMODE_NATIVE = &H1
Public Const IME_CMODE_HANGEUL = IME_CMODE_NATIVE
Public Const IME_CMODE_ALPHANUMERIC = &H0
Public Const IME_SMODE_NONE = &H0
Declare Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As Long) As Long
Declare Function ImmSetConversionStatus Lib "imm32.dll" _
(ByVal hIMC As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
'테스트
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
'/키보드 후킹부분
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Const KEYEVENTF_KEYUP = &H2
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long

'/한글조합위한 후킹 api함수 부분

Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Const WH_KEYBOARD = 2
Global hHook As Long

' Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
'현재 훅 Chain상의 다음 훅 프로시저에게 정보를 전달
'Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
'후킹하는 부분
'Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
'메세지 정리
'Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
'Public Const WH_KEYBOARD = 2
'Global HWND_HOOK As Long
'Public Const HC_ACTION = 0

Public Const MSH_MOUSEWHEEL = "MSWHEEL_ROLLMSG"
Public Declare Function RegisterWindowMessage& Lib "user32" Alias _
 "RegisterWindowMessageA" (ByVal lpString As String)
Public IMWHEEL_MSG As Long

'/가상키값
Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const VK_CANCEL = &H3
Public Const VK_MBUTTON = &H4
Public Const VK_BACK = &H8
Public Const VK_TAB = &H9
Public Const VK_CLEAR = &HC
Public Const VK_RETURN = &HD
Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11
Public Const VK_MENU = &H12
Public Const VK_PAUSE = &H13
Public Const VK_CAPITAL = &H14
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_SELECT = &H29
Public Const VK_PRINT = &H2A
Public Const VK_EXECUTE = &H2B
Public Const VK_SNAPSHOT = &H2C
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E
Public Const VK_HELP = &H2F
Public Const VK_0 = &H30
Public Const VK_1 = &H31
Public Const VK_2 = &H32
Public Const VK_3 = &H33
Public Const VK_4 = &H34
Public Const VK_5 = &H35
Public Const VK_6 = &H36
Public Const VK_7 = &H37
Public Const VK_8 = &H38
Public Const VK_9 = &H39
Public Const VK_A = &H41
Public Const VK_B = &H42
Public Const VK_C = &H43
Public Const VK_D = &H44
Public Const VK_E = &H45
Public Const VK_F = &H46
Public Const VK_G = &H47
Public Const VK_H = &H48
Public Const VK_I = &H49
Public Const VK_J = &H4A
Public Const VK_K = &H4B
Public Const VK_L = &H4C
Public Const VK_M = &H4D
Public Const VK_N = &H4E
Public Const VK_O = &H4F
Public Const VK_P = &H50
Public Const VK_Q = &H51
Public Const VK_R = &H52
Public Const VK_S = &H53
Public Const VK_T = &H54
Public Const VK_U = &H55
Public Const VK_V = &H56
Public Const VK_W = &H57
Public Const VK_X = &H58
Public Const VK_Y = &H59
Public Const VK_Z = &H5A
Public Const VK_STARTKEY = &H5B
Public Const VK_CONTEXTKEY = &H5D
Public Const VK_NUMPAD0 = &H60
Public Const VK_NUMPAD1 = &H61
Public Const VK_NUMPAD2 = &H62
Public Const VK_NUMPAD3 = &H63
Public Const VK_NUMPAD4 = &H64
Public Const VK_NUMPAD5 = &H65
Public Const VK_NUMPAD6 = &H66
Public Const VK_NUMPAD7 = &H67
Public Const VK_NUMPAD8 = &H68
Public Const VK_NUMPAD9 = &H69
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const VK_F13 = &H7C
Public Const VK_F14 = &H7D
Public Const VK_F15 = &H7E
Public Const VK_F16 = &H7F
Public Const VK_F17 = &H80
Public Const VK_F18 = &H81
Public Const VK_F19 = &H82
Public Const VK_F20 = &H83
Public Const VK_F21 = &H84
Public Const VK_F22 = &H85
Public Const VK_F23 = &H86
Public Const VK_F24 = &H87
Public Const VK_NUMLOCK = &H90
Public Const VK_OEM_SCROLL = &H91
Public Const VK_OEM_1 = &HBA
Public Const VK_OEM_PLUS = &HBB
Public Const VK_OEM_COMMA = &HBC
Public Const VK_OEM_MINUS = &HBD
Public Const VK_OEM_PERIOD = &HBE
Public Const VK_OEM_2 = &HBF
Public Const VK_OEM_3 = &HC0
Public Const VK_OEM_4 = &HDB
Public Const VK_OEM_5 = &HDC
Public Const VK_OEM_6 = &HDD
Public Const VK_OEM_7 = &HDE
Public Const VK_OEM_8 = &HDF
Public Const VK_ICO_F17 = &HE0
Public Const VK_ICO_F18 = &HE1
Public Const VK_OEM102 = &HE2
Public Const VK_ICO_HELP = &HE3
Public Const VK_ICO_00 = &HE4
Public Const VK_ICO_CLEAR = &HE6
Public Const VK_OEM_RESET = &HE9
Public Const VK_OEM_JUMP = &HEA
Public Const VK_OEM_PA1 = &HEB
Public Const VK_OEM_PA2 = &HEC
Public Const VK_OEM_PA3 = &HED
Public Const VK_OEM_WSCTRL = &HEE
Public Const VK_OEM_CUSEL = &HEF
Public Const VK_OEM_ATTN = &HF0
Public Const VK_OEM_FINNISH = &HF1
Public Const VK_OEM_COPY = &HF2
Public Const VK_OEM_AUTO = &HF3
Public Const VK_OEM_ENLW = &HF4
Public Const VK_OEM_BACKTAB = &HF5
Public Const VK_ATTN = &HF6
Public Const VK_CRSEL = &HF7
Public Const VK_EXSEL = &HF8
Public Const VK_EREOF = &HF9
Public Const VK_PLAY = &HFA
Public Const VK_ZOOM = &HFB
Public Const VK_NONAME = &HFC
Public Const VK_PA1 = &HFD
Public Const VK_OEM_CLEAR = &HFE
Public key_use As Integer '포커스 조절 0이면 web에 1이면 text컨트롤에 포커스
Public Language_Set As Integer '1한글 2영어
Dim a As Long
Dim b As Long
Public Url As String
Public GsKeyBoardResult As String

Public Function SKeyDown(virtualkey As Byte)
    keybd_event virtualkey, MapVirtualKey(virtualkey, 0), 0, 0
End Function
    
Public Function SKeyUp(virtualkey As Byte)
    'KEYBD_EVENT 전달 인수 값 ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long
    '                               가상키값,    스캔코드값, flags specifying various function options, additional data associated with keystroke
    'MAPVIRTUALKEY 전달인수 값 ByVal wCode As Long, ByVal wMapType As Long) As Long -스캔 코드로 변환
    keybd_event virtualkey, MapVirtualKey(virtualkey, 0), KEYEVENTF_KEYUP, 0
    ' keybd_event virtualkey, MapVirtualKey(virtualkey, 0), KEYEVENTF_EXTENDEDKEY, 0
End Function

Public Function FKeyboardProc(ByVal idHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim iReturn As Integer

    If idHook < 0 Then
        FKeyboardProc = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
    Else
    
        iReturn = fiOperatingVersion
        
        Select Case iReturn
            Case 1
                If (wParam = 229 And lParam = -2147483647) Then
                    FKeyboardProc = 1
                        Exit Function
                    End If
            Case 2
                If (wParam = 229 And lParam = -2147483648#) Then
                    FKeyboardProc = 1
                    Exit Function
                End If
        End Select
        
        FKeyboardProc = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
    End If

End Function

Public Function fbKeyBoardShow(cControl As Control, Optional Language As String = "한글") As Boolean
'**************************************************************************************************
'목           적 : KeyBoard를 실행시키고 그 결과 값을 전달한다.
'Input Parameter : cControl - 결과 값을 전달할 Control
'Return Value    : String
'작  성   일  자 : 2002년 1월 24일
'작     성    자 : 박 대 영
'**************************************************************************************************
On Error GoTo fbKeyBoardShow_Err

    Language_Set = IIf(Language = "한글", 1, 2)
         
    CKeyBoard.Show vbModal
    If TypeOf cControl Is TextBox Then
        If Len(Trim(GsKeyBoardResult)) > 0 Then
            cControl.Text = GsKeyBoardResult
        End If
    End If
    
    fbKeyBoardShow = True
    Exit Function
fbKeyBoardShow_Err:
    fbKeyBoardShow = False
End Function






Public Function fiOperatingVersion() As Integer
'**************************************************************************************************
'목           적 : 현재 Operation System을 체크하여 운영체제에 따른 유연성을 기른다.
'Input Parameter : .
'Return Value    : 0 : Windows 32, 1 : Windows 95, 2 : Windows NT(2000)
'작  성   일  자 : 2001년 12월 14일
'작     성    자 : 박대영
'**************************************************************************************************
Dim iReturn     As Integer
Dim sDosVersion As String
Dim sWinVersion As String
Dim sMajor      As String
Dim sMinor      As String
Dim sBuild      As String
Dim VerInfo     As OSVerInfo
      
    ' Get operating system and version.
    VerInfo.dwOSVersionInfoSize = Len(VerInfo)
    iReturn = GetVersionEx(VerInfo)

    If iReturn = 0 Then
        Exit Function
    End If
      
    fiOperatingVersion = VerInfo.dwPlatformId
  
End Function

Public Function fsOperatingVersion() As String
Dim iReturn     As Integer
    
    iReturn = fiOperatingVersion()
    Select Case iReturn
        Case 0
            fsOperatingVersion = "Windows 32s "
        Case 1
            fsOperatingVersion = "Windows 95/98 "
        Case 2
            fsOperatingVersion = "Windows NT/2000 "
        Case 3
            fsOperatingVersion = "Windows XP "
        Case Else
            fsOperatingVersion = "Microsoft Windows Flatform "
    End Select
  
End Function
 

