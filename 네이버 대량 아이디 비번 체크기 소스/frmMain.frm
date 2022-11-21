VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   BorderStyle     =   1  '단일 고정
   Caption         =   "VPN Proxy IP Server Settings"
   ClientHeight    =   2145
   ClientLeft      =   12645
   ClientTop       =   7125
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4365
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6360
      Top             =   3960
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5400
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin Project1.CandyButton cmdProxyOn 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "적용"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Project1.CandyButton cmdProxyOff 
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "해제"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "프록시 찾기"
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.ComboBox cbProxyServer 
         Height          =   300
         Left            =   1440
         Style           =   2  '드롭다운 목록
         TabIndex        =   3
         Top             =   720
         Width           =   2775
      End
      Begin VB.ComboBox cbCountry 
         Height          =   300
         Left            =   1440
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
      Begin Project1.CandyButton cmdSearch 
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Search"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "국적선택 :"
         Height          =   180
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "우회할 프록시 :"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   60
      End
   End
   Begin VB.Label Label4 
      Height          =   135
      Left            =   4680
      TabIndex        =   10
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   5040
      Width           =   4575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Const HKEY_CURRENT_USER = &H80000001
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Private Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByVal lBufferLength As Long) As Integer
Const INTERNET_OPTION_SETTINGS_CHANGED As Long = 39
Const INTERNET_OPTION_REFRESH As Long = 37
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim WinHttp As New WinHttpRequest
Dim Value() As String

Function ProxySetting(ByVal ProxyServer As String, ByVal ProxyEnable As Boolean)
    Dim ret As Long
    
    RegCreateKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", ret
    RegSetValueEx ret, "ProxyServer", 0, REG_SZ, ByVal ProxyServer, Len(ProxyServer)
    
    If ProxyEnable = True Then
        RegSetValueEx ret, "ProxyEnable", 0, REG_DWORD, Abs(CLng(ProxyEnable)), 0
    Else
        RegSetValueEx ret, "ProxyEnable", 0, REG_DWORD, Abs(CLng(ProxyEnable)), 5
    End If

    Call InternetSetOption(0, INTERNET_OPTION_SETTINGS_CHANGED, 0, 0)
    Call InternetSetOption(0, INTERNET_OPTION_REFRESH, 0, 0)
End Function

Private Sub CandyButton1_Click()
Form1.Show
End Sub

Private Sub cmdProxyOff_Click()
    cbCountry.Enabled = Not cbCountry.Enabled
    cbProxyServer.Enabled = Not cbProxyServer.Enabled
    cmdSearch.Enabled = Not cmdSearch.Enabled

    cmdProxyOn.Enabled = Not cmdProxyOn.Enabled
    cmdProxyOff.Enabled = Not cmdProxyOff.Enabled
    Call ProxySetting("127.0.0.1:80", False)

End Sub

Private Sub cmdProxyOn_Click()
    cbCountry.Enabled = Not cbCountry.Enabled
    cbProxyServer.Enabled = Not cbProxyServer.Enabled
    cmdSearch.Enabled = Not cmdSearch.Enabled

    cmdProxyOn.Enabled = Not cmdProxyOn.Enabled
   ' cmdProxyOff.Enabled = Not cmdProxyOff.Enabled
    Call ProxySetting(cbProxyServer, True)
    MsgBox "적용완료 되었습니다", vbInformation, "안내"
End Sub

Private Sub cmdSearch_Click()
    Dim Temp() As String
    Dim ProxyServer As String
    
    cbProxyServer.Clear
    cbCountry.Enabled = Not cbCountry.Enabled
    cbProxyServer.Enabled = Not cbProxyServer.Enabled
    cmdSearch.Enabled = Not cmdSearch.Enabled
    lblStatus.Caption = "프록시 서버를 찾고있습니다."

    WinHttp.Open "GET", "http://www.xroxy.com/proxylist.php?country=" & Value(cbCountry.ListIndex + 1), True
    WinHttp.Send
    WinHttp.WaitForResponse
    
    Temp = Split(WinHttp.ResponseText, "host=")

    For i = 1 To UBound(Temp)
        ProxyServer = Split(Temp(i), "&")(0) & ":" & Split(Split(Temp(i), "&port=")(1), "&")(0)

        If ProxyTester(ProxyServer) = True Then
            cbProxyServer.AddItem ProxyServer
        End If
    Next i
    
    If cbProxyServer.ListCount = 0 Then
        MsgBox "프록시 서버가 발견되지않았습니다." & vbCrLf & "다른 국적을 선택해주세요.", vbExclamation, "이용안내"
    Else
        MsgBox cbProxyServer.ListCount & "개의 프록시서버가 발견되었습니다.", vbInformation, "이용안내"
        'cmdProxyOn.Enabled = Not cmdProxyOn.Enabled
    End If
    
    lblStatus.Caption = vbNullString
    cbCountry.Enabled = Not cbCountry.Enabled
    cbProxyServer.Enabled = Not cbProxyServer.Enabled
    cmdSearch.Enabled = Not cmdSearch.Enabled
End Sub
Public Sub IPLoad()
    Dim Temp As String
    Dim Country() As String
    Dim a As String

WinHttp.Open "GET", "http://ipip.kr"
WinHttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0)"
WinHttp.Send

a = Split(Split(StrConv(WinHttp.ResponseBody, vbUnicode), "</span><span style=""font-size:11pt;"">")(1), "</span>")(0)

    WinHttp.Open "GET", "http://www.xroxy.com/proxylist.php", True
    WinHttp.Send
    WinHttp.WaitForResponse
    
    Temp = Split(Split(WinHttp.ResponseText, "<option selected='selected' value=''>Any country</option>")(1), "</select>")(0)
    Country = Split(Temp, "<option value='")
    
    ReDim Value(UBound(Country)) As String
    
    For i = 1 To UBound(Country)
        Value(i) = Split(Country(i), "'>")(0)
        cbCountry.AddItem Split(Split(Country(i), "'>")(1), "</option>")(0)
    Next i
Label3.Caption = "UdpIP : " & Winsock1.LocalIP & ""
Label3.Caption = Label3.Caption & " WebIP:" & a & "입니다."
End Sub

Private Sub Command1_Click()
Unload frmMain
End Sub

Private Sub Form_Load()
IPLoad
Dim WinHttp As Object '한번더 참조해요
Set WinHttp = CreateObject("Winhttp.WinHttpRequest.5.1")
WinHttp.Open "GET", "http://thoyeon.dothome.co.kr/CafeHelper.txt" '자기의 FTP주소.gm.txt를 써주고 FTP주소에 gm.txt를 만들고 ON 이라 적으면 서버ON되고  OFF라 적으면 OFF가 된다.
WinHttp.Send '그정보를 보냄
Label2.Caption = StrConv(WinHttp.ResponseBody, vbUnicode) '레이블글자에 나오게 해요
If Label2.Caption = "ON" Then '레이블이 ON이라 되있다면요
Else 'ON이 아닌 다른 모든 경우라면
MsgBox "프로그램이 새로 업데이트되었습니다.개발자에게 문의바랍니다.", vbCritical, "hollywoodst4r@nate.com"
Unload Me ' 나를닫는다
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call ProxySetting("127.0.0.1:80", False)
    End
End Sub

Private Sub Timer1_Timer()
Dim WinHttp As Object '한번더 참조해요
Set WinHttp = CreateObject("Winhttp.WinHttpRequest.5.1")
WinHttp.Open "GET", "http://thoyeon.dothome.co.kr/CafeHelper.txt" '자기의 FTP주소.gm.txt를 써주고 FTP주소에 gm.txt를 만들고 ON 이라 적으면 서버ON되고  OFF라 적으면 OFF가 된다.
WinHttp.Send '그정보를 보냄
Label2.Caption = StrConv(WinHttp.ResponseBody, vbUnicode) '레이블글자에 나오게 해요
If Label2.Caption = "ON" Then '레이블이 ON이라 되있다면요
Else 'ON이 아닌 다른 모든 경우라면
MsgBox "프로그램이 새로 업데이트되었습니다.개발자에게 문의바랍니다.", vbCritical, "hollywoodst4r@nate.com"
Unload Me ' 나를닫는다
End If
End Sub
