VERSION 5.00
Begin VB.UserControl CubeSkin 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   ClientHeight    =   3420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5820
   ScaleHeight     =   3420
   ScaleWidth      =   5820
   Begin VB.PictureBox bUnload 
      AutoSize        =   -1  'True
      BorderStyle     =   0  '없음
      Height          =   195
      Left            =   5040
      Picture         =   "CubeSkin.ctx":0000
      ScaleHeight     =   195
      ScaleWidth      =   345
      TabIndex        =   1
      ToolTipText     =   "닫기"
      Top             =   50
      Width           =   345
   End
   Begin VB.PictureBox bMaxmize 
      AutoSize        =   -1  'True
      BorderStyle     =   0  '없음
      Enabled         =   0   'False
      Height          =   195
      Left            =   4680
      Picture         =   "CubeSkin.ctx":03EC
      ScaleHeight     =   195
      ScaleWidth      =   345
      TabIndex        =   4
      ToolTipText     =   "최대화"
      Top             =   50
      Width           =   345
   End
   Begin VB.PictureBox bMinimize 
      AutoSize        =   -1  'True
      BorderStyle     =   0  '없음
      Height          =   195
      Left            =   4320
      Picture         =   "CubeSkin.ctx":07D8
      ScaleHeight     =   195
      ScaleWidth      =   345
      TabIndex        =   0
      ToolTipText     =   "최소화"
      Top             =   50
      Width           =   345
   End
   Begin VB.Label lblDrag 
      BackStyle       =   0  '투명
      Height          =   285
      Left            =   -120
      TabIndex        =   2
      Top             =   0
      Width           =   4305
   End
   Begin VB.Image bMaxmize3 
      Height          =   195
      Left            =   4680
      Picture         =   "CubeSkin.ctx":0BC4
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image bMaxmize2 
      Height          =   195
      Left            =   4680
      Picture         =   "CubeSkin.ctx":0FB0
      Top             =   720
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image bMaxmize1 
      Height          =   195
      Left            =   4680
      Picture         =   "CubeSkin.ctx":139C
      Top             =   480
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image UnderBar 
      Height          =   45
      Left            =   0
      Picture         =   "CubeSkin.ctx":1788
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   4245
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Cube"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   -240
      TabIndex        =   3
      Top             =   0
      Width           =   4215
   End
   Begin VB.Image TitleBar 
      Height          =   285
      Left            =   240
      Picture         =   "CubeSkin.ctx":18AE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3765
   End
   Begin VB.Image LTitle 
      Height          =   285
      Left            =   0
      Picture         =   "CubeSkin.ctx":1E94
      Top             =   0
      Width           =   285
   End
   Begin VB.Image RTitle 
      Height          =   285
      Left            =   3960
      Picture         =   "CubeSkin.ctx":247A
      Top             =   0
      Width           =   285
   End
   Begin VB.Image bMinimize1 
      Height          =   195
      Left            =   4320
      Picture         =   "CubeSkin.ctx":2930
      Top             =   480
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image bMinimize2 
      Height          =   195
      Left            =   4320
      Picture         =   "CubeSkin.ctx":2D1C
      Top             =   720
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image bMinimize3 
      Height          =   195
      Left            =   4320
      Picture         =   "CubeSkin.ctx":3108
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image bUnload1 
      Height          =   195
      Left            =   5040
      Picture         =   "CubeSkin.ctx":34F4
      Top             =   480
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image bUnload2 
      Height          =   195
      Left            =   5040
      Picture         =   "CubeSkin.ctx":38E0
      Top             =   720
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image bUnload3 
      Height          =   195
      Left            =   5040
      Picture         =   "CubeSkin.ctx":3CCC
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image LHeight 
      Height          =   3045
      Left            =   0
      Picture         =   "CubeSkin.ctx":40B8
      Stretch         =   -1  'True
      Top             =   240
      Width           =   45
   End
   Begin VB.Image RHeight 
      Height          =   3285
      Left            =   4200
      Picture         =   "CubeSkin.ctx":41DE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   45
   End
End
Attribute VB_Name = "CubeSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ::        Cube Skin         ::
' :: 제작자 arshica@naver.com ::
' ::      이미지 군고구마     ::
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private CName As String
Dim FrmFull As Boolean
Private Sub bMinimize_Click()
bMinimize.Picture = bMinimize1.Picture
Parent.WindowState = 1
End Sub
Private Sub bUnload_Click()
End
End Sub
Private Sub UserControl_Resize()
On Error Resume Next
'상단 타이틀
LTitle.Left = 0
LTitle.Top = 0
RTitle.Left = Width - RTitle.Width
RTitle.Top = 0
TitleBar.Top = 0
TitleBar.Left = LTitle.Width
TitleBar.Width = Width - LTitle.Width - RTitle.Width
'중간
LHeight.Left = 0
LHeight.Top = TitleBar.Height
LHeight.Height = Height - TitleBar.Height
RHeight.Left = Width - RHeight.Width
RHeight.Top = TitleBar.Height
RHeight.Height = Height - TitleBar.Height
'하단
UnderBar.Left = LHeight.Width
UnderBar.Top = Height - UnderBar.Height
UnderBar.Width = Width - RHeight.Width - LHeight.Width
'버튼
bUnload.Left = Width - bUnload.Width - 50
bMaxmize.Left = Width - bUnload.Width - bMaxmize.Width - 60
bMinimize.Left = Width - bUnload.Width - bMaxmize.Width - bMinimize.Width - 70
'제목
lblCaption.Width = Width
'드래그
lblDrag.Width = Width
End Sub
Private Sub bMinimize_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If (x < 0) Or (Y < 0) Or (x > bMinimize.Width) Or (Y > bMinimize.Height) Then
ReleaseCapture
bMinimize.Picture = bMinimize1.Picture
ElseIf GetCapture() <> bMinimize.hWnd Then
SetCapture bMinimize.hWnd
bMinimize.Picture = bMinimize2.Picture
End If
End Sub
Private Sub bMinimize_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
bMinimize.Picture = bMinimize3.Picture
End Sub
Private Sub bUnload_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If (x < 0) Or (Y < 0) Or (x > bUnload.Width) Or (Y > bUnload.Height) Then
ReleaseCapture
bUnload.Picture = bUnload1.Picture
ElseIf GetCapture() <> bUnload.hWnd Then
SetCapture bUnload.hWnd
bUnload.Picture = bUnload2.Picture
End If
End Sub
Private Sub bUnload_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
bUnload.Picture = bUnload3.Picture
End Sub
Private Sub bMaxmize_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If (x < 0) Or (Y < 0) Or (x > bMaxmize.Width) Or (Y > bMaxmize.Height) Then
ReleaseCapture
bMaxmize.Picture = bMaxmize1.Picture
ElseIf GetCapture() <> bMaxmize.hWnd Then
SetCapture bMaxmize.hWnd
bMaxmize.Picture = bMaxmize2.Picture
End If
End Sub
Private Sub lblDrag_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Parent.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
CName = PropBag.ReadProperty("Caption", UserControl.Name)
lblCaption = CName
Parent.Caption = CName
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Caption", CName, Empty)
End Sub
Public Property Get Caption() As String
Caption = CName
End Property
Public Property Let Caption(Str As String)
CName = Str
lblCaption = CName
End Property
