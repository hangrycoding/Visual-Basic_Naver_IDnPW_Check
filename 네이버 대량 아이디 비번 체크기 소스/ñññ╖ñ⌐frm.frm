VERSION 5.00
Begin VB.Form CKeyBoard 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  '없음
   Caption         =   "KeyBoard"
   ClientHeight    =   5070
   ClientLeft      =   9000
   ClientTop       =   6300
   ClientWidth     =   11985
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ㄴㅇㄹfrm.frx":0000
   ScaleHeight     =   5070
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3120
      Top             =   2160
   End
   Begin VB.TextBox Txt_address 
      BackColor       =   &H00400000&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   11775
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label4 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Image Img_oem5 
      Height          =   765
      Left            =   10620
      Tag             =   "N\"
      Top             =   1680
      Width           =   1245
   End
   Begin VB.Image Img_add 
      Height          =   765
      Left            =   10224
      Tag             =   "N="
      Top             =   850
      Width           =   795
   End
   Begin VB.Image Img_minus 
      Height          =   765
      Left            =   9387
      Tag             =   "N-"
      Top             =   850
      Width           =   795
   End
   Begin VB.Image Img_0 
      Height          =   765
      Left            =   8550
      Tag             =   "N)"
      Top             =   850
      Width           =   795
   End
   Begin VB.Image Img_9 
      Height          =   765
      Left            =   7713
      Tag             =   "N("
      Top             =   850
      Width           =   795
   End
   Begin VB.Image Img_8 
      Height          =   765
      Left            =   6876
      Tag             =   "N*"
      Top             =   850
      Width           =   795
   End
   Begin VB.Image Img_7 
      Height          =   765
      Left            =   6039
      Tag             =   "N&"
      Top             =   850
      Width           =   795
   End
   Begin VB.Image Img_6 
      Height          =   765
      Left            =   5202
      Tag             =   "N^"
      Top             =   850
      Width           =   795
   End
   Begin VB.Image Img_5 
      Height          =   765
      Left            =   4365
      Tag             =   "N%"
      Top             =   850
      Width           =   795
   End
   Begin VB.Image Img_4 
      Height          =   765
      Left            =   3528
      Tag             =   "N$"
      Top             =   850
      Width           =   795
   End
   Begin VB.Image Img_3 
      Height          =   765
      Left            =   2691
      Tag             =   "N#"
      Top             =   850
      Width           =   795
   End
   Begin VB.Image Img_2 
      Height          =   765
      Left            =   1854
      Tag             =   "N@"
      Top             =   850
      Width           =   795
   End
   Begin VB.Image Img_1 
      Height          =   765
      Left            =   1017
      Tag             =   "N!"
      Top             =   850
      Width           =   795
   End
   Begin VB.Image Img_oem3 
      Height          =   765
      Left            =   180
      Tag             =   "N`"
      Top             =   850
      Width           =   795
   End
   Begin VB.Image Img_back1 
      Height          =   765
      Left            =   11070
      Tag             =   "NBK"
      Top             =   850
      Width           =   795
   End
   Begin VB.Image Img_A 
      Height          =   765
      Left            =   945
      Tag             =   "CA"
      Top             =   2490
      Width           =   795
   End
   Begin VB.Image Img_S 
      Height          =   765
      Left            =   1785
      Tag             =   "CS"
      Top             =   2490
      Width           =   795
   End
   Begin VB.Image Img_D 
      Height          =   765
      Left            =   2625
      Tag             =   "CD"
      Top             =   2490
      Width           =   795
   End
   Begin VB.Image Img_F 
      Height          =   765
      Left            =   3465
      Tag             =   "CF"
      Top             =   2490
      Width           =   795
   End
   Begin VB.Image Img_G 
      Height          =   765
      Left            =   4305
      Tag             =   "CG"
      Top             =   2490
      Width           =   795
   End
   Begin VB.Image Img_H 
      Height          =   765
      Left            =   5145
      Tag             =   "CH"
      Top             =   2490
      Width           =   795
   End
   Begin VB.Image Img_J 
      Height          =   765
      Left            =   5985
      Tag             =   "CJ"
      Top             =   2490
      Width           =   795
   End
   Begin VB.Image Img_K 
      Height          =   765
      Left            =   6825
      Tag             =   "CK"
      Top             =   2490
      Width           =   795
   End
   Begin VB.Image Img_L 
      Height          =   765
      Left            =   7665
      Tag             =   "CL"
      Top             =   2490
      Width           =   795
   End
   Begin VB.Image Img_Q 
      Height          =   765
      Left            =   520
      Tag             =   "CQ"
      Top             =   1680
      Width           =   795
   End
   Begin VB.Image Img_W 
      Height          =   765
      Left            =   1361
      Tag             =   "CW"
      Top             =   1680
      Width           =   795
   End
   Begin VB.Image Img_E 
      Height          =   765
      Left            =   2202
      Tag             =   "CE"
      Top             =   1680
      Width           =   795
   End
   Begin VB.Image Img_R 
      Height          =   765
      Left            =   3043
      Tag             =   "CR"
      Top             =   1680
      Width           =   795
   End
   Begin VB.Image Img_T 
      Height          =   765
      Left            =   3884
      Tag             =   "CT"
      Top             =   1680
      Width           =   795
   End
   Begin VB.Image img_Y 
      Height          =   765
      Left            =   4725
      Tag             =   "CY"
      Top             =   1680
      Width           =   795
   End
   Begin VB.Image Img_U 
      Height          =   765
      Left            =   5566
      Tag             =   "CU"
      Top             =   1680
      Width           =   795
   End
   Begin VB.Image Img_I 
      Height          =   765
      Left            =   6407
      Tag             =   "CI"
      Top             =   1680
      Width           =   795
   End
   Begin VB.Image Img_O 
      Height          =   765
      Left            =   7248
      Tag             =   "CO"
      Top             =   1680
      Width           =   795
   End
   Begin VB.Image Img_P 
      Height          =   765
      Left            =   8089
      Tag             =   "CP"
      Top             =   1680
      Width           =   795
   End
   Begin VB.Image Img_Z 
      Height          =   765
      Left            =   1365
      Tag             =   "CZ"
      Top             =   3320
      Width           =   795
   End
   Begin VB.Image Img_X 
      Height          =   765
      Left            =   2205
      Tag             =   "CX"
      Top             =   3320
      Width           =   795
   End
   Begin VB.Image Img_C 
      Height          =   765
      Left            =   3045
      Tag             =   "CC"
      Top             =   3320
      Width           =   795
   End
   Begin VB.Image Img_V 
      Height          =   765
      Left            =   3885
      Tag             =   "CV"
      Top             =   3320
      Width           =   795
   End
   Begin VB.Image Img_B 
      Height          =   765
      Left            =   4725
      Tag             =   "CB"
      Top             =   3320
      Width           =   795
   End
   Begin VB.Image Img_N 
      Height          =   765
      Left            =   5565
      Tag             =   "CN"
      Top             =   3320
      Width           =   795
   End
   Begin VB.Image Img_M 
      Height          =   765
      Left            =   6405
      Tag             =   "CM"
      Top             =   3320
      Width           =   795
   End
   Begin VB.Image Img_oem4 
      Height          =   765
      Left            =   8930
      Tag             =   "N["
      Top             =   1680
      Width           =   795
   End
   Begin VB.Image Img_oem6 
      Height          =   765
      Left            =   9780
      Tag             =   "N]"
      Top             =   1680
      Width           =   795
   End
   Begin VB.Image Img_OEM1 
      Height          =   765
      Left            =   8505
      Tag             =   "N;"
      Top             =   2490
      Width           =   795
   End
   Begin VB.Image Img_OEM7 
      Height          =   765
      Left            =   9345
      Tag             =   "N'"
      Top             =   2490
      Width           =   795
   End
   Begin VB.Image Img_OEM_COMMA 
      Height          =   765
      Left            =   7245
      Tag             =   "N,"
      Top             =   3320
      Width           =   795
   End
   Begin VB.Image Img_OEM_PERIOD 
      Height          =   765
      Left            =   8085
      Tag             =   "N."
      Top             =   3320
      Width           =   795
   End
   Begin VB.Image img_OEM2 
      Height          =   765
      Left            =   8925
      Tag             =   "N/"
      Top             =   3320
      Width           =   795
   End
   Begin VB.Image Img_tab 
      Height          =   840
      Left            =   11880
      Picture         =   "ㄴㅇㄹfrm.frx":4DCA
      Top             =   4410
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Image Img_shift 
      Height          =   765
      Left            =   9780
      Tag             =   "SSFT"
      Top             =   3320
      Width           =   2085
   End
   Begin VB.Image Img_enter 
      Height          =   765
      Left            =   10200
      Tag             =   "NRTN"
      Top             =   2490
      Width           =   1665
   End
   Begin VB.Image Img_space 
      Height          =   765
      Left            =   1710
      Tag             =   "NSP"
      Top             =   4110
      Width           =   7155
   End
   Begin VB.Image Img_EngKor 
      Height          =   765
      Left            =   8970
      Tag             =   "NTG"
      Top             =   4110
      Width           =   795
   End
   Begin VB.Image ImgExit 
      Height          =   840
      Left            =   11040
      Picture         =   "ㄴㅇㄹfrm.frx":5EE5
      Stretch         =   -1  'True
      Top             =   4170
      Width           =   825
   End
   Begin VB.Image ImgClear 
      Height          =   840
      Left            =   10170
      Picture         =   "ㄴㅇㄹfrm.frx":80A1
      Stretch         =   -1  'True
      Top             =   4170
      Width           =   825
   End
End
Attribute VB_Name = "CKeyBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim FsShitFlag    As Integer  ' 쉬프트의 상태 체크
Dim FsNumPress    As Integer  ' 한글/영문 변환
Dim FsImcKor      As Long     ' 한글 전환 시스템
Dim FsVKeyCode    As Byte     ' 가상키코드 값GsKeyBoardResult
Dim FsImagePath   As String   ' 이미지경로

Public Function FKeyStatus(FsVKeyCode As Byte, FsShitFlag As Integer) ' 가상키코드값과 쉬프트값 전달

    '쉬프트를 눌렀을때
    If FsShitFlag = 1 Then
        '키다운메세지의 가상키 (쉬프트) 호출
        Call SKeyDown(VK_SHIFT)
       
        '키보드 키다운 메세지 호출
        Call SKeyDown(FsVKeyCode)
        '키보드 키업 메세지 호출
        Call SKeyUp(FsVKeyCode)
        Call SKeyUp(VK_SHIFT)
    Else
        '키보드 키다운 메세지 호출
        Call SKeyDown(FsVKeyCode)
        '키보드 키업 메세지 호출
        Call SKeyUp(FsVKeyCode)
    
    End If
    
End Function

Public Function FChageKorKey()
Dim cControl        As Control
Dim iCount          As Integer

    For Each cControl In Me.Controls
        If TypeOf cControl Is Image Then
            If Mid(cControl.Tag, 1, 1) = "C" Then
                For iCount = 1 To imgChange.ListImages.Count
                    If imgChange.ListImages(iCount).Key = "_" & Mid(cControl.Tag, 2, Len(cControl.Tag) - 1) Then
                        cControl.Picture = imgChange.ListImages(iCount).Picture
                        Exit For
                    End If
                Next iCount
            End If
        End If
    Next cControl
End Function
Public Function FChageEngKey()
Dim cControl        As Control
Dim iCount          As Integer

    For Each cControl In Me.Controls
        If TypeOf cControl Is Image Then
            If Mid(cControl.Tag, 1, 1) = "C" Then
                For iCount = 1 To imgChange.ListImages.Count
                    If imgChange.ListImages(iCount).Key = Mid(cControl.Tag, 2, Len(cControl.Tag) - 1) Then
                        cControl.Picture = imgChange.ListImages(iCount).Picture
                        Exit For
                    End If
                Next iCount
            End If
        End If
    Next cControl
End Function

Public Function FNoneChageKey()
Dim cControl        As Control
Dim iCount          As Integer

    For Each cControl In Me.Controls
        If TypeOf cControl Is Image Then
            If Mid(cControl.Tag, 1, 1) = "N" Then
                For iCount = 1 To imgNoneChg.ListImages.Count
                    If imgNoneChg.ListImages(iCount).Key = Mid(cControl.Tag, 2, Len(cControl.Tag) - 1) Then
                        cControl.Picture = imgNoneChg.ListImages(iCount).Picture
                        Exit For
                    End If
                Next iCount
            End If
        End If
    Next cControl
End Function

Private Sub c0_Click()
    Call FKeyStatus(VK_0, FsShitFlag)
End Sub

Private Sub c1_Click()
    Call FKeyStatus(VK_1, FsShitFlag)
End Sub
Private Sub c2_Click()
    Call FKeyStatus(VK_2, FsShitFlag)
End Sub

Private Sub c3_Click()
    Call FKeyStatus(VK_3, FsShitFlag)
End Sub

Private Sub c4_Click()
    Call FKeyStatus(VK_4, FsShitFlag)
End Sub

Private Sub c5_Click()
    Call FKeyStatus(VK_5, FsShitFlag)
End Sub

Private Sub c6_Click()
    Call FKeyStatus(VK_6, FsShitFlag)
End Sub

Private Sub c7_Click()
    Call FKeyStatus(VK_7, FsShitFlag)
End Sub

Private Sub c8_Click()
    Call FKeyStatus(VK_8, FsShitFlag)
End Sub

Private Sub c9_Click()
    Call FKeyStatus(VK_9, FsShitFlag)
End Sub


Private Sub Form_Load()

Dim TmpNum   As Long
Dim iCount   As Integer

    Me.Top = 3855
    Me.Left = 0
    Me.WindowState = 0
    FsShitFlag = 0
    FsNumPress = 0
    
    '** 바뀌지 않는 Key
    Call FNoneChageKey
    
    '** 초기 Shift Key
    For iCount = 1 To imgNoneChg.ListImages.Count
        If imgNoneChg.ListImages(iCount).Key = Mid(Img_shift.Tag, 2, 3) Then
            Img_shift.Picture = imgNoneChg.ListImages(iCount).Picture
            Exit For
        End If
    Next iCount
    
    If Language_Set = 2 Then
        '** 영문 상태로
        Call FChageEngKey
        
        FsImcKor = ImmGetContext(Me.hwnd)
        ImmSetConversionStatus FsImcKor, IME_CMODE_ALPHANUMERIC, IME_SMODE_NONE
        FsNumPress = 0
        If hHook Then
            TmpNum = UnhookWindowsHookEx(hHook)
            'DoEvents
            hHook = 0
        End If

    Else
        '** 한글 상태로
        FsImcKor = ImmGetContext(Me.hwnd)
        ImmSetConversionStatus FsImcKor, IME_CMODE_HANGEUL, IME_SMODE_NONE
        FsNumPress = 1
        hHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf FKeyboardProc, App.hInstance, App.ThreadID)
        Call FChageKorKey
    End If
     
End Sub

Private Sub img_engkor_Click() ' 한영버튼
    Dim TmpNum   As Long
    
    If FsNumPress = 0 Then ' 영-한
        FsImcKor = ImmGetContext(Me.hwnd)
        ImmSetConversionStatus FsImcKor, IME_CMODE_HANGEUL, IME_SMODE_NONE
        FsNumPress = 1
        hHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf FKeyboardProc, App.hInstance, App.ThreadID)
        Call FChageKorKey
    Else ' 한-영
        Call FChageEngKey
        FsImcKor = ImmGetContext(Me.hwnd)
        ImmSetConversionStatus FsImcKor, IME_CMODE_ALPHANUMERIC, IME_SMODE_NONE
        FsNumPress = 0

        If hHook Then
            TmpNum = UnhookWindowsHookEx(hHook)
            'DoEvents
            hHook = 0
        End If
    End If


End Sub

Private Sub Image2_Click()
    Call FKeyStatus(VK_OEM_CLEAR, FsShitFlag)
End Sub

Private Sub Img_1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_1.BorderStyle = 1
    Call FKeyStatus(VK_1, FsShitFlag)
End Sub

Private Sub Img_1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_1.BorderStyle = 0
End Sub
Private Sub Img_2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_2.BorderStyle = 1
    Call FKeyStatus(VK_2, FsShitFlag)
End Sub

Private Sub Img_2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_2.BorderStyle = 0
End Sub
Private Sub Img_3_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_3.BorderStyle = 1
    Call FKeyStatus(VK_3, FsShitFlag)
End Sub

Private Sub Img_3_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_3.BorderStyle = 0
End Sub
Private Sub Img_4_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_4.BorderStyle = 1
    Call FKeyStatus(VK_4, FsShitFlag)
End Sub

Private Sub Img_4_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_4.BorderStyle = 0
End Sub
Private Sub Img_5_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_5.BorderStyle = 1
    Call FKeyStatus(VK_5, FsShitFlag)
End Sub

Private Sub Img_5_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_5.BorderStyle = 0
End Sub
Private Sub Img_6_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_6.BorderStyle = 1
    Call FKeyStatus(VK_6, FsShitFlag)
End Sub

Private Sub Img_6_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_6.BorderStyle = 0
End Sub
Private Sub Img_7_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_7.BorderStyle = 1
    Call FKeyStatus(VK_7, FsShitFlag)
End Sub

Private Sub Img_7_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_7.BorderStyle = 0
End Sub
Private Sub Img_8_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_8.BorderStyle = 1
Call FKeyStatus(VK_8, FsShitFlag)
End Sub

Private Sub Img_8_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_8.BorderStyle = 0
End Sub
Private Sub Img_9_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_9.BorderStyle = 1
    Call FKeyStatus(VK_9, FsShitFlag)
End Sub

Private Sub Img_9_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_9.BorderStyle = 0
End Sub
Private Sub Img_0_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_0.BorderStyle = 1
    Call FKeyStatus(VK_0, FsShitFlag)
End Sub

Private Sub Img_0_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_0.BorderStyle = 0
End Sub

Private Sub Img_EngKor_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_EngKor.BorderStyle = 1
End Sub

Private Sub Img_EngKor_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_EngKor.BorderStyle = 0
End Sub

Private Sub Img_minus_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_minus.BorderStyle = 1
    Call FKeyStatus(VK_OEM_MINUS, FsShitFlag)
End Sub

Private Sub Img_minus_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_minus.BorderStyle = 0
End Sub
Private Sub Img_add_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_add.BorderStyle = 1
    Call FKeyStatus(VK_OEM_PLUS, FsShitFlag)
End Sub
Private Sub Img_add_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_add.BorderStyle = 0
End Sub

Private Sub Img_oem3_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_oem3.BorderStyle = 1
    Call FKeyStatus(VK_OEM_3, FsShitFlag)
End Sub

Private Sub Img_oem3_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_oem3.BorderStyle = 0
End Sub

Private Sub Img_oem5_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_oem5.BorderStyle = 1
    Call FKeyStatus(VK_OEM_5, FsShitFlag)
End Sub
Private Sub Img_oem5_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_oem5.BorderStyle = 0
End Sub
Private Sub Img_back1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_back1.BorderStyle = 1
    If Len(Trim(Txt_address.Text)) = 0 Then Exit Sub
    Call FKeyStatus(VK_BACK, FsShitFlag)
End Sub
Private Sub Img_back1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_back1.BorderStyle = 0
End Sub
Private Sub Img_Q_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_Q.BorderStyle = 1
    Call FKeyStatus(VK_Q, FsShitFlag)
End Sub
Private Sub Img_Q_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_Q.BorderStyle = 0
End Sub
Private Sub Img_w_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_W.BorderStyle = 1
    Call FKeyStatus(VK_W, FsShitFlag)
End Sub
Private Sub Img_w_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_W.BorderStyle = 0
End Sub
Private Sub Img_e_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_E.BorderStyle = 1
    Call FKeyStatus(VK_E, FsShitFlag)
End Sub
Private Sub Img_e_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_E.BorderStyle = 0
End Sub
Private Sub Img_r_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_R.BorderStyle = 1
    Call FKeyStatus(VK_R, FsShitFlag)
End Sub
Private Sub Img_r_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_R.BorderStyle = 0
End Sub
Private Sub Img_t_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_T.BorderStyle = 1
    Call FKeyStatus(VK_T, FsShitFlag)
End Sub
Private Sub Img_t_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_T.BorderStyle = 0
End Sub
Private Sub Img_y_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    img_Y.BorderStyle = 1
    Call FKeyStatus(VK_Y, FsShitFlag)
End Sub
Private Sub Img_y_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    img_Y.BorderStyle = 0
End Sub
Private Sub Img_u_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_U.BorderStyle = 1
    Call FKeyStatus(VK_U, FsShitFlag)
End Sub
Private Sub Img_u_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_U.BorderStyle = 0
End Sub
Private Sub Img_i_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_I.BorderStyle = 1
    Call FKeyStatus(VK_I, FsShitFlag)
End Sub
Private Sub Img_i_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_I.BorderStyle = 0
End Sub
Private Sub Img_o_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_O.BorderStyle = 1
    Call FKeyStatus(VK_O, FsShitFlag)
End Sub
Private Sub Img_o_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_O.BorderStyle = 0
End Sub
Private Sub Img_p_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_P.BorderStyle = 1
    Call FKeyStatus(VK_P, FsShitFlag)
End Sub
Private Sub Img_p_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_P.BorderStyle = 0
End Sub
Private Sub Img_oem4_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_oem4.BorderStyle = 1
    Call FKeyStatus(VK_OEM_4, FsShitFlag)
End Sub
Private Sub Img_oem4_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_oem4.BorderStyle = 0
End Sub
Private Sub Img_oem6_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_oem6.BorderStyle = 1
    Call FKeyStatus(VK_OEM_6, FsShitFlag)
End Sub
Private Sub Img_oem6_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_oem6.BorderStyle = 0
End Sub
Private Sub Img_a_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_A.BorderStyle = 1
    Call FKeyStatus(VK_A, FsShitFlag)
End Sub
Private Sub Img_a_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_A.BorderStyle = 0
End Sub
Private Sub Img_s_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_S.BorderStyle = 1
    Call FKeyStatus(VK_S, FsShitFlag)
End Sub
Private Sub Img_s_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_S.BorderStyle = 0
End Sub
Private Sub Img_d_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_D.BorderStyle = 1
    Call FKeyStatus(VK_D, FsShitFlag)
End Sub
Private Sub Img_d_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_D.BorderStyle = 0
End Sub
Private Sub Img_f_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_F.BorderStyle = 1
    Call FKeyStatus(VK_F, FsShitFlag)
End Sub
Private Sub Img_f_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_F.BorderStyle = 0
End Sub
Private Sub Img_g_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_G.BorderStyle = 1
    Call FKeyStatus(VK_G, FsShitFlag)
End Sub
Private Sub Img_g_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_G.BorderStyle = 0
End Sub
Private Sub Img_h_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_H.BorderStyle = 1
    Call FKeyStatus(VK_H, FsShitFlag)
End Sub
Private Sub Img_h_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_H.BorderStyle = 0
End Sub
Private Sub Img_j_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_J.BorderStyle = 1
    Call FKeyStatus(VK_J, FsShitFlag)
End Sub
Private Sub Img_j_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_J.BorderStyle = 0
End Sub
Private Sub Img_k_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_K.BorderStyle = 1
    Call FKeyStatus(VK_K, FsShitFlag)
End Sub
Private Sub Img_k_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_K.BorderStyle = 0
End Sub
Private Sub Img_L_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_L.BorderStyle = 1
    Call FKeyStatus(VK_L, FsShitFlag)
End Sub
Private Sub Img_L_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_L.BorderStyle = 0
End Sub
Private Sub Img_oem1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_OEM1.BorderStyle = 1
    Call FKeyStatus(VK_OEM_1, FsShitFlag)
End Sub
Private Sub Img_oem1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_OEM1.BorderStyle = 0
End Sub
Private Sub Img_oem7_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_OEM7.BorderStyle = 1
    Call FKeyStatus(VK_OEM_7, FsShitFlag)
End Sub
Private Sub Img_oem7_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_OEM7.BorderStyle = 0
End Sub
Private Sub Img_enter_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_enter.BorderStyle = 1
End Sub
Private Sub Img_enter_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_enter.BorderStyle = 0
    Dim TmpNum   As Long
    GsKeyBoardResult = RTrim(Txt_address)
    Unload Me
End Sub
Private Sub Img_z_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_Z.BorderStyle = 1
    Call FKeyStatus(VK_Z, FsShitFlag)
End Sub
Private Sub Img_z_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_Z.BorderStyle = 0
End Sub
Private Sub Img_x_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_X.BorderStyle = 1
    Call FKeyStatus(VK_X, FsShitFlag)
End Sub
Private Sub Img_x_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_X.BorderStyle = 0
End Sub
Private Sub Img_c_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_C.BorderStyle = 1
    Call FKeyStatus(VK_C, FsShitFlag)
End Sub
Private Sub Img_c_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_C.BorderStyle = 0
End Sub
Private Sub Img_v_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_V.BorderStyle = 1
    Call FKeyStatus(VK_V, FsShitFlag)
End Sub
Private Sub Img_v_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_V.BorderStyle = 0
End Sub
Private Sub Img_b_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_B.BorderStyle = 1
    Call FKeyStatus(VK_B, FsShitFlag)
End Sub
Private Sub Img_b_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_B.BorderStyle = 0
End Sub
Private Sub Img_n_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_N.BorderStyle = 1
    Call FKeyStatus(VK_N, FsShitFlag)
End Sub
Private Sub Img_n_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_N.BorderStyle = 0
End Sub
Private Sub Img_m_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_M.BorderStyle = 1
    Call FKeyStatus(VK_M, FsShitFlag)
End Sub
Private Sub Img_m_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_M.BorderStyle = 0
End Sub
Private Sub Img_oem_comma_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_OEM_COMMA.BorderStyle = 1
    Call FKeyStatus(VK_OEM_COMMA, FsShitFlag)
End Sub
Private Sub Img_oem_comma_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_OEM_COMMA.BorderStyle = 0
End Sub
Private Sub Img_oem_period_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_OEM_PERIOD.BorderStyle = 1
    Call FKeyStatus(VK_OEM_PERIOD, FsShitFlag)
End Sub
Private Sub Img_oem_period_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_OEM_PERIOD.BorderStyle = 0
End Sub
Private Sub Img_oem2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    img_OEM2.BorderStyle = 1
    Call FKeyStatus(VK_OEM_2, FsShitFlag)
End Sub
Private Sub Img_oem2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    img_OEM2.BorderStyle = 0
End Sub
Private Sub Img_tab_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_tab.BorderStyle = 1
    Call FKeyStatus(VK_TAB, FsShitFlag)
End Sub
Private Sub Img_tab_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_tab.BorderStyle = 0
End Sub
Private Sub Img_space_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_space.BorderStyle = 1
    Call FKeyStatus(VK_SPACE, FsShitFlag)
End Sub
Private Sub Img_space_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_space.BorderStyle = 0
End Sub

Private Sub Img_shift_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim iCount      As Integer

    Img_shift.BorderStyle = 1
    If FsShitFlag = 0 Then '쉬프트 처음상태
        FsShitFlag = 1
        For iCount = 1 To imgNoneChg.ListImages.Count
            If imgNoneChg.ListImages(iCount).Key = "_" & Mid(Img_shift.Tag, 2, 3) Then
                Img_shift.Picture = imgNoneChg.ListImages(iCount).Picture
                Exit For
            End If
        Next iCount
    Else                  '쉬프트가 눌렸을때
        FsShitFlag = 0
        For iCount = 1 To imgNoneChg.ListImages.Count
            If imgNoneChg.ListImages(iCount).Key = Mid(Img_shift.Tag, 2, 3) Then
                Img_shift.Picture = imgNoneChg.ListImages(iCount).Picture
                Exit For
            End If
        Next iCount
    End If

End Sub
Private Sub Img_shift_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Img_shift.BorderStyle = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim TmpNum As Long


    If FsNumPress = 1 Then ' 한글상태 -> ENG
        FsImcKor = ImmGetContext(Me.hwnd)
        ImmSetConversionStatus FsImcKor, IME_CMODE_ALPHANUMERIC, IME_SMODE_NONE
        If hHook Then
            TmpNum = UnhookWindowsHookEx(hHook)
        '    'DoEvents
            hHook = 0
        End If
        FsNumPress = 0
    End If
End Sub

Private Sub ImgClear_Click()
    Txt_address.Text = ""
End Sub

Private Sub ImgClear_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ImgClear.BorderStyle = 1
End Sub

Private Sub ImgClear_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ImgClear.BorderStyle = 0
End Sub

Private Sub ImgExit_Click()
    GsKeyBoardResult = Space(1)
    Unload Me

End Sub

Private Sub ImgExit_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ImgExit.BorderStyle = 1
End Sub

Private Sub ImgExit_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ImgExit.BorderStyle = 0
End Sub



Private Sub Txt_address_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       Url = Txt_address.Text
    End If
End Sub

