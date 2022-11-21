VERSION 5.00
Begin VB.Form KBoardForm 
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   2340
   ClientTop       =   1740
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   6825
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1740
      Width           =   3585
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   435
      Left            =   960
      ScaleHeight     =   435
      ScaleWidth      =   705
      TabIndex        =   0
      Top             =   1200
      Width           =   705
   End
End
Attribute VB_Name = "KBoardForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'사용방법
'BugMoney@nate.com
'박대영


Private Sub Picture1_Click()
    Text1.Text = ""
    If fbKeyBoardShow(Text1) Then
    
    End If
    Set CKeyBoard = Nothing
    
         
End Sub
