VERSION 5.00
Begin VB.Form main 
   Caption         =   "单词学习21.0"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14955
   LinkTopic       =   "Form3"
   ScaleHeight     =   8250
   ScaleWidth      =   14955
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command4 
      Caption         =   "关"
      Height          =   495
      Left            =   6600
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "换肤"
      Height          =   495
      Left            =   13320
      TabIndex        =   2
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "背默系统"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3960
      TabIndex        =   0
      Top             =   4440
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "单词学习系统"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4440
      TabIndex        =   1
      Top             =   960
      Width           =   5535
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SkinH_AttachEx Lib "SkinH_VB6.dll" (ByVal lpSkinFile As String, ByVal lpPasswd As String) As Long
Dim aa1 As Integer


Private Sub Command2_Click()
mx.Show
Me.Hide
End Sub

Private Sub Command3_Click()
Dim aa(1 To 10) As String
aa1 = aa1 + 1
aa(2) = "black": aa(1) = "darkroyale": aa(3) = "china"
SkinH_AttachEx App.Path & "/皮肤/" & aa(aa1) & ".she", ""
If aa1 = 3 Then aa1 = 0
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()
SkinH_AttachEx App.Path & "/皮肤/china.she", ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
