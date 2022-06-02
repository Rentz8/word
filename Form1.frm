VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12975
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8160
   ScaleMode       =   0  'User
   ScaleWidth      =   17691.91
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command29 
      Caption         =   "关闭"
      Height          =   495
      Left            =   11040
      TabIndex        =   22
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   1920
      Picture         =   "Form1.frx":14F856
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   1920
      Picture         =   "Form1.frx":14FD9A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   1920
      Picture         =   "Form1.frx":1502DE
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Height          =   615
      Left            =   1920
      Picture         =   "Form1.frx":150822
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5880
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Height          =   615
      Left            =   1920
      Picture         =   "Form1.frx":150D66
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6840
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Height          =   615
      Left            =   5040
      Picture         =   "Form1.frx":1512AA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Height          =   615
      Left            =   5040
      Picture         =   "Form1.frx":1517EE
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Height          =   615
      Left            =   5040
      Picture         =   "Form1.frx":151D32
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Height          =   615
      Left            =   5040
      Picture         =   "Form1.frx":152276
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5880
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Height          =   615
      Left            =   5040
      Picture         =   "Form1.frx":1527BA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6840
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Height          =   615
      Left            =   8160
      Picture         =   "Form1.frx":152CFE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Height          =   615
      Left            =   8160
      Picture         =   "Form1.frx":153242
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Height          =   615
      Left            =   8160
      Picture         =   "Form1.frx":153786
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      Height          =   615
      Left            =   8160
      Picture         =   "Form1.frx":153CCA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Height          =   615
      Left            =   8160
      Picture         =   "Form1.frx":15420E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      Height          =   615
      Left            =   11400
      Picture         =   "Form1.frx":154752
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      Height          =   615
      Left            =   11400
      Picture         =   "Form1.frx":154C96
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command18 
      Height          =   615
      Left            =   11400
      Picture         =   "Form1.frx":1551DA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton Command19 
      Height          =   615
      Left            =   11400
      Picture         =   "Form1.frx":15571E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5880
      Width           =   615
   End
   Begin VB.CommandButton Command20 
      Height          =   615
      Left            =   11400
      Picture         =   "Form1.frx":155C62
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3960
      Top             =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      TabIndex        =   21
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "按钮单击发音一次               长按循环发音                    双击空白处切换元音辅音"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7680
      TabIndex        =   20
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "Winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public num As String

Private Sub Command29_Click()

Me.Hide
End Sub

Private Sub Form_DblClick()
    Form1.Hide
    Form2.Show
    
End Sub


Private Sub Timer1_Timer()
   'Form1.Print "hello" '测试文字
    'Label1.Caption = num
    wav (num) '计时器每隔1s播放一次音频
End Sub

Private Function wav(num)
    Label1.Caption = num
    SoundFile = App.Path & "\wav\" + num + ".wav" '相对路径调用工程下wav文件夹的1.wav
    Result = sndPlaySound(SoundFile, 1)
End Function

Private Sub Command1_Click()
    num = 1
    wav (num)
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 1 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub

Private Sub Command2_Click()
    wav (num)
End Sub

Private Sub Command3_Click()
    wav (num)
End Sub

Private Sub Command4_Click()
    wav (num)
End Sub

Private Sub Command5_Click()
    wav (num)
End Sub
Private Sub Command6_Click()
    wav (num)
End Sub
Private Sub Command7_Click()
    wav (num)
End Sub
Private Sub Command8_Click()
    wav (num)
End Sub

Private Sub Command9_Click()
    wav (num)
End Sub

Private Sub Command10_Click()
    wav (num)
End Sub

Private Sub Command11_Click()
    wav (num)
End Sub
Private Sub Command12_Click()
    wav (num)
End Sub
Private Sub Command13_Click()
    wav (num)
End Sub
Private Sub Command14_Click()
    wav (num)
End Sub

Private Sub Command15_Click()
    wav (num)
End Sub

Private Sub Command16_Click()
    wav (num)
End Sub

Private Sub Command17_Click()
    wav (num)
End Sub
Private Sub Command18_Click()
    wav (num)
End Sub
Private Sub Command19_Click()
    wav (num)
End Sub
Private Sub Command20_Click()
    wav (num)
End Sub
Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 2 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub
Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 3 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub
Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 4 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub
Private Sub Command5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 5 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub
Private Sub Command6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 6 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub
Private Sub Command7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 7 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub
Private Sub Command8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 8 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub
Private Sub Command9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 9 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub
Private Sub Command10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 10 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub
Private Sub Command11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 11 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub
Private Sub Command12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 12 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub
Private Sub Command13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 13 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub
Private Sub Command14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 14 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub
Private Sub Command15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 15 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub
Private Sub Command16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 16 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub
Private Sub Command17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 17 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command17_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub
Private Sub Command18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 18 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command18_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub
Private Sub Command19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 19 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command19_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub
Private Sub Command20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    num = 20 '音频顺序名
    Timer1.Enabled = True  '按下按钮开启计时器
End Sub

Private Sub Command20_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False  '按下按钮关闭计时器
End Sub


