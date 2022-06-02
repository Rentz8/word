VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form mx 
   Caption         =   "单词背默系统"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   18960
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command10 
      Caption         =   "音标"
      Height          =   495
      Left            =   14280
      TabIndex        =   21
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "错词重新开始"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   20
      Top             =   6600
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "word.frx":0000
      Left            =   600
      List            =   "word.frx":00DF
      TabIndex        =   19
      Text            =   "1"
      Top             =   7800
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "背诵 "
      Height          =   495
      Left            =   9360
      TabIndex        =   18
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "关闭"
      Height          =   495
      Left            =   17040
      TabIndex        =   16
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "音标提示"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "判断2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   12
      Top             =   6480
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Text            =   "错词输入处"
      Top             =   5760
      Width           =   9615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "抽取错词"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   5820
      Left            =   10560
      TabIndex        =   9
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "重新开始"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   7
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "提示"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   6
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "判断1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "word.frx":020D
      Top             =   4200
      Width           =   6495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "抽取"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Text            =   "此处输入单词"
      Top             =   1920
      Width           =   9615
   End
   Begin MSForms.CommandButton cmdcmd1 
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   3240
      Width           =   6495
      Size            =   "11456;1508"
      FontName        =   "Cambria"
      FontEffects     =   1073741825
      FontHeight      =   435
      FontCharSet     =   0
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label5 
      Caption         =   "词性参照表"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14760
      TabIndex        =   14
      Top             =   960
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   4950
      Left            =   12600
      Picture         =   "word.frx":0216
      Top             =   2280
      Width           =   6360
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   13
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   7440
      TabIndex        =   8
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "请点击“抽取”或“抽取错词”进行单词的抽取。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   9375
   End
End
Attribute VB_Name = "mx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(1 To 5000) As String  '英文
Dim b(1 To 5000) As String  '中文
Dim c(1 To 5000) As String  '章节
Dim e(1 To 5000) As Integer '正确次数统计
Dim f(1 To 5000) As Integer '错误次数统计
Dim g(1 To 5000) As Integer '序号
Dim d(1 To 5000) As Boolean '是否已被抽取标记
Dim h(1 To 5000) As String '读取音标
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim n As Integer
Dim m As Integer
Dim i As Integer
Dim tishi As Integer
Dim bs As Boolean

Dim sum As Integer





Private Sub Check1_Click()
  bs = Check1.Value
End Sub

Private Sub Command10_Click()
Form1.Show
End Sub

Private Sub Command7_Click()
On Error Resume Next
If i <> 0 Then
'Image2.Picture = LoadPicture("image\" & g(i) & ".jpg")
cmdcmd1.Caption = h(i)
End If
End Sub




Private Sub Command8_Click()
For j = 1 To n
   If e(j) < 1 And f(j) > 3 Then d(j) = False
Next j
Command1.Enabled = True
Command5.Enabled = True
m = 0
End Sub

Private Sub Command9_Click()
main.Show
Me.Hide
End Sub

Private Sub Form_Load() '程序初始化
Check1.Value = False
conn.ConnectionString = "Provider= Microsoft.ACE.OLEDB.12.0;DATA Source=" & App.Path & "\1111.mdb"
conn.Open
Set rs.ActiveConnection = conn
rs.Open "select * from word where  chapter= 1"
n = 0
m = 0
Do While Not rs.EOF
  n = n + 1
  a(n) = rs.Fields(1)
  b(n) = rs.Fields(2)
  c(n) = rs.Fields(3)
  e(n) = rs.Fields(4)
  f(n) = rs.Fields(5)
  g(n) = rs.Fields(0)
  h(n) = rs.Fields(6)
  rs.MoveNext
Loop
rs.Close
conn.Close
Me.KeyPreview = True
End Sub
Private Sub Combo1_Click() '更换章节
Dim j As Integer
conn.ConnectionString = "Provider= Microsoft.ACE.OLEDB.12.0;DATA Source=" & App.Path & "\1111.mdb"
conn.Open
Set rs.ActiveConnection = conn
rs.Open "select * from word where  chapter= " & Combo1.Text
For j = 1 To n
 a(j) = 0
 b(j) = 0
 c(j) = 0
 e(j) = 0
 f(j) = 0
 g(j) = 0
 h(j) = ""
 d(j) = False
Next j
m = 0
n = 0
On Error Resume Next
Do While Not rs.EOF
  n = n + 1
  a(n) = rs.Fields(1)
  b(n) = rs.Fields(2)
  c(n) = rs.Fields(3)
  e(n) = rs.Fields(4)
  f(n) = rs.Fields(5)
  g(n) = rs.Fields(0)
  h(n) = rs.Fields(6)
  rs.MoveNext
Loop
rs.Close
conn.Close
Command1.Enabled = True
Command5.Enabled = True
End Sub

Private Sub Command1_Click() '抽取
Dim j As Integer, aa As Integer
Randomize
List1.Clear
Text1.Text = ""
Label2.Caption = ""
aa = 0
m = m + 1

For j = 1 To n
  If d(j) = False And e(j) < 1 Then aa = aa + 1
  List1.AddItem Str(j) + Str(d(j)) + Str(e(j)) + Str(aa)
Next j

If aa = 0 Then Command1.Enabled = False: Exit Sub
i = Int(Rnd * n) + 1
Do While d(i) = True Or e(i) >= 1
  i = Int(Rnd * n) + 1
Loop
Label1.Caption = Str(m) + ". " + b(i)
d(i) = True
If bs = True Then
  Text2.Text = a(i)
  On Error Resume Next
  cmdcmd1.Caption = h(i)
End If
End Sub
Private Sub Command2_Click() '判断1
Dim ch As String, j As Integer
Dim c As String
Dim s1 As String, s2 As String
ch = Text1.Text
For j = 1 To Len(a(i))
  c = Mid(a(i), j, 1)
  If c >= "A" And c <= "Z" Then
     c = Chr(Asc(c) + 32)
  End If
  s1 = s1 + c
Next j
For j = 1 To Len(ch)
  c = Mid(ch, j, 1)
  If c >= "A" And c <= "Z" Then
     c = Chr(Asc(c) + 32)
  End If
  s2 = s2 + c
Next j
If s2 = s1 Then
  Label2.Caption = "正确"
  e(i) = e(i) + 1
Else
  Label2.Caption = "错误"
End If
End Sub

Private Sub Command5_Click() '抽取错词
Dim j As Integer, a As Integer
Randomize
List1.Clear
Text3.Text = ""
Label4.Caption = ""
a = 0
m = m + 1

For j = 1 To n
  If d(j) = False And e(j) < 1 And f(j) > 3 Then a = a + 1
  List1.AddItem Str(j) + Str(d(j)) + Str(e(j)) + Str(a)
Next j

If a = 0 Then Command5.Enabled = False: Exit Sub
i = Int(Rnd * n) + 1
Do While d(i) = True Or e(i) >= 1 Or f(i) = 0
  i = Int(Rnd * n) + 1
Loop
Label1.Caption = Str(m) + ". " + b(i)
d(i) = True
End Sub
Private Sub Command6_Click() '判断2
Dim ch As String, j As Integer
Dim c As String
Dim s1 As String, s2 As String
ch = Text3.Text
For j = 1 To Len(a(i))
  c = Mid(a(i), j, 1)
  If c >= "A" And c <= "Z" Then
     c = Chr(Asc(c) + 32)
  End If
  s1 = s1 + c
Next j
For j = 1 To Len(ch)
  c = Mid(ch, j, 1)
  If c >= "A" And c <= "Z" Then
     c = Chr(Asc(c) + 32)
  End If
  s2 = s2 + c
Next j
If s2 = s1 Then
  Label4.Caption = "正确"
  e(i) = e(i) + 1
Else
  Label4.Caption = "错误"
End If
End Sub


Private Sub Command3_Click() '提示
If i <> 0 Then
Text2.Text = a(i)
'On Error Resume Next
'Image2.Picture = LoadPicture("image\" & g(i) & ".jpg")
cmdcmd1.Caption = h(i)

tishi = tishi + 1
Label3.Caption = "提示：" + Str(tishi) + "次"
f(i) = f(i) + 1
e(i) = e(i) - 1

conn.ConnectionString = "Provider= Microsoft.ACE.OLEDB.12.0;DATA Source=" & App.Path & "\1111.mdb"
conn.Open

rs.Open "select * from word where num=" & g(i), conn, 1, 3
rs("wrong").Value = rs("wrong").Value + 1
rs.Update

rs.Close
conn.Close
End If
Text1.Text = ""
End Sub

Private Sub Command4_Click() '重新开始
Dim j As Integer
For j = 1 To n
  d(j) = False
Next j
Command1.Enabled = True
Command5.Enabled = True
m = 0
End Sub



Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) = 1 Then
   Text2.Text = ""
   cmdcmd1.Caption = ""
End If
End Sub

Private Sub Text1_Click()
Text1.Text = ""
Text2.Text = ""
cmdcmd1.Caption = ""
'Image2.Picture = LoadPicture("image\30.jpg")
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 59 Then Command2_Click
If KeyAscii = 13 Then Command1_Click
'Print KeyAscii
End Sub


Private Sub Text3_Click()
Text3.Text = ""
Text2.Text = ""
cmdcmd1.Caption = ""
'Image2.Picture = LoadPicture("image\30.jpg")
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 59 Then Command6_Click
If KeyAscii = 13 Then Command5_Click
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then   'F1
    Command3_Click
ElseIf KeyCode = vbKey1 And Shift = 2 Then   'Ctrl+1
    Command3_Click
ElseIf KeyCode = vbKey1 Then
    Command3_Click
End If
End Sub
