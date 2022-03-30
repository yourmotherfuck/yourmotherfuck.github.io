VERSION 5.00
Begin VB.Form 借书界面 
   Caption         =   "借书界面"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   9600
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   11
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "新增借书"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   10
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   3120
      TabIndex        =   9
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3120
      TabIndex        =   8
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3120
      TabIndex        =   7
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3120
      TabIndex        =   6
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3120
      TabIndex        =   5
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "输入借书数量："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   720
      TabIndex        =   4
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "输入借书时间："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   720
      TabIndex        =   3
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "输入借书人（单位）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "输入书名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "输入书籍号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "借书界面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'这个语句用于报错了就跳过这个语句
'On Error Resume Next

'按照惯例，先检查一遍
If Trim(Text1(0)) = empy Then MsgBox "请输入书籍号", 48, "注意": Exit Sub
If Trim(Text1(1)) = empy Then MsgBox "请输入书名", 48, "注意": Exit Sub
If Trim(Text1(2)) = empy Then MsgBox "请输入借书人（单位）", 48, "注意": Exit Sub
If Trim(Text1(3)) = empy Then Text1(3) = Date
If Trim(Text1(4)) = empy Then MsgBox "请输入借书数量", 48, "注意": Exit Sub
If IsDate(Trim(Text1(3))) = False Then MsgBox "请输入正确的时间格式,eg:2022/3/22": Exit Sub

'排错机制（要确定书库里有这样对应的书号和书名）
Dim rs As New ADODB.Recordset
Dim sql As String
sql = "select 书号,书名 from 书库系统 where 书号 =" & "'" & Trim(Text1(0)) & "'" & "and 书名 = " & "'" & Trim(Text1(1)) & "'"
rs.Open sql, 主界面加查找界面.conn, 1, 3
    If rs.EOF Then MsgBox "书库里没有这样的数，请认真检查书籍号和与他对应的书名是否正确", 48, "注意"
rs.Close


'这里开始同步更改数据
sql = "select * from 书库系统 where 书号 =  " & "'" & Trim(Text1(0)) & "'"
rs.Open sql, 主界面加查找界面.conn, 1, 3
    '先拿到下文需要的书籍信息变更表的专业和变更前的数据
    Dim vocation As String, value_before As String
    vocation = rs.Fields("专业").Value
    value_before = rs.Fields("库存").Value
    
    reserve = rs.Fields("库存").Value - Val(Trim(Text1(4)))
    If reserve < 0 Then MsgBox "借书数量不能大于库存", 48, "警告": Exit Sub
    rs.Fields("库存") = reserve
    rs.Update
rs.Close


'将数据加入借书记录表
rs.Open "借书记录表", 主界面加查找界面.conn, 1, 3
    rs.AddNew
    rs.Fields("书号") = Trim(Text1(0))
    rs.Fields("书名") = Trim(Text1(1))
    rs.Fields("借书人（单位）") = Trim(Text1(2))
    rs.Fields("借书时间") = Trim(Text1(3))
    rs.Fields("借书数量") = Trim(Text1(4))
    rs.Update
rs.Close

'同时在书籍信息变更表里面加入对应的东西
temp = "借" + Trim(Text1(4)) + "本书"
rs.Open "书籍信息变更记录表", 主界面加查找界面.conn, 1, 3
    rs.AddNew
    rs.Fields("书号") = Trim(Text1(0))
    rs.Fields("书名") = Trim(Text1(1))
    rs.Fields("专业") = vocation
    rs.Fields("更改时间") = Trim(Text1(3))
    rs.Fields("更改的类型") = "借书"
    rs.Fields("变更的数据") = temp
    rs.Fields("更改前") = value_before
    rs.Fields("更改后") = reserve
    rs.Update
rs.Close


MsgBox "借书成功，借书内容为：" & Chr(13) & "书籍号：" & Trim(Text1(0)) & Space(5) & "书名：" & Trim(Text1(1)) & Space(5) & "借书人（单位）：" & Trim(Text1(2)) & Chr(13) & "借书时间：" & Trim(Text1(3)) & Space(5) & "借书数量：" & Trim(Text1(4)), , "通过"

'完成后回到主界面
Unload Me
主界面加查找界面.Enabled = True
主界面加查找界面.Show
主界面加查找界面.maindata_load " select * from  书库系统 order by " & Trim(主界面加查找界面.Combo2.Text), True
End Sub

Private Sub Command2_Click()
Unload Me
主界面加查找界面.Enabled = True
主界面加查找界面.Show
'主界面加查找界面.maindata_load (" select * from  书库系统 order by " & Trim(主界面加查找界面.Combo2.Text) & " desc ")
End Sub

Private Sub Form_Load()
'主界面加查找界面.conn
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 0
    主界面加查找界面.Enabled = True
    主界面加查找界面.Show
    '主界面加查找界面.maindata_load (" select * from  书库系统 order by " & Trim(主界面加查找界面.Combo2.Text) & " desc ")
End Sub
