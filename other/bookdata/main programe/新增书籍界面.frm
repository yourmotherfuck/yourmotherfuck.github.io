VERSION 5.00
Begin VB.Form 新增书籍界面 
   Caption         =   "新增书籍"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   8625
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "新增书籍界面.frx":0000
      Left            =   5280
      List            =   "新增书籍界面.frx":0013
      TabIndex        =   15
      Text            =   "这里输入B段的专业"
      Top             =   2280
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "新增书籍界面.frx":004A
      Left            =   2760
      List            =   "新增书籍界面.frx":0066
      TabIndex        =   14
      Text            =   "这里输入A段的专业"
      Top             =   2280
      Width           =   2295
   End
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
      TabIndex        =   12
      Top             =   5280
      Width           =   1935
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
      Index           =   5
      Left            =   2760
      TabIndex        =   5
      Top             =   4320
      Width           =   1815
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
      Left            =   2760
      TabIndex        =   4
      Top             =   3600
      Width           =   1815
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
      Left            =   2760
      TabIndex        =   3
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "新建入库"
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
      Left            =   600
      TabIndex        =   2
      Top             =   5280
      Width           =   2295
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
      Left            =   2760
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
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
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "格式：2022/3/22"
      Height          =   255
      Left            =   4800
      TabIndex        =   13
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "输入库存："
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
      Index           =   5
      Left            =   600
      TabIndex        =   11
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "输入入库时间：（默认为今天）"
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
      Left            =   600
      TabIndex        =   10
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "输入出版社："
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
      Left            =   600
      TabIndex        =   9
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "输入专业："
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
      Left            =   600
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
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
      Left            =   600
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
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
      Left            =   600
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "新增书籍界面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

'先检查一遍数据都写了没
If Trim(Text1(0)) = empy Then MsgBox "请输入书号", 48, "错误": Exit Sub: Text1(0).SetFocus
If Trim(Text1(1)) = empy Then MsgBox "请输入书名", 48, "错误": Exit Sub: Text1(1).SetFocus
If Trim(Combo1.Text) = "这里输入A段的专业" And Trim(Combo2.Text) = "这里输入B段的专业" Then MsgBox "请选择专业", 48, "错误": Exit Sub
If Trim(Text1(3)) = empy Then MsgBox "请输入出版社信息", 48, "错误": Exit Sub: Text1(3).SetFocus
If Trim(Text1(4)) = empy Then Text1(4) = Date
If Trim(Text1(5)) = empy Then MsgBox "请输入库存", 48, "错误": Exit Sub: Text1(5).SetFocus
If Format(Trim(Text1(4)), "yyyy/m/d") <> Trim(Text1(4)) Then MsgBox "请输入正确的时间格式,eg:2022/3/22": Exit Sub

'因为combolbox控件有两个所以设置变量来平衡
Dim vocation As String
If Trim(Combo1.Text) <> "这里输入A段的专业" And Trim(Combo2.Text) <> "这里输入B段的专业" Then MsgBox "两个框只能选择一个哦！另一个请选择下拉第一个项", 48, "错误": Exit Sub
If Trim(Combo1.Text) = "这里输入A段的专业" Then vocation = Trim(Combo2.Text)
If Trim(Combo2.Text) = "这里输入B段的专业" Then vocation = Trim(Combo1.Text)

'增加排重机制
Dim rs As New ADODB.Recordset
Dim sql As String
sql = " select 书号,书名 from 书库系统 where 书号 = " & "'" & Trim(Text1(0)) & "'"
rs.Open sql, 主界面加查找界面.conn, 1, 3
    '这里开始正式排错
        If rs.EOF = False Then MsgBox "检查到重复，请认真排查书籍号": Exit Sub
rs.Close

'这里开始导入数据
rs.Open "书库系统", 主界面加查找界面.conn, 1, 3
    rs.AddNew
    rs.Fields("书号") = Trim(Text1(0))
    rs.Fields("书名") = Trim(Text1(1))
    rs.Fields("专业") = vocation
    rs.Fields("出版社") = Trim(Text1(3))
    rs.Fields("入库时间") = Trim(Text1(4))
    rs.Fields("库存") = Trim(Text1(5))
    rs.Update
rs.Close

'添加完后记得加入历史记录表
rs.Open "书籍信息变更记录表", 主界面加查找界面.conn, 1, 3
    rs.AddNew
    rs.Fields("书号") = Trim(Text1(0))
    rs.Fields("书名") = Trim(Text1(1))
    rs.Fields("专业") = vocation
    rs.Fields("更改时间") = Trim(Text1(4))
    rs.Fields("更改的类型") = "新增"
    rs.Fields(5) = "全部"
    rs.Fields(6) = "无"
    rs.Fields(7) = "有"
    rs.Update
rs.Close
MsgBox "添加成功，你加入的内容为：" & Chr(13) & "书号：" & Text1(0) & Space(5) & "书号：" & Text1(1) & Space(5) & "专业：" & vocation & Chr(13) & "出版社：" & Text1(3) & Space(5) & "入库时间：" & Text1(4) & Space(5) & "库存：" & Text1(5), , "提示"

'全部完成后退出此窗口
Unload Me
主界面加查找界面.Enabled = True
主界面加查找界面.Show
主界面加查找界面.maindata_load (" select * from  书库系统 order by " & Trim(主界面加查找界面.Text2) & " desc ")
End Sub

Private Sub Command2_Click()
Unload Me
主界面加查找界面.Enabled = True
主界面加查找界面.Show
主界面加查找界面.maindata_load (" select * from  书库系统 order by " & Trim(主界面加查找界面.Text2) & " desc ")
End Sub

Private Sub Form_Click()
'主界面加查找界面.conn
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 0
    主界面加查找界面.Enabled = True
    主界面加查找界面.Show
    主界面加查找界面.maindata_load (" select * from  书库系统 order by " & Trim(主界面加查找界面.Text2) & " desc ")
End Sub
