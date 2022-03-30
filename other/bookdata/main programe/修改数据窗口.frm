VERSION 5.00
Begin VB.Form 修改数据窗口 
   Caption         =   "修改数据窗口"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   10695
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "修改数据窗口.frx":0000
      Left            =   3000
      List            =   "修改数据窗口.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3600
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "补全"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   16
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   14
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   13
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "使用补全只需要输入书名或书号就可以了，列名写上后才会去匹配变更前的数据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   18
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "书籍号，列名，变更后必写，书名，专业，变更前可以不写（补全）,变更时间默认今天变更前和专业选填"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6360
      TabIndex        =   15
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "变更后："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   720
      TabIndex        =   6
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "变更前："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   720
      TabIndex        =   5
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "变更数据的列名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
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
      Caption         =   "更改时间："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
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
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "专业："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
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
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "输入书名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
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
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "请输入书籍号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
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
      Width           =   1695
   End
End
Attribute VB_Name = "修改数据窗口"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'老规矩先排错，变更前和专业选填
If Trim(Text1) = empy Then MsgBox "请输入书籍号！", 48, "警告": Exit Sub
If Trim(Text2) = empy Then MsgBox "请输入书名！", 48, "警告": Exit Sub
If Trim(Text4) = empy Then Text4 = Date
If Combo1.Text = Combo1.List(0) Then MsgBox "请选择要更改的列名", 48, "警告": Exit Sub
If Trim(Text6) = empy Then MsgBox "请输入变更后的数！", 48, "警告": Exit Sub
If Trim(Text4) <> empy And IsDate(Trim(Text4)) = False Then MsgBox "请输入正确的时间格式", 48, "警告"
If Combo1.Text = "库存" Then If CStr(Int((Val(Trim(Text6))))) <> Trim(Text6) Then MsgBox "更改后的数据请输入整数！", 48, "警告"
If Combo1.Text = "入库时间" Then If IsDate(Trim(Text6)) = -False Then MsgBox "更改后的数据也要按照时间格式写 eg:2022/3/3", 48, "警告"

Dim rs As New ADODb.Recordset
Dim sql As String
'这里要先确定有这个书，顺便告诉他们这本书的内容
sql = "select * from 书库系统  where 书号 ='" & Trim(Text1) & "'"
rs.Open sql, 主界面加查找界面.conn, 1, 3
    If rs.EOF = True Then MsgBox "书库系统里面没有对应的书号，请确认后再次输入", 48, "警告": Exit Sub
    MsgBox "此行内容为:" & Chr(13) & " 书号=" & rs.Fields("书号").Value & "  书名=" & rs.Fields("书名").Value & "  专业=" & rs.Fields("专业").Value & "  " & Chr(13) & "出版社=" & rs.Fields("出版社").Value & "  入库时间=" & rs.Fields("入库时间").Value & "  库存=" & rs.Fields("库存").Value
rs.Close

'这里开始更改数据
a = MsgBox("书号为% " & Trim(Text1) & " %的" & Combo1.Text & "数据将会变为  " & Trim(Text6) & "  " & Chr(13) & "你确定要这么做吗？", vbYesNo, "注意")
If a = vbNo Then Exit Sub

If Combo1.Text = "入库时间" Then
    sql = " update 书库系统 set " & Combo1.Text & " = #" & Trim(Text6) & "# where 书号 = '" & Trim(Text1) & "'"
ElseIf Combo1.Text = "库存" Then
    sql = " update 书库系统 set " & Combo1.Text & " = " & Trim(Text6) & " where 书号 = '" & Trim(Text1) & "'"
Else
    sql = " update 书库系统 set " & Combo1.Text & " = '" & Trim(Text6) & "' where 书号 = '" & Trim(Text1) & "'"
End If
rs.Open sql, 主界面加查找界面.conn, 1, 3


'修改完数据后记得加入到信息历史记录表里面
rs.Open "书籍信息变更记录表", 主界面加查找界面.conn, 1, 3
    rs.AddNew
    rs.Fields("书号") = Trim(Text1)
    rs.Fields("书名") = Trim(Text2)
    rs.Fields("专业") = Trim(Text3)
    rs.Fields("更改时间") = Trim(Text4)
    rs.Fields("更改的类型") = "手动修改数据"
    rs.Fields("变更的数据") = Combo1.Text
    rs.Fields("更改前") = Trim(Text5)
    rs.Fields("更改后") = Trim(Text6)
    rs.Update
rs.Close

'全部完成后退出此窗口
Unload Me
主界面加查找界面.Enabled = True
主界面加查找界面.Show
主界面加查找界面.maindata_load " select * from  书库系统 order by " & Trim(主界面加查找界面.Combo2.Text), True
End Sub

Private Sub Command2_Click()
Unload Me
主界面加查找界面.Enabled = True
主界面加查找界面.Show
End Sub

Private Sub Command3_Click()
If Trim(Text1) = empy And Trim(Text2) = empy Then MsgBox "请输入书名或书号！", 48, "警告": Exit Sub
If Trim(Text4) = empy Then Text4 = Date

Dim rs As New ADODb.Recordset
Dim sql As String
If Trim(Text2) <> empy Then sql = "select * from 书库系统 where 书名 = '" & Trim(Text2) & "'"
If Trim(Text1) <> empy Then sql = "select * from 书库系统 where 书号 = '" & Trim(Text1) & "'"

rs.Open sql, 主界面加查找界面.conn, 1, 3
    Text2 = rs.Fields("书名").Value
    Text3 = rs.Fields("专业").Value
    If Combo1.Text <> Combo1.List(0) Then Text5 = rs.Fields(Combo1.Text).Value
rs.Close
End Sub


Private Sub Form_Load()
Combo1.Text = Combo1.List(0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 0
Unload Me
主界面加查找界面.Enabled = True
主界面加查找界面.Show
End Sub
