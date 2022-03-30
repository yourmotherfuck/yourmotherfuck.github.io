VERSION 5.00
Begin VB.Form 详细查找界面 
   Caption         =   "详细查找界面"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   10155
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check1 
      Caption         =   "搜索时间区间"
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
      Left            =   7560
      TabIndex        =   18
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
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
      Left            =   5160
      TabIndex        =   16
      Text            =   "这里默认不勾选"
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
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
      Left            =   7320
      TabIndex        =   15
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查找"
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
      Left            =   5280
      TabIndex        =   14
      Top             =   5520
      Width           =   1575
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
      Left            =   2520
      TabIndex        =   12
      Top             =   4560
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
      Index           =   4
      Left            =   2520
      TabIndex        =   11
      Top             =   3840
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
      Index           =   3
      Left            =   2520
      TabIndex        =   10
      Top             =   3120
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
      Index           =   2
      Left            =   2520
      TabIndex        =   9
      Top             =   2400
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
      Index           =   1
      Left            =   2520
      TabIndex        =   8
      Top             =   1680
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
      Index           =   0
      Left            =   2520
      TabIndex        =   7
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "时间格子规则：勾选后悔自动补全当前年份，如果要实现早于和晚于效果在对应的地方输入一个很大的时间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5520
      TabIndex        =   19
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   4680
      TabIndex        =   17
      Top             =   3960
      Width           =   270
   End
   Begin VB.Label Label2 
      Caption         =   "注意：此界面采用松散的匹配方法，并且所有框都是可选的            不选的项会变成通配符*"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5520
      TabIndex        =   13
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "入库时间"
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
      Left            =   840
      TabIndex        =   6
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "库存"
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
      Left            =   840
      TabIndex        =   5
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "出版社"
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
      Left            =   840
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "专业"
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
      Left            =   840
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label1 
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
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "书名"
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
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "书号"
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
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "详细查找界面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 0 Then
'取消勾选时固定的
    Text2.Enabled = False
    Text2 = "这里默认不勾选"
    Text1(4) = empy
Else
    Text2.Enabled = True
    '两种情况，前面那段为空和不为空
    If Text1(4) = empy Then
            Text2.Text = Year(Date) & "/" & Month(Date) & "/"
            Text1(4).Text = CStr(Year(Date)) & "/" & Month(Date) & "/"
    Else
        '如果起始时间不为空那么检查他的成分
        If IsDate(Trim(Text1(4))) = False Then MsgBox "请输入正确的时间格式！eg:2022/3/3", 48, "警告": Text1(4) = empy: Check1.Value = 0: Exit Sub
        '两种情况，如果满足isdate后是 #2022/2#格式 直接报错 ， #3/4# 格式时在前面补全上年份
        If Year(CDate(Text1(4))) & "/" & Month(Text1(4)) = Trim(Text1(4)) Then MsgBox "请输入正确的时间格式！eg:2022/3/3", 48, "警告": Text1(4) = empy: Check1.Value = 0: Exit Sub
        If CStr(Month(Trim(Text1(4)))) + "/" + CStr(Day(Trim(Text1(4)))) = Trim(Text1(4)) Then
            Text1(4) = CStr(Year(Date)) + "/" + Trim(Text1(4))
            Text2.SetFocus
        Else
            Text1(4).SetFocus
        End If
        Text2 = CStr(Year(CDate(Trim(Text1(4))))) + "/"
    End If
End If
End Sub

Private Sub Command1_Click()
'老规矩，先排错
If Trim(Text1(5).Text) <> empy And Val(Trim(Text1(5).Text)) <> Trim(Text1(5).Text) Then MsgBox "想查找库存请输入纯数字", 48, "警告！": Exit Sub
If Trim(Text1(0).Text) = empy And Trim(Text1(1).Text) = empy And Trim(Text1(2).Text) = empy And Trim(Text1(3).Text) = empy And Trim(Text1(4).Text) = empy And Trim(Text1(5).Text) = empy Then MsgBox "请输入最少一个项", 48, "警告", 48, "警告！": Exit Sub

If Trim(Text1(4)) <> empy And IsDate(Trim(Text1(4).Text)) = False Then MsgBox "请输入正确的时间格式,eg:2022/3/22", 48, "警告！": Exit Sub

If Check1.Value = 1 Then
    '这个选项勾选后必须写时间，懂？
    If Trim(Text2) = empy Or Trim(Trim(Text1(4))) = empy Then MsgBox "勾选后两个文本框都必须输入时间！", 48, "警告！": Exit Sub
    '后面的时间不能大于前面的，懂？
    If CDate(Trim(Text2)) < CDate(Trim(Text1(4))) Then MsgBox "后面的时间要大于前面的！", 48, "警告！": Exit Sub
End If

Dim sql As String
sql = "select * from 书库系统 where "

If Trim(Text1(0).Text) <> empy Then sql = sql & " 书号 like '%" & Trim(Text1(0).Text) & "%'"

If Trim(Text1(1).Text) <> empy Then
    If Trim(Text1(0).Text) = empy Then
        sql = sql & " 书名 like '%" & Trim(Text1(1).Text) & "%'"
    Else
        sql = sql & " and 书名 like '%" & Trim(Text1(1).Text) & "%'"
    End If
End If

If Trim(Text1(2).Text) <> empy Then
    If Trim(Text1(0).Text) = empy And Trim(Text1(1).Text) = empy Then
        sql = sql & " 专业 like '%" & Trim(Text1(2).Text) & "%'"
    Else
        sql = sql & " and 专业 like '%" & Trim(Text1(2).Text) & "%'"
    End If
End If

If Trim(Text1(3).Text) <> empy Then
    If Trim(Text1(0).Text) = empy And Trim(Text1(1).Text) = empy And Trim(Text1(2).Text) = empy Then
        sql = sql & " 出版社 like '%" & Trim(Text1(3).Text) & "%'"
    Else
        sql = sql & " and 出版社 like '%" & Trim(Text1(3).Text) & "%'"
    End If
End If

If Trim(Text1(4).Text) <> empy Then
    If Trim(Text1(0).Text) = empy And Trim(Text1(1).Text) = empy And Trim(Text1(2).Text) = empy And Trim(Text1(3).Text) = empy Then
        If Check1.Value = 0 Then
            '如果不勾选时间区间选项
            sql = sql & " 入库时间 like #" & Trim(Text1(4).Text) & "# "
        Else
            '勾选后按区间来算
            sql = sql & " 入库时间 >= #" & Trim(Text1(4).Text) & "# and 入库时间 <= #" & Trim(Text2) & "# "
        End If
    Else
        If Check1.Value = 0 Then
            sql = sql & " and 入库时间 like #" & Trim(Text1(4).Text) & "# "
        Else
            sql = sql & " and 入库时间 >= #" & Trim(Text1(4).Text) & "# and 入库时间 <= #" & Trim(Text2) & "# "
        End If
    End If
End If

If Trim(Text1(5).Text) <> empy Then
    If Trim(Text1(0).Text) = empy And Trim(Text1(1).Text) = empy And Trim(Text1(2).Text) = empy And Trim(Text1(3).Text) = empy And Trim(Text1(4).Text) = empy Then
        sql = sql & " 库存 = " & Trim(Text1(5).Text)
    Else
        sql = sql & " and 库存 = " & Trim(Text1(5).Text)
    End If
End If
sql = sql + " order by " & Trim(主界面加查找界面.Combo2.Text)

Dim rs As New ADODb.Recordset
rs.Open sql, 主界面加查找界面.conn, 1, 3
    If rs.EOF Then MsgBox "书库中没有与之匹配的数据！请检查后再次输入", 48, "警告": Exit Sub
rs.Close

主界面加查找界面.maindata_load sql, True
Unload Me
主界面加查找界面.Enabled = True
主界面加查找界面.Show
End Sub

Private Sub Command2_Click()
Unload Me
主界面加查找界面.Enabled = True
主界面加查找界面.Show
'主界面加查找界面.maindata_load (" select * from  书库系统 order by " & Trim(主界面加查找界面.Combo2.Text) & " desc ")
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 0
    主界面加查找界面.Enabled = True
    主界面加查找界面.Show
    '主界面加查找界面.maindata_load (" select * from  书库系统 order by " & Trim(主界面加查找界面.Combo2.Text) & " desc ")
End Sub

