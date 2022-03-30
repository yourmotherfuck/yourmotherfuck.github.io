VERSION 5.00
Begin VB.Form 书籍记录表详细查找界面 
   Caption         =   "0"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   10995
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
      Left            =   8040
      TabIndex        =   22
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
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
      Left            =   5520
      TabIndex        =   21
      Text            =   "这里默认不勾选"
      Top             =   3000
      Width           =   2055
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
      Height          =   735
      Left            =   7800
      TabIndex        =   18
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "开始查找"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      TabIndex        =   17
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox Text8 
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
      Left            =   2760
      TabIndex        =   16
      Top             =   6360
      Width           =   2055
   End
   Begin VB.TextBox Text7 
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
      Left            =   2760
      TabIndex        =   15
      Top             =   5520
      Width           =   2055
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
      Left            =   2760
      TabIndex        =   14
      Top             =   4680
      Width           =   2055
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
      Left            =   2760
      TabIndex        =   13
      Top             =   3840
      Width           =   2055
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
      Left            =   2760
      TabIndex        =   12
      Top             =   3000
      Width           =   2055
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
      Left            =   2760
      TabIndex        =   11
      Top             =   2160
      Width           =   2055
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
      Left            =   2760
      TabIndex        =   10
      Top             =   1320
      Width           =   2055
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
      Left            =   2760
      TabIndex        =   9
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   20
      Top             =   3120
      Width           =   240
   End
   Begin VB.Label Label10 
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
      Left            =   6000
      TabIndex        =   19
      Top             =   4080
      Width           =   3135
   End
   Begin VB.Label Label9 
      Caption         =   "更改后："
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
      Left            =   600
      TabIndex        =   8
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "更改前："
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
      Left            =   600
      TabIndex        =   7
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "变更的数据："
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
      Left            =   600
      TabIndex        =   6
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "更改的类型："
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
      Left            =   600
      TabIndex        =   5
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label4 
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
      Left            =   600
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label3 
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
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "书名："
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
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "书号："
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
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "没什么规则，全部可选，最少输入一个列的数据，自动搜索"
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
      Left            =   6000
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
End
Attribute VB_Name = "书籍记录表详细查找界面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 0 Then
'取消勾选时固定的
    Text9.Enabled = False
    Text9 = "这里默认不勾选"
    Text4 = empy
Else
    Text9.Enabled = True
    '两种情况，前面那段为空和不为空
    If Text4 = empy Then
            Text9.Text = Year(Date) & "/" & Month(Date) & "/"
            Text4.Text = Year(Date) & "/" & Month(Date) & "/"
    Else
        '如果起始时间不为空那么检查他的成分
        If IsDate(Trim(Text4)) = False Then MsgBox "请输入正确的时间格式！eg:2022/3/3", 48, "警告": Text4 = empy: Check1.Value = 0: Exit Sub
        '两种情况，如果满足isdate后是 #2022/2#格式 直接报错 ， #3/4# 格式时在前面补全上年份
        If Year(CDate(Text4)) & "/" & Month(Text4) = Trim(Text4) Then MsgBox "请输入正确的时间格式！eg:2022/3/3", 48, "警告": Text4 = empy: Check1.Value = 0: Exit Sub
        If CStr(Month(Trim(Text4))) + "/" + CStr(Day(Trim(Text4))) = Trim(Text4) Then
            Text4 = CStr(Year(Date)) + "/" + Trim(Text4)
            Text9.SetFocus
        Else
            Text4.SetFocus
        End If
        Text9 = CStr(Year(CDate(Trim(Text4)))) + "/"
    End If
End If
End Sub

Private Sub Command1_Click()
书籍信息变更记录查看窗口.Text1 = ""
'先排错
If Trim(Text1) = empy And Trim(Text9) = empy And Trim(Text3) = empy And Trim(Text4) = empy And Trim(Text5) = empy And Trim(Text6) = empy And Trim(Text7) = empy And Trim(Text8) = empy Then MsgBox "不能全部为空值", 48, "警告": Exit Sub
If Trim(Text4) <> empy And IsDate(Trim(Text4)) = False Then MsgBox "请输入正确的时间格式！", 48, "警告": Exit Sub
'时间排错
If Check1.Value = 1 Then
    '这个选项勾选后必须写时间，懂？
    If Trim(Text9) = empy Or Trim(Trim(Text4)) = empy Then MsgBox "勾选后两个文本框都必须输入时间！", 48, "警告！": Exit Sub
    '后面的时间不能大于前面的，懂？
    If CDate(Trim(Text9)) < CDate(Trim(Text4)) Then MsgBox "后面的时间要大于前面的！", 48, "警告！": Exit Sub
End If


Dim sql As String
'开始调配SQL语句
sql = "select * from 书籍信息变更记录表 where "

If Trim(Text1) <> empy Then sql = sql & " 书号 like '%" & Trim(Text1.Text) & "%'"

If Trim(Text2) <> empy Then
    If Trim(Text1) = empy Then
        sql = sql & " 书名 like '%" & Trim(Text2.Text) & "%' "
    Else
        sql = sql & " and 书名 like '%" & Trim(Text2.Text) & "%' "
    End If
End If

If Trim(Text3) <> empy Then
    If Trim(Text1) = empy And Trim(Text2) = empy Then
        sql = sql & " 专业 like '%" & Trim(Text3.Text) & "%' "
    Else
        sql = sql & " and 专业 like '%" & Trim(Text3.Text) & "%' "
    End If
End If

If Trim(Text4.Text) <> empy Then
    If Trim(Text1) = empy And Trim(Text2) = empy And Trim(Text3) = empy Then
            If Check1.Value = 0 Then
                '如果不勾选时间区间选项
                sql = sql & " 更改时间 like #" & Trim(Text4.Text) & "# "
            Else
                '勾选后按区间来算
                sql = sql & " 更改时间 >= #" & Trim(Text4.Text) & "# and 更改时间 <= #" & Trim(Text9) & "# "
            End If
    Else
            If Check1.Value = 0 Then
                sql = sql & " and 更改时间 like #" & Trim(Text4.Text) & "# "
            Else
                sql = sql & " and 更改时间 >= #" & Trim(Text4.Text) & "# and 更改时间 <= #" & Trim(Text9) & "# "
            End If
    End If
End If

'If Trim(Text4) <> empy Then
'    If Trim(Text1) = empy And Trim(Text2) = empy And Trim(Text3) = empy Then
'        sql = sql & " 更改时间 like #" & Trim(Text4.Text) & "# "
'    Else
'        sql = sql & " and 更改时间 like #" & Trim(Text4.Text) & "# "
'    End If
'End If

If Trim(Text5) <> empy Then
    If Trim(Text1) = empy And Trim(Text2) = empy And Trim(Text3) = empy And Trim(Text4) = empy Then
        sql = sql & " 更改的类型 like '%" & Trim(Text5.Text) & "%' "
    Else
        sql = sql & " and 更改的类型 like '%" & Trim(Text5.Text) & "%' "
    End If
End If

If Trim(Text6) <> empy Then
    If Trim(Text1) = empy And Trim(Text2) = empy And Trim(Text3) = empy And Trim(Text4) = empy And Trim(Text5) = empy Then
        sql = sql & " 变更的数据 like '%" & Trim(Text5.Text) & "%' "
    Else
        sql = sql & " and 变更的数据 like '%" & Trim(Text5.Text) & "%' "
    End If
End If

If Trim(Text7) <> empy Then
    If Trim(Text1) = empy And Trim(Text2) = empy And Trim(Text3) = empy And Trim(Text4) = empy And Trim(Text5) = empy And Trim(Text6) = empy Then
        sql = sql & " 更改前 like '%" & Trim(Text5.Text) & "%' "
    Else
        sql = sql & " and 更改前 like '%" & Trim(Text5.Text) & "%' "
    End If
End If

If Trim(Text8) <> empy Then
    If Trim(Text1) = empy And Trim(Text2) = empy And Trim(Text3) = empy And Trim(Text4) = empy And Trim(Text5) = empy And Trim(Text6) = empy And Trim(Text7) = empy Then
        sql = sql & " 更改后 like '%" & Trim(Text5.Text) & "%' "
    Else
        sql = sql & " and 更改后 like '%" & Trim(Text5.Text) & "%' "
    End If
End If

sql = sql & "order by " & 书籍信息变更记录查看窗口.Combo1.Text

书籍信息变更记录查看窗口.bookhistory_load sql, True

Unload Me
书籍信息变更记录查看窗口.Enabled = True
书籍信息变更记录查看窗口.Show
End Sub

Private Sub Command2_Click()
Unload Me
书籍信息变更记录查看窗口.Show
书籍信息变更记录查看窗口.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
书籍信息变更记录查看窗口.Enabled = True
书籍信息变更记录查看窗口.Show
End Sub

Private Sub Text4_GotFocus()
Text4.Text = Year(Date) & "/" & Month(Date) & "/"
End Sub
