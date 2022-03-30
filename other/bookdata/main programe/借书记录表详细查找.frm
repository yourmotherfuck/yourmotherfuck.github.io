VERSION 5.00
Begin VB.Form 借书记录表详细查找 
   Caption         =   "借书记录表详细查找"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   7935
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
      Left            =   5160
      TabIndex        =   16
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox Text6 
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
      Left            =   5160
      TabIndex        =   13
      Text            =   "这里默认不勾选"
      Top             =   3000
      Width           =   1935
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
      Left            =   2760
      TabIndex        =   11
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
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
      Height          =   495
      Left            =   600
      TabIndex        =   10
      Top             =   4680
      Width           =   1815
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
      Left            =   2640
      TabIndex        =   9
      Top             =   3720
      Width           =   1935
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
      Left            =   2640
      TabIndex        =   8
      Top             =   3000
      Width           =   1935
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
      Left            =   2640
      TabIndex        =   7
      Top             =   2280
      Width           =   1935
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
      Left            =   2640
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
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
      Left            =   2640
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "时间格式，懂得都懂！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   15
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label7 
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
      Left            =   4800
      TabIndex        =   14
      Top             =   3120
      Width           =   240
   End
   Begin VB.Label Label6 
      Caption         =   "没什么规则，全部可选，最少输入一个列的数据，自动搜索"
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
      Left            =   5040
      TabIndex        =   12
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "借书数量"
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
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "借书时间"
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
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "借书人（单位）"
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
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
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
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
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
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "借书记录表详细查找"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 0 Then
'取消勾选时固定的
    Text6.Enabled = False
    Text6 = "这里默认不勾选"
    Text4 = empy
Else
    Text6.Enabled = True
    '两种情况，前面那段为空和不为空
    If Text4 = empy Then
            Text6.Text = Year(Date) & "/" & Month(Date) & "/"
            Text4.Text = Year(Date) & "/" & Month(Date) & "/"
    Else
        '如果起始时间不为空那么检查他的成分
        If IsDate(Trim(Text4)) = False Then MsgBox "请输入正确的时间格式！eg:2022/3/3", 48, "警告": Text4 = empy: Check1.Value = 0: Exit Sub
        '两种情况，如果满足isdate后是 #2022/2#格式 直接报错 ， #3/4# 格式时在前面补全上年份
        If Year(CDate(Text4)) & "/" & Month(Text4) = Trim(Text4) Then MsgBox "请输入正确的时间格式！eg:2022/3/3", 48, "警告": Text4 = empy: Check1.Value = 0: Exit Sub
        If CStr(Month(Trim(Text4))) + "/" + CStr(Day(Trim(Text4))) = Trim(Text4) Then
            Text4 = CStr(Year(Date)) + "/" + Trim(Text4)
            Text6.SetFocus
        Else
            Text4.SetFocus
        End If
        Text6 = CStr(Year(CDate(Trim(Text4)))) + "/"
    End If
End If
End Sub

Private Sub Command1_Click()
'先排错
If Trim(Text1) = empy And Trim(Text2) = empy And Trim(Text3) = empy And Trim(Text4) = empy And Trim(Text5) = empy Then MsgBox "不能全部为空值", 48, "警告": Exit Sub
If Trim(Text4) <> empy And IsDate(Trim(Text4)) = False Then MsgBox "请输入正确的时间格式！", 48, "警告": Exit Sub
If Trim(Text5) <> empy And CStr(Int(Val(Trim(Text5)))) <> Trim(Text5) Then MsgBox "借书数量请输入整数！", 48, "警告": Exit Sub

'时间排错
If Check1.Value = 1 Then
    '这个选项勾选后必须写时间，懂？
    If Trim(Text6) = empy Or Trim(Trim(Text4)) = empy Then MsgBox "勾选后两个文本框都必须输入时间！", 48, "警告！": Exit Sub
    '后面的时间不能大于前面的，懂？
    If CDate(Trim(Text6)) < CDate(Trim(Text4)) Then MsgBox "后面的时间要大于前面的！", 48, "警告！": Exit Sub
End If


Dim sql As String
'开始调配SQL语句
sql = "select * from 借书记录表 where "

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
        sql = sql & " 借书人（单位） like '%" & Trim(Text3.Text) & "%' "
    Else
        sql = sql & " and 借书人（单位） like '%" & Trim(Text3.Text) & "%' "
    End If
End If

If Trim(Text4.Text) <> empy Then
    If Trim(Text1) = empy And Trim(Text2) = empy And Trim(Text3) = empy Then
            If Check1.Value = 0 Then
                '如果不勾选时间区间选项
                sql = sql & " 借书时间 like #" & Trim(Text4.Text) & "# "
            Else
                '勾选后按区间来算
                sql = sql & " 借书时间 >= #" & Trim(Text4.Text) & "# and 借书时间 <= #" & Trim(Text6) & "# "
            End If
    Else
            If Check1.Value = 0 Then
                sql = sql & " and 借书时间 like #" & Trim(Text4.Text) & "# "
            Else
                sql = sql & " and 借书时间 >= #" & Trim(Text4.Text) & "# and 借书时间 <= #" & Trim(Text6) & "# "
            End If
    End If
End If

'If Trim(Text4) <> empy Then
'    If Trim(Text1) = empy And Trim(Text2) = empy And Trim(Text3) = empy Then
'        sql = sql & " 借书时间 like #" & Trim(Text4.Text) & "# "
'    Else
'        sql = sql & " and 借书时间 like #" & Trim(Text4.Text) & "# "
'    End If
'End If

If Trim(Text5) <> empy Then
    If Trim(Text1) = empy And Trim(Text2) = empy And Trim(Text3) = empy And Trim(Text4) = empy Then
        sql = sql & " 借书数量 = " & Trim(Text5.Text) & " "
    Else
        sql = sql & " and 借书数量 = " & Trim(Text5.Text) & " "
    End If
End If

sql = sql & "order by " & 借书记录查看窗口.Combo1.Text

借书记录查看窗口.senthistory_load sql, True

Unload Me
借书记录查看窗口.Show
借书记录查看窗口.Enabled = True
End Sub

Private Sub Command2_Click()
Unload Me
借书记录查看窗口.Show
借书记录查看窗口.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
借书记录查看窗口.Enabled = True
借书记录查看窗口.Show
End Sub


Private Sub Text4_GotFocus()
Text4.Text = Year(Date) & "/" & Month(Date) & "/"
End Sub
