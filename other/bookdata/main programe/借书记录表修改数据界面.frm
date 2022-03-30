VERSION 5.00
Begin VB.Form 借书记录表修改数据界面 
   Caption         =   "借书记录表修改数据界面"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   12375
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "更改"
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
      Height          =   735
      Left            =   6960
      TabIndex        =   4
      Top             =   5520
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
      Height          =   735
      Left            =   9360
      TabIndex        =   3
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "精准查找窗口"
      Height          =   5175
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   6015
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
         TabIndex        =   13
         Top             =   840
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
         Left            =   2760
         TabIndex        =   12
         Top             =   1560
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
         Left            =   2760
         TabIndex        =   11
         Top             =   2280
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
         Left            =   2760
         TabIndex        =   10
         Top             =   3000
         Width           =   1935
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
         TabIndex        =   9
         Top             =   3720
         Width           =   1935
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
         Left            =   720
         TabIndex        =   18
         Top             =   840
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
         Left            =   720
         TabIndex        =   17
         Top             =   1560
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
         Left            =   720
         TabIndex        =   16
         Top             =   2280
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
         Left            =   720
         TabIndex        =   15
         Top             =   3000
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
         Left            =   720
         TabIndex        =   14
         Top             =   3720
         Width           =   1695
      End
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   9240
      TabIndex        =   1
      Top             =   4680
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "借书记录表修改数据界面.frx":0000
      Left            =   9120
      List            =   "借书记录表修改数据界面.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Label100 
      Caption         =   "这里的数据更改不会记录到任何表中，没有还原的可能，请谨慎更改或删除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7680
      TabIndex        =   8
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label99 
      Caption         =   "前面这些是用于校准行列的"
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
      Left            =   7680
      TabIndex        =   7
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label Label11 
      Caption         =   "将数据变更为："
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
      Left            =   7080
      TabIndex        =   6
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label12 
      Caption         =   "要变更的列名："
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
      Left            =   7080
      TabIndex        =   5
      Top             =   3960
      Width           =   1695
   End
End
Attribute VB_Name = "借书记录表修改数据界面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'老规矩先排错（在这里面所有的东西都是必选的）
If Trim(Text1) = empy Then MsgBox "请输入校准的数据，（全部都要）！", 48, "警告": Exit Sub
If Trim(Text2) = empy Then MsgBox "请输入校准的数据，（全部都要）！", 48, "警告": Exit Sub
If Trim(Text3) = empy Then MsgBox "请输入校准的数据，（全部都要）！", 48, "警告": Exit Sub
If Trim(Text4) = empy Then MsgBox "请输入校准的数据，（全部都要）！", 48, "警告": Exit Sub
If IsDate(Trim(Text4)) = False Then MsgBox "请输入正确的时间格式", 48, "警告": Exit Sub
If Trim(Text5) = empy Then MsgBox "请输入校准的数据，（全部都要）！", 48, "警告": Exit Sub

If Trim(Combo2.Text) = Combo2.List(0) Then MsgBox "请选择要更改数据的列名！", 48, "警告": Exit Sub
If Trim(Combo2.Text) = Combo2.List(4) And IsDate(Trim(Text10)) = False Then MsgBox "更改后的数据请为时间格式，eg:2022/3/3", 48, "警告": Exit Sub
If Trim(Text10) = empy Then MsgBox "请输入更改后的数据", 48, "警告": Exit Sub

'先确定是否有这行
Dim temp, sql As String
Dim rs As New ADODb.Recordset
temp = " where " & Trim(Label1.Caption) & " ='" & Trim(Text1) & "' and " & Trim(Label2.Caption) & " ='" & Trim(Text2) & "' and " & Trim(Label3.Caption) & " ='" & Trim(Text3) & "' and " & Trim(Label4.Caption) & " = #" & Trim(Text4) & "# and " & Trim(Label5.Caption) & " =" & Trim(Text5)
'先确认有没有这一行
rs.Open "select * from 借书记录表 " & temp, 主界面加查找界面.conn, 1, 3
    If rs.EOF = True Then MsgBox "没有此行数据，请确认后重新输入！", 48, "警告": Exit Sub
    If CStr(rs.Fields(Combo2.Text).Value) = Trim(Text10) Then MsgBox "更改前和更改后的数据一样！", 48, "警告": Exit Sub
rs.Close

a = MsgBox("您确定要删除吗？删除后数据无法找回，请谨慎删除！", vbYesNo, "警告")
If a = vbYes Then
    If Combo2.Text = Combo2.List(4) Then
        sql = "update 借书记录表 set " & Combo2.Text & " = #" & Trim(Text10) & "# " & temp
    ElseIf Combo2.Text = Combo2.List(5) Then
        sql = "update 借书记录表 set " & Combo2.Text & " = " & Trim(Text10) & " " & temp
    Else
        sql = "update 借书记录表 set " & Combo2.Text & " = '" & Trim(Text10) & "' " & temp
    End If

     借书记录查看窗口.senthistory_load sql, False
Else
    Exit Sub
End If

MsgBox "数据更改成功!", , "注意"
Unload Me
借书记录查看窗口.Enabled = True
借书记录查看窗口.senthistory_load "select * from 借书记录表 order by " & 借书记录查看窗口.Combo1.Text, True
借书记录查看窗口.Show
End Sub

Private Sub Command2_Click()
Unload Me
借书记录查看窗口.Show
借书记录查看窗口.Enabled = True
End Sub

Private Sub Form_Load()
Combo2.Text = Combo2.List(0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
借书记录查看窗口.Show
借书记录查看窗口.Enabled = True
End Sub

