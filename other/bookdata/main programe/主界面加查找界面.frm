VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form 主界面加查找界面 
   Caption         =   "主界面"
   ClientHeight    =   9705
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   16815
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   16815
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command6 
      Caption         =   "退出查找状态"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15000
      TabIndex        =   10
      Top             =   840
      Width           =   1215
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
      ItemData        =   "主界面加查找界面.frx":0000
      Left            =   12480
      List            =   "主界面加查找界面.frx":0019
      TabIndex        =   9
      Text            =   "选择要查找的字段"
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   8880
      TabIndex        =   8
      Text            =   "入库时间"
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "test"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "查找书籍"
      Height          =   495
      Left            =   15000
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      TabIndex        =   4
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "修改数据"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "新建借书"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "删除此书"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7455
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   13150
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   18
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "选择以哪个字段排序：默认是入库时间"
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
      Left            =   6360
      TabIndex        =   7
      Top             =   600
      Width           =   2175
   End
   Begin VB.Menu 书籍管理 
      Caption         =   "书籍管理"
      Begin VB.Menu 添加书籍 
         Caption         =   "添加书籍"
      End
      Begin VB.Menu 删除书籍 
         Caption         =   "删除书籍"
      End
      Begin VB.Menu 查找书籍 
         Caption         =   "查找书籍（详细）"
      End
   End
   Begin VB.Menu 借书系统 
      Caption         =   "借书系统"
      Begin VB.Menu 新增借书记录 
         Caption         =   "新增借书记录"
      End
      Begin VB.Menu 新增还书记录 
         Caption         =   "新增还书记录"
      End
      Begin VB.Menu 修改记录 
         Caption         =   "修改记录（不登记历史记录）"
      End
   End
   Begin VB.Menu 历史记录查询 
      Caption         =   "历史记录查询"
      Begin VB.Menu 借书或还书记录查询 
         Caption         =   "借书或还书记录查询"
      End
      Begin VB.Menu 书籍变更记录查询 
         Caption         =   "书籍变更记录查询"
      End
   End
End
Attribute VB_Name = "主界面加查找界面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public conn As New ADODB.Connection


Public Sub maindata_load(strsql As String)
            'MsgBox strsql
    Dim nothiscolumn As Boolean
    Dim rs As New ADODB.Recordset
    rs.Open strsql, conn, 1, 3
    For Each i In rs.Fields
        If Text2 = i.Name Then nothiscolumn = True
    Next i
    If nothiscolumn Then
        If rs.EOF Then MsgBox "没有与之匹配的数据", 48, "错误": Exit Sub
        Set DataGrid1.DataSource = rs
    Else
        MsgBox "数据库中没有这个列名，请检查后输入", 48, "注意"
    End If
End Sub


Private Sub Command4_Click()
If Combo1.Text = "选择要查找的字段" Then MsgBox "请先输入要查找的字段": Exit Sub
If Combo1.Text = "入库时间" And Trim(Text1.Text) <> Format(Trim(Text1.Text), "yyyy/m/d") Then MsgBox "请输入正确的日期格式。如：2022/3/3", 48, "注意": Exit Sub
If Combo1.Text = "库存" And Val(Trim(Text1.Text)) <> Trim(Text1.Text) Then MsgBox "想查找库存请输入纯数字": Exit Sub
If Combo1.Text = "入库时间" Then
    maindata_load ("select * from 书库系统 where 入库时间 = #" & Trim(Text1.Text) & "#")
ElseIf Combo1.Text = "库存" Then
    maindata_load (" select * from 书库系统 where 库存 =  " & Val(Trim(Text1.Text)))
Else
    maindata_load (" select * from 书库系统 where " & Trim(Combo1.Text) & " = '" & Trim(Text1.Text) & "'")
End If

End Sub

Private Sub Command5_Click()
maindata_load ("select * from 书库系统 where 入库时间 = #2020/3/3# ")
End Sub

Private Sub Command6_Click()
maindata_load (" select * from  书库系统 order by " & Trim(Text2) & " desc ")
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    yes = MsgBox("你确定要更改数据吗？", vbYesNo, "注意")
    If yes = vbYes Then
        Dim rs As New ADODB.Recordset
        rs.Open "书库系统", conn, 1, 3
            rs.Update
        rs.Close
        maindata_load (" select * from  书库系统 order by " & Trim(Text2) & " desc ")
        MsgBox "更改成功"
    End If
End If
End Sub

Private Sub Form_Load()
'这里连接数据库
conn.Provider = "Microsoft.Jet.OLEDB.4.0"              '这行默认不变
conn.Open "..\value\书库.mdb"                             '这里选择数据库路径

'用来搞什么什么光标的东西
conn.CursorLocation = adUseClient
    Dim rs As New ADODB.Recordset
    rs.Open " select * from  书库系统 order by " & Trim(Text2) & " desc ", conn, 1, 3
    Set DataGrid1.DataSource = rs
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
a = MsgBox("确定要退出吗？", vbYesNo, "QAQ")
If a = vbYes Then
    Cancel = 0
Else
    Cancel = 1
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then maindata_load (" select * from  书库系统 order by " & Trim(Text2) & " desc ")
End Sub

Private Sub 查找书籍_Click()
详细查找界面.Show
主界面加查找界面.Enabled = False
End Sub

Private Sub 添加书籍_Click()
新增书籍界面.Show
主界面加查找界面.Enabled = False
End Sub

Private Sub 新增借书记录_Click()
借书界面.Show
主界面加查找界面.Enabled = False
End Sub
