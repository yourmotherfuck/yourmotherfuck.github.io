VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form 书籍信息变更记录查看窗口 
   Caption         =   "书籍信息变更记录查看窗口"
   ClientHeight    =   9900
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15645
   LinkTopic       =   "Form1"
   ScaleHeight     =   9900
   ScaleWidth      =   15645
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "删除数据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      TabIndex        =   10
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16560
      TabIndex        =   9
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "修改数据（不计入历史记录的）"
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
      Left            =   13800
      TabIndex        =   8
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出查找模式"
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
      Left            =   10800
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
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
      Left            =   10800
      TabIndex        =   6
      Top             =   360
      Width           =   1695
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
      ItemData        =   "书籍信息变更记录查看窗口.frx":0000
      Left            =   7800
      List            =   "书籍信息变更记录查看窗口.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "书籍信息变更记录查看窗口.frx":007B
      Left            =   3240
      List            =   "书籍信息变更记录查看窗口.frx":0097
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   2295
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
      Left            =   7800
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6855
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   12091
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
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
      Caption         =   "降序Z-A"
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
      Left            =   5880
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "选择已哪个字段排序：默认入库时间"
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
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   2415
   End
   Begin VB.Menu 详细查找界面 
      Caption         =   "详细查找就界面"
   End
End
Attribute VB_Name = "书籍信息变更记录查看窗口"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public last_strsql As String
Public last_sort As String
Public first As Boolean

Public Sub bookhistory_load(strsql As String, is_sort As Boolean)
last_strsql = strsql
    If is_sort = True Then
            If Label2.Caption = "降序Z-A" Then
                strsql = strsql & " desc"
            ElseIf Label2.Caption = "升序A-Z" Then
                strsql = strsql & " asc"
            End If
    End If
        'MsgBox strsql
    Dim rs As New ADODb.Recordset
    rs.Open strsql, 主界面加查找界面.conn, 1, 3
    
    If is_sort Then
        If rs.EOF Then MsgBox "没有与之匹配的数据", 48, "错误": Exit Sub
    End If
    
    Set DataGrid1.DataSource = rs
        DataGrid1.ScrollBars = dbgVertical
        DataGrid1.Width = 书籍信息变更记录查看窗口.Width - DataGrid1.Left * 2
        DataGrid1.Height = 书籍信息变更记录查看窗口.Height - DataGrid1.Top * 1.7
        
            '同步完设置datagrid的列宽
            For i = 0 To DataGrid1.Columns.Count - 1
                DataGrid1.Columns(i).Width = (DataGrid1.Width / DataGrid1.Columns.Count) - 100
            Next i
    
End Sub

Private Sub Combo1_Click()
If first = True Then
    'MsgBox last_strsql
        last_strsql = Mid(last_strsql, 1, Len(last_strsql) - Len(last_sort) - 1)
    'MsgBox last_strsql
                bookhistory_load last_strsql & " " & Trim(Combo1.Text), True
            last_sort = Trim(Combo1.Text)
End If
End Sub

Private Sub Command1_Click()
If Combo2.Text = "请选择要查找的字段" Then MsgBox "请选择要查找的列名": Exit Sub

If Combo2.Text = "更改时间" Then
    If IsDate(Trim(Text1)) = False Then MsgBox "请输入正确的时间格式！eg:2022/3/3", 48, "警告": Exit Sub
    bookhistory_load "select * from 书籍信息变更记录表 where " & Combo2.Text & " like #" & Text1 & "# order by " & Combo1.Text, True
Else
    bookhistory_load "select * from 书籍信息变更记录表 where " & Combo2.Text & " like '%" & Text1 & "%' order by " & Combo1.Text & " ", True
End If

End Sub

Private Sub Command2_Click()
书籍信息变更记录查看窗口.bookhistory_load "select * from 书籍信息变更记录表 order by " & 书籍信息变更记录查看窗口.Combo1.Text, True
Text1 = ""
End Sub

Private Sub Command3_Click()
书籍记录表修改数据界面.Show
书籍信息变更记录查看窗口.Enabled = False
End Sub

Private Sub Command4_Click()
Unload Me
主界面加查找界面.Enabled = True
主界面加查找界面.Show
End Sub



Private Sub Command5_Click()
a = MsgBox("您确定要删除吗，删除后数据无法找回切不会记录到任何表中！", vbYesNo, "警告")
If a = vbNo Then Exit Sub

    'MsgBox " delete from 书籍信息变更记录表 where 书号 = '" & DataGrid1.Columns(0).CellText(DataGrid1.Bookmark) & "' " & " and 书名 = '" & DataGrid1.Columns("书名").CellText(DataGrid1.Bookmark) & "' " & " and 专业 = '" & DataGrid1.Columns("专业").CellText(DataGrid1.Bookmark) & "' " & " and 更改时间 = #" & DataGrid1.Columns("更改时间").CellText(DataGrid1.Bookmark) & "# " & " and 更改的类型 = '" & DataGrid1.Columns("更改的类型").CellText(DataGrid1.Bookmark) & "' " & " and 变更的数据 = '" & DataGrid1.Columns("变更的数据").CellText(DataGrid1.Bookmark) & "' " & " and 更改前 = '" & DataGrid1.Columns("更改前").CellText(DataGrid1.Bookmark) & "' " & " and 更改后 = '" & DataGrid1.Columns("更改后").CellText(DataGrid1.Bookmark) & "' "
    bookhistory_load " delete from 书籍信息变更记录表 where 书号 = '" & DataGrid1.Columns("书号").CellText(DataGrid1.Bookmark) & "' " & " and 书名 = '" & DataGrid1.Columns("书名").CellText(DataGrid1.Bookmark) & "' " & " and 专业 = '" & DataGrid1.Columns("专业").CellText(DataGrid1.Bookmark) & "' " & " and 更改时间 = #" & DataGrid1.Columns("更改时间").CellText(DataGrid1.Bookmark) & "# " & " and 更改的类型 = '" & DataGrid1.Columns("更改的类型").CellText(DataGrid1.Bookmark) & "' " & " and 变更的数据 = '" & DataGrid1.Columns("变更的数据").CellText(DataGrid1.Bookmark) & "' " & " and 更改前 = '" & DataGrid1.Columns("更改前").CellText(DataGrid1.Bookmark) & "' " & " and 更改后 = '" & DataGrid1.Columns("更改后").CellText(DataGrid1.Bookmark) & "' ", False
    MsgBox "删除成功"
    
    '删除后重新加载
    bookhistory_load "select * from 书籍信息变更记录表 order by " & 书籍信息变更记录查看窗口.Combo1.Text, True
End Sub

Private Sub DataGrid1_DblClick()
书籍记录表修改数据界面.Text1 = DataGrid1.Columns(书籍记录表修改数据界面.Label1.Caption).CellText(DataGrid1.Bookmark)
书籍记录表修改数据界面.Text2 = DataGrid1.Columns(书籍记录表修改数据界面.Label2.Caption).CellText(DataGrid1.Bookmark)
书籍记录表修改数据界面.Text3 = DataGrid1.Columns(书籍记录表修改数据界面.Label3.Caption).CellText(DataGrid1.Bookmark)
书籍记录表修改数据界面.Text4 = DataGrid1.Columns(书籍记录表修改数据界面.Label4.Caption).CellText(DataGrid1.Bookmark)
书籍记录表修改数据界面.Text5 = DataGrid1.Columns(书籍记录表修改数据界面.Label5.Caption).CellText(DataGrid1.Bookmark)
书籍记录表修改数据界面.Text6 = DataGrid1.Columns(书籍记录表修改数据界面.Label6.Caption).CellText(DataGrid1.Bookmark)
书籍记录表修改数据界面.Text7 = DataGrid1.Columns(书籍记录表修改数据界面.Label7.Caption).CellText(DataGrid1.Bookmark)
书籍记录表修改数据界面.Text8 = DataGrid1.Columns(书籍记录表修改数据界面.Label8.Caption).CellText(DataGrid1.Bookmark)
书籍记录表修改数据界面.Show
书籍记录表修改数据界面.Combo2.Text = DataGrid1.Columns(DataGrid1.Col).DataField
书籍记录表修改数据界面.Text10.SetFocus
书籍信息变更记录查看窗口.Enabled = False

End Sub

Private Sub Form_Activate()
DataGrid1.ScrollBars = dbgVertical
DataGrid1.Width = 书籍信息变更记录查看窗口.Width - DataGrid1.Left * 2
DataGrid1.Height = 书籍信息变更记录查看窗口.Height - DataGrid1.Top * 1.7

    '同步完设置datagrid的列宽
    For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).Width = (DataGrid1.Width / DataGrid1.Columns.Count) - 100
    Next i
End Sub

Private Sub Form_Load()
Combo1.Text = Combo1.List(3)
Combo2.Text = Combo2.List(0)
书籍信息变更记录查看窗口.WindowState = 2
bookhistory_load "select * from 书籍信息变更记录表 order by " & 书籍信息变更记录查看窗口.Combo1.Text, True
书籍信息变更记录查看窗口.first = True
书籍信息变更记录查看窗口.last_sort = 书籍信息变更记录查看窗口.Combo1.Text

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
主界面加查找界面.Enabled = True
主界面加查找界面.Show
End Sub

Private Sub Label2_Click()
If Label2.Caption = "降序Z-A" Then
    Label2.Caption = "升序A-Z"
ElseIf Label2.Caption = "升序A-Z" Then
    Label2.Caption = "降序Z-A"
End If

'用上一次的sql语句后面跟上排序语句
    'MsgBox last_strsql
        last_strsql = Mid(last_strsql, 1, Len(last_strsql) - Len(last_sort) - 1)
    'MsgBox last_strsql
                bookhistory_load last_strsql & " " & Trim(Combo1.Text), True
            last_sort = Trim(Combo1.Text)
End Sub

Private Sub 详细查找界面_Click()
书籍记录表详细查找界面.Show
书籍信息变更记录查看窗口.Enabled = False
End Sub
