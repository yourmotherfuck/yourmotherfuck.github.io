VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form 主界面加查找界面 
   Caption         =   "主界面"
   ClientHeight    =   10680
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17400
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
   ScaleHeight     =   10680
   ScaleWidth      =   17400
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command7 
      Caption         =   "退出"
      Height          =   615
      Left            =   17520
      TabIndex        =   11
      Top             =   240
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   8175
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   14420
      _Version        =   393216
      AllowUpdate     =   -1  'True
      DefColWidth     =   40
      HeadLines       =   1
      RowHeight       =   18
      RowDividerStyle =   6
      AllowDelete     =   -1  'True
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
   Begin VB.ComboBox Combo2 
      Height          =   360
      ItemData        =   "主界面加查找界面.frx":0000
      Left            =   8280
      List            =   "主界面加查找界面.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   720
      Width           =   1815
   End
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
      Left            =   14880
      TabIndex        =   8
      Top             =   960
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
      ItemData        =   "主界面加查找界面.frx":0044
      Left            =   12000
      List            =   "主界面加查找界面.frx":005D
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "查找书籍"
      Height          =   495
      Left            =   14880
      TabIndex        =   5
      ToolTipText     =   "除了入库时间和库存外其他可以使用模糊匹配的方式查找"
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
      Left            =   12000
      TabIndex        =   4
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "修改数据"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      ToolTipText     =   "点击修改后修改对应单元格的数据然后点击回车来完成修改"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "新建借书"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      ToolTipText     =   "新建借书"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "删除此书"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      ToolTipText     =   "删除此书"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "降序Z-A"
      Height          =   255
      Left            =   10200
      TabIndex        =   10
      Top             =   720
      Width           =   1215
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
      Left            =   5760
      TabIndex        =   6
      Top             =   720
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
      Begin VB.Menu 修改数据 
         Caption         =   "修改数据"
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
   End
   Begin VB.Menu 历史记录查询 
      Caption         =   "历史记录查询"
      Begin VB.Menu 书籍变更记录查询 
         Caption         =   "书籍变更记录查询"
      End
      Begin VB.Menu 借书或还书记录查询 
         Caption         =   "借书或还书记录查询"
      End
   End
End
Attribute VB_Name = "主界面加查找界面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public conn As New ADODb.Connection
Private first As Boolean
Private last_strsql As String
Private last_sort As String

Public Sub maindata_load(strsql As String, is_sort As Boolean)
    last_strsql = strsql
    If is_sort = True Then
            If Label1.Caption = "降序Z-A" Then
                strsql = strsql & " desc"
            ElseIf Label1.Caption = "升序A-Z" Then
                strsql = strsql & " asc"
            End If
    Else
    End If
    
            'MsgBox strsql
    Dim nothiscolumn As Boolean
    Dim rs As New ADODb.Recordset
    rs.Open strsql, conn, 1, 3
    '先检查有没有这个列名
    For Each i In rs.Fields
        If Trim(Combo2.Text) = i.Name Then nothiscolumn = False:            'MsgBox strsql
    Next i
    If nothiscolumn = False Then
        If is_sort Then
            If rs.EOF Then MsgBox "没有与之匹配的数据", 48, "错误": Exit Sub
            '数据后读取后将他与datagrid数据表同步
            Set DataGrid1.DataSource = rs
            
            
                DataGrid1.ScrollBars = dbgVertical
                DataGrid1.Width = 主界面加查找界面.Width - DataGrid1.Left * 2
                DataGrid1.Height = 主界面加查找界面.Height - DataGrid1.Top * 1.6
                'DataGrid1.DefColWidth = DataGrid1.Width / 6
                            '同步完设置datagrid的列宽
                            For i = 0 To DataGrid1.Columns.Count - 1
                                DataGrid1.Columns(i).Width = (DataGrid1.Width / DataGrid1.Columns.Count) - 100
                            Next i
                            
        End If
    Else
        MsgBox "数据库中没有这个列名，请检查后输入", 48, "注意"
    End If
End Sub



Private Sub Combo1_Click()
If Combo1.Text = Combo1.List(5) Then Text1 = Year(Date) & "/" & Month(Date) & "/"
End Sub

Private Sub Combo2_Click()
If first Then
    first = False
Else
    last_strsql = Mid(last_strsql, 1, Len(last_strsql) - Len(last_sort))
                maindata_load last_strsql & Trim(Combo2.Text), True
        last_sort = Trim(Combo2.Text)

End If
End Sub

Private Sub Command1_Click()
temp_booknumber = DataGrid1.Columns("书号").CellText(DataGrid1.Bookmark)
temp_bookname = DataGrid1.Columns("书名").CellText(DataGrid1.Bookmark)
temp_vocation = DataGrid1.Columns("专业").CellText(DataGrid1.Bookmark)

a = MsgBox("你确定要删除吗？", vbYesNo, "警告")
If a = vbYes Then
    '表格先删除（表格和数据库只是读取关系，暂时不知修改的办法）
    DataGrid1.Columns.Remove ColIndex
    '再删后台的数据
    maindata_load " delete from 书库系统 where 书号 =  " & "'" & Trim(temp_booknumber) & "'", False
    
    '删完后记得记录下来，以后算账
    Dim rs As New ADODb.Recordset
    rs.Open "书籍信息变更记录表", conn, 1, 3
    rs.AddNew
    rs.Fields("书号") = temp_booknumber
    rs.Fields("书名") = temp_bookname
    rs.Fields("专业") = temp_vocation
    rs.Fields("更改时间") = Date
    rs.Fields("更改的类型") = "删除"
    rs.Fields("变更的数据") = "全部"
    rs.Fields("更改前") = "有"
    rs.Fields("更改后") = "无"
    rs.Update
    rs.Close
    
     '到这来没问题就是真的没问题了呢，因为数据有变更所以需要再次重新加载
    MsgBox "已删除"
    maindata_load " select * from  书库系统 order by " & Trim(主界面加查找界面.Combo2.Text), True
Else
    Exit Sub
End If
End Sub

Private Sub Command2_Click()
借书界面.Show
借书界面.Text1(0) = DataGrid1.Columns("书号").CellText(DataGrid1.Bookmark)
借书界面.Text1(1) = DataGrid1.Columns("书名").CellText(DataGrid1.Bookmark)
借书界面.Text1(3) = Date
主界面加查找界面.Enabled = False
End Sub

Private Sub Command3_Click()
If DataGrid1.Row = 0 And DataGrid1.Col = 0 Then a = MsgBox("是否要修改第一行第一列的这个数据？", vbYesNo, "提示")
If a = vbNo Then Exit Sub

修改数据窗口.Text1 = DataGrid1.Columns("书号").CellText(DataGrid1.Bookmark)
修改数据窗口.Text2 = DataGrid1.Columns("书名").CellText(DataGrid1.Bookmark)
修改数据窗口.Text3 = DataGrid1.Columns("专业").CellText(DataGrid1.Bookmark)
修改数据窗口.Combo1.Text = 修改数据窗口.Combo1.List(DataGrid1.Col + 1)
修改数据窗口.Text5 = DataGrid1.Columns(DataGrid1.Col).CellText(DataGrid1.Bookmark)

修改数据窗口.Text4 = Date

修改数据窗口.Show
主界面加查找界面.Enabled = False
修改数据窗口.Text6.SetFocus
End Sub

Private Sub Command4_Click()
If Combo1.Text = "请选择要查找的字段" Then MsgBox "请先输入要查找的字段": Exit Sub
If Combo1.Text = "入库时间" And IsDate(Trim(Text1.Text)) = False Then MsgBox "请输入正确的日期格式。如：2022/3/3", 48, "注意": Exit Sub
If Combo1.Text = "库存" And Val(Trim(Text1.Text)) <> Trim(Text1.Text) Then MsgBox "想查找库存请输入纯数字": Exit Sub
If Combo1.Text = "入库时间" Then
    maindata_load "select * from 书库系统 where 入库时间 like #" & Trim(Text1.Text) & "#" & " order by " & Trim(Combo2.Text), True
ElseIf Combo1.Text = "库存" Then
    maindata_load " select * from 书库系统 where 库存 =  " & Val(Trim(Text1.Text)) & " order by " & Trim(Combo2.Text), True
Else
    maindata_load " select * from 书库系统 where " & Trim(Combo1.Text) & " like '%" & Trim(Text1.Text) & "%'" & " order by " & Trim(Combo2.Text), True
End If

End Sub

Private Sub Command5_Click()
Print Replace(Replace(last_strsql, "%'", "#"), "'%", "#")
End Sub

Private Sub Command6_Click()
maindata_load " select * from  书库系统 order by " & Trim(Combo2.Text), True
End Sub




Private Sub Command7_Click()
Unload Me
End Sub

Private Sub DataGrid1_DblClick()
修改数据窗口.Text1 = DataGrid1.Columns("书号").CellText(DataGrid1.Bookmark)
修改数据窗口.Text2 = DataGrid1.Columns("书名").CellText(DataGrid1.Bookmark)
修改数据窗口.Text3 = DataGrid1.Columns("专业").CellText(DataGrid1.Bookmark)
修改数据窗口.Combo1.Text = 修改数据窗口.Combo1.List(DataGrid1.Col + 1)
修改数据窗口.Text5 = DataGrid1.Columns(DataGrid1.Col).CellText(DataGrid1.Bookmark)

修改数据窗口.Text4 = Date

修改数据窗口.Show
主界面加查找界面.Enabled = False
修改数据窗口.Text6.SetFocus
End Sub

'Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
'
'If KeyAscii = 13 Then
'    yes = MsgBox("你确定要更改数据吗？", vbYesNo, "注意")
'    If yes = vbYes Then
'        '先看看如果更改后的数据没有变化就不管他，而且如果连续两次点击行的话他会报错显示下标错误，这个可以缓解问题
'        If temp_column = DataGrid1.Columns(DataGrid1.Col).CellText(DataGrid1.Bookmark) Then MsgBox 132: Exit Sub
'
'        DataGrid1.Refresh
'        MsgBox "updata 书库系统 set  " & DataGrid1.Columns(DataGrid1.Col).DataField & " = '" & DataGrid1.Columns(DataGrid1.Col).CellText(DataGrid1.Bookmark) & "' where 书号 = '" & DataGrid1.Columns("书号").CellText(DataGrid1.Bookmark) & "'"
'        'DataGrid1.Columns(DataGrid1.Col).DataField     返回的是当前选择单元格的列名
'        'temp_column是修改前的数据      DataGrid1.Columns(DataGrid1.Col).CellText(DataGrid1.Bookmark)是修改后的数据
'        '其中where还是采用书号校准（书号是绝对唯一的，这个重复就全家完蛋）
'        'maindata_load "updata 书库系统 set  " & DataGrid1.Columns(DataGrid1.Col).DataField & " = '" & DataGrid1.Columns(DataGrid1.Col).CellText(DataGrid1.Bookmark) & "' where 书号 = '" & DataGrid1.Columns("书号").CellText(DataGrid1.Bookmark) & "'"0
'
'        '记得加上变更记录信息在变更记录表
'
'
'        '只有点了回车后才不会触发
'        'DataGrid1.DataChanged = False
'
'    End If
'Else
'    Exit Sub
'End If
'End Sub
'
'
'
'Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
''如果连续两次点击行的话会报错，所以要把报错取消掉
''On Error Resume Next
'
''Print DataGrid1.DataChanged
''
''If DataGrid1.DataChanged = True Then
''    MsgBox "修改数据时请不要点击鼠标，修改完后按回车确认更改！", 48, "警告"
''    DataGrid1.Col = temp_col
''    DataGrid1.Row = temp_row
''    DataGrid1.Columns(temp_col).CellValue(temp_row + 1) = temp_column
''Else
''    temp_col = DataGrid1.Col
''    temp_row = DataGrid1.Row
''
''    '储存上次点击的单元格数据
''    temp_column = DataGrid1.Columns(DataGrid1.Col).CellText(DataGrid1.Bookmark)
''
''
'''    MsgBox temp_col
'''    MsgBox temp_row
'''    MsgBox temp_column
''End If
'
'temp_column = DataGrid1.Columns(DataGrid1.Col)
'MsgBox temp_column
'Print DataGrid1.DataChanged
'If DataGrid1.DataChanged = True Then
'    MsgBox "修改数据时请不要点击鼠标，修改完后按回车确认更改！", 48, "警告"
'    DataGrid1.Refresh
'End If
'End Sub

Private Sub Form_Activate()
DataGrid1.ScrollBars = dbgVertical
DataGrid1.Width = 主界面加查找界面.Width - DataGrid1.Left * 3
DataGrid1.Height = 主界面加查找界面.Height - DataGrid1.Top * 1.6
'DataGrid1.DefColWidth = DataGrid1.Width / 6
            '同步完设置datagrid的列宽
            For i = 0 To DataGrid1.Columns.Count - 1
                DataGrid1.Columns(i).Width = (DataGrid1.Width / DataGrid1.Columns.Count) - 100
            Next i
'用来搞什么什么光标的东西

'    Dim rs As New ADODB.Recordset
'    rs.Open " select * from  书库系统 order by " & Trim(Combo2.Text) & " desc ", conn, 1, 3
'    Set DataGrid1.DataSource = rs


    
End Sub

Private Sub Form_Load()
'先将窗口最大化
主界面加查找界面.WindowState = 2

first = True
Combo1.Text = Combo1.List(0)
Combo2.Text = "入库时间"
last_sort = "入库时间"

DataGrid1.AllowUpdate = False
DataGrid1.AllowDelete = True

'这里连接数据库
conn.Provider = "Microsoft.Jet.OLEDB.4.0"              '这行默认不变
conn.Open "..\value\书库.mdb"                             '这里选择数据库路径
conn.CursorLocation = adUseClient                       '设置什么光标（常规操作）

    maindata_load " select * from  书库系统 order by " & Trim(Combo2.Text), True
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
a = MsgBox("确定要退出吗？", vbYesNo, "QAQ")
If a = vbYes Then
    Cancel = 0
Else
    Cancel = 1
End If
End Sub

Private Sub Label1_Click()
If Label1.Caption = "降序Z-A" Then
    Label1.Caption = "升序A-Z"
        last_strsql = Mid(last_strsql, 1, Len(last_strsql) - Len(last_sort))
                maindata_load last_strsql & Trim(Combo2.Text), True
        last_sort = Trim(Combo2.Text)
ElseIf Label1.Caption = "升序A-Z" Then
    Label1.Caption = "降序Z-A"
        last_strsql = Mid(last_strsql, 1, Len(last_strsql) - Len(last_sort))
                maindata_load last_strsql & Trim(Combo2.Text), True
        last_sort = Trim(Combo2.Text)
End If
End Sub

Private Sub 查找书籍_Click()
详细查找界面.Show
主界面加查找界面.Enabled = False
End Sub

Private Sub 借书或还书记录查询_Click()
借书记录查看窗口.senthistory_load "select * from 借书记录表 order by " & 借书记录查看窗口.Combo1.Text, True
借书记录查看窗口.Show
主界面加查找界面.Enabled = False
End Sub

Private Sub 书籍变更记录查询_Click()

书籍信息变更记录查看窗口.Show
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

Private Sub 修改数据_Click()
修改数据窗口.Show
主界面加查找界面.Enabled = False
End Sub
