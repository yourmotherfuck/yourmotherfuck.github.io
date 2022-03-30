VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form 借书记录查看窗口 
   Caption         =   "借书记录查看窗口"
   ClientHeight    =   9375
   ClientLeft      =   1125
   ClientTop       =   870
   ClientWidth     =   18720
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   18720
   Begin VB.CommandButton Command5 
      Caption         =   "删除选中数据"
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
      Left            =   13800
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   615
      Left            =   16560
      TabIndex        =   9
      Top             =   360
      Width           =   1815
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
      TabIndex        =   5
      Top             =   960
      Width           =   2415
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
      ItemData        =   "借书记录查看窗口.frx":0000
      Left            =   3240
      List            =   "借书记录查看窗口.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Width           =   2295
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
      ItemData        =   "借书记录查看窗口.frx":0047
      Left            =   7800
      List            =   "借书记录查看窗口.frx":005D
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   2655
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
      TabIndex        =   2
      Top             =   360
      Width           =   1695
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
      Height          =   615
      Left            =   10800
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
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
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6855
      Left            =   360
      TabIndex        =   6
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
      TabIndex        =   8
      Top             =   600
      Width           =   2415
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
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.Menu 详细查找界面 
      Caption         =   "详细查找界面"
   End
End
Attribute VB_Name = "借书记录查看窗口"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public first As Boolean
Public last_sort, last_strsql As String

Public Sub senthistory_load(strsql As String, is_sort As Boolean)
last_strsql = strsql
    If is_sort = True Then
            If Label2.Caption = "降序Z-A" Then
                strsql = strsql & " desc"
            ElseIf Label2.Caption = "升序A-Z" Then
                strsql = strsql & " asc"
            End If
    End If
        'MsgBox last_strsql & Chr(13) & strsql
    Dim rs As New ADODb.Recordset
    rs.Open strsql, 主界面加查找界面.conn, 1, 3
    
    If is_sort Then
        If rs.EOF Then MsgBox "没有与之匹配的数据", 48, "错误": Exit Sub
    End If
    
    Set DataGrid1.DataSource = rs
    
End Sub

Private Sub Combo1_Click()
If first = True Then
    'MsgBox last_strsql
        last_strsql = Mid(last_strsql, 1, Len(last_strsql) - Len(last_sort) - 1)
    'MsgBox last_strsql
                senthistory_load last_strsql & " " & Trim(Combo1.Text), True
            last_sort = Trim(Combo1.Text)
End If
End Sub

Private Sub Command1_Click()
If Combo2.Text = "请选择要查找的字段" Then MsgBox "请选择要查找的列名": Exit Sub
If Combo2.Text = "借书时间" Then
    If IsDate(Trim(Text1)) = False Then MsgBox "请输入正确的时间格式！eg:2022/3/3", 48, "警告": Exit Sub
    senthistory_load "select * from 借书记录表 where " & Combo2.Text & " like #" & Text1 & "# order by " & Combo1.Text, True
Else
    senthistory_load "select * from 借书记录表 where " & Combo2.Text & " like '%" & Text1 & "%' order by " & Combo1.Text & " ", True
End If

End Sub

Private Sub Command2_Click()
借书记录查看窗口.senthistory_load "select * from 借书记录表 order by " & 借书记录查看窗口.Combo1.Text, True
End Sub

Private Sub Command3_Click()
借书记录表修改数据界面.Show
借书记录查看窗口.Enabled = False
End Sub

Private Sub Command4_Click()
Unload Me
主界面加查找界面.Enabled = True
主界面加查找界面.Show
End Sub

Private Sub Command5_Click()
a = MsgBox("您确定要删除吗，删除后数据无法找回切不会记录到任何表中！", vbYesNo, "警告")
If a = vbNo Then Exit Sub

    
'MsgBox " delete from 借书记录表 where 书号 = '" & DataGrid1.Columns(0).CellText(DataGrid1.Bookmark) & "' " & "and 书名 = '" & DataGrid1.Columns("书名").CellText(DataGrid1.Bookmark) & "' " & "and 借书人（单位） = '" & DataGrid1.Columns("借书人（单位）").CellText(DataGrid1.Bookmark) & "' " & "and 借书时间 = '" & DataGrid1.Columns("借书时间").CellText(DataGrid1.Bookmark) & "' " & "and 借书数量 = '" & DataGrid1.Columns("借书数量").CellText(DataGrid1.Bookmark) & "' "
senthistory_load " delete from 借书记录表 where 书号 = '" & DataGrid1.Columns("书号").CellText(DataGrid1.Bookmark) & "' " & "and 书名 = '" & DataGrid1.Columns("书名").CellText(DataGrid1.Bookmark) & "' " & "and 借书人（单位） = '" & DataGrid1.Columns("借书人（单位）").CellText(DataGrid1.Bookmark) & "' " & "and 借书时间 = #" & DataGrid1.Columns("借书时间").CellText(DataGrid1.Bookmark) & "# " & "and 借书数量 = " & DataGrid1.Columns("借书数量").CellText(DataGrid1.Bookmark) & " ", False

    '把表内的先删掉
    DataGrid1.Columns.Remove ColIndex
    
'删完后重新加载一次
senthistory_load "select * from 借书记录表 order by " & Combo1.Text, True
End Sub

Private Sub DataGrid1_DblClick()
借书记录表修改数据界面.Show
借书记录表修改数据界面.Text1 = DataGrid1.Columns(借书记录表修改数据界面.Label1).CellText(DataGrid1.Bookmark)
借书记录表修改数据界面.Text2 = DataGrid1.Columns(借书记录表修改数据界面.Label2).CellText(DataGrid1.Bookmark)
借书记录表修改数据界面.Text3 = DataGrid1.Columns(借书记录表修改数据界面.Label3).CellText(DataGrid1.Bookmark)
借书记录表修改数据界面.Text4 = DataGrid1.Columns(借书记录表修改数据界面.Label4).CellText(DataGrid1.Bookmark)
借书记录表修改数据界面.Text5 = DataGrid1.Columns(借书记录表修改数据界面.Label5).CellText(DataGrid1.Bookmark)
借书记录查看窗口.Enabled = False
借书记录表修改数据界面.Combo2.Text = DataGrid1.Columns(DataGrid1.Col).DataField
借书记录表修改数据界面.Text10.SetFocus
End Sub

Private Sub Form_Activate()
DataGrid1.Width = 借书记录查看窗口.Width - DataGrid1.Left * 3
DataGrid1.Height = 借书记录查看窗口.Height - DataGrid1.Top * 1.7
DataGrid1.DefColWidth = DataGrid1.Width / Combo1.ListCount
End Sub

Private Sub Form_Load()
Combo1.Text = Combo1.List(3)
Combo2.Text = Combo2.List(0)
借书记录查看窗口.Left = 2000
senthistory_load "select * from 借书记录表 order by " & Combo1.Text, True
first = True
last_sort = Combo1.Text
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
'错误示范了属实是
'        last_strsql = Mid(last_strsql, 1, Len(last_strsql) - Len(last_sort))
'                senthistory_load last_strsql & Trim(Combo1.Text), True
'            last_sort = Trim(Combo1.Text)
    'MsgBox last_strsql
        last_strsql = Mid(last_strsql, 1, Len(last_strsql) - Len(last_sort) - 1)
    'MsgBox last_strsql
                senthistory_load last_strsql & " " & Trim(Combo1.Text), True
            last_sort = Trim(Combo1.Text)
End Sub


Private Sub 详细查找界面_Click()
借书记录表详细查找.Show
借书记录查看窗口.Enabled = False
End Sub
