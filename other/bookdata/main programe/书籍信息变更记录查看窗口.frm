VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form �鼮��Ϣ�����¼�鿴���� 
   Caption         =   "�鼮��Ϣ�����¼�鿴����"
   ClientHeight    =   9900
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15645
   LinkTopic       =   "Form1"
   ScaleHeight     =   9900
   ScaleWidth      =   15645
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command5 
      Caption         =   "ɾ������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�޸����ݣ���������ʷ��¼�ģ�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�˳�����ģʽ"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ʼ����"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "�鼮��Ϣ�����¼�鿴����.frx":0000
      Left            =   7800
      List            =   "�鼮��Ϣ�����¼�鿴����.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "�鼮��Ϣ�����¼�鿴����.frx":007B
      Left            =   3240
      List            =   "�鼮��Ϣ�����¼�鿴����.frx":0097
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Caption         =   "����Z-A"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ѡ�����ĸ��ֶ�����Ĭ�����ʱ��"
      BeginProperty Font 
         Name            =   "����"
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
   Begin VB.Menu ��ϸ���ҽ��� 
      Caption         =   "��ϸ���Ҿͽ���"
   End
End
Attribute VB_Name = "�鼮��Ϣ�����¼�鿴����"
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
            If Label2.Caption = "����Z-A" Then
                strsql = strsql & " desc"
            ElseIf Label2.Caption = "����A-Z" Then
                strsql = strsql & " asc"
            End If
    End If
        'MsgBox strsql
    Dim rs As New ADODb.Recordset
    rs.Open strsql, ������Ӳ��ҽ���.conn, 1, 3
    
    If is_sort Then
        If rs.EOF Then MsgBox "û����֮ƥ�������", 48, "����": Exit Sub
    End If
    
    Set DataGrid1.DataSource = rs
        DataGrid1.ScrollBars = dbgVertical
        DataGrid1.Width = �鼮��Ϣ�����¼�鿴����.Width - DataGrid1.Left * 2
        DataGrid1.Height = �鼮��Ϣ�����¼�鿴����.Height - DataGrid1.Top * 1.7
        
            'ͬ��������datagrid���п�
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
If Combo2.Text = "��ѡ��Ҫ���ҵ��ֶ�" Then MsgBox "��ѡ��Ҫ���ҵ�����": Exit Sub

If Combo2.Text = "����ʱ��" Then
    If IsDate(Trim(Text1)) = False Then MsgBox "��������ȷ��ʱ���ʽ��eg:2022/3/3", 48, "����": Exit Sub
    bookhistory_load "select * from �鼮��Ϣ�����¼�� where " & Combo2.Text & " like #" & Text1 & "# order by " & Combo1.Text, True
Else
    bookhistory_load "select * from �鼮��Ϣ�����¼�� where " & Combo2.Text & " like '%" & Text1 & "%' order by " & Combo1.Text & " ", True
End If

End Sub

Private Sub Command2_Click()
�鼮��Ϣ�����¼�鿴����.bookhistory_load "select * from �鼮��Ϣ�����¼�� order by " & �鼮��Ϣ�����¼�鿴����.Combo1.Text, True
Text1 = ""
End Sub

Private Sub Command3_Click()
�鼮��¼���޸����ݽ���.Show
�鼮��Ϣ�����¼�鿴����.Enabled = False
End Sub

Private Sub Command4_Click()
Unload Me
������Ӳ��ҽ���.Enabled = True
������Ӳ��ҽ���.Show
End Sub



Private Sub Command5_Click()
a = MsgBox("��ȷ��Ҫɾ����ɾ���������޷��һ��в����¼���κα��У�", vbYesNo, "����")
If a = vbNo Then Exit Sub

    'MsgBox " delete from �鼮��Ϣ�����¼�� where ��� = '" & DataGrid1.Columns(0).CellText(DataGrid1.Bookmark) & "' " & " and ���� = '" & DataGrid1.Columns("����").CellText(DataGrid1.Bookmark) & "' " & " and רҵ = '" & DataGrid1.Columns("רҵ").CellText(DataGrid1.Bookmark) & "' " & " and ����ʱ�� = #" & DataGrid1.Columns("����ʱ��").CellText(DataGrid1.Bookmark) & "# " & " and ���ĵ����� = '" & DataGrid1.Columns("���ĵ�����").CellText(DataGrid1.Bookmark) & "' " & " and ��������� = '" & DataGrid1.Columns("���������").CellText(DataGrid1.Bookmark) & "' " & " and ����ǰ = '" & DataGrid1.Columns("����ǰ").CellText(DataGrid1.Bookmark) & "' " & " and ���ĺ� = '" & DataGrid1.Columns("���ĺ�").CellText(DataGrid1.Bookmark) & "' "
    bookhistory_load " delete from �鼮��Ϣ�����¼�� where ��� = '" & DataGrid1.Columns("���").CellText(DataGrid1.Bookmark) & "' " & " and ���� = '" & DataGrid1.Columns("����").CellText(DataGrid1.Bookmark) & "' " & " and רҵ = '" & DataGrid1.Columns("רҵ").CellText(DataGrid1.Bookmark) & "' " & " and ����ʱ�� = #" & DataGrid1.Columns("����ʱ��").CellText(DataGrid1.Bookmark) & "# " & " and ���ĵ����� = '" & DataGrid1.Columns("���ĵ�����").CellText(DataGrid1.Bookmark) & "' " & " and ��������� = '" & DataGrid1.Columns("���������").CellText(DataGrid1.Bookmark) & "' " & " and ����ǰ = '" & DataGrid1.Columns("����ǰ").CellText(DataGrid1.Bookmark) & "' " & " and ���ĺ� = '" & DataGrid1.Columns("���ĺ�").CellText(DataGrid1.Bookmark) & "' ", False
    MsgBox "ɾ���ɹ�"
    
    'ɾ�������¼���
    bookhistory_load "select * from �鼮��Ϣ�����¼�� order by " & �鼮��Ϣ�����¼�鿴����.Combo1.Text, True
End Sub

Private Sub DataGrid1_DblClick()
�鼮��¼���޸����ݽ���.Text1 = DataGrid1.Columns(�鼮��¼���޸����ݽ���.Label1.Caption).CellText(DataGrid1.Bookmark)
�鼮��¼���޸����ݽ���.Text2 = DataGrid1.Columns(�鼮��¼���޸����ݽ���.Label2.Caption).CellText(DataGrid1.Bookmark)
�鼮��¼���޸����ݽ���.Text3 = DataGrid1.Columns(�鼮��¼���޸����ݽ���.Label3.Caption).CellText(DataGrid1.Bookmark)
�鼮��¼���޸����ݽ���.Text4 = DataGrid1.Columns(�鼮��¼���޸����ݽ���.Label4.Caption).CellText(DataGrid1.Bookmark)
�鼮��¼���޸����ݽ���.Text5 = DataGrid1.Columns(�鼮��¼���޸����ݽ���.Label5.Caption).CellText(DataGrid1.Bookmark)
�鼮��¼���޸����ݽ���.Text6 = DataGrid1.Columns(�鼮��¼���޸����ݽ���.Label6.Caption).CellText(DataGrid1.Bookmark)
�鼮��¼���޸����ݽ���.Text7 = DataGrid1.Columns(�鼮��¼���޸����ݽ���.Label7.Caption).CellText(DataGrid1.Bookmark)
�鼮��¼���޸����ݽ���.Text8 = DataGrid1.Columns(�鼮��¼���޸����ݽ���.Label8.Caption).CellText(DataGrid1.Bookmark)
�鼮��¼���޸����ݽ���.Show
�鼮��¼���޸����ݽ���.Combo2.Text = DataGrid1.Columns(DataGrid1.Col).DataField
�鼮��¼���޸����ݽ���.Text10.SetFocus
�鼮��Ϣ�����¼�鿴����.Enabled = False

End Sub

Private Sub Form_Activate()
DataGrid1.ScrollBars = dbgVertical
DataGrid1.Width = �鼮��Ϣ�����¼�鿴����.Width - DataGrid1.Left * 2
DataGrid1.Height = �鼮��Ϣ�����¼�鿴����.Height - DataGrid1.Top * 1.7

    'ͬ��������datagrid���п�
    For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).Width = (DataGrid1.Width / DataGrid1.Columns.Count) - 100
    Next i
End Sub

Private Sub Form_Load()
Combo1.Text = Combo1.List(3)
Combo2.Text = Combo2.List(0)
�鼮��Ϣ�����¼�鿴����.WindowState = 2
bookhistory_load "select * from �鼮��Ϣ�����¼�� order by " & �鼮��Ϣ�����¼�鿴����.Combo1.Text, True
�鼮��Ϣ�����¼�鿴����.first = True
�鼮��Ϣ�����¼�鿴����.last_sort = �鼮��Ϣ�����¼�鿴����.Combo1.Text

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
������Ӳ��ҽ���.Enabled = True
������Ӳ��ҽ���.Show
End Sub

Private Sub Label2_Click()
If Label2.Caption = "����Z-A" Then
    Label2.Caption = "����A-Z"
ElseIf Label2.Caption = "����A-Z" Then
    Label2.Caption = "����Z-A"
End If

'����һ�ε�sql����������������
    'MsgBox last_strsql
        last_strsql = Mid(last_strsql, 1, Len(last_strsql) - Len(last_sort) - 1)
    'MsgBox last_strsql
                bookhistory_load last_strsql & " " & Trim(Combo1.Text), True
            last_sort = Trim(Combo1.Text)
End Sub

Private Sub ��ϸ���ҽ���_Click()
�鼮��¼����ϸ���ҽ���.Show
�鼮��Ϣ�����¼�鿴����.Enabled = False
End Sub
