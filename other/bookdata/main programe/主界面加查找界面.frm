VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ������Ӳ��ҽ��� 
   Caption         =   "������"
   ClientHeight    =   10680
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17400
   BeginProperty Font 
      Name            =   "����"
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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command7 
      Caption         =   "�˳�"
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
   Begin VB.ComboBox Combo2 
      Height          =   360
      ItemData        =   "������Ӳ��ҽ���.frx":0000
      Left            =   8280
      List            =   "������Ӳ��ҽ���.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "�˳�����״̬"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "������Ӳ��ҽ���.frx":0044
      Left            =   12000
      List            =   "������Ӳ��ҽ���.frx":005D
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�����鼮"
      Height          =   495
      Left            =   14880
      TabIndex        =   5
      ToolTipText     =   "�������ʱ��Ϳ������������ʹ��ģ��ƥ��ķ�ʽ����"
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�޸�����"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      ToolTipText     =   "����޸ĺ��޸Ķ�Ӧ��Ԫ�������Ȼ�����س�������޸�"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�½�����"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      ToolTipText     =   "�½�����"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ɾ������"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      ToolTipText     =   "ɾ������"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "����Z-A"
      Height          =   255
      Left            =   10200
      TabIndex        =   10
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "ѡ�����ĸ��ֶ�����Ĭ�������ʱ��"
      BeginProperty Font 
         Name            =   "����"
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
   Begin VB.Menu �鼮���� 
      Caption         =   "�鼮����"
      Begin VB.Menu ����鼮 
         Caption         =   "����鼮"
      End
      Begin VB.Menu ɾ���鼮 
         Caption         =   "ɾ���鼮"
      End
      Begin VB.Menu �޸����� 
         Caption         =   "�޸�����"
      End
      Begin VB.Menu �����鼮 
         Caption         =   "�����鼮����ϸ��"
      End
   End
   Begin VB.Menu ����ϵͳ 
      Caption         =   "����ϵͳ"
      Begin VB.Menu ���������¼ 
         Caption         =   "���������¼"
      End
   End
   Begin VB.Menu ��ʷ��¼��ѯ 
      Caption         =   "��ʷ��¼��ѯ"
      Begin VB.Menu �鼮�����¼��ѯ 
         Caption         =   "�鼮�����¼��ѯ"
      End
      Begin VB.Menu ��������¼��ѯ 
         Caption         =   "��������¼��ѯ"
      End
   End
End
Attribute VB_Name = "������Ӳ��ҽ���"
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
            If Label1.Caption = "����Z-A" Then
                strsql = strsql & " desc"
            ElseIf Label1.Caption = "����A-Z" Then
                strsql = strsql & " asc"
            End If
    Else
    End If
    
            'MsgBox strsql
    Dim nothiscolumn As Boolean
    Dim rs As New ADODb.Recordset
    rs.Open strsql, conn, 1, 3
    '�ȼ����û���������
    For Each i In rs.Fields
        If Trim(Combo2.Text) = i.Name Then nothiscolumn = False:            'MsgBox strsql
    Next i
    If nothiscolumn = False Then
        If is_sort Then
            If rs.EOF Then MsgBox "û����֮ƥ�������", 48, "����": Exit Sub
            '���ݺ��ȡ������datagrid���ݱ�ͬ��
            Set DataGrid1.DataSource = rs
            
            
                DataGrid1.ScrollBars = dbgVertical
                DataGrid1.Width = ������Ӳ��ҽ���.Width - DataGrid1.Left * 2
                DataGrid1.Height = ������Ӳ��ҽ���.Height - DataGrid1.Top * 1.6
                'DataGrid1.DefColWidth = DataGrid1.Width / 6
                            'ͬ��������datagrid���п�
                            For i = 0 To DataGrid1.Columns.Count - 1
                                DataGrid1.Columns(i).Width = (DataGrid1.Width / DataGrid1.Columns.Count) - 100
                            Next i
                            
        End If
    Else
        MsgBox "���ݿ���û��������������������", 48, "ע��"
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
temp_booknumber = DataGrid1.Columns("���").CellText(DataGrid1.Bookmark)
temp_bookname = DataGrid1.Columns("����").CellText(DataGrid1.Bookmark)
temp_vocation = DataGrid1.Columns("רҵ").CellText(DataGrid1.Bookmark)

a = MsgBox("��ȷ��Ҫɾ����", vbYesNo, "����")
If a = vbYes Then
    '�����ɾ�����������ݿ�ֻ�Ƕ�ȡ��ϵ����ʱ��֪�޸ĵİ취��
    DataGrid1.Columns.Remove ColIndex
    '��ɾ��̨������
    maindata_load " delete from ���ϵͳ where ��� =  " & "'" & Trim(temp_booknumber) & "'", False
    
    'ɾ���ǵü�¼�������Ժ�����
    Dim rs As New ADODb.Recordset
    rs.Open "�鼮��Ϣ�����¼��", conn, 1, 3
    rs.AddNew
    rs.Fields("���") = temp_booknumber
    rs.Fields("����") = temp_bookname
    rs.Fields("רҵ") = temp_vocation
    rs.Fields("����ʱ��") = Date
    rs.Fields("���ĵ�����") = "ɾ��"
    rs.Fields("���������") = "ȫ��"
    rs.Fields("����ǰ") = "��"
    rs.Fields("���ĺ�") = "��"
    rs.Update
    rs.Close
    
     '������û����������û�������أ���Ϊ�����б��������Ҫ�ٴ����¼���
    MsgBox "��ɾ��"
    maindata_load " select * from  ���ϵͳ order by " & Trim(������Ӳ��ҽ���.Combo2.Text), True
Else
    Exit Sub
End If
End Sub

Private Sub Command2_Click()
�������.Show
�������.Text1(0) = DataGrid1.Columns("���").CellText(DataGrid1.Bookmark)
�������.Text1(1) = DataGrid1.Columns("����").CellText(DataGrid1.Bookmark)
�������.Text1(3) = Date
������Ӳ��ҽ���.Enabled = False
End Sub

Private Sub Command3_Click()
If DataGrid1.Row = 0 And DataGrid1.Col = 0 Then a = MsgBox("�Ƿ�Ҫ�޸ĵ�һ�е�һ�е�������ݣ�", vbYesNo, "��ʾ")
If a = vbNo Then Exit Sub

�޸����ݴ���.Text1 = DataGrid1.Columns("���").CellText(DataGrid1.Bookmark)
�޸����ݴ���.Text2 = DataGrid1.Columns("����").CellText(DataGrid1.Bookmark)
�޸����ݴ���.Text3 = DataGrid1.Columns("רҵ").CellText(DataGrid1.Bookmark)
�޸����ݴ���.Combo1.Text = �޸����ݴ���.Combo1.List(DataGrid1.Col + 1)
�޸����ݴ���.Text5 = DataGrid1.Columns(DataGrid1.Col).CellText(DataGrid1.Bookmark)

�޸����ݴ���.Text4 = Date

�޸����ݴ���.Show
������Ӳ��ҽ���.Enabled = False
�޸����ݴ���.Text6.SetFocus
End Sub

Private Sub Command4_Click()
If Combo1.Text = "��ѡ��Ҫ���ҵ��ֶ�" Then MsgBox "��������Ҫ���ҵ��ֶ�": Exit Sub
If Combo1.Text = "���ʱ��" And IsDate(Trim(Text1.Text)) = False Then MsgBox "��������ȷ�����ڸ�ʽ���磺2022/3/3", 48, "ע��": Exit Sub
If Combo1.Text = "���" And Val(Trim(Text1.Text)) <> Trim(Text1.Text) Then MsgBox "����ҿ�������봿����": Exit Sub
If Combo1.Text = "���ʱ��" Then
    maindata_load "select * from ���ϵͳ where ���ʱ�� like #" & Trim(Text1.Text) & "#" & " order by " & Trim(Combo2.Text), True
ElseIf Combo1.Text = "���" Then
    maindata_load " select * from ���ϵͳ where ��� =  " & Val(Trim(Text1.Text)) & " order by " & Trim(Combo2.Text), True
Else
    maindata_load " select * from ���ϵͳ where " & Trim(Combo1.Text) & " like '%" & Trim(Text1.Text) & "%'" & " order by " & Trim(Combo2.Text), True
End If

End Sub

Private Sub Command5_Click()
Print Replace(Replace(last_strsql, "%'", "#"), "'%", "#")
End Sub

Private Sub Command6_Click()
maindata_load " select * from  ���ϵͳ order by " & Trim(Combo2.Text), True
End Sub




Private Sub Command7_Click()
Unload Me
End Sub

Private Sub DataGrid1_DblClick()
�޸����ݴ���.Text1 = DataGrid1.Columns("���").CellText(DataGrid1.Bookmark)
�޸����ݴ���.Text2 = DataGrid1.Columns("����").CellText(DataGrid1.Bookmark)
�޸����ݴ���.Text3 = DataGrid1.Columns("רҵ").CellText(DataGrid1.Bookmark)
�޸����ݴ���.Combo1.Text = �޸����ݴ���.Combo1.List(DataGrid1.Col + 1)
�޸����ݴ���.Text5 = DataGrid1.Columns(DataGrid1.Col).CellText(DataGrid1.Bookmark)

�޸����ݴ���.Text4 = Date

�޸����ݴ���.Show
������Ӳ��ҽ���.Enabled = False
�޸����ݴ���.Text6.SetFocus
End Sub

'Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
'
'If KeyAscii = 13 Then
'    yes = MsgBox("��ȷ��Ҫ����������", vbYesNo, "ע��")
'    If yes = vbYes Then
'        '�ȿ���������ĺ������û�б仯�Ͳ���������������������ε���еĻ����ᱨ����ʾ�±����������Ի�������
'        If temp_column = DataGrid1.Columns(DataGrid1.Col).CellText(DataGrid1.Bookmark) Then MsgBox 132: Exit Sub
'
'        DataGrid1.Refresh
'        MsgBox "updata ���ϵͳ set  " & DataGrid1.Columns(DataGrid1.Col).DataField & " = '" & DataGrid1.Columns(DataGrid1.Col).CellText(DataGrid1.Bookmark) & "' where ��� = '" & DataGrid1.Columns("���").CellText(DataGrid1.Bookmark) & "'"
'        'DataGrid1.Columns(DataGrid1.Col).DataField     ���ص��ǵ�ǰѡ��Ԫ�������
'        'temp_column���޸�ǰ������      DataGrid1.Columns(DataGrid1.Col).CellText(DataGrid1.Bookmark)���޸ĺ������
'        '����where���ǲ������У׼������Ǿ���Ψһ�ģ�����ظ���ȫ���군��
'        'maindata_load "updata ���ϵͳ set  " & DataGrid1.Columns(DataGrid1.Col).DataField & " = '" & DataGrid1.Columns(DataGrid1.Col).CellText(DataGrid1.Bookmark) & "' where ��� = '" & DataGrid1.Columns("���").CellText(DataGrid1.Bookmark) & "'"0
'
'        '�ǵü��ϱ����¼��Ϣ�ڱ����¼��
'
'
'        'ֻ�е��˻س���Ų��ᴥ��
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
''����������ε���еĻ��ᱨ������Ҫ�ѱ���ȡ����
''On Error Resume Next
'
''Print DataGrid1.DataChanged
''
''If DataGrid1.DataChanged = True Then
''    MsgBox "�޸�����ʱ�벻Ҫ�����꣬�޸���󰴻س�ȷ�ϸ��ģ�", 48, "����"
''    DataGrid1.Col = temp_col
''    DataGrid1.Row = temp_row
''    DataGrid1.Columns(temp_col).CellValue(temp_row + 1) = temp_column
''Else
''    temp_col = DataGrid1.Col
''    temp_row = DataGrid1.Row
''
''    '�����ϴε���ĵ�Ԫ������
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
'    MsgBox "�޸�����ʱ�벻Ҫ�����꣬�޸���󰴻س�ȷ�ϸ��ģ�", 48, "����"
'    DataGrid1.Refresh
'End If
'End Sub

Private Sub Form_Activate()
DataGrid1.ScrollBars = dbgVertical
DataGrid1.Width = ������Ӳ��ҽ���.Width - DataGrid1.Left * 3
DataGrid1.Height = ������Ӳ��ҽ���.Height - DataGrid1.Top * 1.6
'DataGrid1.DefColWidth = DataGrid1.Width / 6
            'ͬ��������datagrid���п�
            For i = 0 To DataGrid1.Columns.Count - 1
                DataGrid1.Columns(i).Width = (DataGrid1.Width / DataGrid1.Columns.Count) - 100
            Next i
'������ʲôʲô���Ķ���

'    Dim rs As New ADODB.Recordset
'    rs.Open " select * from  ���ϵͳ order by " & Trim(Combo2.Text) & " desc ", conn, 1, 3
'    Set DataGrid1.DataSource = rs


    
End Sub

Private Sub Form_Load()
'�Ƚ��������
������Ӳ��ҽ���.WindowState = 2

first = True
Combo1.Text = Combo1.List(0)
Combo2.Text = "���ʱ��"
last_sort = "���ʱ��"

DataGrid1.AllowUpdate = False
DataGrid1.AllowDelete = True

'�����������ݿ�
conn.Provider = "Microsoft.Jet.OLEDB.4.0"              '����Ĭ�ϲ���
conn.Open "..\value\���.mdb"                             '����ѡ�����ݿ�·��
conn.CursorLocation = adUseClient                       '����ʲô��꣨���������

    maindata_load " select * from  ���ϵͳ order by " & Trim(Combo2.Text), True
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
a = MsgBox("ȷ��Ҫ�˳���", vbYesNo, "QAQ")
If a = vbYes Then
    Cancel = 0
Else
    Cancel = 1
End If
End Sub

Private Sub Label1_Click()
If Label1.Caption = "����Z-A" Then
    Label1.Caption = "����A-Z"
        last_strsql = Mid(last_strsql, 1, Len(last_strsql) - Len(last_sort))
                maindata_load last_strsql & Trim(Combo2.Text), True
        last_sort = Trim(Combo2.Text)
ElseIf Label1.Caption = "����A-Z" Then
    Label1.Caption = "����Z-A"
        last_strsql = Mid(last_strsql, 1, Len(last_strsql) - Len(last_sort))
                maindata_load last_strsql & Trim(Combo2.Text), True
        last_sort = Trim(Combo2.Text)
End If
End Sub

Private Sub �����鼮_Click()
��ϸ���ҽ���.Show
������Ӳ��ҽ���.Enabled = False
End Sub

Private Sub ��������¼��ѯ_Click()
�����¼�鿴����.senthistory_load "select * from �����¼�� order by " & �����¼�鿴����.Combo1.Text, True
�����¼�鿴����.Show
������Ӳ��ҽ���.Enabled = False
End Sub

Private Sub �鼮�����¼��ѯ_Click()

�鼮��Ϣ�����¼�鿴����.Show
������Ӳ��ҽ���.Enabled = False
End Sub

Private Sub ����鼮_Click()
�����鼮����.Show
������Ӳ��ҽ���.Enabled = False
End Sub

Private Sub ���������¼_Click()
�������.Show
������Ӳ��ҽ���.Enabled = False
End Sub

Private Sub �޸�����_Click()
�޸����ݴ���.Show
������Ӳ��ҽ���.Enabled = False
End Sub
