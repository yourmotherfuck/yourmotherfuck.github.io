VERSION 5.00
Begin VB.Form �����鼮���� 
   Caption         =   "�����鼮"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   8625
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "�����鼮����.frx":0000
      Left            =   5280
      List            =   "�����鼮����.frx":0013
      TabIndex        =   15
      Text            =   "��������B�ε�רҵ"
      Top             =   2280
      Width           =   2295
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
      ItemData        =   "�����鼮����.frx":004A
      Left            =   2760
      List            =   "�����鼮����.frx":0066
      TabIndex        =   14
      Text            =   "��������A�ε�רҵ"
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
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
      Left            =   3360
      TabIndex        =   12
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox Text1 
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
      Index           =   5
      Left            =   2760
      TabIndex        =   5
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox Text1 
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
      Index           =   4
      Left            =   2760
      TabIndex        =   4
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
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
      Index           =   3
      Left            =   2760
      TabIndex        =   3
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�½����"
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
      Left            =   600
      TabIndex        =   2
      Top             =   5280
      Width           =   2295
   End
   Begin VB.TextBox Text1 
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
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text1 
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
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "��ʽ��2022/3/22"
      Height          =   255
      Left            =   4800
      TabIndex        =   13
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "�����棺"
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
      Index           =   5
      Left            =   600
      TabIndex        =   11
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�������ʱ�䣺��Ĭ��Ϊ���죩"
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
      Index           =   4
      Left            =   600
      TabIndex        =   10
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "��������磺"
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
      Index           =   3
      Left            =   600
      TabIndex        =   9
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "����רҵ��"
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
      Index           =   2
      Left            =   600
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "����������"
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
      Index           =   1
      Left            =   600
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�����鼮�ţ�"
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
      Index           =   0
      Left            =   600
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "�����鼮����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

'�ȼ��һ�����ݶ�д��û
If Trim(Text1(0)) = empy Then MsgBox "���������", 48, "����": Exit Sub: Text1(0).SetFocus
If Trim(Text1(1)) = empy Then MsgBox "����������", 48, "����": Exit Sub: Text1(1).SetFocus
If Trim(Combo1.Text) = "��������A�ε�רҵ" And Trim(Combo2.Text) = "��������B�ε�רҵ" Then MsgBox "��ѡ��רҵ", 48, "����": Exit Sub
If Trim(Text1(3)) = empy Then MsgBox "�������������Ϣ", 48, "����": Exit Sub: Text1(3).SetFocus
If Trim(Text1(4)) = empy Then Text1(4) = Date
If Trim(Text1(5)) = empy Then MsgBox "��������", 48, "����": Exit Sub: Text1(5).SetFocus
If Format(Trim(Text1(4)), "yyyy/m/d") <> Trim(Text1(4)) Then MsgBox "��������ȷ��ʱ���ʽ,eg:2022/3/22": Exit Sub

'��Ϊcombolbox�ؼ��������������ñ�����ƽ��
Dim vocation As String
If Trim(Combo1.Text) <> "��������A�ε�רҵ" And Trim(Combo2.Text) <> "��������B�ε�רҵ" Then MsgBox "������ֻ��ѡ��һ��Ŷ����һ����ѡ��������һ����", 48, "����": Exit Sub
If Trim(Combo1.Text) = "��������A�ε�רҵ" Then vocation = Trim(Combo2.Text)
If Trim(Combo2.Text) = "��������B�ε�רҵ" Then vocation = Trim(Combo1.Text)

'�������ػ���
Dim rs As New ADODB.Recordset
Dim sql As String
sql = " select ���,���� from ���ϵͳ where ��� = " & "'" & Trim(Text1(0)) & "'"
rs.Open sql, ������Ӳ��ҽ���.conn, 1, 3
    '���￪ʼ��ʽ�Ŵ�
        If rs.EOF = False Then MsgBox "��鵽�ظ����������Ų��鼮��": Exit Sub
rs.Close

'���￪ʼ��������
rs.Open "���ϵͳ", ������Ӳ��ҽ���.conn, 1, 3
    rs.AddNew
    rs.Fields("���") = Trim(Text1(0))
    rs.Fields("����") = Trim(Text1(1))
    rs.Fields("רҵ") = vocation
    rs.Fields("������") = Trim(Text1(3))
    rs.Fields("���ʱ��") = Trim(Text1(4))
    rs.Fields("���") = Trim(Text1(5))
    rs.Update
rs.Close

'������ǵü�����ʷ��¼��
rs.Open "�鼮��Ϣ�����¼��", ������Ӳ��ҽ���.conn, 1, 3
    rs.AddNew
    rs.Fields("���") = Trim(Text1(0))
    rs.Fields("����") = Trim(Text1(1))
    rs.Fields("רҵ") = vocation
    rs.Fields("����ʱ��") = Trim(Text1(4))
    rs.Fields("���ĵ�����") = "����"
    rs.Fields(5) = "ȫ��"
    rs.Fields(6) = "��"
    rs.Fields(7) = "��"
    rs.Update
rs.Close
MsgBox "��ӳɹ�������������Ϊ��" & Chr(13) & "��ţ�" & Text1(0) & Space(5) & "��ţ�" & Text1(1) & Space(5) & "רҵ��" & vocation & Chr(13) & "�����磺" & Text1(3) & Space(5) & "���ʱ�䣺" & Text1(4) & Space(5) & "��棺" & Text1(5), , "��ʾ"

'ȫ����ɺ��˳��˴���
Unload Me
������Ӳ��ҽ���.Enabled = True
������Ӳ��ҽ���.Show
������Ӳ��ҽ���.maindata_load (" select * from  ���ϵͳ order by " & Trim(������Ӳ��ҽ���.Text2) & " desc ")
End Sub

Private Sub Command2_Click()
Unload Me
������Ӳ��ҽ���.Enabled = True
������Ӳ��ҽ���.Show
������Ӳ��ҽ���.maindata_load (" select * from  ���ϵͳ order by " & Trim(������Ӳ��ҽ���.Text2) & " desc ")
End Sub

Private Sub Form_Click()
'������Ӳ��ҽ���.conn
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 0
    ������Ӳ��ҽ���.Enabled = True
    ������Ӳ��ҽ���.Show
    ������Ӳ��ҽ���.maindata_load (" select * from  ���ϵͳ order by " & Trim(������Ӳ��ҽ���.Text2) & " desc ")
End Sub
