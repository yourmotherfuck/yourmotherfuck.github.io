VERSION 5.00
Begin VB.Form ������� 
   Caption         =   "�������"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   9600
   StartUpPosition =   3  '����ȱʡ
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
      TabIndex        =   11
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��������"
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
      Left            =   840
      TabIndex        =   10
      Top             =   4680
      Width           =   2175
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
      Left            =   3120
      TabIndex        =   9
      Top             =   3600
      Width           =   2655
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
      Left            =   3120
      TabIndex        =   8
      Top             =   2880
      Width           =   2655
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
      Index           =   2
      Left            =   3120
      TabIndex        =   7
      Top             =   2160
      Width           =   2655
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
      Left            =   3120
      TabIndex        =   6
      Top             =   1440
      Width           =   2655
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
      Left            =   3120
      TabIndex        =   5
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "�������������"
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
      Left            =   720
      TabIndex        =   4
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "�������ʱ�䣺"
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
      Left            =   720
      TabIndex        =   3
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "��������ˣ���λ��"
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
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
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
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
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
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "�������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'���������ڱ����˾�����������
'On Error Resume Next

'���չ������ȼ��һ��
If Trim(Text1(0)) = empy Then MsgBox "�������鼮��", 48, "ע��": Exit Sub
If Trim(Text1(1)) = empy Then MsgBox "����������", 48, "ע��": Exit Sub
If Trim(Text1(2)) = empy Then MsgBox "����������ˣ���λ��", 48, "ע��": Exit Sub
If Trim(Text1(3)) = empy Then Text1(3) = Date
If Trim(Text1(4)) = empy Then MsgBox "�������������", 48, "ע��": Exit Sub
If IsDate(Trim(Text1(3))) = False Then MsgBox "��������ȷ��ʱ���ʽ,eg:2022/3/22": Exit Sub

'�Ŵ���ƣ�Ҫȷ���������������Ӧ����ź�������
Dim rs As New ADODB.Recordset
Dim sql As String
sql = "select ���,���� from ���ϵͳ where ��� =" & "'" & Trim(Text1(0)) & "'" & "and ���� = " & "'" & Trim(Text1(1)) & "'"
rs.Open sql, ������Ӳ��ҽ���.conn, 1, 3
    If rs.EOF Then MsgBox "�����û���������������������鼮�ź�������Ӧ�������Ƿ���ȷ", 48, "ע��"
rs.Close


'���￪ʼͬ����������
sql = "select * from ���ϵͳ where ��� =  " & "'" & Trim(Text1(0)) & "'"
rs.Open sql, ������Ӳ��ҽ���.conn, 1, 3
    '���õ�������Ҫ���鼮��Ϣ������רҵ�ͱ��ǰ������
    Dim vocation As String, value_before As String
    vocation = rs.Fields("רҵ").Value
    value_before = rs.Fields("���").Value
    
    reserve = rs.Fields("���").Value - Val(Trim(Text1(4)))
    If reserve < 0 Then MsgBox "�����������ܴ��ڿ��", 48, "����": Exit Sub
    rs.Fields("���") = reserve
    rs.Update
rs.Close


'�����ݼ�������¼��
rs.Open "�����¼��", ������Ӳ��ҽ���.conn, 1, 3
    rs.AddNew
    rs.Fields("���") = Trim(Text1(0))
    rs.Fields("����") = Trim(Text1(1))
    rs.Fields("�����ˣ���λ��") = Trim(Text1(2))
    rs.Fields("����ʱ��") = Trim(Text1(3))
    rs.Fields("��������") = Trim(Text1(4))
    rs.Update
rs.Close

'ͬʱ���鼮��Ϣ�������������Ӧ�Ķ���
temp = "��" + Trim(Text1(4)) + "����"
rs.Open "�鼮��Ϣ�����¼��", ������Ӳ��ҽ���.conn, 1, 3
    rs.AddNew
    rs.Fields("���") = Trim(Text1(0))
    rs.Fields("����") = Trim(Text1(1))
    rs.Fields("רҵ") = vocation
    rs.Fields("����ʱ��") = Trim(Text1(3))
    rs.Fields("���ĵ�����") = "����"
    rs.Fields("���������") = temp
    rs.Fields("����ǰ") = value_before
    rs.Fields("���ĺ�") = reserve
    rs.Update
rs.Close


MsgBox "����ɹ�����������Ϊ��" & Chr(13) & "�鼮�ţ�" & Trim(Text1(0)) & Space(5) & "������" & Trim(Text1(1)) & Space(5) & "�����ˣ���λ����" & Trim(Text1(2)) & Chr(13) & "����ʱ�䣺" & Trim(Text1(3)) & Space(5) & "����������" & Trim(Text1(4)), , "ͨ��"

'��ɺ�ص�������
Unload Me
������Ӳ��ҽ���.Enabled = True
������Ӳ��ҽ���.Show
������Ӳ��ҽ���.maindata_load " select * from  ���ϵͳ order by " & Trim(������Ӳ��ҽ���.Combo2.Text), True
End Sub

Private Sub Command2_Click()
Unload Me
������Ӳ��ҽ���.Enabled = True
������Ӳ��ҽ���.Show
'������Ӳ��ҽ���.maindata_load (" select * from  ���ϵͳ order by " & Trim(������Ӳ��ҽ���.Combo2.Text) & " desc ")
End Sub

Private Sub Form_Load()
'������Ӳ��ҽ���.conn
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 0
    ������Ӳ��ҽ���.Enabled = True
    ������Ӳ��ҽ���.Show
    '������Ӳ��ҽ���.maindata_load (" select * from  ���ϵͳ order by " & Trim(������Ӳ��ҽ���.Combo2.Text) & " desc ")
End Sub
