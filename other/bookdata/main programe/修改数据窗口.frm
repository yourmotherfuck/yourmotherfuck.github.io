VERSION 5.00
Begin VB.Form �޸����ݴ��� 
   Caption         =   "�޸����ݴ���"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   10695
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "�޸����ݴ���.frx":0000
      Left            =   3000
      List            =   "�޸����ݴ���.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3600
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ȫ"
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
      Left            =   5880
      TabIndex        =   16
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
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
      Left            =   8160
      TabIndex        =   14
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Default         =   -1  'True
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
      Left            =   5880
      TabIndex        =   13
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox Text6 
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
      Left            =   3120
      TabIndex        =   12
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox Text5 
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
      Left            =   3120
      TabIndex        =   11
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox Text4 
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
      Left            =   3120
      TabIndex        =   10
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox Text3 
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
      Left            =   3120
      TabIndex        =   9
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text2 
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
      Left            =   3120
      TabIndex        =   8
      Top             =   1440
      Width           =   2175
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
      Left            =   3120
      TabIndex        =   7
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "ʹ�ò�ȫֻ��Ҫ������������žͿ����ˣ�����д�Ϻ�Ż�ȥƥ����ǰ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   18
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "�鼮�ţ�������������д��������רҵ�����ǰ���Բ�д����ȫ��,���ʱ��Ĭ�Ͻ�����ǰ��רҵѡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6360
      TabIndex        =   15
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "�����"
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
      Index           =   6
      Left            =   720
      TabIndex        =   6
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "���ǰ��"
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
      Index           =   5
      Left            =   720
      TabIndex        =   5
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "������ݵ�������"
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
      Index           =   4
      Left            =   720
      TabIndex        =   4
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "����ʱ�䣺"
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
      Index           =   3
      Left            =   720
      TabIndex        =   3
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "רҵ��"
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
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "����������"
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
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "�������鼮�ţ�"
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
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "�޸����ݴ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'�Ϲ�����Ŵ����ǰ��רҵѡ��
If Trim(Text1) = empy Then MsgBox "�������鼮�ţ�", 48, "����": Exit Sub
If Trim(Text2) = empy Then MsgBox "������������", 48, "����": Exit Sub
If Trim(Text4) = empy Then Text4 = Date
If Combo1.Text = Combo1.List(0) Then MsgBox "��ѡ��Ҫ���ĵ�����", 48, "����": Exit Sub
If Trim(Text6) = empy Then MsgBox "���������������", 48, "����": Exit Sub
If Trim(Text4) <> empy And IsDate(Trim(Text4)) = False Then MsgBox "��������ȷ��ʱ���ʽ", 48, "����"
If Combo1.Text = "���" Then If CStr(Int((Val(Trim(Text6))))) <> Trim(Text6) Then MsgBox "���ĺ������������������", 48, "����"
If Combo1.Text = "���ʱ��" Then If IsDate(Trim(Text6)) = -False Then MsgBox "���ĺ������ҲҪ����ʱ���ʽд eg:2022/3/3", 48, "����"

Dim rs As New ADODb.Recordset
Dim sql As String
'����Ҫ��ȷ��������飬˳����������Ȿ�������
sql = "select * from ���ϵͳ  where ��� ='" & Trim(Text1) & "'"
rs.Open sql, ������Ӳ��ҽ���.conn, 1, 3
    If rs.EOF = True Then MsgBox "���ϵͳ����û�ж�Ӧ����ţ���ȷ�Ϻ��ٴ�����", 48, "����": Exit Sub
    MsgBox "��������Ϊ:" & Chr(13) & " ���=" & rs.Fields("���").Value & "  ����=" & rs.Fields("����").Value & "  רҵ=" & rs.Fields("רҵ").Value & "  " & Chr(13) & "������=" & rs.Fields("������").Value & "  ���ʱ��=" & rs.Fields("���ʱ��").Value & "  ���=" & rs.Fields("���").Value
rs.Close

'���￪ʼ��������
a = MsgBox("���Ϊ% " & Trim(Text1) & " %��" & Combo1.Text & "���ݽ����Ϊ  " & Trim(Text6) & "  " & Chr(13) & "��ȷ��Ҫ��ô����", vbYesNo, "ע��")
If a = vbNo Then Exit Sub

If Combo1.Text = "���ʱ��" Then
    sql = " update ���ϵͳ set " & Combo1.Text & " = #" & Trim(Text6) & "# where ��� = '" & Trim(Text1) & "'"
ElseIf Combo1.Text = "���" Then
    sql = " update ���ϵͳ set " & Combo1.Text & " = " & Trim(Text6) & " where ��� = '" & Trim(Text1) & "'"
Else
    sql = " update ���ϵͳ set " & Combo1.Text & " = '" & Trim(Text6) & "' where ��� = '" & Trim(Text1) & "'"
End If
rs.Open sql, ������Ӳ��ҽ���.conn, 1, 3


'�޸������ݺ�ǵü��뵽��Ϣ��ʷ��¼������
rs.Open "�鼮��Ϣ�����¼��", ������Ӳ��ҽ���.conn, 1, 3
    rs.AddNew
    rs.Fields("���") = Trim(Text1)
    rs.Fields("����") = Trim(Text2)
    rs.Fields("רҵ") = Trim(Text3)
    rs.Fields("����ʱ��") = Trim(Text4)
    rs.Fields("���ĵ�����") = "�ֶ��޸�����"
    rs.Fields("���������") = Combo1.Text
    rs.Fields("����ǰ") = Trim(Text5)
    rs.Fields("���ĺ�") = Trim(Text6)
    rs.Update
rs.Close

'ȫ����ɺ��˳��˴���
Unload Me
������Ӳ��ҽ���.Enabled = True
������Ӳ��ҽ���.Show
������Ӳ��ҽ���.maindata_load " select * from  ���ϵͳ order by " & Trim(������Ӳ��ҽ���.Combo2.Text), True
End Sub

Private Sub Command2_Click()
Unload Me
������Ӳ��ҽ���.Enabled = True
������Ӳ��ҽ���.Show
End Sub

Private Sub Command3_Click()
If Trim(Text1) = empy And Trim(Text2) = empy Then MsgBox "��������������ţ�", 48, "����": Exit Sub
If Trim(Text4) = empy Then Text4 = Date

Dim rs As New ADODb.Recordset
Dim sql As String
If Trim(Text2) <> empy Then sql = "select * from ���ϵͳ where ���� = '" & Trim(Text2) & "'"
If Trim(Text1) <> empy Then sql = "select * from ���ϵͳ where ��� = '" & Trim(Text1) & "'"

rs.Open sql, ������Ӳ��ҽ���.conn, 1, 3
    Text2 = rs.Fields("����").Value
    Text3 = rs.Fields("רҵ").Value
    If Combo1.Text <> Combo1.List(0) Then Text5 = rs.Fields(Combo1.Text).Value
rs.Close
End Sub


Private Sub Form_Load()
Combo1.Text = Combo1.List(0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 0
Unload Me
������Ӳ��ҽ���.Enabled = True
������Ӳ��ҽ���.Show
End Sub
