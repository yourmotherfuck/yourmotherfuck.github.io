VERSION 5.00
Begin VB.Form ��ϸ���ҽ��� 
   Caption         =   "��ϸ���ҽ���"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   10155
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox Check1 
      Caption         =   "����ʱ������"
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
      Left            =   7560
      TabIndex        =   18
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
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
      Left            =   5160
      TabIndex        =   16
      Text            =   "����Ĭ�ϲ���ѡ"
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   15
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   14
      Top             =   5520
      Width           =   1575
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
      Left            =   2520
      TabIndex        =   12
      Top             =   4560
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
      Index           =   4
      Left            =   2520
      TabIndex        =   11
      Top             =   3840
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
      Index           =   3
      Left            =   2520
      TabIndex        =   10
      Top             =   3120
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
      Index           =   2
      Left            =   2520
      TabIndex        =   9
      Top             =   2400
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
      Index           =   1
      Left            =   2520
      TabIndex        =   8
      Top             =   1680
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
      Index           =   0
      Left            =   2520
      TabIndex        =   7
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "ʱ����ӹ��򣺹�ѡ����Զ���ȫ��ǰ��ݣ����Ҫʵ�����ں�����Ч���ڶ�Ӧ�ĵط�����һ���ܴ��ʱ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5520
      TabIndex        =   19
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   4680
      TabIndex        =   17
      Top             =   3960
      Width           =   270
   End
   Begin VB.Label Label2 
      Caption         =   "ע�⣺�˽��������ɢ��ƥ�䷽�����������п��ǿ�ѡ��            ��ѡ�������ͨ���*"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5520
      TabIndex        =   13
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "���ʱ��"
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
      Left            =   840
      TabIndex        =   6
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "���"
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
      Left            =   840
      TabIndex        =   5
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "������"
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
      Left            =   840
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "רҵ"
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
      Left            =   840
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label1 
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
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "����"
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
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "���"
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
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "��ϸ���ҽ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 0 Then
'ȡ����ѡʱ�̶���
    Text2.Enabled = False
    Text2 = "����Ĭ�ϲ���ѡ"
    Text1(4) = empy
Else
    Text2.Enabled = True
    '���������ǰ���Ƕ�Ϊ�պͲ�Ϊ��
    If Text1(4) = empy Then
            Text2.Text = Year(Date) & "/" & Month(Date) & "/"
            Text1(4).Text = CStr(Year(Date)) & "/" & Month(Date) & "/"
    Else
        '�����ʼʱ�䲻Ϊ����ô������ĳɷ�
        If IsDate(Trim(Text1(4))) = False Then MsgBox "��������ȷ��ʱ���ʽ��eg:2022/3/3", 48, "����": Text1(4) = empy: Check1.Value = 0: Exit Sub
        '����������������isdate���� #2022/2#��ʽ ֱ�ӱ��� �� #3/4# ��ʽʱ��ǰ�油ȫ�����
        If Year(CDate(Text1(4))) & "/" & Month(Text1(4)) = Trim(Text1(4)) Then MsgBox "��������ȷ��ʱ���ʽ��eg:2022/3/3", 48, "����": Text1(4) = empy: Check1.Value = 0: Exit Sub
        If CStr(Month(Trim(Text1(4)))) + "/" + CStr(Day(Trim(Text1(4)))) = Trim(Text1(4)) Then
            Text1(4) = CStr(Year(Date)) + "/" + Trim(Text1(4))
            Text2.SetFocus
        Else
            Text1(4).SetFocus
        End If
        Text2 = CStr(Year(CDate(Trim(Text1(4))))) + "/"
    End If
End If
End Sub

Private Sub Command1_Click()
'�Ϲ�أ����Ŵ�
If Trim(Text1(5).Text) <> empy And Val(Trim(Text1(5).Text)) <> Trim(Text1(5).Text) Then MsgBox "����ҿ�������봿����", 48, "���棡": Exit Sub
If Trim(Text1(0).Text) = empy And Trim(Text1(1).Text) = empy And Trim(Text1(2).Text) = empy And Trim(Text1(3).Text) = empy And Trim(Text1(4).Text) = empy And Trim(Text1(5).Text) = empy Then MsgBox "����������һ����", 48, "����", 48, "���棡": Exit Sub

If Trim(Text1(4)) <> empy And IsDate(Trim(Text1(4).Text)) = False Then MsgBox "��������ȷ��ʱ���ʽ,eg:2022/3/22", 48, "���棡": Exit Sub

If Check1.Value = 1 Then
    '���ѡ�ѡ�����дʱ�䣬����
    If Trim(Text2) = empy Or Trim(Trim(Text1(4))) = empy Then MsgBox "��ѡ�������ı��򶼱�������ʱ�䣡", 48, "���棡": Exit Sub
    '�����ʱ�䲻�ܴ���ǰ��ģ�����
    If CDate(Trim(Text2)) < CDate(Trim(Text1(4))) Then MsgBox "�����ʱ��Ҫ����ǰ��ģ�", 48, "���棡": Exit Sub
End If

Dim sql As String
sql = "select * from ���ϵͳ where "

If Trim(Text1(0).Text) <> empy Then sql = sql & " ��� like '%" & Trim(Text1(0).Text) & "%'"

If Trim(Text1(1).Text) <> empy Then
    If Trim(Text1(0).Text) = empy Then
        sql = sql & " ���� like '%" & Trim(Text1(1).Text) & "%'"
    Else
        sql = sql & " and ���� like '%" & Trim(Text1(1).Text) & "%'"
    End If
End If

If Trim(Text1(2).Text) <> empy Then
    If Trim(Text1(0).Text) = empy And Trim(Text1(1).Text) = empy Then
        sql = sql & " רҵ like '%" & Trim(Text1(2).Text) & "%'"
    Else
        sql = sql & " and רҵ like '%" & Trim(Text1(2).Text) & "%'"
    End If
End If

If Trim(Text1(3).Text) <> empy Then
    If Trim(Text1(0).Text) = empy And Trim(Text1(1).Text) = empy And Trim(Text1(2).Text) = empy Then
        sql = sql & " ������ like '%" & Trim(Text1(3).Text) & "%'"
    Else
        sql = sql & " and ������ like '%" & Trim(Text1(3).Text) & "%'"
    End If
End If

If Trim(Text1(4).Text) <> empy Then
    If Trim(Text1(0).Text) = empy And Trim(Text1(1).Text) = empy And Trim(Text1(2).Text) = empy And Trim(Text1(3).Text) = empy Then
        If Check1.Value = 0 Then
            '�������ѡʱ������ѡ��
            sql = sql & " ���ʱ�� like #" & Trim(Text1(4).Text) & "# "
        Else
            '��ѡ����������
            sql = sql & " ���ʱ�� >= #" & Trim(Text1(4).Text) & "# and ���ʱ�� <= #" & Trim(Text2) & "# "
        End If
    Else
        If Check1.Value = 0 Then
            sql = sql & " and ���ʱ�� like #" & Trim(Text1(4).Text) & "# "
        Else
            sql = sql & " and ���ʱ�� >= #" & Trim(Text1(4).Text) & "# and ���ʱ�� <= #" & Trim(Text2) & "# "
        End If
    End If
End If

If Trim(Text1(5).Text) <> empy Then
    If Trim(Text1(0).Text) = empy And Trim(Text1(1).Text) = empy And Trim(Text1(2).Text) = empy And Trim(Text1(3).Text) = empy And Trim(Text1(4).Text) = empy Then
        sql = sql & " ��� = " & Trim(Text1(5).Text)
    Else
        sql = sql & " and ��� = " & Trim(Text1(5).Text)
    End If
End If
sql = sql + " order by " & Trim(������Ӳ��ҽ���.Combo2.Text)

Dim rs As New ADODb.Recordset
rs.Open sql, ������Ӳ��ҽ���.conn, 1, 3
    If rs.EOF Then MsgBox "�����û����֮ƥ������ݣ�������ٴ�����", 48, "����": Exit Sub
rs.Close

������Ӳ��ҽ���.maindata_load sql, True
Unload Me
������Ӳ��ҽ���.Enabled = True
������Ӳ��ҽ���.Show
End Sub

Private Sub Command2_Click()
Unload Me
������Ӳ��ҽ���.Enabled = True
������Ӳ��ҽ���.Show
'������Ӳ��ҽ���.maindata_load (" select * from  ���ϵͳ order by " & Trim(������Ӳ��ҽ���.Combo2.Text) & " desc ")
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 0
    ������Ӳ��ҽ���.Enabled = True
    ������Ӳ��ҽ���.Show
    '������Ӳ��ҽ���.maindata_load (" select * from  ���ϵͳ order by " & Trim(������Ӳ��ҽ���.Combo2.Text) & " desc ")
End Sub

