VERSION 5.00
Begin VB.Form ��ϸ���ҽ��� 
   Caption         =   "��ϸ���ҽ���"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   8610
   StartUpPosition =   3  '����ȱʡ
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
      Left            =   5400
      TabIndex        =   15
      Top             =   4440
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
      Left            =   5400
      TabIndex        =   14
      Top             =   3480
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
      Index           =   4
      Left            =   2520
      TabIndex        =   11
      Top             =   3840
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
      Index           =   3
      Left            =   2520
      TabIndex        =   10
      Top             =   3120
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
      Index           =   2
      Left            =   2520
      TabIndex        =   9
      Top             =   2400
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
      Index           =   1
      Left            =   2520
      TabIndex        =   8
      Top             =   1680
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
      Index           =   0
      Left            =   2520
      TabIndex        =   7
      Top             =   960
      Width           =   1575
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
      Height          =   1815
      Left            =   4800
      TabIndex        =   13
      Top             =   960
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
Private Sub Command1_Click()
'�Ϲ�أ����Ŵ�
If Format(Trim(Text1(4)), "yyyy/m/d") <> Trim(Text1(4)) Then MsgBox "��������ȷ��ʱ���ʽ,eg:2022/3/22": Exit Sub
If Trim(Text1(5).Text) <> empy And Val(Trim(Text1(5).Text)) <> Trim(Text1(5).Text) Then MsgBox "����ҿ�������봿����": Exit Sub
If Trim(Text1(0).Text) = empy And Trim(Text1(1).Text) = empy And Trim(Text1(2).Text) = empy And Trim(Text1(3).Text) = empy And Trim(Text1(4).Text) = empy And Trim(Text1(5).Text) = empy Then MsgBox "����������һ����", 48, "����": Exit Sub

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
        sql = sql & " ���ʱ�� = #" & Trim(Text1(4).Text) & "#"
    Else
        sql = sql & " and ���ʱ�� = #" & Trim(Text1(4).Text) & "#"
    End If
End If

If Trim(Text1(5).Text) <> empy Then
    If Trim(Text1(0).Text) = empy And Trim(Text1(1).Text) = empy And Trim(Text1(2).Text) = empy And Trim(Text1(3).Text) = empy And Trim(Text1(4).Text) = empy Then
        sql = sql & " ��� = " & Trim(Text1(5).Text)
    Else
        sql = sql & " and ��� = " & Trim(Text1(5).Text)
    End If
End If


������Ӳ��ҽ���.Enabled = True
������Ӳ��ҽ���.Show
������Ӳ��ҽ���.maindata_load (sql)
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
������Ӳ��ҽ���.Enabled = True
������Ӳ��ҽ���.Show
������Ӳ��ҽ���.maindata_load (" select * from  ���ϵͳ order by " & Trim(������Ӳ��ҽ���.Text2) & " desc ")
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 0
    ������Ӳ��ҽ���.Enabled = True
    ������Ӳ��ҽ���.Show
    ������Ӳ��ҽ���.maindata_load (" select * from  ���ϵͳ order by " & Trim(������Ӳ��ҽ���.Text2) & " desc ")
End Sub
