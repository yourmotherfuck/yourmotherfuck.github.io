VERSION 5.00
Begin VB.Form �鼮��¼����ϸ���ҽ��� 
   Caption         =   "0"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   10995
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
      Left            =   8040
      TabIndex        =   22
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
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
      Left            =   5520
      TabIndex        =   21
      Text            =   "����Ĭ�ϲ���ѡ"
      Top             =   3000
      Width           =   2055
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
      Height          =   735
      Left            =   7800
      TabIndex        =   18
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
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
      Height          =   735
      Left            =   5520
      TabIndex        =   17
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox Text8 
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
      Left            =   2760
      TabIndex        =   16
      Top             =   6360
      Width           =   2055
   End
   Begin VB.TextBox Text7 
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
      Left            =   2760
      TabIndex        =   15
      Top             =   5520
      Width           =   2055
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
      Left            =   2760
      TabIndex        =   14
      Top             =   4680
      Width           =   2055
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
      Left            =   2760
      TabIndex        =   13
      Top             =   3840
      Width           =   2055
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
      Left            =   2760
      TabIndex        =   12
      Top             =   3000
      Width           =   2055
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
      Left            =   2760
      TabIndex        =   11
      Top             =   2160
      Width           =   2055
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
      Left            =   2760
      TabIndex        =   10
      Top             =   1320
      Width           =   2055
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
      Left            =   2760
      TabIndex        =   9
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   20
      Top             =   3120
      Width           =   240
   End
   Begin VB.Label Label10 
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
      Left            =   6000
      TabIndex        =   19
      Top             =   4080
      Width           =   3135
   End
   Begin VB.Label Label9 
      Caption         =   "���ĺ�"
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
      Left            =   600
      TabIndex        =   8
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "����ǰ��"
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
      Left            =   600
      TabIndex        =   7
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "��������ݣ�"
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
      Left            =   600
      TabIndex        =   6
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "���ĵ����ͣ�"
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
      Left            =   600
      TabIndex        =   5
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label4 
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
      Left            =   600
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label3 
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
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
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
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "��ţ�"
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
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "ûʲô����ȫ����ѡ����������һ���е����ݣ��Զ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6000
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
End
Attribute VB_Name = "�鼮��¼����ϸ���ҽ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 0 Then
'ȡ����ѡʱ�̶���
    Text9.Enabled = False
    Text9 = "����Ĭ�ϲ���ѡ"
    Text4 = empy
Else
    Text9.Enabled = True
    '���������ǰ���Ƕ�Ϊ�պͲ�Ϊ��
    If Text4 = empy Then
            Text9.Text = Year(Date) & "/" & Month(Date) & "/"
            Text4.Text = Year(Date) & "/" & Month(Date) & "/"
    Else
        '�����ʼʱ�䲻Ϊ����ô������ĳɷ�
        If IsDate(Trim(Text4)) = False Then MsgBox "��������ȷ��ʱ���ʽ��eg:2022/3/3", 48, "����": Text4 = empy: Check1.Value = 0: Exit Sub
        '����������������isdate���� #2022/2#��ʽ ֱ�ӱ��� �� #3/4# ��ʽʱ��ǰ�油ȫ�����
        If Year(CDate(Text4)) & "/" & Month(Text4) = Trim(Text4) Then MsgBox "��������ȷ��ʱ���ʽ��eg:2022/3/3", 48, "����": Text4 = empy: Check1.Value = 0: Exit Sub
        If CStr(Month(Trim(Text4))) + "/" + CStr(Day(Trim(Text4))) = Trim(Text4) Then
            Text4 = CStr(Year(Date)) + "/" + Trim(Text4)
            Text9.SetFocus
        Else
            Text4.SetFocus
        End If
        Text9 = CStr(Year(CDate(Trim(Text4)))) + "/"
    End If
End If
End Sub

Private Sub Command1_Click()
�鼮��Ϣ�����¼�鿴����.Text1 = ""
'���Ŵ�
If Trim(Text1) = empy And Trim(Text9) = empy And Trim(Text3) = empy And Trim(Text4) = empy And Trim(Text5) = empy And Trim(Text6) = empy And Trim(Text7) = empy And Trim(Text8) = empy Then MsgBox "����ȫ��Ϊ��ֵ", 48, "����": Exit Sub
If Trim(Text4) <> empy And IsDate(Trim(Text4)) = False Then MsgBox "��������ȷ��ʱ���ʽ��", 48, "����": Exit Sub
'ʱ���Ŵ�
If Check1.Value = 1 Then
    '���ѡ�ѡ�����дʱ�䣬����
    If Trim(Text9) = empy Or Trim(Trim(Text4)) = empy Then MsgBox "��ѡ�������ı��򶼱�������ʱ�䣡", 48, "���棡": Exit Sub
    '�����ʱ�䲻�ܴ���ǰ��ģ�����
    If CDate(Trim(Text9)) < CDate(Trim(Text4)) Then MsgBox "�����ʱ��Ҫ����ǰ��ģ�", 48, "���棡": Exit Sub
End If


Dim sql As String
'��ʼ����SQL���
sql = "select * from �鼮��Ϣ�����¼�� where "

If Trim(Text1) <> empy Then sql = sql & " ��� like '%" & Trim(Text1.Text) & "%'"

If Trim(Text2) <> empy Then
    If Trim(Text1) = empy Then
        sql = sql & " ���� like '%" & Trim(Text2.Text) & "%' "
    Else
        sql = sql & " and ���� like '%" & Trim(Text2.Text) & "%' "
    End If
End If

If Trim(Text3) <> empy Then
    If Trim(Text1) = empy And Trim(Text2) = empy Then
        sql = sql & " רҵ like '%" & Trim(Text3.Text) & "%' "
    Else
        sql = sql & " and רҵ like '%" & Trim(Text3.Text) & "%' "
    End If
End If

If Trim(Text4.Text) <> empy Then
    If Trim(Text1) = empy And Trim(Text2) = empy And Trim(Text3) = empy Then
            If Check1.Value = 0 Then
                '�������ѡʱ������ѡ��
                sql = sql & " ����ʱ�� like #" & Trim(Text4.Text) & "# "
            Else
                '��ѡ����������
                sql = sql & " ����ʱ�� >= #" & Trim(Text4.Text) & "# and ����ʱ�� <= #" & Trim(Text9) & "# "
            End If
    Else
            If Check1.Value = 0 Then
                sql = sql & " and ����ʱ�� like #" & Trim(Text4.Text) & "# "
            Else
                sql = sql & " and ����ʱ�� >= #" & Trim(Text4.Text) & "# and ����ʱ�� <= #" & Trim(Text9) & "# "
            End If
    End If
End If

'If Trim(Text4) <> empy Then
'    If Trim(Text1) = empy And Trim(Text2) = empy And Trim(Text3) = empy Then
'        sql = sql & " ����ʱ�� like #" & Trim(Text4.Text) & "# "
'    Else
'        sql = sql & " and ����ʱ�� like #" & Trim(Text4.Text) & "# "
'    End If
'End If

If Trim(Text5) <> empy Then
    If Trim(Text1) = empy And Trim(Text2) = empy And Trim(Text3) = empy And Trim(Text4) = empy Then
        sql = sql & " ���ĵ����� like '%" & Trim(Text5.Text) & "%' "
    Else
        sql = sql & " and ���ĵ����� like '%" & Trim(Text5.Text) & "%' "
    End If
End If

If Trim(Text6) <> empy Then
    If Trim(Text1) = empy And Trim(Text2) = empy And Trim(Text3) = empy And Trim(Text4) = empy And Trim(Text5) = empy Then
        sql = sql & " ��������� like '%" & Trim(Text5.Text) & "%' "
    Else
        sql = sql & " and ��������� like '%" & Trim(Text5.Text) & "%' "
    End If
End If

If Trim(Text7) <> empy Then
    If Trim(Text1) = empy And Trim(Text2) = empy And Trim(Text3) = empy And Trim(Text4) = empy And Trim(Text5) = empy And Trim(Text6) = empy Then
        sql = sql & " ����ǰ like '%" & Trim(Text5.Text) & "%' "
    Else
        sql = sql & " and ����ǰ like '%" & Trim(Text5.Text) & "%' "
    End If
End If

If Trim(Text8) <> empy Then
    If Trim(Text1) = empy And Trim(Text2) = empy And Trim(Text3) = empy And Trim(Text4) = empy And Trim(Text5) = empy And Trim(Text6) = empy And Trim(Text7) = empy Then
        sql = sql & " ���ĺ� like '%" & Trim(Text5.Text) & "%' "
    Else
        sql = sql & " and ���ĺ� like '%" & Trim(Text5.Text) & "%' "
    End If
End If

sql = sql & "order by " & �鼮��Ϣ�����¼�鿴����.Combo1.Text

�鼮��Ϣ�����¼�鿴����.bookhistory_load sql, True

Unload Me
�鼮��Ϣ�����¼�鿴����.Enabled = True
�鼮��Ϣ�����¼�鿴����.Show
End Sub

Private Sub Command2_Click()
Unload Me
�鼮��Ϣ�����¼�鿴����.Show
�鼮��Ϣ�����¼�鿴����.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
�鼮��Ϣ�����¼�鿴����.Enabled = True
�鼮��Ϣ�����¼�鿴����.Show
End Sub

Private Sub Text4_GotFocus()
Text4.Text = Year(Date) & "/" & Month(Date) & "/"
End Sub
