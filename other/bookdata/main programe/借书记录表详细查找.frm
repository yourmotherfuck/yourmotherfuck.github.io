VERSION 5.00
Begin VB.Form �����¼����ϸ���� 
   Caption         =   "�����¼����ϸ����"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   7935
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
      Left            =   5160
      TabIndex        =   16
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox Text6 
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
      Left            =   5160
      TabIndex        =   13
      Text            =   "����Ĭ�ϲ���ѡ"
      Top             =   3000
      Width           =   1935
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
      Left            =   2760
      TabIndex        =   11
      Top             =   4680
      Width           =   1815
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
      Left            =   600
      TabIndex        =   10
      Top             =   4680
      Width           =   1815
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
      Left            =   2640
      TabIndex        =   9
      Top             =   3720
      Width           =   1935
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
      Left            =   2640
      TabIndex        =   8
      Top             =   3000
      Width           =   1935
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
      Left            =   2640
      TabIndex        =   7
      Top             =   2280
      Width           =   1935
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
      Left            =   2640
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
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
      Left            =   2640
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "ʱ���ʽ�����ö�����"
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
      Left            =   4920
      TabIndex        =   15
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label7 
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
      Left            =   4800
      TabIndex        =   14
      Top             =   3120
      Width           =   240
   End
   Begin VB.Label Label6 
      Caption         =   "ûʲô����ȫ����ѡ����������һ���е����ݣ��Զ�����"
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
      Left            =   5040
      TabIndex        =   12
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "��������"
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
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "����ʱ��"
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
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "�����ˣ���λ��"
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
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
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
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
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
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "�����¼����ϸ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 0 Then
'ȡ����ѡʱ�̶���
    Text6.Enabled = False
    Text6 = "����Ĭ�ϲ���ѡ"
    Text4 = empy
Else
    Text6.Enabled = True
    '���������ǰ���Ƕ�Ϊ�պͲ�Ϊ��
    If Text4 = empy Then
            Text6.Text = Year(Date) & "/" & Month(Date) & "/"
            Text4.Text = Year(Date) & "/" & Month(Date) & "/"
    Else
        '�����ʼʱ�䲻Ϊ����ô������ĳɷ�
        If IsDate(Trim(Text4)) = False Then MsgBox "��������ȷ��ʱ���ʽ��eg:2022/3/3", 48, "����": Text4 = empy: Check1.Value = 0: Exit Sub
        '����������������isdate���� #2022/2#��ʽ ֱ�ӱ��� �� #3/4# ��ʽʱ��ǰ�油ȫ�����
        If Year(CDate(Text4)) & "/" & Month(Text4) = Trim(Text4) Then MsgBox "��������ȷ��ʱ���ʽ��eg:2022/3/3", 48, "����": Text4 = empy: Check1.Value = 0: Exit Sub
        If CStr(Month(Trim(Text4))) + "/" + CStr(Day(Trim(Text4))) = Trim(Text4) Then
            Text4 = CStr(Year(Date)) + "/" + Trim(Text4)
            Text6.SetFocus
        Else
            Text4.SetFocus
        End If
        Text6 = CStr(Year(CDate(Trim(Text4)))) + "/"
    End If
End If
End Sub

Private Sub Command1_Click()
'���Ŵ�
If Trim(Text1) = empy And Trim(Text2) = empy And Trim(Text3) = empy And Trim(Text4) = empy And Trim(Text5) = empy Then MsgBox "����ȫ��Ϊ��ֵ", 48, "����": Exit Sub
If Trim(Text4) <> empy And IsDate(Trim(Text4)) = False Then MsgBox "��������ȷ��ʱ���ʽ��", 48, "����": Exit Sub
If Trim(Text5) <> empy And CStr(Int(Val(Trim(Text5)))) <> Trim(Text5) Then MsgBox "��������������������", 48, "����": Exit Sub

'ʱ���Ŵ�
If Check1.Value = 1 Then
    '���ѡ�ѡ�����дʱ�䣬����
    If Trim(Text6) = empy Or Trim(Trim(Text4)) = empy Then MsgBox "��ѡ�������ı��򶼱�������ʱ�䣡", 48, "���棡": Exit Sub
    '�����ʱ�䲻�ܴ���ǰ��ģ�����
    If CDate(Trim(Text6)) < CDate(Trim(Text4)) Then MsgBox "�����ʱ��Ҫ����ǰ��ģ�", 48, "���棡": Exit Sub
End If


Dim sql As String
'��ʼ����SQL���
sql = "select * from �����¼�� where "

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
        sql = sql & " �����ˣ���λ�� like '%" & Trim(Text3.Text) & "%' "
    Else
        sql = sql & " and �����ˣ���λ�� like '%" & Trim(Text3.Text) & "%' "
    End If
End If

If Trim(Text4.Text) <> empy Then
    If Trim(Text1) = empy And Trim(Text2) = empy And Trim(Text3) = empy Then
            If Check1.Value = 0 Then
                '�������ѡʱ������ѡ��
                sql = sql & " ����ʱ�� like #" & Trim(Text4.Text) & "# "
            Else
                '��ѡ����������
                sql = sql & " ����ʱ�� >= #" & Trim(Text4.Text) & "# and ����ʱ�� <= #" & Trim(Text6) & "# "
            End If
    Else
            If Check1.Value = 0 Then
                sql = sql & " and ����ʱ�� like #" & Trim(Text4.Text) & "# "
            Else
                sql = sql & " and ����ʱ�� >= #" & Trim(Text4.Text) & "# and ����ʱ�� <= #" & Trim(Text6) & "# "
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
        sql = sql & " �������� = " & Trim(Text5.Text) & " "
    Else
        sql = sql & " and �������� = " & Trim(Text5.Text) & " "
    End If
End If

sql = sql & "order by " & �����¼�鿴����.Combo1.Text

�����¼�鿴����.senthistory_load sql, True

Unload Me
�����¼�鿴����.Show
�����¼�鿴����.Enabled = True
End Sub

Private Sub Command2_Click()
Unload Me
�����¼�鿴����.Show
�����¼�鿴����.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
�����¼�鿴����.Enabled = True
�����¼�鿴����.Show
End Sub


Private Sub Text4_GotFocus()
Text4.Text = Year(Date) & "/" & Month(Date) & "/"
End Sub
