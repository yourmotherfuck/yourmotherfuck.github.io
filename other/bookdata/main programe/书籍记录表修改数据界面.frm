VERSION 5.00
Begin VB.Form �鼮��¼���޸����ݽ��� 
   Caption         =   "�鼮��¼���޸����ݽ���"
   ClientHeight    =   9390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13260
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   13260
   StartUpPosition =   3  '����ȱʡ
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
      ItemData        =   "�鼮��¼���޸����ݽ���.frx":0000
      Left            =   9240
      List            =   "�鼮��¼���޸����ݽ���.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   4680
      Width           =   2655
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   9360
      TabIndex        =   22
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "��׼���Ҵ���"
      Height          =   7455
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   6015
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
         TabIndex        =   12
         Top             =   600
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
         TabIndex        =   11
         Top             =   1440
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
         TabIndex        =   10
         Top             =   2280
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
         TabIndex        =   9
         Top             =   3120
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
         TabIndex        =   8
         Top             =   3960
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
         TabIndex        =   7
         Top             =   4800
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
         TabIndex        =   6
         Top             =   5640
         Width           =   2055
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
         TabIndex        =   5
         Top             =   6480
         Width           =   2055
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
         TabIndex        =   20
         Top             =   600
         Width           =   1815
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
         TabIndex        =   19
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label3 
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
         Left            =   600
         TabIndex        =   18
         Top             =   2280
         Width           =   1815
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
         TabIndex        =   17
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "���ĵ�����"
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
         TabIndex        =   16
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "���������"
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
         TabIndex        =   15
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "����ǰ"
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
         TabIndex        =   14
         Top             =   5640
         Width           =   1815
      End
      Begin VB.Label Label8 
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
         TabIndex        =   13
         Top             =   6480
         Width           =   1815
      End
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
      Left            =   9480
      TabIndex        =   1
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
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
      Height          =   735
      Left            =   7080
      TabIndex        =   0
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label12 
      Caption         =   "Ҫ�����������"
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
      Left            =   7200
      TabIndex        =   23
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "�����ݱ��Ϊ��"
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
      Left            =   7200
      TabIndex        =   21
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label99 
      Caption         =   "ǰ����Щ������У׼���е�"
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
      TabIndex        =   3
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label Label100 
      Caption         =   "��������ݸ��Ĳ����¼���κα��У�û�л�ԭ�Ŀ��ܣ���������Ļ�ɾ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7800
      TabIndex        =   2
      Top             =   1320
      Width           =   3015
   End
End
Attribute VB_Name = "�鼮��¼���޸����ݽ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'�Ϲ�����Ŵ������������еĶ������Ǳ�ѡ�ģ�
If Trim(Text1) = empy Then MsgBox "������У׼�����ݣ���ȫ����Ҫ����", 48, "����": Exit Sub
If Trim(Text2) = empy Then MsgBox "������У׼�����ݣ���ȫ����Ҫ����", 48, "����": Exit Sub
If Trim(Text3) = empy Then MsgBox "������У׼�����ݣ���ȫ����Ҫ����", 48, "����": Exit Sub
If Trim(Text4) = empy Then MsgBox "������У׼�����ݣ���ȫ����Ҫ����", 48, "����": Exit Sub
If IsDate(Trim(Text4)) = False Then MsgBox "��������ȷ��ʱ���ʽ", 48, "����": Exit Sub
If Trim(Text5) = empy Then MsgBox "������У׼�����ݣ���ȫ����Ҫ����", 48, "����": Exit Sub
If Trim(Text6) = empy Then MsgBox "������У׼�����ݣ���ȫ����Ҫ����", 48, "����": Exit Sub
If Trim(Text7) = empy Then MsgBox "������У׼�����ݣ���ȫ����Ҫ����", 48, "����": Exit Sub
If Trim(Text8) = empy Then MsgBox "������У׼�����ݣ���ȫ����Ҫ����", 48, "����": Exit Sub

If Trim(Combo2.Text) = Combo2.List(0) Then MsgBox "��ѡ��Ҫ�������ݵ�������", 48, "����": Exit Sub
If Trim(Combo2.Text) = Combo2.List(4) And IsDate(Trim(Text10)) = False Then MsgBox "���ĺ��������Ϊʱ���ʽ��eg:2022/3/3", 48, "����": Exit Sub
If Trim(Text10) = empy Then MsgBox "��������ĺ������", 48, "����": Exit Sub

Dim rs As New ADODb.Recordset
Dim sql As String
Dim temp As String

'��ȷ���Ƿ�������
temp = " where " & Trim(Label1.Caption) & " ='" & Trim(Text1) & "' and " & Trim(Label2.Caption) & " ='" & Trim(Text2) & "' and " & Trim(Label3.Caption) & " ='" & Trim(Text3) & "' and " & Trim(Label4.Caption) & " = #" & Trim(Text4) & "# and " & Trim(Label5.Caption) & " ='" & Trim(Text5) & "' and " & Trim(Label6.Caption) & " ='" & Trim(Text6) & "' and " & Trim(Label7.Caption) & " ='" & Trim(Text7) & "' and " & Trim(Label8.Caption) & " ='" & Trim(Text8) & "' "
'��ȷ����û����һ��
rs.Open "select * from �鼮��Ϣ�����¼�� " & temp, ������Ӳ��ҽ���.conn, 1, 3
    If rs.EOF = True Then MsgBox "û�д������ݣ���ȷ�Ϻ��������룡", 48, "����": Exit Sub
    If CStr(rs.Fields(Combo2.Text).Value) = Trim(Text10) Then MsgBox "����ǰ�͸��ĺ������һ����", 48, "����": Exit Sub
    '�������ǰ������Ҳ��ʱ��������ҲҪУ׼
    If IsDate(CStr(rs.Fields(Combo2.Text).Value)) Then If IsDate(Trim(Text10)) = False Then MsgBox "Ҫ���ĵ�������ʱ�����͵ģ�������ʱ������ eg:2022/3/4", 48, "����": Exit Sub
rs.Close

a = MsgBox("��ȷ��Ҫɾ����ɾ���������޷��һأ������ɾ����", vbYesNo, "����")
If a = vbYes Then
    If Combo2.Text = Combo2.List(4) Then
        sql = "update �鼮��Ϣ�����¼�� set " & Combo2.Text & " = #" & Trim(Text10) & "# " & temp
    Else
        sql = "update �鼮��Ϣ�����¼�� set " & Combo2.Text & " = '" & Trim(Text10) & "' " & temp
    End If

    �鼮��Ϣ�����¼�鿴����.bookhistory_load sql, False
Else
    Exit Sub
End If

MsgBox "���ݸ��ĳɹ�!", , "ע��"
Unload Me
�鼮��Ϣ�����¼�鿴����.Show
�鼮��Ϣ�����¼�鿴����.Enabled = True
�鼮��Ϣ�����¼�鿴����.bookhistory_load "select * from �鼮��Ϣ�����¼�� order by " & �鼮��Ϣ�����¼�鿴����.Combo1.Text, True
End Sub

Private Sub Command2_Click()
Unload Me
�鼮��Ϣ�����¼�鿴����.Show
�鼮��Ϣ�����¼�鿴����.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
�鼮��Ϣ�����¼�鿴����.Show
�鼮��Ϣ�����¼�鿴����.Enabled = True
End Sub

