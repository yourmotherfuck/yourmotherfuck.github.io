VERSION 5.00
Begin VB.Form �����¼���޸����ݽ��� 
   Caption         =   "�����¼���޸����ݽ���"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   12375
   StartUpPosition =   3  '����ȱʡ
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
      Left            =   6960
      TabIndex        =   4
      Top             =   5520
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
      Height          =   735
      Left            =   9360
      TabIndex        =   3
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "��׼���Ҵ���"
      Height          =   5175
      Left            =   480
      TabIndex        =   2
      Top             =   840
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
         TabIndex        =   13
         Top             =   840
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
         Left            =   2760
         TabIndex        =   12
         Top             =   1560
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
         Left            =   2760
         TabIndex        =   11
         Top             =   2280
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
         Left            =   2760
         TabIndex        =   10
         Top             =   3000
         Width           =   1935
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
         TabIndex        =   9
         Top             =   3720
         Width           =   1935
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
         Left            =   720
         TabIndex        =   18
         Top             =   840
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
         Left            =   720
         TabIndex        =   17
         Top             =   1560
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
         Left            =   720
         TabIndex        =   16
         Top             =   2280
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
         Left            =   720
         TabIndex        =   15
         Top             =   3000
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
         Left            =   720
         TabIndex        =   14
         Top             =   3720
         Width           =   1695
      End
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   9240
      TabIndex        =   1
      Top             =   4680
      Width           =   2055
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
      ItemData        =   "�����¼���޸����ݽ���.frx":0000
      Left            =   9120
      List            =   "�����¼���޸����ݽ���.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3960
      Width           =   2655
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
      Left            =   7680
      TabIndex        =   8
      Top             =   840
      Width           =   3015
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
      Left            =   7680
      TabIndex        =   7
      Top             =   2760
      Width           =   3135
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
      Left            =   7080
      TabIndex        =   6
      Top             =   4680
      Width           =   1695
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
      Left            =   7080
      TabIndex        =   5
      Top             =   3960
      Width           =   1695
   End
End
Attribute VB_Name = "�����¼���޸����ݽ���"
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

If Trim(Combo2.Text) = Combo2.List(0) Then MsgBox "��ѡ��Ҫ�������ݵ�������", 48, "����": Exit Sub
If Trim(Combo2.Text) = Combo2.List(4) And IsDate(Trim(Text10)) = False Then MsgBox "���ĺ��������Ϊʱ���ʽ��eg:2022/3/3", 48, "����": Exit Sub
If Trim(Text10) = empy Then MsgBox "��������ĺ������", 48, "����": Exit Sub

'��ȷ���Ƿ�������
Dim temp, sql As String
Dim rs As New ADODb.Recordset
temp = " where " & Trim(Label1.Caption) & " ='" & Trim(Text1) & "' and " & Trim(Label2.Caption) & " ='" & Trim(Text2) & "' and " & Trim(Label3.Caption) & " ='" & Trim(Text3) & "' and " & Trim(Label4.Caption) & " = #" & Trim(Text4) & "# and " & Trim(Label5.Caption) & " =" & Trim(Text5)
'��ȷ����û����һ��
rs.Open "select * from �����¼�� " & temp, ������Ӳ��ҽ���.conn, 1, 3
    If rs.EOF = True Then MsgBox "û�д������ݣ���ȷ�Ϻ��������룡", 48, "����": Exit Sub
    If CStr(rs.Fields(Combo2.Text).Value) = Trim(Text10) Then MsgBox "����ǰ�͸��ĺ������һ����", 48, "����": Exit Sub
rs.Close

a = MsgBox("��ȷ��Ҫɾ����ɾ���������޷��һأ������ɾ����", vbYesNo, "����")
If a = vbYes Then
    If Combo2.Text = Combo2.List(4) Then
        sql = "update �����¼�� set " & Combo2.Text & " = #" & Trim(Text10) & "# " & temp
    ElseIf Combo2.Text = Combo2.List(5) Then
        sql = "update �����¼�� set " & Combo2.Text & " = " & Trim(Text10) & " " & temp
    Else
        sql = "update �����¼�� set " & Combo2.Text & " = '" & Trim(Text10) & "' " & temp
    End If

     �����¼�鿴����.senthistory_load sql, False
Else
    Exit Sub
End If

MsgBox "���ݸ��ĳɹ�!", , "ע��"
Unload Me
�����¼�鿴����.Enabled = True
�����¼�鿴����.senthistory_load "select * from �����¼�� order by " & �����¼�鿴����.Combo1.Text, True
�����¼�鿴����.Show
End Sub

Private Sub Command2_Click()
Unload Me
�����¼�鿴����.Show
�����¼�鿴����.Enabled = True
End Sub

Private Sub Form_Load()
Combo2.Text = Combo2.List(0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
�����¼�鿴����.Show
�����¼�鿴����.Enabled = True
End Sub

