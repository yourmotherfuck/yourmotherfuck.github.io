VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ������Ӳ��ҽ��� 
   Caption         =   "������"
   ClientHeight    =   9705
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   16815
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
   ScaleHeight     =   9705
   ScaleWidth      =   16815
   StartUpPosition =   3  '����ȱʡ
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
      Left            =   15000
      TabIndex        =   10
      Top             =   840
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
      ItemData        =   "������Ӳ��ҽ���.frx":0000
      Left            =   12480
      List            =   "������Ӳ��ҽ���.frx":0019
      TabIndex        =   9
      Text            =   "ѡ��Ҫ���ҵ��ֶ�"
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   8880
      TabIndex        =   8
      Text            =   "���ʱ��"
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "test"
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
      Left            =   10920
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�����鼮"
      Height          =   495
      Left            =   15000
      TabIndex        =   5
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
      Left            =   12480
      TabIndex        =   4
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�޸�����"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�½�����"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ɾ������"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7455
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   13150
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   18
      AllowAddNew     =   -1  'True
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
      Left            =   6360
      TabIndex        =   7
      Top             =   600
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
      Begin VB.Menu �����鼮 
         Caption         =   "�����鼮����ϸ��"
      End
   End
   Begin VB.Menu ����ϵͳ 
      Caption         =   "����ϵͳ"
      Begin VB.Menu ���������¼ 
         Caption         =   "���������¼"
      End
      Begin VB.Menu ���������¼ 
         Caption         =   "���������¼"
      End
      Begin VB.Menu �޸ļ�¼ 
         Caption         =   "�޸ļ�¼�����Ǽ���ʷ��¼��"
      End
   End
   Begin VB.Menu ��ʷ��¼��ѯ 
      Caption         =   "��ʷ��¼��ѯ"
      Begin VB.Menu ��������¼��ѯ 
         Caption         =   "��������¼��ѯ"
      End
      Begin VB.Menu �鼮�����¼��ѯ 
         Caption         =   "�鼮�����¼��ѯ"
      End
   End
End
Attribute VB_Name = "������Ӳ��ҽ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public conn As New ADODB.Connection


Public Sub maindata_load(strsql As String)
            'MsgBox strsql
    Dim nothiscolumn As Boolean
    Dim rs As New ADODB.Recordset
    rs.Open strsql, conn, 1, 3
    For Each i In rs.Fields
        If Text2 = i.Name Then nothiscolumn = True
    Next i
    If nothiscolumn Then
        If rs.EOF Then MsgBox "û����֮ƥ�������", 48, "����": Exit Sub
        Set DataGrid1.DataSource = rs
    Else
        MsgBox "���ݿ���û��������������������", 48, "ע��"
    End If
End Sub


Private Sub Command4_Click()
If Combo1.Text = "ѡ��Ҫ���ҵ��ֶ�" Then MsgBox "��������Ҫ���ҵ��ֶ�": Exit Sub
If Combo1.Text = "���ʱ��" And Trim(Text1.Text) <> Format(Trim(Text1.Text), "yyyy/m/d") Then MsgBox "��������ȷ�����ڸ�ʽ���磺2022/3/3", 48, "ע��": Exit Sub
If Combo1.Text = "���" And Val(Trim(Text1.Text)) <> Trim(Text1.Text) Then MsgBox "����ҿ�������봿����": Exit Sub
If Combo1.Text = "���ʱ��" Then
    maindata_load ("select * from ���ϵͳ where ���ʱ�� = #" & Trim(Text1.Text) & "#")
ElseIf Combo1.Text = "���" Then
    maindata_load (" select * from ���ϵͳ where ��� =  " & Val(Trim(Text1.Text)))
Else
    maindata_load (" select * from ���ϵͳ where " & Trim(Combo1.Text) & " = '" & Trim(Text1.Text) & "'")
End If

End Sub

Private Sub Command5_Click()
maindata_load ("select * from ���ϵͳ where ���ʱ�� = #2020/3/3# ")
End Sub

Private Sub Command6_Click()
maindata_load (" select * from  ���ϵͳ order by " & Trim(Text2) & " desc ")
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    yes = MsgBox("��ȷ��Ҫ����������", vbYesNo, "ע��")
    If yes = vbYes Then
        Dim rs As New ADODB.Recordset
        rs.Open "���ϵͳ", conn, 1, 3
            rs.Update
        rs.Close
        maindata_load (" select * from  ���ϵͳ order by " & Trim(Text2) & " desc ")
        MsgBox "���ĳɹ�"
    End If
End If
End Sub

Private Sub Form_Load()
'�����������ݿ�
conn.Provider = "Microsoft.Jet.OLEDB.4.0"              '����Ĭ�ϲ���
conn.Open "..\value\���.mdb"                             '����ѡ�����ݿ�·��

'������ʲôʲô���Ķ���
conn.CursorLocation = adUseClient
    Dim rs As New ADODB.Recordset
    rs.Open " select * from  ���ϵͳ order by " & Trim(Text2) & " desc ", conn, 1, 3
    Set DataGrid1.DataSource = rs
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
a = MsgBox("ȷ��Ҫ�˳���", vbYesNo, "QAQ")
If a = vbYes Then
    Cancel = 0
Else
    Cancel = 1
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then maindata_load (" select * from  ���ϵͳ order by " & Trim(Text2) & " desc ")
End Sub

Private Sub �����鼮_Click()
��ϸ���ҽ���.Show
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
