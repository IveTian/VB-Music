VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ر�ʱ"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5580
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5580
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CheckBox Ck 
      Caption         =   "��������"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton Op2 
      Caption         =   "��̨����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.OptionButton Op1 
      Caption         =   "�رճ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "�����¹رհ�ť��ϣ������:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a%, b%

Private Sub Command1_Click()
If Ck.Value = 1 Then
  Open App.path & "\End.ini" For Output As #1
   Print Op1.Value
   Print Op2.Value
  Close #1
Else
  If Dir(App.path & "End.ini", vbNormal + vbDirectory) <> "" Then Kill (App.path & "End.ini")
End If
If Op1.Visible = True Then
    Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
   'Shell (App.path & "/Route/ɾ��.bat")
   Call save
   Form1.WindowsMediaPlayer1.Controls.stop 'ֹͣ����
   Form1.WindowsMediaPlayer1.close 'WindowsMediaPlayer1.�ر�
   Set Form1 = Nothing '�ͷŴ������
   End '��������
End If
If Op2.Value = True Then
  Form1.BorderStyle = 1
  Unload Me
End If
End Sub

Private Sub Command2_Click()
Form1.Show
Unload Me
End Sub

Private Sub Form_Load()
Dim tag As String
a = Me.Height
b = Me.Width
If Dir(App.path & "End.ini", vbNormal + vbDirectory) <> "" Then
   Open App.path & "\End.ini" For Input As #1
     Line Input #1, tag
      If tag = "True" Then Op1.Value = True
     Line Input #1, tag
      If tag = "True" Then Op2.Value = True
  Close #1
  Exit Sub
End If
Op2.Visible = True
End Sub

Private Sub Form_Resize()
Me.Height = a
Me.Width = b
End Sub
