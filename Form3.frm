VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ϸ��Ϣ"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9210
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   9210
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer T3 
      Interval        =   70
      Left            =   2760
      Top             =   120
   End
   Begin VB.CommandButton C2 
      Caption         =   "����"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Timer T2 
      Interval        =   70
      Left            =   2040
      Top             =   120
   End
   Begin VB.Timer T1 
      Interval        =   70
      Left            =   1080
      Top             =   120
   End
   Begin VB.CommandButton C1 
      Caption         =   "��ϵ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label S2 
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1335
      Index           =   1
      Left            =   4680
      TabIndex        =   8
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label S1 
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1335
      Index           =   1
      Left            =   4800
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label xian2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "��л"
      BeginProperty Font 
         Name            =   "����"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   3840
      TabIndex        =   6
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Line xian1 
      BorderStyle     =   4  'Dash-Dot
      X1              =   0
      X2              =   9240
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label S2 
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1335
      Index           =   0
      Left            =   3120
      TabIndex        =   4
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label S1 
      BackStyle       =   0  'Transparent
      Caption         =   "�s"
      BeginProperty Font 
         Name            =   "����"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1335
      Index           =   0
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label L2 
      BackStyle       =   0  'Transparent
      Caption         =   "������Ա��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label L1 
      BackStyle       =   0  'Transparent
      Caption         =   "���߼��༭:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim S%, l1t%, l2t%, xian%, kuan%, gao%, an%
'***********�ƶ�����API**********************
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Const WM_NCLBUTTONDOWN = &HA1 '���峣�� WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2 '���峣�� HTCAPTION = 2

Private Sub C1_Click()
MsgBox "����QQ:1273367387  ��ɼ�Ⱥ251924436", vbOK, "��ϵ��"
End Sub

Private Sub C2_Click()
MsgBox "�ٳ���ʦ ��һЩ�����������Ͼ��������Ұ������ش���ָ�л", vbYes, "����"
End Sub

Private Sub Form_Load()
kuan = Form3.Width
gao = Form3.Height
S = S1(0).Width
l1t = L1.Top
l2t = L2.Top
xian = xian1.Y1
an = C1.Left
'-----------------------------
L1.Top = -300
L2.Top = gao + 300

For i = 0 To 1
  S2(i).Width = 0
  S1(i).Width = 0
Next i

xian1.Y1 = gao + 300
xian1.Y2 = gao + 300
xian2.Top = gao + 300
C1.Left = kuan + 300
C2.Left = kuan + 300
End Sub
Private Sub Form_Resize()
If Form3.Height > gao Then Form3.Height = gao
If Form3.Width > kuan Then Form3.Width = kuan
End Sub
'��갴����������϶�����
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = 1 Then '��� Button = 1 ��갴���������
      Call ReleaseCapture '���� ReleaseCapture
      Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&) '��ʼ�϶�����
   End If
End Sub




Private Sub T1_Timer()

If L1.Top < l1t Then
  L1.Top = L1.Top + 100
  Exit Sub
Else
  L1.Top = l1t
End If

If S1(0).Width < S Then
  S1(0).Width = S1(0).Width + 100
  Exit Sub
Else
  S1(0).Width = S
End If

If S1(1).Width < S Then
  S1(1).Width = S1(1).Width + 100
  Exit Sub
Else
  S1(1).Width = S
End If

If C1.Left > an Then
   C1.Left = C1.Left - 100
   Exit Sub
Else
   C1.Left = an
End If
T1.Enabled = False
End Sub

Private Sub T2_Timer()
Dim y%
If T1.Enabled = True Then Exit Sub
If xian1.Y1 > xian Then
  xian1.Y1 = xian1.Y1 - 100
  y = xian1.Y1
  xian1.Y2 = y
  xian2.Top = y
  Exit Sub
Else
  xian1.Y1 = xian
  y = xian1.Y1
  xian1.Y2 = y
  xian2.Top = y
End If
T2.Enabled = False
End Sub

Private Sub T3_Timer()

If T2.Enabled = True Then Exit Sub

If L2.Top > l2t Then
  L2.Top = L2.Top - 100
  Exit Sub
Else
  L2.Top = l2t
End If

If S2(0).Width < S Then
  S2(0).Width = S2(0).Width + 100
  Exit Sub
Else
  S2(0).Width = S
End If

If S2(1).Width < S Then
  S2(1).Width = S2(1).Width + 100
  Exit Sub
Else
  S2(1).Width = S
End If

If C2.Left > an Then
   C2.Left = C2.Left - 100
   Exit Sub
Else
   C2.Left = an
End If

T3.Enabled = False
End Sub
