VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Music(�s������)"
   ClientHeight    =   10470
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   8835
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10470
   ScaleWidth      =   8835
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton SS 
      BackColor       =   &H8000000D&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   120
      Width           =   375
   End
   Begin VB.Timer Tge 
      Interval        =   500
      Left            =   3720
      Top             =   7680
   End
   Begin VB.CommandButton gedan 
      Caption         =   "�������赥"
      Height          =   375
      Index           =   2
      Left            =   3720
      TabIndex        =   55
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton gedan 
      Caption         =   "ɾ���赥"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   54
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton gedan 
      Caption         =   "��Ӹ赥"
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   53
      Top             =   5040
      Width           =   1095
   End
   Begin VB.ComboBox Cb 
      Height          =   300
      ItemData        =   "Form1.frx":324A
      Left            =   2280
      List            =   "Form1.frx":3254
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Timer ��һ�� 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8160
      Top             =   6000
   End
   Begin VB.CommandButton C 
      Caption         =   "���Ϊ��ʼ����"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   5760
      TabIndex        =   51
      Top             =   9240
      Width           =   1455
   End
   Begin VB.TextBox ��һ�� 
      Height          =   270
      Left            =   2400
      TabIndex        =   50
      Text            =   "������һ����Ϣ"
      Top             =   8880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer ͸���� 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   9840
   End
   Begin VB.TextBox touming 
      Height          =   270
      Left            =   0
      TabIndex        =   49
      Text            =   "255"
      Top             =   9480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer �ö� 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7440
      Top             =   6120
   End
   Begin VB.Timer suiji2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   5640
   End
   Begin VB.Timer suiji 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   5640
   End
   Begin VB.CommandButton C 
      Caption         =   "�ز�"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   4080
      TabIndex        =   48
      Top             =   9240
      Width           =   1455
   End
   Begin VB.CommandButton Command 
      Caption         =   "���к�"
      Height          =   375
      Index           =   35
      Left            =   5880
      TabIndex        =   47
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton C 
      Caption         =   "�ر�����"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   720
      TabIndex        =   46
      Top             =   9240
      Width           =   1455
   End
   Begin VB.CommandButton C 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   4
      Left            =   7320
      Picture         =   "Form1.frx":3269
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   9000
      Width           =   975
   End
   Begin VB.CommandButton caozuo 
      Caption         =   "��ѡ����ĿĿ¼"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Index           =   1
      Left            =   6720
      MaskColor       =   &H8000000F&
      TabIndex        =   44
      TabStop         =   0   'False
      ToolTipText     =   "����ѡ�����Ŀ����Ŀ¼"
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton caozuo 
      Caption         =   "����ָ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Index           =   2
      Left            =   5760
      MaskColor       =   &H8000000F&
      TabIndex        =   43
      TabStop         =   0   'False
      ToolTipText     =   "����ѡ��ĸ���"
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   7800
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "ѡ������ļ�"
      Filter          =   "MP3(*.mp3)|*.mp3|MIDI(*.midi)|*.midi|CD Audio(*.cda)|*.cda|WAV(*.wav)|*.wav|WMA(*.wma)|*.wma|����(*.*)|*.*|"
   End
   Begin VB.CommandButton caozuo 
      Caption         =   "ɾ��/���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Index           =   3
      Left            =   5040
      TabIndex        =   42
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton caozuo 
      Caption         =   "����赥�Լ���Ŀ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Index           =   4
      Left            =   6720
      TabIndex        =   40
      ToolTipText     =   "���������ӳ��ĸ������´γ�������ʱ�Զ����"
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton caozuo 
      Caption         =   "��Ӹ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Index           =   0
      Left            =   5040
      MaskColor       =   &H8000000F&
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   "�����Լ��ĸ�����ӵ�����б�"
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   3300
      ItemData        =   "Form1.frx":4F33
      Left            =   120
      List            =   "Form1.frx":4F35
      MultiSelect     =   1  'Simple
      TabIndex        =   37
      Top             =   5520
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   8880
   End
   Begin VB.CommandButton Command 
      Caption         =   "���Ƕ�һ��"
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
      Index           =   29
      Left            =   6720
      TabIndex        =   36
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command 
      Caption         =   "��ɶ������"
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
      Index           =   24
      Left            =   6720
      TabIndex        =   35
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command 
      Caption         =   "���ܰ��ֵ�"
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
      Index           =   28
      Left            =   5040
      TabIndex        =   34
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command 
      Caption         =   "���Ĺ���"
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
      Index           =   23
      Left            =   5040
      TabIndex        =   33
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command 
      Caption         =   "ʢ�ĵĹ�ʵ"
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
      Index           =   27
      Left            =   3360
      TabIndex        =   32
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command 
      Caption         =   "���ڵȵ���"
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
      Index           =   22
      Left            =   3360
      TabIndex        =   31
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command 
      Caption         =   "û��ô��"
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
      Index           =   26
      Left            =   1680
      TabIndex        =   30
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command 
      Caption         =   "˭ Control"
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
      Index           =   21
      Left            =   1680
      TabIndex        =   29
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command 
      Caption         =   "ͻȻ������"
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
      Index           =   25
      Left            =   0
      TabIndex        =   28
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command 
      Caption         =   "Miss Mystery"
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
      Index           =   20
      Left            =   0
      TabIndex        =   27
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command 
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
      Index           =   19
      Left            =   6840
      TabIndex        =   26
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      Caption         =   "aLIEz"
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
      Index           =   18
      Left            =   5160
      TabIndex        =   25
      ToolTipText     =   "�˱�����"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      Caption         =   "������ɢ"
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
      Index           =   17
      Left            =   3360
      TabIndex        =   24
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      Caption         =   "���鳵վ"
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
      Index           =   16
      Left            =   1680
      TabIndex        =   23
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      Caption         =   "���˵ĺ�"
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
      Index           =   15
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      Caption         =   "ÿ���˶�������ϣ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   34
      Left            =   6720
      TabIndex        =   21
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command 
      Caption         =   "��һ���˺���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   33
      Left            =   5040
      TabIndex        =   20
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command 
      Caption         =   "K��֮��"
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
      Index           =   14
      Left            =   6840
      TabIndex        =   19
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      Caption         =   "�������"
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
      Index           =   13
      Left            =   5160
      TabIndex        =   18
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command 
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
      Index           =   9
      Left            =   6840
      TabIndex        =   17
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command 
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
      Index           =   8
      Left            =   5160
      TabIndex        =   16
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command 
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
      Index           =   4
      Left            =   6840
      TabIndex        =   15
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      Caption         =   "����ҹ��"
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
      Left            =   5160
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton C 
      Caption         =   "��ͣ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   2400
      TabIndex        =   12
      Top             =   9240
      Width           =   1455
   End
   Begin VB.CommandButton Command 
      Caption         =   "��Ȼ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   32
      Left            =   3360
      TabIndex        =   11
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command 
      Caption         =   "��Ѱ������˭"
      Height          =   615
      Index           =   31
      Left            =   1680
      TabIndex        =   10
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command 
      Caption         =   "ʧ�������Ĳ���"
      Height          =   615
      Index           =   30
      Left            =   0
      TabIndex        =   9
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command 
      Caption         =   "��İ���"
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
      Index           =   12
      Left            =   3360
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      Caption         =   "����һ��"
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
      Index           =   11
      Left            =   1680
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command 
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
      Index           =   10
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command 
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
      Index           =   7
      Left            =   3360
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      Caption         =   "��"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      Caption         =   "��"
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
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command 
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
      Index           =   2
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command 
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
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command 
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label ��ʾ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��ʾ :������Ӻ�������ĸ�������Ŀ¼���޸ĺ����������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3600
      TabIndex        =   41
      Top             =   8160
      Width           =   4695
   End
   Begin VB.Label L1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�û������Ŀ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   600
      TabIndex        =   38
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   8880
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8760
      Y1              =   8880
      Y2              =   8880
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   873
      _cy             =   450
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu ���ڲ���ʲô 
         Caption         =   "���ڲ���ʲô"
         Shortcut        =   {F1}
      End
      Begin VB.Menu ����Ŀ¼ 
         Caption         =   "�򿪳���Ŀ¼"
      End
      Begin VB.Menu close 
         Caption         =   "ɾ��/���"
      End
      Begin VB.Menu SaveMusic 
         Caption         =   "�����������"
         Shortcut        =   ^D
      End
      Begin VB.Menu Save���� 
         Caption         =   "��������"
         Shortcut        =   ^S
      End
      Begin VB.Menu end 
         Caption         =   "�˳�"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu MenuTray 
      Caption         =   "����"
      Begin VB.Menu shower 
         Caption         =   "��ʾ"
      End
      Begin VB.Menu end2 
         Caption         =   "�رճ���"
      End
      Begin VB.Menu wu3 
         Caption         =   "-"
      End
      Begin VB.Menu Ctrl 
         Caption         =   "��һ��"
         Index           =   0
      End
      Begin VB.Menu Ctrl 
         Caption         =   "��һ��"
         Index           =   1
      End
      Begin VB.Menu Ctrl 
         Caption         =   "��ͣ"
         Index           =   2
      End
      Begin VB.Menu Ctrl 
         Caption         =   "����"
         Index           =   3
      End
      Begin VB.Menu Ctrl 
         Caption         =   "�ز�"
         Index           =   4
      End
      Begin VB.Menu Ctrl 
         Caption         =   "ֹͣ"
         Index           =   5
      End
   End
   Begin VB.Menu ��Ϣ 
      Caption         =   "��Ϣ"
      Begin VB.Menu ��ϸ��Ϣ 
         Caption         =   "��ϸ��Ϣ"
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu open 
         Caption         =   "������"
      End
      Begin VB.Menu �� 
         Caption         =   "-"
      End
      Begin VB.Menu ���� 
         Caption         =   "���屳��ɫ"
      End
   End
   Begin VB.Menu Skin 
      Caption         =   "Ƥ��"
      Begin VB.Menu SkinEx 
         Caption         =   "��ɫ����"
         Index           =   0
      End
      Begin VB.Menu SkinEx 
         Caption         =   "��ľ���"
         Index           =   1
      End
      Begin VB.Menu SkinEx 
         Caption         =   "��ɫ����"
         Index           =   2
      End
      Begin VB.Menu SkinEx 
         Caption         =   "�����ۺ�"
         Index           =   3
      End
      Begin VB.Menu SkinEx 
         Caption         =   "��ɫ�ഺ"
         Index           =   4
      End
      Begin VB.Menu SkinEx 
         Caption         =   "��ɫ����"
         Index           =   5
      End
      Begin VB.Menu SkinEx 
         Caption         =   "�й���"
         Index           =   6
      End
      Begin VB.Menu SkinEx 
         Caption         =   "��ɫ����"
         Index           =   7
      End
      Begin VB.Menu wu 
         Caption         =   "-"
      End
      Begin VB.Menu �Զ���Ƥ�� 
         Caption         =   "ʹ���Զ���Ƥ��"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim houzhui As String, suoyin%, sui%, l%, TS%, bofang As String, chongbo As String, Play As Boolean, mulu As String, biaoji As String, youji As Boolean, gaibian As Boolean
'**********�����ö���API**************
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'***********����͸����API************************
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Const WS_EX_LAYERED = &H80000 '���峣�� WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20) '���峣�� GWL_EXSTYLE =(-20)
Const LWA_ALPHA = &H2 '���峣�� LWA_ALPHA = &H2
Const LWA_COLORKEY = &H1 '���峣�� LWA_COLORKEY = &H1
'***********�ƶ�����API**********************
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Const WM_NCLBUTTONDOWN = &HA1 '���峣�� WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2 '���峣�� HTCAPTION = 2

Private Sub C_Click(Index As Integer)
'4��������ť

Select Case Index
'�ر�����***************************
Case Is = 0
a = MsgBox("�����Ҫ�ر����ڲ��ŵ�������", vbYesNo, "Music")
If a = vbNo Then Exit Sub
chongbo = "F"
WindowsMediaPlayer1.Controls.stop
bofang = "����ʲô��û��Ŷ����ȥѡ����Ҫ���ĸ����ɣ�"
C2.Caption = "��ͣ"
C2.Enabled = False
C3.Enabled = False
C1.Enabled = False
'��ͣ*********************************
Case Is = 1
If C(1).Caption = "��ͣ" And WindowsMediaPlayer1.URL <> "" Then
     chongbo = "F"
     WindowsMediaPlayer1.Controls.pause '��ͣ����
     C(1).Caption = "����"
Else
     chongbo = "T"
     C(1).Caption = "��ͣ"
     WindowsMediaPlayer1.Controls.Play
End If
'�ز�************************************
Case Is = 2
C(1).Caption = "��ͣ"
WindowsMediaPlayer1.Controls.pause
chongbo = "T"
'���Ϊ��ʼ����******************************
Case Is = 3
If C(4).Caption = "���Ϊ��ʼ����" Then
   For i = 0 To Command.Count - 1
     Command(i).Font.Bold = False
   Next i
   biaoji = bofang
   kaishi = WindowsMediaPlayer1.URL
   For i = 0 To Command.Count - 1
     If Command(i).Caption = ��һ��.Text Then Command(i).Font.Bold = True
   Next i
   MsgBox "��ǳɹ���", vbYes, "Music"
Else
   kaishi = AppDisk & "aLTEz" & houzhui
End If
'����****************************
Case Is = 4
Form2.Show
'����**************************
End Select


End Sub

Private Sub caozuo_Click(Index As Integer)
Select Case Index
Case Is = 0   '���

Dim FullName As String, ShortName As String, i As Long, chongfu&
CD1.InitDir = App.path
chongfu = 1
CD1.Action = 1
On Error GoTo ErrHandler
If CD1.FileName = "" Then Exit Sub
FullName = CD1.FileName
ShortName = Right(FullName, Len(FullName) - InStrRev(FullName, "\"))
For i = List1.ListCount - 1 To 0 Step -1
    If List1.List(i) = ShortName + "(" & chongfu & ")" Or List1.List(i) = ShortName Then
        chongfu = chongfu + 1
    End If
Next i

If chongfu > 1 Then
      a = MsgBox("����������ӹ����Ƿ������", vbYesNo, "Music")
        If a = vbYes Then ShortName = ShortName + "(" & chongfu & ")" Else Exit Sub
End If

List1.AddItem ShortName, suoyin
suoyin = suoyin + 1
   If Dir(App.path & "/Route/" & ShortName & ".txt", vbNormal + vbDirectory) <> "" Then Kill (App.path & "/Route/" & ShortName & ".txt")
     Open App.path & "/Route/" & ShortName & ".txt" For Output As #1
     Print #1, FullName
     Close #1
CD1.FileName = ""
WindowsMediaPlayer1.URL = FullName
ErrHandler:
                '�û����ˡ�ȡ������ť��
                Exit Sub
'----------------------------------------
Case Is = 1  'Ŀ¼
��ʾ.Caption = "��ʾ:��ѡ�������Ĭ�ϴ����һ��Ŀ¼,��δѡ����򿪳���Ŀ¼"
mulu = App.path + "\"
Shell "explorer.exe mulu", 1
'------------------------------------------
Case Is = 2   '����
Dim S As String
��ʾ.Caption = "��ʾ:��ѡ�������Ĭ�ϲ��ŵ�һ��"

   For i = List1.ListCount - 1 To 0 Step -1
     If List1.Selected(i) = True Then S = List1.List(i)
   Next
   If S = "" Then
      MsgBox "����δѡ�����", vbYes, "����"
      Exit Sub
   End If
If Dir(App.path & "/Route/" & S & ".txt", vbNormal + vbDirectory) <> "" Then
   Open App.path & "/Route/" & S & ".txt" For Input As #2
   Do While Not EOF(2)
   InputData = Input(1, #2)
   mulu = mulu + InputData
   Loop
   Close #2
Else
   MsgBox "����δ֪ԭ�򣬳��ִ��󣬸ø������������", vbOK, "����"
   Exit Sub
End If
   Play = False
   ��һ��.Text = WindowsMediaPlayer1.URL
   WindowsMediaPlayer1.URL = mulu
   chongbo = "T"
   mulu = ""
   bongfang = S
   C4.Enabled = False
   C2.Enabled = True
   C3.Enabled = True
   C1.Enabled = True
'---------------------------------------------
Case Is = 3   'ɾ��

a = MsgBox("[ɾ��ѡ������] ��ѡȷ����[����б�] ��ѡȡ��", vbYesNo, "ɾ�������")
If a = vbNo Then b = MsgBox("��ȷ������б���", vbYesNo, "��ʾ")
If a = vbYes Then
   For i = List1.ListCount - 1 To 0 Step -1
     suoyin = suoyin - 1
     If List1.Selected(i) = True Then
         Kill (App.path & "\Route\" & List1.List(i) & ".txt")
         List1.RemoveItem (i)
     End If
   Next i
   MsgBox "ɾ��ָ�������ɹ���", vbYes, "Music"
Else
    If b = vbYes Then List1.Clear
        suoyin = 0
        For i = List1.ListCount - 1 To 0 Step -1
           If List1.Selected(i) = True Then List1.RemoveItem (i)
           Kill (App.path & "\Route\" & List1.List(i) & ".txt")
        Next i
        Open App.path & "\SaveMusic.ini" For Output As #1
        Close #1
        MsgBox "����б�ɹ���", vbYes, "Music"
End If

'-----------------------------------------------
Case Is = 4  '����

Dim n As String
Open App.path & "\SaveMusic.ini" For Output As #1
For i = 0 To List1.ListCount - 1
   n = List1.List(i)
   Print #1, n '������Ŀ
Next i
Close #1
MsgBox "[�û������Ŀ]����ɹ���", vbYes, "��ʾ"
'-----------------------------------------------
End Select
End Sub





Private Sub Command_Click(Index As Integer)
    ��һ��.Text = WindowsMediaPlayer1.URL
    WindowsMediaPlayer1.URL = AppDisk & Command(Index).Caption & houzhui
    Play = True
    If ��һ��.Text = "" Then ��һ��.Text = Command(Index).Caption
    bofang = Command(Index).Caption
    If Command(Index).Caption = biao Then C(3).Caption = "ȡ����Ϊ��ʼ����" Else C(3).Caption = "���Ϊ��ʼ����"
chongbo = "T"
For i = 0 To 3
C(i).Enabled = True
Next i
End Sub

Private Sub Command1_Click()
MsgBox (Cb.Text)
End Sub

Private Sub Ctrl_Click(Index As Integer)
On Error Resume Next

Select Case Index

 Case 0
   If ��һ��.Text = "" Then Exit Sub
    If Dir(��һ��.Text, vbNormal + vbDirectory) <> "" Then
         WindowsMediaPlayer1.URL = ��һ��.Text
     Else
        MsgBox "����Դ�ļ��ƶ���ɾ�����޷�������һ�ף������¶�λ�����ļ�", vbOK, "����"
        Exit Sub
     End If
   chongbo = "T"
 Case 1
    If Form2.S1 = "ѭ������(Ĭ��)" And Form2.Ck2.Value = 1 Then
       MsgBox "�㵱ǰģʽΪ[ѭ������]������[һ������ѭ��״̬],�޷�������һ��", vbYes, "����"
       Exit Sub
    End If
    If bofang = "" Then
       MsgBox "�㻹û�п�ʼ����", vbOK, "����"
       Exit Sub
    End If
    ��һ��.Enabled = True
 Case 2
    WindowsMediaPlayer1.Controls.pause
    C2.Caption = "����"
 Case 3
    WindowsMediaPlayer1.Controls.Play
    C2.Caption = "��ͣ"
 Case 4
    C2.Caption = "��ͣ"
    WindowsMediaPlayer1.Controls.pause
    chongbo = "T"
    
 Case 5
   a = MsgBox("�����Ҫ�ر����ڲ��ŵ�������", vbYesNo, "Music")
   If a = vbNo Then Exit Sub
   chongbo = "F"
   WindowsMediaPlayer1.Controls.stop
   bofang = "����ʲô��û��Ŷ����ȥѡ����Ҫ���ĸ����ɣ�"
   C2.Caption = "��ͣ"
   C2.Enabled = False
   C3.Enabled = False
   C1.Enabled = False
 End Select

End Sub

Private Sub End_Click()
    Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
   'Shell (App.path & "/Route/ɾ��.bat")
   Call save
End
End Sub

Private Sub end2_Click()
    Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
   'Shell (App.path & "/Route/ɾ��.bat")
   Call save
End
End Sub

Private Sub Form_Load()
With nfIconData
.hWnd = Me.hWnd
.uID = Me.Icon
.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
.uCallbackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon.Handle
'��������ƶ���������ʱ��ʾ��Tip
.szTip = "(Music v2.21 BY:�s��)" & vbNullChar
.cbSize = Len(nfIconData)
End With
Call Shell_NotifyIcon(NIM_ADD, nfIconData)

suoyin2 = 1
Cb.Text = Cb.List(0)
shower.Visible = False
end2.Visible = False
wu3.Visible = False
pifu = False
��һ��.Text = ""
chongbo = "T"
houzhui = ".mp3"
kaishi = AppDisk & "��Ȼ����" & houzhui
biaoji = ""
saves = False
w = Me.Width
h = Me.Height
Call ��ʼ��
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 1 Then
   youji = False
ElseIf Button = 2 Then
   youji = True
Else
   youji = False
End If
End Sub

'��갴����������϶�����
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = 1 Then '��� Button = 1 ��갴���������
      Call ReleaseCapture '���� ReleaseCapture
      Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&) '��ʼ�϶�����
   End If
   
   Dim lMsg As Single
    lMsg = X / Screen.TwipsPerPixelX
   Select Case lMsg
    Case WM_LBUTTONUP
    '�����������ʾ����
     shower.Visible = False
     end2.Visible = False
     wu3.Visible = False
     If GetForegroundWindow <> Me.hWnd Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
     'ShowWindow Me.hWnd, SW_RESTORE
   Case WM_RBUTTONUP
      shower.Visible = True
      end2.Visible = True
      wu3.Visible = True
      PopupMenu MenuTray '�������ϵͳTrayͼ���ϵ��Ҽ����򵯳��˵�MenuTray
   Case WM_MOUSEMOVE
   Case WM_LBUTTONDOWN
   Case WM_LBUTTONDBLCLK
   Case WM_RBUTTONDOWN
   Case WM_RBUTTONDBLCLK
   Case Else
  End Select
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If youji = True Then PopupMenu ����
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
   'Shell (App.path & "/Route/ɾ��.bat")
   Call save
   WindowsMediaPlayer1.Controls.stop 'ֹͣ����
   WindowsMediaPlayer1.close 'WindowsMediaPlayer1.�ر�
   Set Form1 = Nothing '�ͷŴ������
   End '��������
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim tag As String
Cancel = 1
If Dir(App.path & "End.ini", vbNormal + vbDirectory) <> "" Then
   Open App.path & "\End.ini" For Input As #1
     Line Input #1, tag
      If tag = "True" Then End
     Line Input #1, tag
       If tag = "True" Then Form1.BorderStyle = 1
   Close #1
   Exit Sub
End If
Call save
Form4.Show
Me.Hide
End Sub


Private Sub gedan_Click(Index As Integer)
Dim p As String
Select Case Index
Case 0
   p = InputBox("����赥��", "�½�")
   If p = "" Then Exit Sub
   For i = 0 To suoyin2
    If Cb.List(i) = p Then
      MsgBox "�赥���ظ�", vbYes, "����"
      Exit Sub
    End If
   Next i
   suoyin2 = suoyin2 + 1
   Cb.AddItem p, suoyin2
   
Case 1
   p = Cb.ListIndex
   a = MsgBox("�����Ҫɾ���赥[" & Cb.List(p) & "]��", vbYesNo, "Music")
   If a = vbYes Then Cb.RemoveItem p
   Cb.Text = Cb.List(0)
   suoyin2 = suoyin2 - 1
Case 2
   p = Cb.ListIndex
   a = InputBox("��Ϊ�赥[" & Cb.List(p) & "]����һ���µ�����", "������")
   If a = "" Then Exit Sub
   Cb.List(p) = a
End Select
End Sub

Private Sub List1_DblClick()
Dim S As String, tag As String
sb = List1.ListIndex '��0����ĵڼ���
S = List1.List(sb)
If Dir(App.path & "/Route/" & S & ".txt", vbNormal + vbDirectory) <> "" Then
   Open App.path & "/Route/" & S & ".txt" For Input As #1
   Line Input #1, tag
   WindowsMediaPlayer1.URL = tag
   Close #1
Else
   MsgBox "��ʷ�ļ�δ�ҵ�������ʷ�޷�����", vbOK, "����"
   Exit Sub
End If
End Sub

Private Sub open_Click()
Form2.Show
End Sub


Private Sub SaveMusic_Click()
Dim i%, n As String
Open App.path & "\SaveMusic.ini" For Output As #1
For i = 0 To List1.ListCount - 1
   n = List1.List(i)
   Print #1, n '������Ŀ
Next i
Close #1
MsgBox "[�û������Ŀ]����ɹ���", vbYes, "��ʾ"
End Sub

Private Sub Save����_Click()
Call ��������
End Sub

Private Sub shower_Click()
Form1.Show
End Sub

Private Sub SkinEx_Click(Index As Integer)
pifu = False
SSkin = SkinEx(Index).Caption
path = App.path & "\Skin\" & SSkin & ".she"
Call Ƥ��
End Sub

Private Sub SS_Click()
If SS.Caption = "��" Then
   SS.Caption = "��"
   suofang = True
   For i = 0 To 34
     Command(i).Visible = False
   Next i
   L1.Top = L1.Top - 4920
   Cb.Top = L1.Top
   For i = 0 To 2
      gedan(i).Top = gedan(i).Top - 4920
   Next i
   List1.Top = List1.Top - 4920
   For i = 0 To 4
      caozuo(i).Top = caozuo(i).Top - 4920
   Next i
   For i = 0 To 4
     C(i).Top = C(i).Top - 4920
   Next i
   ��ʾ.Top = ��ʾ.Top - 4920
   Line2.Visible = False
   Line1.Y1 = Line1.Y1 - 4920
   Line1.Y2 = Line1.Y1
   Form1.Height = Form1.Height - 4920
Else
   SS.Caption = "��"
   suofang = False
   For i = 0 To 34
     Command(i).Visible = True
   Next i
      L1.Top = L1.Top + 4920
   Cb.Top = L1.Top
   For i = 0 To 2
      gedan(i).Top = gedan(i).Top + 4920
   Next i
   List1.Top = List1.Top + 4920
   For i = 0 To 4
      caozuo(i).Top = caozuo(i).Top + 4920
   Next i
   For i = 0 To 4
     C(i).Top = C(i).Top + 4920
   Next i
   ��ʾ.Top = ��ʾ.Top + 4920
   Line2.Visible = True
   Line1.Y1 = Line1.Y1 + 4920
   Line1.Y2 = Line1.Y1
   Form1.Height = Form1.Height + 4920
End If
End Sub

Private Sub suiji_Timer()
Dim mulu As String
If chongbo = "F" Then Exit Sub
If WindowsMediaPlayer1.playState = wmppsStopped Or WindowsMediaPlayer1.playState = wmppsStopped Then
       Randomize
    If Form2.S2.Text = "�����Դ�" Then
       chongbo = "F"
       WindowsMediaPlayer1.Controls.pause
       sui = Int(Rnd * 36)
       WindowsMediaPlayer1.URL = App.path & Command(sui).Caption & houzhui
       bofang = Command(sui).Caption & houzhui
       TS = 0
    ElseIf Form2.S2.Text = "�û����" Then
         If List1.ListCount = 0 And TS <> 1 Then
             MsgBox "�㻹û������κθ�����", vbOK, "��ѡ��������Χ��[�û����]"
             TS = 1
         Else
             l = List1.ListCount + 1
             sui = Int(Rnd * l)
                Open App.path & "/Route/" & List1.List(sui) & ".txt" For Input As #3
                 Do While Not EOF(3)
                 InputData = Input(1, #3)
                 mulu = mulu + InputData
                 Loop
                Close #3
            WindowsMediaPlayer1.URL = mulu
            bofang = List1.List(sui)
         End If
    ElseIf Form2.S2.Text = "ȫ��" Then
        TS = 36 + List1.ListCount
        sui = Int(Rnd * TS)
          If sui <= 35 Then
             WindowsMediaPlayer1.URL = App.path & Command(sui).Caption & houzhui
             bofang = Command(sui).Caption & houzhui
          Else
             If List1.ListCount = 0 Then
                sui = Int(Rnd * 36)
                WindowsMediaPlayer1.URL = App.path & Command(sui).Caption & houzhui
                bofang = Command(sui).Caption & houzhui
             Else
                sui = sui - 36
                Open App.path & "/Route/" & List1.List(sui) & ".txt" For Input As #3
                 Do While Not EOF(3)
                 InputData = Input(1, #3)
                 mulu = mulu + InputData
                 Loop
                Close #3
                WindowsMediaPlayer1.URL = mulu
                bofang = List1.List(sui)
            End If
         End If
    End If
End If '���� ���
End Sub

Private Sub suiji2_Timer()
Dim mulu As String, shi%
Static ci%
If chongbo = "F" Then Exit Sub
shi = Form2.T1.Text
If WindowsMediaPlayer1.playState = wmppsStopped Or WindowsMediaPlayer1.playState = wmppsStopped Then
    If ci < Val(shi) Then
       WindowsMediaPlayer1.Controls.Play
       ci = ci + 1
       Exit Sub
    End If
    ci = 0
     Randomize
    If Form2.S2.Text = "�����Դ�" Then
       chongbo = "F"
       WindowsMediaPlayer1.Controls.pause
       sui = Int(Rnd * 36)
       WindowsMediaPlayer1.URL = App.path & Command(sui).Caption & houzhui
       bofang = Command(sui).Caption & houzhui
       TS = 0
    ElseIf Form2.S2.Text = "�û����" Then
         If List1.ListCount = 0 And TS <> 1 Then
             MsgBox "�㻹û������κθ�����", vbOK, "��ѡ��������Χ��[�û����]"
             TS = 1
         Else
             l = List1.ListCount + 1
             sui = Int(Rnd * l)
                Open App.path & "/Route/" & List1.List(sui) & ".txt" For Input As #3
                 Do While Not EOF(3)
                 InputData = Input(1, #3)
                 mulu = mulu + InputData
                 Loop
                Close #3
            WindowsMediaPlayer1.URL = mulu
            bofang = List1.List(sui)
         End If
    ElseIf Form2.S2.Text = "ȫ��" Then
        TS = 36 + List1.ListCount
        sui = Int(Rnd * TS)
          If sui <= 35 Then
             WindowsMediaPlayer1.URL = App.path & Command(sui).Caption & houzhui
             bofang = Command(sui).Caption & houzhui
          Else
             If List1.ListCount = 0 Then
                sui = Int(Rnd * 36)
                WindowsMediaPlayer1.URL = App.path & Command(sui).Caption & houzhui
                bofang = Command(sui).Caption & houzhui
             Else
                sui = sui - 36
                Open App.path & "/Route/" & List1.List(sui) & ".txt" For Input As #3
                 Do While Not EOF(3)
                 InputData = Input(1, #3)
                 mulu = mulu + InputData
                 Loop
                Close #3
                WindowsMediaPlayer1.URL = mulu
                bofang = List1.List(sui)
            End If
         End If
    End If
End If '���� ���
End Sub

Private Sub Tge_Timer()
If Dir(App.path & "�赥.ini", vbNormal + vbDirectory) <> "" Then Exit Sub
   Open App.path & "/Route/" & ShortName & ".txt" For Output As #1
    Print #1, FullName
   Close #1
End Sub

Private Sub Timer1_Timer()
If WindowsMediaPlayer1.playState = wmppsStopped Or WindowsMediaPlayer1.playState = wmppsReady Then
   If chongbo = "T" Then WindowsMediaPlayer1.Controls.Play   '��ʼѭ���ٴβ���
End If '���� ���
End Sub




Private Sub ����_Click()
a = MsgBox("����ָ� ����Ĭ����ɫ ���� �Զ����µ���ɫ ��������,ǰ��ѡ[ȷ��],����ѡ[ȡ��]", vbYesNo, "����ɫ")
If a = vbYes Then
  Form1.BackColor = &HC0C0C0
  Exit Sub
End If
'�� Cancel ���ó� True��
                CD1.CancelError = True
                On Error GoTo ErrHandler
                '���� Flags ���ԡ�
                CD1.Flags = cdlCCRGBInit
                '��ʾ����ɫ���Ի���
                CD1.Action = 3
                '������ı�����ɫ���ó�ѡ����'��ɫ��
                Form1.BackColor = CD1.Color
                Exit Sub

ErrHandler:
                '�û����ˡ�ȡ������ť��
                Exit Sub
End Sub


Private Sub ͸����_Timer()
Dim S%
S = Val(touming.Text)
    Ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, Ret
    SetLayeredWindowAttributes Form1.hWnd, 0, S, LWA_ALPHA
    ͸����.Enabled = False
End Sub

Private Sub ��һ��_Timer()
Dim music%, S As String
If Form2.S1.Text = "�������" Then
 Randomize
  If Form2.S2.Text = "����" Then
    If List1.ListCount = 0 Then
       music = Command.Count
       music = Int(Rnd * music)
       Call Command_Click(music)
    Else
       music = Command.Count + List1.ListCount
       music = Int(Rnd * music)
       If music <= Command.Count - 1 Then
          Call Command_Click(music)
       Else
          S = List1.List(music)
          If Dir(App.path & "/Route/" & S & ".txt", vbNormal + vbDirectory) <> "" Then
                   Open App.path & "/Route/" & S & ".txt" For Input As #2
                   Line Input #2, mulu
                   Close #2
          Else
                  MsgBox "����δ֪ԭ�򣬳��ִ��󣬸ø������������", vbOK, "����"
                  Exit Sub
          End If
          Play = False
          WindowsMediaPlayer1.URL = mulu
          ��һ��.Text = S
          chongbo = "T"
          mulu = ""
          bongfang = S
      End If
    End If
  ElseIf Form2.S2.Text = "�����Դ�" Then
     music = Command.Count
     music = Int(Rnd * music)
     Call Command_Click(music)
  ElseIf Form2.S2.Text = "�û����" Then
     music = Command.Count + List1.ListCount
     music = Int(Rnd * music)
     If music <= Command.Count - 1 Then
         Call Command_Click(music)
     Else
          S = List1.List(music)
          If Dir(App.path & "/Route/" & S & ".txt", vbNormal + vbDirectory) <> "" Then
                   Open App.path & "/Route/" & S & ".txt" For Input As #2
                   Line Input #2, mulu
                   Close #2
          Else
                  MsgBox "����δ֪ԭ�򣬳��ִ��󣬸ø������������", vbOK, "����"
                  Exit Sub
          End If
          Play = False
          WindowsMediaPlayer1.URL = mulu
          ��һ��.Text = S
          chongbo = "T"
          mulu = ""
          bongfang = S
    End If
  End If
Else
  S = bofang - ".mp3"
  For i = 0 To Command.Count - 1
    If i < Command.Count - 1 Then
      If Command(i).Caption = S Then Call Command_Click(i + 1)
      ��һ��.Enabled = False
      Exit Sub
    Else
      If Command(i).Caption = S Then Call Command_Click(1)
      ��һ��.Enabled = False
      Exit Sub
    End If
  Next i
  For i = 0 To List1.ListCount - 1
      If List1.List(i) = bofang Then
         List1.Selected(i + 1) = True
         Call caozuo_Click(2)
      End If
  Next i
End If
  S = ""
  ��һ��.Enabled = False
End Sub

Private Sub ��ϸ��Ϣ_Click()
Form3.Show
End Sub

Private Sub ���ڲ���ʲô_Click()
MsgBox "���ڲ��ŵ���:" & bofang, vbOK, "����"
End Sub

Private Sub �ö�_Timer()
   If GetForegroundWindow <> Me.hWnd Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
End Sub
Private Sub Ƥ��()
On Error Resume Next
If Dir(path) <> "" Then
   SkinH_AttachEx path, ""
End If
End Sub

Private Sub ��ʼ��()
Dim tag As String

On Error Resume Next

If Dir(App.path & "\Save\Save.ini") <> "" Then
   Open App.path & "\Save.ini" For Input As #1
   Line Input #1, tag
   kaishi = tag
   Line Input #1, tag
   If tag = "True" Then pifu = True
   Line Input #1, tag
   If pifu = False Then
     path = App.path & "\Skin\" & tag & ".she"
     SSkin = tag
     Call Ƥ��
   Else
     path = tag
     Call Ƥ��
   End If
   Line Input #1, tag
     If tag = True Then Call SS_Click
   Close #1
Else
   path = App.path & "\Skin\�й���.she"
   Call Ƥ��
End If

If kaishi = "" Then
  bofang = "����ʲô��û��Ŷ����ȥѡ����Ҫ���ĸ����ɣ�"
Else
  bofang = kaishi
  WindowsMediaPlayer1.URL = kaishi
  C1.Enabled = True
  C2.Enabled = True
  C3.Enabled = True
  C4.Enabled = True
End If
Call �����б�
End Sub

Public Sub �����б�()
   Open App.path & "\Save\SaveMusic.ini" For Input As #1
   Do While Not EOF(1)
   Line Input #1, mulu
   List1.AddItem mulu
   Loop
   Close #1
   mulu = ""
Call ��ȡ����
End Sub

Public Sub ��ȡ����()
Dim tag As String
   Open App.path & "\Save\����.ini" For Input As #1
   Line Input #1, tag
   Form2.S1.Text = tag
   Line Input #1, tag
   Form2.S2.Text = tag
   Line Input #1, tag
   Form2.S3.Text = tag
   Line Input #1, tag
   Form2.T1.Text = tag
   Line Input #1, tag
   Form2.T2.Text = tag
   Line Input #1, tag
   Form2.chexiao.Text = tag
   Form2.HS1.Value = Val(tag)
   Line Input #1, tag
     If tag = True Then Form2.Ck1.Value = 1 Else Form2.Ck1.Value = 0
   Line Input #1, tag
     If tag = True Then Form2.Ck2.Value = 1 Else Form2.Ck2.Value = 0
   Line Input #1, tag
     If tag = True Then Form2.zhiding.Value = 1 Else Form2.zhiding.Value = 0
   Line Input #1, tag
     If tag = True Then Form2.Ck4.Value = 1 Else Form2.Ck4.Value = 0
   Close #1
End Sub

Private Sub �Զ���Ƥ��_Click()
On Error GoTo ErrHandler
CD1.InitDir = App.path & "\Skin"
CD1.Filter = "SHE(*.she)|*.she|"
CD1.Action = 1
If CD1.FileName = "" Then Exit Sub
pifu = True
path = CD1.FileName
Call Ƥ��

ErrHandler:
    Exit Sub
End Sub

