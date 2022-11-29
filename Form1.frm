VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Music(榮哥制作)"
   ClientHeight    =   10470
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   8835
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10470
   ScaleWidth      =   8835
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton SS 
      BackColor       =   &H8000000D&
      Caption         =   "简"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "重命名歌单"
      Height          =   375
      Index           =   2
      Left            =   3720
      TabIndex        =   55
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton gedan 
      Caption         =   "删除歌单"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   54
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton gedan 
      Caption         =   "添加歌单"
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
   Begin VB.Timer 下一首 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8160
      Top             =   6000
   End
   Begin VB.CommandButton C 
      Caption         =   "标记为开始乐曲"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
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
   Begin VB.TextBox 上一首 
      Height          =   270
      Left            =   2400
      TabIndex        =   50
      Text            =   "保存上一首信息"
      Top             =   8880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer 透明度 
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
   Begin VB.Timer 置顶 
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
      Caption         =   "重播"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "老男孩"
      Height          =   375
      Index           =   35
      Left            =   5880
      TabIndex        =   47
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton C 
      Caption         =   "关闭音乐"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "设置"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "打开选定曲目目录"
      BeginProperty Font 
         Name            =   "宋体"
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
      ToolTipText     =   "打开你选择的曲目所在目录"
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton caozuo 
      Caption         =   "播放指定"
      BeginProperty Font 
         Name            =   "宋体"
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
      ToolTipText     =   "播放选择的歌曲"
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
      DialogTitle     =   "选择添加文件"
      Filter          =   "MP3(*.mp3)|*.mp3|MIDI(*.midi)|*.midi|CD Audio(*.cda)|*.cda|WAV(*.wav)|*.wav|WMA(*.wma)|*.wma|其他(*.*)|*.*|"
   End
   Begin VB.CommandButton caozuo 
      Caption         =   "删除/清空"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "保存歌单以及曲目"
      BeginProperty Font 
         Name            =   "宋体"
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
      ToolTipText     =   "保存左侧添加出的歌曲，下次程序启动时自动添加"
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton caozuo 
      Caption         =   "添加歌曲"
      BeginProperty Font 
         Name            =   "宋体"
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
      ToolTipText     =   "将你自己的歌曲添加到左侧列表"
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
      Caption         =   "我们都一样"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "闹啥子嘛闹"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "奔跑吧兄弟"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "爱的供养"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "盛夏的果实"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "终于等到你"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "没那么简单"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "谁 Control"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "突然的自我"
      BeginProperty Font 
         Name            =   "宋体"
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
         Name            =   "宋体"
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
      Caption         =   "打工行"
      BeginProperty Font 
         Name            =   "宋体"
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
         Name            =   "宋体"
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
      ToolTipText     =   "核爆神曲"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      Caption         =   "曲终人散"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "心情车站"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "男人的好"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "每个人都看到了希望"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "爱一个人好难"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "K歌之王"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "海阔天空"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "剑心"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "大哥"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "李白"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "巴黎夜雨"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "暂停"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "依然爱你"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "你把爱情给了谁"
      Height          =   615
      Index           =   31
      Left            =   1680
      TabIndex        =   10
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command 
      Caption         =   "失恋达人四部曲"
      Height          =   615
      Index           =   30
      Left            =   0
      TabIndex        =   9
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command 
      Caption         =   "真的爱你"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "和你一样"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "离歌"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "浮夸"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "大海"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "变"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "王妃"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "父亲"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "北京北京"
      BeginProperty Font 
         Name            =   "宋体"
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
   Begin VB.Label 提示 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "提示 :保存添加后请勿更改歌曲保存目录，修改后请重新添加"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "用户添加曲目"
      BeginProperty Font 
         Name            =   "宋体"
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
   Begin VB.Menu 操作 
      Caption         =   "操作"
      Begin VB.Menu 正在播放什么 
         Caption         =   "正在播放什么"
         Shortcut        =   {F1}
      End
      Begin VB.Menu 自身目录 
         Caption         =   "打开程序目录"
      End
      Begin VB.Menu close 
         Caption         =   "删除/清空"
      End
      Begin VB.Menu SaveMusic 
         Caption         =   "保存添加乐曲"
         Shortcut        =   ^D
      End
      Begin VB.Menu Save设置 
         Caption         =   "保存设置"
         Shortcut        =   ^S
      End
      Begin VB.Menu end 
         Caption         =   "退出"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu MenuTray 
      Caption         =   "控制"
      Begin VB.Menu shower 
         Caption         =   "显示"
      End
      Begin VB.Menu end2 
         Caption         =   "关闭程序"
      End
      Begin VB.Menu wu3 
         Caption         =   "-"
      End
      Begin VB.Menu Ctrl 
         Caption         =   "上一首"
         Index           =   0
      End
      Begin VB.Menu Ctrl 
         Caption         =   "下一首"
         Index           =   1
      End
      Begin VB.Menu Ctrl 
         Caption         =   "暂停"
         Index           =   2
      End
      Begin VB.Menu Ctrl 
         Caption         =   "继续"
         Index           =   3
      End
      Begin VB.Menu Ctrl 
         Caption         =   "重播"
         Index           =   4
      End
      Begin VB.Menu Ctrl 
         Caption         =   "停止"
         Index           =   5
      End
   End
   Begin VB.Menu 信息 
      Caption         =   "信息"
      Begin VB.Menu 详细信息 
         Caption         =   "详细信息"
      End
   End
   Begin VB.Menu 设置 
      Caption         =   "设置"
      Begin VB.Menu open 
         Caption         =   "打开设置"
      End
      Begin VB.Menu 无 
         Caption         =   "-"
      End
      Begin VB.Menu 背景 
         Caption         =   "窗体背景色"
      End
   End
   Begin VB.Menu Skin 
      Caption         =   "皮肤"
      Begin VB.Menu SkinEx 
         Caption         =   "黑色酷炫"
         Index           =   0
      End
      Begin VB.Menu SkinEx 
         Caption         =   "积木风格"
         Index           =   1
      End
      Begin VB.Menu SkinEx 
         Caption         =   "蓝色经典"
         Index           =   2
      End
      Begin VB.Menu SkinEx 
         Caption         =   "浪漫粉红"
         Index           =   3
      End
      Begin VB.Menu SkinEx 
         Caption         =   "绿色青春"
         Index           =   4
      End
      Begin VB.Menu SkinEx 
         Caption         =   "绿色活力"
         Index           =   5
      End
      Begin VB.Menu SkinEx 
         Caption         =   "中国红"
         Index           =   6
      End
      Begin VB.Menu SkinEx 
         Caption         =   "蓝色淡雅"
         Index           =   7
      End
      Begin VB.Menu wu 
         Caption         =   "-"
      End
      Begin VB.Menu 自定义皮肤 
         Caption         =   "使用自定义皮肤"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim houzhui As String, suoyin%, sui%, l%, TS%, bofang As String, chongbo As String, Play As Boolean, mulu As String, biaoji As String, youji As Boolean, gaibian As Boolean
'**********保持置顶的API**************
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'***********窗体透明度API************************
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Const WS_EX_LAYERED = &H80000 '定义常量 WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20) '定义常量 GWL_EXSTYLE =(-20)
Const LWA_ALPHA = &H2 '定义常量 LWA_ALPHA = &H2
Const LWA_COLORKEY = &H1 '定义常量 LWA_COLORKEY = &H1
'***********移动窗体API**********************
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Const WM_NCLBUTTONDOWN = &HA1 '定义常量 WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2 '定义常量 HTCAPTION = 2

Private Sub C_Click(Index As Integer)
'4个操作按钮

Select Case Index
'关闭音乐***************************
Case Is = 0
a = MsgBox("你真的要关闭正在播放的音乐吗？", vbYesNo, "Music")
If a = vbNo Then Exit Sub
chongbo = "F"
WindowsMediaPlayer1.Controls.stop
bofang = "现在什么都没放哦，快去选择你要听的歌曲吧！"
C2.Caption = "暂停"
C2.Enabled = False
C3.Enabled = False
C1.Enabled = False
'暂停*********************************
Case Is = 1
If C(1).Caption = "暂停" And WindowsMediaPlayer1.URL <> "" Then
     chongbo = "F"
     WindowsMediaPlayer1.Controls.pause '暂停播放
     C(1).Caption = "继续"
Else
     chongbo = "T"
     C(1).Caption = "暂停"
     WindowsMediaPlayer1.Controls.Play
End If
'重播************************************
Case Is = 2
C(1).Caption = "暂停"
WindowsMediaPlayer1.Controls.pause
chongbo = "T"
'标记为开始音乐******************************
Case Is = 3
If C(4).Caption = "标记为开始乐曲" Then
   For i = 0 To Command.Count - 1
     Command(i).Font.Bold = False
   Next i
   biaoji = bofang
   kaishi = WindowsMediaPlayer1.URL
   For i = 0 To Command.Count - 1
     If Command(i).Caption = 上一首.Text Then Command(i).Font.Bold = True
   Next i
   MsgBox "标记成功！", vbYes, "Music"
Else
   kaishi = AppDisk & "aLTEz" & houzhui
End If
'设置****************************
Case Is = 4
Form2.Show
'结束**************************
End Select


End Sub

Private Sub caozuo_Click(Index As Integer)
Select Case Index
Case Is = 0   '添加

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
      a = MsgBox("这首您已添加过，是否继续？", vbYesNo, "Music")
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
                '用户按了“取消”按钮。
                Exit Sub
'----------------------------------------
Case Is = 1  '目录
提示.Caption = "提示:若选择多项则默认打开随机一项目录,若未选择则打开程序目录"
mulu = App.path + "\"
Shell "explorer.exe mulu", 1
'------------------------------------------
Case Is = 2   '播放
Dim S As String
提示.Caption = "提示:若选择多项则默认播放第一项"

   For i = List1.ListCount - 1 To 0 Step -1
     If List1.Selected(i) = True Then S = List1.List(i)
   Next
   If S = "" Then
      MsgBox "您还未选择歌曲", vbYes, "错误"
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
   MsgBox "由于未知原因，出现错误，该歌曲请重新添加", vbOK, "错误"
   Exit Sub
End If
   Play = False
   上一首.Text = WindowsMediaPlayer1.URL
   WindowsMediaPlayer1.URL = mulu
   chongbo = "T"
   mulu = ""
   bongfang = S
   C4.Enabled = False
   C2.Enabled = True
   C3.Enabled = True
   C1.Enabled = True
'---------------------------------------------
Case Is = 3   '删除

a = MsgBox("[删除选定歌曲] 请选确定，[清空列表] 请选取消", vbYesNo, "删除或清空")
If a = vbNo Then b = MsgBox("你确定清空列表吗？", vbYesNo, "提示")
If a = vbYes Then
   For i = List1.ListCount - 1 To 0 Step -1
     suoyin = suoyin - 1
     If List1.Selected(i) = True Then
         Kill (App.path & "\Route\" & List1.List(i) & ".txt")
         List1.RemoveItem (i)
     End If
   Next i
   MsgBox "删除指定歌曲成功！", vbYes, "Music"
Else
    If b = vbYes Then List1.Clear
        suoyin = 0
        For i = List1.ListCount - 1 To 0 Step -1
           If List1.Selected(i) = True Then List1.RemoveItem (i)
           Kill (App.path & "\Route\" & List1.List(i) & ".txt")
        Next i
        Open App.path & "\SaveMusic.ini" For Output As #1
        Close #1
        MsgBox "清空列表成功！", vbYes, "Music"
End If

'-----------------------------------------------
Case Is = 4  '保存

Dim n As String
Open App.path & "\SaveMusic.ini" For Output As #1
For i = 0 To List1.ListCount - 1
   n = List1.List(i)
   Print #1, n '保存曲目
Next i
Close #1
MsgBox "[用户添加曲目]保存成功！", vbYes, "提示"
'-----------------------------------------------
End Select
End Sub





Private Sub Command_Click(Index As Integer)
    上一首.Text = WindowsMediaPlayer1.URL
    WindowsMediaPlayer1.URL = AppDisk & Command(Index).Caption & houzhui
    Play = True
    If 上一首.Text = "" Then 上一首.Text = Command(Index).Caption
    bofang = Command(Index).Caption
    If Command(Index).Caption = biao Then C(3).Caption = "取消作为开始音乐" Else C(3).Caption = "标记为开始音乐"
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
   If 上一首.Text = "" Then Exit Sub
    If Dir(上一首.Text, vbNormal + vbDirectory) <> "" Then
         WindowsMediaPlayer1.URL = 上一首.Text
     Else
        MsgBox "由于源文件移动或被删除，无法播放上一首，请重新定位歌曲文件", vbOK, "错误"
        Exit Sub
     End If
   chongbo = "T"
 Case 1
    If Form2.S1 = "循环播放(默认)" And Form2.Ck2.Value = 1 Then
       MsgBox "你当前模式为[循环播放]，且是[一首无限循环状态],无法播放下一首", vbYes, "错误"
       Exit Sub
    End If
    If bofang = "" Then
       MsgBox "你还没有开始播放", vbOK, "错误"
       Exit Sub
    End If
    下一首.Enabled = True
 Case 2
    WindowsMediaPlayer1.Controls.pause
    C2.Caption = "继续"
 Case 3
    WindowsMediaPlayer1.Controls.Play
    C2.Caption = "暂停"
 Case 4
    C2.Caption = "暂停"
    WindowsMediaPlayer1.Controls.pause
    chongbo = "T"
    
 Case 5
   a = MsgBox("你真的要关闭正在播放的音乐吗？", vbYesNo, "Music")
   If a = vbNo Then Exit Sub
   chongbo = "F"
   WindowsMediaPlayer1.Controls.stop
   bofang = "现在什么都没放哦，快去选择你要听的歌曲吧！"
   C2.Caption = "暂停"
   C2.Enabled = False
   C3.Enabled = False
   C1.Enabled = False
 End Select

End Sub

Private Sub End_Click()
    Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
   'Shell (App.path & "/Route/删除.bat")
   Call save
End
End Sub

Private Sub end2_Click()
    Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
   'Shell (App.path & "/Route/删除.bat")
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
'定义鼠标移动到托盘上时显示的Tip
.szTip = "(Music v2.21 BY:榮哥)" & vbNullChar
.cbSize = Len(nfIconData)
End With
Call Shell_NotifyIcon(NIM_ADD, nfIconData)

suoyin2 = 1
Cb.Text = Cb.List(0)
shower.Visible = False
end2.Visible = False
wu3.Visible = False
pifu = False
上一首.Text = ""
chongbo = "T"
houzhui = ".mp3"
kaishi = AppDisk & "依然爱你" & houzhui
biaoji = ""
saves = False
w = Me.Width
h = Me.Height
Call 初始化
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

'鼠标按下左键可以拖动窗体
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = 1 Then '如果 Button = 1 鼠标按下了左键则
      Call ReleaseCapture '呼叫 ReleaseCapture
      Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&) '开始拖动窗体
   End If
   
   Dim lMsg As Single
    lMsg = X / Screen.TwipsPerPixelX
   Select Case lMsg
    Case WM_LBUTTONUP
    '单击左键，显示窗体
     shower.Visible = False
     end2.Visible = False
     wu3.Visible = False
     If GetForegroundWindow <> Me.hWnd Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
     'ShowWindow Me.hWnd, SW_RESTORE
   Case WM_RBUTTONUP
      shower.Visible = True
      end2.Visible = True
      wu3.Visible = True
      PopupMenu MenuTray '如果是在系统Tray图标上点右键，则弹出菜单MenuTray
   Case WM_MOUSEMOVE
   Case WM_LBUTTONDOWN
   Case WM_LBUTTONDBLCLK
   Case WM_RBUTTONDOWN
   Case WM_RBUTTONDBLCLK
   Case Else
  End Select
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If youji = True Then PopupMenu 操作
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
   'Shell (App.path & "/Route/删除.bat")
   Call save
   WindowsMediaPlayer1.Controls.stop '停止播放
   WindowsMediaPlayer1.close 'WindowsMediaPlayer1.关闭
   Set Form1 = Nothing '释放窗体对象
   End '结束程序
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
   p = InputBox("输入歌单名", "新建")
   If p = "" Then Exit Sub
   For i = 0 To suoyin2
    If Cb.List(i) = p Then
      MsgBox "歌单名重复", vbYes, "警告"
      Exit Sub
    End If
   Next i
   suoyin2 = suoyin2 + 1
   Cb.AddItem p, suoyin2
   
Case 1
   p = Cb.ListIndex
   a = MsgBox("你真的要删除歌单[" & Cb.List(p) & "]吗", vbYesNo, "Music")
   If a = vbYes Then Cb.RemoveItem p
   Cb.Text = Cb.List(0)
   suoyin2 = suoyin2 - 1
Case 2
   p = Cb.ListIndex
   a = InputBox("请为歌单[" & Cb.List(p) & "]输入一个新的名字", "重命名")
   If a = "" Then Exit Sub
   Cb.List(p) = a
End Select
End Sub

Private Sub List1_DblClick()
Dim S As String, tag As String
sb = List1.ListIndex '从0起算的第几笔
S = List1.List(sb)
If Dir(App.path & "/Route/" & S & ".txt", vbNormal + vbDirectory) <> "" Then
   Open App.path & "/Route/" & S & ".txt" For Input As #1
   Line Input #1, tag
   WindowsMediaPlayer1.URL = tag
   Close #1
Else
   MsgBox "历史文件未找到，该历史无法播放", vbOK, "错误"
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
   Print #1, n '保存曲目
Next i
Close #1
MsgBox "[用户添加曲目]保存成功！", vbYes, "提示"
End Sub

Private Sub Save设置_Click()
Call 保存设置
End Sub

Private Sub shower_Click()
Form1.Show
End Sub

Private Sub SkinEx_Click(Index As Integer)
pifu = False
SSkin = SkinEx(Index).Caption
path = App.path & "\Skin\" & SSkin & ".she"
Call 皮肤
End Sub

Private Sub SS_Click()
If SS.Caption = "简" Then
   SS.Caption = "放"
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
   提示.Top = 提示.Top - 4920
   Line2.Visible = False
   Line1.Y1 = Line1.Y1 - 4920
   Line1.Y2 = Line1.Y1
   Form1.Height = Form1.Height - 4920
Else
   SS.Caption = "简"
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
   提示.Top = 提示.Top + 4920
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
    If Form2.S2.Text = "程序自带" Then
       chongbo = "F"
       WindowsMediaPlayer1.Controls.pause
       sui = Int(Rnd * 36)
       WindowsMediaPlayer1.URL = App.path & Command(sui).Caption & houzhui
       bofang = Command(sui).Caption & houzhui
       TS = 0
    ElseIf Form2.S2.Text = "用户添加" Then
         If List1.ListCount = 0 And TS <> 1 Then
             MsgBox "你还没有添加任何歌曲！", vbOK, "你选择的随机范围是[用户添加]"
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
    ElseIf Form2.S2.Text = "全部" Then
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
End If '结束 如果
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
    If Form2.S2.Text = "程序自带" Then
       chongbo = "F"
       WindowsMediaPlayer1.Controls.pause
       sui = Int(Rnd * 36)
       WindowsMediaPlayer1.URL = App.path & Command(sui).Caption & houzhui
       bofang = Command(sui).Caption & houzhui
       TS = 0
    ElseIf Form2.S2.Text = "用户添加" Then
         If List1.ListCount = 0 And TS <> 1 Then
             MsgBox "你还没有添加任何歌曲！", vbOK, "你选择的随机范围是[用户添加]"
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
    ElseIf Form2.S2.Text = "全部" Then
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
End If '结束 如果
End Sub

Private Sub Tge_Timer()
If Dir(App.path & "歌单.ini", vbNormal + vbDirectory) <> "" Then Exit Sub
   Open App.path & "/Route/" & ShortName & ".txt" For Output As #1
    Print #1, FullName
   Close #1
End Sub

Private Sub Timer1_Timer()
If WindowsMediaPlayer1.playState = wmppsStopped Or WindowsMediaPlayer1.playState = wmppsReady Then
   If chongbo = "T" Then WindowsMediaPlayer1.Controls.Play   '开始循环再次播放
End If '结束 如果
End Sub




Private Sub 背景_Click()
a = MsgBox("您想恢复 程序默认颜色 还是 自定义新的颜色 做背景？,前者选[确定],后者选[取消]", vbYesNo, "背景色")
If a = vbYes Then
  Form1.BackColor = &HC0C0C0
  Exit Sub
End If
'将 Cancel 设置成 True。
                CD1.CancelError = True
                On Error GoTo ErrHandler
                '设置 Flags 属性。
                CD1.Flags = cdlCCRGBInit
                '显示“颜色”对话框。
                CD1.Action = 3
                '将窗体的背景颜色设置成选定的'颜色。
                Form1.BackColor = CD1.Color
                Exit Sub

ErrHandler:
                '用户按了“取消”按钮。
                Exit Sub
End Sub


Private Sub 透明度_Timer()
Dim S%
S = Val(touming.Text)
    Ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, Ret
    SetLayeredWindowAttributes Form1.hWnd, 0, S, LWA_ALPHA
    透明度.Enabled = False
End Sub

Private Sub 下一首_Timer()
Dim music%, S As String
If Form2.S1.Text = "随机播放" Then
 Randomize
  If Form2.S2.Text = "所有" Then
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
                  MsgBox "由于未知原因，出现错误，该歌曲请重新添加", vbOK, "错误"
                  Exit Sub
          End If
          Play = False
          WindowsMediaPlayer1.URL = mulu
          上一首.Text = S
          chongbo = "T"
          mulu = ""
          bongfang = S
      End If
    End If
  ElseIf Form2.S2.Text = "程序自带" Then
     music = Command.Count
     music = Int(Rnd * music)
     Call Command_Click(music)
  ElseIf Form2.S2.Text = "用户添加" Then
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
                  MsgBox "由于未知原因，出现错误，该歌曲请重新添加", vbOK, "错误"
                  Exit Sub
          End If
          Play = False
          WindowsMediaPlayer1.URL = mulu
          上一首.Text = S
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
      下一首.Enabled = False
      Exit Sub
    Else
      If Command(i).Caption = S Then Call Command_Click(1)
      下一首.Enabled = False
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
  下一首.Enabled = False
End Sub

Private Sub 详细信息_Click()
Form3.Show
End Sub

Private Sub 正在播放什么_Click()
MsgBox "正在播放的是:" & bofang, vbOK, "帮助"
End Sub

Private Sub 置顶_Timer()
   If GetForegroundWindow <> Me.hWnd Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
End Sub
Private Sub 皮肤()
On Error Resume Next
If Dir(path) <> "" Then
   SkinH_AttachEx path, ""
End If
End Sub

Private Sub 初始化()
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
     Call 皮肤
   Else
     path = tag
     Call 皮肤
   End If
   Line Input #1, tag
     If tag = True Then Call SS_Click
   Close #1
Else
   path = App.path & "\Skin\中国红.she"
   Call 皮肤
End If

If kaishi = "" Then
  bofang = "现在什么都没放哦，快去选择你要听的歌曲吧！"
Else
  bofang = kaishi
  WindowsMediaPlayer1.URL = kaishi
  C1.Enabled = True
  C2.Enabled = True
  C3.Enabled = True
  C4.Enabled = True
End If
Call 加载列表
End Sub

Public Sub 加载列表()
   Open App.path & "\Save\SaveMusic.ini" For Input As #1
   Do While Not EOF(1)
   Line Input #1, mulu
   List1.AddItem mulu
   Loop
   Close #1
   mulu = ""
Call 读取设置
End Sub

Public Sub 读取设置()
Dim tag As String
   Open App.path & "\Save\设置.ini" For Input As #1
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

Private Sub 自定义皮肤_Click()
On Error GoTo ErrHandler
CD1.InitDir = App.path & "\Skin"
CD1.Filter = "SHE(*.she)|*.she|"
CD1.Action = 1
If CD1.FileName = "" Then Exit Sub
pifu = True
path = CD1.FileName
Call 皮肤

ErrHandler:
    Exit Sub
End Sub

