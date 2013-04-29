VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form StageEditor 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关卡编辑器"
   ClientHeight    =   8685
   ClientLeft      =   4890
   ClientTop       =   4335
   ClientWidth     =   9840
   Icon            =   "StageEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   9840
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton bro 
      Caption         =   "浏览"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Deletes 
      Caption         =   "删除本关"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   25
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton Addnews 
      Caption         =   "增加新关"
      Height          =   375
      Left            =   2400
      TabIndex        =   24
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton Nexts 
      Caption         =   "下一关"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   23
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton Lasts 
      Caption         =   "上一关"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   5760
      Width           =   975
   End
   Begin VB.PictureBox X5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   6
      Left            =   3600
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   92
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox X5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   5
      Left            =   3000
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   91
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox X5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   4
      Left            =   2400
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   90
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox X5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   3
      Left            =   1800
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   89
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox X5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   2
      Left            =   1200
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   88
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox X5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   1
      Left            =   600
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   87
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox X4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   6
      Left            =   3600
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   86
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox X4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   5
      Left            =   3000
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   85
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox X4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   4
      Left            =   2400
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   84
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox X4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   3
      Left            =   1800
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   83
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox X4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   2
      Left            =   1200
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   82
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox X4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   1
      Left            =   600
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   81
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox X3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   6
      Left            =   3600
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   80
      Top             =   3720
      Width           =   495
   End
   Begin VB.PictureBox X3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   5
      Left            =   3000
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   79
      Top             =   3720
      Width           =   495
   End
   Begin VB.PictureBox X3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   4
      Left            =   2400
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   78
      Top             =   3720
      Width           =   495
   End
   Begin VB.PictureBox X3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   3
      Left            =   1800
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   77
      Top             =   3720
      Width           =   495
   End
   Begin VB.PictureBox X3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   2
      Left            =   1200
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   76
      Top             =   3720
      Width           =   495
   End
   Begin VB.PictureBox X3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   1
      Left            =   600
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   75
      Top             =   3720
      Width           =   495
   End
   Begin VB.PictureBox X2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   6
      Left            =   3600
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   74
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox X2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   5
      Left            =   3000
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   73
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox X2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   4
      Left            =   2400
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   72
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox X2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   3
      Left            =   1800
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   71
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox X2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   2
      Left            =   1200
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   70
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox X2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   1
      Left            =   600
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   69
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox X1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   6
      Left            =   3600
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   68
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox X1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   5
      Left            =   3000
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   67
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox X1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   4
      Left            =   2400
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   66
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox X1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   3
      Left            =   1800
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   65
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox X1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   2
      Left            =   1200
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   64
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox X1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   1
      Left            =   600
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   63
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Y6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   3480
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   62
      Top             =   5040
      Width           =   135
   End
   Begin VB.PictureBox Y5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   3480
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   61
      Top             =   4440
      Width           =   135
   End
   Begin VB.PictureBox Y5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   2880
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   60
      Top             =   4440
      Width           =   135
   End
   Begin VB.PictureBox Y5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   2280
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   59
      Top             =   4440
      Width           =   135
   End
   Begin VB.PictureBox Y5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   1680
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   58
      Top             =   4440
      Width           =   135
   End
   Begin VB.PictureBox Y5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   1080
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   57
      Top             =   4440
      Width           =   135
   End
   Begin VB.PictureBox Y6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   2880
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   56
      Top             =   5040
      Width           =   135
   End
   Begin VB.PictureBox Y4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   3480
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   55
      Top             =   3840
      Width           =   135
   End
   Begin VB.PictureBox Y4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   2880
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   54
      Top             =   3840
      Width           =   135
   End
   Begin VB.PictureBox Y4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   2280
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   53
      Top             =   3840
      Width           =   135
   End
   Begin VB.PictureBox Y4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   1680
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   52
      Top             =   3840
      Width           =   135
   End
   Begin VB.PictureBox Y4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   1080
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   51
      Top             =   3840
      Width           =   135
   End
   Begin VB.PictureBox Y6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   2280
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   50
      Top             =   5040
      Width           =   135
   End
   Begin VB.PictureBox Y3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   3480
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   49
      Top             =   3240
      Width           =   135
   End
   Begin VB.PictureBox Y3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   2880
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   48
      Top             =   3240
      Width           =   135
   End
   Begin VB.PictureBox Y3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   2280
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   47
      Top             =   3240
      Width           =   135
   End
   Begin VB.PictureBox Y3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   1680
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   46
      Top             =   3240
      Width           =   135
   End
   Begin VB.PictureBox Y3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   1080
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   45
      Top             =   3240
      Width           =   135
   End
   Begin VB.PictureBox Y6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   1680
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   44
      Top             =   5040
      Width           =   135
   End
   Begin VB.PictureBox Y2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   3480
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   43
      Top             =   2640
      Width           =   135
   End
   Begin VB.PictureBox Y2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   2880
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   42
      Top             =   2640
      Width           =   135
   End
   Begin VB.PictureBox Y2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   2280
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   41
      Top             =   2640
      Width           =   135
   End
   Begin VB.PictureBox Y2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   1680
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   40
      Top             =   2640
      Width           =   135
   End
   Begin VB.PictureBox Y2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   1080
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   39
      Top             =   2640
      Width           =   135
   End
   Begin VB.PictureBox Y6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   1080
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   38
      Top             =   5040
      Width           =   135
   End
   Begin VB.PictureBox Y1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   3480
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   37
      Top             =   2040
      Width           =   135
   End
   Begin VB.PictureBox Y1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   2880
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   36
      Top             =   2040
      Width           =   135
   End
   Begin VB.PictureBox Y1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   2280
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   35
      Top             =   2040
      Width           =   135
   End
   Begin VB.PictureBox Y1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   1680
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   34
      Top             =   2040
      Width           =   135
   End
   Begin VB.PictureBox Y1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   1080
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   33
      Top             =   2040
      Width           =   135
   End
   Begin VB.TextBox Path 
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton br 
      Caption         =   "打开"
      Height          =   375
      Left            =   8640
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Start 
      Caption         =   "测试"
      Height          =   375
      Left            =   8640
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton SaveStag 
      Caption         =   "保存"
      Height          =   375
      Left            =   8640
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox names 
      Height          =   270
      Left            =   1200
      TabIndex        =   5
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox counts 
      Height          =   270
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "1"
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox maker 
      Height          =   270
      Left            =   1200
      TabIndex        =   7
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox rems 
      Height          =   1455
      Left            =   4800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   120
      Width           =   3735
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "云端图标战士(*.csd)|*.csd"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "属性"
      Height          =   4935
      Left            =   4560
      TabIndex        =   93
      Top             =   1920
      Width           =   5295
      Begin VB.CommandButton Saveset 
         Caption         =   "保存修改"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   4440
         Width           =   975
      End
      Begin VB.ComboBox sType 
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "StageEditor.frx":0442
         Left            =   600
         List            =   "StageEditor.frx":0455
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   4455
      End
      Begin VB.Frame tp3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   120
         TabIndex        =   97
         Top             =   840
         Visible         =   0   'False
         Width           =   5055
         Begin VB.TextBox sWoodmax 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   1320
            TabIndex        =   10
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox sWoodids 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   960
            TabIndex        =   9
            Top             =   0
            Width           =   3975
         End
         Begin VB.Label lbs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "物品最大数量"
            Height          =   180
            Index           =   12
            Left            =   120
            TabIndex        =   108
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label lbs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "物品编号"
            Height          =   180
            Index           =   11
            Left            =   120
            TabIndex        =   107
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.Frame tp1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   120
         TabIndex        =   96
         Top             =   840
         Visible         =   0   'False
         Width           =   5055
         Begin VB.TextBox sDP 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   360
            TabIndex        =   21
            Top             =   3240
            Width           =   4575
         End
         Begin VB.TextBox sAP 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   360
            TabIndex        =   20
            Top             =   2880
            Width           =   4575
         End
         Begin VB.TextBox sMP 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   360
            TabIndex        =   19
            Top             =   2520
            Width           =   4575
         End
         Begin VB.TextBox sHP 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   360
            TabIndex        =   18
            Top             =   2160
            Width           =   4575
         End
         Begin VB.TextBox sWoodNum 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   270
            Left            =   1560
            TabIndex        =   17
            Top             =   1800
            Width           =   3375
         End
         Begin VB.TextBox sWoodID 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   270
            Left            =   1200
            TabIndex        =   16
            Top             =   1440
            Width           =   3735
         End
         Begin VB.TextBox sOPath 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   270
            Left            =   1560
            TabIndex        =   14
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox sPicID 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   960
            TabIndex        =   13
            Top             =   360
            Width           =   3975
         End
         Begin VB.TextBox sName 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   600
            TabIndex        =   12
            Top             =   0
            Width           =   4335
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "设置奖励物品"
            Height          =   180
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Width           =   4815
         End
         Begin VB.Label lbs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DP"
            Height          =   180
            Index           =   10
            Left            =   120
            TabIndex        =   106
            Top             =   3240
            Width           =   180
         End
         Begin VB.Label lbs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AP"
            Height          =   180
            Index           =   9
            Left            =   120
            TabIndex        =   105
            Top             =   2880
            Width           =   180
         End
         Begin VB.Label lbs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MP"
            Height          =   180
            Index           =   8
            Left            =   120
            TabIndex        =   104
            Top             =   2520
            Width           =   180
         End
         Begin VB.Label lbs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HP"
            Height          =   180
            Index           =   7
            Left            =   120
            TabIndex        =   103
            Top             =   2160
            Width           =   180
         End
         Begin VB.Label lbs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "物品最大数量"
            Enabled         =   0   'False
            Height          =   180
            Index           =   6
            Left            =   360
            TabIndex        =   102
            Top             =   1800
            Width           =   1080
         End
         Begin VB.Label lbs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "物品编号"
            Enabled         =   0   'False
            Height          =   180
            Index           =   5
            Left            =   360
            TabIndex        =   101
            Top             =   1440
            Width           =   720
         End
         Begin VB.Label lbs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "外部图片地址"
            Enabled         =   0   'False
            Height          =   180
            Index           =   4
            Left            =   360
            TabIndex        =   100
            Top             =   720
            Width           =   1080
         End
         Begin VB.Label lbs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "图片编号"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   99
            ToolTipText     =   "如果需要外部图片本栏不填"
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lbs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "名称"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   98
            Top             =   0
            Width           =   360
         End
      End
      Begin VB.Label zby 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   840
         TabIndex        =   111
         Top             =   240
         Width           =   90
      End
      Begin VB.Label zbd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ","
         Height          =   180
         Left            =   720
         TabIndex        =   110
         Top             =   240
         Width           =   90
      End
      Begin VB.Label zbx 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   600
         TabIndex        =   109
         Top             =   240
         Width           =   90
      End
      Begin VB.Label lbs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "类型"
         Enabled         =   0   'False
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   95
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lbs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "坐标"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   94
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "1          2            3         4           5           6          "
      Height          =   180
      Left            =   600
      TabIndex        =   121
      Top             =   8400
      Width           =   6210
   End
   Begin VB.Image WoodPic 
      Height          =   480
      Index           =   0
      Left            =   480
      Picture         =   "StageEditor.frx":0485
      Top             =   7920
      Width           =   480
   End
   Begin VB.Image WoodPic 
      Height          =   480
      Index           =   1
      Left            =   1560
      Picture         =   "StageEditor.frx":08C7
      Top             =   7920
      Width           =   480
   End
   Begin VB.Image WoodPic 
      Height          =   480
      Index           =   2
      Left            =   2640
      Picture         =   "StageEditor.frx":0D09
      Top             =   7920
      Width           =   480
   End
   Begin VB.Image WoodPic 
      Height          =   480
      Index           =   3
      Left            =   3600
      Picture         =   "StageEditor.frx":114B
      Top             =   7920
      Width           =   480
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "补血散"
      Height          =   180
      Index           =   1
      Left            =   960
      TabIndex        =   120
      Top             =   7920
      Width           =   540
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "提神剂"
      Height          =   180
      Index           =   2
      Left            =   2040
      TabIndex        =   119
      Top             =   7920
      Width           =   540
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "复活灯"
      Height          =   180
      Index           =   3
      Left            =   3120
      TabIndex        =   118
      Top             =   7920
      Width           =   540
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "免战旗"
      Height          =   180
      Index           =   4
      Left            =   4080
      TabIndex        =   117
      Top             =   7920
      Width           =   540
   End
   Begin VB.Image WoodPic 
      Height          =   480
      Index           =   4
      Left            =   4680
      Picture         =   "StageEditor.frx":158D
      Top             =   7920
      Width           =   480
   End
   Begin VB.Image WoodPic 
      Height          =   480
      Index           =   5
      Left            =   5760
      Picture         =   "StageEditor.frx":19CF
      Top             =   7920
      Width           =   480
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "战神锤"
      Height          =   180
      Index           =   5
      Left            =   5160
      TabIndex        =   116
      Top             =   7920
      Width           =   540
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "金钢铠"
      Height          =   180
      Index           =   6
      Left            =   6240
      TabIndex        =   115
      Top             =   7920
      Width           =   540
   End
   Begin VB.Label nb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   180
      Left            =   360
      TabIndex        =   114
      Top             =   1800
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "关"
      Height          =   180
      Index           =   6
      Left            =   720
      TabIndex        =   113
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "第"
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   112
      Top             =   1800
      Width           =   180
   End
   Begin VB.Image Wood 
      Height          =   495
      Index           =   0
      Left            =   480
      Picture         =   "StageEditor.frx":1E11
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   495
   End
   Begin VB.Image Wood 
      Height          =   495
      Index           =   1
      Left            =   960
      Picture         =   "StageEditor.frx":26DB
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   495
   End
   Begin VB.Image Wood 
      Height          =   495
      Index           =   2
      Left            =   1440
      Picture         =   "StageEditor.frx":2FA5
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "关卡注释："
      Height          =   180
      Index           =   4
      Left            =   3960
      TabIndex        =   32
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "关卡作者  ："
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   31
      Top             =   1320
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "关卡总关数："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   30
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "关卡名称  ："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   29
      Top             =   600
      Width           =   1080
   End
   Begin VB.Image P6 
      Height          =   495
      Index           =   5
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   495
   End
   Begin VB.Image P6 
      Height          =   495
      Index           =   4
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   495
   End
   Begin VB.Image P6 
      Height          =   495
      Index           =   3
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   495
   End
   Begin VB.Image P6 
      Height          =   495
      Index           =   2
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   495
   End
   Begin VB.Image P6 
      Height          =   495
      Index           =   1
      Left            =   600
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   495
   End
   Begin VB.Image P5 
      Height          =   495
      Index           =   6
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   495
   End
   Begin VB.Image P5 
      Height          =   495
      Index           =   5
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   495
   End
   Begin VB.Image P5 
      Height          =   495
      Index           =   4
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   495
   End
   Begin VB.Image P5 
      Height          =   495
      Index           =   3
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   495
   End
   Begin VB.Image P5 
      Height          =   495
      Index           =   2
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   495
   End
   Begin VB.Image P5 
      Height          =   495
      Index           =   1
      Left            =   600
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   495
   End
   Begin VB.Image P4 
      Height          =   495
      Index           =   6
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image P4 
      Height          =   495
      Index           =   5
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image P4 
      Height          =   495
      Index           =   4
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image P4 
      Height          =   495
      Index           =   3
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image P4 
      Height          =   495
      Index           =   2
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image P4 
      Height          =   495
      Index           =   1
      Left            =   600
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image P3 
      Height          =   495
      Index           =   5
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   495
   End
   Begin VB.Image P3 
      Height          =   495
      Index           =   4
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   495
   End
   Begin VB.Image P3 
      Height          =   495
      Index           =   3
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   495
   End
   Begin VB.Image P3 
      Height          =   495
      Index           =   2
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   495
   End
   Begin VB.Image P3 
      Height          =   495
      Index           =   1
      Left            =   600
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   495
   End
   Begin VB.Image P2 
      Height          =   495
      Index           =   6
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   495
   End
   Begin VB.Image P2 
      Height          =   495
      Index           =   5
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   495
   End
   Begin VB.Image P2 
      Height          =   495
      Index           =   4
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   495
   End
   Begin VB.Image P2 
      Height          =   495
      Index           =   3
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   495
   End
   Begin VB.Image P2 
      Height          =   495
      Index           =   2
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   495
   End
   Begin VB.Image P2 
      Height          =   495
      Index           =   1
      Left            =   600
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   495
   End
   Begin VB.Image P1 
      Height          =   495
      Index           =   6
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image P1 
      Height          =   495
      Index           =   5
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image P1 
      Height          =   495
      Index           =   4
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image P1 
      Height          =   495
      Index           =   3
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image P1 
      Height          =   495
      Index           =   2
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image P1 
      Height          =   495
      Index           =   1
      Left            =   600
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image P6 
      Height          =   495
      Index           =   6
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "关卡文件  ："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   1080
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   17
      Left            =   8160
      Picture         =   "StageEditor.frx":386F
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   16
      Left            =   7680
      Picture         =   "StageEditor.frx":3CB1
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   15
      Left            =   7200
      Picture         =   "StageEditor.frx":40F3
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   14
      Left            =   6720
      Picture         =   "StageEditor.frx":4535
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   13
      Left            =   6240
      Picture         =   "StageEditor.frx":4977
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   12
      Left            =   5760
      Picture         =   "StageEditor.frx":4DB9
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   11
      Left            =   5280
      Picture         =   "StageEditor.frx":51FB
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   10
      Left            =   4800
      Picture         =   "StageEditor.frx":563D
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   9
      Left            =   4320
      Picture         =   "StageEditor.frx":5A7F
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   8
      Left            =   3840
      Picture         =   "StageEditor.frx":5EC1
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   7
      Left            =   3360
      Picture         =   "StageEditor.frx":7843
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   6
      Left            =   2880
      Picture         =   "StageEditor.frx":7C85
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   5
      Left            =   2400
      Picture         =   "StageEditor.frx":7F8F
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   4
      Left            =   1920
      Picture         =   "StageEditor.frx":8319
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   3
      Left            =   1440
      Picture         =   "StageEditor.frx":AABB
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   2
      Left            =   960
      Picture         =   "StageEditor.frx":ADC5
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   1
      Left            =   480
      Picture         =   "StageEditor.frx":B0CF
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "1     2    3    4     5    6    7     8    9    10    11   12   13   14    15   16    17"
      Height          =   180
      Left            =   600
      TabIndex        =   27
      Top             =   7680
      Width           =   7920
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      Height          =   3495
      Left            =   600
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Image P3 
      Height          =   495
      Index           =   6
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   495
   End
End
Attribute VB_Name = "StageEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Addnews_Click()
counts = Val(counts) + 1
N = Val(counts)
ReDrawMap (N)
nb = N
If Val(counts) > 1 Then Deletes.Enabled = True Else Deletes.Enabled = False
If N = 1 Then Lasts.Enabled = False Else Lasts.Enabled = True
If N = Val(counts) Then Nexts.Enabled = False Else Nexts.Enabled = True
End Sub

Private Sub br_Click()
Open Path.Text For Binary As #1 Len = 32767
Get #1, 1, tMap
Close #1
names = tMap(1).Head.StageName
counts = tMap(1).Head.StagesCount
maker = tMap(1).Head.maker
rems = tMap(1).Head.Texts
ReDrawMap (1)
If Val(counts) > 1 Then Deletes.Enabled = True Else Deletes.Enabled = False
If N = 1 Then Lasts.Enabled = False Else Lasts.Enabled = True
If N < Val(counts) Then Nexts.Enabled = True Else Nexts.Enabled = False
End Sub

Private Sub bro_Click()
On Error GoTo errHandler1
cd.ShowOpen
Path.Text = cd.FileName
errHandler1:
End Sub

Private Sub Check1_Click()
lbs(5).Enabled = Check1.Value
lbs(6).Enabled = Check1.Value
sWoodID.Enabled = Check1.Value
sWoodNum.Enabled = Check1.Value
End Sub

Private Sub Deletes_Click()
counts = Val(counts) - 1
N = N - 1
ReDrawMap (N)
nb = N
If Val(counts) > 1 Then Deletes.Enabled = True Else Deletes.Enabled = False
If N = 1 Then Lasts.Enabled = False Else Lasts.Enabled = True
If N = Val(counts) Then Nexts.Enabled = False Else Nexts.Enabled = True
End Sub

Private Sub Lasts_Click()
N = N - 1
ReDrawMap (N)
nb = N
If Val(counts) > 1 Then Deletes.Enabled = True Else Deletes.Enabled = False
If N = 1 Then Lasts.Enabled = False Else Lasts.Enabled = True
If N = Val(counts) Then Nexts.Enabled = False Else Nexts.Enabled = True
End Sub

Private Sub Nexts_Click()
N = N + 1
ReDrawMap (N)
nb = N
If Val(counts) > 1 Then Deletes.Enabled = True Else Deletes.Enabled = False
If N = 1 Then Lasts.Enabled = False Else Lasts.Enabled = True
If N = Val(counts) Then Nexts.Enabled = False Else Nexts.Enabled = True
End Sub

Private Sub P1_Click(Index As Integer)
zbx.Caption = 1
zby.Caption = Index
sType.Enabled = True
lbs(1).Enabled = True
OpenS
End Sub

Private Sub P2_Click(Index As Integer)
zbx.Caption = 2
zby.Caption = Index
sType.Enabled = True
lbs(1).Enabled = True
OpenS
End Sub

Private Sub P3_Click(Index As Integer)
zbx.Caption = 3
zby.Caption = Index
sType.Enabled = True
lbs(1).Enabled = True
OpenS
End Sub

Private Sub P4_Click(Index As Integer)
zbx.Caption = 4
zby.Caption = Index
sType.Enabled = True
lbs(1).Enabled = True
OpenS
End Sub

Private Sub P5_Click(Index As Integer)
zbx.Caption = 5
zby.Caption = Index
sType.Enabled = True
lbs(1).Enabled = True
OpenS
End Sub

Private Sub P6_Click(Index As Integer)
zbx.Caption = 6
zby.Caption = Index
sType.Enabled = True
lbs(1).Enabled = True
OpenS
End Sub

Private Sub Saveset_Click()
tMap(N).P(zbx, zby).isWhat = sType.ListIndex
tMap(N).P(zbx, zby).Name = sName
tMap(N).P(zbx, zby).PicID = Val(sPicID)
tMap(N).P(zbx, zby).OtherPic = sOPath
tMap(N).P(zbx, zby).WoodType = Val(sWoodID)
tMap(N).P(zbx, zby).WoodNum = Val(sWoodNum)
tMap(N).P(zbx, zby).HP = Val(sHP)
tMap(N).P(zbx, zby).MP = Val(sMP)
tMap(N).P(zbx, zby).AP = Val(sAP)
tMap(N).P(zbx, zby).DP = Val(sDP)
ReDrawMap (N)
End Sub

Private Sub SaveStag_Click()
tMap(1).Head.StageName = names.Text
tMap(1).Head.StagesCount = counts.Text
tMap(1).Head.maker = maker.Text
tMap(1).Head.Texts = rems.Text
Open Path.Text For Binary As #1 Len = 32767
Put #1, 1, tMap
Close #1
End Sub

Private Sub Form_Load()
N = 1
ReDrawMap (1)
End Sub

Private Function COLOR(c As Boolean) As Long
If c Then COLOR = &H0& Else COLOR = &HC0C0C0
End Function

Private Sub sPicID_Change()
If Val(sPicID) = 0 Then
lbs(4).Enabled = True
sOPath.Enabled = True
Else
lbs(4).Enabled = False
sOPath.Enabled = False
End If
End Sub

Private Sub Start_Click()
If Val(MainForm.mxhpz.Caption) = 0 Then
MainForm.mxhpz.Caption = 10000
MainForm.mxmpz.Caption = 10000
MainForm.mapz.Caption = 1000
MainForm.mdpz.Caption = 1000
MainForm.Self.Picture = MainForm.MSP(6).Picture
End If
StageForm.Show 1
End Sub

Private Sub ReDrawMap(N As Byte)
On Error Resume Next
Randomize
For i = 1 To 6
X1(i).BackColor = COLOR(tMap(N).x(1, i))
Next
For i = 1 To 6
X2(i).BackColor = COLOR(tMap(N).x(2, i))
Next
For i = 1 To 6
X3(i).BackColor = COLOR(tMap(N).x(3, i))
Next
For i = 1 To 6
X4(i).BackColor = COLOR(tMap(N).x(4, i))
Next
For i = 1 To 6
X5(i).BackColor = COLOR(tMap(N).x(5, i))
Next

For i = 1 To 5
Y1(i).BackColor = COLOR(tMap(N).Y(1, i))
Next
For i = 1 To 5
Y2(i).BackColor = COLOR(tMap(N).Y(2, i))
Next
For i = 1 To 5
Y3(i).BackColor = COLOR(tMap(N).Y(3, i))
Next
For i = 1 To 5
Y4(i).BackColor = COLOR(tMap(N).Y(4, i))
Next
For i = 1 To 5
Y5(i).BackColor = COLOR(tMap(N).Y(5, i))
Next
For i = 1 To 5
Y6(i).BackColor = COLOR(tMap(N).Y(6, i))
Next

For i = 1 To 6
Select Case tMap(N).P(1, i).isWhat
Case 0
P1(i).Picture = LoadPicture
P1(i).Tag = ""
Case 1
    PicLib(0).Picture = LoadPicture(tMap(N).P(1, i).OtherPic)
    P1(i).Picture = PicLib(tMap(N).P(1, i).PicID)
    P1(i).Tag = "s"
Case 2
    PicLib(0).Picture = LoadPicture(tMap(N).P(1, i).OtherPic)
    P1(i).Picture = PicLib(tMap(N).P(1, i).PicID)
    P1(i).Tag = "n"
Case 3
    P1(i).Picture = Wood(Fix(Rnd * 100) Mod 2 + 1).Picture
    P1(i).Tag = "w"
Case 4
    P1(i).Picture = Wood(0).Picture
    Self.Left = P1(i).Left
    Self.Top = P1(i).Top
    P1(i).Tag = "f"
    sX = 1
    sY = i
End Select
P1(i).ToolTipText = tMap(N).P(1, i).Name
Next

For i = 1 To 6
Select Case tMap(N).P(2, i).isWhat
Case 0
P2(i).Picture = LoadPicture
P2(i).Tag = ""
Case 1
    PicLib(0).Picture = LoadPicture(tMap(N).P(2, i).OtherPic)
    P2(i).Picture = PicLib(tMap(N).P(2, i).PicID)
    P2(i).Tag = "s"
Case 2
    PicLib(0).Picture = LoadPicture(tMap(N).P(2, i).OtherPic)
    P2(i).Picture = PicLib(tMap(N).P(2, i).PicID)
    P2(i).Tag = "n"
Case 3
    P2(i).Picture = Wood(Fix(Rnd * 100) Mod 2 + 1).Picture
    P2(i).Tag = "w"
Case 4
    P2(i).Picture = Wood(0).Picture
    Self.Left = P2(i).Left
    Self.Top = P2(i).Top
    P1(i).Tag = "f"
    sX = 2
    sY = i
End Select
P2(i).ToolTipText = tMap(N).P(2, i).Name
Next

For i = 1 To 6
Select Case tMap(N).P(3, i).isWhat
Case 0
P3(i).Picture = LoadPicture
P3(i).Tag = ""
Case 1
    PicLib(0).Picture = LoadPicture(tMap(N).P(3, i).OtherPic)
    P3(i).Picture = PicLib(tMap(N).P(3, i).PicID)
    P3(i).Tag = "s"
Case 2
    PicLib(0).Picture = LoadPicture(tMap(N).P(3, i).OtherPic)
    P3(i).Picture = PicLib(tMap(N).P(3, i).PicID)
    P3(i).Tag = "n"
Case 3
    P3(i).Picture = Wood(Fix(Rnd * 100) Mod 2 + 1).Picture
    P3(i).Tag = "w"
Case 4
    P3(i).Picture = Wood(0).Picture
    Self.Left = P3(i).Left
    Self.Top = P3(i).Top
    P1(i).Tag = "f"
    sX = 3
    sY = i
End Select
P3(i).ToolTipText = tMap(N).P(3, i).Name
Next

For i = 1 To 6
Select Case tMap(N).P(4, i).isWhat
Case 0
P4(i).Picture = LoadPicture
P4(i).Tag = ""
Case 1
    PicLib(0).Picture = LoadPicture(tMap(N).P(4, i).OtherPic)
    P4(i).Picture = PicLib(tMap(N).P(4, i).PicID)
    P4(i).Tag = "s"
Case 2
    PicLib(0).Picture = LoadPicture(tMap(N).P(4, i).OtherPic)
    P4(i).Picture = PicLib(tMap(N).P(4, i).PicID)
    P4(i).Tag = "n"
Case 3
    P4(i).Picture = Wood(Fix(Rnd * 100) Mod 2 + 1).Picture
    P4(i).Tag = "w"
Case 4
    P4(i).Picture = Wood(0).Picture
    Self.Left = P4(i).Left
    Self.Top = P4(i).Top
    P1(i).Tag = "f"
    sX = 4
    sY = i
End Select
P4(i).ToolTipText = tMap(N).P(4, i).Name
Next

For i = 1 To 6
Select Case tMap(N).P(5, i).isWhat
Case 0
P5(i).Picture = LoadPicture
P5(i).Tag = ""
Case 1
    PicLib(0).Picture = LoadPicture(tMap(N).P(5, i).OtherPic)
    P5(i).Picture = PicLib(tMap(N).P(5, i).PicID)
    P5(i).Tag = "s"
Case 2
    PicLib(0).Picture = LoadPicture(tMap(N).P(5, i).OtherPic)
    P5(i).Picture = PicLib(tMap(N).P(5, i).PicID)
    P5(i).Tag = "n"
Case 3
    P5(i).Picture = Wood(Fix(Rnd * 100) Mod 2 + 1).Picture
    P5(i).Tag = "w"
Case 4
    P5(i).Picture = Wood(0).Picture
    Self.Left = P5(i).Left
    Self.Top = P5(i).Top
    P1(i).Tag = "f"
    sX = 5
    sY = i
End Select
P5(i).ToolTipText = tMap(N).P(5, i).Name
Next

For i = 1 To 6
Select Case tMap(N).P(6, i).isWhat
Case 0
P6(i).Picture = LoadPicture
P6(i).Tag = ""
Case 1
    PicLib(0).Picture = LoadPicture(tMap(N).P(6, i).OtherPic)
    P6(i).Picture = PicLib(tMap(N).P(6, i).PicID)
    P6(i).Tag = "s"
Case 2
    PicLib(0).Picture = LoadPicture(tMap(N).P(6, i).OtherPic)
    P6(i).Picture = PicLib(tMap(N).P(6, i).PicID)
    P6(i).Tag = "n"
Case 3
    P6(i).Picture = Wood(Fix(Rnd * 100) Mod 2 + 1).Picture
    P6(i).Tag = "w"
Case 4
    P6(i).Picture = Wood(0).Picture
    Self.Left = P6(i).Left
    Self.Top = P6(i).Top
    P1(i).Tag = "f"
    sX = 6
    sY = i
End Select
P6(i).ToolTipText = tMap(N).P(6, i).Name
Next
End Sub

Private Sub sType_Click()
Select Case sType.ListIndex
Case 0
tp1.Visible = False
tp3.Visible = False
Case 1
tp1.Visible = True
tp3.Visible = False
Case 2
tp1.Visible = True
tp3.Visible = False
Case 3
tp3.Visible = True
tp1.Visible = False
Case 4
tp1.Visible = False
tp3.Visible = False
End Select

End Sub


Private Sub sWoodids_Change()
sWoodID.Text = sWoodids
End Sub

Private Sub sWoodmax_Change()
sWoodNum = sWoodmax
End Sub

Private Sub X1_Click(Index As Integer)
tMap(N).x(1, Index) = Not tMap(N).x(1, Index)
X1(Index).BackColor = COLOR(tMap(N).x(1, Index))
End Sub

Private Sub X2_Click(Index As Integer)
tMap(N).x(2, Index) = Not tMap(N).x(2, Index)
X2(Index).BackColor = COLOR(tMap(N).x(2, Index))
End Sub

Private Sub X3_Click(Index As Integer)
tMap(N).x(3, Index) = Not tMap(N).x(3, Index)
X3(Index).BackColor = COLOR(tMap(N).x(3, Index))
End Sub

Private Sub X4_Click(Index As Integer)
tMap(N).x(4, Index) = Not tMap(N).x(4, Index)
X4(Index).BackColor = COLOR(tMap(N).x(4, Index))
End Sub

Private Sub X5_Click(Index As Integer)
tMap(N).x(5, Index) = Not tMap(N).x(5, Index)
X5(Index).BackColor = COLOR(tMap(N).x(5, Index))
End Sub

Private Sub Y1_Click(Index As Integer)
tMap(N).Y(1, Index) = Not tMap(N).Y(1, Index)
Y1(Index).BackColor = COLOR(tMap(N).Y(1, Index))
End Sub

Private Sub Y2_Click(Index As Integer)
tMap(N).Y(2, Index) = Not tMap(N).Y(2, Index)
Y2(Index).BackColor = COLOR(tMap(N).Y(2, Index))
End Sub

Private Sub Y3_Click(Index As Integer)
tMap(N).Y(3, Index) = Not tMap(N).Y(3, Index)
Y3(Index).BackColor = COLOR(tMap(N).Y(3, Index))
End Sub

Private Sub Y4_Click(Index As Integer)
tMap(N).Y(4, Index) = Not tMap(N).Y(4, Index)
Y4(Index).BackColor = COLOR(tMap(N).Y(4, Index))
End Sub

Private Sub Y5_Click(Index As Integer)
tMap(N).Y(5, Index) = Not tMap(N).Y(5, Index)
Y5(Index).BackColor = COLOR(tMap(N).Y(5, Index))
End Sub

Private Sub Y6_Click(Index As Integer)
tMap(N).Y(6, Index) = Not tMap(N).Y(6, Index)
Y6(Index).BackColor = COLOR(tMap(N).Y(6, Index))
End Sub

Private Sub OpenS()
sType.ListIndex = tMap(N).P(zbx, zby).isWhat
sName = tMap(N).P(zbx, zby).Name
sPicID = tMap(N).P(zbx, zby).PicID
sOPath = tMap(N).P(zbx, zby).OtherPic
sWoodID = tMap(N).P(zbx, zby).WoodType
sWoodNum = tMap(N).P(zbx, zby).WoodNum
sHP = tMap(N).P(zbx, zby).HP
sMP = tMap(N).P(zbx, zby).MP
sAP = tMap(N).P(zbx, zby).AP
sDP = tMap(N).P(zbx, zby).DP
sWoodids = tMap(N).P(zbx, zby).WoodType
sWoodmax = tMap(N).P(zbx, zby).WoodNum
End Sub
