VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form StageForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   Icon            =   "StageForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9360
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer WaitBack 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Tag             =   "0"
      Top             =   5280
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   110
      Top             =   3960
      Width           =   3615
      Begin VB.Label State 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   120
         TabIndex        =   111
         Top             =   240
         Width           =   90
      End
   End
   Begin VB.Frame YD 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   3840
      TabIndex        =   102
      Top             =   3960
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton Flagy 
         Caption         =   "使用“免战旗”免战"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   2775
      End
      Begin VB.CommandButton Battle 
         Caption         =   "进入战斗"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Typeo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1080
         TabIndex        =   106
         Top             =   600
         Width           =   90
      End
      Begin VB.Label Nameof 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1080
         TabIndex        =   105
         Top             =   240
         Width           =   90
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "敌人类型："
         Height          =   180
         Index           =   14
         Left            =   120
         TabIndex        =   104
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "敌人名称："
         Height          =   180
         Index           =   13
         Left            =   120
         TabIndex        =   103
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "当前物品"
      Height          =   3495
      Left            =   6120
      TabIndex        =   83
      Top             =   360
      Width           =   3135
      Begin VB.CommandButton Use2 
         Caption         =   "使用"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Use1 
         Caption         =   "使用"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Ws 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Index           =   0
         Left            =   1440
         TabIndex        =   101
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Ws 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Index           =   5
         Left            =   1440
         TabIndex        =   100
         Top             =   2640
         Width           =   90
      End
      Begin VB.Label Ws 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Index           =   4
         Left            =   1440
         TabIndex        =   99
         Top             =   2160
         Width           =   90
      End
      Begin VB.Label Ws 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Index           =   3
         Left            =   1440
         TabIndex        =   98
         Top             =   1680
         Width           =   90
      End
      Begin VB.Label Ws 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Index           =   2
         Left            =   1440
         TabIndex        =   97
         Top             =   1200
         Width           =   90
      End
      Begin VB.Label Ws 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Index           =   1
         Left            =   1440
         TabIndex        =   96
         Top             =   720
         Width           =   90
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "个"
         Height          =   180
         Index           =   12
         Left            =   1800
         TabIndex        =   95
         Top             =   2640
         Width           =   180
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "个"
         Height          =   180
         Index           =   11
         Left            =   1800
         TabIndex        =   94
         Top             =   2160
         Width           =   180
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "个"
         Height          =   180
         Index           =   10
         Left            =   1800
         TabIndex        =   93
         Top             =   1680
         Width           =   180
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "个"
         Height          =   180
         Index           =   9
         Left            =   1800
         TabIndex        =   92
         Top             =   1200
         Width           =   180
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "个"
         Height          =   180
         Index           =   8
         Left            =   1800
         TabIndex        =   91
         Top             =   720
         Width           =   180
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "个"
         Height          =   180
         Index           =   7
         Left            =   1800
         TabIndex        =   90
         Top             =   240
         Width           =   180
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "金钢铠"
         Height          =   180
         Index           =   6
         Left            =   600
         TabIndex        =   89
         Top             =   2640
         Width           =   540
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "战神锤"
         Height          =   180
         Index           =   5
         Left            =   600
         TabIndex        =   88
         Top             =   2160
         Width           =   540
      End
      Begin VB.Image WoodPic 
         Height          =   480
         Index           =   5
         Left            =   120
         Picture         =   "StageForm.frx":0442
         Top             =   2640
         Width           =   480
      End
      Begin VB.Image WoodPic 
         Height          =   480
         Index           =   4
         Left            =   120
         Picture         =   "StageForm.frx":0884
         Top             =   2160
         Width           =   480
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "免战旗"
         Height          =   180
         Index           =   4
         Left            =   600
         TabIndex        =   87
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "复活灯"
         Height          =   180
         Index           =   3
         Left            =   600
         TabIndex        =   86
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提神剂"
         Height          =   180
         Index           =   2
         Left            =   600
         TabIndex        =   85
         Top             =   720
         Width           =   540
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "补血散"
         Height          =   180
         Index           =   1
         Left            =   600
         TabIndex        =   84
         Top             =   240
         Width           =   540
      End
      Begin VB.Image WoodPic 
         Height          =   480
         Index           =   3
         Left            =   120
         Picture         =   "StageForm.frx":0CC6
         Top             =   1680
         Width           =   480
      End
      Begin VB.Image WoodPic 
         Height          =   480
         Index           =   2
         Left            =   120
         Picture         =   "StageForm.frx":1108
         Top             =   1200
         Width           =   480
      End
      Begin VB.Image WoodPic 
         Height          =   480
         Index           =   1
         Left            =   120
         Picture         =   "StageForm.frx":154A
         Top             =   720
         Width           =   480
      End
      Begin VB.Image WoodPic 
         Height          =   480
         Index           =   0
         Left            =   120
         Picture         =   "StageForm.frx":198C
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame sf 
      BackColor       =   &H00FFFFFF&
      Caption         =   "当前状况"
      Height          =   1695
      Left            =   3840
      TabIndex        =   74
      Top             =   360
      Width           =   2175
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP"
         Height          =   180
         Left            =   120
         TabIndex        =   82
         Top             =   240
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MP"
         Height          =   180
         Left            =   120
         TabIndex        =   81
         Top             =   600
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AP"
         Height          =   180
         Left            =   120
         TabIndex        =   80
         Top             =   960
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DP"
         Height          =   180
         Left            =   120
         TabIndex        =   79
         Top             =   1320
         Width           =   180
      End
      Begin VB.Label mapz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   1080
         TabIndex        =   78
         Top             =   960
         Width           =   90
      End
      Begin VB.Label mdpz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   1080
         TabIndex        =   77
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label mzhpz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1080
         TabIndex        =   76
         Top             =   240
         Width           =   90
      End
      Begin VB.Label mzmpz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1080
         TabIndex        =   75
         Top             =   600
         Width           =   90
      End
      Begin VB.Shape mhp 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   240
         Width           =   1695
      End
      Begin VB.Shape mmp 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   600
         Width           =   1695
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   240
         Width           =   1695
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFFF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   600
         Width           =   1695
      End
      Begin VB.Shape mapp 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   960
         Width           =   1695
      End
      Begin VB.Shape mdpp 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   1320
         Width           =   1695
      End
   End
   Begin VB.Frame op 
      BackColor       =   &H00FFFFFF&
      Caption         =   "敌人状况"
      Height          =   1695
      Left            =   3840
      TabIndex        =   65
      Top             =   2160
      Visible         =   0   'False
      Width           =   2175
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DP"
         Height          =   180
         Left            =   120
         TabIndex        =   73
         Top             =   1320
         Width           =   180
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AP"
         Height          =   180
         Left            =   120
         TabIndex        =   72
         Top             =   960
         Width           =   180
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MP"
         Height          =   180
         Left            =   120
         TabIndex        =   71
         Top             =   600
         Width           =   180
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP"
         Height          =   180
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   180
      End
      Begin VB.Label dapz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   1080
         TabIndex        =   69
         Top             =   960
         Width           =   90
      End
      Begin VB.Label ddpz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   1080
         TabIndex        =   68
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label dzhpz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1080
         TabIndex        =   67
         Top             =   240
         Width           =   90
      End
      Begin VB.Label dzmpz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1080
         TabIndex        =   66
         Top             =   600
         Width           =   90
      End
      Begin VB.Shape dhp 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   240
         Width           =   1695
      End
      Begin VB.Shape dmp 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   600
         Width           =   1695
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FFFF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   600
         Width           =   1695
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   240
         Width           =   1695
      End
      Begin VB.Shape dapp 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   960
         Width           =   1695
      End
      Begin VB.Shape ddpp 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   1320
         Width           =   1695
      End
   End
   Begin VB.PictureBox Y1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   63
      Top             =   120
      Width           =   135
   End
   Begin VB.PictureBox Y1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   1200
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   62
      Top             =   120
      Width           =   135
   End
   Begin VB.PictureBox Y1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   61
      Top             =   120
      Width           =   135
   End
   Begin VB.PictureBox Y1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   2400
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   60
      Top             =   120
      Width           =   135
   End
   Begin VB.PictureBox Y1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   3000
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   59
      Top             =   120
      Width           =   135
   End
   Begin VB.PictureBox Y6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   58
      Top             =   3120
      Width           =   135
   End
   Begin VB.PictureBox Y2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   57
      Top             =   720
      Width           =   135
   End
   Begin VB.PictureBox Y2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   1200
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   56
      Top             =   720
      Width           =   135
   End
   Begin VB.PictureBox Y2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   55
      Top             =   720
      Width           =   135
   End
   Begin VB.PictureBox Y2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   2400
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   54
      Top             =   720
      Width           =   135
   End
   Begin VB.PictureBox Y2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   3000
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   53
      Top             =   720
      Width           =   135
   End
   Begin VB.PictureBox Y6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   1200
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   52
      Top             =   3120
      Width           =   135
   End
   Begin VB.PictureBox Y3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   51
      Top             =   1320
      Width           =   135
   End
   Begin VB.PictureBox Y3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   1200
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   50
      Top             =   1320
      Width           =   135
   End
   Begin VB.PictureBox Y3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   49
      Top             =   1320
      Width           =   135
   End
   Begin VB.PictureBox Y3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   2400
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   48
      Top             =   1320
      Width           =   135
   End
   Begin VB.PictureBox Y3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   3000
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   47
      Top             =   1320
      Width           =   135
   End
   Begin VB.PictureBox Y6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   46
      Top             =   3120
      Width           =   135
   End
   Begin VB.PictureBox Y4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   45
      Top             =   1920
      Width           =   135
   End
   Begin VB.PictureBox Y4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   1200
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   44
      Top             =   1920
      Width           =   135
   End
   Begin VB.PictureBox Y4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   43
      Top             =   1920
      Width           =   135
   End
   Begin VB.PictureBox Y4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   2400
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   42
      Top             =   1920
      Width           =   135
   End
   Begin VB.PictureBox Y4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   3000
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   41
      Top             =   1920
      Width           =   135
   End
   Begin VB.PictureBox Y6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   2400
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   40
      Top             =   3120
      Width           =   135
   End
   Begin VB.PictureBox Y5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   39
      Top             =   2520
      Width           =   135
   End
   Begin VB.PictureBox Y5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   1200
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   38
      Top             =   2520
      Width           =   135
   End
   Begin VB.PictureBox Y5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   37
      Top             =   2520
      Width           =   135
   End
   Begin VB.PictureBox Y5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   2400
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   36
      Top             =   2520
      Width           =   135
   End
   Begin VB.PictureBox Y5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   3000
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   35
      Top             =   2520
      Width           =   135
   End
   Begin VB.PictureBox Y6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   3000
      ScaleHeight     =   495
      ScaleWidth      =   135
      TabIndex        =   34
      Top             =   3120
      Width           =   135
   End
   Begin VB.PictureBox X1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   1
      Left            =   120
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   33
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox X1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   2
      Left            =   720
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   32
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox X1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   3
      Left            =   1320
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   31
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox X1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   4
      Left            =   1920
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   30
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox X1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   5
      Left            =   2520
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   29
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox X1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   6
      Left            =   3120
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   28
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox X2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   1
      Left            =   120
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   27
      Top             =   1200
      Width           =   495
   End
   Begin VB.PictureBox X2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   2
      Left            =   720
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   26
      Top             =   1200
      Width           =   495
   End
   Begin VB.PictureBox X2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   3
      Left            =   1320
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   25
      Top             =   1200
      Width           =   495
   End
   Begin VB.PictureBox X2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   4
      Left            =   1920
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   24
      Top             =   1200
      Width           =   495
   End
   Begin VB.PictureBox X2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   5
      Left            =   2520
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   23
      Top             =   1200
      Width           =   495
   End
   Begin VB.PictureBox X2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   6
      Left            =   3120
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   22
      Top             =   1200
      Width           =   495
   End
   Begin VB.PictureBox X3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   1
      Left            =   120
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   21
      Top             =   1800
      Width           =   495
   End
   Begin VB.PictureBox X3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   2
      Left            =   720
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   20
      Top             =   1800
      Width           =   495
   End
   Begin VB.PictureBox X3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   3
      Left            =   1320
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   19
      Top             =   1800
      Width           =   495
   End
   Begin VB.PictureBox X3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   4
      Left            =   1920
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   18
      Top             =   1800
      Width           =   495
   End
   Begin VB.PictureBox X3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   5
      Left            =   2520
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   17
      Top             =   1800
      Width           =   495
   End
   Begin VB.PictureBox X3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   6
      Left            =   3120
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   16
      Top             =   1800
      Width           =   495
   End
   Begin VB.PictureBox X4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   1
      Left            =   120
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   15
      Top             =   2400
      Width           =   495
   End
   Begin VB.PictureBox X4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   2
      Left            =   720
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   14
      Top             =   2400
      Width           =   495
   End
   Begin VB.PictureBox X4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   3
      Left            =   1320
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   13
      Top             =   2400
      Width           =   495
   End
   Begin VB.PictureBox X4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   4
      Left            =   1920
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   12
      Top             =   2400
      Width           =   495
   End
   Begin VB.PictureBox X4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   5
      Left            =   2520
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   11
      Top             =   2400
      Width           =   495
   End
   Begin VB.PictureBox X4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   6
      Left            =   3120
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   10
      Top             =   2400
      Width           =   495
   End
   Begin VB.PictureBox X5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   1
      Left            =   120
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox X5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   2
      Left            =   720
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox X5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   3
      Left            =   1320
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   7
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox X5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   4
      Left            =   1920
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox X5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   5
      Left            =   2520
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox X5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   6
      Left            =   3120
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   3000
      Width           =   495
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   2880
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "云端图标战士存档文件(*.cdv)|*.cdv"
   End
   Begin VB.Label zy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   180
      Left            =   4920
      TabIndex        =   109
      Top             =   120
      Width           =   90
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ","
      Height          =   180
      Left            =   4800
      TabIndex        =   108
      Top             =   120
      Width           =   90
   End
   Begin VB.Label zx 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   180
      Left            =   4680
      TabIndex        =   107
      Top             =   120
      Width           =   90
   End
   Begin VB.Image Imgsave 
      Height          =   480
      Left            =   240
      Picture         =   "StageForm.frx":1DCE
      Stretch         =   -1  'True
      ToolTipText     =   "储存进度"
      Top             =   5400
      Width           =   480
   End
   Begin VB.Image Self 
      Height          =   495
      Left            =   2280
      Picture         =   "StageForm.frx":2698
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Imgopen 
      Height          =   480
      Left            =   720
      Picture         =   "StageForm.frx":8EEA
      Stretch         =   -1  'True
      ToolTipText     =   "读取进度"
      Top             =   5400
      Width           =   480
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "当前坐标："
      Height          =   180
      Index           =   0
      Left            =   3840
      TabIndex        =   64
      Top             =   120
      Width           =   900
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   24
      Left            =   3000
      Top             =   3000
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   23
      Left            =   2400
      Top             =   3000
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   22
      Left            =   1800
      Top             =   3000
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   21
      Left            =   1200
      Top             =   3000
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   20
      Left            =   1200
      Top             =   2400
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   19
      Left            =   1800
      Top             =   2400
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   18
      Left            =   2400
      Top             =   1800
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   17
      Left            =   1800
      Top             =   1800
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   16
      Left            =   3000
      Top             =   1800
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   15
      Left            =   3000
      Top             =   2400
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   14
      Left            =   1800
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   13
      Left            =   2400
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   12
      Left            =   3000
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   11
      Left            =   600
      Top             =   1200
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   10
      Left            =   1200
      Top             =   1200
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   9
      Left            =   1800
      Top             =   1200
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   8
      Left            =   2400
      Top             =   1200
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   7
      Left            =   3000
      Top             =   1200
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   6
      Left            =   600
      Top             =   1800
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   5
      Left            =   600
      Top             =   2400
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   4
      Left            =   600
      Top             =   3000
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   3
      Left            =   1200
      Top             =   1800
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   2
      Left            =   1200
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   1
      Left            =   600
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Ss 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      DrawMode        =   1  'Blackness
      Height          =   135
      Index           =   0
      Left            =   2400
      Top             =   2400
      Width           =   135
   End
   Begin VB.Image P3 
      Height          =   495
      Index           =   6
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   495
   End
   Begin VB.Image P6 
      Height          =   495
      Index           =   6
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   495
   End
   Begin VB.Image P1 
      Height          =   495
      Index           =   1
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
   Begin VB.Image P1 
      Height          =   495
      Index           =   2
      Left            =   720
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
   Begin VB.Image P1 
      Height          =   495
      Index           =   3
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
   Begin VB.Image P1 
      Height          =   495
      Index           =   4
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
   Begin VB.Image P1 
      Height          =   495
      Index           =   5
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
   Begin VB.Image P1 
      Height          =   495
      Index           =   6
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
   Begin VB.Image P2 
      Height          =   495
      Index           =   1
      Left            =   120
      Stretch         =   -1  'True
      Top             =   720
      Width           =   495
   End
   Begin VB.Image P2 
      Height          =   495
      Index           =   2
      Left            =   720
      Stretch         =   -1  'True
      Top             =   720
      Width           =   495
   End
   Begin VB.Image P2 
      Height          =   495
      Index           =   3
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   720
      Width           =   495
   End
   Begin VB.Image P2 
      Height          =   495
      Index           =   4
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   720
      Width           =   495
   End
   Begin VB.Image P2 
      Height          =   495
      Index           =   5
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   720
      Width           =   495
   End
   Begin VB.Image P2 
      Height          =   495
      Index           =   6
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   720
      Width           =   495
   End
   Begin VB.Image P3 
      Height          =   495
      Index           =   1
      Left            =   120
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   495
   End
   Begin VB.Image P3 
      Height          =   495
      Index           =   2
      Left            =   720
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   495
   End
   Begin VB.Image P3 
      Height          =   495
      Index           =   3
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   495
   End
   Begin VB.Image P3 
      Height          =   495
      Index           =   4
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   495
   End
   Begin VB.Image P3 
      Height          =   495
      Index           =   5
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   495
   End
   Begin VB.Image P4 
      Height          =   495
      Index           =   1
      Left            =   120
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image P4 
      Height          =   495
      Index           =   2
      Left            =   720
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image P4 
      Height          =   495
      Index           =   3
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image P4 
      Height          =   495
      Index           =   4
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image P4 
      Height          =   495
      Index           =   5
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image P4 
      Height          =   495
      Index           =   6
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image P5 
      Height          =   495
      Index           =   1
      Left            =   120
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   495
   End
   Begin VB.Image P5 
      Height          =   495
      Index           =   2
      Left            =   720
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   495
   End
   Begin VB.Image P5 
      Height          =   495
      Index           =   3
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   495
   End
   Begin VB.Image P5 
      Height          =   495
      Index           =   4
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   495
   End
   Begin VB.Image P5 
      Height          =   495
      Index           =   5
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   495
   End
   Begin VB.Image P5 
      Height          =   495
      Index           =   6
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   495
   End
   Begin VB.Image P6 
      Height          =   495
      Index           =   1
      Left            =   120
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   495
   End
   Begin VB.Image P6 
      Height          =   495
      Index           =   2
      Left            =   720
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   495
   End
   Begin VB.Image P6 
      Height          =   495
      Index           =   3
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   495
   End
   Begin VB.Image P6 
      Height          =   495
      Index           =   4
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   495
   End
   Begin VB.Image P6 
      Height          =   495
      Index           =   5
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   17
      Left            =   2160
      Picture         =   "StageForm.frx":97B4
      Stretch         =   -1  'True
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   16
      Left            =   1680
      Picture         =   "StageForm.frx":9BF6
      Stretch         =   -1  'True
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   15
      Left            =   1200
      Picture         =   "StageForm.frx":A038
      Stretch         =   -1  'True
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   14
      Left            =   720
      Picture         =   "StageForm.frx":A47A
      Stretch         =   -1  'True
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Wood 
      Height          =   495
      Index           =   2
      Left            =   1200
      Picture         =   "StageForm.frx":A8BC
      Stretch         =   -1  'True
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Wood 
      Height          =   495
      Index           =   1
      Left            =   720
      Picture         =   "StageForm.frx":B186
      Stretch         =   -1  'True
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Wood 
      Height          =   495
      Index           =   0
      Left            =   240
      Picture         =   "StageForm.frx":BA50
      Stretch         =   -1  'True
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   13
      Left            =   240
      Picture         =   "StageForm.frx":C31A
      Stretch         =   -1  'True
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   12
      Left            =   2640
      Picture         =   "StageForm.frx":C75C
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   11
      Left            =   2160
      Picture         =   "StageForm.frx":CB9E
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   10
      Left            =   1680
      Picture         =   "StageForm.frx":CFE0
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   9
      Left            =   1200
      Picture         =   "StageForm.frx":D422
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   8
      Left            =   720
      Picture         =   "StageForm.frx":D864
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   7
      Left            =   240
      Picture         =   "StageForm.frx":F1E6
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   6
      Left            =   2640
      Picture         =   "StageForm.frx":F628
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   5
      Left            =   2160
      Picture         =   "StageForm.frx":F932
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   4
      Left            =   1680
      Picture         =   "StageForm.frx":FCBC
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   3
      Left            =   1200
      Picture         =   "StageForm.frx":1245E
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   2
      Left            =   720
      Picture         =   "StageForm.frx":12768
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   1
      Left            =   240
      Picture         =   "StageForm.frx":12A72
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image PicLib 
      Height          =   495
      Index           =   0
      Left            =   2640
      Stretch         =   -1  'True
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      Height          =   3495
      Left            =   120
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "StageForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sX As Byte, sY As Byte
Dim B1 As Boolean, b2 As Boolean, b3 As Boolean, B4 As Boolean, B5 As Boolean, B6 As Boolean



Private Sub Battle_Click()
Me.Hide
MainForm.Show
MainForm.mxhpz = mzhpz
MainForm.mxmpz = mzmpz
MainForm.mzhpz = mzhpz
MainForm.mzmpz = mzmpz
MainForm.mapz = mapz + Val(Ws(4).Caption) * 30
MainForm.mdpz = mdpz + Val(Ws(5).Caption) * 30
MainForm.dxhpz = dzhpz
MainForm.dxmpz = dzmpz
MainForm.dzhpz = dzhpz
MainForm.dzmpz = dzmpz
MainForm.dapz = dapz
MainForm.ddpz = ddpz
MainForm.dhp.Width = 1695
MainForm.dmp.Width = 1695
dAP = Val(MainForm.dapz.Caption)
dDP = Val(MainForm.ddpz.Caption)
dxHP = Val(MainForm.dxhpz.Caption)
dxMP = Val(MainForm.dxmpz.Caption)
dzHP = Val(MainForm.dzhpz.Caption)
dzMP = Val(MainForm.dzmpz.Caption)
MainForm.mhp.Width = 1695
MainForm.mmp.Width = 1695
Map = Val(MainForm.mapz.Caption)
mDP = Val(MainForm.mdpz.Caption)
mxHP = Val(MainForm.mxhpz.Caption)
mxMP = Val(MainForm.mxmpz.Caption)
mzHP = Val(MainForm.mzhpz.Caption)
mzMP = Val(MainForm.mzmpz.Caption)
MainForm.oper.Picture = PicLib(tMap(N).P(sX, sY).PicID).Picture
MainForm.op.Caption = tMap(N).P(sX, sY).Name
MainForm.Att.Enabled = True
WaitBack.Enabled = True
WaitBack.Tag = 0
YD.Visible = False
End Sub

Private Sub Flagy_Click()
Locker = False
Ws(3).Caption = Val(Ws(3).Caption) - 1
If Val(Ws(3).Caption) = 0 Then Flagy.Enabled = False
op.Visible = False
YD.Visible = False
End Sub

Private Sub GetSelfXY()
For i = 1 To 6
    For j = 1 To 6
        If tMap(N).P(i, j).isWhat = 4 Then
        sX = i
        sY = j
        zx = sX
        zy = sY
        Exit Sub
        End If
    Next j
Next i
End Sub

Private Sub Form_Load()
Load Y1(6)
Load Y2(6)
Load Y3(6)
Load Y4(6)
Load Y5(6)
Load Y6(6)
If StageOpenMode = 1 Then Call CheckOpen: Exit Sub
N = 1
GetSelfXY
Me.Caption = tMap(1).Head.StageName & " 第" & N & "关"
ReDrawMap (N)
Self.Picture = MainForm.Self.Picture
mzhpz.Caption = MainForm.mxhpz.Caption
mzmpz.Caption = MainForm.mxmpz.Caption
mapz.Caption = MainForm.mapz.Caption
mdpz.Caption = MainForm.mdpz.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub CheckOpen()
T = Split(tMap(256).P(1, 1).Name, "&")
For i = 0 To 5
Ws(i).Caption = T(i)
Next
N = Val(T(6))
MagicWL = Val(T(7))
If Ws(0).Enabled Then Use1.Enabled = True
If Ws(1).Enabled Then Use2.Enabled = True
mzhpz.Caption = T(9)
mzmpz.Caption = T(10)
mapz.Caption = T(11)
mdpz.Caption = T(12)
zx.Caption = T(13)
zy.Caption = T(14)
ReDrawMap (N)
sX = Val(zx)
sY = Val(zy)
Self.Left = P1(sY).Left
Select Case sX
Case 1
Self.Top = P1(sX).Top
Case 2
Self.Top = P2(sX).Top
Case 3
Self.Top = P3(sX).Top
Case 4
Self.Top = P4(sX).Top
Case 5
Self.Top = P5(sX).Top
Case 6
Self.Top = P6(sX).Top
End Select
Me.Caption = tMap(1).Head.StageName & " 第" & N & "关"
End Sub

Private Sub Imgopen_Click()
If Locker = True Then Exit Sub
On Error GoTo errH
Dim T() As String
cd.ShowOpen
Open cd.FileName For Binary As #1 Len = 32767
Get #1, 1, tMap
Close #1
CheckOpen
errH:
End Sub

Private Sub Imgsave_Click()
If Locker = True Then Exit Sub
On Error GoTo errH
Dim T(0 To 14) As String
For i = 0 To 5
T(i) = Ws(i).Caption
Next
T(6) = N
T(7) = MagicWL
T(9) = mzhpz.Caption
T(10) = mzmpz.Caption
T(11) = mapz.Caption
T(12) = mdpz.Caption
T(13) = zx.Caption
T(14) = zy.Caption
tMap(256).P(1, 1).Name = Join(T, "&")
cd.ShowSave
Open cd.FileName For Binary As #1 Len = 32767
Put #1, 1, tMap
Close #1
errH:
End Sub

Private Sub P1_Click(Index As Integer)
If Locker = True Then Exit Sub
B1 = (1 - sX = 1 And sY = Index) _
  Or (sX - 1 = 1 And sY = Index And X1(Index).Visible = False) _
  Or (sX = 1 And Index - sY = 1 And Y1(sY).Visible = False) _
  Or (sX = 1 And sY - Index = 1 And Y1(Index).Visible = False)

If B1 Then
Self.Left = P1(Index).Left
Self.Top = P1(Index).Top
sX = 1
sY = Index
Getpo
End If
End Sub

Private Sub P2_Click(Index As Integer)
If Locker = True Then Exit Sub
b2 = (2 - sX = 1 And sY = Index And X1(Index).Visible = False) _
  Or (sX - 2 = 1 And sY = Index And X2(Index).Visible = False) _
  Or (sX = 2 And Index - sY = 1 And Y2(sY).Visible = False) _
  Or (sX = 2 And sY - Index = 1 And Y2(Index).Visible = False)
'下上右左
If b2 Then
Self.Left = P2(Index).Left
Self.Top = P2(Index).Top
sX = 2
sY = Index
Getpo
End If
End Sub

Private Sub P3_Click(Index As Integer)
If Locker = True Then Exit Sub
b3 = (3 - sX = 1 And sY = Index And X2(Index).Visible = False) _
  Or (sX - 3 = 1 And sY = Index And X3(Index).Visible = False) _
  Or (sX = 3 And Index - sY = 1 And Y3(sY).Visible = False) _
  Or (sX = 3 And sY - Index = 1 And Y3(Index).Visible = False)
If b3 Then
Self.Left = P3(Index).Left
Self.Top = P3(Index).Top
sX = 3
sY = Index
Getpo
End If
End Sub

Private Sub P4_Click(Index As Integer)
If Locker = True Then Exit Sub
B4 = (4 - sX = 1 And sY = Index And X3(Index).Visible = False) _
  Or (sX - 4 = 1 And sY = Index And X4(Index).Visible = False) _
  Or (sX = 4 And Index - sY = 1 And Y4(sY).Visible = False) _
  Or (sX = 4 And sY - Index = 1 And Y4(Index).Visible = False)
If B4 Then
Self.Left = P4(Index).Left
Self.Top = P4(Index).Top
sX = 4
sY = Index
Getpo
End If
End Sub

Private Sub P5_Click(Index As Integer)
If Locker = True Then Exit Sub
B5 = (5 - sX = 1 And sY = Index And X4(Index).Visible = False) _
  Or (sX - 5 = 1 And sY = Index And X5(Index).Visible = False) _
  Or (sX = 5 And Index - sY = 1 And Y5(sY).Visible = False) _
  Or (sX = 5 And sY - Index = 1 And Y5(Index).Visible = False)
If B5 Then
Self.Left = P5(Index).Left
Self.Top = P5(Index).Top
sX = 5
sY = Index
Getpo
End If
End Sub

Private Sub P6_Click(Index As Integer)
If Locker = True Then Exit Sub
B6 = (6 - sX = 1 And sY = Index And X5(Index).Visible = False) _
  Or (sX - 6 = 1 And sY = Index) _
  Or (sX = 6 And Index - sY = 1 And Y6(sY).Visible = False) _
  Or (sX = 6 And sY - Index = 1 And Y6(Index).Visible = False)

If B6 Then
Self.Left = P6(Index).Left
Self.Top = P6(Index).Top
sX = 6
sY = Index
Getpo
End If
End Sub


Private Sub ReDrawMap(N As Byte)
On Error Resume Next
Randomize
For i = 1 To 6
X1(i).Visible = tMap(N).x(1, i)
Next
For i = 1 To 6
X2(i).Visible = tMap(N).x(2, i)
Next
For i = 1 To 6
X3(i).Visible = tMap(N).x(3, i)
Next
For i = 1 To 6
X4(i).Visible = tMap(N).x(4, i)
Next
For i = 1 To 6
X5(i).Visible = tMap(N).x(5, i)
Next

For i = 1 To 5
Y1(i).Visible = tMap(N).Y(1, i)
Next
For i = 1 To 5
Y2(i).Visible = tMap(N).Y(2, i)
Next
For i = 1 To 5
Y3(i).Visible = tMap(N).Y(3, i)
Next
For i = 1 To 5
Y4(i).Visible = tMap(N).Y(4, i)
Next
For i = 1 To 5
Y5(i).Visible = tMap(N).Y(5, i)
Next
For i = 1 To 5
Y6(i).Visible = tMap(N).Y(6, i)
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


Private Sub Getpo()
State.Caption = ""
Randomize
zx = sX
zy = sY
Select Case tMap(N).P(sX, sY).isWhat
Case 1
YD.Visible = True
Nameof = tMap(N).P(sX, sY).Name
Typeo = "普通敌人"
op.Visible = True
dzhpz = tMap(N).P(sX, sY).HP
dzmpz = tMap(N).P(sX, sY).MP
dapz = tMap(N).P(sX, sY).AP
ddpz = tMap(N).P(sX, sY).DP
If Val(Ws(3)) > 0 Then Flagy.Enabled = True
Locker = True
Case 2
YD.Visible = True
Nameof = tMap(N).P(sX, sY).Name
Typeo = "BOSS"
op.Visible = True
dzhpz = tMap(N).P(sX, sY).HP
dzmpz = tMap(N).P(sX, sY).MP
dapz = tMap(N).P(sX, sY).AP
ddpz = tMap(N).P(sX, sY).DP
Flagy.Enabled = False
Locker = True
Case 3
YD.Visible = False
op.Visible = False
If tMap(N).P(sX, sY).WoodNum > 0 Then
If tMap(N).P(sX, sY).WoodType = 0 And tMap(N).P(sX, sY).WoodNum > 0 Then tMap(N).P(sX, sY).WoodType = Fix(Rnd * 6) + 1
State = "你得到了 " & tMap(N).P(sX, sY).WoodNum & " 个" & lb(tMap(N).P(sX, sY).WoodType).Caption
Ws(tMap(N).P(sX, sY).WoodType - 1).Caption = Val(Ws(tMap(N).P(sX, sY).WoodType - 1).Caption) + tMap(N).P(sX, sY).WoodNum
tMap(N).P(sX, sY).WoodNum = 0
If Ws(0).Caption > 0 Then Use1.Enabled = True
If Ws(1).Caption > 0 Then Use2.Enabled = True
If Ws(3).Caption > 0 Then Flagy.Enabled = True
tMap(N).P(zx, zy).isWhat = 0
End If
End Select
End Sub

Private Sub State_Change()
ReDrawMap (N)
sX = Val(zx)
sY = Val(zy)
Self.Left = P1(sY).Left
Select Case sX
Case 1
Self.Top = P1(sX).Top
Case 2
Self.Top = P2(sX).Top
Case 3
Self.Top = P3(sX).Top
Case 4
Self.Top = P4(sX).Top
Case 5
Self.Top = P5(sX).Top
Case 6
Self.Top = P6(sX).Top
End Select
End Sub

Private Sub Use1_Click()
mzhpz.Caption = Int(Val(mzhpz.Caption) + 100 * (MagicWL / 8 + 0.8))
Ws(0).Caption = Val(Ws(0).Caption) - 1
If Ws(0).Caption = 0 Then Use1.Enabled = False
End Sub

Private Sub Use2_Click()
mzmpz.Caption = Int(Val(mzmpz.Caption) + 100 * (MagicWL / 8 + 0.8))
Ws(1).Caption = Val(Ws(1).Caption) - 1
If Ws(1).Caption = 0 Then Use2.Enabled = False
End Sub

Private Sub WaitBack_Timer()
If WaitBack.Tag <> 0 Then
WaitBack.Enabled = False
op.Visible = False
    If WaitBack.Tag = 2 Then
    N = N + 1
    Me.Caption = tMap(1).Head.StageName & " 第" & N & "关"
    ReDrawMap (N)
    GetSelfXY
    sX = Val(zx)
    sY = Val(zy)
    Self.Left = P1(sY).Left
    Select Case sX
    Case 1
    Self.Top = P1(sX).Top
    Case 2
    Self.Top = P2(sX).Top
    Case 3
    Self.Top = P3(sX).Top
    Case 4
    Self.Top = P4(sX).Top
    Case 5
    Self.Top = P5(sX).Top
    Case 6
    Self.Top = P6(sX).Top
    End Select
    
    End If
End If
End Sub
