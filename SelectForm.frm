VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SelectForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "闯关模式"
   ClientHeight    =   3510
   ClientLeft      =   3600
   ClientTop       =   2565
   ClientWidth     =   6225
   Icon            =   "SelectForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6225
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton os 
      Caption         =   "打开存档文件"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox rems 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   1200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox maker 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox counts 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox names 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton Start 
      Caption         =   "开始"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton br 
      Caption         =   "打开关卡文件"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Path 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   3495
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "关卡注释："
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "关卡制作者："
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "关卡名称："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "关卡文件："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "SelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Closemode As Boolean
Dim Head As FileHead


Private Sub br_Click()
On Error GoTo errHandler1
cd.ShowOpen
Path.Text = cd.FileName
Open Path.Text For Binary As #1 Len = 32767
Get #1, 1, tMap
Close #1
names = tMap(1).Head.StageName
counts = tMap(1).Head.StagesCount
maker = tMap(1).Head.maker
rems = tMap(1).Head.Texts
Start.Enabled = True
StageOpenMode = 0
errHandler1:
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Closemode = False Then
MainForm.mod1.Visible = True
MainForm.mod2.Visible = True
Beginning = False
End If
End Sub


Private Sub os_Click()
On Error GoTo errHandler1
cd.ShowOpen
Path.Text = cd.FileName
Open Path.Text For Binary As #1 Len = 32767
Get #1, 1, tMap
Close #1
names = tMap(1).Head.StageName
counts = tMap(1).Head.StagesCount
maker = tMap(1).Head.maker
rems = tMap(1).Head.Texts
Start.Enabled = True
StageOpenMode = 1
errHandler1:
End Sub

Private Sub Start_Click()
Closemode = True
StageForm.Show
Unload Me
End Sub
