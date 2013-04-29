VERSION 5.00
Begin VB.Form AboutForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "关于 云端图标战士"
   ClientHeight    =   1905
   ClientLeft      =   8115
   ClientTop       =   6690
   ClientWidth     =   4320
   Icon            =   "AboutForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4320
   Begin VB.CommandButton OK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "免费游戏 欢迎自由传播"
      Height          =   180
      Left            =   840
      TabIndex        =   5
      Top             =   840
      Width           =   1890
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email：gjb66@yeah.net"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Width           =   1890
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "作者：云端"
      Height          =   180
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Ver 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ver"
      Height          =   180
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "云端图标战士"
      Height          =   180
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "AboutForm.frx":0442
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long


Private Sub Form_Load()
Ver.Caption = "Ver" & App.Major & "." & App.Minor & " Build " & App.Revision
End Sub

Private Sub Label3_Click()
Start "mailto:gjb66@yeah.net"
End Sub

Private Sub OK_Click()
Unload Me
End Sub

Private Function Start(Path As String) As Long
Start = ShellExecute(GetDesktopWindow(), "", Path, "", "", 1)
End Function
