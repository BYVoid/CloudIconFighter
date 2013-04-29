VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "云端图标战士"
   ClientHeight    =   5280
   ClientLeft      =   7785
   ClientTop       =   5610
   ClientWidth     =   4695
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   4695
   StartUpPosition =   2  '屏幕中心
   Begin VB.OptionButton mod2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "闯关模式"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
   Begin VB.OptionButton mod1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "试练模式"
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   0
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.Timer Cat 
      Enabled         =   0   'False
      Interval        =   1250
      Left            =   3240
      Top             =   360
   End
   Begin VB.Timer Att 
      Enabled         =   0   'False
      Interval        =   1250
      Left            =   1080
      Top             =   360
   End
   Begin VB.ListBox progress 
      Appearance      =   0  'Flat
      Height          =   1650
      ItemData        =   "MainForm.frx":0442
      Left            =   120
      List            =   "MainForm.frx":0444
      TabIndex        =   11
      Top             =   3000
      Width           =   4455
   End
   Begin VB.Frame op 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   2400
      TabIndex        =   6
      Top             =   840
      Width           =   2175
      Begin VB.Label dzmpz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1200
         TabIndex        =   31
         Top             =   600
         Width           =   90
      End
      Begin VB.Label dzhpz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1200
         TabIndex        =   30
         Top             =   240
         Width           =   90
      End
      Begin VB.Label labf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   5
         Left            =   1080
         TabIndex        =   26
         Top             =   600
         Width           =   90
      End
      Begin VB.Label labf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   4
         Left            =   1080
         TabIndex        =   25
         Top             =   240
         Width           =   90
      End
      Begin VB.Label ddpz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   1080
         TabIndex        =   22
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label dapz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   1080
         TabIndex        =   21
         Top             =   960
         Width           =   90
      End
      Begin VB.Label dxmpz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   720
         TabIndex        =   20
         Top             =   600
         Width           =   90
      End
      Begin VB.Label dxhpz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   720
         TabIndex        =   19
         Top             =   240
         Width           =   90
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   180
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MP"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   180
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AP"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DP"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   180
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
   End
   Begin VB.Frame sf 
      BackColor       =   &H00FFFFFF&
      Caption         =   "my"
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2175
      Begin VB.Label mzmpz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1200
         TabIndex        =   29
         Top             =   600
         Width           =   90
      End
      Begin VB.Label mzhpz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1200
         TabIndex        =   28
         Top             =   240
         Width           =   90
      End
      Begin VB.Label mxhpz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   720
         TabIndex        =   27
         Top             =   240
         Width           =   90
      End
      Begin VB.Label labf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   1080
         TabIndex        =   24
         Top             =   600
         Width           =   90
      End
      Begin VB.Label labf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   1080
         TabIndex        =   23
         Top             =   240
         Width           =   90
      End
      Begin VB.Label mdpz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   1080
         TabIndex        =   18
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label mapz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   1080
         TabIndex        =   17
         Top             =   960
         Width           =   90
      End
      Begin VB.Label mxmpz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   720
         TabIndex        =   16
         Top             =   600
         Width           =   90
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DP"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AP"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MP"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   180
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   600
         Width           =   1695
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton Begins 
      Caption         =   "开始/终止"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton dop 
      Caption         =   "选择对手"
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton mop 
      Caption         =   "选择战士"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "图标文件(*.ico)|*.ico|所有文件(*.*)|*.*"
   End
   Begin VB.Label Doing 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   34
      Top             =   2640
      Width           =   120
   End
   Begin VB.Shape Shape5 
      Height          =   255
      Left            =   120
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Image Mag1 
      Height          =   495
      Left            =   360
      Stretch         =   -1  'True
      Top             =   240
      Width           =   495
   End
   Begin VB.Image Mag2 
      Height          =   495
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   240
      Width           =   495
   End
   Begin VB.Label dsz 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3720
      TabIndex        =   33
      Top             =   240
      Width           =   90
   End
   Begin VB.Label msz 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   840
      TabIndex        =   32
      Top             =   240
      Width           =   90
   End
   Begin VB.Image MSP 
      Height          =   480
      Index           =   6
      Left            =   3840
      Picture         =   "MainForm.frx":0446
      ToolTipText     =   "双击打开关卡编辑器"
      Top             =   4680
      Width           =   480
   End
   Begin VB.Image MSP 
      Height          =   480
      Index           =   5
      Left            =   3240
      Picture         =   "MainForm.frx":6C98
      Top             =   4680
      Width           =   480
   End
   Begin VB.Image MSP 
      Height          =   480
      Index           =   4
      Left            =   2640
      Picture         =   "MainForm.frx":70DA
      Top             =   4680
      Width           =   480
   End
   Begin VB.Image MSP 
      Height          =   480
      Index           =   3
      Left            =   2040
      Picture         =   "MainForm.frx":751C
      Top             =   4680
      Width           =   480
   End
   Begin VB.Image MSP 
      Height          =   480
      Index           =   2
      Left            =   1440
      Picture         =   "MainForm.frx":795E
      Top             =   4680
      Width           =   480
   End
   Begin VB.Image MSP 
      Height          =   480
      Index           =   1
      Left            =   840
      Picture         =   "MainForm.frx":7C68
      Top             =   4680
      Width           =   480
   End
   Begin VB.Image MSP 
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   "MainForm.frx":80AA
      Top             =   4680
      Width           =   480
   End
   Begin VB.Image oper 
      Height          =   495
      Left            =   4080
      Stretch         =   -1  'True
      Top             =   240
      Width           =   495
   End
   Begin VB.Image self 
      Height          =   480
      Left            =   120
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Temp As String, Ltmp As Long, Fx As Long, Fm As Long, Fa As Long, Fd As Long
Dim T1 As Boolean, T2 As Boolean

Private Sub Att_Timer()
Mag1.Picture = LoadPicture
Mag2.Picture = LoadPicture
msz.Caption = ""
dsz.Caption = ""
Map = Val(mapz.Caption)
mDP = Val(mdpz.Caption)
mxHP = Val(mxhpz.Caption)
mxMP = Val(mxmpz.Caption)
mzHP = Val(mzhpz.Caption)
mzMP = Val(mzmpz.Caption)
dAP = Val(dapz.Caption)
dDP = Val(ddpz.Caption)
dxHP = Val(dxhpz.Caption)
dxMP = Val(dxmpz.Caption)
dzHP = Val(dzhpz.Caption)
dzMP = Val(dzmpz.Caption)
Att.Enabled = False
progress.AddItem Doing.Caption
If mxHP = 0 Then Over (False): Exit Sub
AttackForm.Show 1
End Sub

Private Sub Begins_Click()
Beginning = Not Beginning
mop.Enabled = Not Beginning
dop.Enabled = Not Beginning
mod1.Visible = False: mod2.Visible = False
If mod1.Value Then Att.Enabled = Beginning Else SelectForm.Show: Me.Hide: Begins.Visible = False
End Sub

Private Sub Cat_Timer()
Mag1.Picture = LoadPicture
Mag2.Picture = LoadPicture
msz.Caption = ""
dsz.Caption = ""
Map = Val(mapz.Caption)
mDP = Val(mdpz.Caption)
mxHP = Val(mxhpz.Caption)
mxMP = Val(mxmpz.Caption)
mzHP = Val(mzhpz.Caption)
mzMP = Val(mzmpz.Caption)
dAP = Val(dapz.Caption)
dDP = Val(ddpz.Caption)
dxHP = Val(dxhpz.Caption)
dxMP = Val(dxmpz.Caption)
dzHP = Val(dzhpz.Caption)
dzMP = Val(dzmpz.Caption)

Cat.Enabled = False
progress.AddItem Doing.Caption
If dxHP = 0 Then Over (True): Exit Sub
Computer_Attack

If Val(MainForm.mxhpz.Caption) < 0 Then MainForm.mxhpz.Caption = 0
MainForm.mhp.Width = Val(MainForm.mxhpz.Caption) / Val(MainForm.mzhpz.Caption) * 1695
MainForm.mmp.Width = Val(MainForm.mxmpz.Caption) / Val(MainForm.mzmpz.Caption) * 1695
MainForm.dhp.Width = Val(MainForm.dxhpz.Caption) / Val(MainForm.dzhpz.Caption) * 1695
MainForm.dmp.Width = Val(MainForm.dxmpz.Caption) / Val(MainForm.dzmpz.Caption) * 1695


Att.Enabled = True

End Sub

Private Sub Form_Load()
Randomize
MagicWL = 1.6
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then AboutForm.Show 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Timer Mod 5 = 0 Then AboutForm.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mod1_Click()
dop.Visible = True
dop.Enabled = True
If (Not T1) Or (Not T2) Then Begins.Enabled = False Else Begins.Enabled = True
BTmode = 0
End Sub

Private Sub mod2_Click()
dop.Visible = False
If mod2.Value And T1 Then Begins.Enabled = True
BTmode = 1
End Sub

Private Sub mop_Click()
On Error GoTo Errh1
cd.FileName = ""
cd.ShowOpen
self.Picture = LoadPicture(cd.FileName)
Fx = 0
Fm = 0
Fa = 0
Fd = 0
Dim q As New Scripting.FileSystemObject, w As TextStream
Set w = q.OpenTextFile(cd.FileName)
Temp = w.ReadAll
w.Close
Ltmp = Len(Temp)
For i = 1 To Ltmp
Fx = Fx + Asc(Mid(Temp, i, 1))
Next
For i = 1 To Ltmp Step 2
Fm = Fm + Asc(Mid(Temp, i, 1))
Next
For i = 1 To Ltmp Step 3
Fa = Fa + Asc(Mid(Temp, i, 1))
Next
For i = 1 To Ltmp Step 4
Fd = Fd + Asc(Mid(Temp, i, 1))
Next
mzhpz.Caption = Abs(Fx) Mod 253 + Int(Rnd * 30)
mzmpz.Caption = Abs(Fm) Mod 253 + Int(Rnd * 30)
mapz.Caption = Abs(Fa) Mod 47 + Int(Rnd * 30)
mdpz.Caption = Abs(Fd) Mod 19 + Int(Rnd * 30)
mxhpz.Caption = mzhpz.Caption
mxmpz.Caption = mzmpz.Caption
mhp.Width = 1695
mmp.Width = 1695
sf.Caption = cd.FileTitle
Map = Val(mapz.Caption)
mDP = Val(mdpz.Caption)
mxHP = Val(mxhpz.Caption)
mxMP = Val(mxmpz.Caption)
mzHP = Val(mzhpz.Caption)
mzMP = Val(mzmpz.Caption)
T1 = True
If T1 And T2 And mod1.Value Then Begins.Enabled = True
If mod2.Value Then Begins.Enabled = True
Errh1:
End Sub

Private Sub dop_Click()
On Error GoTo Errh2
cd.FileName = ""
cd.ShowOpen
oper.Picture = LoadPicture(cd.FileName)
Fx = 0
Fm = 0
Fa = 0
Fd = 0
Dim q As New Scripting.FileSystemObject, w As TextStream
Set w = q.OpenTextFile(cd.FileName)
Temp = w.ReadAll
w.Close
Ltmp = Len(Temp)
For i = 1 To Ltmp
Fx = Fx + Asc(Mid(Temp, i, 1))
Next
For i = 1 To Ltmp Step 2
Fm = Fm + Asc(Mid(Temp, i, 1))
Next
For i = 1 To Ltmp Step 3
Fa = Fa + Asc(Mid(Temp, i, 1))
Next
For i = 1 To Ltmp Step 4
Fd = Fd + Asc(Mid(Temp, i, 1))
Next
dzhpz.Caption = Abs(Fx) Mod 253 + Int(Rnd * 30)
dzmpz.Caption = Abs(Fm) Mod 253 + Int(Rnd * 30)
dapz.Caption = Abs(Fa) Mod 47 + Int(Rnd * 30)
ddpz.Caption = Abs(Fd) Mod 19 + Int(Rnd * 30)
dxhpz.Caption = dzhpz.Caption
dxmpz.Caption = dzmpz.Caption
dhp.Width = 1695
dmp.Width = 1695
op.Caption = cd.FileTitle
dAP = Val(dapz.Caption)
dDP = Val(ddpz.Caption)
dxHP = Val(dxhpz.Caption)
dxMP = Val(dxmpz.Caption)
dzHP = Val(dzhpz.Caption)
dzMP = Val(dzmpz.Caption)
T2 = True
If T1 And T2 And mod1.Value Then Begins.Enabled = True
Errh2:
End Sub

Private Sub Computer_Attack()
Dim m(1 To 6) As Long
Randomize

If dxMP >= Ms(4) Then
If dxHP / dzHP <= 0.25 Then
dsz.Caption = Ma4(dzMP) + Int(Rnd * 10)
dxmpz.Caption = Val(dxmpz.Caption) - Ms(4)
dxhpz.Caption = Val(dxhpz.Caption) + Val(dsz.Caption)
Mag2.Picture = MainForm.MSP(4).Picture
Doing.Caption = op.Caption & "用血气方刚" & " HP+" & dsz
Exit Sub
End If
End If

If dxMP >= Ms(6) Then
msz.Caption = Ma6(dzMP)
dxmpz.Caption = Val(dxmpz.Caption) - Ms(6)
mxhpz.Caption = Val(mxhpz.Caption) - Val(msz.Caption)
Mag1.Picture = MainForm.MSP(6).Picture
Doing.Caption = op.Caption & "用天命玄鸟攻击" & sf.Caption & " HP-" & msz
Exit Sub
End If

If dxHP >= Ms(5) Then
If dxMP < Ms(1) And Ma0(dAP, mDP) <= Ma1(dzMP) Then
dsz.Caption = Ma5(dzMP) + Int(Rnd * 10)
dxhpz.Caption = Val(dxhpz.Caption) - Ms(5)
dxmpz.Caption = Val(dxmpz.Caption) + Val(dsz.Caption)
Mag2.Picture = MainForm.MSP(5).Picture
Doing.Caption = op.Caption & "燃血补法" & "MP+" & dsz
Exit Sub
End If
End If

Select Case Timer Mod 10
Case 0 To 3
GoTo b2
Case 4 To 6
GoTo b3
End Select


If dxMP >= Ms(3) Then
msz.Caption = Ma3(dzMP) + Int(Rnd * 10)
dxmpz.Caption = Val(dxmpz.Caption) - Ms(3)
mxhpz.Caption = Val(mxhpz.Caption) - Val(msz.Caption)
Mag1.Picture = MainForm.MSP(3).Picture
Doing.Caption = op.Caption & "用醉卧云端攻击" & sf.Caption & " HP-" & msz
Exit Sub
End If

b2:
If dxMP >= Ms(2) Then
msz.Caption = Ma2(dzMP) + Int(Rnd * 10)
dxmpz.Caption = Val(dxmpz.Caption) - Ms(2)
mxhpz.Caption = Val(mxhpz.Caption) - Val(msz.Caption)
Mag1.Picture = MainForm.MSP(2).Picture
Doing.Caption = op.Caption & "用怒火之拳攻击" & sf.Caption & " HP-" & msz
Exit Sub
End If

b3:
If dxMP >= Ms(1) Then
msz.Caption = Ma1(dzMP) + Int(Rnd * 10)
dxmpz.Caption = Val(dxmpz.Caption) - Ms(1)
mxhpz.Caption = Val(mxhpz.Caption) - Val(msz.Caption)
Mag1.Picture = MainForm.MSP(1).Picture
Doing.Caption = op.Caption & "用雷霆霹雳攻击" & sf.Caption & " HP-" & msz
Exit Sub
End If

If Ma0(dAP, mDP) > 0 Then
msz.Caption = Ma0(dAP, mDP) + Int(Rnd * 10)
mxhpz.Caption = Val(mxhpz.Caption) - Val(msz.Caption)
Mag1.Picture = MainForm.MSP(0).Picture
Doing.Caption = op.Caption & "普通攻击" & sf.Caption & " HP-" & msz
Exit Sub
End If

Doing.Caption = op.Caption & "休息"
End Sub

Private Sub Over(Who As Boolean)
On Error Resume Next
Randomize
If Who Then
Doing.Caption = sf.Caption & "打败了" & op.Caption
Map = Val(mapz.Caption)
mDP = Val(mdpz.Caption)
mxHP = Val(mxhpz.Caption)
mxMP = Val(mxmpz.Caption)
mzHP = Val(mzhpz.Caption)
mzMP = Val(mzmpz.Caption)
mhp.Width = 1695
mmp.Width = 1695
MagicWL = MagicWL + 0.1
If BTmode = 0 Then
T2 = False
dop.Enabled = True
mzhpz.Caption = Int(Val(mzhpz.Caption) * 1.15)
mzmpz.Caption = Int(Val(mzmpz.Caption) * 1.15)
mapz.Caption = Int(Val(mapz.Caption) * 1.2)
mdpz.Caption = Int(Val(mdpz.Caption) * 1.2)
mxhpz.Caption = Int(mzhpz.Caption)
mxmpz.Caption = Int(mzmpz.Caption)
End If

If BTmode = 1 Then
    Locker = False
    Me.Hide
    StageForm.Show
    StageForm.mzhpz = Int(mxhpz * (dzhpz / mzhpz / 16 + 1))
    StageForm.mzmpz = Int(mxmpz * (dzmpz / mzmpz / 16 + 1))
    StageForm.mapz = Int(mapz * (dapz / mapz / 48 + 1)) - Val(StageForm.Ws(4).Caption) * 30
    StageForm.mdpz = Int(mdpz * (ddpz / mdpz / 48 + 1)) - Val(StageForm.Ws(5).Caption) * 30
    
    StageForm.WaitBack.Tag = 1
    
    If tMap(N).P(StageForm.zx.Caption, StageForm.zy.Caption).isWhat = 2 Then '进入下一关
    StageForm.WaitBack.Tag = 2
    End If
    
    tMap(N).P(StageForm.zx, StageForm.zy).isWhat = 0
    StageForm.State = "你获胜了"
    
    If tMap(N).P(StageForm.zx.Caption, StageForm.zy.Caption).WoodNum > 0 Then '如果有奖励物品
        If tMap(N).P(StageForm.zx.Caption, StageForm.zy.Caption).WoodType = 0 And tMap(N).P(StageForm.zx.Caption, StageForm.zy.Caption).WoodNum > 0 Then tMap(N).P(StageForm.zx.Caption, StageForm.zy.Caption).WoodType = Fix(Rnd * 6) + 1 '指定随机物品
        StageForm.State = "你得到了 " & tMap(N).P(StageForm.zx.Caption, StageForm.zy.Caption).WoodNum & " 个" & StageForm.lb(tMap(N).P(StageForm.zx.Caption, StageForm.zy.Caption).WoodType).Caption '显示标签
        StageForm.Ws(tMap(N).P(StageForm.zx.Caption, StageForm.zy.Caption).WoodType - 1).Caption = Val(StageForm.Ws(tMap(N).P(StageForm.zx.Caption, StageForm.zy.Caption).WoodType - 1).Caption) + tMap(N).P(StageForm.zx.Caption, StageForm.zy.Caption).WoodNum '增加物品
        '判断是否可以使用
        If StageForm.Ws(0).Caption > 0 Then StageForm.Use1.Enabled = True
        If StageForm.Ws(1).Caption > 0 Then StageForm.Use2.Enabled = True
        If StageForm.Ws(3).Caption > 0 Then StageForm.Flagy.Enabled = True
      
        StageForm.op.Visible = False
        tageForm.YD.Visible = False
    End If
    
    
End If



Else '被打败

Doing.Caption = op.Caption & "打败了" & sf.Caption
dAP = Val(dapz.Caption)
dDP = Val(ddpz.Caption)
dxHP = Val(dxhpz.Caption)
dxMP = Val(dxmpz.Caption)
dzHP = Val(dzhpz.Caption)
dzMP = Val(dzmpz.Caption)
dhp.Width = 1695
dmp.Width = 1695
If BTmode = 0 Then
T1 = False
mop.Enabled = True
dzhpz.Caption = Int(Val(dzhpz.Caption) * 1.15)
dzmpz.Caption = Int(Val(dzmpz.Caption) * 1.15)
dapz.Caption = Int(Val(dapz.Caption) * 1.2)
ddpz.Caption = Int(Val(ddpz.Caption) * 1.2)
dxhpz.Caption = Int(dzhpz.Caption)
dxmpz.Caption = Int(dzmpz.Caption)
MagicWL = 1.6
End If

If BTmode = 1 Then
Locker = False
Me.Hide
    If StageForm.Ws(2) > 0 Then
    MsgBox "复活灯显灵！", vbInformation
    StageForm.Ws(2) = StageForm.Ws(2) - 1
    StageForm.mzhpz = Int(mzhpz * 0.8)
    StageForm.mzmpz = Int(mzmpz * 0.8)
    StageForm.mapz = Int(mapz * 0.8)
    StageForm.mdpz = Int(mdpz * 0.8)
    StageForm.op.Visible = False
    StageForm.YD.Visible = False
    StageForm.State = "你复活了"
    StageForm.Show
    Else
    If MsgBox("你被打败了！而你没有复活灯！你要继续游戏吗？" & vbCrLf & "按“是”继续游戏" & vbCrLf & "按“否”退出游戏", vbExclamation + vbYesNo) = vbYes Then
    SelectForm.Show
    MagicWL = 1.6
    Else
    End
    End If
    End If
End If
End If
progress.AddItem Doing.Caption
Begins.Enabled = False
Beginning = False
End Sub

Private Sub MSP_DblClick(Index As Integer)
If Index = 6 Then StageEditor.Show
End Sub
