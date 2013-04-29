VERSION 5.00
Begin VB.Form AttackForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "攻击"
   ClientHeight    =   1560
   ClientLeft      =   8280
   ClientTop       =   8925
   ClientWidth     =   4005
   Icon            =   "AttackForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4005
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "休息"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   2880
      TabIndex        =   7
      Top             =   720
      Width           =   360
   End
   Begin VB.Image Nor 
      Height          =   375
      Left            =   2760
      Picture         =   "AttackForm.frx":0442
      Top             =   600
      Width           =   1155
   End
   Begin VB.Image Dis 
      Height          =   375
      Left            =   2760
      Picture         =   "AttackForm.frx":1B2C
      Top             =   1080
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "天命玄鸟"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   6
      Left            =   2880
      TabIndex        =   6
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "燃血补法"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   5
      Left            =   1560
      TabIndex        =   5
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "血气方刚"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   4
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "醉卧云端"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   3
      Left            =   1560
      TabIndex        =   3
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "怒火之拳"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "雷霆霹雳"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "普通攻击"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
   Begin VB.Image m 
      Height          =   375
      Index           =   6
      Left            =   2760
      Top             =   120
      Width           =   1155
   End
   Begin VB.Image m 
      Height          =   375
      Index           =   5
      Left            =   1440
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Image m 
      Height          =   375
      Index           =   4
      Left            =   1440
      Top             =   600
      Width           =   1155
   End
   Begin VB.Image m 
      Height          =   375
      Index           =   3
      Left            =   1440
      Top             =   120
      Width           =   1155
   End
   Begin VB.Image m 
      Height          =   375
      Index           =   2
      Left            =   120
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Image m 
      Height          =   375
      Index           =   1
      Left            =   120
      Top             =   600
      Width           =   1155
   End
   Begin VB.Image m 
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   1155
   End
   Begin VB.Image Clk 
      Height          =   375
      Left            =   2760
      Picture         =   "AttackForm.frx":3216
      Top             =   840
      Visible         =   0   'False
      Width           =   1155
   End
End
Attribute VB_Name = "AttackForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Wds As Boolean


Private Sub Form_Load()
Randomize
Label1(0).ToolTipText = "使用武器攻击对方，给予对方大约" & Ma0(Map, dDP) & "HP的伤害。"
Label1(1).ToolTipText = "以强大的霹雳电击对方，给予对方大约" & Ma1(mzMP) & "的HP伤害。消耗MP" & ms(1) '20
Label1(2).ToolTipText = "召唤天之怒火，给予对方大约" & Ma2(mzMP) & "的HP伤害。消耗MP" & ms(2) '30
Label1(3).ToolTipText = "醉卧云端笑人间，给予对方大约" & Ma3(mzMP) & "的HP伤害。消耗MP" & ms(3) '50
Label1(4).ToolTipText = "命悬一线，血气方刚，补充大约" & Ma4(mzMP) & "的HP。消耗MP" & ms(4) '30
Label1(5).ToolTipText = "消耗HP以补充MP。补充大约" & Ma5(mzMP) & "的MP。消耗HP" & ms(5) '30
Label1(6).ToolTipText = "天命玄鸟，降而生商。使对方受到大约" & Ma6(mzMP) & "的HP伤害。消耗MP" & ms(6) '500
For i = 0 To 6
m(i).Picture = Dis.Picture
m(i).ToolTipText = Label1(i).ToolTipText
Label1(i).Enabled = False
m(i).Enabled = False
Next i
If Ma0(Map, dDP) > 0 Then m(0).Picture = Nor.Picture: Label1(0).Enabled = True: m(0).Enabled = True
If mxMP >= ms(1) Then m(1).Picture = Nor.Picture: Label1(1).Enabled = True: m(1).Enabled = True
If mxMP >= ms(2) Then m(2).Picture = Nor.Picture: Label1(2).Enabled = True: m(2).Enabled = True
If mxMP >= ms(3) Then m(3).Picture = Nor.Picture: Label1(3).Enabled = True: m(3).Enabled = True
If mxMP >= ms(4) Then m(4).Picture = Nor.Picture: Label1(4).Enabled = True: m(4).Enabled = True
If mxHP >= ms(5) Then m(5).Picture = Nor.Picture: Label1(5).Enabled = True: m(5).Enabled = True
If mxMP >= ms(6) Then m(6).Picture = Nor.Picture: Label1(6).Enabled = True: m(6).Enabled = True
Wds = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not Wds Then Cancel = 1
End Sub

Private Sub Label1_Click(Index As Integer)
m_Click (Index)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
m(Index).Picture = Clk.Picture
End Sub

Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
m(Index).Picture = Clk.Picture
End Sub

Private Sub Label2_Click()
Nor_Click
End Sub

Private Sub m_Click(Index As Integer)
Select Case Index
Case 0
MainForm.dsz = Ma0(Map, dDP) + Int(Rnd * 10)
MainForm.Mag2.Picture = MainForm.msp(0).Picture
MainForm.Doing.Caption = MainForm.sf.Caption & "普通攻击" & MainForm.op.Caption & " HP-" & MainForm.dsz
Case 1
MainForm.dsz = Ma1(mzMP) + Int(Rnd * 10)
MainForm.mxmpz.Caption = MainForm.mxmpz.Caption - ms(1)
MainForm.Mag2.Picture = MainForm.msp(1).Picture
MainForm.Doing.Caption = MainForm.sf.Caption & "用雷霆霹雳攻击" & MainForm.op.Caption & " HP-" & MainForm.dsz
Case 2
MainForm.dsz = Ma2(mzMP) + Int(Rnd * 10)
MainForm.mxmpz.Caption = MainForm.mxmpz.Caption - ms(2)
MainForm.Mag2.Picture = MainForm.msp(2).Picture
MainForm.Doing.Caption = MainForm.sf.Caption & "用怒火之拳攻击" & MainForm.op.Caption & " HP-" & MainForm.dsz
Case 3
MainForm.dsz = Ma3(mzMP) + Int(Rnd * 10)
MainForm.mxmpz.Caption = MainForm.mxmpz.Caption - ms(3)
MainForm.Mag2.Picture = MainForm.msp(3).Picture
MainForm.Doing.Caption = MainForm.sf.Caption & "用醉卧云端攻击" & MainForm.op.Caption & " HP-" & MainForm.dsz
Case 4
MainForm.msz = Ma4(mzMP) + Int(Rnd * 10)
MainForm.mxmpz.Caption = MainForm.mxmpz.Caption - ms(4)
MainForm.mxhpz.Caption = Val(MainForm.mxhpz.Caption) + Val(MainForm.msz)
MainForm.Mag1.Picture = MainForm.msp(4).Picture
MainForm.Doing.Caption = MainForm.sf.Caption & "用血气方刚" & " HP+" & MainForm.msz
Case 5
MainForm.msz = Ma5(mzMP) + Int(Rnd * 10)
MainForm.mxhpz.Caption = MainForm.mxhpz.Caption - ms(5)
MainForm.mxmpz.Caption = Val(MainForm.mxmpz.Caption) + Val(MainForm.msz)
MainForm.Mag1.Picture = MainForm.msp(5).Picture
MainForm.Doing.Caption = MainForm.sf.Caption & "燃血补法" & "MP+" & MainForm.msz
Case 6
MainForm.dsz = Ma6(mzMP) + Int(Rnd * 10)
MainForm.mxmpz.Caption = MainForm.mxmpz.Caption - ms(6)
MainForm.Mag2.Picture = MainForm.msp(6).Picture
MainForm.Doing.Caption = MainForm.sf.Caption & "用天命玄鸟攻击" & MainForm.op.Caption & " HP-" & MainForm.dsz
End Select
MainForm.dxhpz.Caption = Val(MainForm.dxhpz.Caption) - Val(MainForm.dsz)
If Val(MainForm.dxhpz.Caption) < 0 Then MainForm.dxhpz.Caption = 0
MainForm.mhp.Width = Val(MainForm.mxhpz.Caption) / Val(MainForm.mzhpz.Caption) * 1695
MainForm.mmp.Width = Val(MainForm.mxmpz.Caption) / Val(MainForm.mzmpz.Caption) * 1695
MainForm.dhp.Width = Val(MainForm.dxhpz.Caption) / Val(MainForm.dzhpz.Caption) * 1695
MainForm.dmp.Width = Val(MainForm.dxmpz.Caption) / Val(MainForm.dzmpz.Caption) * 1695
MainForm.Cat.Enabled = True
Wds = True
Unload Me
End Sub

Private Sub m_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
m(Index).Picture = Clk.Picture
End Sub

Private Sub m_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
m(Index).Picture = Nor.Picture
End Sub

Private Sub Nor_Click()
MainForm.Doing.Caption = MainForm.sf.Caption & "休息"
Wds = True
MainForm.Cat.Enabled = True
Unload Me
End Sub
