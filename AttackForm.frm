VERSION 5.00
Begin VB.Form AttackForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����"
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
      Caption         =   "��Ϣ"
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
      Caption         =   "��������"
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
      Caption         =   "ȼѪ����"
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
      Caption         =   "Ѫ������"
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
      Caption         =   "�����ƶ�"
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
      Caption         =   "ŭ��֮ȭ"
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
      Caption         =   "��������"
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
      Caption         =   "��ͨ����"
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
Label1(0).ToolTipText = "ʹ�����������Է�������Է���Լ" & Ma0(Map, dDP) & "HP���˺���"
Label1(1).ToolTipText = "��ǿ�����������Է�������Է���Լ" & Ma1(mzMP) & "��HP�˺�������MP" & ms(1) '20
Label1(2).ToolTipText = "�ٻ���֮ŭ�𣬸���Է���Լ" & Ma2(mzMP) & "��HP�˺�������MP" & ms(2) '30
Label1(3).ToolTipText = "�����ƶ�Ц�˼䣬����Է���Լ" & Ma3(mzMP) & "��HP�˺�������MP" & ms(3) '50
Label1(4).ToolTipText = "����һ�ߣ�Ѫ�����գ������Լ" & Ma4(mzMP) & "��HP������MP" & ms(4) '30
Label1(5).ToolTipText = "����HP�Բ���MP�������Լ" & Ma5(mzMP) & "��MP������HP" & ms(5) '30
Label1(6).ToolTipText = "�������񣬽������̡�ʹ�Է��ܵ���Լ" & Ma6(mzMP) & "��HP�˺�������MP" & ms(6) '500
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
MainForm.Doing.Caption = MainForm.sf.Caption & "��ͨ����" & MainForm.op.Caption & " HP-" & MainForm.dsz
Case 1
MainForm.dsz = Ma1(mzMP) + Int(Rnd * 10)
MainForm.mxmpz.Caption = MainForm.mxmpz.Caption - ms(1)
MainForm.Mag2.Picture = MainForm.msp(1).Picture
MainForm.Doing.Caption = MainForm.sf.Caption & "��������������" & MainForm.op.Caption & " HP-" & MainForm.dsz
Case 2
MainForm.dsz = Ma2(mzMP) + Int(Rnd * 10)
MainForm.mxmpz.Caption = MainForm.mxmpz.Caption - ms(2)
MainForm.Mag2.Picture = MainForm.msp(2).Picture
MainForm.Doing.Caption = MainForm.sf.Caption & "��ŭ��֮ȭ����" & MainForm.op.Caption & " HP-" & MainForm.dsz
Case 3
MainForm.dsz = Ma3(mzMP) + Int(Rnd * 10)
MainForm.mxmpz.Caption = MainForm.mxmpz.Caption - ms(3)
MainForm.Mag2.Picture = MainForm.msp(3).Picture
MainForm.Doing.Caption = MainForm.sf.Caption & "�������ƶ˹���" & MainForm.op.Caption & " HP-" & MainForm.dsz
Case 4
MainForm.msz = Ma4(mzMP) + Int(Rnd * 10)
MainForm.mxmpz.Caption = MainForm.mxmpz.Caption - ms(4)
MainForm.mxhpz.Caption = Val(MainForm.mxhpz.Caption) + Val(MainForm.msz)
MainForm.Mag1.Picture = MainForm.msp(4).Picture
MainForm.Doing.Caption = MainForm.sf.Caption & "��Ѫ������" & " HP+" & MainForm.msz
Case 5
MainForm.msz = Ma5(mzMP) + Int(Rnd * 10)
MainForm.mxhpz.Caption = MainForm.mxhpz.Caption - ms(5)
MainForm.mxmpz.Caption = Val(MainForm.mxmpz.Caption) + Val(MainForm.msz)
MainForm.Mag1.Picture = MainForm.msp(5).Picture
MainForm.Doing.Caption = MainForm.sf.Caption & "ȼѪ����" & "MP+" & MainForm.msz
Case 6
MainForm.dsz = Ma6(mzMP) + Int(Rnd * 10)
MainForm.mxmpz.Caption = MainForm.mxmpz.Caption - ms(6)
MainForm.Mag2.Picture = MainForm.msp(6).Picture
MainForm.Doing.Caption = MainForm.sf.Caption & "���������񹥻�" & MainForm.op.Caption & " HP-" & MainForm.dsz
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
MainForm.Doing.Caption = MainForm.sf.Caption & "��Ϣ"
Wds = True
MainForm.Cat.Enabled = True
Unload Me
End Sub
