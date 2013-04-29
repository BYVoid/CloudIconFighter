Attribute VB_Name = "Functions"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Map As Long, mDP As Long, mxHP As Long, mxMP As Long, mzHP As Long, mzMP As Long
Public dAP As Long, dDP As Long, dxHP As Long, dxMP As Long, dzHP As Long, dzMP As Long
Public Beginning As Boolean, BTmode As Byte, Locker As Boolean
Public N As Byte

Public Type Unit
Name As String
HP As Long
MP As Long
AP As Long
DP As Long
PicID As Integer
OtherPic As String
isWhat As Byte '0 Space  1 Standard  2 NextStage  3 Woods  4 Home
WoodType As Byte
WoodNum As Integer
End Type

Public Type FileHead
StageName As String
StagesCount As Byte
maker As String
Texts As String
End Type

Public Type Map
x(5, 6) As Boolean
Y(6, 5) As Boolean
P(1 To 6, 1 To 6) As Unit
Head As FileHead
End Type
Public MagicWL As Single, StageOpenMode As Byte
Public tMap(1 To 256) As Map

Public Function Ma0(MMap As Long, DDDP As Long) As Long 'ÆÕ¹¥
Ma0 = MMap - DDDP
End Function

Public Function Ma1(MMZMP As Long) As Long '20 Åùö¨
Ma1 = (MMZMP Mod 15 + 14) * MagicWL + (MMZMP Mod 3) * 2
End Function

Public Function Ma2(MMZMP As Long) As Long '30 ÁÒ»ð
Ma2 = (MMZMP Mod 10 + 21) * (MagicWL + 0.3) + (MMZMP Mod 3) * 2
End Function

Public Function Ma3(MMZMP As Long) As Long '40 ÔÆ¶Ë
Ma3 = (MMZMP Mod 23 + 22) * (MagicWL + 0.5) + (MMZMP Mod 3) * 2
End Function

Public Function Ma4(MMZMP As Long) As Long '30 ÑªÆø
Ma4 = (MMZMP Mod 20 + 23) * (MagicWL + 0.52) + (MMZMP Mod 3) * 2
End Function

Public Function Ma5(MMZMP As Long) As Long '30 È¼Ñª
Ma5 = (MMZMP Mod 24 + 21) * (MagicWL + 0.52) + (MMZMP Mod 3) * 2
End Function

Public Function Ma6(MMZMP As Long) As Long '500 ÐþÄñ
Ma6 = (MMZMP Mod 5 + 26) + (MagicWL * 1.25 + 2) * 97.7 + (MMZMP Mod 30) * 2
End Function

Public Function Ms(Magicnumber As Byte) As Long
Select Case Magicnumber
Case 1 To 3
Ms = Int((MagicWL * 8 + 7.2) * (0.5 + 0.5 * Magicnumber + MagicWL / 10))
Case 4
Ms = Int((MagicWL * 8 + 7.2) * (0.5 + 0.5 * 3 + MagicWL / 10))
Case 5
Ms = Int((MagicWL * 8 + 7.2) * (0.5 + 0.5 * 3.1 + MagicWL / 11))
Case 6
Ms = Int(((MagicWL + 0.4) / 6 * (Int(MagicWL * 1000 + N) Mod 2 + 6.325) * 35) + (MagicWL + 2) * 100)
End Select
End Function

Public Function Start(Path As String) As Long
Start = ShellExecute(GetDesktopWindow(), "", Path, "", "", 1)
End Function

