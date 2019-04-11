Attribute VB_Name = "S_Settings"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Public skeletal1 As Range
Public skeletal2 As Range
Public skeletal3 As Range
Public skeletal4 As Range
Public skeletal5 As Range
Public skeletal6 As Range
Public skeletal7 As Range

Public fantomas As Range
Public fantomas2 As Range
Public fantomas3 As Range

Public anim1 As Range
Public anim2 As Range

Public pahar1 As Range
Public pahar2 As Range
Public pahar3 As Range

Public para1 As Range
Public para2 As Range
Public para3 As Range

Public beer1 As Range
Public beer2 As Range
Public beer3 As Range

Public cola1 As Range
Public cola2 As Range
Public cola3 As Range

Public cig1 As Range
Public cig2 As Range
Public cig3 As Range

Public ochi1 As Range
Public ochi2 As Range

Public mobs(1 To 15) As Range

Public anim As Range
Public waitfortalk As Integer
Public waitformouth As Integer
Public waitforeyes As Integer
Public waitformove As Integer
Public noMobs As Integer
Public u As Integer

Sub Settings()
Set skeletal1 = Sheets(2).Range("M4:AE20")
Set skeletal2 = Sheets(2).Range("M22:AE39")
Set skeletal3 = Sheets(2).Range("M41:AE59")
Set skeletal4 = Sheets(2).Range("M61:AE80")
Set skeletal5 = Sheets(2).Range("M84:AE104")
Set skeletal6 = Sheets(2).Range("M108:AE129")
Set skeletal7 = Sheets(2).Range("M133:AE155")

Set fantomas = Sheets(2).Range("AQ4:BE20")
Set fantomas2 = Sheets(2).Range("AQ21:BE36")
Set fantomas3 = Sheets(2).Range("AQ38:BE54")

Set pahar1 = Sheets(2).Range("BK4:BY17")
Set pahar2 = Sheets(2).Range("BK21:BY34")
Set pahar3 = Sheets(2).Range("BK38:BY51")
Set mobs(13) = pahar1
Set mobs(14) = pahar2
Set mobs(15) = pahar3

Set para1 = Sheets(2).Range("CI4:CT21")
Set para2 = Sheets(2).Range("CI23:CT40")
Set para3 = Sheets(2).Range("CI42:CT59")
Set mobs(10) = para1
Set mobs(11) = para2
Set mobs(12) = para3

Set cola1 = Sheets(2).Range("DU4:EC18")
Set cola2 = Sheets(2).Range("DU23:EC37")
Set cola3 = Sheets(2).Range("DU43:EC57")
Set mobs(1) = cola1
Set mobs(2) = cola2
Set mobs(3) = cola3

'tigarea
Set mobs(4) = Sheets(2).Range("EI4:FA22")
Set mobs(5) = Sheets(2).Range("EI23:FA41")
Set mobs(6) = Sheets(2).Range("EI44:FA62")

'gaura neagra
Set mobs(7) = Sheets(2).Range("FK4:GE24")
Set mobs(8) = Sheets(2).Range("FK27:GE47")
Set mobs(9) = Sheets(2).Range("FK49:GE69")

u = Sheets(1).Range("A2").Value

Set ochi1 = Sheets(1).Range("N12:O12")
Set ochi2 = Sheets(1).Range("W12:X12")
'Set anim1 = Sheets(1).Cells(Range("C3").Offset(0, 0).Value, Range("D3").Offset(0, 0).Value)
'Set anim2 = Sheets(1).Cells(Range("C3").Offset(0, 2).Value, Range("D3").Offset(0, 2).Value)
waitfortalk = 60
waitformouth = 80
waitforeyes = 2
waitformove = 4
End Sub


Sub CreateMobs()

noMobs = InputBox("How many mobs?")

For i = 1 To noMobs
'Aici setam celulele de citire pentru coordonatele de plecare a fiecarui mob
Sheets(1).Range("C3").Offset(0, 2 * (i - 1)).Value = 31
Sheets(1).Range("D3").Offset(0, 2 * (i - 1)).Value = i * 47
'Aici setam celulele de citire pentru coordonatele precedente, necesare la animarea mobului
Sheets(1).Range("E2").Offset(0, 2 * (i - 1)).Value = 31
Sheets(1).Range("F2").Offset(0, 2 * (i - 1)).Value = i * 47
Next i

Sheets(1).Range("A2").Value = noMobs

End Sub

