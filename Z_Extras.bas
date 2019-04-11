Attribute VB_Name = "Z_Extras"
''As putea sa abstractizez Move intr-o functie parametrata cu pana la 3 cadre de animatie si pozitia de plecare
'Sub PearMove()
'Dim randnum As Integer
'Dim prevnum As Integer
'
''Pozitia de plecare
'Set anim = Sheets(1).Cells(Range("C1").Value, Range("D1").Value)
'prevnum = Range("E1").Value
'
''Stergerea cadrului anterior
'Select Case prevnum
'    Case 1
'        anim.Offset(1, 0).Resize(para1.Rows.Count, para1.Columns.Count).Interior.Color = rgbWhite
'    Case 2
'        anim.Offset(-1, 0).Resize(para1.Rows.Count, para1.Columns.Count).Interior.Color = rgbWhite
'    Case 3
'        anim.Offset(0, -1).Resize(para1.Rows.Count, para1.Columns.Count).Interior.Color = rgbWhite
'    Case 4
'        anim.Offset(0, 1).Resize(para1.Rows.Count, para1.Columns.Count).Interior.Color = rgbWhite
'End Select
'
''Fixarea unei directii aleatoare
''TODO SETAREA TRAIECTORIEI PAREI IN FUNCTIE DE POZITIA FANTOMEI IN 4 CAZURI POSIBILE - CELE 4 CADRANE INCONJURATOARE ALE PAREI
'Randomize
'randnum = Int(4 * Rnd) + 1
'
''Desenarea cadrelor curente
'Select Case randnum
'    Case 1
'        Call moveloop(-1, 0, para1)
'    Case 2
'        Call moveloop(1, 0, para2)
'    Case 3
'        Call moveloop(0, 1, para3)
'    Case 4
'        Call moveloop(0, -1, para1)
'End Select
'
''Scrierea coordonatelor pentru cadrul urmator
'Range("C1").Value = anim.Row
'Range("D1").Value = anim.Column
'Range("E1").Value = randnum
'
'End Sub


'Sub skelanim2()
'Call Settings
'
'skeletal4.Copy anim
'Sleep (waitformouth)
'skeletal5.Copy anim
'Sleep (waitformouth)
'skeletal6.Copy anim
'Sleep (waitformouth)
'skeletal7.Copy anim
'Sleep (waitformouth)
'For j = 1 To 255
'ochi1.Interior.Color = RGB(255, 255 - j, 0)
'ochi2.Interior.Color = RGB(255, 255 - j, 0)
'Sleep (waitforeyes)
'Next j
'
'End Sub
'
'Sub spriteanim()
'Call Settings
'
'For i = 1 To 4
'
'skeletal1.Copy anim
'Sleep (waitfortalk)
'skeletal2.Copy anim
'Sleep (waitfortalk)
'skeletal3.Copy anim
'Sleep (waitfortalk)
'Next i
'
'MsgBox "8===============D"
'Call skelanim2
'Sleep (800)
'anim.Resize(skeletal7.Rows.Count, skeletal7.Columns.Count).Clear
'
'End Sub
