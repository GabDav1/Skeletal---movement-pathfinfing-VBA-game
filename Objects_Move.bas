Attribute VB_Name = "Objects_Move"
Sub SkelMove(nmbs As Integer, toPointx As Integer, toPointy As Integer, iterator As Variant)

'SETAREA TRAIECTORIEI PAREI IN FUNCTIE DE POZITIA FANTOMEI IN 4 CAZURI POSIBILE - CELE 4 CADRANE INCONJURATOARE ALE PAREI
xs = toPointx
ys = toPointy

'Pozitia de plecare
Set anim = Sheets(1).Cells(Range("C3").Offset(0, 2 * nmbs).Value, Range("D3").Offset(0, 2 * nmbs).Value)

'Stabilirea directiei in functie de coordonatele celulei destinatie
If xs > anim.Row And ys < anim.Column Then
    randnum = 1
ElseIf xs > anim.Row And ys > anim.Column Then
    randnum = 2
ElseIf xs < anim.Row And ys < anim.Column Then
    randnum = 3
ElseIf xs < anim.Row And ys > anim.Column Then
    randnum = 4
Else:
    Randomize
    randnum = Int(Rnd * 4) + 1
End If

'Stabilirea cadrului actual de animatie
If iterator Mod 3 = 0 Then
    Set tempAnim = mobs(1 + nmbs * 3)
ElseIf iterator Mod 3 = 1 Then
    Set tempAnim = mobs(2 + nmbs * 3)
Else: Set tempAnim = mobs(3 + nmbs * 3)
End If

'Desenarea cadrelor in functie de directia stabilita
Select Case randnum
    Case 1
        Call moveloop(nmbs, 1, -1, tempAnim)
    Case 2
        Call moveloop(nmbs, 1, 1, tempAnim)
    Case 3
        Call moveloop(nmbs, -1, -1, tempAnim)
    Case 4
        Call moveloop(nmbs, -1, 1, tempAnim)
End Select

'Scrierea coordonatelor pentru cadrul urmator
Range("C3").Offset(0, 2 * nmbs).Value = anim.Row
Range("D3").Offset(0, 2 * nmbs).Value = anim.Column

End Sub

Sub moveloop(nmbs As Integer, x As Integer, y As Integer, z As Range)

prevnumX = Sheets(1).Range("E2").Offset(0, 2 * nmbs).Value
prevnumY = Sheets(1).Range("F2").Offset(0, 2 * nmbs).Value

'Stergerea cadrului anterior
Sheets(1).Cells(prevnumX, prevnumY).Resize(z.Rows.Count + 1, z.Columns.Count + 1).Interior.Color = rgbWhite

isRight = False
isLeft = False
isDown = False
isUp = False
isBlocked = False

'Set pozVerf = anim.Offset(x, y)
'Verificam daca pozitiile deasupra, stanga, dreapta si jos sunt libere
'Deasupra
If anim.Offset(-1, 0).Interior.Color <> rgbWhite Or anim.Offset(-1, z.Columns.Count).Interior.Color <> rgbWhite Or _
anim.Offset(-1, z.Columns.Count / 2).Interior.Color <> rgbWhite Then isUp = True
'Stanga
If anim.Offset(0, -1).Interior.Color <> rgbWhite Or anim.Offset(z.Rows.Count, -1).Interior.Color <> rgbWhite Or _
anim.Offset(z.Rows.Count / 2, -1).Interior.Color <> rgbWhite Then isLeft = True
'Dreapta
If anim.Offset(0, 1 + z.Columns.Count).Interior.Color <> rgbWhite Or anim.Offset(z.Rows.Count, 1 + z.Columns.Count).Interior.Color <> rgbWhite Or _
anim.Offset(z.Rows.Count / 2, 1 + z.Columns.Count).Interior.Color <> rgbWhite Then isRight = True
'Jos
If anim.Offset(1 + z.Rows.Count, 0).Interior.Color <> rgbWhite Or anim.Offset(1 + z.Rows.Count, z.Columns.Count).Interior.Color <> rgbWhite Or _
anim.Offset(1 + z.Rows.Count, z.Columns.Count / 2).Interior.Color <> rgbWhite Then isDown = True

'E blocat din una sau mai multe din cele 4 directii.
If (isUp And x = -1) Or (isDown And x = 1) Or (isLeft And y = -1) Or (isRight And y = 1) Then isBlocked = True

'Copierea cadrului curent
z.Copy anim

'Setarea directiilor posibile
If isBlocked Then
    Dim coll As New Collection
    Dim direction As Integer
    
    'Alegem intre sus si jos in functie de coordonata x a destinatiei
    If isUp = False And x = -1 Then
    'If isUp = False And isDown = False And x = -1 Then
        coll.Add anim.Offset(-1, 0)
    ElseIf isDown = False And x = 1 Then
    'ElseIf isUp = False And isDown = False And x = 1 Then
        coll.Add anim.Offset(1, 0)
    Else:
    'Verificam daca directia de deplasare e libera
       'If isUp = False Then coll.Add anim.Offset(-1, 0)
       'If isDown = False Then coll.Add anim.Offset(1, 0)
    End If
    'Alegem intre stanga si dreapta in functie de coordonata y a destinatiei
    If isLeft = False And y = -1 Then
    'If isLeft = False And isRight = False And y = -1 Then
        coll.Add anim.Offset(0, -1)
    ElseIf isRight = False And y = 1 Then
    'ElseIf isLeft = False And isRight = False And y = 1 Then
        coll.Add anim.Offset(0, 1)
    Else:
    'Verificam daca directia de deplasare e libera
        'If isLeft = False Then coll.Add anim.Offset(0, -1)
        'If isRight = False Then coll.Add anim.Offset(0, 1)
    End If
    
    Randomize
    direction = Int(Rnd * coll.Count) + 1
    If coll.Count <> 0 Then
        Set tEmp = coll(direction)
    Else:
        'Sa stea pe loc
        Set tEmp = Cells(prevnumX, prevnumY)
    End If
    
Else: Set tEmp = anim.Offset(x, y)
End If

'Cazul in care cadrul stabilit a fost parcurs deja: una din pozitiile libere daca nu a fost parcursa deja
If prevColl(nmbs).Exists(CStr(tEmp.Row) & CStr(tEmp.Column)) Then

    'incrementam valoarea celulelor deja parcurse
    prevColl(nmbs)(CStr(tEmp.Row) & CStr(tEmp.Column)) = prevColl(nmbs)(CStr(tEmp.Row) & CStr(tEmp.Column)) + 1

    valeft = 0
    varight = 0
    vaup = 0
    vadown = 0

    'Daca directia e disponibila primeste o evaluare, altfel ramane cu 0
    If isLeft = False And evalPos(prevColl(nmbs)(CStr(anim.Offset(0, -1).Row) & CStr(anim.Offset(0, -1).Column))) > 0 Then
        'coll.Add anim.Offset(0, -1)
        valeft = evalPos(prevColl(nmbs)(CStr(anim.Offset(0, -1).Row) & CStr(anim.Offset(0, -1).Column)))
    End If
'    'If isLeft = False And Not (prevColl.Exists(anim.Offset(0, -1))) Then coll.Add anim.Offset(0, -1)
    If isRight = False And evalPos(prevColl(nmbs)(CStr(anim.Offset(0, 1).Row) & CStr(anim.Offset(0, 1).Column))) > 0 Then
        'coll.Add anim.Offset(0, 1)
        varight = evalPos(prevColl(nmbs)(CStr(anim.Offset(0, 1).Row) & CStr(anim.Offset(0, 1).Column)))
    End If
'    'If isLeft = False And Not (prevColl.Exists(anim.Offset(0, 1))) Then coll.Add anim.Offset(0, 1)
    If isUp = False And evalPos(prevColl(nmbs)(CStr(anim.Offset(-1, 0).Row) & CStr(anim.Offset(-1, 0).Column))) > 0 Then
        'coll.Add anim.Offset(-1, 0)
        vaup = evalPos(prevColl(nmbs)(CStr(anim.Offset(-1, 0).Row) & CStr(anim.Offset(-1, 0).Column)))
    End If
'    'If isLeft = False And Not (prevColl.Exists(anim.Offset(-1, 0))) Then coll.Add anim.Offset(-1, 0)
    If isDown = False And evalPos(prevColl(nmbs)(CStr(anim.Offset(1, 0).Row) & CStr(anim.Offset(1, 0).Column))) > 0 Then
        'coll.Add anim.Offset(1, 0)
        vadown = evalPos(prevColl(nmbs)(CStr(anim.Offset(1, 0).Row) & CStr(anim.Offset(1, 0).Column)))
    End If
'    'If isLeft = False And Not (prevColl.Exists(anim.Offset(1, 0))) Then coll.Add anim.Offset(1, 0)

    vatotal = valeft + varight + vaup + vadown
    Randomize
    direction = Int(Rnd * vatotal) + 1

    If valeft <> 0 And direction <= valeft Then
        Set anim = anim.Offset(0, -1)
        'MsgBox "left"
        nbtest = valeft
    ElseIf varight <> 0 And direction <= valeft + varight And direction > valeft Then
        Set anim = anim.Offset(0, 1)
        'MsgBox "right"
        nbtest = varight
    ElseIf vaup <> 0 And direction <= valeft + varight + vaup And direction > valeft + varight Then
        Set anim = anim.Offset(-1, 0)
        'MsgBox "up"
        nbtest = vaup
    ElseIf vadown <> 0 And direction <= valeft + varight + vaup + vadown And direction > valeft + varight + vaup Then
        Set anim = anim.Offset(1, 0)
        'MsgBox "down"
        nbtest = vadown
    End If
Else:
    Set anim = tEmp
End If

'lasarea unei urme pe alta foaie
Set anim2 = Sheets(3).Cells(anim.Row, anim.Column)
anim2.Interior.Color = RGB(0, 0, nmbs * 50)
anim2.Value = nbtest

'Odata ce cadrul a fost desenat, ii adaugam si pozitia in dictionar
If Not prevColl(nmbs).Exists(CStr(anim.Row) & CStr(anim.Column)) Then
    'prevColl.Add CStr(anim.Row) & CStr(anim.Column), anim
    prevColl(nmbs).Add CStr(anim.Row) & CStr(anim.Column), 0
End If

Sheets(1).Range("E2").Offset(0, 2 * nmbs).Value = anim.Row
Sheets(1).Range("F2").Offset(0, 2 * nmbs).Value = anim.Column

End Sub

Function evalPos(nrParc As Integer)
'mecanismul statistic de deplasare, sa-l schimb cu un model exponential? TODO o chestie mai desteapta de stabilire a directiei
If nrParc <= 3 Then
    evalPos = (3 - nrParc) ^ (3 - nrParc)
Else: evalPos = 0
End If

End Function
