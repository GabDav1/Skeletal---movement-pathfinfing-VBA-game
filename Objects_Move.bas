Attribute VB_Name = "Objects_Move"
Sub SkelMove(nmbs As Integer, toPointx As Integer, toPointy As Integer, iterator As Variant)

Dim randnum As Integer
Dim temp As Range

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
    Set temp = mobs(1 + nmbs * 3)
ElseIf iterator Mod 3 = 1 Then
    Set temp = mobs(2 + nmbs * 3)
Else: Set temp = mobs(3 + nmbs * 3)
End If

'Desenarea cadrelor in functie de directia stabilita
Select Case randnum
    Case 1
        Call moveloop(nmbs, 1, -1, temp)
    Case 2
        Call moveloop(nmbs, 1, 1, temp)
    Case 3
        Call moveloop(nmbs, -1, -1, temp)
    Case 4
        Call moveloop(nmbs, -1, 1, temp)
End Select

'Scrierea coordonatelor pentru cadrul urmator
Range("C3").Offset(0, 2 * nmbs).Value = anim.Row
Range("D3").Offset(0, 2 * nmbs).Value = anim.Column

End Sub

Sub moveloop(nmbs As Integer, x As Integer, y As Integer, z As Range)
Dim pozVerf As Range
Dim isRight As Boolean
Dim isLeft As Boolean
Dim isDown As Boolean
Dim isUp As Boolean
Dim isBlocked As Boolean
Dim prevnumX As Integer
Dim prevnumY As Integer

prevnumX = Sheets(1).Range("E2").Offset(0, 2 * nmbs).Value
prevnumY = Sheets(1).Range("F2").Offset(0, 2 * nmbs).Value

'Stergerea cadrului anterior
Sheets(1).Cells(prevnumX, prevnumY).Resize(z.Rows.Count + 1, z.Columns.Count + 1).Interior.Color = rgbWhite

isRight = False
isLeft = False
isDown = False
isUp = False
isBlocked = False

'TODO : de facut mecanism STATISTIC de deplasare: toate celulele care au fost vizitate sunt memorate si primesc rating:
' cu cat sunt vizitate mai des, cu atat ratingul creste. Cu cat ratingul e mai ridicat cu atat sunt mai mici sansele de a fi vizitata din nou.

Set pozVerf = anim.Offset(x, y)
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
    If isUp = False And isDown = False And x = -1 Then
        coll.Add anim.Offset(-1, 0)
    ElseIf isUp = False And isDown = False And x = 1 Then
        coll.Add anim.Offset(1, 0)
    Else:
    'Verificam daca directia de deplasare e libera si NU e aceeasi cu directia precedenta
        If isUp = False And anim.Offset(-1, 0).Row <> prevnumX And anim.Offset(-1, 0).Column <> prevnumY Then coll.Add anim.Offset(-1, 0)
        If isDown = False And anim.Offset(1, 0).Row <> prevnumX And anim.Offset(1, 0).Column <> prevnumY Then coll.Add anim.Offset(1, 0)
    End If
    'Alegem intre stanga si dreapta in functie de coordonata y a destinatiei
    If isLeft = False And isRight = False And y = -1 Then
        coll.Add anim.Offset(0, -1)
    ElseIf isLeft = False And isRight = False And y = 1 Then
        coll.Add anim.Offset(0, 1)
    Else:
    'Verificam daca directia de deplasare e libera si NU e aceeasi cu directia precedenta
        If isLeft = False And anim.Offset(0, -1).Row <> prevnumX And anim.Offset(0, -1).Column <> prevnumY Then coll.Add anim.Offset(0, -1)
        If isRight = False And anim.Offset(0, 1).Row <> prevnumX And anim.Offset(0, 1).Column <> prevnumY Then coll.Add anim.Offset(0, 1)
    End If
    
    Randomize
    direction = Int(Rnd * coll.Count) + 1
    If coll.Count <> 0 Then
    'TODO: INCA O COLECTIE CU TOATE POZITIILE PRECEDENTE SETATA AICI, COLECTIA ASTA VA FI FOLOSITA IN LOC DE PREVNUMX SI PREVNUMY
        Set anim = coll(direction)
    Else:
        'Sa mearga in directia opusa daca nu a gasit nici o directie valabila
        Set anim = anim.Offset(-x, -y)
    End If
    
Else: Set anim = anim.Offset(x, y)
End If

Sheets(1).Range("E2").Offset(0, 2 * nmbs).Value = anim.Row
Sheets(1).Range("F2").Offset(0, 2 * nmbs).Value = anim.Column

End Sub
