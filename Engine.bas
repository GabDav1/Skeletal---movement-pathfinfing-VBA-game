Attribute VB_Name = "Engine"

Sub ghostmove()
Dim slope As Integer
Dim b As Integer
Dim deltaY As Integer
Dim deltaX As Integer
Dim lastY As Integer
Dim lastX As Integer
Dim pozitieF As Range
Dim reverseP As Boolean
Dim counter As Integer
Dim endA As Integer
Dim arrayP() As Range

Call Settings

Set pozitieF = Cells(Range("A1").Value, Range("B1").Value)

' Aici setam punctul de origine si punctul destinatie; lastx si lasty sunt pentru "persistarea" ultimei destinatii
xA = ActiveCell.Column
xF = pozitieF.Column
yA = ActiveCell.Row
yF = pozitieF.Row
lastX = ActiveCell.Column
lastY = ActiveCell.Row
reverseP = False

' Cazul in care traiectoria merge spre STANGA-SUS sau dreapta-sus-sus
If (xA <= xF And yA <= yF) Or (xA > xF And yA < yF) Then
    xA = pozitieF.Column
    xF = ActiveCell.Column
    yA = pozitieF.Row
    yF = ActiveCell.Row
    reverseP = True
End If
  
'Inclinatia este raportul dintre diferenta coordonatelor deci functia va arata in felul urmator===> x=(y-b)/m DONC y=mx + b
'Oriunde apare inclinatia in calcule, valoarea ei absoluta(slope) va fi inlocuita cu raportul ei (cele 2 delta)
deltaY = (yA - yF)
deltaX = (xA - xF)
    
If deltaX = 0 Then
deltaX = 1
End If

'b reprezinta intersectul functiei, calculat dupa unul din cele 2 puncte care definesc functia
b = yF - ((deltaY * xF) / deltaX)

counter = 0

If xA > xF Then
' Se va calcula functia dupa numarul maxim de puncte a dreptunghiului a carui diagonala este linia pe care o calculam.
    If xA - xF <= yA - yF Then
        endA = yA - yF
        ReDim arrayP(endA)
        For y = yF To yA
        x = ((y - b) * deltaX) / deltaY
        Set arrayP(counter) = Cells(y, x)
        counter = counter + 1
        Next y
    ElseIf xA - xF > yA - yF Then
        endA = xA - xF
        ReDim arrayP(endA)
        For x = xF To xA
        y = (deltaY * x) / deltaX + b
        Set arrayP(counter) = Cells(y, x)
        counter = counter + 1
        Next x
    End If
' daca deplasarea e spre stanga
ElseIf xA <= xF Then
   If xF - xA <= yA - yF Then
        endA = yA - yF
        ReDim arrayP(endA)
        For y = yF To yA
        x = ((y - b) * deltaX) / deltaY
        Set arrayP(counter) = Cells(y, x)
        counter = counter + 1
        Next y
    ElseIf xF - xA > yA - yF Then
        endA = xF - xA
        ReDim arrayP(endA)
        For x = xA To xF
        y = (deltaY * x) / deltaX + b
        Set arrayP(counter) = Cells(y, x)
        counter = counter + 1
        Next x
    End If
End If

' De aici incepe animatia, mai intai verificam sensul de deplasare
If (reverseP = True And xA > xF) Or (xF - xA > yA - yF And reverseP = False) Or (reverseP = True And xA <= xF And xF - xA <= yA - yF) Then
    For i = UBound(arrayP) To LBound(arrayP) Step -1
        'Intercalam animatiile celorlalte obiecte in deplasarea personajului
        For u = 0 To noMobs - 1
            Call SkelMove(u, arrayP(i).Row, arrayP(i).Column, i)
        Next u
        DoEvents
        'Stergem cadrul anterior
        If i <= endA - 1 Then arrayP(i + 1).Resize(fantomas3.Rows.Count, fantomas3.Columns.Count).Interior.Color = rgbWhite
        'Alegem cadrul de animatie
        If i Mod 3 = 0 Then
            fantomas3.Copy arrayP(i)
        ElseIf i Mod 3 = 1 Then
            fantomas2.Copy arrayP(i)
        ElseIf i Mod 3 = 2 Then
            fantomas.Copy arrayP(i)
        End If
        DoEvents
    Next i
Else
    For i = LBound(arrayP) To UBound(arrayP)
        'Intercalam animatiile celorlalte obiecte in deplasarea personajului
        For u = 0 To noMobs - 1
            Call SkelMove(u, arrayP(i).Row, arrayP(i).Column, i)
        Next u
        DoEvents
        'Stergem cadrul anterior
        If i >= 1 Then arrayP(i - 1).Resize(fantomas3.Rows.Count, fantomas3.Columns.Count).Interior.Color = rgbWhite
        'Alegem cadrul de animatie
        If i Mod 3 = 0 Then
            fantomas3.Copy arrayP(i)
        ElseIf i Mod 3 = 1 Then
            fantomas2.Copy arrayP(i)
        ElseIf i Mod 3 = 2 Then
            fantomas.Copy arrayP(i)
        End If
        DoEvents
    Next i
End If

'TODO sa isi adapteze coordonatele de plecare in interiorul buclei de animatie
Range("A1").Value = lastY
Range("B1").Value = lastX

End Sub
