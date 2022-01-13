Option Compare Database
'Zadatak 1
Public Sub UnosMatrice(ByRef mat() As Integer, ByVal brRed As Integer, ByVal brKol As Integer)
    Dim i As Integer
    Dim j As Integer
    'Koristimo dvostruki ciklus (za svaku promenu vrednosti promenljive i, promenljiva j ima vrednosti 1..brKol)
    For i = 1 To brRed
        For j = 1 To brKol
            'Unesite element [red,kol]:
            mat(i, j) = InputBox("Unesite element [" & i & "," & j & "]:")
        Next j
    Next i
End Sub
'Zadatak 1
Public Function PrikazMatrice(ByRef mat() As Integer, ByVal brRed As Integer, ByVal brKol As Integer) As String
    Dim i As Integer
    Dim j As Integer
    Dim rez As String
    rez = ""
    For i = 1 To brRed
        For j = 1 To brKol
            rez = rez & mat(i, j) & vbTab
        Next j
        'Na kraju svakog reda dodajemo novu liniju vbNewLine (tj. vrsimo prelazak u novi red)
        rez = rez & vbNewLine
    Next i
    PrikazMatrice = rez
End Function
'Zadatak 2
Public Function SumaZadKol(ByRef mat() As Integer, ByVal brRed As Integer, ByVal zadKol As Integer) As Integer
    'Kolona je zadata (tj. fiksna), a redovi se menjaju (prolazimo kroz sve redove matrice za jednu zadatu kolonu)
    Dim suma As Integer
    Dim i As Integer
    suma = 0
    For i = 1 To brRed
        'Prolazimo kroz sve redove i sumiramo elemente zadate kolone
        suma = suma + mat(i, zadKol)
    Next i
    SumaZadKol = suma
End Function
'Zadatak 3 (prvi nacin: funkcija vraca max sumu kolone)
Public Function MaxSumaKol(ByRef mat() As Integer, ByVal brRed As Integer, ByVal brKol As Integer) As Integer
    Dim j As Integer
    Dim max As Integer
    Dim suma As Integer
    'Na pocetku pretpostavimo da je suma prve kolone najveca
    max = SumaZadKol(mat, brRed, 1)
    For j = 2 To brKol
        'Racunamo sumu elemenata j-te kolone
        suma = SumaZadKol(mat, brRed, j)
        If suma > max Then
            'Nasli smo novu max sumu kolone - to je suma j-te kolone
            max = suma
        End If
    Next j
    MaxSumaKol = max
End Function
'Zadatak 3 (drugi nacin: funkcija vraca redni broj kolone sa max sumom)
Public Function MaxSumaKolIndex(ByRef mat() As Integer, ByVal brRed As Integer, ByVal brKol As Integer) As Integer
    Dim j As Integer
    Dim max As Integer
    Dim maxIndex As Integer
    Dim suma As Integer
    'Na pocetku pretpostavimo da je suma prve kolone najveca
    max = SumaZadKol(mat, brRed, 1)
    maxIndex = 1
    For j = 2 To brKol
        'Racunamo sumu elemenata j-te kolone
        suma = SumaZadKol(mat, brRed, j)
        If suma > max Then
            'Nasli smo novu max sumu kolone - to je suma j-te kolone
            max = suma
            maxIndex = j
        End If
    Next j
    MaxSumaKolIndex = maxIndex
End Function





