Option Compare Database
Private Type Student
    ImePrezime As String
    BrojIndeksa As String
    GodinaStudija As Integer
End Type

Private Type Nastavnik
    ImePrezime As String
    Adresa As String
    Predmet As String
End Type


Private Sub BtnPokreniSL1_Click()
    Dim st As Student
    Dim brojStudenata As Integer
    
    st.ImePrezime = InputBox("Unesite vase ime i prezime")
    st.BrojIndeksa = InputBox("Unesi broj indeksa")
    st.GodinaStudija = InputBox("Unesi godinu studija")
    MsgBox ("Uneti student se zove " & st.ImePrezime & ", broj indeksa mu je " & st.BrojIndeksa & " i na " & st.GodinaStudija & ". godini studija je!")
    
End Sub

Private Sub BtnPokreniSL2_Click()
    Dim st(3) As Student
    Dim brojStudenata As Integer
    brojStudenata = InputBox("Unesite broj studenata koje zelite da unesete ovom prilikom!")
    
    For i = 1 To brojStudenata
        st(i).ImePrezime = InputBox("Unesite vase ime i prezime")
        st(i).BrojIndeksa = InputBox("Unesi broj indeksa")
        st(i).GodinaStudija = InputBox("Unesi godinu studija")
    Next i
    
    For i = 1 To brojStudenata
          MsgBox ("Uneti student se zove " & st(i).ImePrezime & ", broj indeksa mu je " & st(i).BrojIndeksa & " i na " & st(i).GodinaStudija & ". godini studija je!")
    Next i
   
End Sub

Private Sub BtnPokreniZSL1_Click()
    Dim nt(3) As Nastavnik
    Dim rez As String
    
    rez = ""
    
    For i = 1 To 3
        nt(i).ImePrezime = InputBox("Unesite vase ime i prezime")
        nt(i).Adresa = InputBox("Unesi adresu")
        nt(i).Predmet = InputBox("Unesi predmet")
    Next i
    
    
    For i = 1 To 3
    
    If nt(i).Predmet = "Matematika" Or nt(i).Predmet = "matematika" Then
        rez = rez & nt(i).ImePrezime & ", "
    End If
    Next i
    
    MsgBox ("Profesori koji predaju matu su: " & rez)
    
    
End Sub

    ' Dim nizStudenata(5) As Student
    ' Dim BrojIndeksa As String
    ' Dim ImePrezime As String
    ' Dim GodinaStudija As Integer
    ' Dim i, j As Integer
    ' For i = 1 To 5
    
    '     nizStudenata(i).ImePrezime = InputBox("Unesi ime i prezime")
    '     nizStudenata(i).BrojIndeksa = InputBox("Unesi br indeksa")
    '     nizStudenata(i).GodinaStudija = InputBox("Unesi godinu studija")
    ' Next i
    
    
    ' For i = 1 To 4
    '     For j = i + 1 To 5
    '     If StrComp(nizStudenata(i).BrojIndeksa, nizStudenata(j).BrojIndeksa) = 1 Then
    '         BrojIndeksa = nizStudenata(i).BrojIndeksa
    '         ImePrezime = nizStudenata(i).ImePrezime
    '         GodinaStudija = nizStudenata(i).GodinaStudija
            
    '         nizStudenata(i).BrojIndeksa = nizStudenata(j).BrojIndeksa
    '         nizStudenata(i).ImePrezime = nizStudenata(j).ImePrezime
    '         nizStudenata(i).GodinaStudija = nizStudenata(j).GodinaStudija
            
    '         nizStudenata(j).BrojIndeksa = BrojIndeksa
    '         nizStudenata(j).ImePrezime = ImePrezime
    '         nizStudenata(j).GodinaStudija = GodinaStudija
    '     End If
    '     Next j
    ' Next i
