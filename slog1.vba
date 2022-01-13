Option Compare Database

'Svi zadaci - modelujemo isplatu
Type Isplata
    SifraRadnika As Integer
    ImePrezime As String
    Mesec As Integer
    Godina As Integer
    Iznos As Double
End Type
'Zadatak 1
Public Sub DodajIsplatu(ByRef niz() As Isplata, ByRef n As Integer, ByRef isp As Isplata)
    n = n + 1
    niz(n) = isp
End Sub
'Zadatak 1
Public Function PrikazIsplate(ByRef isp As Isplata) As String
    Dim rez As String
    rez = isp.SifraRadnika & vbTab & isp.ImePrezime & vbTab & isp.Godina & vbTab & isp.Mesec & vbTab & isp.Iznos
    PrikazIsplate = rez
End Function
'Zadatak 2
Public Function MaxZarada(ByRef niz() As Isplata, ByVal n As Integer, ByVal zadataGodina As Integer) As Isplata
    'Povratna vrednost funkcije je tipa Isplata - imamo kompletan podatak o najvecoj isplati (tj. znamo radnika, mesec, godinu i iznos)
    Dim max As Isplata
    Dim prvaIsplataZaGodinu As Boolean
    Dim i As Integer
    'Pretpostavimo da u nizu ne postoje isplate za zadatu godinu
    prvaIsplataZaGodinu = False
    For i = 1 To n
        If niz(i).Godina = zadataGodina And prvaIsplataZaGodinu = False Then
            'Nasli smo prvu isplatu za zadatu godinu i nju postavljamo za max
            max = niz(i)
            prvaIsplataZaGodinu = True
        End If
                
        'Pitamo da li je tekuca vrednost veca od max vrednosti
        If niz(i).Iznos > max.Iznos And niz(i).Godina = zadataGodina Then
            'Nasli smo novi max u zadatoj godini
            max = niz(i)
        End If
    Next i
    MaxZarada = max
End Function
'Zadatak 3
Public Function ProsecnaZaradaRadnikGodina(ByRef niz() As Isplata, ByVal n As Integer, ByVal SifraRadnika As Integer, ByVal zadataGodina As Integer) As Double
    Dim i As Integer
    Dim suma As Double
    Dim brojac As Integer
    Dim arSred As Double
    suma = 0
    brojac = 0
    arSred = 0
    For i = 1 To n
        'Proveravamo da li je isplata iz zadate godine i da li odgovara zadatom radniku
        If niz(i).Godina = zadataGodina And niz(i).SifraRadnika = SifraRadnika Then
            'Nasli smo isplatu zadatog radnika u zadatoj godini -> povecavamo sumu i brojac
            suma = suma + niz(i).Iznos
            brojac = brojac + 1
        End If
    Next i
    If brojac > 0 Then
        'Racunamo prosecnu zaradu radnika u godini
        arSred = suma / brojac
    End If
    ProsecnaZaradaRadnikGodina = arSred
End Function
