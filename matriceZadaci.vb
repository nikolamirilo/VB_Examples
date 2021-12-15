Option Compare Database


Private Sub BtnPokreniMA1_Click()

Dim mat(10, 10) As Integer

napuniMatricu mat(), 2, 2, "A"

prikaziMatricu mat(), 2, 2

End Sub

Private Sub BtnPokreniMA2_Click()
Dim A(2, 3) As Integer
Dim B(2, 3) As Integer
Dim C(2, 3) As Integer

napuniMatricu A(), 2, 3, "A"
napuniMatricu B(), 2, 3, "B"

sabiranjeMatrice A, B, C, 2, 3
MsgBox ("Matrica C koja je nastala sabiranjem matrice A i matrice B je: ")
prikaziMatricu C(), 2, 3


End Sub
Private Sub BtnPokreniMA3_Click()
Dim A(4, 4) As Integer

napuniMatricu A(), 4, 4, "A"
MsgBox ("Matrica A pre sortiranja: ")
prikaziMatricu A(), 4, 4
sortiranjeMatrice A(), 4, 4
MsgBox ("Matrica A posle sortiranja: ")
prikaziMatricu A(), 4, 4

End Sub

Private Sub BtnPokreniMA4_Click()
Dim X(3, 4) As Integer
Dim V(4, 1) As Integer
Dim R(3, 4) As Integer

napuniMatricu X(), 3, 4, "X"
napuniMatricu V(), 4, 1, "V"

mnozenjeMatriceVektorom X(), V(), R(), 3, 4

prikaziMatricu R(), 3, 1



End Sub



Private Sub BtnPokreniMA5_Click()

Dim Y(2, 3) As Integer
Dim Z(3, 4) As Integer
Dim Q(2, 4) As Integer

napuniMatricu Y(), 2, 3, "Y"
prikaziMatricu Y(), 2, 3
napuniMatricu Z(), 3, 4, "Z"
prikaziMatricu Z(), 3, 4
mnozenjeMatriceMatricom Y(), Z(), Q(), 2, 3, 4
prikaziMatricu Q(), 2, 4
End Sub
