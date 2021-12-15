Option Compare Database

Sub napuniMatricu(mat() As Integer, ByVal brojRedova As Integer, ByVal brojKolona As Integer, ByVal nazivMatrice As String)

Dim i As Integer
Dim j As Integer

For i = 1 To brojRedova
    For j = 1 To brojKolona
    mat(i, j) = InputBox("Unesi [" & i & " ," & j & "] element matrice " & nazivMatrice)
    Next j
Next i
End Sub


Sub prikaziMatricu(mat() As Integer, ByVal brojRedova As Integer, ByVal brojKolona As Integer)

Dim i As Integer
Dim j As Integer
Dim rez As String

rez = ""

For i = 1 To brojRedova
    For j = 1 To brojKolona
        'MsgBox ("Element [" & i & " ," & j & "] unete matrice je: " & mat(i, j))
        rez = rez & mat(i, j) & vbTab 'prikaz preko tabele, posle svakog novog clana u istom redu baci tab
    Next j
     rez = rez & vbNewLine 'posle svakog novog reda baci new line
Next i
MsgBox (rez)
End Sub


Sub sabiranjeMatrice(mat1() As Integer, mat2() As Integer, mat3() As Integer, ByVal brojRedova As Integer, ByVal brojKolona As Integer)

    Dim i As Integer
    Dim j As Integer
    
    For i = 1 To brojRedova
        For j = 1 To brojKolona
            mat3(i, j) = mat1(i, j) + mat2(i, j)
        Next j
    Next i
    
End Sub

Sub sortiranjeMatrice(mat() As Integer, ByVal brojRedova As Integer, ByVal brojKolona As Integer)

Dim i, j As Integer
Dim pom As Integer
For i = 1 To brojRedova - 1
    For j = i + 1 To brojRedova
        If mat(i, i) > mat(j, j) Then
            pom = mat(i, i)
            mat(i, i) = mat(j, j)
            mat(j, j) = pom
        End If
    Next j
Next i
End Sub


Sub mnozenjeMatriceVektorom(mat1() As Integer, mat2() As Integer, mat3() As Integer, ByVal brojRedova As Integer, ByVal brojKolona As Integer)

    Dim i As Integer
    Dim j As Integer
    
    For i = 1 To brojRedova
        mat3(i, 1) = 0
        For j = 1 To brojKolona
            mat3(i, 1) = mat3(i, 1) + (mat1(i, j) * mat2(j, 1))
        Next j
    Next i
    
End Sub

Sub mnozenjeMatriceMatricom(mat1() As Integer, mat2() As Integer, mat3() As Integer, ByVal brojRedova As Integer, ByVal brojRedovaKolona As Integer, ByVal brojKolona As Integer)

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    
    For i = 1 To brojRedova
        For j = 1 To brojKolona
        mat3(i, j) = 0
            For k = 1 To brojRedovaKolona
                mat3(i, j) = mat3(i, j) + (mat1(i, k) * mat2(k, j))
            Next k
        Next j
    Next i
    
End Sub

