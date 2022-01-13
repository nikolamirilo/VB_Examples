Option Compare Database

Private Sub BtnPokreniDAT1_Click()
Dim niz(4) As Integer
Dim pom As Integer
Dim brojElem As Integer

For i = 1 To 4
    niz(i) = InputBox("Unesi broj")
Next i

Set fs = CreateObject("Scripting.FileSystemObject")
Set dat = fs.createTextFile("data.txt", True)

dat.WriteLine 4

For i = 1 To 4
    dat.WriteLine niz(i)
Next i

dat.Close

Set dat = fs.OpenTextFile("data.txt", 1, 0)
brojElem = dat.ReadLine
MsgBox ("Parni brojevi")
For i = 1 To brojElem
    pom = dat.ReadLine
    If (pom Mod 2 = 0) Then
        MsgBox pom
    End If
Next i
dat.Close
End Sub

Private Sub BtnPokreniDAT2_Click()
Dim brojImena As Integer
Dim ime As String

Set fs = CreateObject("Scripting.FileSystemObject")
Set dat = fs.createTextFile("data1.txt", True)
brojImena = InputBox("Unesite broj imena koje zelis da zapamtis")
dat.WriteLine brojImena

For i = 1 To brojImena
    ime = InputBox("Unesi " & i & ". ime")
    dat.WriteLine ime
Next i
dat.Close

Set dat = fs.OpenTextFile("data1.txt", 1, 0)
brojImena = dat.ReadLine
MsgBox ("Imena koja pocinju na S")
For i = 1 To brojImena
    ime = dat.ReadLine
    If Left(ime, 1) = "S" Then
    MsgBox (ime)
    End If
Next i
dat.Close
End Sub

Private Sub BtnPokreniDAT3_Click()

End Sub
