Option Compare Database
Option Explicit

' Copyright 2009-2012 Denis SCHEIDT
' Ce programme est distribué sous Licence LGPL

'    This file is part of libMAIL

'    libMAIL is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Lesser General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.

'    libMAIL is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Lesser General Public License for more details.

'    You should have received a copy of the GNU Lesser General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.



Private Const MAX32 As Currency = 4294967296@
Private Const MAX24 As Currency = 16777216@
Private Const MAX16 As Currency = 65536@
Private Const MAX08 As Currency = 256@

Dim o() As Byte

' VBA ne possède pas de type entier long non signé.
' Il faut passer par Currency pour avoir une capacité suffisante.


' Limité à une longeur de 2^32 bits
'
' Retourne une chaine de 16 octets.
Function MD5(Param As String) As String
    Dim lLng As Currency, l0 As Currency, l1 As Currency
    Dim a  As Currency, b  As Currency, c  As Currency, d  As Currency, n As Currency
    Dim aa As Currency, bb As Currency, cc As Currency, dd As Currency
    Dim tSin() As Currency

    ' Passer la chaine dans un tableau d'octets
    o = StrConv(Param, vbFromUnicode)

    ' Etape 1 : compléter pour avoir une longueur multiple de 512 bits (64 octets)
    lLng = UBound(o) + 1                                    ' Nb d'octets du tableau d'entrée
    l0 = lLng Mod 64                                        ' Nb d'octets du dernier bloc
    l1 = lLng + (64 - l0)                                   ' Compléter à 64 octets
    If l0 >= 56 Then l1 = l1 + 64                           ' Ajouter un bloc
    ReDim Preserve o(l1 - 1)
    o(lLng) = 128                                           ' Mettre le bit suivant à 1, le reste est déjà à 0

    ' Etape 2 : écrire la longueur de la chaine, en bits, dans les 8 derniers octets,
    l0 = l1 - 8
    ' le mot de poids faible en tête
    o(l0 + 0) = (lLng * 8) And &HFF&
    o(l0 + 1) = ((lLng * 8) And &HFF00&) \ &H100&
    o(l0 + 2) = ((lLng * 8) And &HFF0000) \ &H10000
    o(l0 + 3) = Int((lLng * 8) And &HFF000000) / &H1000000
    ' puis le mot de poids fort, qui reste à zéro, car dépassement de capacité...
    ' ...

    ' Valeurs de hachage
    a = 1732584193@     '&H67452301
    b = 4023233417@     '&HEFCDAB89
    c = 2562383102@     '&H98BADCFE
    d = 271733878@      '&H10325476

    ' Initialiser le tableau des Sinus
    Call tblSin(tSin)

    ' Traiter les blocs de 16 mots (64 octets, 512 bits)
    For n = 0 To l1 - 1 Step 64
        aa = a: bb = b: cc = c: dd = d

        FTr1 a, b, c, d, Cv4O2C(n + 4@ * 0), 7, tSin(1)
        FTr1 d, a, b, c, Cv4O2C(n + 4@ * 1), 12, tSin(2)
        FTr1 c, d, a, b, Cv4O2C(n + 4@ * 2), 17, tSin(3)
        FTr1 b, c, d, a, Cv4O2C(n + 4@ * 3), 22, tSin(4)
        FTr1 a, b, c, d, Cv4O2C(n + 4@ * 4), 7, tSin(5)
        FTr1 d, a, b, c, Cv4O2C(n + 4@ * 5), 12, tSin(6)
        FTr1 c, d, a, b, Cv4O2C(n + 4@ * 6), 17, tSin(7)
        FTr1 b, c, d, a, Cv4O2C(n + 4@ * 7), 22, tSin(8)
        FTr1 a, b, c, d, Cv4O2C(n + 4@ * 8), 7, tSin(9)
        FTr1 d, a, b, c, Cv4O2C(n + 4@ * 9), 12, tSin(10)
        FTr1 c, d, a, b, Cv4O2C(n + 4@ * 10), 17, tSin(11)
        FTr1 b, c, d, a, Cv4O2C(n + 4@ * 11), 22, tSin(12)
        FTr1 a, b, c, d, Cv4O2C(n + 4@ * 12), 7, tSin(13)
        FTr1 d, a, b, c, Cv4O2C(n + 4@ * 13), 12, tSin(14)
        FTr1 c, d, a, b, Cv4O2C(n + 4@ * 14), 17, tSin(15)
        FTr1 b, c, d, a, Cv4O2C(n + 4@ * 15), 22, tSin(16)

        FTr2 a, b, c, d, Cv4O2C(n + 4@ * 1), 5, tSin(17)
        FTr2 d, a, b, c, Cv4O2C(n + 4@ * 6), 9, tSin(18)
        FTr2 c, d, a, b, Cv4O2C(n + 4@ * 11), 14, tSin(19)
        FTr2 b, c, d, a, Cv4O2C(n + 4@ * 0), 20, tSin(20)
        FTr2 a, b, c, d, Cv4O2C(n + 4@ * 5), 5, tSin(21)
        FTr2 d, a, b, c, Cv4O2C(n + 4@ * 10), 9, tSin(22)
        FTr2 c, d, a, b, Cv4O2C(n + 4@ * 15), 14, tSin(23)
        FTr2 b, c, d, a, Cv4O2C(n + 4@ * 4), 20, tSin(24)
        FTr2 a, b, c, d, Cv4O2C(n + 4@ * 9), 5, tSin(25)
        FTr2 d, a, b, c, Cv4O2C(n + 4@ * 14), 9, tSin(26)
        FTr2 c, d, a, b, Cv4O2C(n + 4@ * 3), 14, tSin(27)
        FTr2 b, c, d, a, Cv4O2C(n + 4@ * 8), 20, tSin(28)
        FTr2 a, b, c, d, Cv4O2C(n + 4@ * 13), 5, tSin(29)
        FTr2 d, a, b, c, Cv4O2C(n + 4@ * 2), 9, tSin(30)
        FTr2 c, d, a, b, Cv4O2C(n + 4@ * 7), 14, tSin(31)
        FTr2 b, c, d, a, Cv4O2C(n + 4@ * 12), 20, tSin(32)

        FTr3 a, b, c, d, Cv4O2C(n + 4@ * 5), 4, tSin(33)
        FTr3 d, a, b, c, Cv4O2C(n + 4@ * 8), 11, tSin(34)
        FTr3 c, d, a, b, Cv4O2C(n + 4@ * 11), 16, tSin(35)
        FTr3 b, c, d, a, Cv4O2C(n + 4@ * 14), 23, tSin(36)
        FTr3 a, b, c, d, Cv4O2C(n + 4@ * 1), 4, tSin(37)
        FTr3 d, a, b, c, Cv4O2C(n + 4@ * 4), 11, tSin(38)
        FTr3 c, d, a, b, Cv4O2C(n + 4@ * 7), 16, tSin(39)
        FTr3 b, c, d, a, Cv4O2C(n + 4@ * 10), 23, tSin(40)
        FTr3 a, b, c, d, Cv4O2C(n + 4@ * 13), 4, tSin(41)
        FTr3 d, a, b, c, Cv4O2C(n + 4@ * 0), 11, tSin(42)
        FTr3 c, d, a, b, Cv4O2C(n + 4@ * 3), 16, tSin(43)
        FTr3 b, c, d, a, Cv4O2C(n + 4@ * 6), 23, tSin(44)
        FTr3 a, b, c, d, Cv4O2C(n + 4@ * 9), 4, tSin(45)
        FTr3 d, a, b, c, Cv4O2C(n + 4@ * 12), 11, tSin(46)
        FTr3 c, d, a, b, Cv4O2C(n + 4@ * 15), 16, tSin(47)
        FTr3 b, c, d, a, Cv4O2C(n + 4@ * 2), 23, tSin(48)

        FTr4 a, b, c, d, Cv4O2C(n + 4@ * 0), 6, tSin(49)
        FTr4 d, a, b, c, Cv4O2C(n + 4@ * 7), 10, tSin(50)
        FTr4 c, d, a, b, Cv4O2C(n + 4@ * 14), 15, tSin(51)
        FTr4 b, c, d, a, Cv4O2C(n + 4@ * 5), 21, tSin(52)
        FTr4 a, b, c, d, Cv4O2C(n + 4@ * 12), 6, tSin(53)
        FTr4 d, a, b, c, Cv4O2C(n + 4@ * 3), 10, tSin(54)
        FTr4 c, d, a, b, Cv4O2C(n + 4@ * 10), 15, tSin(55)
        FTr4 b, c, d, a, Cv4O2C(n + 4@ * 1), 21, tSin(56)
        FTr4 a, b, c, d, Cv4O2C(n + 4@ * 8), 6, tSin(57)
        FTr4 d, a, b, c, Cv4O2C(n + 4@ * 15), 10, tSin(58)
        FTr4 c, d, a, b, Cv4O2C(n + 4@ * 6), 15, tSin(59)
        FTr4 b, c, d, a, Cv4O2C(n + 4@ * 13), 21, tSin(60)
        FTr4 a, b, c, d, Cv4O2C(n + 4@ * 4), 6, tSin(61)
        FTr4 d, a, b, c, Cv4O2C(n + 4@ * 11), 10, tSin(62)
        FTr4 c, d, a, b, Cv4O2C(n + 4@ * 2), 15, tSin(63)
        FTr4 b, c, d, a, Cv4O2C(n + 4@ * 9), 21, tSin(64)

        a = myMod(a + aa)
        b = myMod(b + bb)
        c = myMod(c + cc)
        d = myMod(d + dd)
    Next n

    ' Sortir le MD5.
    MD5 = CvC2O_le(a) & CvC2O_le(b) & CvC2O_le(c) & CvC2O_le(d)

    Erase tSin
End Function



' Calcul utilisé dans le tour 1
Private Sub FTr1(a As Currency, b As Currency, c As Currency, d As Currency, xk As Currency, s As Integer, ti As Currency)
    a = myMod(b + RotGauche(a + fF(b, c, d) + xk + ti, s))
End Sub

' Calcul utilisé dans le tour 2
Private Sub FTr2(a As Currency, b As Currency, c As Currency, d As Currency, xk As Currency, s As Integer, ti As Currency)
    a = myMod(b + RotGauche(a + fG(b, c, d) + xk + ti, s))
End Sub

' Calcul utilisé dans le tour 3
Private Sub FTr3(a As Currency, b As Currency, c As Currency, d As Currency, xk As Currency, s As Integer, ti As Currency)
    a = myMod(b + RotGauche(a + fH(b, c, d) + xk + ti, s))
End Sub

' Calcul utilisé dans le tour 4
Private Sub FTr4(a As Currency, b As Currency, c As Currency, d As Currency, xk As Currency, s As Integer, ti As Currency)
    a = myMod(b + RotGauche(a + fI(b, c, d) + xk + ti, s))
End Sub



Private Function fF(x As Currency, y As Currency, z As Currency) As Currency
    fF = Or32(And32(x, y), And32(Not32(x), z))
End Function

Private Function fG(x As Currency, y As Currency, z As Currency) As Currency
    fG = Or32(And32(x, z), And32(y, Not32(z)))
End Function

Private Function fH(x As Currency, y As Currency, z As Currency) As Currency
    fH = XOr32(XOr32(x, y), z)
End Function

Private Function fI(x As Currency, y As Currency, z As Currency) As Currency
    fI = XOr32(y, Or32(x, Not32(z)))
End Function



' Retourne un Currency calculé à l'aide de 4 octets consécutifs du tableau.
Private Function Cv4O2C(n As Currency) As Currency
    Cv4O2C = o(n + 0@) + _
             o(n + 1@) * MAX08 + _
             o(n + 2@) * MAX16 + _
             o(n + 3@) * MAX24
End Function

' Génère le tableau des sinus
Private Sub tblSin(tSin() As Currency)
    Dim j As Byte

    ReDim tSin(64)

    For j = 1 To 64
        tSin(j) = Int(Abs(Sin(j)) * MAX32)
    Next j
End Sub