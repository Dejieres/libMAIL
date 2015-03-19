Option Compare Database
Option Explicit

' Copyright 2009-2013 Denis SCHEIDT
' Ce programme est distribu� sous Licence LGPL

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





' D�code une cha�ne en Base64
Function Dec_Base64(sEncB64 As String) As String
    Dim l As Long, tOctets() As Byte, i As Integer, j As Long
    Dim lNbCarE As Long, l3octets As Long, oSortie() As Byte, iPos As Long

    lNbCarE = Len(sEncB64)                                      ' Nb de caract�res en entr�e

    If lNbCarE = 0 Then Exit Function

    j = lNbCarE Mod 4
    If j <> 0 Then l = lNbCarE + 4 - j Else l = lNbCarE

    ' Calcul brutal : la cha�ne de sortie repr�sente 75% de la chaine d'entr�e.
    ' S'il y a des retours chariots, elle sera trop longue. Pas grave ;)
    ReDim oSortie(l * 3 / 4 - 1)

    tOctets = StrConv(sEncB64, vbFromUnicode)                   ' Convertir brutalement en octets

    iPos = -1
    For l = 0 To lNbCarE - 1
        j = tOctets(l)
        Select Case j
            Case 65 To 90:      j = j - 65                      ' ABCDEFGHIJKLMNOPQRSTUVWXYZ
            Case 97 To 122:     j = j - 71                      ' abcdefghijklmnopqrstuvwxyz
            Case 48 To 57:      j = j + 4                       ' 0123456789
            Case 43:            j = 62                          ' +
            Case 47:            j = 63                          ' /
            Case Else:          j = -1                          ' Autre...
        End Select

        If j >= 0 Then
            i = i + 1                                           ' Car. Base64 valide
            Select Case i
                Case 1:     l3octets = l3octets + j * &H40000   ' 6 bits de poids fort
                Case 2:     l3octets = l3octets + j * &H1000&   ' .
                Case 3:     l3octets = l3octets + j * &H40&     ' .
                Case 4                                          ' 6 bits de poids faible
                    l3octets = l3octets + j

                    ' Convertir en 3 caract�res 8 bits
                    iPos = iPos + 1: oSortie(iPos) = (l3octets And &HFF0000) / &H10000
                    iPos = iPos + 1: oSortie(iPos) = (l3octets And &HFF00&) / &H100&
                    iPos = iPos + 1: oSortie(iPos) = l3octets And &HFF&

                    i = 0                                       ' RAZ pour le prochain tour
                    l3octets = 0&
            End Select
        End If
    Next l

    If i <> 0 Then                                              ' Il reste un bout � traiter.
        ' Convertir en 3 caract�res 8 bits
        iPos = iPos + 1: oSortie(iPos) = (l3octets And &HFF0000) / &H10000
        iPos = iPos + 1: oSortie(iPos) = (l3octets And &HFF00&) / &H100&
        iPos = iPos + 1: oSortie(iPos) = l3octets And &HFF&
        i = -4 + i                                              ' Ignorer les '=' de fin, le cas �ch�ant
    End If

    Erase tOctets

    If (iPos + i) >= 0 Then
        ReDim Preserve oSortie(iPos + i)                        ' Tronquer � la longueur utile.
        Dec_Base64 = StrConv(oSortie, vbUnicode)
    End If

    Erase oSortie
End Function

' D�code une cha�ne encod�e en Quoted-Printable
' Lorsque QEncoding est True, '_' devient ' '.
Function Dec_QP(sChQP As String, Optional QEncoding As Boolean = False) As String
    Dim bi() As Byte, bo() As Byte, bCar As Long, lCar As Long
    Dim i As Long, j As Long, l As Long

    ' Tableau de conversion
    Dim bASC(255) As Long
    ' On itialise le tableau � 256. De cette mani�re, un �gal suivi d'un caract�re autre que 0-9, A-F
    ' produira un r�sultat sup�rieur � 255, et ne sera pas d�cod�.
    For i = 0 To 255
        bASC(i) = 256
    Next i
    ' Positions correspondant aux chiffres Hex (0-9, A-F)
    bASC(48) = 0:   bASC(49) = 1:   bASC(50) = 2:   bASC(51) = 3
    bASC(52) = 4:   bASC(53) = 5:   bASC(54) = 6:   bASC(55) = 7
    bASC(56) = 8:   bASC(57) = 9:   bASC(65) = 10:  bASC(66) = 11
    bASC(67) = 12:  bASC(68) = 13:  bASC(69) = 14:  bASC(70) = 15

    bi = StrConv(sChQP, vbFromUnicode)                          ' Convertir la cha�ne en tableau d'octets.
    l = UBound(bi)
    ReDim bo(l)                                                 ' Pr�allouer le tableau de sortie.

    i = 0
    Do
        bCar = bi(i)
        Select Case bCar
            Case 9&, 32& To 60&, 62& To 126&                    ' Caract�res sans conversion.
                If bCar = 95& And QEncoding Then bCar = 32& ' '_' --> ' '

                bo(j) = bCar
                j = j + 1&

            Case 61&                                    ' '='
                ' Il faut voir les 2 caract�res qui suivent le '='.
                lCar = bASC(bi(i + 1&)) * 16& + bASC(bi(i + 2&))

                Select Case lCar
                    Case 0& To 217&, 219& To 255&
                        bo(j) = lCar                            ' C'est un nombre Hex --> caract�re encod�.
                        j = j + 1&
                        i = i + 2&

                    Case 218&                                   ' C'est un CRLF qui suit le '=' --> Soft Break.
                        i = i + 2&                              ' On ignore...

                    Case Else                                   ' Caract�res invalides apr�s le '='.
                        i = i + 2&                              ' On ignore...

                End Select

            Case 13&                                            ' CR
                Select Case bi(i + 1&)
                    Case 10&                                    ' LF
                        bo(j) = bCar                            ' Ecrire le CRLF
                        bo(j + 1&) = bi(i + 1&)
                        i = i + 1&
                        j = j + 2&

                    Case Else                                   ' Ignorer.
                        ' Rien

                End Select

        End Select

        i = i + 1&
    Loop While i <= l

    ReDim Preserve bo(j - 1&)
    Dec_QP = StrConv(bo, vbUnicode)
End Function

' Conversion d'une cha�ne de caract�re en Base64.
' Cette fonction a �t� optimis�e en occupation m�moire et en vitesse de traitement
' par l'emploi de la pr�-allocation de cha�ne.
' Voir http://support.microsoft.com/?scid=kb%3Ben-us%3B170964&x=19&y=13 pour une discussion
' sur les performances (d�sastreuses) des concat�nations de cha�nes.

' Une description de la conversion en Base64 : http://fr.wikipedia.org/wiki/Base64
'
' Lorsque lCRLF est diff�rent de 0, la fonction ins�re un retour chariot apr�s oCRLF caract�res.
' lCRLF est ramen� au multiple de 4 le plus proche.
Function Enc_Base64(sEntree As String, Optional ByVal lCRLF As Long = 76) As String
    Dim tOctets() As Byte, lLong As Long, oSortie() As Byte
    Dim i As Long, j As Long, l As Long
    Dim sB64() As Byte

    sB64 = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
                   "abcdefghijklmnopqrstuvwxyz" & _
                   "0123456789+/", vbFromUnicode)

    ' Pr�-allocation du tableau de sortie. On calcule la taille exacte.
    ' Seules les positions paires du tableau de sortie sont �crites, pour
    ' tenir compte de la conversion Unicode lors de l'affectation finale.
    l = Len(sEntree)

    If l = 0 Then Exit Function

    tOctets = StrConv(sEntree, vbFromUnicode)                   ' Convertir toute la chaine en octets
    If (l Mod 3) <> 0 Then
        ReDim Preserve tOctets(l + 3 - (l Mod 3) - 1)           ' Ajuster au multiple de 3 sup�rieur
    End If
    j = (UBound(tOctets) + 1) * 4 / 3                           ' 3 octets en entr�e = 4 octets en sortie.

    If lCRLF > 0 Then                                           ' Pr�voir des retours chariots tous les n car de sortie.
        lCRLF = Int((lCRLF / 4) + 0.5) * 4                      ' Ramener au multiple de 4 le plus proche.
        j = j + (Int((j - 1) / lCRLF)) * 2
    End If
    ReDim oSortie(j - 1)                                        ' Taille d�finitive du tableau de sortie

    j = 0                                                       ' Position d'�criture dans la cha�ne de sortie
    l = 0                                                       ' Compteur de caract�res sortis
    For i = 0 To UBound(tOctets) Step 3                         ' Pour chaque groupe de 3 octets.
        If l >= lCRLF And lCRLF > 0 Then
            oSortie(j + 0) = 13                                 ' Retour chariot apr�s n caract�res de sortie.
            oSortie(j + 1) = 10
            j = j + 2                                           ' Ajuster le pointeur � la position suivante.
            l = 0
        End If

        ' Calcul de l'entier long (24 bits)
        lLong = &H10000 * tOctets(i) + &H100& * tOctets(i + 1) + tOctets(i + 2)

        ' Utiliser les groupes de 6 bits comme indices dans la cha�ne de conversion.
        ' Les groupes de 6 bits donnent 4 caract�res Base64.
        ' Ecrire dans le tableau de sortie, � la position du pointeur.
        oSortie(j + 0) = sB64((lLong And &O77000000) / &O1000000)
        oSortie(j + 1) = sB64((lLong And &O770000) / &O10000)
        oSortie(j + 2) = sB64((lLong And &O7700) / &O100)
        oSortie(j + 3) = sB64((lLong And &O77))

        j = j + 4                                               ' D�placer le pointeur de sortie.
        l = l + 4                                               ' Quatre caract�res ecrits.
    Next i

    Erase tOctets, sB64                                         ' Nettoyage et lib�ration de m�moire

    i = Len(sEntree) Mod 3
    If i <> 0 Then
        ' Il reste 8 ou 16 bits, � compl�ter par == ou =
        oSortie(j - 1) = 61
        If i = 1 Then oSortie(j - 2) = 61
    End If

    Enc_Base64 = StrConv(oSortie(), vbUnicode)

    Erase oSortie                                               ' Nettoyage et lib�ration de m�moire
End Function

' Encode une cha�ne de caract�res en Quoted-printable
' Lorsque QEncoding est True, le caract�re "?" est encod� aussi,
' et l'espace est remplac� par "_".
' Ceci est n�cessaire pour pouvoir encoder correctement l'objet du message ainsi que les
' noms des pi�ces jointes.
' Lorsque QEncoding = True, il n'y a pas d'ajout de SoftBreak (CRLF) pour couper les lignes � 76 car. max.
Function Enc_QP(sChaine As String, Optional QEncoding As Boolean = False) As String
    Dim sQP() As Byte, bCar As Long, i As Long, l As Long, j As Long, n As Long, k As Long
    Dim b() As Byte, bEncode As Long

    ' Tableau des valeurs Hex de 00 � FF. Ce tableau comporte 512 lignes.
    '   La ligne 2*CodeAsc donne le premier caract�re Hex,
    '   la ligne 2*CodeAsc+1 le second.
    Dim bASC() As Byte
    bASC = StrConv("000102030405060708090A0B0C0D0E0F" & _
                   "101112131415161718191A1B1C1D1E1F" & _
                   "202122232425262728292A2B2C2D2E2F" & _
                   "303132333435363738393A3B3C3D3E3F" & _
                   "404142434445464748494A4B4C4D4E4F" & _
                   "505152535455565758595A5B5C5D5E5F" & _
                   "606162636465666768696A6B6C6D6E6F" & _
                   "707172737475767778797A7B7C7D7E7F" & _
                   "808182838485868788898A8B8C8D8E8F" & _
                   "909192939495969798999A9B9C9D9E9F" & _
                   "A0A1A2A3A4A5A6A7A8A9AAABACADAEAF" & _
                   "B0B1B2B3B4B5B6B7B8B9BABBBCBDBEBF" & _
                   "C0C1C2C3C4C5C6C7C8C9CACBCCCDCECF" & _
                   "D0D1D2D3D4D5D6D7D8D9DADBDCDDDEDF" & _
                   "E0E1E2E3E4E5E6E7E8E9EAEBECEDEEEF" & _
                   "F0F1F2F3F4F5F6F7F8F9FAFBFCFDFEFF", vbFromUnicode)

    b = StrConv(sChaine, vbFromUnicode)
    l = UBound(b)                                               ' Nombre de caract�res.
    If l = -1 Then Exit Function                                ' Chaine d'entr�e vide.

    ' On allonge de deux octets pour �viter les tests de d�bordement.
    ReDim Preserve b(l + 2)
    ReDim sQP((l + 1 + l \ 76) * 3)                             ' Pr�-allouer l'espace maximal n�cessaire.

    Do
        bCar = b(i)                                             ' Extraire un caract�re.

        Select Case bCar
            Case 33& To 60&, 62&, 64& To 94&, 96& To 126&       ' Caract�res � ne pas encoder.
                bEncode = 0&

            Case 63&, 95&                                       ' Cas du '?' et du '_'
                bEncode = Abs(QEncoding)

            Case 13&                                            ' CR.
                bEncode = 1&                                    ' Encoder, en principe,
                If b(i + 1&) = 10 Then bEncode = 2&             ' sauf si c'est un CRLF.

            Case 9&, 32&                                        ' Espace et tabulation.
                bEncode = 0&                                    ' Ne pas encoder, normalement

                If QEncoding And bCar = 32& Then                ' Lorsque QEncoding = True, Espace --> '_'
                    bCar = 95&

                Else                                            ' sauf si suivi d'un CRLF ou en fin de chaine.
                    If b(i + 1&) = 13 And b(i + 2&) = 10 Or i = l Then bEncode = 1&

                End If

            Case Else                                           ' Autre caract�re.
                bEncode = 1&                                    ' Encoder.

        End Select

        ' Longueur maxi de ligne, Soft Break compris : 76.
        Select Case bEncode
            Case 0&:    k = 75&                                 ' On va ajouter 1 car.
            Case 1&:    k = 73&                                 ' On va ajouter 3 car.
            Case 2&:    k = 74&                                 ' On va ajouter 2 car.
        End Select

        ' Ins�rer un 'Soft Break' si QEncoding est 'Faux', pour �viter de tronquer les objets de messages
        ' et les noms de pi�ces jointes.
        If n >= k And Not QEncoding Then
            ' Si le caract�re courant est un point (46) suivi de CRLF, on n'ins�re pas
            ' de SoftBreak, car CRLF . CRLF serait interpr�t� comme la fin des DATA.
            If Not (b(i) = 46 And b(i + 1) = 13 And b(i + 2) = 10) Then
                sQP(j) = 61                                     ' "="
                sQP(j + 1&) = 13                                ' vbCrLf
                sQP(j + 2&) = 10
                j = j + 3&
            End If
            n = 0&                                              ' Caract�res d'une ligne.
        End If

        ' Proc�der � l'encodage.
        Select Case bEncode
            Case 0&, 2&                                         ' Ne pas encoder le caract�re.
                sQP(j) = bCar                                   ' Caract�re non encod�.
                j = j + 1&                                      ' Position de sortie suivante.
                n = n + 1&                                      ' Caract�res d'une ligne.

                ' C'est un CRLF, on ajoute le LF non encod�.
                If bEncode = 2& Then
                    i = i + 1&
                    sQP(j) = b(i)                               ' Caract�re non encod�.
                    j = j + 1&                                  ' Position de sortie suivante.
                    n = 0&                                      ' Caract�res d'une ligne.
                End If

            Case 1                                              ' Encoder le caract�re.
                sQP(j) = 61:            j = j + 1&              ' '='

                k = bCar * 2&
                sQP(j) = bASC(k):       j = j + 1&

                sQP(j) = bASC(k + 1&):  j = j + 1&

                n = n + 3&                                      ' Caract�res d'une ligne.

        End Select

        i = i + 1&                                              ' Caract�re d'entr�e suivant.
    Loop While i <= l

    ReDim Preserve sQP(j - 1&)                                  ' Ne garder que la partie utile.
    Enc_QP = StrConv(sQP, vbUnicode)
End Function

' Reformate la chaine en ligne de 998 caract�res maxi hors CRLF.
' La chaine d'entr�e est suppos�e �tre en 7bit.
' Aucun contr�le n'est effectu�.
Function Enc_TXT(sChaine As String) As String
    Dim bi() As Byte, bo() As Byte, l As Long, i As Long, j As Long, n As Long

    bi = StrConv(sChaine, vbFromUnicode)
    l = UBound(bi)
    If l = -1 Then Exit Function                                ' Chaine d'entr�e vide.

    ReDim Preserve bi(l + 1)                                    ' Pour �viter les tests de d�bordement.
    ReDim bo(l + 2& * l \ 998&)

    Do While i <= l
        ' Si on passe sur un CRLF, on remet le compteur � z�ro.
        If bi(i) = 13 Then If bi(i + 1&) = 10 Then n = -1&

        bo(j) = bi(i)                                           ' Copier le caract�re.
        i = i + 1&
        j = j + 1&
        n = n + 1&

        If n = 999& Then                                        ' Longueur maximale d'une ligne.
            bo(j) = 13: j = j + 1&
            bo(j) = 10: j = j + 1&
            n = 0&
        End If
    Loop

    If j > 0 Then ReDim Preserve bo(j - 1&)
    Enc_TXT = StrConv(bo, vbUnicode)
End Function

' Retourne la repr�sentation hexad�cimale de la chaine MD5
Function myHEX(Param As String) As String
    Dim b() As Byte, i As Long, s As String

    b = StrConv(Param, vbFromUnicode)

    myHEX = Space$(2 * (UBound(b) + 1))
    For i = 0 To UBound(b)
        s = Hex$(b(i))
        If Len(s) < 2 Then s = "0" & s                          ' S'assurer qu'on a bien deux caract�res.
        Mid$(myHEX, 2 * i + 1, 2) = LCase$(s)
    Next i

    Erase b
End Function

' Convertit une cha�ne de caract�res Unicode en son �quivalent UTF8.
Function UaUTF8(sChaine As String) As String
    Dim i As Long, l As Long, j As Long
    Dim bIn() As Byte, bOut() As Byte, b() As Byte

    bIn = sChaine
    l = UBound(bIn)
    ReDim bOut((l + 1) * 2)                                     ' Pr�-allouer de l'espace pour le pire des cas.

    For i = 0 To l Step 2
        Call UTF8Car(256& * bIn(i + 1) + bIn(i), b)             ' Convertir le caract�re en UTF-8

        Select Case UBound(b)
            Case 0                                              ' Sur 1 octet
                bOut(j + 0) = b(0)
                j = j + 1

            Case 1                                              ' Sur 2 octets
                bOut(j + 0) = b(0)
                bOut(j + 1) = b(1)
                j = j + 2

            Case 2                                              ' Sur 3 octets
                bOut(j + 0) = b(0)
                bOut(j + 1) = b(1)
                bOut(j + 2) = b(2)
                j = j + 3

            Case 3                                              ' Sur 4 octets
                bOut(j + 0) = b(0)
                bOut(j + 1) = b(1)
                bOut(j + 2) = b(2)
                bOut(j + 3) = b(3)
                j = j + 4

        End Select
    Next i

    Erase bIn, b                                                ' Lib�rer de la m�moire

    ReDim Preserve bOut(j - 1)                                  ' Tronquer � la partie utile.
    UaUTF8 = StrConv(bOut, vbUnicode)

    Erase bOut
End Function

' Convertit une cha�ne UTF8 en une cha�ne Unicode.
' Un caract�re invalide (248 � 255) dans la chaine sera simplement ignor�.
Function UTF8aU(sChUTF8 As String) As String
    Dim l As Long, i As Long, j As Long, lx As Long, lCar As Long
    Dim b() As Byte, bOut() As Byte

    If Len(sChUTF8) = 0 Then Exit Function

    b = StrConv(sChUTF8, vbFromUnicode)
    l = UBound(b)
    ReDim bOut((l + 1) * 2)

    For i = 0& To l
        Select Case b(i)
            Case 0 To 127                                       ' Caract�re ASCII normal
                lx = 1&                                         ' 1 seul octet
                lCar = b(i)

            Case 192 To 223                                     ' 1er car., 2 octets
                lx = 2&
                lCar = (b(i) And &H1F&) * &H40&                 ' Placer les 5 bits utiles

            Case 224 To 239                                     ' 1er car., 3 octets
                lx = 3&
                lCar = (b(i) And &HF&) * &H1000&                ' Placer les 4 bits utiles

            Case 240 To 247                                     ' 1er car., 4 octets
                lx = 4&
                lCar = (b(i) And &H7&) * &H40000                ' Placer les 3 bits utiles

            Case 128 To 191                                     ' Car. suivant
                lx = lx - 1
                Select Case lx
                    Case 1: lCar = lCar Or (b(i) And &H3F&)     ' 6 derniers bits
                    Case 2: lCar = lCar Or (b(i) And &H3F&) * &H40&
                    Case 3: lCar = lCar Or (b(i) And &H3F&) * &H1000&
                End Select
        End Select

        If lx = 1& Then
            ' On a le code Unicode complet, on peut l'�crire dans la chaine
            ' Si lCar est sup�rieur � 65535, le fichier d'entr�e n'�tait pas en UTF-8.
            If lCar <= &HFFFF& Then
                bOut(j) = lCar Mod 256&
                bOut(j + 1) = lCar \ 256&
            End If
            j = j + 2&                                          ' Position de sortie suivante
            lCar = 0&
            lx = 0&
        End If
    Next i

    Erase b

    ReDim Preserve bOut(j - 1)
    UTF8aU = bOut

    Erase bOut
End Function