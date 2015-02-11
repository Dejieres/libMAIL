Option Compare Database
Option Explicit
Option Private Module

' Copyright 2011-2013 Denis SCHEIDT
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


Public Const MAX32 As Currency = 4294967296@
Public Const MAX24 As Currency = 16777216@
Public Const MAX16 As Currency = 65536@
Public Const MAX08 As Currency = 256@


' Retourne une cha�ne compos�e de n octets al�atoires.
Function Alea(ByVal iNbOctets As Integer) As String
    Dim b() As Byte

    Randomize

    iNbOctets = iNbOctets * 2 - 1
    ReDim b(iNbOctets)
    iNbOctets = iNbOctets - 1
    Do
        b(iNbOctets) = Fix(Rnd * 256)
        iNbOctets = iNbOctets - 2
    Loop While iNbOctets >= 0
    Alea = b
End Function

' Les fonctions sur bits convertissent la valeur Currency sur 2 Long (32 bits vers 2 fois 16 bits),
' effectuent des op�rations partielles sur chaque Long,
' puis recollent les morceaux...
'
' AND sur des valeurs 32 bits, non sign�
Function And32(x As Currency, y As Currency) As Currency
    Dim x2 As Long, x1 As Long, y2 As Long, y1 As Long

    x2 = Int(x / MAX16):   x1 = x - x2 * MAX16
    y2 = Int(y / MAX16):   y1 = y - y2 * MAX16
    And32 = (x2 And y2) * MAX16 + (x1 And y1)               ' Recombiner les AND partiels
End Function

' Convertit un Currency en une chaine de 4 octets, poids FORTS en t�te (big endian).
Function CvC2O_be(x As Currency) As String
    CvC2O_be = Chr$(And32(x, 4278190080@) / 16777216) & _
               Chr$(And32(x, 16711680) / 65536) & _
               Chr$(And32(x, 65280) / 256) & _
               Chr$(And32(x, 255))
End Function

' Convertit un Currency en une chaine de 4 octets, poids FAIBLES en t�te (little endian)
Function CvC2O_le(x As Currency) As String
    CvC2O_le = Chr$(And32(x, 255)) & _
               Chr$(And32(x, 65280) / 256) & _
               Chr$(And32(x, 16711680) / 65536) & _
               Chr$(And32(x, 4278190080#) / 16777216)
End Function

' NOT sur des valeurs 32 bits, non sign�
Function Not32(x As Currency) As Currency
    Dim x2 As Long, x1 As Long

    x2 = Int(x / MAX16):   x1 = x - x2 * MAX16
    ' And &HFFFF& --> Reset des bits de poids fort, pour �viter de passer en n�gatif
    ' On ne garde que les 16 bits de poids faible. Les autres sont mis � 0.
    Not32 = ((Not x2) And &HFFFF&) * MAX16 + ((Not x1) And &HFFFF&) ' Recombiner
End Function

' OR sur des valeurs 32 bits, non sign�
Function Or32(x As Currency, y As Currency) As Currency
    Dim x2 As Long, x1 As Long, y2 As Long, y1 As Long

    x2 = Int(x / MAX16):   x1 = x - x2 * MAX16
    y2 = Int(y / MAX16):   y1 = y - y2 * MAX16
    Or32 = (x2 Or y2) * MAX16 + (x1 Or y1)                  ' Recombiner les OR partiels
End Function

' Modulo sur n bits, non sign�
Function myMod(x As Currency, Optional oBits As Byte = 32) As Currency
    Dim nMod As Currency

    ' Plus rapide que l'�l�vation � la puissance...
    ' Les plus fr�quents en premier...
    Select Case oBits
        Case 32:    nMod = MAX32
        Case 8:     nMod = MAX08
        Case 16:    nMod = MAX16
        Case Else:  nMod = 2@ ^ oBits
    End Select

    ' Pas la peine de calculer le modulo si la valeur est d�j� inf�rieure.
    Select Case x
        Case Is < nMod:     myMod = x
        Case Else:          myMod = x - Int(x / nMod) * nMod
    End Select
End Function

' Rotation � gauche d'une valeur
Function RotGauche(ByVal x As Currency, ByVal oRot As Integer, Optional oBits As Byte = 32) As Currency
    Dim bpF As Currency, cInterm As Currency, nDecal As Currency

    x = myMod(x, oBits)                             ' On commence par arrondir...
    oRot = oRot Mod oBits                           ' Pas la peine de faire plusieurs tours ;)

    nDecal = 2 ^ (oBits - oRot)
    ' Extraire les oRot bits de poids fort (par d�calage � droite)
    bpF = Int(x / nDecal)                           ' Division enti�re impossible, d�passement de capacit�...

    cInterm = bpF * nDecal                          ' Bits de poids fort, d�cal�s � gauche, le reste � 0
    cInterm = x - cInterm                           ' Retirer les bits de poids fort de la valeur d'origine
    cInterm = cInterm * 2 ^ oRot                    ' D�caler � gauche
    RotGauche = cInterm + bpF                       ' Ins�rer les bits de gauche par la droite.
End Function

' XOR sur des valeurs 32 bits, non sign�
Function XOr32(x As Currency, y As Currency) As Currency
    Dim x2 As Long, x1 As Long, y2 As Long, y1 As Long

    x2 = Int(x / MAX16):   x1 = x - x2 * MAX16
    y2 = Int(y / MAX16):   y1 = y - y2 * MAX16
    XOr32 = (x2 Xor y2) * MAX16 + (x1 Xor y1)       ' Recombiner les XOR partiels
End Function