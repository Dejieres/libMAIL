Option Compare Database
Option Explicit
Option Private Module


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


' Formatage d'une date pour la RFC821
Function DateMail(ByVal dt As Date) As String
    DateMail = JourSemaineUS(WeekDay(dt, vbMonday)) & ", " _
                & Format(Day(dt), "00") & " " & MoisUS(Month(dt)) & " " & Format(Year(dt), "00") _
                & " " & Format(dt, "hh:nn:ss") _
                & " " & FormatDH(DecalageUTC())
End Function

' Mois au format US
Function MoisUS(ByVal Mois As Integer) As String
    Dim v As Variant

    v = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    MoisUS = v(Mois - 1)
End Function

' Jour au format US
Function JourSemaineUS(ByVal Jour As Integer) As String
    Dim v As Variant

    v = Array("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
    JourSemaineUS = v(Jour - 1)
End Function

' Donne le décalage par rapport à UTC. La valeur de retour est en minutes.
Function DecalageUTC() As Integer
    Dim TZI As TIME_ZONE_INFORMATION, l As Long

    l = GetTimeZoneInformation(TZI)

    ' Additionner le fuseau horaire et le décalage heure d'été.
    With TZI
        DecalageUTC = -(.Bias + .StandardBias - (l = TIME_ZONE_DAYLIGHT) * .DaylightBias)
    End With
End Function

' Retourne le décalage horaire formaté "+hhnn" ou "-hhnn".
' iDecalage est en minutes. Tous les calculs sont faits sur des minutes.
Function FormatDH(iDecalage As Integer) As String
    Dim i As Single

    i = iDecalage Mod 1440                          ' Limiter à un tour de globe :)

    Select Case i
        Case Is < -720:     i = i + 1440
        Case -720 To 720:                           ' Rien, on garde la valeur
        Case Is > 720:      i = i - 1440
    End Select

    ' Conversion en heures et formatage.
    ' Cette syntaxe ne fonctionne plus à partir d'Access 2000,
    'FormatDH = Format$(i / 1440, "+hhnn;-hhnn;+hhnn")
    ' Format$ ne tenant compte que de la première section pour les dates.
    i = i / 1440
    FormatDH = Choose(Sgn(i) + 2, "-", "+", "+") & Format$(i, "hhnn")
End Function

' Retourne la date et l'heure système, à la milliseconde.
Function HoroDatage() As String
    Dim dtuST As SYSTEMTIME

    Call GetSystemTime(dtuST)

    With dtuST
        ' Heure système (UTC) ajustée en fonction du fuseau horaire.
        HoroDatage = DateAdd("n", DecalageUTC(), DateSerial(.wYear, .wMonth, .wDay) & " " & _
                                                 TimeSerial(.wHour, .wMinute, .wSecond)) & "." & _
                     Format$(.wMillisecond, "000")
    End With
End Function