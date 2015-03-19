Option Compare Database
Option Explicit
Option Private Module


' Copyright 2009-2012 Denis SCHEIDT
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

Type tuCorpsMIME
    ContentType                 As String
    ContentID                   As String
    Charset                     As String
    Name                        As String
    ContentTransferEncoding     As String
    ContentDisposition          As String
    FileName                    As String
    Donnee                      As String
End Type


' Analyse le corps MIME d'un message et renseigne un tableau de type tuCorpsMIME
' Le tableau comporte une ligne pour chaque partie de message.
' Lorsque bText=True, seule la partie text/plain est extraite.
Sub MIMEAnalyse(sMsgMIME As String, dtuCM() As tuCorpsMIME, _
                Optional psDelim As String = "")
    Dim sDelim As String, d As Long, f As Long

    ' Chercher un multipart
    d = InStr(sMsgMIME, "Content-Type: multipart/")

    If d > 0 Then                                               ' Traiter multipart.
        ' R�cup�rer le d�limiteur.
        d = InStr(d + 24, sMsgMIME, "boundary=") + 9            ' Premier caract�re du d�limiteur.
        f = InStr(d, sMsgMIME, vbCrLf) - 1                      ' Dernier caract�re du d�limiteur.
        sDelim = Mid$(sMsgMIME, d, f - d + 1)                   ' D�limiteur.
        ' Si le d�limiteur est encadr� par des ", on les �limine.
        If sDelim Like """*""" Then sDelim = Mid$(sDelim, 2, f - d - 1)
        sDelim = vbCrLf & "--" & sDelim                         ' Compl�ter le d�limiteur.

        ' Trouver le d�but de la partie (premier caract�re apr�s d�limiteur).
        d = InStr(f + 1, sMsgMIME, sDelim) + Len(sDelim) + 2

        ' Trouver la fin de la partie (dernier caract�re).
        f = InStr(d, sMsgMIME, sDelim & "--") - 1
        If f < 0 Then f = Len(sMsgMIME)

        ' Traiter r�cursivement la partie.
        Call MIMEAnalyse(Mid$(sMsgMIME, d, f - d + 1), dtuCM(), sDelim)

        ' Se placer apr�s la partie qui vient d'�tre trait�e r�cursivement.
        d = f + Len(sDelim) + Len(psDelim) + 7

    Else
        d = 1

    End If

    If Len(psDelim) = 0 And d <> 1 Then Exit Sub

    ' Traiter parties correspondant � un d�limiteur.
    Do
        f = InStr(d, sMsgMIME, psDelim) - 1                     ' Chercher la fin de la partie.
        If f < 1 Then f = Len(sMsgMIME)                         ' D�limiteur non trouv�.

        ' M�moriser les donn�es de la partie.
        If f > d Then Call MIMEAnalysePartie(Mid$(sMsgMIME, d, f - d + 1), dtuCM())

        d = f + Len(psDelim) + 3                                ' D�but de la partie suivante.
    Loop While d < Len(sMsgMIME)

End Sub

' Retourne la premi�re partie correspondant � ConteType et ContentDisposition.
Function MIMEPartie(sContentType As String, sContentDisposition As String, dtuCM() As tuCorpsMIME) As tuCorpsMIME
    Dim i As Integer

    For i = 0 To UBound(dtuCM)
        If dtuCM(i).ContentType = sContentType And dtuCM(i).ContentDisposition = sContentDisposition Then
            MIMEPartie = dtuCM(i)
            Exit For
        End If
    Next i
End Function

' D�code l'objet du message.
Function DecObjet(sChaine As String) As String
    Dim s As String, i As Long

    ' Extraire la partie utile de la chaine.
    If sChaine Like "=[?]*[?][BQ][?]*[?]=" Then
        i = InStr(sChaine, "?Q?")
        If i = 0 Then i = InStr(sChaine, "?B?")
        s = Mid$(sChaine, i + 3, Len(sChaine) - 4 - i)

    Else
        s = sChaine

    End If

    ' Effectuer le bon d�codage.
    If sChaine Like "*[?]Q[?]*" Then
        DecObjet = UTF8aU(Dec_QP(s, True))
    ElseIf sChaine Like "*[?]B[?]*" Then
        DecObjet = UTF8aU(Dec_Base64(s))
    Else
        DecObjet = s
    End If
End Function

' Extrait une portion text/plain du corps du message.
Function DecCorps(sCorps As String) As String
    Dim dtu() As tuCorpsMIME, s As tuCorpsMIME

    ReDim dtu(0)
    Call MIMEAnalyse(sCorps, dtu())                         ' Analyse le message.

    s = MIMEPartie("text/plain", "", dtu())                 ' Extraire la partie text/plain.
    DecCorps = Left$(DecPartie(s), 100)
End Function

' D�code une partie de message.
Function DecPartie(dtuCM As tuCorpsMIME) As String
    Select Case dtuCM.ContentTransferEncoding
        Case "7bit":                DecPartie = dtuCM.Donnee
        Case "quoted-printable":    DecPartie = Dec_QP(dtuCM.Donnee)
        Case "base64":              DecPartie = Dec_Base64(dtuCM.Donnee)
        Case Else:                  DecPartie = Traduit("dec_part", "libMAIL ne g�re pas le Content-Transfer-Encoding '%s'\n=== Partie non d�cod�e. ===", dtuCM.ContentTransferEncoding) & _
                                                vbCrLf & "" & vbCrLf & dtuCM.Donnee
    End Select

    Select Case dtuCM.Charset
        Case "utf-8":               DecPartie = UTF8aU(DecPartie)
    End Select
End Function


' Conversion HTML vers Texte.
' Tente de conserver un semblant de mise en forme...
Function HTMLaTexte(sChaine As String) As String
    ' Suppression de la pollution ajout�e par les traitements de textes.
    HTMLaTexte = SupprBalise(sChaine, "<HEAD>", "</HEAD>", True)

    ' Retirer les CRLF et les TAB.
    HTMLaTexte = Remplacer(HTMLaTexte, vbCr, " ")
    HTMLaTexte = Remplacer(HTMLaTexte, vbLf, " ")
    HTMLaTexte = Remplacer(HTMLaTexte, vbTab, " ")

    ' Supprimer les balises HTML.
    HTMLaTexte = SupprBalise(HTMLaTexte, "<HTML", ">", True)
'    HTMLaTexte = SupprBalise(HTMLaTexte, "<!DOCTYPE ", ">", True)
    HTMLaTexte = SupprBalise(HTMLaTexte, "<!--", "-->", True)
    HTMLaTexte = SupprBalise(HTMLaTexte, "<SCRIPT", "</SCRIPT>", True)
    HTMLaTexte = SupprBalise(HTMLaTexte, "<STYLE", "</STYLE>", True)

    ' Listes.
    HTMLaTexte = Remplacer(HTMLaTexte, "<LI><P", "<LI", , , vbTextCompare)
    HTMLaTexte = Remplacer(HTMLaTexte, "<LI", vbCrLf & "* <LI", , , vbTextCompare)
    ' Retours.
    HTMLaTexte = Remplacer(HTMLaTexte, "<BR>", vbCrLf & "<BR>", , , vbTextCompare)
    HTMLaTexte = Remplacer(HTMLaTexte, "<BR ", vbCrLf & "<BR ", , , vbTextCompare)
    HTMLaTexte = Remplacer(HTMLaTexte, "<P", vbCrLf & "<P", , , vbTextCompare)
    HTMLaTexte = Remplacer(HTMLaTexte, "<H", vbCrLf & "<H", , , vbTextCompare)
    ' Tableaux.
    HTMLaTexte = Remplacer(HTMLaTexte, "<TR", vbCrLf & "<TR", , , vbTextCompare)
    HTMLaTexte = Remplacer(HTMLaTexte, "<TD", vbTab & "<TD", , , vbTextCompare)

    ' Retirer toutes les balises restantes.
    HTMLaTexte = SupprBalise(HTMLaTexte, "<", ">", True)

    ' Caract�res sp�ciaux.
    If HTMLaTexte Like "*&*;*" Then Call RemplCarSpec(HTMLaTexte)

    ' Remplacer les espaces cons�cutifs.
    Do While HTMLaTexte Like "*  *"
        HTMLaTexte = Remplacer(HTMLaTexte, "  ", " ")
    Loop
    HTMLaTexte = Remplacer(HTMLaTexte, " " & vbCrLf, vbCrLf)

End Function



' Supprime toutes les occurrences de la balise, ainsi que son contenu (si bContenu).
Function SupprBalise(sChaine As String, sBDebut As String, sBFin As String, bContenu As Boolean) As String
    Dim iBD As Long, iBF As Long, iDeb As Long, iFin As Long, i As Long, l As Long

    If Len(sChaine) = 0 Then Exit Function

    iDeb = 1                                                            ' D�but de lecture.
    iBD = 1
    i = 1                                                               ' Position d'�criture.
    SupprBalise = Space$(Len(sChaine))                                  ' Pr�-allouer de l'espace.

    Do
        ' Chercher un bloc entre balises.
        iBD = InStr(iBD, sChaine, sBDebut)                              ' Chercher la balise de d�but.
        If iBD = 0 Then iBD = Len(sChaine) + 1
        iBF = InStr(iBD, sChaine, sBFin)                                ' Chercher la balise de fin.
        If iBF = 0 Then iBF = Len(sChaine) + 1

        ' Copier le bloc situ� AVANT la balise de d�but.
        iFin = iBD
        l = iFin - iDeb                                                 ' Longueur � copier.
        Mid$(SupprBalise, i, l) = Mid$(sChaine, iDeb, l)                ' Copier le bloc.
        i = i + l                                                       ' D�placer le pointeur de sortie.

        If Not bContenu Then
            ' Copier le bloc situ� ENTRE les balises, SANS les balises.
            l = iBF - (iBD + Len(sBDebut))                              ' Longueur � copier.
            If l > 0 Then
                Mid$(SupprBalise, i, l) = Mid$(sChaine, iBD + Len(sBDebut), l) ' Copier le bloc.
                i = i + l                                               ' D�placer le pointeur de sortie.
            End If

        End If

        ' Se replacer pour chercher les balises suivantes.
        iBD = iBF + IIf(Len(sBFin) = 0, 1, Len(sBFin))
        iDeb = iBD

    Loop While iDeb < Len(sChaine)

    SupprBalise = Left$(SupprBalise, i - 1)                             ' Garder la partie utile.
End Function






' Traite les s�quences sp�ciales (&...;) d'une chaine HTML.
' La chaine d'entr�e est modifi�e !
Private Sub RemplCarSpec(sChaine As String)
    Dim i As Long, j As Long, l As Long, sC As String, sR As String

    i = InStr(sChaine, "&")
    If i > 0 Then j = InStr(i, sChaine, ";")
    l = j - i + 1

    Do While l > 1 And l < 10
        sC = Mid$(sChaine, i, l)
        sR = ""

        Select Case sC
            Case "&acirc;":     sR = "�"
            Case "&agrave;":    sR = "�"
            Case "&amp;":       sR = "&"
            Case "&ccedil;":    sR = "�"
            Case "&eacute;":    sR = "�"
            Case "&ecirc;":     sR = "�"
            Case "&egrave;":    sR = "�"
            Case "&gt;":        sR = ">"
            Case "&lt;":        sR = "<"
            Case "&nbsp;":      sR = " "
            Case "&quot;":      sR = """"
            Case "&ugrave;":    sR = "�"
            Case Else:          If sC Like "&[#]*;" Then sR = Chr$(Mid$(sC, 3, Len(sC) - 3))
        End Select

        Mid$(sChaine, i, 2) = sR & "&"

        i = InStr(j, sChaine, "&")                                      ' Occurrence suivante.
        If i > 0 Then j = InStr(i, sChaine, ";") Else j = 0
        l = j - i + 1
    Loop

    sChaine = SupprBalise(sChaine, "&", ";", True)                      ' Nettoyer les r�sidus de s�quences.
End Sub

Private Function RemplTout(sChaine As String, sSep1 As String, sSep2 As String) As String
    RemplTout = Joindre(Scinder(sChaine, sSep1), sSep2)
End Function


' S�pare les diff�rents �l�ments de la partie et les m�morise dans un �l�ment de tableau.
 Sub MIMEAnalysePartie(ByVal sPartie As String, dtuCM() As tuCorpsMIME)
    Dim i As Long, l As Long, vET As Variant, v0 As Variant
    Dim sN As String, sV As String

    ' Extraire l'en-t�te.
    i = InStr(sPartie, vbCrLf & vbCrLf)
    sN = Left$(sPartie, i - 1)
    ' Pour simplifier le traitement, on remplace ':' par '=', et '; ' par CRLF.
    sN = Remplacer(sN, ": ", "=")
    sN = Remplacer(sN, "; ", vbCrLf)

    ' Scinder la chaine pour obtenir une ligne par en-t�te ou param�tre.
    vET = Scinder(sN, vbCrLf)

    ' Augmenter le tableau.
    i = UBound(dtuCM)
    If Not (i = 0 And Len(dtuCM(0).ContentType) = 0) Then
        i = i + 1
        ReDim Preserve dtuCM(i)
    End If

    ' Pour chaque ligne.
    For Each v0 In vET
        ' Extraire le nom de l'en-t�te.
        v0 = Trim(v0)
        l = InStr(v0, "=")
        If l > 0 Then
            sN = Trim$(Left$(v0, l - 1))
            sV = Trim$(Mid$(v0, l + 1))
            ' Si la valeur est encadr�e par des ", on les �limine.
            If sV Like """*""" Then sV = Mid$(sV, 2, Len(sV) - 2)

            Select Case sN
                Case "Content-Type":                dtuCM(i).ContentType = sV
                Case "Content-ID":                  dtuCM(i).ContentID = sV
                Case "Content-Transfer-Encoding":   dtuCM(i).ContentTransferEncoding = sV
                Case "Content-Disposition":         dtuCM(i).ContentDisposition = sV
                Case "Charset":                     dtuCM(i).Charset = sV
                Case "Name":                        dtuCM(i).Name = sV
                Case "FileName":                    dtuCM(i).FileName = sV
            End Select
        End If
    Next v0

    ' Donn�es de la partie.
    l = InStr(sPartie, vbCrLf & vbCrLf)
    If l > 0 Then dtuCM(i).Donnee = Mid$(sPartie, l + 4)
End Sub