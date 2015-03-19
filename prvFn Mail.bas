Option Compare Database
Option Explicit
Option Private Module

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





' --------------------------------------------------------------------
' TU pour la transmission des donn�es vers le formulaire frm_EditeMail
' Ce Type n'est pas export�.
Type tuMAIL
    Identifiant                 As String           ' Identifiant unique du message
    De                          As String           ' Exp�diteur
    a                           As String           ' Destinataires
    cc                          As String           ' Dest. en copie
    BCC                         As String           ' Dest. en copie cach�e
    Objet                       As String           ' Objet du message
    Message                     As tuMessageMIME    ' Corps du message
    PJ                          As Variant          ' Tableau des pi�ces jointes
    OptionsMSG                  As tuESMTP_MSG      ' Options �tendues de message
    Utilisateur                 As String           ' Utilisateur ayant cr�� le message
    Differer                    As Date             ' Envoi diff�r� du message
    Conserver                   As Date             ' Conserver le message 'V' jusqu'au...
End Type



' Type d'encodage pour les corps de messages et les pi�ces jointes.
Public Const lmlEncTXT As Long = 0&
Public Const lmlEncQP  As Long = 1&
Public Const lmlEncB64 As Long = 2&


' --------------------------------------------------------------------
' TU pour la gestion du multipart/related
Private Type tuMpRel
    bal         As String                           ' Balise d'origine
    src         As String                           ' Chaine src de la balise <IMG ...>
    cid         As String                           ' Content-ID pour l'image jointe au message.
    Nom         As String                           ' Nom du fichier image (sans extension).
    Ext         As String                           ' Type de fichier (extension).
    Doublon     As Boolean                          ' Fichier apparaissant plusieurs fois dans le corps HTML.
End Type

Public Const DELIM         As String = ";"          ' S�parateur d'adresses de messagerie





' Encode la DTU d'options de messages sous forme de cha�ne Base64,
' pour m�morisation dans la table
Function DTU_MSG_Enc(vDTU As tuESMTP_MSG) As String
    Dim sFichier As String, nFich As Integer, l As Long

    sFichier = FichTemp()
    nFich = FreeFile()
    Open sFichier For Binary Access Write As #nFich
    Put #nFich, , vDTU                              ' Ecrit la DTU dans le fichier
    Close #nFich                                    ' Ecrire le tampon sur le disque

    Open sFichier For Binary Access Read As #nFich
    l = LOF(nFich)
    DTU_MSG_Enc = Enc_Base64(Input(l, #nFich))      ' Relit la DTU dans une chaine et l'encode
    Close #nFich

    Kill sFichier                                   ' Nettoyage
End Function

' D�code la cha�ne Base64 et retourne une variable DTU
Function DTU_MSG_Dec(sDTU_B64 As String) As tuESMTP_MSG
    Dim sFichier As String, nFich As Integer, s As String

    sFichier = FichTemp()
    nFich = FreeFile()
    Open sFichier For Binary Access Write As #nFich
    s = Dec_Base64(sDTU_B64)                        ' Il faut une variable pour le Put
    Put #nFich, , s                                 ' Ecrit la DTU d�cod�e dans le fichier
    Close #nFich                                    ' Ecrire le tampon sur le disque

    Open sFichier For Binary Access Read As #nFich
    Get #nFich, , DTU_MSG_Dec                       ' Relit la DTU
    Close #nFich

    Kill sFichier                                   ' Nettoyage
End Function


' Retourne une chaine constitu�e des lignes d'en-t�te.
' Chaque ligne doit �tre termin�e par CR+LF.
' La date n'est fournie que par la proc�dure d'envoi (EnvoieTout).
' Pour l'export EML, elle est lue depuis la table.
Function MSGEnTete(rs As DAO.Recordset, ESMTP_MSG As tuESMTP_MSG, _
                   Optional bEML As Boolean = False, Optional ByVal dDate As Date = 0) As String
    Dim s As String

    ESMTP_MSG = DTU_MSG_Dec(Nz(rs!ESMTP))                           ' Lecture des options �tendues li�es au message.

    If dDate = 0 Then dDate = Nz(rs!DateEnvoi, 0)                   ' D�terminer la date d'envoi.
    MSGEnTete = "From: " & rs!Expediteur & vbCrLf & _
                "To: " & Delims(Nz(rs!Destinataires)) & vbCrLf & _
                "Date: " & DateMail(dDate) & vbCrLf

    If Not IsNull(rs!cc) Then MSGEnTete = MSGEnTete & "CC: " & Delims(rs!cc) & vbCrLf

    ' Pour l'export EML, on ajoute �galement les destinataires cach�s.
    If bEML Then
        If Not IsNull(rs!BCC) Then MSGEnTete = MSGEnTete & "BCC: " & Delims(rs!BCC) & vbCrLf
    End If

    With ESMTP_MSG.ORG
        ' Reponse, si n�cessaire.
        If Len(.Repondre) > 0 Then MSGEnTete = MSGEnTete & "Reply-To: " & Delims(.Repondre) & vbCrLf

        ' Envoyeur, si diff�rent de l'exp�diteur.
        If (Len(.Envoyeur) > 0) And (.Envoyeur <> rs!Expediteur) Then
            MSGEnTete = MSGEnTete & "Sender: " & .Envoyeur & vbCrLf
        End If
    End With

    ' Ajouter l'objet du message.
    If Not IsNull(rs!Objet) Then MSGEnTete = MSGEnTete & "Subject: " & rs!Objet & vbCrLf

    s = Trim$(ESMTP_MSG.MDN.Notification)                           ' Avis de (non-)lecture.
    If Len(s) > 0 Then MSGEnTete = MSGEnTete & "Disposition-Notification-To: " & s & vbCrLf

    s = Trim$(ESMTP_MSG.MDN.Reception)                              ' Accus� de r�ception.
    If Len(s) > 0 Then MSGEnTete = MSGEnTete & "Return-Receipt-To: " & s & vbCrLf

    ' En-t�tes personnalis�s (non normalis�s).
    MSGEnTete = MSGEnTete & "X-Mailer: " & Enc_nASCII("Biblioth�que VBA libMAIL v." & VersionProg()) & vbCrLf

    ' Priorit� du message.
    If ESMTP_MSG.Priorite <> 0 Then MSGEnTete = MSGEnTete & "X-Priority: " & ESMTP_MSG.Priorite & vbCrLf

    ' Identifiant (Message-Id).
    s = "<libMAIL_" & HoroDatage(True) & "@" & myComputerName & ">"
    MSGEnTete = MSGEnTete & "Message-Id: " & s & vbCrLf

    ' User-agent.
    MSGEnTete = MSGEnTete & "User-Agent: libMAILVBA/" & VersionProg() & vbCrLf
End Function

' Uniformise les d�limiteurs, et supprime les d�limiteurs cons�cutifs
Function Delims(ByVal sListe As String, Optional sDelim As String = DELIM) As String
    Dim i As Integer, i1 As Integer

    For i = 1 To Len(sListe)
        If Mid$(sListe, i, 1) Like "[,;]" Then
            If i1 = 0 Then
                Mid$(sListe, i, 1) = sDelim                         ' On �crit le s�parateur unique
            Else
                Mid$(sListe, i, 1) = " "                            ' On efface les s�parateurs cons�cutifs
            End If
            i1 = 1                                                  ' On vient de lire un s�parateur
        Else
            i1 = 0                                                  ' On est sur un caract�re 'normal'. On ne fait rien de plus
        End If
    Next i
    Delims = sListe
End Function


' Enregistre le message dans la table.
' Modifie la premi�re colonne du tableau des P.J. en cas d'erreur.
Sub SauveMail(dtuMail As tuMAIL)
    Dim rs As DAO.Recordset, SQL As String
    Dim sDelim As String, sCorps As String, sPartiePJ As String, sPartieIMG As String
    Dim i As Integer, j As Long

    ' Commencer le corps du message
    sCorps = "MIME-Version: 1.0 (" & Enc_nASCII("Biblioth�que VBA libMAIL v." & VersionProg()) & ")" & vbCrLf


    ' -- MULTIPART/MIXED ------------------------------------------------------------------------------------
    '
    ' D�terminer s'il y a des pi�ces jointes
    If Not IsEmpty(dtuMail.PJ) Then
        For i = 0 To UBound(dtuMail.PJ, 2)
            If (Len(dtuMail.PJ(0, i)) <> 0) Then                ' On ignore les PJ qui n'ont pas de nom
                j = 1
                Exit For
            End If
        Next i
    End If

    ' S'il y a au moins une pi�ce jointe,
    If j > 0 Then
        ' Ajouter un en-t�te multipart/mixed.
        If Len(sDelim) <> 0 Then sCorps = sCorps & vbCrLf & "--" & sDelim & vbCrLf
        sDelim = Delimiteur()                                   ' Nouveau d�limiteur pour la partie.

        sCorps = sCorps & "Content-Type: multipart/mixed; boundary=""" & sDelim & """" & vbCrLf & _
                 vbCrLf & _
                 "Ce message au format MIME comporte plusieurs parties." & vbCrLf & _
                 vbCrLf

        ' Pr�parer la partie PJ.
        sPartiePJ = PJ_Partie(dtuMail.PJ, sDelim)
    End If


    ' -- MULTIPART/RELATED ----------------------------------------------------------------------------------
    '
    ' Chercher la pr�sence d'au moins une balise IMG dans le corps HTML.
    If InStr(dtuMail.Message.HTML, "<IMG ") > 0 Then
        If Len(sDelim) <> 0 Then sCorps = sCorps & vbCrLf & "--" & sDelim & vbCrLf

        sDelim = Delimiteur()                                   ' Nouveau d�limiteur pour la partie.
        ' Ajouter un en-t�te multipart/related
        sCorps = sCorps & "Content-Type: multipart/related; boundary=""" & sDelim & """" & vbCrLf

        ' Pr�parer la partie images incorpor�es.
        sPartieIMG = IMG_Partie(dtuMail.Message.HTML, sDelim)
    End If


    ' -- MULTIPART/ALTERNATIVE ------------------------------------------------------------------------------
    '
    ' Si le message comporte plusieurs versions du texte (autre que texte brut)
    If Len(dtuMail.Message.HTML) <> 0 Then
        If Len(sDelim) <> 0 Then sCorps = sCorps & vbCrLf & "--" & sDelim & vbCrLf

        sDelim = Delimiteur()                                   ' Nouveau d�limiteur pour la partie.
        ' Ajouter un en-t�te multipart/alternative
        sCorps = sCorps & "Content-Type: multipart/alternative; boundary=""" & sDelim & """" & vbCrLf

    End If

    ' ------------------------------------------------------------------------------------------------------
    ' Ajouter le text/plain. Il est toujours ajout�, m�me vide.
    sCorps = sCorps & TEXT_Ajouter(sDelim, "text/plain", dtuMail.Message.Texte)

    ' Ajouter la partie text/html si n�cessaire.
    If Len(dtuMail.Message.HTML) <> 0 Then
        sCorps = sCorps & TEXT_Ajouter(sDelim, "text/html", dtuMail.Message.HTML)

        ' D�limiteur de fin de la partie alternative.
        sCorps = sCorps & vbCrLf & "--" & sDelim & "--" & vbCrLf
    End If


    ' -------------------------------------------------------------------------------------------------------
    ' ----- Enregistrement proprement dit dans la table.
    SQL = "SELECT * FROM " & TableMail() & " WHERE Identifiant='" & dtuMail.Identifiant & "'"
    Set rs = CurrentDb.OpenRecordset(SQL, dbOpenDynaset)
    With rs
        If .RecordCount = 0 Then                                ' Non trouv�, on cr�e
            .AddNew
            dtuMail.Identifiant = IDMail()                      ' Identifiant unique du message
            !Identifiant = dtuMail.Identifiant
        Else
            .MoveFirst
            .Edit                                               ' On modifie
        End If

        ' Nom de l'utilisateur, s'il n'est pas fourni
        If Len(dtuMail.Utilisateur) = 0 Then !Utilisateur = myCurrentUser() Else !Utilisateur = dtuMail.Utilisateur
        !DateMsg = Now()
        !Etat = "E"                                             ' Message en Bo�te d'envoi
        !Expediteur = dtuMail.De
        !Destinataires = dtuMail.a
        !cc = IIf(Len(dtuMail.cc) = 0, Null, dtuMail.cc)
        !BCC = IIf(Len(dtuMail.BCC) = 0, Null, dtuMail.BCC)
        !Objet = IIf(Len(dtuMail.Objet) = 0, Null, Enc_nASCII(dtuMail.Objet))
        !CorpsMsg.AppendChunk sCorps
        !CorpsMsg.AppendChunk sPartieIMG
        !CorpsMsg.AppendChunk sPartiePJ
        !ESMTP = DTU_MSG_Enc(dtuMail.OptionsMSG)                ' Options �tendues
        !Differer = dtuMail.Differer
        !Conserver = dtuMail.Conserver
        !DateEnvoi = 0
        .Update

        .Close
    End With
    Set rs = Nothing
End Sub

' Retourne la cha�ne UTF-8 pour un caract�re Unicode
' Le caract�re converti est retourn� dans b()
Sub UTF8Car(lUnicode As Long, b() As Byte)
    Dim bNbOctets As Integer, i As Integer, lMasque As Long

    ' D�terminer la longueur de la chaine UTF8
    Select Case lUnicode
        Case Is < 128&
            ReDim b(0)
            b(0) = lUnicode
            Exit Sub
        Case 128& To 2047&:                     bNbOctets = 2: i = 1
        Case 2048& To 55295, 57344 To 65533:    bNbOctets = 3: i = 2
        Case 65536 To 2097151:                  bNbOctets = 4: i = 3
        Case Else                                           ' Caract�re invalide n'ayant pas �t� encod�.
            ReDim b(8)                                      ' Dimension 'invalide', caract�re non encod�.
            Exit Sub
    End Select

    ReDim b(i)
    lMasque = &H1&
    Do
        b(i) = &H80& + ((lUnicode And (lMasque * &H3F&)) \ lMasque)
        lMasque = lMasque * &H40&
        i = i - 1
    Loop While i >= 0

    Select Case bNbOctets                                   ' Ajuster le premier octet
        Case 2: b(0) = b(0) Or &HC0&
        Case 3: b(0) = b(0) Or &HE0&
        Case 4: b(0) = b(0) Or &HF0&
    End Select
End Sub

' Encode les chaines contenant des caract�res non ASCII
' Utilis� pour l'objet du message et les noms des pi�ces jointes.
' Section encoded-word de la RFC2047.
Function Enc_nASCII(sObjet As String) As String
    Dim s As String

    s = Enc_QP(UaUTF8(sObjet), True)
    If StrComp(s, sObjet) <> 0 Then
'        Enc_nASCII = "=?utf-8?B?" & Enc_Base64(sObjet) & "?="
        Enc_nASCII = "=?utf-8?Q?" & s & "?="

    Else
        Enc_nASCII = sObjet

    End If
End Function

' Encode une cha�ne en xtext.
' Tous les caract�res � l'ext�rieur de l'intervalle 33-126, ainsi que '+' et '=' sont
' encod�s sous la forme '+NN'.
'
' bUSASCII = True : tous les caract�res non USASCII sont ignor�s et �limin�s de la chaine.
'                   pour l'encodage de l'IDEnveloppe.
Function Enc_XText(sChaine As String, Optional bUSASCII As Boolean = False) As String
    Dim XText() As Byte, bCar As Long, i As Long, l As Long, j As Long
    Dim b() As Byte

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
    ReDim XText((l + 1) * 3)                                    ' Pr�-allouer l'espace maximal n�cessaire.

    Do
        bCar = b(i)                                             ' Extraire un caract�re.

        Select Case bCar
            Case 33& To 42&, 44& To 60&, 62& To 126&            ' Caract�res � ne pas encoder (de ! � ~, sauf + et =).
                XText(j) = bCar
                j = j + 1                                       ' Position d'�criture suivante

            Case Else                                           ' Autre caract�re, � encoder.
                If bUSASCII And (bCar = 43& Or bCar = 61&) Or Not bUSASCII Then
                    XText(j) = 43&                              ' Ins�rer un '+'.
                    XText(j + 1&) = bASC(bCar * 2&)             ' Premier car. Hex.
                    XText(j + 2&) = bASC(bCar * 2& + 1&)        ' Second car. Hex.
                    j = j + 3&
                End If

        End Select

        i = i + 1&                                              ' Caract�re d'entr�e suivant.
    Loop While i <= l

    ReDim Preserve XText(j - 1&)                                ' Ne garder que la partie utile.
    Enc_XText = StrConv(XText, vbUnicode)
End Function

' Retourne un identifiant unique construit � partir de l'heure courante.
Function IDMail(Optional ByVal dDateHeure As Date = 0) As String
    Randomize Timer
    If dDateHeure = 0 Then dDateHeure = Now()
    IDMail = Format$(dDateHeure, "yyyymmddhhnnss") & Left$(Format$(Rnd() * 10000, "0000"), 4)
End Function

' D�termine le type d'encodage le plus compact.
' lmlEncTXT, lmlEncB64 ou lmlEncQP
Function TypeEnc(sChaine As String) As Long
    Dim b() As Byte, lASCII As Long, lNASCII As Long, l As Long, i As Long
    Dim bCar As Long

    b = StrConv(sChaine, vbFromUnicode)                         ' Transforme la chaine en un tableau d'octets.
    l = UBound(b)

    ' Parcourt tout le tableau pour d�terminer la r�partition entre caract�res ASCII et non ASCII.
    Do While i <= l
        bCar = b(i)
        Select Case bCar
            Case 9&, 32& To 126&
                lASCII = lASCII + 1&                            ' ASCII

            Case 13&
                bCar = b(i + 1&)
                If bCar = 10& Then
                    lASCII = lASCII + 2&
                    i = i + 1&
                End If

        End Select

        i = i + 1&
    Loop

    lNASCII = UBound(b) + 1& - lASCII                           ' Le reste est non ASCII.

    If lNASCII = 0& Then
        TypeEnc = lmlEncTXT                                     ' ASCII pur.
    ElseIf (lASCII / lNASCII) < 5& Then
        TypeEnc = lmlEncB64                                     ' Beaucoup de non-ASCII.
    Else
        TypeEnc = lmlEncQP                                      ' Peu de non-ASCII.
    End If
End Function

' Charge les 64000 premiers caract�res du fichier HTML pass� en param�tre.
' Supprime la balise <HEAD>...</HEAD>
' Pr�-traite les balises <IMG en �crivant le chemin absolu dans le champ src.
Function HTMLCharge(sURL As String) As String
    Dim i As Long, n As Long, dtuMR() As tuMpRel, s As String, sRep As String, sBal As String
    Dim bUTF As Boolean

    ' Extraire le r�pertoire de l'URL
    i = Len(sURL)
    Do While Mid$(sURL, i, 1) <> "\"
        i = i - 1
        If i = 0 Then
            i = 1
            Exit Do
        End If
    Loop
    sRep = Left$(sURL, i - 1)

    ' Charger le fichier HTML.
    s = PJFichier(sURL, 64000)

    ' Le fichier HTML est-il en UTF-8 ?
    bUTF = s Like "*charset=utf-8*"

    ' Suppression de la pollution.
    s = SupprBalise(s, "<HEAD>", "</HEAD>", True)


    ' Les balises sp�ciales plac�es dans des commentaires par Word emp�chent Outlook d'afficher correctement
    ' une image incorpor�e.
    s = SupprBalise(s, "<!--", "-->", True)

    n = IMG_Balises(s, dtuMR())                                 ' Extraire toutes les balises IMG.

    For i = 0 To n
        ' Modifier la chaine HTML et cr�er les parties.
        ' Cr�er la nouvelle balise.
        sBal = Remplacer(dtuMR(i).bal, dtuMR(i).src, sRep & "\" & CorrigeURL(dtuMR(i).src), , 1, vbTextCompare)
        ' Remplacer l'ancienne balise par la nouvelle.
        s = Remplacer(s, dtuMR(i).bal, sBal, , 1, vbTextCompare)
    Next i

    ' Recr�er une balise <HEAD> simplifi�e.
    s = "<HEAD>" & vbCrLf & _
        "<META HTTP-EQUIV=""CONTENT-TYPE"" CONTENT=""text/html; charset=utf-8"">" & vbCrLf & _
        "</HEAD>" & vbCrLf & _
        s
    If bUTF Then
        HTMLCharge = UTF8aU(s)                                  ' Conversion UTF8 � Unicode.
    Else
        HTMLCharge = s                                          ' Sinon, on ne convertit pas.
    End If
End Function

' S�pare le Chemin, le nom et l'extension de la sp�cification de fichier.
' Retourne Vrai si la sp�cification a pu �tre s�par�e.
Function AnaSpecFich(sSpecFich As String, sChem As String, sNom As String, sExt As String) As Boolean
    Dim i As Long, j As Long

    ' Eliminer les URL.
    If sSpecFich Like "*://*.*/*" Then Exit Function

    AnaSpecFich = True

    ' Chercher le dernier antislash.
    i = InStrFin(sSpecFich, "\", , vbBinaryCompare)

    sChem = Left$(sSpecFich, i)                                 ' A gauche, c'est le chemin.

    ' Chercher le dernier point de la cha�ne.
    j = InStrFin(sSpecFich, ".", , vbBinaryCompare)

    If j > i Then                                               ' Il y a bien un '.' � droite du dernier '\'.
        sNom = Mid$(sSpecFich, i + 1, j - i - 1)                ' Nom � gauche.
        sExt = Mid$(sSpecFich, j + 1)                           ' Extension � droite.

    Else
        sNom = Mid$(sSpecFich, i + 1)                           ' Pas d'extension.

    End If
End Function

' Exporte un objet Access vers un fichier temporaire.
' Retourne le code d'erreur.
' Le param�tres est modifi� par la fonction (retourne la sp�cification de fichier temporaire).
Function PJOA_GenFichier(sNomFichier As String) As Integer
    Dim vObjXS  As Variant

    vObjXS = Scinder(sNomFichier, "/")
    sNomFichier = FichTemp("OBJXS")                                 ' Nom pour le fichier temporaire

    On Error Resume Next
    If vObjXS(0) = 1 Then
        Application.SaveAsText vObjXS(1), vObjXS(3), sNomFichier    ' Exporter la d�finition
    Else
        DoCmd.OutputTo vObjXS(1), vObjXS(3), vObjXS(2), sNomFichier ' Exporter les donn�es
    End If

    PJOA_GenFichier = Err.Number                                    ' Erreur �ventuelle

    On Error GoTo 0
End Function


' ============================================================================================================
'
' Ajoute une partie texte au corps du message.
Private Function TEXT_Ajouter(sDelim As String, sContentType As String, sTexte As String) As String
    Dim sTxtCorps As String, lTE As Long, sCT As String, sCTE As String

    ' D�terminer le type d'encodage � utiliser
    lTE = TypeEnc(sTexte)

    Select Case lTE
        Case lmlEncTXT
            sCT = sContentType & "; charset=""us-ascii"""
            sCTE = "7bit"
            sTxtCorps = Enc_TXT(sTexte)

        Case lmlEncQP
            sCT = sContentType & "; charset=""utf-8"""
            sCTE = "quoted-printable"
            sTxtCorps = Enc_QP(UaUTF8(sTexte))

    End Select

    TEXT_Ajouter = IIf(Len(sDelim) <> 0, vbCrLf & "--" & sDelim & vbCrLf, "") & _
                   "Content-Type: " & sCT & vbCrLf & _
                   "Content-Transfer-Encoding: " & sCTE & vbCrLf & _
                    vbCrLf & _
                    sTxtCorps
End Function

' Construit la partie des pi�ces jointes.
Private Function PJ_Partie(PJ As Variant, sDelim As String) As String
    Dim sNomPJ As String, sNomFichier As String, i As Integer, j As Long

    ' Traiter les pi�ces jointes
    For i = 0 To UBound(PJ, 2)                                  ' Ajout des pi�ces jointes
        If (Len(PJ(0, i)) <> 0) Then                            ' On ignore les PJ qui n'ont pas de nom
            sNomPJ = PJ(0, i)
            j = InStr(sNomPJ, ":")                              ' Y'a-t-il d�j� un code d'erreur ?
            If j <> 0 Then sNomPJ = Mid$(sNomPJ, j + 1)         ' Oter le code erreur existant

            sNomFichier = PJ(1, i)                              ' Chemin et nom du fichier

            j = 0                                               ' Indicateur d'erreur
            If sNomFichier Like "#/*/*/*" Then
                ' Joindre un objet Access ---------------------------------------------------------
                j = PJOA_GenFichier(sNomFichier)

            Else
                ' Joindre un fichier disque -------------------------------------------------------
                If Not FichierExiste(sNomFichier) Then          ' Le fichier existe ?
                    On Error Resume Next
                    j = GetAttr(sNomFichier)
                    j = Err.Number                              ' Pourquoi on ne peut pas joindre le fichier ?
                    On Error GoTo 0
                End If

            End If

            If j = 0 Then                                       ' Si pas d'erreur...
                ' Ajouter la pi�ce jointe au corps du message
                PJ_Partie = PJ_Partie & PJ_Ajouter(sDelim, sNomPJ, sNomFichier)

            Else
                ' La colonne 0 (Nom de la PJ) est pr�fix�e avec le code d'erreur
                PJ(0, i) = Format$(j, "00000") & ":" & sNomPJ

            End If
            If PJ(1, i) Like "#/*/*/*" Then                         ' Dans le cas d'une PJOA,
                If FichierExiste(sNomFichier) Then Kill sNomFichier ' effacer le fichier temporaire.
            End If

        End If
    Next i

    ' D�limiteur de fin de partie.
    PJ_Partie = PJ_Partie & vbCrLf & "--" & sDelim & "--"
End Function

' Ajoute un fichier de pi�ce jointe, encod� en Base64, au corps du message.
Private Function PJ_Ajouter(sDelim As String, sNomPJ As String, sNomFichier As String) As String
    Dim sContenuPJ As String, lTE As Long, sCT As String, sCTE As String

    ' Lire le fichier
    sContenuPJ = PJFichier(sNomFichier)

    ' D�terminer le type d'encodage � utiliser
    lTE = TypeEnc(sContenuPJ)

    Select Case lTE
        Case lmlEncTXT
            sCT = "text/plain; charset=""us-ascii"""
            sCTE = "7bit"
            sContenuPJ = Enc_TXT(sContenuPJ)

        Case lmlEncQP
            sCT = "text/plain; charset=""utf-8"""
            sCTE = "quoted-printable"
            sContenuPJ = Enc_QP(UaUTF8(sContenuPJ))

        Case lmlEncB64
            sCT = "application/octet-stream"
            sCTE = "base64"
            sContenuPJ = Enc_Base64(sContenuPJ)

    End Select

    PJ_Ajouter = vbCrLf & "--" & sDelim & vbCrLf & _
                 "Content-Type: " & sCT & "; name=""" & Enc_nASCII(sNomPJ) & """" & vbCrLf & _
                 "Content-Transfer-Encoding: " & sCTE & vbCrLf & _
                 "Content-Disposition: attachment; filename=""" & Enc_nASCII(sNomPJ) & """" & vbCrLf & _
                 vbCrLf & _
                 sContenuPJ
End Function

' Gestion des images incorpor�es dans un corps HTML.
Private Function IMG_Partie(sHTML As String, sDelim As String) As String
    Dim dtuMR() As tuMpRel, i As Long, j As Long, n As Long
    Dim sBal As String, sCID As String

    ' R�cup�rer toutes les balises IMG de la chaine HTML.
    n = IMG_Balises(sHTML, dtuMR())

    ' G�n�rer les Content-ID, de mani�re unique pour des src identiques.
    ' Une m�me image ne sera jointe qu'une fois au message, m�me si elle y
    ' appara�t plusieurs fois.
    For i = 0 To n
        If Len(dtuMR(i).cid) = 0 Then                   ' Si cet �l�ment n'a pas encore �t� initialis�.
            sCID = Delimiteur("Alternative.") & "@" & myComputerName()

            ' Parcourir tout le tableau � partir de la position courante et
            ' mettre � jour le cid, si ce n'est d�j� fait, lorsque qu'il a le m�me src.
            For j = i To n
                If dtuMR(j).src = dtuMR(i).src Then     ' src identique � l'�l�ment courant.
                    dtuMR(j).cid = sCID
                    dtuMR(j).Doublon = (j > i)          ' Marquer les fichiers en double.
                End If
            Next j
        End If

        ' Modifier la chaine HTML et cr�er les parties.
        ' Cr�er la nouvelle balise.
        sBal = Remplacer(dtuMR(i).bal, dtuMR(i).src, "cid:" & dtuMR(i).cid, , 1, vbTextCompare)
        ' Remplacer l'ancienne balise par la nouvelle.
        sHTML = Remplacer(sHTML, dtuMR(i).bal, sBal, , 1, vbTextCompare)

        ' N'ajouter qu'une fois la partie pour une m�me src.
        If Not dtuMR(i).Doublon Then IMG_Partie = IMG_Partie & IMG_Ajouter(sDelim, dtuMR(i))
    Next i

    ' D�limiteur de fin de partie.
    IMG_Partie = IMG_Partie & vbCrLf & "--" & sDelim & "--"
End Function

' Ajoute une image incorpor�e.
Private Function IMG_Ajouter(sDelim As String, dtuMR As tuMpRel) As String
    With dtuMR
        IMG_Ajouter = vbCrLf & "--" & sDelim & vbCrLf & _
                      "Content-Type: image/" & .Ext & "; name=""" & .Nom & "." & .Ext & """" & vbCrLf & _
                      "Content-ID: <" & .cid & ">" & vbCrLf & _
                      "Content-Transfer-Encoding: base64" & vbCrLf & _
                      "Content-Disposition: inline; filename=""" & .Nom & "." & .Ext & """" & vbCrLf & _
                      vbCrLf & _
                      Enc_Base64(PJFichier(dtuMR.src))
    End With
End Function

' Renseigne un tableau avec toutes les balises IMG.
' Retourne le ubound, ou -1 si aucun �l�ment.
Private Function IMG_Balises(sHTML As String, dtuMR() As tuMpRel) As Long
    Dim i As Long, j As Long, i0 As Long, j0 As Long, n As Long
    Dim sBal As String, sSrc As String, s As String

    ReDim dtuMR(100)                                    ' 100 images.

    n = -1
    ' Chercher et stocker toutes les balises IMG dans un tableau.
    i = InStr(sHTML, "<IMG ")
    Do While i > 0
        j = InStr(i, sHTML, ">")                        ' Fin de la balise.
        If j = 0 Then j = Len(sHTML)

        sBal = Mid$(sHTML, i, j - i + 1)                ' Extraire la balise <IMG ....>

        i0 = InStr(sBal, "src=""")                      ' Chercher le champ 'src='
        If i0 = 0 Then i0 = 1
        j0 = InStr(i0 + 5, sBal, """")
        If j0 = 0 Then j0 = Len(sBal) - 1
        sSrc = Mid$(sBal, i0 + 5, j0 - i0 - 5)

        n = n + 1                                       ' Ecrire ici...
        With dtuMR(n)
            .bal = sBal                                 ' Balise <IMG ....> compl�te.
            .src = sSrc                                 ' Contenu du champ src=.
            Call AnaSpecFich(.src, s, .Nom, .Ext)       ' S�pare nom et extension.
        End With

        i = InStr(j, sHTML, "<IMG ")                    ' Balise suivante.
    Loop

    If n > -1 Then ReDim Preserve dtuMR(n)              ' Ne garder que la partie utile.
    IMG_Balises = n
End Function

' Remplace les s�quences %xx par le caract�re correspondant et / par \
Private Function CorrigeURL(sURL As String) As String
    Dim i As Long, j As Long, s As String

    CorrigeURL = Space$(Len(sURL))

    i = 1: j = 1
    Do While i <= Len(sURL)
        s = Mid$(sURL, i, 1)

        Select Case s
            Case "/":   s = "\"
            Case "%"
                If i < Len(sURL) - 1 Then
                    ' Convertir la valeur hexa qui suit le %.
                    s = Chr$(Val("&H" & Mid$(sURL, i + 1, 2)))
                    i = i + 2
                End If
        End Select

        Mid$(CorrigeURL, j, 1) = s
        i = i + 1
        j = j + 1
    Loop
    CorrigeURL = Left$(CorrigeURL, j - 1)
End Function

' Cr�e un d�limiteur de parties de message.
Private Function Delimiteur(Optional sRacine As String = "----Separateur=_") As String
    Randomize Timer
    Delimiteur = sRacine & CDec(Date * 24 * 3600) & "." & Int(Rnd * 1000000)
End Function