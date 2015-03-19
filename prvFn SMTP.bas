Option Compare Database
Option Explicit
Option Private Module

' Copyright 2009-2014 Denis SCHEIDT
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





' Constantes pour l'attente du serveur distant
Public Const REP_AUCUNE    As Integer = 0
Public Const REP_DELAI     As Integer = -2
Public Const REP_INFINI    As Integer = -3

' Variable globale d'état du serveur. Le type est Public, mais la variable est locale à la bibliothèque.
Public dtuEtatSRV_RAZ           As tuEtatSRV


' TU pour l'analyse du défi du serveur
Private Type tuD_MD5_S
    realm               As String           ' Opt.  Mult.   Déf: serveur smtp fourni à libMAIL
    nonce               As String           ' Req.  Uniq.
    qop                 As String           ' Opt.          Déf: "auth"
    maxbuf              As Long             ' Opt.  Uniq.   Déf: 65536
    Charset             As String           ' Opt.  Uniq.   Déf: ISO 8859-1
    algorithm           As String           ' Req.  Uniq.
End Type




' Gère la phase de connexion au serveur SMTP distant.
' Retour :
'        0 : Connexion TCP/IP OK. lSock contient le n° du socket.
'        n : Erreur de connexion.
'            Dans ce cas, lSock
'            -1 : impossible d'initialiser Winsock
'             0 : impossible de créer un socket.
'             n : n° de socket s'il a pu être créé.
Function CnxServ(sNomSrv As String, lPort As Long, lSock As Long) As Long
    Dim WSAdata As WSA_DATA, SrvSmtp As SOCK_ADDR
    Dim lRet As Long, sIP As String

    ' Initialisation de WinSock (version 1.1)
    lSock = -1                                                      ' Impossible d'initialiser Winsock
    CnxServ = WSAStartup(&H101, WSAdata)                            ' Remonter le code d'erreur.
    Call Journal(Traduit("¤cnx_wsockinit", "Initialisation de Winsock...  LastError=[%s]", WSAGetLastError()))
    If CnxServ <> 0 Then Exit Function                              ' Erreur...


    ' Création d'un Socket
    lSock = socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)
    Call Journal(Traduit("¤cnx_wsockcreate", "Création du socket... (%s). LastError=[%s]", lSock, WSAGetLastError()))
    If lSock = INVALID_SOCKET Then
        CnxServ = lSock                                             ' Remonter le code d'erreur.
        lSock = 0                                                   ' Impossible de créer le socket.
        Exit Function                                               ' Erreur...
    End If


    ' Tentative de connexion au serveur
    ' Tenter de convertir directement, c'est peut-être une adresse IP.
    lRet = inet_addr(sNomSrv)
    ' Résolution de l'adresse IP, si nécessaire.
    If lRet = INADDR_ANY Or lRet = INADDR_NONE Then
        sIP = AdresseSrv(sNomSrv)                                   ' Conversion nom --> adresse IP.
        If Len(sIP) = 0 Then                                        ' Résolution impossible...
            CnxServ = SOCKET_ERROR                                  ' Erreur de connexion...
            Call Journal(Traduit("¤Résolution de nom impossible pour [%s].", sNomSrv))
            Exit Function
        End If
        lRet = inet_addr(sIP)
    End If

    With SrvSmtp
        .sin_family = PF_INET
        .sin_addr.S_addr = lRet
        .sin_port = htons(lPort)
        .sin_zero(0) = 0
    End With
    lRet = connect(lSock, SrvSmtp, Len(SrvSmtp))
    Call Journal(Traduit("¤cnx_connect", "Connexion au serveur %s [%s], sur le port %s... LastError=[%s]", sNomSrv, sIP, lPort, WSAGetLastError()))
    If lRet = SOCKET_ERROR Then
        Call Journal(Traduit("¤cnx_error", "*** Connexion impossible..."))
        CnxServ = lRet                                              ' Remonter le code d'erreur.
        Exit Function
    End If


    ' Mise en place de la détection de données
    ' Va mettre le socket en mode non bloquant.
    lRet = ioctlsocket(lSock, FIOBION, &H1&)
    Call Journal(Traduit("¤cnx_socket", "WinSock : Mode non-bloquant. LastError=[%s]", WSAGetLastError()))
    If lRet = SOCKET_ERROR Then
        CnxServ = lRet
        Exit Function
    End If

    CnxServ = 0                                                      ' Connexion établie.
End Function


' Gère la fin de connexion.
' Fermeture du socket et libération des ressources.
Sub CnxFin(lSock As Long)
    Dim lRet As Long

    If lSock > 0 Then
        ' Fermeture du Socket
        lRet = closesocket(lSock)
        Call Journal(Traduit("¤cnx_close", "Fermeture du socket... (%s). LastError=[%s]", lSock, WSAGetLastError()))
    End If

    If lSock > -1 Then
        ' Nettoyage final
        lRet = WSACleanup()
        Call Journal(Traduit("¤cnx_cleanup", "Libération des ressources. LastError=[%s]", WSAGetLastError()))
    End If
End Sub

' Envoie une commande au serveur SMTP distant et attend la réponse
'
' Paramètres :
'   sCmd             :  chaîne de commande.
'                       Si <Null>, n'envoie pas de commande, se met en attente de réponse.
'   bRepAttendue     :  REP_INFINI  attend indéfiniment une réponse du serveur
'                       REP_AUCUNE  n'attend pas de réponse
'                       REP_DELAI   attend iDelaiRep secondes
'   bLogAUTH         :  booléen - empêche la journalisation des informations d'authentification.
'
' Valeurs de retour :
'   Retourne le code de réponse du serveur (entre 1 et 5)
'       La chaine complète renvoyée par le serveur peut être récupérée dans sRepSrv
'   Si aucune réponse n'est attendue,   retourne REP_AUCUNE   ( 0) -- utilisé pour la partie DATA.
'   En cas d'erreur de socket,          retourne SOCKET_ERROR (-1)
'   En cas de dépassement de délai,     retourne REP_DELAI    (-2)
Function EnvoiCMD(lSock As Long, ByVal sCmd As Variant, _
                          Optional bRepAttendue As Integer = REP_DELAI, Optional bLogAUTH As Boolean = True, _
                          Optional sRepSrv As String = "") As Integer
    Dim lNbCar As Long, bTampon() As Byte, i As Long, lRet As Long
    Dim sDelai As Single, s As String, sCmd0 As String
    Dim dtuTimeVal As timeval, dtuFD_Read As fd_set, dtuFD_Write As fd_set, dtuFD_Except As fd_set, dtuFD_RAZ As fd_set


    If Not IsNull(sCmd) Then ' ***** Partie ENVOI *****************************************************
        sCmd0 = sCmd & vbCrLf                                       ' Ajouter le CRLF requis par SMTP

        dtuTimeVal.tv_sec = 0                                       ' Délai en secondes
        dtuTimeVal.tv_usec = 0                                      ' Délai pour select(), en µs.
                                                                    ' Avec {0,0} WSSelect n'attend pas.

        Do While Len(sCmd0) > 0                                     ' Jusqu'à ce que tout soit envoyé
            i = 0                                                   ' Code erreur de WSAGetLastError

            ' Vérifier que le socket est prêt ---------------------------------------------------------
            ' Attendre qu'un socket soit prêt, à concurrence de iDelaiRep.
            ' Si WSSelect retourne 0, soit aucun socket n'est prêt,
            ' soit le délai dtuTimeVal a expiré.
            sDelai = Timer
            Do
                dtuFD_Read = dtuFD_RAZ
                dtuFD_Write = dtuFD_RAZ
                dtuFD_Except = dtuFD_RAZ
                With dtuFD_Write
                    .fd_count = 1                                   ' Nombre de sockets à contrôler
                    .fd_array(0) = lSock
                End With

                lRet = WSSelect(0, dtuFD_Read, dtuFD_Write, dtuFD_Except, dtuTimeVal)
                If dtuFD_Write.fd_count = 0 Then
                    Call myDoEvents                                 ' Il faut attendre un peu...
                End If
            Loop While dtuFD_Write.fd_count = 0 And (Abs(Timer - sDelai) < dtuEtatSyst.Serveur.DelaiReponse)

            If lRet <= 0 Then                                       ' Sortie sur SOCKET_ERROR ou TimeOut
                i = WSAGetLastError()
                Exit Do
            End If

            lNbCar = Len(sCmd0)                                     ' Nombre de caractères à envoyer
            bTampon = StrConv(sCmd0, vbFromUnicode)                 ' Chaine vers tableau d'octets
            ReDim Preserve bTampon(lNbCar)                          ' Agrandir d'un octet dont la valeur est 0

            lRet = send(lSock, bTampon(0), lNbCar, 0)               ' Envoi au serveur SMTP

            If lRet = SOCKET_ERROR Then                             ' Erreur de socket
                i = WSAGetLastError()
                If i <> WSAEWOULDBLOCK Then Exit Do                 ' Sortie sur autre erreur
                Call myDoEvents

            Else                                                    ' Décrémenter le compteur
                sCmd0 = Mid$(sCmd0, lRet + 1)                       ' Soumettre les caractères restants

            End If
        Loop

        If bLogAUTH Then                                            ' Journaliser la commande
            Call Journal("--> " & sCmd & ", LastError=[" & i & "]")
        Else
            Call Journal(Traduit("¤cmd_authdata", "--> <*Données d'authentification*>, LastError=[%s]", i))
        End If

        If lRet = SOCKET_ERROR Then                                 ' Erreur de socket. On sort.
            EnvoiCMD = lRet
            Call Journal(Traduit("¤cmd_senderror", "*** Erreur de socket sur send()."))
            Exit Function
        End If

        If lRet = 0 Then                                            ' Dépassement de délai
            EnvoiCMD = REP_DELAI
            Call Journal(Traduit("¤cmd_sendtimeout", "*** Dépassement de délai sur WSSelect()."))
            Exit Function
        End If

    End If ' ==========================================================================================

    If bRepAttendue = REP_AUCUNE Then                               ' Aucune réponse du serveur n'est attendue
        EnvoiCMD = REP_AUCUNE
        Exit Function
    End If

    sDelai = Timer ' ***** Partie RECEPTION (attente réponse) *****************************************
    lRet = -999999999
    lNbCar = 5000
    ReDim bTampon(lNbCar)

    Do
        Call myDoEvents

        dtuFD_Read = dtuFD_RAZ
        dtuFD_Write = dtuFD_RAZ
        dtuFD_Except = dtuFD_RAZ
        With dtuFD_Read
            .fd_count = 1                                           ' Nombre de sockets à contrôler
            .fd_array(0) = lSock
        End With

        lRet = WSSelect(0, dtuFD_Read, dtuFD_Write, dtuFD_Except, dtuTimeVal)
        If dtuFD_Read.fd_count = 0 Then
            Call myDoEvents                                         ' Il faut attendre un peu...
        End If
    Loop While dtuFD_Read.fd_count = 0 And (Abs(Timer - sDelai) < dtuEtatSyst.Serveur.DelaiReponse)

    If lRet <= 0 Then                                               ' Sortie sur SOCKET_ERROR ou TimeOut
        i = WSAGetLastError()

    Else
        lRet = recv(lSock, bTampon(0), lNbCar, 0)
        i = WSAGetLastError()                                       ' Erreur de socket

    End If

    Select Case lRet
        Case SOCKET_ERROR                                           ' Erreur de socket. On sort.
            Call Journal(Traduit("¤cmd_rcverror", "*** Erreur de socket en réception, LastError=[%s]", i))
            EnvoiCMD = SOCKET_ERROR

        Case 0                                                      ' Fermeture de connexion par le serveur distant
            Call Journal(Traduit("¤cmd_cnxclose", "*** Fermeture de connexion par le serveur distant, LastError=[%s]", i))
            EnvoiCMD = SOCKET_ERROR

        Case Is > 0                                                 ' Données reçues normalement
            s = StrConv(bTampon(), vbUnicode)
            s = Left$(s, lRet - 2)                                  ' Retirer le CrLf final de la réponse.

            Call Journal("<-- " & s & ", LastError=[" & i & "]")
            sRepSrv = s                                             ' Renvoyer la réponse complète du serveur
            EnvoiCMD = Val(Left$(s, 1))                             ' Ne garder que le premier chiffre de la réponse

        Case Else                                                   ' Sortie sur dépassement de délai
            Call Journal(Traduit("¤cmd_rcvtimeout", "*** Dépassement de délai de réception."))
            EnvoiCMD = REP_DELAI
    End Select ' ======================================================================================
End Function



' Ecriture du journal de connexion.
Sub Journal(sTexte As String)
    Dim i As Integer, lNbC As Long, s As String

    If Not dtuEtatSyst.Journal.LogComm Then Exit Sub

    ' Journalise dans la variable.
    With dtuEtatSyst.Journal
        ' Crée une chaine au format dd/mm/yyyy hh:nn:ss.xxxx
        s = HoroDatage() & " : " & sTexte & vbCrLf
        lNbC = Len(s)

        If .IxDebut = 0 Then Call SMTPJnlRAZ                    ' Si Journal est appelée avant SMTPJnlRAZ.
        Mid$(.Journal, .IxDebut, lNbC) = s
        .IxDebut = .IxDebut + lNbC                              ' Ajuster le pointeur à la position suivante.

        If .IxDebut > lLnMaxJnl Then                            ' Il faut boucler.
            lNbC = .IxDebut - lLnMaxJnl - 1                     ' Nombre de caractère à reprendre.
            Mid$(.Journal, 1, lNbC) = Right$(s, lNbC)
            .IxDebut = lNbC + 1
        End If
    End With

    If Len(dtuEtatSyst.Journal.FichierJnl) = 0 Then Exit Sub    ' Pas de journal fichier.

    On Error Resume Next                                        ' Au cas où...
    ' Journal Fichier
    i = FreeFile()
    Open dtuEtatSyst.Journal.FichierJnl For Append Access Write Shared As #i
    Print #i, Left$(s, Len(s) - 2)
    Close #i
    On Error GoTo 0
End Sub

Sub myDoEvents()
    Dim tuMsg As MSG

    Do While PeekMessage(tuMsg, 0&, 0&, 0&, PM_REMOVE)
        TranslateMessage tuMsg
        DispatchMessage tuMsg
    Loop
End Sub

' Timer haute précision
' Utilise le type Decimal pour pouvoir contenir le compteur...
' Donne le temps écoulé depuis le démarrage de Windows
Function HPC() As Variant
    Dim cT1 As Currency
    Static cTF As Currency

    If cTF = 0 Then Call QueryPerformanceFrequency(cTF)

    Call QueryPerformanceCounter(cT1)
    HPC = CDec(cT1 / cTF)
End Function

' Extraction du nom et du port du serveur (ex.: smtp.nom_fai.fr[:25])
' La variable sNomSrv est modifiée par la procédure !!!
' Le port est 25, par défaut.
Sub ServPort(sNomSrv As String, lPort As Long)
    Dim i As Integer, s As String

    i = InStr(sNomSrv, ":")
    lPort = 25                                                  ' Par défaut
    If i <> 0 Then
        ' Extraire d'abord le port
        s = Mid$(sNomSrv, i + 1)
        If IsNumeric(s) Then lPort = Val(s) Mod 65536
        ' puis le nom du serveur.
        sNomSrv = Trim$(Left$(sNomSrv, i - 1))
    End If
End Sub

' Nom de la méthode d'authentification
Function NomMethodeAuth(lMethode As Integer) As Variant
    Select Case lMethode
        Case lmlESMTPAuthAucune:    NomMethodeAuth = Traduit("auth_none", "Aucune")
        Case lmlESMTPAuthLogin:     NomMethodeAuth = "LOGIN"
        Case lmlESMTPAuthPlain:     NomMethodeAuth = "PLAIN"
        Case lmlESMTPAuthCRAMMD5:   NomMethodeAuth = "CRAM-MD5"
        Case lmlESMTPAuthDIGESTMD5: NomMethodeAuth = "DIGEST-MD5"
        Case lmlESMTPAuthSTARTTLS:  NomMethodeAuth = "STARTTLS"
        Case Else:                  NomMethodeAuth = Null
    End Select
End Function

' Calcule la réponse pour une authentification de type CRAM-MD5 (RFC-2195)
Function CRAM_MD5(ByVal spDefiSrv As String) As String
    Dim s1 As String

    ' Décoder le défi décodé du serveur
    If spDefiSrv Like "### *" Then Mid$(spDefiSrv, 1, 3) = "   "
    spDefiSrv = Dec_Base64(spDefiSrv)

    ' Obtenir le HMAC
    s1 = HMAC_MD5(dtuEtatSyst.Serveur.OptionsESMTP.AUTH.MotDePasse, spDefiSrv)

    ' Concaténer UserName et HMAC
    s1 = dtuEtatSyst.Serveur.OptionsESMTP.AUTH.Utilisateur & " " & s1

    ' Convertir en Base64 et sortir le résultat
    CRAM_MD5 = Enc_Base64(s1)
End Function

' Analyse le défi envoyé par le serveur et calcule la réponse adéquate.
' Si le défi est invalide, retourne ""
Function DIGEST_MD5(ByVal spDefiSrv As String) As String
    Dim v As Variant, i As Integer, j As Integer, sChamp As String, sVal As String, dtuDefi As tuD_MD5_S
    Dim bAbandon As Boolean
    Dim A1 As String, A2 As String, sCNONCE As String, sReponse As String

    ' Etape 1. Décoder le défi
    ' ========================
    spDefiSrv = Dec_Base64(spDefiSrv)
    spDefiSrv = Remplacer(spDefiSrv, """", "")                      ' Retirer les quotes.

    ' Séparer les différents éléments du défi
    v = Scinder(spDefiSrv, ",")
    For i = 0 To UBound(v)
        j = InStr(v(i), "=")
        sChamp = Left$(v(i), j - 1)                                 ' Nom du champ
        sVal = Mid$(v(i), j + 1)                                    ' Valeur du champ

        With dtuDefi
            Select Case sChamp
                Case "realm"
                    .realm = sVal

                Case "nonce"
                    ' Doublons interdits
                    If Len(.nonce) = 0 Then .nonce = sVal Else bAbandon = True

                Case "qop"
                    .qop = sVal
                    If .qop <> "auth" Then .qop = ""

                Case "maxbuf"
                    ' Doublons interdits
                    If .maxbuf = 0 Then .maxbuf = Val(sVal) Else bAbandon = True

                Case "charset"
                    ' Doublons interdits
                    If Len(.Charset) = 0 Then .Charset = sVal Else bAbandon = True

                Case "algorithm"
                    ' Doublons interdits
                    If Len(.algorithm) = 0 Then .algorithm = sVal Else bAbandon = True

            End Select

        End With
    Next i

    ' Contrôler la présence des champs obligatoires.
    If Len(dtuDefi.nonce) = 0 Then bAbandon = True
    If Len(dtuDefi.algorithm) = 0 Then bAbandon = True

    ' Doublon ou champ requis manquant. On abandonne la connexion.
    If bAbandon Then Exit Function

    ' Compléter les valeurs par défaut pour les champs facultatifs qui n'ont pas été renseignés par le serveur
    With dtuDefi
        If Len(.realm) = 0 Then .realm = dtuEtatSyst.Serveur.NomSrv
        If Len(.qop) = 0 Then .qop = "auth"
        If .maxbuf = 0 Then .maxbuf = 65536
        If Len(.Charset) = 0 Then .Charset = "ISO 8859-1"
    End With

    ' Etape 2. Préparer la réponse
    ' ============================

    sCNONCE = Enc_Base64(Alea(32))                          ' Créer une chaine aléatoire de 32 caractères.

    A1 = MD5(dtuEtatSyst.Serveur.OptionsESMTP.AUTH.Utilisateur & ":" & _
             dtuDefi.realm & ":" & _
             dtuEtatSyst.Serveur.OptionsESMTP.AUTH.MotDePasse) & ":" & _
         dtuDefi.nonce & ":" & sCNONCE

    A2 = "AUTHENTICATE:smtp/" & dtuDefi.realm

    sReponse = myHEX(MD5(myHEX(MD5(A1)) & ":" & dtuDefi.nonce & ":" & "00000001" & ":" & sCNONCE & ":" & dtuDefi.qop & ":" & myHEX(MD5(A2))))

    DIGEST_MD5 = "username=""" & dtuEtatSyst.Serveur.OptionsESMTP.AUTH.Utilisateur & """," & _
                 "realm=""" & dtuDefi.realm & """," & _
                 "nonce=""" & dtuDefi.nonce & """," & _
                 "cnonce=""" & sCNONCE & """," & _
                 "nc=00000001," & _
                 "qop=auth," & _
                 "charset=utf-8," & _
                 "digest-uri=""smtp/" & dtuDefi.realm & """," & _
                 "response=" & sReponse
End Function

' Application de la RFC-2104
Function HMAC_MD5(sSecret As String, sTexte As String) As String
    Dim s1 As String, s2 As String, s3 As String

    Const b     As Long = 64
    Const ipad  As Byte = &H36
    Const opad  As Byte = &H5C

    s1 = String$(b, &H0)

    ' 1. Compléter le mot de passe avec des 0 à concurrence de B octets
    Mid$(s1, 1, Len(sSecret)) = sSecret

    ' 2. Faire un OU exclusif entre (1) et ipad
    s2 = strXOR(s1, ipad)

    ' 3. Concaténer (2) et le texte
    s2 = s2 & sTexte

    ' 4. Calculer le MD5 de (3)
    s2 = MD5(s2)

    ' 5. Faire un OU exclusif entre (1) et opad
    s3 = strXOR(s1, opad)

    ' 6. Concaténer (5) et (4)
    s3 = s3 & s2

    ' 7. Calculer le MD5 de (6)
    HMAC_MD5 = myHEX(MD5(s3))
End Function





' Cherche l'adresse IP en fonction du nom
Private Function AdresseSrv(ByVal sNom As String) As String
#If Vba7 Then
    Dim ptrHosent As LongPtr, ptrAdresse As LongPtr
#Else
    Dim ptrHosent As Long, ptrAdresse As Long
#End If
    Dim sAdresse As String, ptrAdrIP As Long

    sAdresse = Space(4)
    ptrHosent = gethostbyname(sNom & vbNullChar)                    ' Retourne un pointeur vers la structure

    ' Résolution de nom impossible...
    If ptrHosent = 0 Then Exit Function

#If Win64 Then
    ptrAdresse = ptrHosent + 24
#Else
    ptrAdresse = ptrHosent + 12                                     ' L'adresse IP est 12 octets après le début
#End If

    CopyMemory ptrAdresse, ByVal ptrAdresse, 4
    CopyMemory ptrAdrIP, ByVal ptrAdresse, 4
    CopyMemory ByVal sAdresse, ByVal ptrAdrIP, 4

    ' Convertir en chaine, séparée par des points
    AdresseSrv = CStr(Asc(sAdresse)) & "." & CStr(Asc(Mid$(sAdresse, 2, 1))) & "." & _
                 CStr(Asc(Mid$(sAdresse, 3, 1))) & "." & CStr(Asc(Mid$(sAdresse, 4, 1)))
End Function

' Effectue un OU exclusif entre la chaine et la valeur
Private Function strXOR(sCh1 As String, oCh2 As Byte) As String
    Dim b1() As Byte, i As Long

    b1 = StrConv(sCh1, vbFromUnicode)
    For i = 0 To UBound(b1)
        b1(i) = (b1(i) Xor oCh2)
    Next i
    strXOR = StrConv(b1, vbUnicode)
End Function