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





' Constantes pour l'attente du serveur distant
Public Const REP_AUCUNE    As Integer = 0
Public Const REP_DELAI     As Integer = -2
Public Const REP_INFINI    As Integer = -3

' Variable globale d'�tat du serveur. Le type est Public, mais la variable est locale � la biblioth�que.
Public dtuEtatSRV_RAZ           As tuEtatSRV


' TU pour l'analyse du d�fi du serveur
Private Type tuD_MD5_S
    realm               As String           ' Opt.  Mult.   D�f: serveur smtp fourni � libMAIL
    nonce               As String           ' Req.  Uniq.
    qop                 As String           ' Opt.          D�f: "auth"
    maxbuf              As Long             ' Opt.  Uniq.   D�f: 65536
    Charset             As String           ' Opt.  Uniq.   D�f: ISO 8859-1
    algorithm           As String           ' Req.  Uniq.
End Type




' G�re la phase de connexion au serveur SMTP distant.
' Retour :
'        0 : Connexion TCP/IP OK. lSock contient le n� du socket.
'        n : Erreur de connexion.
'            Dans ce cas, lSock
'            -1 : impossible d'initialiser Winsock
'             0 : impossible de cr�er un socket.
'             n : n� de socket s'il a pu �tre cr��.
Function CnxServ(sNomSrv As String, lPort As Long, lSock As Long) As Long
    Dim WSAdata As WSA_DATA, SrvSmtp As SOCK_ADDR
    Dim lRet As Long, sIP As String

    ' Initialisation de WinSock (version 1.1)
    lSock = -1                                                      ' Impossible d'initialiser Winsock
    CnxServ = WSAStartup(&H101, WSAdata)                            ' Remonter le code d'erreur.
    Call Journal("Initialisation de Winsock... LastError=[" & WSAGetLastError() & "]")
    If CnxServ <> 0 Then Exit Function                              ' Erreur...


    ' Cr�ation d'un Socket
    lSock = socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)
    Call Journal("Cr�ation du socket... (" & lSock & "), LastError=[" & WSAGetLastError() & "]")
    If lSock = INVALID_SOCKET Then
        CnxServ = lSock                                             ' Remonter le code d'erreur.
        lSock = 0                                                   ' Impossible de cr�er le socket.
        Exit Function                                               ' Erreur...
    End If


    ' Tentative de connexion au serveur
    ' Tenter de convertir directement, c'est peut-�tre une adresse IP.
    lRet = inet_addr(sNomSrv)
    ' R�solution de l'adresse IP, si n�cessaire.
    If lRet = INADDR_ANY Or lRet = INADDR_NONE Then
        sIP = AdresseSrv(sNomSrv)                                   ' Conversion nom --> adresse IP.
        If Len(sIP) = 0 Then                                        ' R�solution impossible...
            CnxServ = SOCKET_ERROR                                  ' Erreur de connexion...
            Call Journal("R�solution de nom impossible pour [" & sNomSrv & "].")
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
    Call Journal("Connexion au serveur " & sNomSrv & " [" & sIP & "], sur le port " & lPort & "... LastError=[" & WSAGetLastError() & "]")
    If lRet = SOCKET_ERROR Then
        Call Journal("*** Connexion impossible...")
        CnxServ = lRet                                              ' Remonter le code d'erreur.
        Exit Function
    End If


    ' Mise en place de la d�tection de donn�es
    ' Va mettre le socket en mode non bloquant.
    lRet = ioctlsocket(lSock, FIOBION, &H1&)
    Call Journal("WinSock : Mode non-bloquant. LastError=[" & WSAGetLastError() & "]")
    If lRet = SOCKET_ERROR Then
        CnxServ = lRet
        Exit Function
    End If

    CnxServ = 0                                                      ' Connexion �tablie.
End Function


' G�re la fin de connexion.
' Fermeture du socket et lib�ration des ressources.
Sub CnxFin(lSock As Long)
    Dim lRet As Long

    If lSock > 0 Then
        ' Fermeture du Socket
        lRet = closesocket(lSock)
        Call Journal("Fermeture du socket... (" & lSock & "), LastError=[" & WSAGetLastError() & "]")
    End If

    If lSock > -1 Then
        ' Nettoyage final
        lRet = WSACleanup()
        Call Journal("Lib�ration des ressources. LastError=[" & WSAGetLastError() & "]")
    End If
End Sub

' Envoie une commande au serveur SMTP distant et attend la r�ponse
'
' Param�tres :
'   sCmd             :  cha�ne de commande.
'                       Si <Null>, n'envoie pas de commande, se met en attente de r�ponse.
'   bRepAttendue     :  REP_INFINI  attend ind�finiment une r�ponse du serveur
'                       REP_AUCUNE  n'attend pas de r�ponse
'                       REP_DELAI   attend iDelaiRep secondes
'   bLogAUTH         :  bool�en - emp�che la journalisation des informations d'authentification.
'
' Valeurs de retour :
'   Retourne le code de r�ponse du serveur (entre 1 et 5)
'       La chaine compl�te renvoy�e par le serveur peut �tre r�cup�r�e dans sRepSrv
'   Si aucune r�ponse n'est attendue,   retourne REP_AUCUNE   ( 0) -- utilis� pour la partie DATA.
'   En cas d'erreur de socket,          retourne SOCKET_ERROR (-1)
'   En cas de d�passement de d�lai,     retourne REP_DELAI    (-2)
Function EnvoiCMD(lSock As Long, ByVal sCmd As Variant, _
                          Optional bRepAttendue As Integer = REP_DELAI, Optional bLogAUTH As Boolean = True, _
                          Optional sRepSrv As String = "") As Integer
    Dim lNbCar As Long, bTampon() As Byte, i As Long, lRet As Long
    Dim sDelai As Single, s As String, sCmd0 As String
    Dim dtuTimeVal As timeval, dtuFD_Read As fd_set, dtuFD_Write As fd_set, dtuFD_Except As fd_set, dtuFD_RAZ As fd_set


    If Not IsNull(sCmd) Then ' ***** Partie ENVOI *****************************************************
        sCmd0 = sCmd & vbCrLf                                       ' Ajouter le CRLF requis par SMTP

        dtuTimeVal.tv_sec = 0                                       ' D�lai en secondes
        dtuTimeVal.tv_usec = 0                                      ' D�lai pour select(), en �s.
                                                                    ' Avec {0,0} WSSelect n'attend pas.

        Do While Len(sCmd0) > 0                                     ' Jusqu'� ce que tout soit envoy�
            i = 0                                                   ' Code erreur de WSAGetLastError

            ' V�rifier que le socket est pr�t ---------------------------------------------------------
            ' Attendre qu'un socket soit pr�t, � concurrence de iDelaiRep.
            ' Si WSSelect retourne 0, soit aucun socket n'est pr�t,
            ' soit le d�lai dtuTimeVal a expir�.
            sDelai = Timer
            Do
                dtuFD_Read = dtuFD_RAZ
                dtuFD_Write = dtuFD_RAZ
                dtuFD_Except = dtuFD_RAZ
                With dtuFD_Write
                    .fd_count = 1                                   ' Nombre de sockets � contr�ler
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

            lNbCar = Len(sCmd0)                                     ' Nombre de caract�res � envoyer
            bTampon = StrConv(sCmd0, vbFromUnicode)                 ' Chaine vers tableau d'octets
            ReDim Preserve bTampon(lNbCar)                          ' Agrandir d'un octet dont la valeur est 0

            lRet = send(lSock, bTampon(0), lNbCar, 0)               ' Envoi au serveur SMTP

            If lRet = SOCKET_ERROR Then                             ' Erreur de socket
                i = WSAGetLastError()
                If i <> WSAEWOULDBLOCK Then Exit Do                 ' Sortie sur autre erreur
                Call myDoEvents

            Else                                                    ' D�cr�menter le compteur
                sCmd0 = Mid$(sCmd0, lRet + 1)                       ' Soumettre les caract�res restants

            End If
        Loop

        If bLogAUTH Then                                            ' Journaliser la commande
            Call Journal("--> " & sCmd & ", LastError=[" & i & "]")
        Else
            Call Journal("--> <*Donn�es d'authentification*>" & ", LastError=[" & i & "]")
        End If

        If lRet = SOCKET_ERROR Then                                 ' Erreur de socket. On sort.
            EnvoiCMD = lRet
            Call Journal("*** Erreur de socket sur send().")
            Exit Function
        End If

        If lRet = 0 Then                                            ' D�passement de d�lai
            EnvoiCMD = REP_DELAI
            Call Journal("*** D�passement de d�lai sur WSSelect().")
            Exit Function
        End If

    End If ' ==========================================================================================

    If bRepAttendue = REP_AUCUNE Then                               ' Aucune r�ponse du serveur n'est attendue
        EnvoiCMD = REP_AUCUNE
        Exit Function
    End If

    sDelai = Timer ' ***** Partie RECEPTION (attente r�ponse) *****************************************
    lRet = -999999999
    lNbCar = 5000
    ReDim bTampon(lNbCar)

    Do
        Call myDoEvents

        dtuFD_Read = dtuFD_RAZ
        dtuFD_Write = dtuFD_RAZ
        dtuFD_Except = dtuFD_RAZ
        With dtuFD_Read
            .fd_count = 1                                           ' Nombre de sockets � contr�ler
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
            Call Journal("*** Erreur de socket en r�ception, LastError=[" & i & "]")
            EnvoiCMD = SOCKET_ERROR

        Case 0                                                      ' Fermeture de connexion par le serveur distant
            Call Journal("*** Fermeture de connexion par le serveur distant, LastError=[" & i & "]")
            EnvoiCMD = SOCKET_ERROR

        Case Is > 0                                                 ' Donn�es re�ues normalement
            s = StrConv(bTampon(), vbUnicode)
            s = Left$(s, lRet - 2)                                  ' Retirer le CrLf final de la r�ponse.

            Call Journal("<-- " & s & ", LastError=[" & i & "]")
            sRepSrv = s                                             ' Renvoyer la r�ponse compl�te du serveur
            EnvoiCMD = Val(Left$(s, 1))                             ' Ne garder que le premier chiffre de la r�ponse

        Case Else                                                   ' Sortie sur d�passement de d�lai
            Call Journal("*** D�passement de d�lai de r�ception.")
            EnvoiCMD = REP_DELAI
    End Select ' ======================================================================================
End Function



' Ecriture du journal de connexion.
Sub Journal(sTexte As String)
    Dim i As Integer, lNbC As Long, s As String

    If Not dtuEtatSyst.Journal.LogComm Then Exit Sub

    ' Journalise dans la variable.
    With dtuEtatSyst.Journal
        ' Cr�e une chaine au format dd/mm/yyyy hh:nn:ss.xxxx
        s = HoroDatage() & " : " & sTexte & vbCrLf
        lNbC = Len(s)

        If .IxDebut = 0 Then Call SMTPJnlRAZ                    ' Si Journal est appel�e avant SMTPJnlRAZ.
        Mid$(.Journal, .IxDebut, lNbC) = s
        .IxDebut = .IxDebut + lNbC                              ' Ajuster le pointeur � la position suivante.

        If .IxDebut > lLnMaxJnl Then                            ' Il faut boucler.
            lNbC = .IxDebut - lLnMaxJnl - 1                     ' Nombre de caract�re � reprendre.
            Mid$(.Journal, 1, lNbC) = Right$(s, lNbC)
            .IxDebut = lNbC + 1
        End If
    End With

    If Len(dtuEtatSyst.Journal.FichierJnl) = 0 Then Exit Sub    ' Pas de journal fichier.

    On Error Resume Next                                        ' Au cas o�...
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

' Timer haute pr�cision
' Utilise le type Decimal pour pouvoir contenir le compteur...
' Donne le temps �coul� depuis le d�marrage de Windows
Function HPC() As Variant
    Dim cT1 As Currency
    Static cTF As Currency

    If cTF = 0 Then Call QueryPerformanceFrequency(cTF)

    Call QueryPerformanceCounter(cT1)
    HPC = CDec(cT1 / cTF)
End Function

' Extraction du nom et du port du serveur (ex.: smtp.nom_fai.fr[:25])
' La variable sNomSrv est modifi�e par la proc�dure !!!
' Le port est 25, par d�faut.
Sub ServPort(sNomSrv As String, lPort As Long)
    Dim i As Integer, s As String

    i = InStr(sNomSrv, ":")
    lPort = 25                                                  ' Par d�faut
    If i <> 0 Then
        ' Extraire d'abord le port
        s = Mid$(sNomSrv, i + 1)
        If IsNumeric(s) Then lPort = Val(s) Mod 65536
        ' puis le nom du serveur.
        sNomSrv = Trim$(Left$(sNomSrv, i - 1))
    End If
End Sub

' Nom de la m�thode d'authentification
Function NomMethodeAuth(lMethode As Integer) As Variant
    Select Case lMethode
        Case lmlESMTPAuthAucune:    NomMethodeAuth = "Aucune"
        Case lmlESMTPAuthLogin:     NomMethodeAuth = "LOGIN"
        Case lmlESMTPAuthPlain:     NomMethodeAuth = "PLAIN"
        Case lmlESMTPAuthCRAMMD5:   NomMethodeAuth = "CRAM-MD5"
        Case lmlESMTPAuthDIGESTMD5: NomMethodeAuth = "DIGEST-MD5"
        Case lmlESMTPAuthSTARTTLS:  NomMethodeAuth = "STARTTLS"
        Case Else:                  NomMethodeAuth = Null
    End Select
End Function

' Calcule la r�ponse pour une authentification de type CRAM-MD5 (RFC-2195)
Function CRAM_MD5(ByVal spDefiSrv As String) As String
    Dim s1 As String

    ' D�coder le d�fi d�cod� du serveur
    If spDefiSrv Like "### *" Then Mid$(spDefiSrv, 1, 3) = "   "
    spDefiSrv = Dec_Base64(spDefiSrv)

    ' Obtenir le HMAC
    s1 = HMAC_MD5(dtuEtatSyst.Serveur.OptionsESMTP.AUTH.MotDePasse, spDefiSrv)

    ' Concat�ner UserName et HMAC
    s1 = dtuEtatSyst.Serveur.OptionsESMTP.AUTH.Utilisateur & " " & s1

    ' Convertir en Base64 et sortir le r�sultat
    CRAM_MD5 = Enc_Base64(s1)
End Function

' Analyse le d�fi envoy� par le serveur et calcule la r�ponse ad�quate.
' Si le d�fi est invalide, retourne ""
Function DIGEST_MD5(ByVal spDefiSrv As String) As String
    Dim v As Variant, i As Integer, j As Integer, sChamp As String, sVal As String, dtuDefi As tuD_MD5_S
    Dim bAbandon As Boolean
    Dim A1 As String, A2 As String, sCNONCE As String, sReponse As String

    ' Etape 1. D�coder le d�fi
    ' ========================
    spDefiSrv = Dec_Base64(spDefiSrv)
    spDefiSrv = Remplacer(spDefiSrv, """", "")                      ' Retirer les quotes.

    ' S�parer les diff�rents �l�ments du d�fi
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

    ' Contr�ler la pr�sence des champs obligatoires.
    If Len(dtuDefi.nonce) = 0 Then bAbandon = True
    If Len(dtuDefi.algorithm) = 0 Then bAbandon = True

    ' Doublon ou champ requis manquant. On abandonne la connexion.
    If bAbandon Then Exit Function

    ' Compl�ter les valeurs par d�faut pour les champs facultatifs qui n'ont pas �t� renseign�s par le serveur
    With dtuDefi
        If Len(.realm) = 0 Then .realm = dtuEtatSyst.Serveur.NomSrv
        If Len(.qop) = 0 Then .qop = "auth"
        If .maxbuf = 0 Then .maxbuf = 65536
        If Len(.Charset) = 0 Then .Charset = "ISO 8859-1"
    End With

    ' Etape 2. Pr�parer la r�ponse
    ' ============================

    sCNONCE = Enc_Base64(Alea(32))                          ' Cr�er une chaine al�atoire de 32 caract�res.

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

    ' 1. Compl�ter le mot de passe avec des 0 � concurrence de B octets
    Mid$(s1, 1, Len(sSecret)) = sSecret

    ' 2. Faire un OU exclusif entre (1) et ipad
    s2 = strXOR(s1, ipad)

    ' 3. Concat�ner (2) et le texte
    s2 = s2 & sTexte

    ' 4. Calculer le MD5 de (3)
    s2 = MD5(s2)

    ' 5. Faire un OU exclusif entre (1) et opad
    s3 = strXOR(s1, opad)

    ' 6. Concat�ner (5) et (4)
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

    ' R�solution de nom impossible...
    If ptrHosent = 0 Then Exit Function

#If Win64 Then
    ptrAdresse = ptrHosent + 24
#Else
    ptrAdresse = ptrHosent + 12                                     ' L'adresse IP est 12 octets apr�s le d�but
#End If

    CopyMemory ptrAdresse, ByVal ptrAdresse, 4
    CopyMemory ptrAdrIP, ByVal ptrAdresse, 4
    CopyMemory ByVal sAdresse, ByVal ptrAdrIP, 4

    ' Convertir en chaine, s�par�e par des points
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
