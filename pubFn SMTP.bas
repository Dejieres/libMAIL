Option Compare Database
Option Explicit

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





' Proc�dure permettant d'initialiser et de lancer le serveur SMTP
Sub SMTPLance(sNomSrv As String, _
              Optional sHELOdomain As Variant, _
              Optional bLogData As Boolean = False, _
              Optional bLogComm As Boolean = True, _
              Optional sFichJnl As Variant, _
              Optional bEnvoiQuitte As Boolean = True, _
              Optional iDelaiVerif As Integer = 30, _
              Optional iDelaiReponse As Integer = 300)
    Dim OptESMTP As tuESMTP

    Call ESMTPLance(sNomSrv, OptESMTP, sHELOdomain, bLogData, bLogComm, sFichJnl, bEnvoiQuitte, iDelaiVerif, iDelaiReponse)
End Sub


' Proc�dure permettant d'initialiser et de lancer le serveur SMTP, avec options �tendues
'
' Un param�tre type utilisateur ne peut �tre optionnel...
' Utiliser SMTPLance si l'authentification n'est pas n�cessaire.
Sub ESMTPLance(sNomSrv As String, _
               OptionsESMTP As tuESMTP, _
               Optional sHELOdomain As Variant, _
               Optional bLogData As Boolean = False, _
               Optional bLogComm As Boolean = True, _
               Optional sFichJnl As Variant, _
               Optional bEnvoiQuitte As Boolean = True, _
               Optional ByVal iDelaiVerif As Integer = 30, _
               Optional ByVal iDelaiReponse As Integer = 300)

    Dim s As String

    If Not VerifieBAL() Then Exit Sub                               ' La table BoiteMail doit exister

    ' Si une connexion ou un envoi est en cours, cette commande est invalide.
    If dtuEtatSyst.EtatSrv.Etat = lmlSrvEnCours Or _
       dtuEtatSyst.EtatSrv.Etat = lmlSrvConnexion Or _
       dtuEtatSyst.EtatSrv.Etat = lmlSrvAnnulation Then Exit Sub

    ' Le fichier journal est initialis� s�par�ment.
    Call SMTPJnlFichier(sFichJnl)

    ' Appliquer les param�tres.
    s = IIf(IsMissing(sHELOdomain), myComputerName(), Nz(sHELOdomain, ""))
    Call SMTPChange(sNomSrv, s, bLogData, bLogComm, , bEnvoiQuitte, iDelaiVerif, iDelaiReponse)

    ' Initialisation de la variable interne avec les param�tres restants.
    With dtuEtatSyst.Serveur
        .OptionsESMTP = OptionsESMTP
        .Annule = False
    End With

    ' Le formulaire est lanc� cach�.
    DoCmd.OpenForm "frm_SMTP", acNormal, , , , acHidden
    Call Application.Forms.frm_SMTP.Demarrer
End Sub

' Annule l'envoi en cours
Sub SMTPAnnule()
    Select Case dtuEtatSyst.EtatSrv.Etat
        Case lmlSrvEnCours, lmlSrvConnexion
            dtuEtatSyst.Serveur.Annule = True
    End Select
End Sub

' Permet de changer certaines options du serveur.
Sub SMTPChange(Optional NomSrv As Variant, _
               Optional HELOdomain As Variant, _
               Optional LogData As Variant, _
               Optional LogComm As Variant, _
               Optional FichJnl As Variant, _
               Optional EnvoiQuitte As Variant, _
               Optional DelaiVerif As Variant, _
               Optional DelaiReponse As Variant)

    ' Seulement si le serveur est en attente, suspendu ou d�charg�.
    If dtuEtatSyst.EtatSrv.Etat <> lmlSrvAttente And _
       dtuEtatSyst.EtatSrv.Etat <> lmlSrvSuspendu And _
       dtuEtatSyst.EtatSrv.Etat <> lmlSrvDecharge Then Exit Sub

    With dtuEtatSyst.Serveur
        If Not IsMissing(NomSrv) Then
            .NomSrv = Nz(NomSrv, "localhost")
            Call ServPort(.NomSrv, .PortSrv)                        ' S�paration du nom du serveur et du port
        End If
        If Not IsMissing(HELOdomain) Then .HELOdomain = Nz(HELOdomain, "")
        If Not IsMissing(EnvoiQuitte) Then .EnvoiQuitte = Nz(EnvoiQuitte, True)
        If Not IsMissing(DelaiVerif) Then
            .DelaiVerif = Nz(DelaiVerif, 5)
            If .DelaiVerif < 5 Then .DelaiVerif = 5                 ' 5 mn minimum.
        End If
        If Not IsMissing(DelaiReponse) Then
            .DelaiReponse = Nz(DelaiReponse, 60)
            If .DelaiReponse < 60 Then .DelaiReponse = 60           ' 60 secondes minimum.
        End If
    End With

    With dtuEtatSyst.Journal
        If Not IsMissing(FichJnl) Then Call SMTPJnlFichier(FichJnl)
        If Not IsMissing(LogData) Then .LogData = Nz(LogData, False)
        If Not IsMissing(LogComm) Then .LogComm = Nz(LogComm, False)
    End With
End Sub

' D�chargement du formulaire.
Sub SMTPDecharge()
    Select Case dtuEtatSyst.EtatSrv.Etat
        Case lmlSrvSuspendu, lmlSrvAttente
            DoCmd.Close acForm, "frm_SMTP", acSaveNo
    End Select
End Sub

' D�clenche une scrutation imm�diatement
Sub SMTPEnvoieMaintenant()
    Select Case dtuEtatSyst.EtatSrv.Etat
        Case lmlSrvAttente
            Call Application.Forms.frm_SMTP.Envoyer
    End Select
End Sub

' Reprend la scrutation, en appliquant �ventuellement un nouveau d�lai (en minutes).
Sub SMTPRelance(Optional iDelai As Integer = 0)
    Select Case dtuEtatSyst.EtatSrv.Etat
        Case lmlSrvSuspendu, lmlSrvAttente
            If iDelai > 0 Then
                If iDelai < 5 Then iDelai = 5
                dtuEtatSyst.Serveur.DelaiVerif = iDelai              ' Appliquer le nouveau d�lai, si fourni.
            End If
            Call Application.Forms.frm_SMTP.Relancer
    End Select
End Sub

' Proc�dure permettant d'arr�ter le serveur SMTP (sans le d�charger)
Sub SMTPSuspend()
    ' Le serveur ne peut �tre suspendu que lorsqu'il est dans l'�tat Attente
    Select Case dtuEtatSyst.EtatSrv.Etat
        Case lmlSrvAttente
            Call Application.Forms.frm_SMTP.Arreter
    End Select
End Sub


' Fonction publique permettant d'interroger l'�tat du serveur.
Function SMTPEtatSrv() As tuEtatSRV
    SMTPEtatSrv = dtuEtatSyst.EtatSrv
End Function

' Affiche / d�charge le formulaire d'�tat du serveur.
Sub SMTPFormEtat(bAffiche As Boolean)
    If bAffiche Then
        DoCmd.OpenForm "frm_EtatSRV"
    Else
        If FrmEstCharge("frm_EtatSRV") Then DoCmd.Close acForm, "frm_EtatSRV", acSaveNo
    End If
End Sub

' Retourne le contenu du membre Journal de dtuJournal.
Function SMTPJournal() As String
    With dtuEtatSyst.Journal
        If .IxDebut > 0 Then SMTPJournal = Trim$(Mid$(.Journal, .IxDebut) & Left$(.Journal, .IxDebut - 1))
    End With
End Function

' Initialise le membre FichierJnl de la structure dtuJournal.
' Cette proc�dure est �galement utilis�e par (E)SMTPLance.
Sub SMTPJnlFichier(Optional sFichJnl As Variant)
    ' Mise en place et contr�le du fichier journal.
    If IsMissing(sFichJnl) Then
        ' Appliquer la valeur par d�faut, si aucun param�tre n'a �t� founi, et que le membre est vide.
        If Len(dtuEtatSyst.Journal.FichierJnl) = 0 Then
            sFichJnl = "C:\Temp\SMTP_SRV.LOG"
        Else
            sFichJnl = dtuEtatSyst.Journal.FichierJnl               ' Conserver le journal existant
        End If
    End If

    ' Contr�le de validit�.
    If Len(sFichJnl) <> 0 Then                                      ' Si une spec de fichier est fournie,
        sFichJnl = VerifieFich((sFichJnl))                          ' v�rifier qu'elle est valide.
        ' Repli sur Environ$("Temp") si la spec demand�e n'est pas valide
        If Len(sFichJnl) = 0 Then sFichJnl = VerifieFich(Environ$("Temp") & "\SMTP_SRV.LOG")
        If Len(sFichJnl) = 0 Then sFichJnl = VerifieFich(Environ$("Tmp") & "\SMTP_SRV.LOG")
    End If

    dtuEtatSyst.Journal.FichierJnl = Nz(sFichJnl, "")               ' Inscrire le nom et le chemin dans le membre de la DTU.
    If Len(dtuEtatSyst.Journal.Journal) = 0 Then Call SMTPJnlRAZ    ' Initialisation des variables.
End Sub

' Remet � z�ro le membre Journal de dtuJournal.
' Si bDisque est True, efface �galement le fichier disque.
Sub SMTPJnlRAZ(Optional bDisque As Boolean = False)
    ' RAZ de la partie 'm�moire'
    With dtuEtatSyst.Journal
        .Journal = Space$(lLnMaxJnl)
        .IxDebut = 1
    End With

    ' Effacer le fichier disque
    If bDisque Then
        On Error Resume Next
        If FichierExiste(dtuEtatSyst.Journal.FichierJnl) Then Kill dtuEtatSyst.Journal.FichierJnl
        On Error GoTo 0
    End If
End Sub

' Affiche le formulaire de lecture du journal
Sub SMTPFormJnl()
    DoCmd.OpenForm "frm_Journal"
End Sub

' Retourne les informations renvoy�es par la commande EHLO
Function SMTPTest(ByVal sNomSrv As String)
    Dim lRet As Long, lPort As Long, lSock As Long, sRepSrv As String, i As Integer, iDelai As Integer, oEtat As Byte

    ' La fonction ne peut �tre appel�e que si...
    With dtuEtatSyst.EtatSrv
        If .Etat <> lmlSrvAttente And .Etat <> lmlSrvDecharge And .Etat <> lmlSrvSuspendu Then
            SMTPTest = "L'�tat actuel du serveur ne permet pas l'ex�cution de cette commande..."
            Exit Function
        End If
    End With

    ' Sauvegarder le d�lai de r�ponse et l'�tat.
    With dtuEtatSyst
        iDelai = .Serveur.DelaiReponse
        .Serveur.DelaiReponse = 60

        oEtat = .EtatSrv.Etat
        .EtatSrv.Etat = lmlSrvConnexion
    End With

    Call ServPort(sNomSrv, lPort)                                   ' Extraire Serveur et Port.

    SMTPTest = "Connexion � " & sNomSrv & " sur le port " & lPort & vbCrLf
    lRet = CnxServ(sNomSrv, lPort, lSock)
    If lRet = 0 Then                                                ' Connexion OK.
        i = EnvoiCMD(lSock, Null, , , sRepSrv)
        SMTPTest = SMTPTest & sRepSrv & vbCrLf

        If i = 2 Then
            i = EnvoiCMD(lSock, "EHLO " & myComputerName, , , sRepSrv)
            SMTPTest = SMTPTest & "--> EHLO " & myComputerName & vbCrLf & sRepSrv & vbCrLf

            If i = 2 Then
                i = EnvoiCMD(lSock, "QUIT", , , sRepSrv)
                SMTPTest = SMTPTest & "--> QUIT" & vbCrLf & sRepSrv & vbCrLf
            Else
                SMTPTest = SMTPTest & i & " Le serveur rejette la commande EHLO... " & vbCrLf & sRepSrv
            End If
        Else
            SMTPTest = SMTPTest & i & " Le serveur refuse la connexion... " & vbCrLf & sRepSrv
        End If
    Else
        SMTPTest = SMTPTest & "Impossible d'�tablir la connexion..." & vbCrLf & "Erreur " & lRet & ", socket " & lSock
    End If

    Call CnxFin(lSock)

    ' Restaurer le d�lai de r�ponse et l'�tat initiaux.
    With dtuEtatSyst
        .Serveur.DelaiReponse = iDelai
        .EtatSrv.Etat = oEtat
    End With
End Function