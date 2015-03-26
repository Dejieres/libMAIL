Option Compare Database
Option Explicit

' Copyright 2009-2015 Denis SCHEIDT
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





' M�thode et param�tre d'authentification. Ce Type n'est pas export�.
Private Type tuESMTP_AUTH
    Methode         As Integer                              ' M�thode d'authentification (constantes ci-dessous)
    Utilisateur     As String                               ' Nom d'utilisateur
    MotDePasse      As String                               ' Mot de passe associ�
End Type

' Options de message propres aux avis de remise. Ce Type n'est pas export�.
Private Type tuESMTP_DSN
    Notification    As Integer                              ' Type de notification (constantes ci-dessous)
    Retour          As Integer                              ' Type de retour (constantes ci-dessous)
    IDEnveloppe     As String                               ' Identifiant de la transaction pour le retour
End Type

' Options MDN - accus�s de r�ception. Ce Type n'est pas export�.
Private Type tuESMTP_MDN
    Notification    As String                               ' Disposition-Notification-To:  avis de (non-)lecture
    Reception       As String                               ' Return-Receipt-To:            accus� de r�ception
End Type

' Champs d'origine du message
Private Type tuSMTP_ORG
    Repondre        As String                               ' Champ Reply-To (0 - n adresses). R�pondre �.
    Envoyeur        As String                               ' Champ Sender, (0 - 1 adresse)
                                                            ' permet d'envoyer le message au nom de quelqu'un d'autre.
End Type


' --------------------------------------------------------------------------------------------------
' Types publics pour la transmission des options.
' --------------------------------------------------------------------------------------------------

' Constantes et structures de donn�es publiques pour ESMTP-AUTH.
' Options d'authentification.
Public Type tuESMTP                                         ' Options ESMTP impl�ment�es
    AUTH        As tuESMTP_AUTH                             ' Options d'authentification
End Type
Public Const lmlESMTPAuthAucune         As Integer = 0
Public Const lmlESMTPAuthLogin          As Integer = 1
Public Const lmlESMTPAuthPlain          As Integer = 2
Public Const lmlESMTPAuthCRAMMD5        As Integer = 3
Public Const lmlESMTPAuthDIGESTMD5      As Integer = 4
Public Const lmlESMTPAuthSTARTTLS       As Integer = 5


' Options propres aux messages.

' Constantes et structures de donn�es publiques pour ESMTP-DSN -------------------------------------
' Avis de remise et Accus�s de r�ception.
Public Type tuESMTP_MSG
    DSN         As tuESMTP_DSN                              ' Avis de remise
    MDN         As tuESMTP_MDN                              ' Accus�s de r�ception/lecture
    Priorite    As Integer                                  ' Priorit� du message
    ORG         As tuSMTP_ORG                               ' Champs 'origine'
End Type

Public Const lmlESMTPDsnNotifDefaut     As Integer = 0      ' Ne pas ajouter de param�tre � la commande RCPT
Public Const lmlESMTPDsnNotifSucces     As Integer = 1      ' Seulement en cas de remise OK.
Public Const lmlESMTPDsnNotifEchec      As Integer = 2      ' Seulement en cas d'�chec.
Public Const lmlESMTPDsnNotifDelai      As Integer = 4      ' Seulement si la remise est diff�r�e.
Public Const lmlESMTPDsnNotifTous       As Integer = 7      ' Dans tous les cas.
Public Const lmlESMTPDsnNotifJamais     As Integer = 128    ' Ne jamais envoyer d'accus�.

Public Const lmlESMTPDsnRetDefaut       As Integer = 0      ' Ne pas ajouter de param�tre � la commande MAIL.
Public Const lmlESMTPDsnRetHdrs         As Integer = 1      ' Retourne seulement l'en-t�te de message dans l'accus� de r�ception
Public Const lmlESMTPDsnRetFull         As Integer = 2      ' Retourne tout le message dans l'accus� de r�ception

Public Const lmlMsgPrioHte              As Integer = 1      ' Priorit� Haute
Public Const lmlMsgPrioNrm              As Integer = 3      '          Normale
Public Const lmlMsgPrioBas              As Integer = 5      '          Faible


' DTU permettant de m�moriser l'�tat du serveur. ---------------------------------------------------
Public Type tuEtatSRV
    Etat                                As Byte             ' Etat du serveur.
    MessageEnCours                      As Integer          ' Position ordinale du message en cours d'envoi.
    MessagesTotal                       As Integer          ' Nombre total de messages � envoyer dans cette session.
    OctetsEnvoyes                       As Long             ' Volume de donn�es envoy� dans cette session.
    OctetsTotal                         As Long             ' Volume total de donn�es � envoyer dans cette session.
    EnvoiDebut                          As Date             ' Date et heure de d�but de l'envoi pour cette session.
    EnvoiFin                            As Date             ' Date et heure de fin de l'envoi pour cette session.
    Resultat                            As Integer          ' Code de r�sultat en fin de traitement.
    ScrutSvte                           As Date             ' Date et heure de la prochaine scrutation.
End Type
' Constantes d'�tat du serveur ---------------------------------------------------------------------
Public Const lmlSrvDecharge             As Byte = 0         ' Le formulaire frm_SMTP n'est pas charg�.
Public Const lmlSrvSuspendu             As Byte = 1         ' La scrutation est suspendue.
Public Const lmlSrvAttente              As Byte = 2         ' En attente de la prochaine scrutation.
Public Const lmlSrvEnCours              As Byte = 3         ' Envoi des messages en cours.
Public Const lmlSrvAnnulation           As Byte = 4         ' Annulation de l'envoi.
Public Const lmlSrvExecCmd              As Byte = 5         ' Etat transitoire, entre l'appel de la commande et le d�clenchement du Timer.
Public Const lmlSrvConnexion            As Byte = 6         ' En cours de connexion.
' Constantes pour le membre Resultat
Public Const lmlSrvResND                As Integer = 0      ' Pas de r�sultat disponible.
Public Const lmlSrvResRien              As Integer = 1      ' Il n'y avait rien � envoyer.
Public Const lmlSrvResCnx               As Integer = 2      ' Erreur de connexion ou annulation pendant la connexion.
Public Const lmlSrvResOK                As Integer = 3      ' Tous les messages ont �t� envoy�s.
Public Const lmlSrvResErr               As Integer = 4      ' Certains messages n'ont pas �t� envoy�s.

' Corps text ou HTML -------------------------------------------------------------------------------
Type tuMessageMIME
    Texte                               As String           ' Partie text/plain du message
    HTML                                As String           ' Partie HTML du message
End Type


' Constantes pour l'activation/d�sactivation du menu -----------------------------------------------
Public Const lmlMnuSspn                 As Integer = 1      ' Option 'Suspendre'
Public Const lmlMnuRlnc                 As Integer = 2      ' Option 'Relancer'
Public Const lmlMnuEnvM                 As Integer = 4      ' Option 'Envoyer maintenant'
Public Const lmlMnuDech                 As Integer = 8      ' Option 'Decharger'
Public Const lmlMnuAnnE                 As Integer = 16     ' Option 'Annuler l'envoi'
Public Const lmlMnuNMsg                 As Integer = 32     ' Option 'Nouveau message'
Public Const lmlMnuGest                 As Integer = 64     ' Option 'Gestionnaire...'
Public Const lmlMnuEtat                 As Integer = 128    ' Option 'Afficher l'�tat'
Public Const lmlMnuAJnl                 As Integer = 256    ' Option 'Afficher le journal'
Public Const lmlMnuLang                 As Integer = 512    ' Option 'Langue'


' Constantes pour le type d'export des PJ
Public Const lmlPJDonnees               As Integer = 0      ' Utiliser DoCmd.OutputTo pour g�n�rer la PJ
Public Const lmlPJSource                As Integer = 1      ' Utiliser Application.SaveAsText pour g�n�rer la PJ.