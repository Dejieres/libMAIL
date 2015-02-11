Option Compare Database
Option Explicit

' Copyright 2009-2013 Denis SCHEIDT
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





' Méthode et paramètre d'authentification. Ce Type n'est pas exporté.
Private Type tuESMTP_AUTH
    Methode         As Integer                              ' Méthode d'authentification (constantes ci-dessous)
    Utilisateur     As String                               ' Nom d'utilisateur
    MotDePasse      As String                               ' Mot de passe associé
End Type

' Options de message propres aux avis de remise. Ce Type n'est pas exporté.
Private Type tuESMTP_DSN
    Notification    As Integer                              ' Type de notification (constantes ci-dessous)
    Retour          As Integer                              ' Type de retour (constantes ci-dessous)
    IDEnveloppe     As String                               ' Identifiant de la transaction pour le retour
End Type

' Options MDN - accusés de réception. Ce Type n'est pas exporté.
Private Type tuESMTP_MDN
    Notification    As String                               ' Disposition-Notification-To:  avis de (non-)lecture
    Reception       As String                               ' Return-Receipt-To:            accusé de réception
End Type

' Champs d'origine du message
Private Type tuSMTP_ORG
    Repondre        As String                               ' Champ Reply-To (0 - n adresses). Répondre à.
    Envoyeur        As String                               ' Champ Sender, (0 - 1 adresse)
                                                            ' permet d'envoyer le message au nom de quelqu'un d'autre.
End Type


' --------------------------------------------------------------------------------------------------
' Types publics pour la transmission des options.
' --------------------------------------------------------------------------------------------------

' Constantes et structures de données publiques pour ESMTP-AUTH.
' Options d'authentification.
Public Type tuESMTP                                         ' Options ESMTP implémentées
    AUTH        As tuESMTP_AUTH                             ' Options d'authentification
End Type
Public Const lmlESMTPAuthAucune         As Integer = 0
Public Const lmlESMTPAuthLogin          As Integer = 1
Public Const lmlESMTPAuthPlain          As Integer = 2
Public Const lmlESMTPAuthCRAMMD5        As Integer = 3
Public Const lmlESMTPAuthDIGESTMD5      As Integer = 4
Public Const lmlESMTPAuthSTARTTLS       As Integer = 5


' Options propres aux messages.

' Constantes et structures de données publiques pour ESMTP-DSN -------------------------------------
' Avis de remise et Accusés de réception.
Public Type tuESMTP_MSG
    DSN         As tuESMTP_DSN                              ' Avis de remise
    MDN         As tuESMTP_MDN                              ' Accusés de réception/lecture
    Priorite    As Integer                                  ' Priorité du message
    ORG         As tuSMTP_ORG                               ' Champs 'origine'
End Type

Public Const lmlESMTPDsnNotifDefaut     As Integer = 0      ' Ne pas ajouter de paramètre à la commande RCPT
Public Const lmlESMTPDsnNotifSucces     As Integer = 1      ' Seulement en cas de remise OK.
Public Const lmlESMTPDsnNotifEchec      As Integer = 2      ' Seulement en cas d'échec.
Public Const lmlESMTPDsnNotifDelai      As Integer = 4      ' Seulement si la remise est différée.
Public Const lmlESMTPDsnNotifTous       As Integer = 7      ' Dans tous les cas.
Public Const lmlESMTPDsnNotifJamais     As Integer = 128    ' Ne jamais envoyer d'accusé.

Public Const lmlESMTPDsnRetDefaut       As Integer = 0      ' Ne pas ajouter de paramètre à la commande MAIL.
Public Const lmlESMTPDsnRetHdrs         As Integer = 1      ' Retourne seulement l'en-tête de message dans l'accusé de réception
Public Const lmlESMTPDsnRetFull         As Integer = 2      ' Retourne tout le message dans l'accusé de réception

Public Const lmlMsgPrioHte              As Integer = 1      ' Priorité Haute
Public Const lmlMsgPrioNrm              As Integer = 3      '          Normale
Public Const lmlMsgPrioBas              As Integer = 5      '          Faible


' DTU permettant de mémoriser l'état du serveur. ---------------------------------------------------
Public Type tuEtatSRV
    Etat                                As Byte             ' Etat du serveur.
    MessageEnCours                      As Integer          ' Position ordinale du message en cours d'envoi.
    MessagesTotal                       As Integer          ' Nombre total de messages à envoyer dans cette session.
    OctetsEnvoyes                       As Long             ' Volume de données envoyé dans cette session.
    OctetsTotal                         As Long             ' Volume total de données à envoyer dans cette session.
    EnvoiDebut                          As Date             ' Date et heure de début de l'envoi pour cette session.
    EnvoiFin                            As Date             ' Date et heure de fin de l'envoi pour cette session.
    Resultat                            As Integer          ' Code de résultat en fin de traitement.
    ScrutSvte                           As Date             ' Date et heure de la prochaine scrutation.
End Type
' Constantes d'état du serveur ---------------------------------------------------------------------
Public Const lmlSrvDecharge             As Byte = 0         ' Le formulaire frm_SMTP n'est pas chargé.
Public Const lmlSrvSuspendu             As Byte = 1         ' La scrutation est suspendue.
Public Const lmlSrvAttente              As Byte = 2         ' En attente de la prochaine scrutation.
Public Const lmlSrvEnCours              As Byte = 3         ' Envoi des messages en cours.
Public Const lmlSrvAnnulation           As Byte = 4         ' Annulation de l'envoi.
Public Const lmlSrvExecCmd              As Byte = 5         ' Etat transitoire, entre l'appel de la commande et le déclenchement du Timer.
Public Const lmlSrvConnexion            As Byte = 6         ' En cours de connexion.
' Constantes pour le membre Resultat
Public Const lmlSrvResND                As Integer = 0      ' Pas de résultat disponible.
Public Const lmlSrvResRien              As Integer = 1      ' Il n'y avait rien à envoyer.
Public Const lmlSrvResCnx               As Integer = 2      ' Erreur de connexion ou annulation pendant la connexion.
Public Const lmlSrvResOK                As Integer = 3      ' Tous les messages ont été envoyés.
Public Const lmlSrvResErr               As Integer = 4      ' Certains messages n'ont pas été envoyés.

' Corps text ou HTML -------------------------------------------------------------------------------
Type tuMessageMIME
    Texte                               As String           ' Partie text/plain du message
    HTML                                As String           ' Partie HTML du message
End Type


' Constantes pour l'activation/désactivation du menu -----------------------------------------------
Public Const lmlMnuSspn                 As Integer = 1      ' Option 'Suspendre'
Public Const lmlMnuRlnc                 As Integer = 2      ' Option 'Relancer'
Public Const lmlMnuEnvM                 As Integer = 4      ' Option 'Envoyer maintenant'
Public Const lmlMnuDech                 As Integer = 8      ' Option 'Decharger'
Public Const lmlMnuAnnE                 As Integer = 16     ' Option 'Annuler l'envoi'
Public Const lmlMnuNMsg                 As Integer = 32     ' Option 'Nouveau message'
Public Const lmlMnuGest                 As Integer = 64     ' Option 'Gestionnaire...'
Public Const lmlMnuEtat                 As Integer = 128    ' Option 'Afficher l'état'
Public Const lmlMnuAJnl                 As Integer = 256    ' Option 'Afficher le journal'


' Constantes pour le type d'export des PJ
Public Const lmlPJDonnees               As Integer = 0      ' Utiliser DoCmd.OutputTo pour générer la PJ
Public Const lmlPJSource                As Integer = 1      ' Utiliser Application.SaveAsText pour générer la PJ.