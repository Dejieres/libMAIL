Option Compare Database
Option Explicit
Option Private Module

' Copyright 2009-2014 Denis SCHEIDT
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




' Ces constantes ne sont accessibles qu'� l'aide des fonctions d�finies plus bas.
' -------------------------------------------------------------------------------
Private Const VersionProgramme  As String = "1.49"
Private Const cCopyRight        As String = "Copyright 2009-2014 - Denis SCHEIDT"

' Si le nom de la table contient des espaces, pensez � le mettre [entre crochets].
Private Const BoiteMail         As String = "BoiteMail"


' Variable d'�tat interne du serveur
' ----------------------------------
'
' D�clarations de Types
' ---------------------
'
' Sous-types
'   -- Serveur -------------------------------------
Type tuServeur                                      ' Le Type est public pour la biblioth�que.
    NomSrv                      As String           ' Nom du serveur de messagerie SMTP.
    PortSrv                     As Long             ' Port de connexion au serveur SMTP (25 par d�faut)
    HELOdomain                  As String           ' Identification du poste de l'�metteur.
    OptionsESMTP                As tuESMTP          ' Param�tres ESMTP transmis par ESMTPLance.
    EnvoiQuitte                 As Boolean          ' Quitter le serveur SMTP � la fin de l'envoi ?
    DelaiVerif                  As Integer          ' Intervalle de scrutation de la BoiteMail, en minutes.
    DelaiReponse                As Integer          ' D�lai de r�ponse accord� au serveur distant (5 mn par d�faut).
    Annule                      As Boolean          ' Demande d'annulation de l'envoi en cours.
End Type
'   -- Journal -------------------------------------
Private Type tuJournal
    FichierJnl                  As String           ' Nom et chemin d'acc�s au fichier journal.
    Journal                     As String           ' Journal de session.
    IxDebut                     As Long             ' Pointeur de d�but de journal (variable circulaire).
    LogComm                     As Boolean          ' D�termine si le journal de communication est �crit.
    LogData                     As Boolean          ' Les commandes "DATA" sont elles �crites dans le journal ?
End Type
'   -- Tray ----------------------------------------
Private Type tuTray
    EtatMenu                    As Long             ' Options de menu activ�es/d�sactiv�es
    nid                         As NOTIFYICONDATA   ' Variable pour l'ic�ne de notification.
End Type

' Type principal -----------------------------------
Type tuEtatSyst
    Serveur                     As tuServeur        ' Options du serveur SMTP
    Journal                     As tuJournal        ' Options du Journal
    EtatSrv                     As tuEtatSRV        ' Etat du serveur SMTP
    Tray                        As tuTray           ' Menu et ic�ne
    IDLang                      As Long             ' Identifiant de langue (1033=Anglais, 1036=Fran�ais, etc...)
End Type


Public dtuEtatSyst              As tuEtatSyst
Public Const lLnMaxJnl          As Long = 64000     ' Le maximum qui puisse tenir dans un TextBox.




' Retourne la cha�ne de version de la biblioth�que
Function VersionProg() As String
    VersionProg = VersionProgramme
End Function

Function CopyRight() As String
    CopyRight = cCopyRight
End Function

' Retourne le nom de la table g�rant la bo�te mail
Function TableMail() As String
    TableMail = BoiteMail
End Function

' Retourne la version Access
Function VersionXS() As String
    VersionXS = SysCmd(acSysCmdAccessVer)
    Select Case Val(VersionXS)
        Case 8:     VersionXS = VersionXS & " (Access 97)"
        Case 9:     VersionXS = VersionXS & " (Access 2000)"
        Case 10:    VersionXS = VersionXS & " (Access 2002/XP)"
        Case 11:    VersionXS = VersionXS & " (Access 2003)"
        Case 12:    VersionXS = VersionXS & " (Access 2007)"
        Case 14:    VersionXS = VersionXS & " (Access 2010)"
        Case 15:    VersionXS = VersionXS & " (Access 2013)"
        Case Else:  ' Rien
    End Select
End Function

' Retourne la version de Windows
' Source : http://msdn.microsoft.com/en-us/library/ms724832%28VS.85%29.aspx
' 5.0 = Windows 2000
' 5.1 = Windows XP
' 5.2 = Windows XP 64 bits, Windows server 2003
' 6.0 = Windows Vista, Windows Server 2008
' 6.1 = Windows 7, Windows Server 2008 R2
' 6.2 = Windows 8
Function VersionWin() As String
    Dim vInfo As OSVERSIONINFO

    vInfo.dwOSVersionInfoSize = Len(vInfo)
    If GetVersionEx(vInfo) = 0 Then
        VersionWin = "0.0"
    Else
        VersionWin = vInfo.dwMajorVersion & "." & vInfo.dwMinorVersion
    End If

    Select Case Val(VersionWin)
        Case 5:     VersionWin = VersionWin & " (Windows 2000)"
        Case 5.1:   VersionWin = VersionWin & " (Windows XP)"
        Case 5.2:   VersionWin = VersionWin & " (Windows XP64/Server 2003/Server 2003 R2)"
        Case 6:     VersionWin = VersionWin & " (Windows Vista/Server 2008)"
        Case 6.1:   VersionWin = VersionWin & " (Windows 7/Server 2008 R2)"
        Case 6.2:   VersionWin = VersionWin & " (Windows 8/Server 2012)"
        Case 6.3:   VersionWin = VersionWin & " (Windows 8.1/Server 2012 R2)"
        Case Else:  ' Rien
    End Select
End Function

' Retourne la chaine de langue du syst�me d'exploitation.
Function LangueSyst(lID As Integer) As String
    Dim s As String, l As Long

    s = String$(511, 0)
    l = VerLanguageName(lID, s, 511)
    LangueSyst = Left$(s, l)
End Function

Function Plateforme() As String
    #If Win64 Then
        Plateforme = "Win64"
    #ElseIf Win32 Then
        Plateforme = "Win32"
    #ElseIf Win16 Then
        Plateforme = "Win16"
    #Else
        Plateforme = Traduit("glb_platform", "Inconnue")
    #End If
End Function