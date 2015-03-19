Version = 17
VersionRequired = 17
Checksum = 306782919
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView = 0
    ScrollBars = 0
    BorderStyle = 1
    PictureAlignment = 2
    DatasheetGridlinesBehavior = 3
    GridY = 10
    Width = 5842
    DatasheetFontHeight = 10
    ItemSuffix = 21
    Left = 1680
    Top = 1590
    Right = 9135
    Bottom = 5730
    TimerInterval = 1000
    DatasheetGridlinesColor = 12632256
    RecSrcDt = Begin
        0x0d17a51c9489e340
    End
    DatasheetFontName ="Arial"
    OnTimer ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle = 0
        End
        Begin Line
            Width = 1701
        End
        Begin CommandButton
            Width = 1701
            Height = 283
            FontSize = 8
            FontWeight = 400
            ForeColor = -2147483630
            FontName ="MS Sans Serif"
        End
        Begin FormHeader
            Height = 0
            BackColor = -2147483633
            Name ="EntêteFormulaire"
        End
        Begin Section
            Height = 1987
            BackColor = -2147483633
            Name ="Détail"
            Begin
                Begin Label
                    BackStyle = 1
                    OverlapFlags = 93
                    TextAlign = 2
                    Left = 1417
                    Top = 680
                    Width = 3240
                    Height = 285
                    FontWeight = 700
                    BackColor = -2147483645
                    ForeColor = -2147483639
                    Name ="lblProgBarFond"
                    Caption =" "
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle = 1
                    OverlapFlags = 223
                    Left = 1416
                    Top = 672
                    Width = 900
                    Height = 270
                    FontSize = 12
                    FontWeight = 700
                    BackColor = -2147483646
                    Name ="lblProgBarEchelle"
                    Caption =" "
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags = 93
                    Left = 56
                    Top = 56
                    Width = 1361
                    Height = 284
                    FontWeight = 700
                    Name ="lblEtat"
                    Caption ="Etat :"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle = 1
                    OverlapFlags = 95
                    TextAlign = 2
                    Left = 1417
                    Top = 56
                    Width = 1871
                    Height = 284
                    FontWeight = 700
                    BackColor = -2147483633
                    Name ="lblEtatSRV"
                    Caption ="-- Inconnu --"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags = 95
                    Left = 56
                    Top = 340
                    Width = 1361
                    Height = 283
                    FontWeight = 700
                    Name ="lblMsg"
                    Caption ="Message :"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags = 87
                    TextAlign = 2
                    Left = 1417
                    Top = 340
                    Width = 1871
                    Height = 283
                    Name ="lblMessage"
                    Caption ="0 sur 0"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags = 223
                    Left = 56
                    Top = 680
                    Width = 1367
                    Height = 283
                    FontWeight = 700
                    Name ="lblProgres"
                    Caption ="Progression :"
                    FontName ="Arial"
                End
                Begin Label
                    SpecialEffect = 3
                    OverlapFlags = 215
                    TextAlign = 2
                    Left = 1417
                    Top = 680
                    Width = 3220
                    Height = 283
                    FontWeight = 700
                    ForeColor = -2147483639
                    Name ="lblProgBarCadre"
                    Caption ="0 %"
                    FontName ="Arial"
                End
                Begin Label
                    SpecialEffect = 3
                    BackStyle = 1
                    OldBorderStyle = 1
                    OverlapFlags = 85
                    TextAlign = 2
                    Left = 1417
                    Top = 1020
                    Width = 3232
                    Height = 238
                    FontWeight = 700
                    BackColor = -2147483646
                    ForeColor = -2147483624
                    Name ="lblProgBarTxt"
                    Caption ="0 Kio sur 0 Kio"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags = 85
                    TextAlign = 2
                    Left = 4710
                    Top = 705
                    Width = 1087
                    Height = 238
                    Name ="lblDebit"
                    Caption ="0 Kio/s"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags = 93
                    Left = 56
                    Top = 1303
                    Width = 1361
                    Height = 284
                    FontWeight = 700
                    Name ="lblTpsEcou"
                    Caption ="Temps écoulé :"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags = 93
                    Left = 3174
                    Top = 1303
                    Width = 1361
                    Height = 284
                    FontWeight = 700
                    Name ="lblTpsRest"
                    Caption ="Temps restant :"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags = 95
                    TextAlign = 2
                    Left = 1417
                    Top = 1303
                    Width = 1139
                    Height = 283
                    Name ="lblTpsEcoule"
                    Caption ="0 s."
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags = 87
                    TextAlign = 2
                    Left = 4535
                    Top = 1303
                    Width = 1139
                    Height = 283
                    Name ="lblTpsRestant"
                    Caption ="0 s."
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags = 85
                    Left = 4535
                    Top = 1644
                    Width = 1130
                    Height = 343
                    Name ="cmdAPropos"
                    Caption ="A propos..."
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags = 95
                    Left = 56
                    Top = 1587
                    Width = 1361
                    Height = 284
                    FontWeight = 700
                    Name ="lblNextScan"
                    Caption ="Prochaine scrut. :"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags = 87
                    TextAlign = 2
                    Left = 1417
                    Top = 1587
                    Width = 1703
                    Height = 283
                    Name ="lblScrutSvte"
                    Caption ="jj/mm/aaaa hh:nn:ss"
                    FontName ="Arial"
                End
            End
        End
        Begin FormFooter
            Height = 510
            BackColor = -2147483633
            Name ="PiedFormulaire"
            Begin
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags = 85
                    Left = 56
                    Top = 113
                    Width = 1644
                    Height = 340
                    Name ="cmdEnvoie"
                    Caption ="Envoie maintenant"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Envoie immédiatement les messages en attente."
                End
                Begin CommandButton
                    OverlapFlags = 85
                    Left = 1814
                    Top = 113
                    Width = 1644
                    Height = 340
                    TabIndex = 1
                    Name ="cmdJournal"
                    Caption ="Voir le journal"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Affiche le journal des événements."
                End
                Begin CommandButton
                    OverlapFlags = 85
                    Left = 4025
                    Top = 113
                    Width = 1644
                    Height = 340
                    TabIndex = 2
                    Name ="cmdFermer"
                    Caption ="Fermer"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Ferme ce formulaire."
                End
                Begin Line
                    OverlapFlags = 85
                    SpecialEffect = 5
                    Left = 60
                    Top = 15
                    Width = 5782
                    Name ="Trait15"
                End
            End
        End
    End
End
CodeBehindForm
Option Compare Database
Option Explicit

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



' Procédure de traduction de l'interface.
Public Sub ChangeLang()
    Static T9N_org() As String

    Call LangueCtls(Me.Form, T9N_org())
End Sub



' Mise à jour des informations d'état
Private Sub MAJ()
    Dim lTps As Long, PC As Single, lDebit As Long, lTpsReste As Long

    With Me.lblEtatSRV
        Select Case dtuEtatSyst.EtatSrv.Etat
            Case lmlSrvDecharge
                .Caption = Traduit("etat_inactif", "** Inactif **")
                .ForeColor = RGB(255, 255, 255) ' Blanc
                .BackColor = 0                  ' sur Noir

            Case lmlSrvSuspendu
                .Caption = Traduit("icn_paused", "Suspendu.")
                .ForeColor = RGB(255, 0, 0)     ' Rouge
                .BackColor = vbButtonFace       ' Couleur du formulaire

            Case lmlSrvAttente
                .Caption = Traduit("icn_wait", "En attente.")
                .ForeColor = RGB(255, 128, 0)   ' Orange
                .BackColor = vbButtonFace       ' Couleur du formulaire

            Case lmlSrvEnCours
                .Caption = Traduit("icn_sending", "Envoi en cours.")
                .ForeColor = RGB(0, 128, 0)     ' Vert
                .BackColor = vbButtonFace       ' Couleur du formulaire

            Case lmlSrvConnexion
                .Caption = Traduit("etat_connect", "Connexion...")
                .ForeColor = 0
                .BackColor = vbButtonFace

            Case lmlSrvExecCmd
                .Caption = Traduit("etat_demarre", "Démarrage...")
                .ForeColor = 0
                .BackColor = vbButtonFace

        End Select
    End With

    Me.lblMessage.Caption = dtuEtatSyst.EtatSrv.MessageEnCours & Traduit("etat_msgsur", " sur") & " " & dtuEtatSyst.EtatSrv.MessagesTotal
    Me.lblProgBarTxt.Caption = dtuEtatSyst.EtatSrv.OctetsEnvoyes \ 1024 & Traduit("etat_kiosur", " Kio sur") & " " & dtuEtatSyst.EtatSrv.OctetsTotal \ 1024 & Traduit("etat_kio", " Kio.")

    If dtuEtatSyst.EtatSrv.EnvoiDebut <> CDate(0) Then
        If dtuEtatSyst.EtatSrv.EnvoiFin = CDate(0) Then
            lTps = DateDiff("s", dtuEtatSyst.EtatSrv.EnvoiDebut, Now())
        Else
            lTps = DateDiff("s", dtuEtatSyst.EtatSrv.EnvoiDebut, dtuEtatSyst.EtatSrv.EnvoiFin)
        End If
    End If
    Me.lblTpsEcoule.Caption = lTps & " s."

    ' Calcul du pourcentage.
    If dtuEtatSyst.EtatSrv.OctetsTotal > 0 Then PC = dtuEtatSyst.EtatSrv.OctetsEnvoyes / dtuEtatSyst.EtatSrv.OctetsTotal
    If PC > 1 Then PC = 1

    ' Barre de progression
    Me.lblProgBarEchelle.Width = Me.lblProgBarCadre.Width * PC
    Me.lblProgBarCadre.Caption = Format$(PC, "0 %")

    ' Débit
    If lTps <> 0 Then lDebit = dtuEtatSyst.EtatSrv.OctetsEnvoyes / 1024 / lTps
    Me.lblDebit.Caption = lDebit & Traduit("etat_kio", " Kio/s.")

    ' Temps restant
    If lDebit <> 0 Then lTpsReste = (dtuEtatSyst.EtatSrv.OctetsTotal - dtuEtatSyst.EtatSrv.OctetsEnvoyes) / 1024 / lDebit
    Me.lblTpsRestant.Caption = lTpsReste & " s."

    Me.lblScrutSvte.Caption = IIf(dtuEtatSyst.EtatSrv.ScrutSvte = 0, _
                                  Traduit("etat_inconnu", "** Inconnue **"), _
                                  Format$(dtuEtatSyst.EtatSrv.ScrutSvte, Traduit("etat_fmtdate", "dd/mm/yyyy hh:nn:ss")))

    Me.cmdAPropos.SetFocus
    ' Bouton 'Envoie maintenant'
    Me.cmdEnvoie.Enabled = (dtuEtatSyst.EtatSrv.Etat = lmlSrvAttente)
End Sub


Private Sub cmdAPropos_Click()
    DoCmd.OpenForm "frm_APropos"
End Sub

Private Sub cmdEnvoie_Click()
    Call SMTPEnvoieMaintenant
End Sub

Private Sub cmdFermer_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub cmdJournal_Click()
    Call SMTPFormJnl
End Sub

Private Sub Form_Load()
    Call Me.ChangeLang

    Me.Caption = "libMAIL - version " & VersionProg()
    Me.lblProgBarEchelle.Width = 0

    Call MAJ
End Sub

Private Sub Form_Timer()
    Call MAJ
End Sub