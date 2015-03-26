Option Compare Database
Option Explicit
Option Private Module


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



' Met � jour le membre EtatSrv.Etat et affiche l'ic�ne appropri�e.
Sub AffIconeNotifSRV(frm As Form, oEtat As Byte)
    Dim b() As Byte

    ' Mettre � jour le membre EtatSrv.Etat
    dtuEtatSyst.EtatSrv.Etat = oEtat

    ' Choix de l'image � utiliser. Elle est copi�e dans le tableau d'octets.
    Select Case dtuEtatSyst.EtatSrv.Etat
        Case lmlSrvAttente:         b = frm.imgAttente.PictureData
        Case lmlSrvSuspendu:        b = frm.imgSuspendu.PictureData
        Case lmlSrvEnCours:         b = frm.imgEncours.PictureData
        Case lmlSrvAnnulation:      b = frm.imgAnnulation.PictureData
        Case lmlSrvConnexion:       b = frm.imgConnexion.PictureData
        Case lmlSrvExecCmd:         b = frm.imgExecCmd.PictureData
        Case Else:                  b = frm.imgNeutre.PictureData
    End Select

    Call AffIconeNotif(frm.hwnd, b)

    Erase b
End Sub

' Affiche un message dans la zone de notification
' lFlags est l'une des valeurs NIIF_xxx, d�terminant l'image � afficher dans la bulle de notifications.
Sub AffMsgNotif(sMsg As String, lFlags As Long)
    ' Affiche une info-bulle.
    With dtuEtatSyst.Tray.nid
        .szInfoTitle = "libMAIL" & " (" & Application.GetOption("Project Name") & ")" & vbNullChar
        .szInfo = sMsg & vbNullChar
        .dwInfoFlags = lFlags
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE Or NIF_INFO
    End With

    Call Shell_NotifyIcon(NIM_MODIFY, dtuEtatSyst.Tray.nid)
End Sub

' nid : structure pour la fonction API
Sub AffIconeNotif(frmHwnd As Long, b() As Byte)
#If Vba7 Then
    Dim h As LongPtr
#Else
    Dim h As Long
#End If
    Dim bAND() As Byte, bXOR() As Byte, bmINFO As BITMAPINFOHEADER

    ' Informations d'en-t�te de l'image.
    With bmINFO
        .lSize = b(3) * &H1000000 Or b(2) * &H10000 Or b(1) * &H100& Or b(0)
        .lWidth = b(7) * &H1000000 Or b(6) * &H10000 Or b(5) * &H100& Or b(4)
        .lHeight = b(11) * &H1000000 Or b(10) * &H10000 Or b(9) * &H100& Or b(8)
        .iPlanes = b(13) * &H100& Or b(12)
        .lBitCount = b(15) * &H100& Or b(14)
        .lSizeImage = b(23) * &H1000000 Or b(22) * &H10000 Or b(21) * &H100& Or b(20)

        Call bmMasqueXOR(b, bmINFO, bXOR)                   ' Convertir l'image et cr�er le masque XOR
        Call bmMasqueET(bXOR, bmINFO, bAND)                 ' Cr�er le masque AND, pour la transparence.

        ' Cr�er l'ic�ne proprement dite, et r�cup�rer son Handle.
        h = CreateIcon(frmHwnd, .lWidth, .lHeight, .iPlanes, .lBitCount, bAND(0), bXOR(0))
    End With

    Select Case dtuEtatSyst.Tray.nid.hIcon
        Case 0                                              ' ***** Premi�re cr�ation de l'ic�ne
            With dtuEtatSyst.Tray.nid                       ' Initialisation
                .cbSize = Len(dtuEtatSyst.Tray.nid)
                .hwnd = frmHwnd
                .uId = 0
                ' NIF_INFO n'est pas ajout� ici, sinon la derni�re notification s'affiche lors du changement d'ic�ne.
                .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
                .uCallBackMessage = WM_MOUSEMOVE
                .hIcon = h
                .szTip = "libMAIL v." & VersionProg() & " (" & Application.GetOption("Project Name") & ")" & vbNullChar
            End With
            Call Shell_NotifyIcon(NIM_ADD, dtuEtatSyst.Tray.nid)
'            With nid
'                .uTimeOut = NOTITYICON_VERSION
'            End With
'            Call Shell_NotifyIcon(NIM_SETVERSION, nid)      ' Mode de fonctionnement, Windows 2000 +

        Case Else                                           ' ***** Modification d'ic�ne.
            With dtuEtatSyst.Tray.nid
                Call DestroyIcon(.hIcon)                    ' D�truire l'ic�ne existante.
                .hIcon = h                                  ' Affecter la nouvelle ic�ne.
                ' NIF_INFO n'est pas ajout� ici, sinon la derni�re notification s'affiche lors du changement d'ic�ne.
                .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
                ' Message au survol de la souris (128 car. maxi)
                .szTip = "libMAIL v." & VersionProg() & " (" & Application.GetOption("Project Name") & ")" & vbCrLf & _
                         "Serv.: " & dtuEtatSyst.Serveur.NomSrv & ":" & dtuEtatSyst.Serveur.PortSrv & vbCrLf & _
                         "Auth.: " & Nz(NomMethodeAuth(dtuEtatSyst.Serveur.OptionsESMTP.AUTH.Methode), Traduit("icn_unknown", "Inconnue")) & vbCrLf & _
                         Choose(dtuEtatSyst.EtatSrv.Etat + 1, Traduit("icn_unload", "D�charg�."), _
                                                              Traduit("icn_paused", "Suspendu."), _
                                                              Traduit("icn_wait", "Attente.") & vbCrLf & Traduit("icn_nextscan", "Prochaine scrut. : ") & dtuEtatSyst.EtatSrv.ScrutSvte, _
                                                              Traduit("icn_sending", "En cours."), _
                                                              Traduit("icn_cancel", "Annulation de l'envoi. Le message en cours se termine.")) & _
                         vbNullChar
            End With
            Call Shell_NotifyIcon(NIM_MODIFY, dtuEtatSyst.Tray.nid) ' Modifier l'ic�ne.

    End Select
End Sub

' Retourne la largeur d'un pixel en twips
Function TwipsPerPixelX() As Single
#If Vba7 Then
  Dim hDC As LongPtr
#Else
  Dim hDC As Long
#End If

  hDC = GetDC(HWND_DESKTOP)
  TwipsPerPixelX = 1440& / GetDeviceCaps(hDC, LOGPIXELSX)
  Call ReleaseDC(HWND_DESKTOP, hDC)
End Function


' Cr�ation du menu
Sub CreeMenu()
    Dim cb As CommandBar, cbc As CommandBarControl, cbc1 As CommandBarPopup, s As String
    Dim rs As DAO.Recordset

    ' V�rifier si la barre de commande existe d�j�.
    ' (d�chargement puis rechargement de frm_SMTP sans quitter l'application, par exemple)
    For Each cb In Application.CommandBars
        If cb.Name = "CB_libMAIL" Then
            s = cb.Name
            Exit For
        End If
    Next cb
    Set cb = Nothing
    If Len(s) <> 0 Then Exit Sub

    ' Ajouter la barre et les options (barre temporaire)
    Set cb = CommandBars.Add("CB_libMAIL", msoBarPopup, , True)
    Set cbc = cb.Controls.Add(msoControlButton)
    With cbc
        .FaceId = 0
'        .Caption = Traduit("mnu_pause", "&Suspendre")
        .OnAction = "SMTPSuspend"
        .Style = msoButtonIconAndCaption
        .FaceId = 189
        .Tag = lmlMnuSspn
    End With

    Set cbc = cb.Controls.Add(msoControlButton)
    With cbc
        .FaceId = 0
'        .Caption = Traduit("mnu_resume", "&Relancer")
        .OnAction = "SMTPRelance"
        .Style = msoButtonIconAndCaption
        .FaceId = 126
        .Tag = lmlMnuRlnc
    End With

    Set cbc = cb.Controls.Add(msoControlButton)
    With cbc
        .FaceId = 0
'        .Caption = Traduit("mnu_sendnow", "&Envoyer maintenant")
        .OnAction = "SMTPEnvoieMaintenant"
        .Style = msoButtonIconAndCaption
        .FaceId = 325
        .Tag = lmlMnuEnvM
    End With

    Set cbc = cb.Controls.Add(msoControlButton)
    With cbc
        .FaceId = 0
'        .Caption = Traduit("mnu_unload", "&D�charger")
        .OnAction = "SMTPDecharge"
        .Style = msoButtonIconAndCaption
        .FaceId = 2186
        .Tag = lmlMnuDech
    End With

    Set cbc = cb.Controls.Add(msoControlButton)
    With cbc
        .FaceId = 0
'        .Caption = Traduit("mnu_cancel", "Ann&uler l'envoi")
        .OnAction = "SMTPAnnule"
        .BeginGroup = True
        .Style = msoButtonIconAndCaption
        .FaceId = 1019
        .Tag = lmlMnuAnnE
    End With

    Set cbc = cb.Controls.Add(msoControlButton)
    With cbc
        .FaceId = 0
'        .Caption = Traduit("mnu_newmsg", "&Nouveau message...")
        .OnAction = "mnuCreeMail"
        .BeginGroup = True
        .Style = msoButtonIconAndCaption
        .FaceId = 719
        .Tag = lmlMnuNMsg
    End With

    Set cbc = cb.Controls.Add(msoControlButton)
    With cbc
        .FaceId = 0
'        .Caption = Traduit("mnu_mbm", "&Gestionnaire...")
        .OnAction = "mnuGestMail"
        .BeginGroup = True
        .Style = msoButtonIconAndCaption
        .FaceId = 721
        .Tag = lmlMnuGest
    End With

    Set cbc = cb.Controls.Add(msoControlButton)
    With cbc
        .FaceId = 0
'        .Caption = Traduit("mnu_status", "&Afficher l'�tat")
        .OnAction = "mnuAffEtat"
        .Style = msoButtonIconAndCaption
        .FaceId = 352
        .Tag = lmlMnuEtat
    End With

    Set cbc = cb.Controls.Add(msoControlButton)
    With cbc
        .FaceId = 0
'        .Caption = Traduit("mnu_log", "Afficher le &journal...")
        .OnAction = "SMTPFormJnl"
        .Style = msoButtonIconAndCaption
        .FaceId = 195
        .Tag = lmlMnuAJnl
    End With

    ' Menu Langues et sous-menus.
    Set cbc1 = cb.Controls.Add(msoControlPopup)
    With cbc1
'        .Caption = "&Langue"
        .BeginGroup = True
        .Tag = lmlMnuLang

        ' Ajouter les autres langues.
        Set rs = CodeDb.OpenRecordset("SELECT DISTINCT IDLang FROM T9N", dbOpenDynaset, 0, dbReadOnly)
        Do While Not rs.EOF
            Set cbc = cbc1.Controls.Add(msoControlButton)
            With cbc
                .Caption = LangueSyst(rs!IDLang) & " - " & rs!IDLang
                .OnAction = "=ChangeLang(" & rs!IDLang & ")"
                .Tag = rs!IDLang
            End With
        
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        Set cbc = cbc1.Controls.Add(msoControlButton)
        With cbc
            .BeginGroup = True
            .Caption = "Fran�ais (France) - 1036"                       ' C'est la langue par d�faut.
            .OnAction = "=ChangeLang(1036)"
            .Tag = 1036
        End With

    End With

    Set cbc = cb.Controls.Add(msoControlButton)
    With cbc
        .FaceId = 0
'        .Caption = Traduit("mnu_about", "A &propos...")
        .OnAction = "mnuAPropos"
        .BeginGroup = True
        .Style = msoButtonIconAndCaption
        .FaceId = 49
        .Tag = -1
    End With

    Call LangueMenu                                                     ' Traduction du menu.

    Set cb = Nothing
    Set cbc = Nothing
End Sub



' Convertit le bitmap dans la bonne profondeur de couleurs.
' Inverse �galement les lignes par la m�me occasion.
' Retourne le bitmap converti, d�barrass� de l'en-t�te.
Private Sub bmMasqueXOR(bIn() As Byte, dtuEnTete As BITMAPINFOHEADER, bOut() As Byte)
#If Vba7 Then
    Dim hDC As LongPtr
#Else
    Dim hDC As Long
#End If
        Dim b() As Byte, lBPPEcran As Long, i As Long, j As Long, k As Long, iBloc As Integer
    Dim dtuRGB As RGBQUAD

    With dtuEnTete
        ' Redimensionner le tableau de travail, en enlevant l'en-t�te
        ReDim b(.lSizeImage - 1)

        ' Premier passage, on inverse les lignes, car l'ic�ne s'affiche la t�te en bas...
        ' et on �limine l'en-t�te.
        k = .lSizeImage
        iBloc = .lWidth * .lBitCount \ 8                ' Une ligne de pixels.
        For i = .lSize To UBound(bIn) Step iBloc
            k = k - iBloc                               ' Position de destination.
            For j = 0 To iBloc - 1                      ' Copier le bloc vers la destination.
                b(k + j) = bIn(i + j)
            Next j
        Next i

        ' R�cup�rer la profondeur de couleur actuelle de l'�cran
        hDC = GetDC(HWND_DESKTOP)
        lBPPEcran = GetDeviceCaps(hDC, BITSPIXEL)
        Call ReleaseDC(HWND_DESKTOP, hDC)

        ' Redimensionner le tableau de l'image de sortie, � l'aide la profondeur �cran.
        ReDim bOut((.lWidth * .lHeight * lBPPEcran \ 8) - 1)

        ' Si les profondeurs de couleurs sont identiques, il n'y a rien de plus � convertir.
        If .lBitCount = lBPPEcran Then
            bOut = b
            Erase b
            Exit Sub
        End If

        ' Second passage, conversion des couleurs.
        i = 0: j = 0
        Do While i < (.lSizeImage - 1)
            Select Case .lBitCount                      ' Lecture des pixels, extraction des couleurs.
                Case 16
                    k = b(i) * &H100 + b(i + 1)
                    With dtuRGB
                        .bRouge = (k And &H7C00& \ &H400&)
                        .bVert = (k And &H3E0&) \ &H20&
                        .bBleu = k And &H1F&
                    End With

                Case 24, 32
                    With dtuRGB
                        .bRouge = b(i + 2)
                        .bVert = b(i + 1)
                        .bBleu = b(i)
                    End With

            End Select

            Select Case lBPPEcran                       ' Ecriture des pixels.
                Case 8
                    With dtuRGB
                        ' Garder 2 bits de poids fort (\ 64)
                        k = CvCoul(.bRouge, 64) * &H10& Or _
                            CvCoul(.bVert, 64) * &H4& Or _
                            CvCoul(.bBleu, 64)
                    End With
                    bOut(j) = k

                Case 16
                    With dtuRGB
                        ' RGB 5:6:5 : 6 bits pour le vert.
                        k = (.bRouge \ 8) * &H800& Or (.bVert \ 4) * &H20& Or (.bBleu \ 8)
'                        ' RGB 5:5:5 : 5 bits par couleur
'                        k = (.bRouge \ 8) * &H800& Or (.bVert \ 8) * &H40& Or (.bBleu \ 8)
                    End With
                    bOut(j + 1) = k \ &H100&            ' Octet de poids fort.
                    bOut(j) = k And &HFF&               ' Octet de poids faible.

                Case 24, 32
                    With dtuRGB
                        bOut(j + 2) = .bRouge
                        bOut(j + 1) = .bVert
                        bOut(j) = .bBleu
                    End With

            End Select

            i = i + .lBitCount \ 8                      ' Bloc suivant.
            j = j + lBPPEcran \ 8                       ' Destination suivante.
        Loop

        ' Mise � jour de l'en-t�te avec la nouvelle profondeur de couleurs.
        .lBitCount = lBPPEcran
        .lSizeImage = UBound(bOut) + 1
    End With

    Erase b
End Sub

' Cr�er le masque AND : 1 bit par pixel
'   Si le bit de masque est 0, le pixel est dessin� normalement
'   Si le bit de masque est 1, le pixel est dessin� avec XOR sur le fond :
'       --> un pixel 0 est TRANSPARENT.
' L'image doit �tre contruite avec un fond NOIR.
Private Sub bmMasqueET(bIn() As Byte, dtuEnTete As BITMAPINFOHEADER, bOut() As Byte)
    Dim i As Byte, j As Byte, iPixel As Integer, h As Long, l As Long, bMasque As Byte

    With dtuEnTete
        ReDim bOut((.lSizeImage \ .lBitCount) - 1)
        i = 128: h = 0
        ' Un pixel peut �tre constitu� de 2 � 4 octets...
        For l = 0 To .lSizeImage - 1 Step .lBitCount \ 8
            iPixel = 0
            For j = 0 To (.lBitCount \ 8) - 1                           ' Valeur du pixel.
                iPixel = iPixel + bIn(l + j)                            ' Additionner les pixels.
            Next j
            If iPixel = 0 Then bMasque = bMasque + i                    ' Masque � 1, c'est le fond --> transparent.

            If i > 1 Then
                i = i / 2                                               ' Bit de masque pr�c�dent.
            Else
                bOut(h) = bMasque                                       ' Ecrire l'octet de masque.
                h = h + 1                                               ' Position suivante.
                bMasque = 0
                i = 128
            End If
        Next l
    End With
End Sub

' Conversion de couleur, selon la profondeur.
Private Function CvCoul(oVal As Byte, oReduc As Byte) As Byte
    CvCoul = oVal \ oReduc
End Function