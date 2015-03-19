Version = 17
VersionRequired = 17
Checksum = 97600285
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView = 0
    ScrollBars = 0
    PictureAlignment = 2
    DatasheetGridlinesBehavior = 3
    GridY = 10
    Width = 11055
    DatasheetFontHeight = 10
    ItemSuffix = 23
    Left = 12840
    Top = 720
    Right = 23655
    Bottom = 6990
    DatasheetGridlinesColor = 12632256
    RecSrcDt = Begin
        0x2f4978e58d8fe340
    End
    Caption ="Gestion de la bo�te mail"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x8905000089050000890500008905000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    PrtDevMode = Begin
        0x50444643726561746f7200000000000000000000000000000000000000000000 ,
        0x010400069c005c0353ef8001010009009a0b3408640001000f00580202000100 ,
        0x5802030001004134000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000010000000000000001000000 ,
        0x0200000001000000000000000000000000000000000000000000000050524956 ,
        0xe230000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000180000000000102710271027 ,
        0x00001027000000000000000088005c0300000000000000000100000000000000 ,
        0x0000000000000000030000000000000000001000503403002888040000000000 ,
        0x000000000000010000000000000000000000000000000000e7b14b4c03000000 ,
        0x05000a00ff000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0100000000000000000000000000000088000000534d544a0000000010007800 ,
        0x500044004600430072006500610074006f00720000005265736f6c7574696f6e ,
        0x00363030647069005061676553697a650041340050616765526567696f6e0000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x000000000000000000000000000000000000000000000000
    End
    PrtDevNames = Begin
        0x080012001e00010077696e73706f6f6c000050444643726561746f7200005044 ,
        0x4643726561746f723a000000000000000000000000000000000000
    End
    OnResize ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle = 0
        End
        Begin CommandButton
            Width = 1701
            Height = 283
            FontSize = 8
            FontWeight = 400
            ForeColor = -2147483630
            FontName ="MS Sans Serif"
        End
        Begin TextBox
            SpecialEffect = 2
            OldBorderStyle = 0
            Width = 1701
            LabelX = -1701
        End
        Begin ListBox
            SpecialEffect = 2
            Width = 1701
            Height = 1417
            LabelX = -1701
        End
        Begin ComboBox
            SpecialEffect = 2
            Width = 1701
            LabelX = -1701
        End
        Begin Subform
            SpecialEffect = 2
            Width = 1701
            Height = 1701
        End
        Begin UnboundObjectFrame
            SpecialEffect = 2
            OldBorderStyle = 1
            Width = 4536
            Height = 2835
        End
        Begin FormHeader
            Height = 1020
            BackColor = -2147483633
            Name ="Ent�teFormulaire"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect = 0
                    OverlapFlags = 85
                    TextAlign = 2
                    BackStyle = 0
                    Top = 56
                    Width = 11049
                    Height = 396
                    FontSize = 12
                    FontWeight = 700
                    Name ="Texte9"
                    ControlSource ="=Traduit(\"gbm_titre\",\"Gestionnaire de la table %s\",TableMail())"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags = 85
                    Left = 56
                    Top = 566
                    Width = 397
                    Height = 396
                    TabIndex = 1
                    Name ="cmdNouveauMSG"
                    Caption ="Commande3"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadada000adaddadadadad070dada ,
                        0xadadadada070adaddadada000070000aadadad077777770d000000000070000a ,
                        0x00fffffff070adad0f0fffff0070dada0ff0fff0f000adad0f0f000f0f0adada ,
                        0x00fffffff00dadad0fffffffff0adadaa0fffffff0adadadda0fffff0adadada ,
                        0xada00000adadadad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Nouveau message..."
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags = 85
                    Left = 509
                    Top = 566
                    Width = 397
                    Height = 396
                    TabIndex = 2
                    Name ="cmdModifMSG"
                    Caption ="Commande3"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadad00adadadaddadadad300000ada ,
                        0xadadadad3bf300ad0000000003bf370a78888888883b30007888888888830004 ,
                        0x77888777888700c47f7870007878808c7ff70fff078ff0c87f80fffff078f00c ,
                        0x780fffffff0780ad70fffffffff070da7fffffffffff00ad77777777777770da ,
                        0xadadadadadadadad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Modifier le message..."
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags = 85
                    Left = 963
                    Top = 566
                    Width = 397
                    Height = 396
                    TabIndex = 3
                    Name ="cmdSupprSEL"
                    Caption ="Commande3"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadada7adadadaddadadad17adadada ,
                        0xadadada11dada71ddadadad117da717aadadadad117d11ad0000000001111ada ,
                        0x00fffffff111adad0f0ffff711117ada0ff0ff11170117ad0f0f000f0f0a117a ,
                        0x00fffffff00da1170fffffffff0adadaa0fffffff0adadadda0fffff0adadada ,
                        0xada00000adadadad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Supprimer le(s) message(s) vers la corbeille"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    RowSourceTypeInt = 1
                    OverlapFlags = 85
                    TextAlign = 1
                    ColumnCount = 2
                    Left = 1587
                    Top = 623
                    Width = 2149
                    Height = 343
                    TabIndex = 4
                    BoundColumn = 1
                    Name ="lstDeplaceMSG"
                    RowSourceType ="Value List"
                    RowSource ="E;Bo�te d'envoi;V;El�ments envoy�s;X;Erreurs;D;Corbeille"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    Format ="&;[Blue]\"D�placer la s�lection...\""
                End
                Begin CommandButton
                    OverlapFlags = 85
                    Left = 3968
                    Top = 566
                    Width = 397
                    Height = 396
                    TabIndex = 5
                    Name ="cmdVideCorbeille"
                    Caption ="Commande3"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadad00adadadaddada00f700dadada ,
                        0xada0f8f88700adaddad08ff888870adaada0f8f822880daddad08ff828280ada ,
                        0xad08f8f878280dadda0f8ff8822870daad08f8f8888880adda0f8f8ff88880da ,
                        0xad0877777ff880addad00777788ff0daadada00778800daddadadad0000adada ,
                        0xadadadadadadadad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Vider la corbeille"
                End
                Begin CommandButton
                    OverlapFlags = 85
                    Left = 4535
                    Top = 566
                    Width = 397
                    Height = 396
                    TabIndex = 6
                    Name ="cmdActualiser"
                    Caption ="Commande3"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00da0000000000000aad0fffffffffff0dda0fffff2fffff0a ,
                        0xad0ffff22fffff0dda0fff22222fff0aad0ffff22ff2ff0dda0fffff2ff2ff0a ,
                        0xad0ff2fffff2ff0dda0ff2ff2fffff0aad0ff2ff22ffff0dda0fff22222fff0a ,
                        0xad0fffff22ffff0dda0fffff2ff0000aad0ffffffff0f0adda0ffffffff00ada ,
                        0xad0000000000adad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Actualiser"
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect = 0
                    OldBorderStyle = 1
                    OverlapFlags = 85
                    Left = 9694
                    Top = 623
                    Width = 1191
                    Height = 284
                    TabIndex = 8
                    BackColor = 0
                    ForeColor = 65535
                    Name ="Texte20"
                    ControlSource ="=Boutons([txtEtat])"
                End
                Begin CommandButton
                    FontItalic = NotDefault
                    OverlapFlags = 85
                    Left = 5102
                    Top = 566
                    Width = 397
                    Height = 396
                    FontWeight = 700
                    TabIndex = 7
                    Name ="cmdExport"
                    Caption ="eml"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial Narrow"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Exporter vers un fichier .eml"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height = 5499
            BackColor = -2147483633
            Name ="D�tail"
            Begin
                Begin Subform
                    OverlapFlags = 85
                    OldBorderStyle = 0
                    SpecialEffect = 0
                    Left = 56
                    Top = 56
                    Width = 2370
                    Height = 2265
                    Name ="sf_GestionBM_Dossiers"
                    SourceObject ="Form.sf_GestionBM_Dossiers"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                End
                Begin Subform
                    OverlapFlags = 85
                    Left = 2497
                    Top = 56
                    Width = 8475
                    Height = 2265
                    TabIndex = 1
                    Name ="sf_GestionBM_Msg"
                    SourceObject ="Form.sf_GestionBM_Msg"
                    LinkChildFields ="Etat"
                    LinkMasterFields ="txtEtat"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect = 0
                    OldBorderStyle = 1
                    OverlapFlags = 93
                    Left = 226
                    Top = 3968
                    Width = 3749
                    Height = 284
                    TabIndex = 2
                    BackColor = 0
                    ForeColor = 65535
                    Name ="txtEtat"
                    ControlSource ="=[sf_GestionBM_Dossiers].[Form].[txtEtat]"
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect = 0
                    OldBorderStyle = 1
                    OverlapFlags = 87
                    Left = 226
                    Top = 4251
                    Width = 3749
                    Height = 284
                    TabIndex = 3
                    BackColor = 0
                    ForeColor = 65535
                    Name ="txtID_MSG"
                    ControlSource ="=[sf_GestionBM_Msg].[Form].[txtID_MSG]"
                End
            End
        End
        Begin FormFooter
            Height = 0
            BackColor = -2147483633
            Name ="PiedFormulaire"
        End
    End
End
CodeBehindForm
Option Compare Database
Option Explicit


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




' Couleurs vive/att�nu�e pour le curseur des sous-formulaires.
Private Const coulNormale   As Long = vbGrayText 'vbInactiveTitleBar
Private Const coulSelect    As Long = vbHighlight ' vbActiveTitleBar
Private Const coulTexte     As Long = vbButtonText 'vbHighlightText


' Proc�dure de traduction de l'interface.
Public Sub ChangeLang()
    Static T9N_org() As String, i As Integer

    Me.lstDeplaceMSG.RowSource = "E;" & Traduit("gbm_E", "Bo�te d'envoi") & _
                                 ";V;" & Traduit("gbm_V", "El�ments envoy�s") & _
                                 ";X;" & Traduit("gbm_X", "Erreurs") & _
                                 ";D;" & Traduit("gbm_D", "Corbeille")

    i = -1
    On Error Resume Next
    i = UBound(T9N_org, 1)
    On Error GoTo 0
    If i <> -1 Then
        ' Les sous-formulaires ne sont pas accessibles par la collectin Forms.
        ' Il faut les rafra�chir manuellement en cas de changement de langue.
        Me.sf_GestionBM_Dossiers.Requery
        Call Me.sf_GestionBM_Dossiers.Form.ChangeLang
        Me.sf_GestionBM_Msg.Form.Requery
        Call Me.sf_GestionBM_Msg.Form.ChangeLang
    End If

    Call LangueCtls(Me.Form, T9N_org())

    Me.Texte9.Requery

End Sub



' Actualise les sous-formulaires
' M�thode - garder Public
Public Sub Actualiser()
    Dim lT0 As Long, lT1 As Long, lH1 As Long

    ' Lors du Requery, la fonction Boutons est d�clench�e.
    ' Il faut sauvegarder les s�lections AVANT les Requery.
    With Me.sf_GestionBM_Dossiers.Form
        lT0 = .SelTop                                   ' Sauvegarder la s�lection
    End With
    With Me.sf_GestionBM_Msg.Form
        lT1 = .SelTop                                   ' Sauvegarder la s�lection
        lH1 = .pSelHeight
    End With

    On Error Resume Next
    With Me.sf_GestionBM_Dossiers.Form
        .Painting = False
        Me.sf_GestionBM_Msg.Form.Painting = False
        .Requery
        .SelTop = lT0                                   ' Restaurer la s�lection
        .Painting = True
    End With
    With Me.sf_GestionBM_Msg.Form
'        .Requery -- pas n�cessaire, car d�clench� automatiquement par le premier Requery
        .SelTop = lT1                                   ' Restaurer la s�lection
        .pSelHeight = lH1
        .Painting = True
    End With
    On Error GoTo 0
End Sub
'
' ---------------------------------------------------------


' Activation des boutons en fonction du contexte.
Private Function Boutons(sEtat As String) As String
    Dim NbEnr As Long

    ' Nombre d'enregistrements du dossier
    NbEnr = DCount("Identifiant", TableMail(), "Etat='" & Me.txtEtat & "'")

    Me.cmdActualiser.SetFocus                           ' Avant de d�sactiver un autre contr�le...

    ' Le bouton Modifier n'est disponible que dans la bo�te d'envoi, si aucun envoi n'est en cours.
    Me.cmdModifMSG.Enabled = ((sEtat = "E") And _
                              (SMTPEtatSrv.Etat <> lmlSrvEnCours) And _
                              (NbEnr > 0))

    ' La liste D�placer, les boutons Supprimer et Exporter ne sont actifs que s'il y a au moins un message
    Me.cmdSupprSEL.Enabled = (NbEnr > 0)
    Me.lstDeplaceMSG.Enabled = (NbEnr > 0)
    Me.cmdExport.Enabled = (NbEnr > 0)

    ' R�initialiser la s�lection lors du changement de dossier,
    ' seulement lorsqu'on clique sur un autre dossier.
    ' Actualiser provoque �galement un appel � la fonction Boutons(), mais dans ce cas,
    ' on conserve la s�lection.
    On Error Resume Next
    If Screen.ActiveControl.Name = "txtCurseur0" Then
        With Me.sf_GestionBM_Msg.Form
            .SelTop = 1
            .pSelHeight = 1
        End With
    End If
    On Error GoTo 0
End Function

' Retourne la liste des identifiants des messages s�lectionn�s.
Private Function IxMsg() As String
    Dim rs As DAO.Recordset, l As Long, SQL As String, lSelDeb As Long, lSelFin As Long

    Set rs = Me.sf_GestionBM_Msg.Form.RecordsetClone

    lSelDeb = Me.sf_GestionBM_Msg.Form.SelTop
    lSelFin = Me.sf_GestionBM_Msg.Form.SelTop + Me.sf_GestionBM_Msg.Form.pSelHeight - 1

    If rs.RecordCount > 0 Then rs.MoveFirst
    Do While Not rs.EOF And l < lSelFin
        l = rs.AbsolutePosition + 1
        ' Vrai si l'enregistrement en cours est s�lectionn�.
        If (l >= lSelDeb And l <= lSelFin) Then
            If Len(SQL) <> 0 Then SQL = SQL & ","
            SQL = SQL & """" & rs!Identifiant & """"
        End If

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    IxMsg = SQL
End Function

' V�rifications avant traitement.
Private Function TrtPossible() As Boolean
    If Me.sf_GestionBM_Msg.Form.pSelHeight = 0 Then
        MsgBox Traduit("gbm_nosel", "La s�lection ne comporte aucune ligne."), vbInformation, "libMAIL"
        Exit Function
    End If

    If SMTPEtatSrv.Etat = lmlSrvEnCours And Me.txtEtat = "E" Then
        MsgBox Traduit("gbm_impossible", "Cette action n'est pas possible pendant que le serveur traite les messages de la boite d'envoi."), vbCritical, "libMAIL"
        Exit Function
    End If

    TrtPossible = True
End Function
'
' ----------------------------------------------------------


Private Sub cmdActualiser_Click()
    Call Actualiser
End Sub

Private Sub cmdExport_Click()
    Dim dtuOFN As OPENFILENAME, sFiltre As String, l As Long, sID As String

    ' R�cup�rer l'ID de message, m�me si le contr�le n'existe pas.
    On Error Resume Next
    sID = Nz(Me.txtID_MSG, "")
    On Error GoTo 0

    If Len(sID) = 0 Then
        MsgBox Traduit("gbm_noseleml", "Vous devez s�lectionner un message � exporter."), vbInformation
        Exit Sub
    End If

    sFiltre = Traduit("gbm_emlfilter01", "Fichiers eml") & vbNullChar & "*.eml" & vbNullChar & _
              Traduit("gbm_emlfiltre02", "Tous les fichiers") & vbNullChar & "*.*"

    With dtuOFN
        .lStructSize = Len(dtuOFN)
        .hWndOwner = Me.hwnd
        .sFilter = sFiltre
        .lFilterIndex = 1
        .sFile = sID & ".eml" & vbNullChar & Space$(16000)
        .lMaxFile = Len(.sFile)
        .sFileTitle = vbNullChar & Space$(512)
        .lMaxFileTitle = Len(.sFileTitle)
'        .sInitialDir = "c:\temp" & vbNullChar & Space$(512) & vbNullChar & vbNullChar
        .sDialogTitle = Traduit("gbm_emlexport", "Exporter le message '%s'.", sID)
        .lFlags = OFN_ENABLESIZING Or OFN_PATHMUSTEXIST Or OFN_EXPLORER Or OFN_OVERWRITEPROMPT
        .sDefFileExt = "*.eml"
    End With

    l = GetSaveFileName(dtuOFN)
    If l = 0 Then Exit Sub

    l = ExporteEML(sID, dtuOFN.sFile)

    If l = 0 Then
        MsgBox Traduit("gbm_emlexpsucces", "Le message a �t� enregistr� dans %s.", dtuOFN.sFile), vbInformation
    Else
        MsgBox Traduit("gbm_emlexporterreur", "Erreur %s, %s lors de l'enregistrement du message dans le fichier %s.", l, Error$(l), dtuOFN.sFile)
    End If
End Sub

Private Sub cmdModifMSG_Click()
    Call ModifieMail(Me.txtID_MSG)
    Call Actualiser
End Sub

Private Sub cmdNouveauMSG_Click()
    Call mnuCreeMail                                        ' Appeler le formulaire d'�dition de mail

    ' Attendre la fermeture du formulaire.
    Do While FrmEstCharge("frm_EditeMail")
        Call myDoEvents
    Loop

    If FrmEstCharge("frm_EditeMail") Then Call Actualiser
End Sub

Private Sub cmdSupprSEL_Click()
    Dim s As String

    If Not TrtPossible() Then Exit Sub

    If Me.txtEtat = "D" Then s = Traduit("gbm_supprdel", "supprimer") Else s = Traduit("gbm_supprtrash", "envoyer vers la corbeille")

    If MsgBox(Traduit("gbm_supprconfirm", "Etes-vous s�r(e) de vouloir %s le(s) message(s) s�lectionn�(s) ?", s), _
              vbYesNo + vbQuestion + vbDefaultButton2, "libMAIL") = vbYes Then
        If Me.txtEtat = "D" Then
            ' Suppression depuis la corbeille
            CurrentDb.Execute "DELETE * FROM " & TableMail() & " WHERE Etat='D' AND Identifiant In (" & IxMsg() & ")"
        Else
            ' Envoyer vers la corbeille
            CurrentDb.Execute "UPDATE " & TableMail() & " SET Etat='D' WHERE Identifiant In (" & IxMsg() & ")"
        End If
    End If

    Call Actualiser
End Sub

Private Sub cmdVideCorbeillle_Click()
    If MsgBox(Traduit("gbm_videconfirm", "Etes-vous s�r(e) de vouloir vider la corbeille ?"), _
              vbYesNo + vbQuestion + vbDefaultButton2, "libMAIL") = vbYes Then
        CurrentDb.Execute "DELETE * FROM " & TableMail() & " WHERE Etat='D'"

        Call Actualiser
    End If
End Sub

Private Sub Form_Load()
    Call Me.ChangeLang

    With Me.sf_GestionBM_Dossiers.Form
        .txtFond.ForeColor = coulNormale
        .txtLibelle.ForeColor = coulTexte
    End With
    With Me.sf_GestionBM_Msg.Form
        .txtFond.ForeColor = coulNormale
        .txtDateMsg.ForeColor = coulTexte
        .txtDestinataires.ForeColor = coulTexte
        .txtObjet.ForeColor = coulTexte
    End With
End Sub

Private Sub Form_Resize()
    Dim l As Long

    Me.Painting = False

    ' Calculer les hauteurs.
    l = Me.InsideHeight - Me.Ent�teFormulaire.Height - Me.PiedFormulaire.Height - 2 * Me.sf_GestionBM_Dossiers.Top
    If l < 240 Then l = 240
    Me.D�tail.Height = l + 2 * Me.sf_GestionBM_Dossiers.Top
    Me.sf_GestionBM_Dossiers.Height = l
    Me.sf_GestionBM_Msg.Height = l

    ' Largeur du sous-formulaire des messages
    l = Me.InsideWidth - Me.sf_GestionBM_Msg.Left - Me.sf_GestionBM_Dossiers.Left
    If l < 240 Then l = 240
    Me.sf_GestionBM_Msg.Width = l

    Me.Texte9.Width = Me.InsideWidth

    ' Lorsque le formulaire est maximis�, Access ne d�clenche pas le Resize du sous-formulaire...
    On Error Resume Next
    Call Me.sf_GestionBM_Msg.Form.myResize
    On Error GoTo 0

    Me.Painting = True
End Sub

Private Sub lstDeplaceMSG_AfterUpdate()
    If Not TrtPossible() Then Exit Sub

    If Me.lstDeplaceMSG.Column(0) = Me.txtEtat Then
        MsgBox Traduit("gbm_deplaceidem", "Les dossiers source et destination sont identiques !"), vbCritical, "libMAIL"
        Exit Sub
    End If

    If MsgBox(Traduit("gbm_deplaceconfirm", "Etes-vous s�r(e) de vouloir d�placer le(s) message(s) s�lectionn�(s) de '%s' vers '%s' ?", Me.sf_GestionBM_Dossiers.Form.Libelle, Me.lstDeplaceMSG), _
              vbYesNo + vbQuestion + vbDefaultButton2, "libMAIL") = vbYes Then
        CurrentDb.Execute "UPDATE " & TableMail() & " SET Etat='" & Me.lstDeplaceMSG.Column(0) & "' WHERE Identifiant In (" & IxMsg() & ")"
    End If

    Call Actualiser

    Me.lstDeplaceMSG = Null
End Sub

Private Sub sf_GestionBM_Dossiers_Enter()
    Me.sf_GestionBM_Dossiers.Form.txtFond.ForeColor = coulSelect
End Sub

Private Sub sf_GestionBM_Dossiers_Exit(Cancel As Integer)
    Me.sf_GestionBM_Dossiers.Form.txtFond.ForeColor = coulNormale
End Sub

Private Sub sf_GestionBM_Msg_Enter()
    Me.sf_GestionBM_Msg.Form.txtFond.ForeColor = coulSelect
End Sub

Private Sub sf_GestionBM_Msg_Exit(Cancel As Integer)
    Me.sf_GestionBM_Msg.Form.txtFond.ForeColor = coulNormale

    ' Re-forcer la s�lection � la sortie du formulaire.
    Me.sf_GestionBM_Msg.Form.pSelHeight = Me.sf_GestionBM_Msg.Form.pSelHeight
End Sub