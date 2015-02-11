Version = 17
VersionRequired = 17
Checksum = -1768622923
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView = 0
    ScrollBars = 0
    BorderStyle = 1
    PictureAlignment = 2
    DatasheetGridlinesBehavior = 3
    GridY = 10
    Width = 9150
    DatasheetFontHeight = 10
    ItemSuffix = 53
    Left = 1650
    Top = 2805
    Right = 11775
    Bottom = 9525
    DatasheetGridlinesColor = 12632256
    PaintPalette = Begin
        0x0003100000000000800000000080000080800000000080008000800000808000 ,
        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
        0xffffff0000000000
    End
    RecSrcDt = Begin
        0xa5111f9015e2e340
    End
    Caption ="Démo libMAIL"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle = 0
        End
        Begin Image
            BackStyle = 0
            OldBorderStyle = 0
            PictureAlignment = 2
            Width = 1701
            Height = 1701
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
        Begin UnboundObjectFrame
            SpecialEffect = 2
            OldBorderStyle = 1
            Width = 4536
            Height = 2835
        End
        Begin Tab
            Width = 5103
            Height = 3402
        End
        Begin Page
            Width = 1701
            Height = 1701
        End
        Begin Section
            Height = 5782
            BackColor = -2147483633
            Name ="Détail"
            Begin
                Begin Tab
                    OverlapFlags = 93
                    Left = 30
                    Top = 30
                    Width = 9120
                    Height = 5355
                    Name ="tabDEMO"
                    FontName ="Arial"
                    Begin
                        Begin Page
                            OverlapFlags = 215
                            Left = 165
                            Top = 450
                            Width = 8850
                            Height = 4800
                            Name ="pgBienvenue"
                            PictureData = Begin
                                0x2800000010000000100000000100040000000000800000000000000000000000 ,
                                0x0000000000000000000000000000800000800000008080008000000080008000 ,
                                0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                                0xffff0000ffffff00dadadadadadadadaadada00000adadaddad00bbbbb00dada ,
                                0xad0bbbbbbbbb0dadd0bbbbbbbbbbb0daa0bbb00000bbb0ad0bbb0bbbbb0bbb0a ,
                                0x0bb0bbbbbbb0bb0d0bbbbbbbbbbbbb0a0bbbbbbbbbbbbb0d0bbb00bbb00bbb0a ,
                                0xa0bb00bbb00bb0add0bbbbbbbbbbb0daad0bbbbbbbbb0daddad00bbbbb00dada ,
                                0xadada00000adadad
                            End
                            Caption ="Bienvenue"
                            Begin
                                Begin Label
                                    OverlapFlags = 215
                                    Left = 170
                                    Top = 570
                                    Width = 8565
                                    Height = 1815
                                    Name ="lblBienvenue"
                                    Caption ="Bienvenue dans l'application de démonstation de libMAIL.\015\012\015\012Cette pe"
                                        "tite application a pour vocation de vous faire découvrir en quelques étapes simp"
                                        "les le fonctionnement de la bibliothèque libMAIL.\015\012\015\012Cliquez sur le "
                                        "bouton Suivant lorsque vous être prêt(e) à commencer."
                                    FontName ="Arial"
                                End
                            End
                        End
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags = 215
                            Left = 165
                            Top = 450
                            Width = 8850
                            Height = 4800
                            Name ="pgBiblio"
                            PictureData = Begin
                                0x2800000010000000100000000100040000000000800000000000000000000000 ,
                                0x0000000000000000000000000000800000800000008080008000000080008000 ,
                                0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                                0xffff0000ffffff00dadadadadadadada000000000000adad0ffffffffff0dada ,
                                0x0f00f00000f0adad0ffffffffff0dada0f00f00000f0adad0ffffffffff0dada ,
                                0x0fffffff0ff0adad0f00fff080f0dad40f080f0808000d440ff0808080888044 ,
                                0x0000080808888844dadad08088888844adadad0888888044dadadad000000a44 ,
                                0xadadadadadadadad
                            End
                            Caption ="Préparation"
                            Begin
                                Begin Label
                                    OverlapFlags = 215
                                    Left = 170
                                    Top = 570
                                    Width = 8674
                                    Height = 737
                                    Name ="lblBiblio00"
                                    Caption ="Pour pouvoir utiliser la bibliothèque à partir de votre application, vous devez "
                                        "la lier à cette dernière.\015\012\015\012Cette opération s'effectue normalement "
                                        "à l'aide de la boite de dialogue 'Références'."
                                    FontName ="Arial"
                                End
                                Begin CommandButton
                                    OverlapFlags = 215
                                    Left = 7370
                                    Top = 1363
                                    Width = 1474
                                    Height = 397
                                    Name ="cmdReferences"
                                    Caption ="Parcourir..."
                                    OnClick ="[Event Procedure]"
                                    FontName ="Arial"
                                End
                                Begin Label
                                    OverlapFlags = 215
                                    Left = 170
                                    Top = 1364
                                    Width = 7088
                                    Height = 564
                                    Name ="lblBiblio01"
                                    Caption ="..."
                                    FontName ="Arial"
                                End
                                Begin Label
                                    OverlapFlags = 215
                                    Left = 170
                                    Top = 2157
                                    Width = 8674
                                    Height = 686
                                    Name ="lblBiblio02"
                                    Caption ="libMAIL a également besoin d'une table pour y stocker les messages.\015\012Cette"
                                        " table sera créée automatiquement pour vous lorsque vous appellerez pour la prem"
                                        "ière fois une fonction de création de message ou de démarrage du serveur."
                                    FontName ="Arial"
                                End
                                Begin Label
                                    OverlapFlags = 215
                                    Left = 170
                                    Top = 2923
                                    Width = 7090
                                    Height = 590
                                    Name ="lblBiblio03"
                                    Caption ="..."
                                    FontName ="Arial"
                                End
                                Begin CommandButton
                                    OverlapFlags = 215
                                    Left = 7370
                                    Top = 2951
                                    Width = 1474
                                    Height = 397
                                    TabIndex = 1
                                    Name ="cmdBoiteMail"
                                    Caption ="Créer la table"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Arial"
                                End
                                Begin Label
                                    OverlapFlags = 215
                                    Left = 170
                                    Top = 4595
                                    Width = 8674
                                    Height = 621
                                    FontWeight = 700
                                    Name ="lblBiblio04"
                                    Caption ="Ces deux opération ne seront à réaliser qu'une seule fois pour votre application"
                                        ".\015\012\015\012Vous pouvez maintenant passer à l'étape suivante : la rédaction"
                                        " du message."
                                    FontName ="Arial"
                                End
                            End
                        End
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags = 215
                            Left = 165
                            Top = 450
                            Width = 8850
                            Height = 4800
                            Name ="pgMSG"
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
                            Caption ="Message"
                            Begin
                                Begin Label
                                    OverlapFlags = 215
                                    Left = 170
                                    Top = 570
                                    Width = 8621
                                    Height = 340
                                    Name ="lblMSG00"
                                    Caption ="Nous avons maintenant tout le nécessaire pour créer notre premier message."
                                    FontName ="Arial"
                                End
                                Begin TextBox
                                    OverlapFlags = 215
                                    TextAlign = 2
                                    Left = 4081
                                    Top = 1023
                                    Width = 3402
                                    Height = 283
                                    Name ="txtEmetteur"
                                    Format ="@;[Blue]\"mon.adresse@monfai.fr\""
                                    OnExit ="[Event Procedure]"
                                    FontName ="Arial"
                                    ControlTipText ="Adresse de l'expéditeur"
                                    Begin
                                        Begin Label
                                            OverlapFlags = 215
                                            Left = 396
                                            Top = 1020
                                            Width = 3636
                                            Height = 276
                                            Name ="Étiquette19"
                                            Caption ="Saisissez votre adresse de courriel ici :"
                                            FontName ="Arial"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags = 215
                                    TextAlign = 2
                                    Left = 4081
                                    Top = 1363
                                    Width = 3402
                                    Height = 284
                                    TabIndex = 1
                                    Name ="txtDest"
                                    Format ="@;[Blue]\"destinataire@domaine.fr\""
                                    OnExit ="[Event Procedure]"
                                    FontName ="Arial"
                                    ControlTipText ="Adresse du destinataire"
                                    Begin
                                        Begin Label
                                            OverlapFlags = 215
                                            Left = 400
                                            Top = 1363
                                            Width = 3636
                                            Height = 276
                                            Name ="Étiquette23"
                                            Caption ="Saisissez l'adresse de courriel du destinataire ici :"
                                            FontName ="Arial"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags = 215
                                    Left = 2157
                                    Top = 1817
                                    Width = 6640
                                    Height = 284
                                    TabIndex = 2
                                    Name ="txtObjet"
                                    FontName ="Arial"
                                    Begin
                                        Begin Label
                                            OverlapFlags = 223
                                            Left = 396
                                            Top = 1817
                                            Width = 1764
                                            Height = 228
                                            Name ="Étiquette29"
                                            Caption ="Objet du message :"
                                            FontName ="Arial"
                                        End
                                    End
                                End
                                Begin TextBox
                                    EnterKeyBehavior = NotDefault
                                    ScrollBars = 2
                                    OverlapFlags = 215
                                    Left = 396
                                    Top = 2383
                                    Width = 8394
                                    Height = 1696
                                    TabIndex = 3
                                    Name ="txtMSG"
                                    FontName ="Arial"
                                    Begin
                                        Begin Label
                                            OverlapFlags = 223
                                            Left = 396
                                            Top = 2157
                                            Width = 3636
                                            Height = 228
                                            Name ="Étiquette25"
                                            Caption ="Tapez votre message ci-dessous :"
                                            FontName ="Arial"
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags = 215
                                    Left = 396
                                    Top = 4141
                                    Width = 6633
                                    Height = 285
                                    Name ="lblMSG01"
                                    Caption ="Lorsque votre message est terminé, cliquez sur le bouton Enregistrer ci-contre."
                                    FontName ="Arial"
                                End
                                Begin CommandButton
                                    OverlapFlags = 215
                                    Left = 7143
                                    Top = 4141
                                    Width = 1637
                                    Height = 340
                                    TabIndex = 4
                                    Name ="cmdEnregistrer"
                                    Caption ="Enregistrer"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Arial"
                                End
                                Begin Label
                                    OverlapFlags = 215
                                    Left = 170
                                    Top = 4595
                                    Width = 8616
                                    Height = 624
                                    FontWeight = 700
                                    Name ="lblMSG02"
                                    Caption ="Le message est maintenant mémorisé dans la table BoiteMail, mais il n'a pas enco"
                                        "re été expédié.\015\012\015\012Vous pouvez maintenant passer à l'étape suivante "
                                        ": l'envoi du message."
                                    FontName ="Arial"
                                End
                            End
                        End
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags = 247
                            Left = 165
                            Top = 450
                            Width = 8850
                            Height = 4800
                            Name ="pgSRV"
                            PictureData = Begin
                                0x2800000010000000100000000100040000000000800000000000000000000000 ,
                                0x0000000000000000000000000000800000800000008080008000000080008000 ,
                                0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                                0xffff0000ffffff00dadadadadadadadaadada00000adadaddad00ee27e00dada ,
                                0xad0eee2727ee0dadd0eeee7272eee0daa0eeee27272ee0ad0eeeee72727eee0a ,
                                0x0ee727272727ee0d0ee272727272e20a0ee727272727270d0ee2727e72e2720a ,
                                0xa0ee27eeeee720add0e27272727270daad0ee72727270daddad00e727200dada ,
                                0xadada00000adadad
                            End
                            Caption ="Serveur SMTP distant"
                            Begin
                                Begin Label
                                    OverlapFlags = 247
                                    Left = 170
                                    Top = 570
                                    Width = 8617
                                    Height = 340
                                    Name ="lblSRV00"
                                    Caption ="Il faut maintenant indiquer à libMAIL à quel serveur SMTP il doit se connecter p"
                                        "our pouvoir envoyer le message."
                                    FontName ="Arial"
                                End
                                Begin TextBox
                                    OverlapFlags = 247
                                    TextAlign = 2
                                    Left = 3798
                                    Top = 967
                                    Width = 2211
                                    Height = 283
                                    Name ="txtSrv"
                                    Format ="@;[Blue]\"smtp.fai.fr\""
                                    OnExit ="[Event Procedure]"
                                    FontName ="Arial"
                                    Begin
                                        Begin Label
                                            OverlapFlags = 247
                                            Left = 963
                                            Top = 967
                                            Width = 2722
                                            Height = 283
                                            Name ="Étiquette33"
                                            Caption ="Nom ou adresse IP du serveur SMTP :"
                                            FontName ="Arial"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags = 247
                                    TextAlign = 2
                                    Left = 3798
                                    Top = 1307
                                    Width = 850
                                    Height = 283
                                    TabIndex = 1
                                    Name ="txtPort"
                                    ValidationRule ="Between 0 And 65535"
                                    ValidationText ="Le port doit être compris entre 0 et 65635"
                                    DefaultValue ="25"
                                    FontName ="Arial"
                                    Begin
                                        Begin Label
                                            OverlapFlags = 247
                                            Left = 963
                                            Top = 1307
                                            Width = 2722
                                            Height = 283
                                            Name ="Étiquette35"
                                            Caption ="Port (25 par défaut) :"
                                            FontName ="Arial"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags = 247
                                    Left = 7200
                                    Top = 1874
                                    Width = 1587
                                    Height = 340
                                    TabIndex = 2
                                    Name ="cmdEnvoyer"
                                    Caption ="Envoyer"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Arial"
                                End
                                Begin Label
                                    OverlapFlags = 247
                                    Left = 170
                                    Top = 1827
                                    Width = 6987
                                    Height = 495
                                    Name ="lblSrv01"
                                    Caption ="Cliquez sur Envoyer pour démarrer le serveur SMTP de libMAIL et déclencher l'exp"
                                        "édition du  message."
                                    FontName ="Arial"
                                End
                                Begin Label
                                    FontItalic = NotDefault
                                    OverlapFlags = 247
                                    Left = 170
                                    Top = 2497
                                    Width = 8617
                                    Height = 2094
                                    Name ="lblSRV02"
                                    Caption ="Le message devrait arriver dans quelques instants dans la messagerie du destinai"
                                        "re, ou dans votre propre boite aux lettres si vous vous êtes mis en destinataire"
                                        ".\015\012\015\012Si le message n'arrive pas, cliquez sur le bouton 'Suivant', pu"
                                        "is sur le bouton Journal.\015\012Parcourez le journal à la recherche d'une erreu"
                                        "r. Vous avez peut-être mal orthographié le nom du serveur SMTP.\015\012\015\012V"
                                        "otre serveur SMTP nécessite peut-être une authentification pour accepter le mess"
                                        "age. Dans ce cas, vous pouvez refaire un essai à partir du formulaire de créatio"
                                        "n et d'envoi de message, accessible sur la page suivante.\015\012Cliquez sur le "
                                        "bouton 'Suivant', puis sur le bouton 'Créer un mail'."
                                    FontName ="Arial"
                                End
                            End
                        End
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags = 247
                            Left = 165
                            Top = 450
                            Width = 8850
                            Height = 4800
                            Name ="pgSuite"
                            PictureData = Begin
                                0x2800000010000000100000000100040000000000800000000000000000000000 ,
                                0x0000000000000000000000000000800000800000008080008000000080008000 ,
                                0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                                0xffff0000ffffff00dadadadadadadadaadadada000adadaddadad005570adada ,
                                0xada0055ff870adadd0055ff008870ada755ff005508870ad75f005555508870a ,
                                0x7005553b55508870755555535555080aa5f55555bb55500dda5f555553bb550a ,
                                0xada5f55535bb5550dada5f55bbb55500adada5f5555500addadada5f5500dada ,
                                0xadadada550adadad
                            End
                            Caption ="Pour aller plus loin..."
                            Begin
                                Begin Label
                                    OverlapFlags = 247
                                    Left = 170
                                    Top = 456
                                    Width = 8674
                                    Height = 685
                                    Name ="lblSuite00"
                                    Caption ="Cette démonstration rapide vous a permis de prendre connaissance des fonctionnal"
                                        "ités de base de libMAIL.\015\012\015\012La bibliothèque possède bien d'autres op"
                                        "tions que vous pourrez découvrir dans le manuel."
                                    FontName ="Arial"
                                End
                                Begin CommandButton
                                    OverlapFlags = 247
                                    Left = 170
                                    Top = 1307
                                    Width = 1304
                                    Height = 397
                                    Name ="cmdJournal"
                                    Caption ="Journal"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Arial"
                                End
                                Begin CommandButton
                                    OverlapFlags = 247
                                    Left = 170
                                    Top = 1784
                                    Width = 1304
                                    Height = 397
                                    TabIndex = 1
                                    Name ="cmdGestMail"
                                    Caption ="Gestionnaire"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Arial"
                                End
                                Begin CommandButton
                                    OverlapFlags = 247
                                    Left = 170
                                    Top = 2261
                                    Width = 1304
                                    Height = 397
                                    TabIndex = 2
                                    Name ="cmdCreeMail"
                                    Caption ="Créer un mail"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Arial"
                                End
                                Begin Label
                                    OverlapFlags = 247
                                    Left = 1530
                                    Top = 1363
                                    Width = 7144
                                    Height = 341
                                    Name ="lblSuite01"
                                    Caption ="Affiche le journal de connexion de libMAIL."
                                    FontName ="Arial"
                                End
                                Begin Label
                                    OverlapFlags = 247
                                    Left = 1530
                                    Top = 1760
                                    Width = 7140
                                    Height = 456
                                    Name ="lblSuite02"
                                    Caption ="Affiche le Gestionnaire de la table BoiteMail. Il vous permettra de déplacer, su"
                                        "pprimer ou modifier des messages."
                                    FontName ="Arial"
                                End
                                Begin Label
                                    OverlapFlags = 247
                                    Left = 1530
                                    Top = 2271
                                    Width = 7144
                                    Height = 396
                                    Name ="lblSuite03"
                                    Caption ="A l'aide de ce formulaire, vous pourrez créer et envoyer un message sans écrire "
                                        "de code."
                                    FontName ="Arial"
                                End
                                Begin Label
                                    OverlapFlags = 255
                                    Left = 165
                                    Top = 2835
                                    Width = 8505
                                    Height = 1275
                                    Name ="lblSuiet04"
                                    Caption ="Vous pouvez également faire un clic droit sur l'icône affichée dans la zone de n"
                                        "otifications de Windows pour accéder à d'autres options.\015\012\015\012libMAIL "
                                        "est une biblitohèque. Toutes les fonctionnalités évoquées ici sont accessibles à"
                                        " l'aide de procédures ou de fonctions, afin que vous puissiez envoyer facilement"
                                        " des messages depuis n'importe quelle application dévelopée sous MS-Access."
                                    FontName ="Arial"
                                End
                                Begin CommandButton
                                    OverlapFlags = 247
                                    Left = 168
                                    Top = 4172
                                    Width = 1304
                                    Height = 517
                                    TabIndex = 3
                                    Name ="cmdSTOP"
                                    Caption ="Arrêt du serveur"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Arial"
                                End
                                Begin Label
                                    OverlapFlags = 247
                                    Left = 1530
                                    Top = 4198
                                    Width = 7152
                                    Height = 456
                                    Name ="Étiquette51"
                                    Caption ="Un clic sur ce bouton terminera le serveur SMTP de libMAIL et retirera l'icône d"
                                        "e la zone de notifications."
                                    FontName ="Arial"
                                End
                                Begin Image
                                    SizeMode = 1
                                    Left = 8447
                                    Top = 2837
                                    Width = 426
                                    Height = 426
                                    Name ="imgAttente"
                                    PictureData = Begin
                                        0x2800000010000000100000000100200000000000000400000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x000000000000ff000000ff000000ff000000ff000000ff000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000ff0000ffff0000ffff008000000000ffff0000ffff000000ff0000000000 ,
                                        0x000000000000000000000000000000000000000000000000000000000000ff00 ,
                                        0x00ffff0000ffff0000ffff008000000000ffff0000ffff0000ffff000000ff00 ,
                                        0x0000000000000000000000000000000000000000000000000000ff0000ffff00 ,
                                        0x00ffff0000ffff0000ffff0000ffff0000ffff0000ffff008000000000ffff00 ,
                                        0x0000ff0000000000000000000000000000000000000000000000ff0000ffff00 ,
                                        0x00ffff0000ffff0000ffff0000ffff0000ffff008000000000ffff0000ffff00 ,
                                        0x0000ff00000000000000000000000000000000000000ff0000ffff0000ffff00 ,
                                        0x00ffff0000ffff0000ffff0000ffff008000000000ffff0000ffff0000ffff00 ,
                                        0x00ffff000000ff000000000080000000800000000000ff008000000080000000 ,
                                        0x00ffff0000ffff0000ffff008000000000ffff0000ffff0000ffff0080000000 ,
                                        0x800000000000ff008000000080000000ffff00000000ff0000ffff0000ffff00 ,
                                        0x00ffff0000ffff0000ffff008000000000ffff0000ffff0000ffff0000ffff00 ,
                                        0x00ffff000000ff0080000000ffff000080000000ffff00000000ff0000ffff00 ,
                                        0x00ffff0000ffff0000ffff008000000000ffff0000ffff0000ffff0000ffff00 ,
                                        0x0000ff000000000080000000ffff0000ffff0000800000000000ff0000ffff00 ,
                                        0x00ffff0000ffff0000ffff008000000000ffff0000ffff0000ffff0000ffff00 ,
                                        0x0000ff000000000080000000ffff0000ffff0000ffff0000800000000000ff00 ,
                                        0x00ffff0000ffff0000ffff008000000000ffff0000ffff0000ffff000000ff00 ,
                                        0x800000000000000080000000ffff0000ffff0000ffff0000ffff000080000000 ,
                                        0x0000ff0000ffff0000ffff0000ffff0000ffff0000ffff000000ff00ffff0000 ,
                                        0x800000000000000080000000ffff0000ffff0000ffff000080000000ffff0000 ,
                                        0xffff00000000ff000000ff000000ff000000ff000000ff00ffff0000ffff0000 ,
                                        0x800000000000000080000000ffff0000ffff000080000000ffff0000ffff0000 ,
                                        0xffff0000ffff0000ffff0000ffff0000ffff000080000000ffff0000ffff0000 ,
                                        0x8000000000000000800000008000000080000000ffff0000ffff0000ffff0000 ,
                                        0xffff0000ffff0000ffff0000ffff0000ffff0000ffff00008000000080000000 ,
                                        0x8000000000000000000000008000000080000000800000008000000080000000 ,
                                        0x8000000080000000800000008000000080000000800000008000000080000000 ,
                                        0x0000000000000000
                                    End
                                    Picture ="C:\\Documents and Settings\\Denis\\Bureau\\TEST\\libMAIL\\lmlAttente.ico"
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags = 215
                    AccessKey = 80
                    Left = 6066
                    Top = 5385
                    Width = 1474
                    Height = 397
                    TabIndex = 1
                    Name ="cmdPrec"
                    Caption ="&Précédent"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags = 215
                    AccessKey = 83
                    Left = 7540
                    Top = 5385
                    Width = 1474
                    Height = 397
                    TabIndex = 2
                    Name ="cmdSvt"
                    Caption ="&Suivant"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                End
            End
        End
    End
End
CodeBehindForm
Option Compare Database
Option Explicit

' Copyright 2009, 2010, 2011 Denis SCHEIDT
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




Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
' Structure de données pour la boîte de dialogue 'Ouvrir'
Private Type OPENFILENAME
    lStructSize             As Long
    hWndOwner               As Long
    hInstance               As Long
    sFilter                 As String
    sCustomFilter           As String
    lMaxCustFilter          As Long
    lFilterIndex            As Long
    sFile                   As String
    lMaxFile                As Long
    sFileTitle              As String
    lMaxFileTitle           As Long
    sInitialDir             As String
    sDialogTitle            As String
    lFlags                  As Long
    iFileOffset             As Integer
    iFileExtension          As Integer
    sDefFileExt             As String
    lCustData               As Long
    lfnHook                 As Long
    sTemplateName           As String
    pvReserved              As Long
    dwReserved              As Long
    lFlagsEx                As Long
End Type
' Constantes OpenFileName
Private Const OFN_ALLOWMULTISELECT       As Long = &H200
Private Const OFN_CREATEPROMPT           As Long = &H2000
Private Const OFN_ENABLEHOOK             As Long = &H20
Private Const OFN_ENABLETEMPLATE         As Long = &H40
Private Const OFN_ENABLETEMPLATEHANDLE   As Long = &H80
Private Const OFN_EXPLORER               As Long = &H80000
Private Const OFN_EXTENSIONDIFFERENT     As Long = &H400
Private Const OFN_FILEMUSTEXIST          As Long = &H1000
Private Const OFN_HIDEREADONLY           As Long = &H4
Private Const OFN_LONGNAMES              As Long = &H200000
Private Const OFN_NOCHANGEDIR            As Long = &H8
Private Const OFN_NODEREFERENCELINKS     As Long = &H100000
Private Const OFN_NOLONGNAMES            As Long = &H40000
Private Const OFN_NONETWORKBUTTON        As Long = &H20000
Private Const OFN_NOREADONLYRETURN       As Long = &H8000&
Private Const OFN_NOTESTFILECREATE       As Long = &H10000
Private Const OFN_NOVALIDATE             As Long = &H100
Private Const OFN_OVERWRITEPROMPT        As Long = &H2
Private Const OFN_PATHMUSTEXIST          As Long = &H800
Private Const OFN_READONLY               As Long = &H1
Private Const OFN_SHAREAWARE             As Long = &H4000
Private Const OFN_SHAREFALLTHROUGH       As Long = 2
Private Const OFN_SHAREWARN              As Long = 0
Private Const OFN_SHARENOWARN            As Long = 1
Private Const OFN_SHOWHELP               As Long = &H10
Private Const OFN_ENABLESIZING           As Long = &H800000
Private Const OFS_MAXPATHNAME            As Long = 260




Private Sub Parcours(iSens As Integer)
    Dim IxPage As Integer

    ' Une variable de niveau module est remise à zéro après ajout d'une référence (recompilation auto ?)
    IxPage = Me.tabDEMO.Value

    Me.tabDEMO.Pages(IxPage + iSens).Visible = True ' Rendre visible la page
    Me.tabDEMO.Pages(IxPage + iSens).SetFocus       ' Placer le focus sur la page

    Me.tabDEMO.Pages(IxPage).Visible = False        ' Rendre invisible la page actuelle

    IxPage = IxPage + iSens                         ' Passer à la page suivante/précédente

    Me.cmdPrec.Enabled = (IxPage > 0)
    Me.cmdSvt.Enabled = (IxPage < Me.tabDEMO.Pages.Count - 1)

    Call Actions
End Sub

Private Sub Actions()
    Dim sMsg As String

    Select Case Me.tabDEMO.Pages(Me.tabDEMO.Value).Name
        Case "pgBienvenue"
            ' Rien

        Case "pgBiblio"
            Me.cmdPrec.SetFocus

            ' Chercher dans les références la présence de la bibliothèque
            If RefLib() Then
                sMsg = "La bibliothèque est correctement référencée." & vbCrLf & _
                       "    Chemin : " & Application.References("libMAIL").FullPath
                Me.cmdReferences.Visible = False
                Me.lblBiblio01.ForeColor = RGB(0, 127, 0)

            Else
                sMsg = "La blibliothèque n'a pas été trouvée dans les références de l'application." & vbCrLf & _
                       "Cliquez sur le bouton ci-contre pour localiser le fichier libMAIL.mdb sur votre poste de travail."
                Me.cmdReferences.Visible = True
                Me.lblBiblio01.ForeColor = RGB(200, 0, 0)

            End If
            Me.lblBiblio01.Caption = sMsg

            ' Présence de la table.
            If RefTbl() Then
                sMsg = "La table BoiteMail existe dans l'application." & vbCrLf
                On Error Resume Next
                sMsg = sMsg & "   Version de la table : " & CurrentDb.TableDefs("BoiteMail").Properties!VersTbl
                On Error GoTo 0
                Me.cmdBoiteMail.Visible = False
                Me.lblBiblio03.ForeColor = RGB(0, 127, 0)

            Else
                sMsg = "La table 'BoiteMail' n'a pas été trouvée dans l'application." & vbCrLf & _
                       "Cliquez sur le bouton ci-contre pour la créer."
                Me.cmdBoiteMail.Visible = RefLib()
                Me.lblBiblio03.ForeColor = RGB(200, 0, 0)

            End If
            Me.lblBiblio03.Caption = sMsg
            Me.lblBiblio03.Visible = RefLib()
            Me.lblBiblio02.Visible = RefLib()

            Me.cmdSvt.Enabled = (RefLib() And RefTbl())
            Me.lblBiblio04.Visible = Me.cmdSvt.Enabled

        Case "pgMSG"
            ' Si les deux adresses mails sont renseignées à peu près correctement...
            Me.cmdEnregistrer.Enabled = (Nz(Me.txtEmetteur, "") Like "*?@?*.??*" And Nz(Me.txtDest, "") Like "*?@?*.??*")

            ' S'il y a au moins un message à envoyer dans la table.
            Me.cmdSvt.Enabled = (DCount("*", "BoiteMail", "Etat='E'") > 0)
            Me.lblMSG02.Visible = Me.cmdSvt.Enabled

        Case "pgSRV"
            Me.cmdEnvoyer.Enabled = Len(Nz(Me.txtSrv, ""))

            Me.lblSrv02.Visible = (DCount("*", "BoiteMail", "Etat='V'") > 0)

    End Select
End Sub

' Est-ce que la référence à libMAIL existe ?
Private Function RefLib() As Boolean
    Dim Ref As Reference

    For Each Ref In Application.References
        If Ref.Name = "libMAIL" Then
            RefLib = True
            Exit For
        End If
    Next Ref
End Function

' Vérifier l'existence de la table BoiteMail
Private Function RefTbl() As Boolean
    Dim tbl As TableDef

    For Each tbl In CurrentDb.TableDefs
        If tbl.Name = "BoiteMail" Then
            RefTbl = True
            Exit For
        End If
    Next tbl
End Function



Private Sub cmdBoiteMail_Click()
    Call VerifieBAL

    Call Actions                                        ' Mettre l'onglet à jour.
End Sub

Private Sub cmdCreeMail_Click()
    Call CreeMail("", "", "", "", , , , , True)
End Sub

Private Sub cmdEnregistrer_Click()
    Dim sID As String

    ' Ce bouton vide la table avant de créer le message.
    CurrentDb.Execute "DELETE * FROM BoiteMail"

    sID = CreeMail(Me.txtDest, Nz(Me.txtObjet, "Test libMAIL"), Nz(Me.txtMSG, ""), Me.txtEmetteur)
    MsgBox "Le message n°" & sID & " a été créé." & vbCrLf & _
           "Cliquez sur le bouton 'Suivant' pour continuer."

    Call Actions
End Sub

Private Sub cmdEnvoyer_Click()
    Call SMTPLance(Me.txtSrv & ":" & Me.txtPort, , , , "", False, 30)
    Me.lblSrv02.Visible = True
End Sub

Private Sub cmdGestMail_Click()
    Call mnuGestMail
End Sub

Private Sub cmdJournal_Click()
    Call SMTPFormJnl
End Sub

Private Sub cmdPrec_Click()
    Call Parcours(-1)
End Sub

Private Sub cmdReferences_Click()
    Dim dtuOFN As OPENFILENAME, sFiltre As String, vFichiers As Variant
    Dim i As Integer, l As Long

    sFiltre = "Bases de données MS-Access" & vbNullChar & "*.md?;*.accd?" & vbNullChar & _
              "Tous les fichiers" & vbNullChar & "*.*"

    With dtuOFN
        .lStructSize = Len(dtuOFN)
        .hWndOwner = Me.Hwnd
        .sFilter = sFiltre
        .lFilterIndex = 1
        .sFile = vbNullChar & Space$(32764)
        .lMaxFile = Len(.sFile)
        .sFileTitle = vbNullChar & Space$(512)
        .lMaxFileTitle = Len(.sFileTitle)
'        .sInitialDir = CurDir() & vbNullChar & Space$(512) & vbNullChar & vbNullChar
        .sDialogTitle = "Trouver la bibliothèque libMAIL..."
        .lFlags = OFN_ENABLESIZING Or OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_EXPLORER
        .sDefFileExt = "*.mdb"
    End With

    l = GetOpenFileName(dtuOFN)

    If l = 0 Then Exit Sub                              ' Annulé par l'utilisateur.

    Application.References.AddFromFile dtuOFN.sFile

    Call Actions                                        ' Mettre l'onglet à jour.
End Sub

Private Sub cmdSTOP_Click()
    Call SMTPDecharge
End Sub

Private Sub cmdSvt_Click()
    Call Parcours(1)
End Sub

Private Sub txtDest_Exit(Cancel As Integer)
    Call Actions
End Sub

Private Sub txtEmetteur_Exit(Cancel As Integer)
    Call Actions
End Sub

Private Sub txtSrv_Exit(Cancel As Integer)
    Call Actions
End Sub