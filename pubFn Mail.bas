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


' Permet de court-circuiter le test sur la table BoiteMail lorsque celle-ci r�side sur un serveur (mySQL, SQLServer, etc...)
#Const mySQL = False




' Cr�e un message dans la table BoiteMail. Passe les options par d�faut � ECreeMail pour OptESMTP
Function CreeMail(sDestinataires As String, _
                  sObjetMsg As String, _
                  sTexteMessage As String, _
                  sExpediteur As String, _
                  Optional sUtilisateur As String = "", _
                  Optional sCC As String = "", _
                  Optional sBCC As String = "", _
                  Optional sPiecesJointes As Variant, _
                  Optional bEditeMail As Boolean = False, _
                  Optional dDifferer As Date = 0, _
                  Optional dConserver As Date = 0) As String
    Dim OptESMTP As tuESMTP_MSG, MsgMIME As tuMessageMIME

    MsgMIME.Texte = sTexteMessage
    CreeMail = ECreeMailMIME(sDestinataires, sObjetMsg, MsgMIME, sExpediteur, OptESMTP, sUtilisateur, sCC, sBCC, sPiecesJointes, bEditeMail, dDifferer, dConserver)
End Function

' Cr�e un message dans la table BoiteMail. Passe les options par d�faut � ECreeMail pour OptESMTP
Function ECreeMail(sDestinataires As String, _
                   sObjetMsg As String, _
                   sTexteMessage As String, _
                   sExpediteur As String, _
                   OptionsESMTP As tuESMTP_MSG, _
                   Optional sUtilisateur As String = "", _
                   Optional sCC As String = "", _
                   Optional sBCC As String = "", _
                   Optional sPiecesJointes As Variant, _
                   Optional bEditeMail As Boolean = False, _
                   Optional dDifferer As Date = 0, _
                   Optional dConserver As Date = 0) As String
    Dim MsgMIME As tuMessageMIME

    MsgMIME.Texte = sTexteMessage
    ECreeMail = ECreeMailMIME(sDestinataires, sObjetMsg, MsgMIME, sExpediteur, OptionsESMTP, sUtilisateur, sCC, sBCC, sPiecesJointes, bEditeMail, dDifferer, dConserver)
End Function

' Cr�e un message dans la table BoiteMail
' Le tableau sPiecesJointes() contient les noms et chemins des pi�ces jointes
'   Colonne 0 : Le nom seul du fichier � joindre
'   Colonne 1 : le chemin complet + le nom du fichier � joindre.
' La colonne est la premi�re dimension afin d'autoriser les ReDim Preserve.
'
' Retourne :
'       : Identifiant du message cr�� sur 18 caract�res
'   ""  : Erreur sur la table BoiteMail
'         Le cas �ch�ant, le code d'erreur sur une PJ est disponible dans la colonne 0 du tableau,
'         sur 5 positions suivi d'un ':'
Function ECreeMailMIME(sDestinataires As String, _
                       sObjetMsg As String, _
                       sTexteMessage As tuMessageMIME, _
                       sExpediteur As String, _
                       OptionsESMTP As tuESMTP_MSG, _
                       Optional sUtilisateur As String = "", _
                       Optional sCC As String = "", _
                       Optional sBCC As String = "", _
                       Optional sPiecesJointes As Variant, _
                       Optional bEditeMail As Boolean = False, _
                       Optional dDifferer As Date = 0, _
                       Optional dConserver As Date = 0) As String

    Dim dtuMail As tuMAIL, i As Integer

    ' La table BoiteMail doit exister
    If Not VerifieBAL() Then
        ECreeMailMIME = ""
        Exit Function
    End If

    With sTexteMessage
        ' Si le membre HTML de sTexteMessage est une sp�cification de fichier, alors charger ce fichier.
        If .HTML Like "?:\*.htm*" Or _
           .HTML Like "\\*\*\*.htm*" Then
             .HTML = HTMLCharge(.HTML)              ' Charge le fichier HTML
        End If

        ' Si le membre texte brut est vide, alors convertir le membre HTML.
        If Len(.Texte) = 0 Then .Texte = HTMLaTexte(.HTML)
    End With

    ' Initialise la DTU pour le passage de param�tres
    With dtuMail
        .a = sDestinataires
        .Objet = sObjetMsg
        .Message = sTexteMessage
        .De = sExpediteur
        .OptionsMSG = OptionsESMTP
        .Utilisateur = sUtilisateur
        .cc = sCC
        .BCC = sBCC
        If Not IsMissing(sPiecesJointes) Then .PJ = sPiecesJointes
        .Differer = dDifferer
        .Conserver = dConserver
    End With

    ' Edition du message ?
    If bEditeMail Then
        DoCmd.OpenForm "frm_EditeMail", acNormal, , , , acWindowNormal
        Forms("frm_EditeMail").MAIL = VarPtr(dtuMail)           ' Transmettre les param�tres au formulaire (pointeur sur la DTU)

    Else
        ' Sauvegarder le message dans la table et remonter l'identifiant du message cr��.
        Call SauveMail(dtuMail)
        ECreeMailMIME = dtuMail.Identifiant
        If Not IsMissing(sPiecesJointes) Then                   ' Remonter les modifs du tableau en cas d'erreur de PJ
            ' Pas d'affectation directe possible... il faut faire une boucle...
            For i = 0 To UBound(dtuMail.PJ, 2)
                If dtuMail.PJ(0, i) Like "#####:*" Then
                    sPiecesJointes(0, i) = dtuMail.PJ(0, i)
                End If
            Next i

        End If

    End If
End Function

' Modification d'un message existant.
Sub ModifieMail(sIdentifiant As String, Optional bAttendre As Boolean = False)

    ' Impossible pendant un envoi, car les messages en �tat 'E' ont d�j� �t� s�lectionn�s par frm_SMTP.
    If dtuEtatSyst.EtatSrv.Etat = lmlSrvEnCours Or dtuEtatSyst.EtatSrv.Etat = lmlSrvConnexion Then
        MsgBox Traduit("mod_impossible", "Il n'est pas possible de modifier un message alors que le serveur traite les messages de la boite d'envoi."), vbExclamation
        Exit Sub
    End If


    DoCmd.OpenForm "frm_EditeMail", acNormal, , , , acWindowNormal
    Forms!frm_EditeMail.IDMail = sIdentifiant                   ' ID du message � modifier


    ' Attendre la fermeture du formulaire d'�dition.
    Do While bAttendre And FrmEstCharge("frm_EditeMail")
        myDoEvents
    Loop
End Sub

' Lit un fichier disque et retourne le contenu dans une variable.
Function PJFichier(sSpecFichier As String, Optional lNbCar As Long = -1) As String
    Dim i As Integer, l As Long

    If FichierExiste(sSpecFichier) Then
        i = FreeFile()
        Open sSpecFichier For Binary Access Read Shared As #i
        l = IIf(lNbCar = -1, LOF(i), Abs(lNbCar))               ' Nombre de caract�res � lire
        PJFichier = Input(l, #i)                                ' On lit tout d'un coup !
        Close #i
    Else
        PJFichier = Traduit("att_notexists", "***** Le fichier '%s' n'existe pas. ***** -\n%s %s", sSpecFichier, Err.Number, Err.Description)
    End If
End Function

' Nombre de mails en attente
Function NbMails(Optional bDiff As Boolean = False) As Long
    Dim sWHERE As String

    sWHERE = "Etat='E'"
    ' Compter les messages diff�r�es ?
    If Not bDiff Then sWHERE = sWHERE & " AND Nz(Differer,0) < " & Format$(Now(), "\#mm-dd-yyyy hh:nn:ss\# ")

    NbMails = DCount("Identifiant", TableMail(), sWHERE)
End Function

' Les enregistrements 'D' sont supprim�s sans distinction.
'
' Purge les enregistrements en �tat 'V' :
' Si vSelection est num�rique,
'       >0 : conserve les n enregistrements les plus r�cents
'       <0 : conserve les n derniers jours
'       =0 : supprime tous les messages 'V'
' Si vSelection est une date, supprime les enregistrements ant�rieurs � la date
'
' Retourne le nombre d'enregistrements supprim�s
Function Purge(vSelection As Variant) As Long
    Dim db As DAO.Database, SQL As String, WHERE As String, n As Long

    Set db = CurrentDb

    ' Requ�te de base
    SQL = "DELETE * FROM " & TableMail() & " WHERE Etat = 'D' OR Etat='V' AND Nz(Conserver,0) < " & Format$(Now(), "\#mm-dd-yyyy hh:nn:ss\# ")

    ' D�terminer les crit�res
    If IsDate(vSelection) Then
        ' Efface tous les messages ant�rieurs � la date fournie.
        WHERE = "AND DateMsg < #" & Format$(vSelection, "mm-dd-yyyy") & "#;"

    ElseIf IsNumeric(vSelection) Then
        n = Val(vSelection)

        Select Case n
            Case Is > 0                 ' Ne garder que les n enregistrements les plus r�cents.
                WHERE = "AND Identifiant Not In (" & _
                            "SELECT TOP " & n & " Identifiant " & _
                            "FROM " & TableMail() & _
                           " WHERE Etat='V' " & _
                            "ORDER BY DateMsg DESC);"

            Case Is < 0                 ' Garder les n derniers jours, par rapport � la date du jour.
                WHERE = "AND DateMsg <= #" & Format$(DateAdd("d", n, Date), "mm-dd-yyyy") & "#;"

            Case Else                   ' Tout effacer (garder 0 enregistrements)
                WHERE = ";"

        End Select

    End If

    If Len(WHERE) <> 0 Then db.Execute SQL & WHERE

    Purge = db.RecordsAffected          ' Nb d'enregistrements supprim�s.

    db.Close
    Set db = Nothing
End Function

' V�rification de l'existence de la table g�rant la bo�te aux lettres
Function VerifieBAL() As Boolean
    Dim db As DAO.Database, tblBoiteMail As DAO.TableDef, ixIndex As Index, rs As DAO.Recordset
    Dim sConnect As String, sNomTable As String, VersATbl As Byte

    sNomTable = TableMail()

    Set db = CurrentDb

    On Error Resume Next
    Set tblBoiteMail = db.TableDefs(sNomTable)
    On Error GoTo 0

    If tblBoiteMail Is Nothing Then                                 ' La table n'existait pas, il faut la cr�er.
#If mySQL Then
        ' Test simplifi� pour une table attach�e mySQL (ou autre).
        MsgBox Traduit("tbl_notexists", "Probl�me : la table '%s'n'a pas �t� trouv�e...", sNomTable), vbCritical + vbOKOnly, "Biblioth�que libMAIL"

#Else
        If Application.GetOption("Project Name") = "libMAIL" Then   ' On ne peut pas la cr�er dans la biblioth�que
            MsgBox Traduit("tbl_nocreate", "Il n'est pas possible de cr�er la table '%s' dans la base de donn�es biblioth�que !\n" & _
                   "Vous devez appeler cette fonction depuis votre application.", sNomTable), vbExclamation

        ' On propose de cr�er la bo�te mail, dans la base active.
        Else
            If MsgBox(Traduit("tbl_create", "La table '%s' n'existe pas dans votre base de donn�es.\nVoulez-vous la cr�er ?", sNomTable), _
                      vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                ' On cr�e la bo�te mail
                db.Execute "CREATE TABLE " & sNomTable & " (Identifiant TEXT (18) CONSTRAINT PrimaryKey PRIMARY KEY, Utilisateur TEXT (25), DateMsg DATETIME, Etat TEXT (1), " & _
                           "Expediteur TEXT (255), Destinataires LONGTEXT, CC LONGTEXT, BCC LONGTEXT, Objet TEXT (255), CorpsMsg LONGTEXT, ESMTP LONGTEXT," & _
                           "Differer DATETIME, Conserver DATETIME, DateEnvoi DATETIME);"
                db.TableDefs.Refresh                                ' Rendre la nouvelle table 'visible'
                VerifieBAL = True

                ' Ajouter le num�ro de version � la table
                Set tblBoiteMail = db.TableDefs(sNomTable)
                tblBoiteMail.Properties.Append tblBoiteMail.CreateProperty("VersTbl", dbByte, VersNTbl)

            End If

        End If

#End If

    Else                                                                ' La table existe d�j�.

#If mySQL Then
        ' On consid�re que la table mySQL est OK.
        VerifieBAL = True

#Else
        ' V�rifier la version de la table
        On Error Resume Next

        VersATbl = tblBoiteMail.Properties!VersTbl                      ' Version de la table actuelle (si la propri�t� existe)
        Err.Clear                                                       ' On ignore l'erreur ici

        If VersATbl < VersNTbl Then                                     ' Diff�rence entre les versions.
            If TableAttachee(tblBoiteMail, sConnect, sNomTable) Then    ' Contr�le si c'est une table attach�e
                Set tblBoiteMail = Nothing
                db.Close
                Set db = Nothing

                Set db = OpenDatabase(sConnect)
                Set tblBoiteMail = db.TableDefs(sNomTable)
            End If

            ' Il faut mettre la structure de la table � jour ===========================================
            With tblBoiteMail
                If VersATbl < 2 Then
                    .Fields.Append .CreateField("Identifiant", dbText, 18)  ' Cl� primaire de la table

                    ' Renseigner le champ cl� sur les enregistrements existants
                    Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & TableMail() & " WHERE Identifiant=''", dbOpenDynaset)
                    With rs
                        Do While Not .EOF
                            .Edit
                            !Identifiant = IDMail(!DateMsg)
                            .Update
                            .MoveNext
                        Loop
                        .Close
                    End With
                    Set rs = Nothing

                    Set ixIndex = .CreateIndex("PrimaryKey")
                    With ixIndex
                        .Fields.Append .CreateField("Identifiant")
                        .Primary = True
                    End With
                    .Indexes.Append ixIndex

                    .Fields.Append .CreateField("ESMTP", dbMemo, 0)
                End If

                If VersATbl < 3 Then
                    .Fields.Append .CreateField("Differer", dbDate)     ' Envoi diff�r� du message
                    .Fields.Append .CreateField("Conserver", dbDate)    ' Conserver le message jusqu'au...
                    .Fields.Append .CreateField("DateEnvoi", dbDate)    ' Date/Heure d'envoi du message.

                End If

            End With
            ' ==========================================================================================

            If Err.Number = 0 Then
                VerifieBAL = True                                       ' Tout s'est bien pass�...
                ' Ecrire le num�ro de version
                If VersATbl = 0 Then                                    ' Ajouter le N� de version, car il n'existait pas encore.
                    tblBoiteMail.Properties.Append tblBoiteMail.CreateProperty("VersTbl", dbByte, VersNTbl)

                Else                                                    ' Mettre � jour le N� de version
                    tblBoiteMail.Properties!VersTbl = VersNTbl

                End If
            Else
                MsgBox Err.Description, vbCritical, "Erreur " & Err.Number

            End If

        Else
            VerifieBAL = True

        End If

        On Error GoTo 0

#End If
    End If

    Set tblBoiteMail = Nothing
    db.Close
    Set db = Nothing
    Set ixIndex = Nothing
End Function

' Compte le nombre d'erreurs de pi�ce jointes.
' Si bRAZ = True, �limine les codes d'erreur par la m�me occasion.
Function ErreursPJ(sPJ() As String, Optional bRAZ As Boolean = False) As Long
    Dim i As Long, nPJ As Long

    For i = 0 To UBound(sPJ, 2)
        If sPJ(0, i) Like "#####:*" Then
            nPJ = nPJ + 1
            If bRAZ Then sPJ(0, i) = Mid$(sPJ(0, i), 7)             ' Elimine le code d'erreur.
        End If
    Next i
    ErreursPJ = nPJ
End Function

' Cr�e la chaine de sp�cification pour l'envoi d'un objet Access
Function PJOA(iTypeObjet As Integer, sFormatExport As String, sNomObjet As String, Optional iTypeExport As Integer = lmlPJDonnees) As String
    PJOA = iTypeExport & "/" & iTypeObjet & "/" & sFormatExport & "/" & sNomObjet
End Function

' Exporte un fichier au format EML.
' Si sSpecFich n'est pas fournie, enregistre dans Mes Documents\ID.eml.
' Retourne le code d'erreur le cas �ch�ant.
Function ExporteEML(sID As String, Optional ByVal sSpecFich As String = "") As Long
    Dim sChem As String, sNom As String, sExt As String, n As Integer
    Dim rs As DAO.Recordset, ESMTP_MSG As tuESMTP_MSG

    ' S�parer les �l�ments du chemin et forcer l'extension � .eml.
    ' Si la sp�cification est invalide ou vide, sauvegarder dans 'Mes Documents\<ID>.eml'
    Call AnaSpecFich(sSpecFich, sChem, sNom, sExt)
    If Len(sChem) = 0 Then sChem = DossierSpecial(CSIDL_MYDOCUMENTS) & "\"
    If Len(sNom) = 0 Then sNom = sID
    sSpecFich = sChem & sNom & ".eml"

    On Error Resume Next

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & TableMail() & " WHERE Identifiant='" & sID & "'", dbOpenDynaset, 0, dbReadOnly)
    If rs.RecordCount > 0 Then
        n = FreeFile()
        Open sSpecFich For Output As n
        Print #n, MSGEnTete(rs, ESMTP_MSG, True); rs!CorpsMsg;
        Close #n

        ExporteEML = Err.Number

    Else
        ExporteEML = -1                                             ' Aucun message ne correspond � l'ID fourni.

    End If
    rs.Close
    Set rs = Nothing

    On Error GoTo 0
End Function