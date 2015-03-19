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

Private Const csNomProp As String = "DateDernModif"



Sub ChargeVB()
    ' Etablit avant tout les r�f�rences n�cessaires au fonctionnement correct de la biblioth�que

    ' GUID pour DAO3x0.DLL  : {00025E01-0000-0000-C000-000000000046}
    ' Pas n�cessaire � partir d'Access 2007 (=ACEDAO.dll)
     If Val(SysCmd(acSysCmdAccessVer)) < 12 Then Call CtrlRefs("{00025E01-0000-0000-C000-000000000046}")
    ' GUID pour MSO(97).DLL : {2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}
    Call CtrlRefs("{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}")

    ' Charge tout le code source depuis le r�pertoire du fichier .MDB
    Call ChargeVBX

    ' Cr�e la table des traductions.
    Call CreeT9N

End Sub

Sub SauveVB(Optional bTout As Boolean = False)
    Dim sRepert As String, i As Integer

    ' R�cup�rer le nom du r�pertoire
    sRepert = CurrentDb.Name
    ' Position du nom du fichier de base de donn�es dans la cha�ne
    i = InStr(sRepert, Dir$(sRepert))
    ' Ne garder que le chemin
    sRepert = Left$(sRepert, i - 1)

    If (ListeMods() = 0) And Not bTout Then
        MsgBox "Aucun document n'a �t� modifi�..."
        Exit Sub
    End If

    Select Case MsgBox("Voulez-vous exporter les documents modifi�s ?" & vbCrLf & _
                       "    Oui : exporter;" & vbCrLf & "    Non : annuler l'indicateur de modification;" & vbCrLf & "    Annuler : Abandonner.", _
                       vbQuestion + vbYesNoCancel + vbDefaultButton3)
        Case vbYes
            ' Sauvegarder les objets modifi�s
            Call SauveDocs(acForm, sRepert, bTout, False)
            Call SauveDocs(acModule, sRepert, bTout, False)
            Debug.Print
            Debug.Print "Export termin�."

        Case vbNo
            ' RAZ des dates m�moris�es
            Call SauveDocs(acForm, sRepert, bTout, True)
            Call SauveDocs(acModule, sRepert, bTout, True)
            Debug.Print
            Debug.Print "Indicateurs de modification r�initialis�s."

        Case vbCancel
            Debug.Print "Abandon de l'export."
    End Select

End Sub





' Liste les objets modifi�s
Private Function ListeMods() As Integer
    Dim db As DAO.Database, Doc As DAO.Document, nbDocs As Integer

    Set db = CurrentDb

    Debug.Print vbCrLf; "Documents modifi�s depuis la derni�re utilisation de SauveVb :"
    Debug.Print "Document____________________LastUpdated____________Modif. pr�c�dente___________"

    nbDocs = DocsMods(db.Containers!Forms.Documents) _
           + DocsMods(db.Containers!Modules.Documents)

    Debug.Print String$(79, "="); vbCrLf

    Set Doc = Nothing
    db.Close
    Set db = Nothing

    ListeMods = nbDocs
End Function

' Liste les documents modifi�s pour un conteneur.
Private Function DocsMods(Docs As DAO.Documents) As Integer
    Dim Doc As DAO.Document, nbDocs As Integer

    On Error Resume Next

    For Each Doc In Docs
        If DateDiff("s", Doc.Properties(csNomProp), Doc.LastUpdated) > 0 Then
            Debug.Print Doc.Name; Tab(29); Doc.LastUpdated; Tab(52);
            If Err.Number = 0 Then
                Debug.Print Doc.Properties(csNomProp)
            Else
                Debug.Print "-- N/A"
                Err.Clear
            End If
            nbDocs = nbDocs + 1
        End If
    Next Doc

    On Error GoTo 0

    DocsMods = nbDocs
End Function

' Contr�le l'existence de la r�f�rence et tente de l'ajouter si elle n'existe pas.
Private Sub CtrlRefs(sGUID As String)
    Dim Ref As Access.Reference

    On Error Resume Next

    Set Ref = Application.References.AddFromGuid(sGUID, 0, 0)
    Select Case Err.Number
        Case 0                                  ' R�f�rence ajout�e correctement
            MsgBox "R�f�rence ajout�e pour " & Ref.Name & " (" & Ref.FullPath & ")", vbInformation

        Case 32813                              ' La r�f�rence existe d�j�
            ' Rien

        Case Else
            MsgBox "Erreur " & Err.Number & "- " & Err.Description & " lors de l'ajout de la r�f�rence " & sGUID, vbCritical

    End Select

    Set Ref = Nothing

    On Error GoTo 0
End Sub

Private Sub ChargeVBX()
    Dim db As DAO.Database, Doc As DAO.Document, sRepert As String, i As Integer, s As String
    Dim nbForms As Integer, nbMods As Integer

    Set db = CurrentDb

    ' V�rifier que le module Outils existe bien
    For Each Doc In db.Containers!Modules.Documents
        If Doc.Name = "Outils" Then
            i = 1
            Exit For
        End If
    Next Doc
    If i = 0 Then
        MsgBox "Vous devez enregistrer le module 'Outils' avant de lancer la proc�dure ChargeVB.", vbCritical
        Exit Sub
    End If

    If MsgBox("Chargement des formulaires et des modules dans la base." _
              & vbCrLf & vbCrLf & "ATTENTION, cette commande va effacer *TOUS* les formulaires et modules avant d'importer les nouvelles versions." _
              & vbCrLf & vbCrLf & "Etes-vous s�r(e) de vouloir faire �a ?", _
              vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

    ' Effacer tous les formulaires et les modules, SAUF celui-ci !
    For Each Doc In db.Containers!Forms.Documents
        DoCmd.DeleteObject acForm, Doc.Name
    Next Doc
    For Each Doc In db.Containers!Modules.Documents
        If Doc.Name <> "Outils" Then DoCmd.DeleteObject acModule, Doc.Name
    Next Doc

    ' R�cup�rer le nom du r�pertoire
    sRepert = db.Name
    ' Position du nom du fichier de base de donn�es dans la cha�ne
    i = InStr(sRepert, Dir$(sRepert))
    ' Ne garder que le chemin
    sRepert = Left$(sRepert, i - 1)

    ' Charger les fichiers .frm et .bas pr�sents dans le m�me r�pertoire
    s = Dir$(sRepert & "*.frm", vbNormal)
    Do While Len(s) <> 0
        LoadFromText acForm, NomSeul(s), sRepert & s
        nbForms = nbForms + 1
        s = Dir$()
    Loop
    s = Dir$(sRepert & "*.bas", vbNormal)
    Do While Len(s) <> 0
        If s <> "Outils.bas" Then
            LoadFromText acModule, NomSeul(s), sRepert & s
            nbMods = nbMods + 1
        End If
        s = Dir$()
    Loop

    ' Garantir que le nom de projet sera toujours le m�me.
    ' Access 2003 plante si on tente de r��crire la m�me valeur dans la propri�t�.
    If Application.GetOption("Project Name") <> "libMAIL" Then Application.SetOption "Project Name", "libMAIL"

    ' Compiler et enregistrer les modules
    DoCmd.RunCommand acCmdCompileAndSaveAllModules

    ' Ajouter une propri�t� personnalis�e de type Date, pour garder la trace de la derni�re date de modification
    ' --> permet de n'exporter que les documents modifi�s
    ' La compilation modifie LastUpdated.
    ' L'ajout de la propri�t� doit �tre le dernier traitement...
    Call AjouteProp("Modules")                  ' Modules
    Call AjouteProp("Forms")                    ' Formulaires

    s = nbForms & " formulaire(s) et " & nbMods & " module(s) import�(s) et compil�(s)." & vbCrLf & vbCrLf
    s = s & "La biblioth�que libMAIL est fournie SANS AUCUNE GARANTIE." & vbCrLf
    s = s & "Ce programme est distribu� sous licence LGPL v3 ou sup�rieure. Vous pouvez le modifier/redistribuer conform�ment aux termes de cette licence." & vbCrLf
    MsgBox s

    db.Close
    Set db = Nothing
End Sub


' Retourne le nom d�barrass� de son extension
Private Function NomSeul(sNomFichier As String) As String
    Dim i As Integer

    ' Recherche le dernier point du nom
    i = Len(sNomFichier)
    Do While Mid$(sNomFichier, i, 1) <> "."
        i = i - 1
        If i = 0 Then
            i = Len(sNomFichier) + 1                ' Aucun point trouv�
            Exit Do
        End If
    Loop
    NomSeul = Left$(sNomFichier, i - 1)
End Function

' Ajoute une propri�t� aux documents.
Private Sub AjouteProp(sContainer As String)
    Dim db As DAO.Database, Doc As DAO.Document

    Set db = CurrentDb

    For Each Doc In db.Containers(sContainer).Documents
        Call CreeProp(Doc)                                              ' Cr�ation �ventuelle de la propri�t�.
        Doc.Properties(csNomProp) = Doc.LastUpdated                     ' Date de derni�re modif.
    Next Doc

    Set Doc = Nothing
    db.Close
    Set db = Nothing
End Sub

' Ajoute la propri�t� au document si n�cessaire.
Private Sub CreeProp(Doc As DAO.Document)
    On Error Resume Next

    ' Ajouter la propri�t�, avec une date � 0. Si la propri�t� existe, elle ne sera pas modifi�e.
    Doc.Properties.Append Doc.CreateProperty(csNomProp, dbDate, CDate(0))

    On Error GoTo 0
End Sub


' Exporte le code source des documents modifi�s
'   bTout : Vrai=exporter tous les documents, m�me non modifi�.
'   bRAZ  : Vrai=ne pas exporter, mais r�initialiser l'indicateur de modif.
Private Sub SauveDocs(lType As Long, sRepert As String, bTout As Boolean, bRAZ As Boolean)
    Dim sContainer As String, sExt As String, db As DAO.Database, Doc As DAO.Document

    Select Case lType
        Case acForm:    sContainer = "Forms":   sExt = ".frm"
        Case acModule:  sContainer = "Modules": sExt = ".bas"
        Case Else:      Exit Sub
    End Select

    Set db = CurrentDb

    For Each Doc In db.Containers(sContainer).Documents
        Call CreeProp(Doc)                                              ' Cr�ation �ventuelle de la propri�t�.

        If (DateDiff("s", Doc.Properties(csNomProp), Doc.LastUpdated) > 0) Or bTout Then      ' Si le document a �t� modifi�...
            If bRAZ Then
                Debug.Print "[" & Doc.Name & "] : Indicateur de modif r�initialis� sans export."
            Else
                SaveAsText lType, Doc.Name, sRepert & Doc.Name & sExt
                Debug.Print "[" & Doc.Name & "] export� vers " & sRepert & Doc.Name & sExt
            End If
            Doc.Properties(csNomProp) = Doc.LastUpdated                 ' Enregistrer la date de derni�re modif
        End If
    Next Doc

    db.Close
    Set db = Nothing
End Sub

' Cr�ation de la table de traductions.
Private Sub CreeT9N()
    Dim db As DAO.Database, td As DAO.TableDef

    Set db = CurrentDb

    On Error Resume Next
    Set td = db.TableDefs("T9N")
    On Error GoTo 0

    If Not td Is Nothing Then db.Execute "DROP TABLE T9N"               ' On supprime la table.

    ' Cr�er la table.
    db.Execute "CREATE TABLE T9N (IDLang LONG, CleMsg TEXT(50), MsgT9N LONGTEXT, CONSTRAINT PrimaryKey PRIMARY KEY (IDLang, CleMsg))"

    ' Chargement des langues.
    ' -----------------------
    ' L'appel doit se faire ici car le module contenant cette proc�dure n'est pas encore charg� lors de l'appel � ChargeVB,
    ' ce qui provoque une erreur de compilation. D�placer l'appel ici �vite cette erreur (� condition que la compilation � la demande soit activ�e).
    Call ChargeT9N

End Sub