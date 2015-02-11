Attribute VB_Name = "Outils"
Option Compare Database
Option Explicit
Option Private Module

' Copyright 2009-2012 Denis SCHEIDT
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

Private Const csNomProp As String = "DateDernModif"



Sub ChargeVB()
    ' Etablit avant tout les références nécessaires au fonctionnement correct de la bibliothèque

    ' GUID pour DAO3x0.DLL  : {00025E01-0000-0000-C000-000000000046}
	' Pas nécessaire à partir d'Access 2007 (=ACEDAO.dll)
     If Val(SysCmd(acSysCmdAccessVer)) < 12 Then Call CtrlRefs("{00025E01-0000-0000-C000-000000000046}")
    ' GUID pour MSO(97).DLL : {2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}
    Call CtrlRefs("{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}")

    ' Charge tout le code source depuis le répertoire du fichier .MDB
    Call ChargeVBX
End Sub

Sub SauveVB(Optional bTout As Boolean = False)
    Dim sRepert As String, i As Integer

    ' Récupérer le nom du répertoire
    sRepert = CurrentDb.Name
    ' Position du nom du fichier de base de données dans la chaîne
    i = InStr(sRepert, Dir$(sRepert))
    ' Ne garder que le chemin
    sRepert = Left$(sRepert, i - 1)

    If (ListeMods() = 0) And Not bTout Then
        MsgBox "Aucun document n'a été modifié..."
        Exit Sub
    End If

    Select Case MsgBox("Voulez-vous exporter les documents modifiés ?" & vbCrLf & _
                       "    Oui : exporter;" & vbCrLf & "    Non : annuler l'indicateur de modification;" & vbCrLf & "    Annuler : Abandonner.", _
                       vbQuestion + vbYesNoCancel + vbDefaultButton3)
        Case vbYes
            ' Sauvegarder les objets modifiés
            Call SauveDocs(acForm, sRepert, bTout, False)
            Call SauveDocs(acModule, sRepert, bTout, False)
            Debug.Print
            Debug.Print "Export terminé."

        Case vbNo
            ' RAZ des dates mémorisées
            Call SauveDocs(acForm, sRepert, bTout, True)
            Call SauveDocs(acModule, sRepert, bTout, True)
            Debug.Print
            Debug.Print "Indicateurs de modification réinitialisés."

        Case vbCancel
            Debug.Print "Abandon de l'export."
    End Select

End Sub




' Liste les objets modifiés
Private Function ListeMods() As Integer
    Dim db As DAO.Database, Doc As DAO.Document, nbDocs As Integer

    Set db = CurrentDb

    Debug.Print vbCrLf; "Documents modifiés depuis la dernière utilisation de SauveVb :"
    Debug.Print "Document____________________LastUpdated____________Modif. précédente___________"

    nbDocs = DocsMods(db.Containers!Forms.Documents) _
           + DocsMods(db.Containers!Modules.Documents)

    Debug.Print String$(79, "="); vbCrLf

    Set Doc = Nothing
    db.Close
    Set db = Nothing

    ListeMods = nbDocs
End Function

' Liste les documents modifiés pour un conteneur.
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

' Contrôle l'existence de la référence et tente de l'ajouter si elle n'existe pas.
Private Sub CtrlRefs(sGUID As String)
    Dim Ref As Access.Reference

    On Error Resume Next

    Set Ref = Application.References.AddFromGuid(sGUID, 0, 0)
    Select Case Err.Number
        Case 0                                  ' Référence ajoutée correctement
            MsgBox "Référence ajoutée pour " & Ref.Name & " (" & Ref.FullPath & ")", vbInformation

        Case 32813                              ' La référence existe déjà
            ' Rien

        Case Else
            MsgBox "Erreur " & Err.Number & "- " & Err.Description & " lors de l'ajout de la référence " & sGUID, vbCritical

    End Select

    Set Ref = Nothing

    On Error GoTo 0
End Sub

Private Sub ChargeVBX()
    Dim db As DAO.Database, Doc As DAO.Document, sRepert As String, i As Integer, s As String
    Dim nbForms As Integer, nbMods As Integer

    Set db = CurrentDb

    ' Vérifier que le module Outils existe bien
    For Each Doc In db.Containers!Modules.Documents
        If Doc.Name = "Outils" Then
            i = 1
            Exit For
        End If
    Next Doc
    If i = 0 Then
        MsgBox "Vous devez enregistrer le module 'Outils' avant de lancer la procédure ChargeVB.", vbCritical
        Exit Sub
    End If

    If MsgBox("Chargement des formulaires et des modules dans la base." _
              & vbCrLf & vbCrLf & "ATTENTION, cette commande va effacer *TOUS* les formulaires et modules avant d'importer les nouvelles versions." _
              & vbCrLf & vbCrLf & "Etes-vous sûr(e) de vouloir faire ça ?", _
              vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

    ' Effacer tous les formulaires et les modules, SAUF celui-ci !
    For Each Doc In db.Containers!Forms.Documents
        DoCmd.DeleteObject acForm, Doc.Name
    Next Doc
    For Each Doc In db.Containers!Modules.Documents
        If Doc.Name <> "Outils" Then DoCmd.DeleteObject acModule, Doc.Name
    Next Doc

    ' Récupérer le nom du répertoire
    sRepert = db.Name
    ' Position du nom du fichier de base de données dans la chaîne
    i = InStr(sRepert, Dir$(sRepert))
    ' Ne garder que le chemin
    sRepert = Left$(sRepert, i - 1)

    ' Charger les fichiers .frm et .bas présents dans le même répertoire
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

    ' Garantir que le nom de projet sera toujours le même.
    ' Access 2003 plante si on tente de réécrire la même valeur dans la propriété.
    If Application.GetOption("Project Name") <> "libMAIL" Then Application.SetOption "Project Name", "libMAIL"

    ' Compiler et enregistrer les modules
    DoCmd.RunCommand acCmdCompileAndSaveAllModules

    ' Ajouter une propriété personnalisée de type Date, pour garder la trace de la dernière date de modification
    ' --> permet de n'exporter que les documents modifiés
    ' La compilation modifie LastUpdated.
    ' L'ajout de la propriété doit être le dernier traitement...
    Call AjouteProp("Modules")                  ' Modules
    Call AjouteProp("Forms")                    ' Formulaires

    s = nbForms & " formulaire(s) et " & nbMods & " module(s) importé(s) et compilé(s)." & vbCrLf & vbCrLf
    s = s & "La bibliothèque libMAIL est fournie SANS AUCUNE GARANTIE." & vbCrLf
    s = s & "Ce programme est distribué sous licence LGPL v3 ou supérieure. Vous pouvez le modifier/redistribuer conformément aux termes de cette licence." & vbCrLf
    MsgBox s

    db.Close
    Set db = Nothing
End Sub


' Retourne le nom débarrassé de son extension
Private Function NomSeul(sNomFichier As String) As String
    Dim i As Integer

    ' Recherche le dernier point du nom
    i = Len(sNomFichier)
    Do While Mid$(sNomFichier, i, 1) <> "."
        i = i - 1
        If i = 0 Then
            i = Len(sNomFichier) + 1                ' Aucun point trouvé
            Exit Do
        End If
    Loop
    NomSeul = Left$(sNomFichier, i - 1)
End Function

' Ajoute une propriété aux documents.
Private Sub AjouteProp(sContainer As String)
    Dim db As DAO.Database, Doc As DAO.Document

    Set db = CurrentDb

    For Each Doc In db.Containers(sContainer).Documents
        Call CreeProp(Doc)                                              ' Création éventuelle de la propriété.
        Doc.Properties(csNomProp) = Doc.LastUpdated                     ' Date de dernière modif.
    Next Doc

    Set Doc = Nothing
    db.Close
    Set db = Nothing
End Sub

' Ajoute la propriété au document si nécessaire.
Private Sub CreeProp(Doc As DAO.Document)
    On Error Resume Next

    ' Ajouter la propriété, avec une date à 0. Si la propriété existe, elle ne sera pas modifiée.
    Doc.Properties.Append Doc.CreateProperty(csNomProp, dbDate, CDate(0))

    On Error GoTo 0
End Sub


' Exporte le code source des documents modifiés
'   bTout : Vrai=exporter tous les documents, même non modifié.
'   bRAZ  : Vrai=ne pas exporter, mais réinitialiser l'indicateur de modif.
Private Sub SauveDocs(lType As Long, sRepert As String, bTout As Boolean, bRAZ As Boolean)
    Dim sContainer As String, sExt As String, db As DAO.Database, Doc As DAO.Document

    Select Case lType
        Case acForm:    sContainer = "Forms":   sExt = ".frm"
        Case acModule:  sContainer = "Modules": sExt = ".bas"
        Case Else:      Exit Sub
    End Select

    Set db = CurrentDb

    For Each Doc In db.Containers(sContainer).Documents
        Call CreeProp(Doc)                                              ' Création éventuelle de la propriété.

        If (DateDiff("s", Doc.Properties(csNomProp), Doc.LastUpdated) > 0) Or bTout Then      ' Si le document a été modifié...
            If bRAZ Then
                Debug.Print "[" & Doc.Name & "] : Indicateur de modif réinitialisé sans export."
            Else
                SaveAsText lType, Doc.Name, sRepert & Doc.Name & sExt
                Debug.Print "[" & Doc.Name & "] exporté vers " & sRepert & Doc.Name & sExt
            End If
            Doc.Properties(csNomProp) = Doc.LastUpdated                 ' Enregistrer la date de dernière modif
        End If
    Next Doc

    db.Close
    Set db = Nothing
End Sub
