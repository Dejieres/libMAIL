Option Compare Database
Option Explicit
Option Private Module

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

Private Const csNomProp As String = "DateDernModif"



Sub ChargeVB()
    ' Charge tout le code source depuis le répertoire du fichier .MDB
    Call ChargeVBX
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

    ' Charger les fichiers .FRM et .MOD présents dans le même répertoire
    s = Dir$(sRepert & "*.FRM", vbNormal)
    Do While Len(s) <> 0
        LoadFromText acForm, NomSeul(s), sRepert & s
        nbForms = nbForms + 1
        s = Dir$()
    Loop
    s = Dir$(sRepert & "*.MOD", vbNormal)
    Do While Len(s) <> 0
        If s <> "Outils.MOD" Then
            LoadFromText acModule, NomSeul(s), sRepert & s
            nbMods = nbMods + 1
        End If
        s = Dir$()
    Loop

    s = nbForms & " formulaire(s) et " & nbMods & " module(s) importé(s)." & vbCrLf & vbCrLf
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