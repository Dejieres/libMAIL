Option Compare Database
Option Explicit
Option Private Module

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




' Version de la table BoiteMail. Incrémenter la constante à chaque nouvelle version.
' ----------------------------------------------------------------------------------
Public Const VersNTbl           As Byte = 3                                 ' Version 3 - le 11/04/2011


' Détermine si c'est une table attachée et retourne les infos pour la table réelle (chemin et nom)
Function TableAttachee(tbl As DAO.TableDef, sConnect As String, sSourceTableName As String) As Boolean
    Dim i As Integer

    If tbl.Attributes And dbAttachedTable = dbAttachedTable Then            ' C'est une table attachée
        sConnect = tbl.connect                                              ' Chemin complet du fichier MDB
        i = InStr(sConnect, ";DATABASE=")
        sConnect = Mid$(sConnect, i + 10)
        i = InStr(sConnect, ";")
        If i > 0 Then sConnect = Left$(sConnect, i - 1)

        sSourceTableName = tbl.SourceTableName

        TableAttachee = True
    End If
End Function

' Vérifie la spécification de fichier
' Retourne sSpecFich si OK
' Retourne "" si invalide.
Function VerifieFich(sSpecFich As String) As String
    Dim i As Integer

    i = FreeFile

    On Error Resume Next

    ' Tente d'ouvrir le fichier. S'il n'existe pas il est créé.
    Open sSpecFich For Append Access Write As #i
    Close #i

    ' Si OK, retourne la spec de fichier, sinon "".
    If Err.Number = 0 Then VerifieFich = sSpecFich

    On Error GoTo 0
End Function

' Retourne une spécification de fichier pour un fichier temporaire.
Function FichTemp(Optional sPrefixe As String = "DTU", Optional sExtension As String = "tmp") As String
    Dim sFichTmp As String, sNomFich As String, i As Integer

    ' Génère un nom aléatoire et unique, pour ne pas risquer de retomber sur un fichier
    ' contenant déjà des données.
    Randomize Timer
    sNomFich = "lml_" & sPrefixe & Hex$(Rnd * &HFFFFF) & "." & sExtension ' Construire le nom avec une partie aléatoire.

    sFichTmp = VerifieFich(Environ$("Temp") & "\" & sNomFich)           ' Rép. temporaire
    If Len(sFichTmp) = 0 Then
        sFichTmp = VerifieFich(Environ$("Tmp") & "\" & sNomFich)        ' Rép. temporaire
        If Len(sFichTmp) = 0 Then
            sFichTmp = CurrentDb.Name                                   ' Rép. de la BDD
            i = Len(sFichTmp)
            Do While Mid$(sFichTmp, i, 1) <> "\" And i > 0
                i = i - 1
            Loop
            sFichTmp = VerifieFich(Left$(sFichTmp, i) & sNomFich)
        End If
    End If
    FichTemp = sFichTmp
End Function

' Renvoie le chemin d'accès à un dossier spécial de Windows, sans \ final.
Function DossierSpecial(lDossier As Long) As String
    Dim sChem As String, dtuID As ITEMIDLIST, r As Long

    ' Chercher le Bureau virtuel.
    r = SHGetSpecialFolderLocation(0&, lDossier, dtuID)

    If r = 0 Then                                                       ' Pas d'erreur...
        sChem = String$(512, Chr$(0))
        r = SHGetPathFromIDList(ByVal dtuID.shellID.SHItem, ByVal sChem)
        If r Then DossierSpecial = Left$(sChem, InStr(sChem, Chr$(0)) - 1)
    End If
End Function