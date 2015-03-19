Option Compare Database
Option Explicit

' Copyright 2009-2015 Denis SCHEIDT
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







' Permet d'activer/désactiver les options du menu
' Spécifier les constantes des options à DESACTIVER
Sub EtatMenu(lEtatMenu As Long)
    dtuEtatSyst.Tray.EtatMenu = lEtatMenu
End Sub


' Affiche le formulaire d'état. Il n'est pas possible d'appeler directement une procédure avec paramètres
' depuis OnAction (pour une fonction, ça marche...).
Sub mnuAffEtat()
    Call SMTPFormEtat(True)
End Sub

' Ouvre le formulaire de création de message, tous les champs sont vides.
Sub mnuCreeMail()
    Call CreeMail("", "", "", "", , , , , True)
End Sub

Sub mnuAPropos()
    DoCmd.OpenForm "frm_APropos"
End Sub

Sub mnuGestMail()
    DoCmd.OpenForm "frm_GestionBoiteMail"
End Sub

' Contrôle l'existence d'un fichier.
Function FichierExiste(sSpecFichier As String) As Boolean
    Dim i As Integer

    On Error Resume Next
    i = GetAttr(sSpecFichier)                                       ' Faire un accès au fichier
    FichierExiste = (Err.Number = 0 And ((i And vbDirectory) <> vbDirectory))   ' Rép. exclus
    On Error GoTo 0
End Function

' Détermine si le formulaire est chargé ou non
Function FrmEstCharge(sNomForm) As Boolean
    FrmEstCharge = SysCmd(acSysCmdGetObjectState, acForm, sNomForm)
End Function

' Retourne le nom de l'ordinateur
' Par défaut lType=3 : ComputerNameDnsFullyQualified = nom complet de l'ordinateur
Function myComputerName(Optional lType As Long = 3&) As String
    Dim s As String, lRet As Long, lNbCar As Long

    s = Space$(512): lNbCar = 510
    lRet = GetComputerNameEx(lType, s, lNbCar)
    If lRet <> 0 Then myComputerName = Left$(s, lNbCar)
End Function

' Retourne le login Windows (nom de l'utilisateur ayant ouvert la session)
Function myCurrentUser() As String
    Dim l As Long, sUtilisateur As String

    sUtilisateur = Space$(256)                                      ' Variable tampon pour le retour de la fonction
    l = Len(sUtilisateur)                                           ' Longueur du tampon
    l = WNetGetUser(vbNullChar, sUtilisateur, l)
    If l = 0 Then
        ' Tronquer la chaîne à la partie utile (après le &H00)
        myCurrentUser = Left$(sUtilisateur, InStr(sUtilisateur, vbNullChar) - 1)
    Else
        ' Valeur de repli
        myCurrentUser = CurrentUser()
    End If
End Function

' Remplace les occurrences d'une chaine dans une autre.
' Le fonctionnement est similaire à la fonction Replace() de VB
Function Remplacer(sExpression As String, sCherche As String, sRemplace As String, _
                   Optional lDebut As Long = 1, _
                   Optional lNbRempl As Long = -1, _
                   Optional iCompare As Integer = VbCompareMethod.vbBinaryCompare) As String
    Dim sResult As String, nbMax As Long, lExpr As Long, lPE As Long, lPS As Long, l As Long, nbRempl As Long


    If Len(sExpression) = 0 Or lDebut > Len(sExpression) Then Exit Function

    If Len(sCherche) = 0 Or lNbRempl = 0 Then
        Remplacer = sExpression
        Exit Function
    End If


    If Len(sRemplace) > Len(sCherche) Then                         ' Pré-allouer de l'espace
        nbMax = Len(sExpression) * 2
    Else
        nbMax = Len(sExpression)
    End If
    sResult = Space$(nbMax)
    lExpr = Len(sExpression)
    lPE = 1                                                         ' Dernière position trouvée
    lPS = 1                                                         ' Position d'écriture

    l = InStr(lDebut, sExpression, sCherche, iCompare)              ' Chercher la première occurrence
    If l = 0 Then l = lExpr + 1                                     ' Si pas trouvée...
    Do
        ' Agrandir la chaine de sortie ?
        If (lPS + l - lPE) > Len(sResult) Then sResult = sResult & Space$(nbMax)

        ' Copier la partie dans la chaine de sortie
        Mid$(sResult, lPS) = Mid$(sExpression, lPE, l - lPE)
        lPS = lPS + l - lPE                                         ' Ajuster la position de sortie

        ' Si une occurrence a été trouvée
        If l <= lExpr Then
            ' Agrandir la chaine de sortie ?
            If (lPS + Len(sRemplace)) > Len(sResult) Then sResult = sResult & Space$(nbMax)

            ' Remplacer l'occurrence trouvée
            Mid$(sResult, lPS) = sRemplace
            lPS = lPS + Len(sRemplace)                              ' Ajuster la position de sortie

            nbRempl = nbRempl + 1                                   ' Compter les remplacements effectués

        End If

        lPE = l + Len(sCherche)                                     ' Avancer le point de départ de la recherche

        If (lNbRempl = -1) Or (nbRempl < lNbRempl) Then             ' Nombre de remplacements atteint ?
            l = InStr(lPE, sExpression, sCherche, iCompare)         ' Chercher l'occurrence suivante
            If l = 0 Then l = lExpr + 1                             ' Si pas trouvée

        Else
            l = lExpr + 1                                           ' On ignore toutes les occurrences suivantes

        End If
    Loop While lPE <= lExpr                                         ' Le pointeur d'entrée a dépassé la fin de la chaine

    Remplacer = Left$(sResult, lPS - 1)                             ' Ne garder que la partie utile
End Function

' Scinde une chaine de caractère et retourne un tableau contenant les éléments individuels.
' Le fonctionnement est similaire à la fonction Split() de VB.
Function Scinder(sChaine As String, _
                 Optional sDelim As String = " ", _
                 Optional iNbFragments As Integer = -1, _
                 Optional iCompare As Integer = VbCompareMethod.vbBinaryCompare) As Variant
    Dim i As Long, j As Long, l As Long
    Dim lDel As Integer, nb As Long, nbMax As Long, sResult() As String

    ' Si iNbFragment vaut zéro, la fonction retourne Empty
    If iNbFragments = 0 Then Exit Function

    nbMax = 32000
    ReDim sResult(nbMax)

    l = Len(sChaine): lDel = Len(sDelim): j = -lDel + 1
    Do
        i = j + lDel                                                ' Départ suivant
        j = InStr(i, sChaine, sDelim, iCompare)                     ' Délimiteur suivant

        If j = 0 Or lDel = 0 Then j = l + 1                         ' Aucun trouvé, ou pas de délimiteur spécifié, on prend tout ce qui reste
        If iNbFragments > 0 Then
            If nb + 1 >= iNbFragments Then j = l + 1
        End If

        If nb > nbMax Then
            nbMax = nbMax + 32000
            ReDim Preserve sResult(nbMax)                           ' Agrandir le tableau
        End If

        sResult(nb) = Mid$(sChaine, i, j - i)                       ' Insérer l'élément
        nb = nb + 1                                                 ' Elément suivant
    Loop While j < l

    ReDim Preserve sResult(nb - 1)                                  ' Eliminer les éléments inutiles
    Scinder = sResult()
    Erase sResult
End Function

' Cherche une chaine dans une autre en partant de la fin.
Function InStrFin(sChaine As String, sCherche As String, _
                  Optional lDebut As Long = -1, Optional iCompare As Integer = VbCompareMethod.vbBinaryCompare) As Long
    Dim i As Long, j As Long

    If Len(sChaine) = 0 Then Exit Function
    If Len(sCherche) = 0 Then InStrFin = lDebut: Exit Function
    If InStr(1, sChaine, sCherche, iCompare) = 0 Then Exit Function ' Aucune correspondance.
    If lDebut < -1 Or lDebut = 0 Then Exit Function

    ' Position de début
    If lDebut = -1 Or lDebut > Len(sChaine) Then
        i = Len(sChaine) - Len(sCherche) + 1
    Else
        i = lDebut
    End If

    Do While i > 0
        j = InStr(i, sChaine, sCherche, iCompare)
        If j = i Then
            InStrFin = j
            Exit Do
        End If
        i = i - 1
    Loop
End Function

' Transforme un tableau en chaine délimitée.
' Equivalent à la fonction Join() de VB.
Function Joindre(sTableau As Variant, Optional sDelim As String = " ") As String
    Dim i As Long, j As Long, l As Long, s As String, sResult As String

    sResult = Space$(65535)                                         ' Pré-allocation d'espace.
    j = 1                                                           ' Position d'écriture.

    For i = 0 To UBound(sTableau)
        s = sTableau(i) & sDelim
        l = Len(s)
        
        ' Agrandir la chaine si nécessaire.
        If (j + l) > Len(sResult) Then sResult = sResult & Space$(65535)
        Mid$(sResult, j, l) = s
        j = j + l
    Next i

    Joindre = Left$(sResult, j - Len(sDelim) - 1)
End Function

' Applique une nouvelle langue. Met à jour tous les formulaires ouverts.
' Doit être Function pour l'appel depuis les menus.
Function ChangeLang(lIDLang As Long)
    Dim frm As Form, lNouvLang As Long

    lNouvLang = LangueExiste(lIDLang)                                       ' Appliquer la nouvelle langue.
    If dtuEtatSyst.IDLang = lNouvLang Then Exit Function                    ' Pas la peine d'appliquer la même langue deux fois de suite.

    dtuEtatSyst.IDLang = lNouvLang

    ' Passer en revue les formulaires ouverts, et y appliquer la nouvelle langue.
    For Each frm In Application.Forms
        On Error Resume Next                                                ' La méthode n'existe que pour les formulaires libMAIL.
        Call frm.ChangeLang
        On Error GoTo 0
    Next frm

    ' Mettre aussi à jour le menu contextuel de l'icône.
    Call LangueMenu
End Function

' Charge la langue spécifiée à partir du fichier <IDLang>.t9n présent dans le répertoire de la bibliothèque.
' Si aucun IDLang n'est fourni, charge tous les fichiers trouvés.
Sub ChargeT9N(Optional ByVal lIDLang As Long = 0)
    Dim sRepert As String, i As Long, s As String, s1 As String, iNF As Integer, sInput As String, sLigne As Variant
    Dim rs As DAO.Recordset

    ' Récupérer le nom du répertoire.
    sRepert = CodeDb.Name
    ' Position du nom du fichier de base de données dans la chaîne.
    i = InStr(sRepert, Dir$(sRepert))
    ' Ne garder que le chemin.
    sRepert = Left$(sRepert, i - 1)

    ' Ouvrir la table
    Set rs = CodeDb.OpenRecordset("T9N", dbOpenTable)
    rs.Index = "PrimaryKey"

    ' Charger les fichiers .t9n présents dans le même répertoire.
    s = Dir$(sRepert & IIf(lIDLang = 0, "*", lIDLang) & ".t9n", vbNormal)
 
    Do While Len(s) <> 0
        iNF = FreeFile()

        Open sRepert & s For Input Access Read As #iNF
        lIDLang = Val(s)                                            ' Récupérer le code langue.
        Do While Not EOF(iNF)
            Line Input #iNF, sInput                                 ' Lire une ligne du fichier.

            sInput = Trim$(sInput)                                  ' Supprimer les espaces de début et de fin.
            ' Si ce n'est pas un commentaire ou une ligne vide.
            If Not (sInput Like ";*") And Len(sInput) > 0 Then
                sLigne = Scinder(sInput, "=", 2)                    ' Séparer clé et valeur.

                If UBound(sLigne) = 1 Then                          ' Il y avait au moins un '=' dans la ligne.
                    ' Créer ou mettre à jour l'enregistrement.
                    s1 = Trim$(Remplacer((sLigne(0)), vbTab, ""))   ' Supprimer d'éventuelles tabulations dans la clé.
                    rs.Seek "=", lIDLang, s1
                    If rs.NoMatch Then
                        rs.AddNew
                        rs!IDLang = lIDLang
                        rs!CleMsg = Trim$(s1)
                    Else
                        rs.Edit
                    End If
                    rs!MsgT9N = sLigne(1)
                    rs.Update
                End If

            End If

        Loop

        Close #iNF

        s = Dir$()                                                  ' Fichier suivant.
    Loop

    rs.Close
    Set rs = Nothing
End Sub