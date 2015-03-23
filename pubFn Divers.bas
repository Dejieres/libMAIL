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
' Si Debut > 1, la chaine est tronquée des Debut-1 premiers caractères.
Function Remplacer(sExpression As String, sCherche As String, sRemplace As String, _
                   Optional lDebut As Long = 1, _
                   Optional lNbRempl As Long = -1, _
                   Optional iCompare As Integer = VbCompareMethod.vbBinaryCompare) As String
#If VBA6 Then
    Remplacer = Replace(sExpression, sCherche, sRemplace, lDebut, lNbRempl, iCompare)

#Else
    Dim sResult As String, nbMax As Long, lExpr As Long, lCher As Long, lRemp As Long, nbRempl As Long, i As Long
    Dim lPe1 As Long, lPe2 As Long, lPs As Long                     ' Pointeurs d'entrée et de sortie.

    lExpr = Len(sExpression)
    lCher = Len(sCherche)
    lRemp = Len(sRemplace)

    If lExpr = 0 Or lDebut > lExpr Then Exit Function

    lPe2 = InStr(lDebut, sExpression, sCherche, iCompare)           ' Chercher la première occurrence à partir de Debut.
    ' Aucune occurrence trouvée, ou on ne cherche rien, ou on ne remplace aucune fois.
    If lPe2 = 0 Or lCher = 0 Or lNbRempl = 0 Then
        Remplacer = Mid$(sExpression, lDebut)
        Exit Function
    End If

    nbMax = 0                                                       ' Pré-allouer de l'espace
    i = lPe2
    Do While i And (lNbRempl = -1 Or nbMax < lNbRempl)              ' Déterminer le nombre de remplacements à faire.
        nbMax = nbMax + 1&
        i = InStr(i + lCher, sExpression, sCherche, iCompare)
    Loop
    sResult = Space$(lExpr - (lCher - lRemp) * nbMax)

    lPs = 1                                                         ' Position d'écriture
    lPe1 = lDebut                                                   ' Début de la lecture.

    Do
        ' Ecrire la portion dans la chaine de sortie.
        Mid$(sResult, lPs, lPe2 - lPe1) = Mid$(sExpression, lPe1, lPe2 - lPe1)
        lPs = lPs + lPe2 - lPe1                                     ' Ajuster le pointeur de sortie.
        If lPs > Len(sResult) Then Exit Do

        ' Ecrire la valeur de remplacement.
        Mid$(sResult, lPs, lRemp) = sRemplace                       ' Ecrire la valeur de remplacement.
        lPs = lPs + lRemp                                           ' Ajuster le pointeur de sortie.
        If lPs > Len(sResult) Then Exit Do

        nbRempl = nbRempl + 1                                       ' Nombre de remplacements effectués.

        lPe1 = lPe2 + lCher                                         ' Ajuster les pointeurs de lecture (entrée).
        lPe2 = InStr(lPe1, sExpression, sCherche, iCompare)         ' Chercher l'occurrence suivante.

        If lNbRempl > -1 And nbRempl >= lNbRempl Then Exit Do       ' Nombre atteint.
    Loop Until lPe2 = 0                                             ' Plus d'autre occurrence

    ' Traiter la partie finale de la chaine, s'il en reste.
    If lPs <= Len(sResult) Then
        If (lExpr + 1 - lPe1 > 0) Then Mid$(sResult, lPs, lExpr + 1 - lPe1) = Mid$(sExpression, lPe1)
    End If

    Remplacer = sResult

#End If
End Function

' Scinde une chaine de caractères et retourne un tableau contenant les éléments individuels.
' Le fonctionnement est similaire à la fonction Split() de VB.
Function Scinder(sChaine As String, _
                 Optional sDelim As String = " ", _
                 Optional iNbFragments As Long = -1, _
                 Optional iCompare As Integer = VbCompareMethod.vbBinaryCompare) As Variant
#If VBA6 Then
    Scinder = Split(sChaine, sDelim, iNbFragments, iCompare)

#Else
    Dim i As Long, j As Long, l As Long
    Dim lDel As Integer, nbMax As Long, sResult() As String

    ' Si iNbFragment vaut zéro, la fonction retourne Empty
    If iNbFragments = 0 Then Exit Function

    lDel = Len(sDelim)                                              ' Longueur du délimiteur.
    If lDel = 0 Then                                                ' Délimiteur = ""
        ReDim sResult(0)
        sResult(0) = sChaine
        Scinder = sResult
        Exit Function
    End If

    ' Tableau
    If iNbFragments > 0 Then                                        ' Taille du tableau
        nbMax = iNbFragments - 1
    Else                                                            ' Compter...
        nbMax = 0
        i = InStr(1, sChaine, sDelim, iCompare)
        Do While i
            nbMax = nbMax + 1&
            i = InStr(i + lDel, sChaine, sDelim, iCompare)
        Loop
    End If

    ReDim sResult(nbMax)

    ' Traiter les n-1 premiers fragments.
    j = 1&: i = -1&
    For l = 0& To nbMax - 1&
        i = InStr(j, sChaine, sDelim, iCompare)                     ' Délimiteur suivant.
        ' On a demandé plus de fragments qu'il n'y a de délimiteurs dans la chaîne.
        If i = 0& Then Exit For

        sResult(l) = Mid$(sChaine, j, i - j)
        j = i + lDel
    Next l
    sResult(l) = Mid$(sChaine, j)                                   ' Ajouter le dernier fragment au tableau.
    If i = 0& Then ReDim Preserve sResult(l)                        ' Supprimer les réservations inutiles.

    Scinder = sResult()
    Erase sResult

#End If
End Function

' Cherche une chaine dans une autre en partant de la fin.
Function InStrFin(sChaine As String, sCherche As String, _
                  Optional lDebut As Long = -1, Optional iCompare As Integer = VbCompareMethod.vbBinaryCompare) As Long
#If VBA6 Then
    InStrFin = InStrRev(sChaine, sCherche, lDebut)

#Else
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

#End If
End Function

' Transforme un tableau en chaine délimitée.
' Equivalent à la fonction Join() de VB.
Function Joindre(sTableau As Variant, Optional sDelim As String = " ") As String
#If VBA6 Then
    Joindre = Join(sTableau, sDelim)

#Else
    Dim i As Long, j As Long, k As Long, l As Long, s As String, sResult As String

    For i = 0 To UBound(sTableau)
        l = l + Len(sTableau(i))
    Next i
    k = Len(sDelim)
    sResult = Space$(l + i * k)                                     ' Pré-allocation d'espace.
    j = 1                                                           ' Position d'écriture.

    For i = 0 To UBound(sTableau)
        s = sTableau(i)
        l = Len(s)
        Mid$(sResult, j, l) = s
        j = j + l

        Mid$(sResult, j, k) = sDelim
        j = j + k
    Next i

    Joindre = Left$(sResult, j - k - 1)                             ' Eliminer le dernier séparateur, inutile.

#End If
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