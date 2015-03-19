Option Compare Database
Option Explicit
Option Private Module

' Copyright 2014 Denis SCHEIDT
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






' Proc�dure appel�e par les m�thodes ChangeLang des formulaires.
'
' Cette proc�dure n'effectue PAS de remplacement de param�tres.
' Elle utilise une cl� de type NomFormulaire.NomControle.Propri�t� et ne peut traduire que les contr�les de formulaire.
Sub LangueCtls(frm As Form, sLangOrg() As String)
    Dim sForm As String, sCtl As Variant, rs As DAO.Recordset, s As String, i As Long, j As Long

    ' Initialiser la langue � utiliser.
    ' Si la langue n'est pas encore d�finie, se base sur la langue du syst�me d'exploitation.
    ' Utiliser la proc�dure ChangeLang() pour forcer une autre langue.
    If dtuEtatSyst.IDLang = 0 Then dtuEtatSyst.IDLang = LangueExiste(GetUserDefaultLangID)


    sForm = Mid$(frm.Name, InStr(frm.Name, "_") + 1)                        ' Eliminer frm_ du nom.

    ' Pr�paration du tableau.
    i = -1
    On Error Resume Next
    i = UBound(sLangOrg, 1)
    On Error GoTo 0
    If i = -1 Then                                                          ' Tableau non initialis�.
        ' Obtenir toutes les cl�s pour le formulaire.
        Set rs = CodeDb.OpenRecordset("SELECT DISTINCT CleMsg FROM T9N WHERE CleMsg Like '" & sForm & ".*'", _
                                       dbOpenDynaset, 0, dbReadOnly)
        If rs.RecordCount > 0 Then rs.MoveLast                              ' Remplir compl�tement le recordset.
        ' 0 : cl�
        ' 1 : Texte d'origine (FR)
        ReDim sLangOrg(1, rs.RecordCount - 1)

        ' Lire les cl�s dans le tableau.
        i = 0
        rs.MoveFirst
        On Error Resume Next
        Do While Not rs.EOF
            s = rs!CleMsg
            sLangOrg(0, i) = s                                              ' Stocker la cl�.

            sCtl = Scinder(s, ".")                                          ' Extraire les diff�rentes parties.
            ' Stocker le texte fran�ais (texte d'origine du contr�le).
            If sCtl(1) = "Caption" Then
                sLangOrg(1, i) = frm.Caption                                ' Caption du formulaire lui-m�me.
            ElseIf sCtl(2) = "Value" Then
                sLangOrg(1, i) = frm.Controls(sCtl(1))                      ' Value des TextBox.
            Else
                sLangOrg(1, i) = frm.Controls(sCtl(1)).Properties(sCtl(2))  ' Autres propri�t�s.
            End If

            rs.MoveNext
            i = i + 1
        Loop
        On Error GoTo 0
        rs.Close

    End If

    ' Chercher les traductions pour les contr�les du formulaire.
    Set rs = CodeDb.OpenRecordset("SELECT CleMsg, MsgT9N FROM T9N WHERE IDLang=" & dtuEtatSyst.IDLang & " AND CleMsg Like '" & sForm & ".*'", _
                                   dbOpenDynaset, 0, dbReadOnly)

    ' Si la langue exacte ne retourne rien, essayer sur le code langue principale.
    If rs.RecordCount = 0 Then
        rs.Close
        Set rs = CodeDb.OpenRecordset("SELECT CleMsg, MsgT9N FROM T9N WHERE IDLang Mod 1024=" & (dtuEtatSyst.IDLang And 1023) & " AND CleMsg Like '" & sForm & ".*'", _
                                       dbOpenDynaset, 0, dbReadOnly)
    End If

    ' Mettre � jour les contr�les avec la traduction, ou FR (depuis le tableau) si la traduction n'existe pas.
    For i = 0 To UBound(sLangOrg, 2)
        rs.FindFirst "CleMsg='" & sLangOrg(0, i) & "'"
        If rs.NoMatch Then
            s = sLangOrg(1, i)                                              ' Texte d'origine.
        Else
            s = TraiteChaine(rs!MsgT9N, Array())                            ' Traduction avec remplacement des caract�res sp�ciaux.
        End If

        sCtl = Scinder(sLangOrg(0, i), ".")                                 ' Extraire les diff�rentes parties de la cl�.

        ' Mettre � jour la propri�t� ad�quate du contr�le.
        On Error Resume Next
        If UBound(sCtl) = 2 Then
            If sCtl(2) = "Value" Then                                       ' Cas particulier des TextBox.
                frm.Controls(sCtl(1)) = s
            Else
                frm.Controls(sCtl(1)).Properties(sCtl(2)) = s               ' Traiter les autres propri�t�s.
            End If
        ElseIf sCtl(1) = "Caption" Then
            frm.Caption = s
        End If
        On Error GoTo 0
    Next i

    rs.Close
    Set rs = Nothing

End Sub

' Fournit la traduction d'un message dans la langue d�finie par la variable syst�me.
'
' Si la langue n'est pas d�finie, retour au fran�ais par d�faut.
'
' Liste des codes langue.
' http://msdn.microsoft.com/en-us/library/windows/desktop/dd318693%28v=vs.85%29.aspx

Function Traduit(ByVal sCle As String, sMsgFR As String, ParamArray Params() As Variant) As String
    Dim s As String, v As Variant, rs As DAO.Recordset, lLang As Long

    ' Initialiser la langue � utiliser.
    ' Si la langue n'est pas encore d�finie, se base sur la langue du syst�me d'exploitation.
    ' Utiliser la proc�dure ChangeLang() pour forcer une autre langue.
    If dtuEtatSyst.IDLang = 0 Then dtuEtatSyst.IDLang = LangueExiste(GetUserDefaultLangID)
    lLang = dtuEtatSyst.IDLang

    ' Le journal n'est �crit qu'en anglais ou fran�ais.
    If sCle Like "�*" Then
        sCle = Mid$(sCle, 2)                                                ' Supprimer le marqueur.
        If lLang <> 1036 Then lLang = 1033                                  ' Forcer l'anglais US si pas Fran�ais.
    End If


    Set rs = CodeDb.OpenRecordset("SELECT IDLang, CleMsg,MsgT9N FROM T9N WHERE CleMsg='" & sCle & "'", dbOpenDynaset, 0, dbReadOnly)

    ' Traduire le message. Chercher d'abord le code lange complet.
    rs.FindFirst "IDLAng=" & lLang
    If rs.NoMatch Then
        ' Si la recherche a �chou�, chercher sur le code langue principal.
        ' (on ignore la partie sublanguage)
        rs.FindFirst "IDLAng Mod 1024=" & (lLang And 1023)
    End If

    If rs.NoMatch Then
        s = sMsgFR                                                          ' Retour au Fran�ais par d�faut.
    Else
        s = rs!MsgT9N
    End If
    rs.Close
    Set rs = Nothing

    v = Params()                                                            ' Seule mani�re de passer correctement le ParamArray.
    Traduit = TraiteChaine(s, v)                                            ' Caract�res sp�ciaux et param�tres.
End Function

' Change la langue du menu.
Sub LangueMenu()
    Dim cbc As CommandBarControl

    For Each cbc In CommandBars("CB_libMAIL").Controls
        With cbc
            Select Case .Tag
                Case lmlMnuSspn:    .Caption = Traduit("mnu_pause", "&Suspendre")
                Case lmlMnuRlnc:    .Caption = Traduit("mnu_resume", "&Relancer")
                Case lmlMnuEnvM:    .Caption = Traduit("mnu_sendnow", "&Envoyer maintenant")
                Case lmlMnuDech:    .Caption = Traduit("mnu_unload", "&D�charger")
                Case lmlMnuAnnE:    .Caption = Traduit("mnu_cancel", "Ann&uler l'envoi")
                Case lmlMnuNMsg:    .Caption = Traduit("mnu_newmsg", "&Nouveau message...")
                Case lmlMnuGest:    .Caption = Traduit("mnu_mbm", "&Gestionnaire...")
                Case lmlMnuEtat:    .Caption = Traduit("mnu_status", "&Afficher l'�tat")
                Case lmlMnuAJnl:    .Caption = Traduit("mnu_log", "Afficher le &journal...")
                Case -2:            .Caption = Traduit("mnu_langue", "&Langue")
                Case -1:            .Caption = Traduit("mnu_about", "A &propos...")
            End Select
        End With
    Next cbc

    Set cbc = Nothing
End Sub


' D�termine la langue qui sera r�ellement utilis�e en fonction des langues existantes.
Function LangueExiste(lIDLang As Long) As Long
    Dim rs As DAO.Recordset

    If lIDLang = 1036 Then                                                  ' Vous avez demand� le fran�ais ? On quitte tout de suite.
        LangueExiste = 1036
        Exit Function
    End If

    ' Chercher tous les codes langues existant dans la table.
    Set rs = CodeDb.OpenRecordset("SELECT DISTINCT IDLang FROM T9N", dbOpenDynaset, 0, dbReadOnly)
    rs.FindFirst "IDLang=" & lIDLang
    If Not rs.NoMatch Then
        LangueExiste = lIDLang                                              ' Langue demand�e trouv�e.

    Else
        rs.FindFirst "IDLang Mod 1024=" & (lIDLang And 1023)                ' Chercher sur la langue principale.
        If Not rs.NoMatch Then
            LangueExiste = rs!IDLang                                        ' Il y a une correspondance sur la langue principale.

        Else
            rs.FindFirst "IDLang=1033"                                      ' Chercher l'anglais (1033).
            If Not rs.NoMatch Then
                LangueExiste = 1033                                         ' L'anglais existe.

            Else
                rs.FindFirst "IDLang Mod 1024=9"                            ' Chercher langue principale = anglais
                If Not rs.NoMatch Then
                    LangueExiste = rs!IDLang                                ' Il y en a une bas�e sur l'anglais.

                End If
            End If
        End If
    End If

    rs.Close
    Set rs = Nothing

    If LangueExiste = 0 Then LangueExiste = 1036                           ' Fran�ais si on ne trouve rien d'autre...

End Function



' Effectue le remplacement des caract�res sp�ciaux et des param�tres dans un message.
Private Function TraiteChaine(sChaine As String, Params As Variant) As String
    Dim s As String, i As Integer

    ' Traiter les caract�res sp�ciaux.
    s = Remplacer(sChaine, "\n", vbCrLf)                                    ' Retour chariot
    s = Remplacer(s, "\t", vbTab)                                           ' Tabulation


    ' Remplacement des param�tres, si n�cessaire. Seul %s est trait�.
    ' S'il y a au moins un param�tre � remplacer et si assez de param�tres ont �t� fournis.
    i = 0
    Do While i <= UBound(Params) And InStr(s, "%s") > 0
        s = Remplacer(s, "%s", CStr(Params(i)), , 1)
        i = i + 1                                                           ' Param�tre suivant.
    Loop

    TraiteChaine = s
End Function