Version = 17
VersionRequired = 17
Checksum = 677395606
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView = 0
    ScrollBars = 0
    PictureAlignment = 2
    DatasheetGridlinesBehavior = 3
    GridY = 10
    Width = 7086
    DatasheetFontHeight = 9
    ItemSuffix = 4
    Left = 1188
    Top = 3492
    Right = 9348
    Bottom = 8916
    DatasheetGridlinesColor = 12632256
    RecSrcDt = Begin
        0x5cf2efcc338be340
    End
    Caption ="Journal de connexion."
    DatasheetFontName ="Arial"
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
        Begin FormHeader
            Height = 0
            BackColor = -2147483633
            Name ="Ent�teFormulaire"
        End
        Begin Section
            Height = 3741
            BackColor = -2147483633
            Name ="D�tail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    ScrollBars = 2
                    OverlapFlags = 85
                    Left = 56
                    Top = 56
                    Width = 4769
                    Height = 3625
                    Name ="txtJournal"
                    FontName ="Arial"
                End
            End
        End
        Begin FormFooter
            Height = 396
            BackColor = -2147483633
            Name ="PiedFormulaire"
            Begin
                Begin CommandButton
                    OverlapFlags = 85
                    Left = 1474
                    Width = 1304
                    Height = 340
                    TabIndex = 1
                    Name ="cmdEfface"
                    Caption ="Effacer"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags = 85
                    Left = 56
                    Width = 1304
                    Height = 340
                    Name ="cmdActualiser"
                    Caption ="Actualiser"
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

' Copyright 2009-2012 Denis SCHEIDT
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



Private Sub cmdEfface_Click()
    Call SMTPJnlRAZ
    Call cmdActualiser_Click
End Sub

Private Sub Form_Load()
    Me.Visible = True
    Call cmdActualiser_Click
End Sub

Private Sub cmdActualiser_Click()
    With Me.txtJournal
        ' SelStart �tant un entier, on ne peut afficher que les 32767 derniers caract�res.
        .Value = Right$(SMTPJournal(), 32767)
        .SetFocus
        .SelStart = Len(.Value)
    End With
End Sub

Private Sub Form_Resize()
    Dim l As Single

    ' Largeur
    l = Me.InsideWidth - 2 * Me.txtJournal.Left
    If l <= 0 Then l = 0
    Me.txtJournal.Width = l
    Me.Width = l

    ' Hauteur
    l = Me.InsideHeight - Me.Ent�teFormulaire.Height - Me.PiedFormulaire.Height - 2 * Me.txtJournal.Top
    If l <= 240 Then l = 240
    Me.txtJournal.Height = l
    Me.D�tail.Height = l
End Sub