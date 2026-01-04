' ExportedAt: 2026-01-04 17:02:17 | Workbook: Planning_2026.xlsm
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "PLANNING TEAM US 1D"
   ClientHeight    =   10185
   ClientLeft      =   -1155
   ClientTop       =   -5880
   ClientWidth     =   3703
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBox3_Change()
'Sub bouton_14()C 19
' Vérifier si la cellule active se trouve dans la plage de "planning" de l'onglet actif
    If Intersect(ActiveCell, ActiveSheet.Range("planning")) Is Nothing Then Exit Sub
    
    ' Écrire la valeur de la cellule W2 de l'onglet "Acceuil" dans la cellule active
    ActiveCell = Sheets("Acceuil").Range("W13").value
    
    ' Changer la couleur de fond de la cellule active en blanc et la couleur du texte en noir
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    
    ' Sélectionner la cellule à droite de la cellule active
    ActiveCell.offset(0, 1).Select
End Sub

Private Sub ComboBox4_Change()
End Sub
Private Sub CommandButton24_Click()
UserForm5.Show
End Sub

Private Sub CommandButton25_Click()
UserForm5.Show
End Sub

Private Sub CommandButton27_Click()
Sheets("AIDE").Activate
End Sub


Private Sub ComboBox1_Change()
End Sub

Private Sub CommandButton1_Click()
Sheets("PLANNING").Activate
Selection.AutoFilter Field:=2, Criteria1:="REEL"
End Sub

Private Sub CommandButton10_Click()
Sheets("Mai").Activate
Unload UserForm1
End Sub

Private Sub CommandButton11_Click()
Sheets("Juin").Activate
Unload UserForm1
End Sub

Private Sub CommandButton12_Click()
Sheets("Juillet").Activate
Unload UserForm1
End Sub

Private Sub CommandButton13_Click()
Sheets("Aout").Activate
Unload UserForm1
End Sub

Private Sub CommandButton14_Click()
Sheets("Sept").Activate
Unload UserForm1
End Sub

Private Sub CommandButton15_Click()
Sheets("Oct").Activate
Unload UserForm1
End Sub

Private Sub CommandButton16_Click()
Sheets("Nov").Activate
Unload UserForm1
End Sub

Private Sub CommandButton17_Click()
Sheets("Dec").Activate
Unload UserForm1
End Sub

Private Sub CommandButton2_Click()
Sheets("PLANNING").Activate
Selection.AutoFilter Field:=2, Criteria1:="PREV"
    Range("B29").Select
End Sub

Private Sub CommandButton21_Click()
End Sub

Private Sub ComboBox5_Change()
ActiveWindow.Zoom = 70
End Sub

Private Sub CommandButton100_Click()
'Sub bouton_22() 12:30 16:30
' Vérifier si la cellule active se trouve dans la plage de "planning" de l'onglet actif
    If Intersect(ActiveCell, ActiveSheet.Range("planning")) Is Nothing Then Exit Sub
    
    ' Écrire la valeur de la cellule W2 de l'onglet "Acceuil" dans la cellule active
    ActiveCell = Sheets("Acceuil").Range("W29").value
    
    ' Changer la couleur de fond de la cellule active en blanc et la couleur du texte en noir
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    
    ' Sélectionner la cellule à droite de la cellule active
    ActiveCell.offset(0, 1).Select
End Sub

Private Sub CommandButton101_Click()
'Sub bouton_01()'8:30 12:45 16:30 20:15
' Vérifier si la cellule active se trouve dans la plage de "planning" de l'onglet actif
    If Intersect(ActiveCell, ActiveSheet.Range("planning")) Is Nothing Then Exit Sub
    
    ' Écrire la valeur de la cellule W2 de l'onglet "Acceuil" dans la cellule active
    ActiveCell = Sheets("Acceuil").Range("W34").value
    
    ' Changer la couleur de fond de la cellule active en blanc et la couleur du texte en noir
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    
    ' Sélectionner la cellule à droite de la cellule active
    ActiveCell.offset(0, 1).Select
End Sub

Private Sub CommandButton102_Click()
'coller code CTR et l adapter par rapport au mois -1
Call InsererCodeCTR
End Sub

Private Sub CommandButton103_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "14 20"
End Sub

Private Sub CommandButton104_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "13 19"
End Sub

Private Sub CommandButton105_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "12:30 16:30"
End Sub

Private Sub CommandButton107_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "10 19"
End Sub

Private Sub CommandButton108_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "9 15:30"
End Sub

Private Sub CommandButton109_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "8 16:30"
End Sub

Private Sub CommandButton110_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "8 14"
End Sub

Private Sub CommandButton111_Click()
'place un astérix si besoin inf niveau du code
Call ToggleAsterisqueCellule
End Sub

Private Sub CommandButton112_Click()
'met la couleur vert foncé code hrel
Call ToggleColorierCelluleVertFonce
End Sub

Private Sub CommandButton113_Click()
    ' Vérifier si la cellule active se trouve dans la plage de "planning" de l'onglet actif
    If Intersect(ActiveCell, ActiveSheet.Range("planning")) Is Nothing Then Exit Sub
    
    ' Écrire la valeur de la cellule W34 de l'onglet "Acceuil" dans la cellule active
    ActiveCell = Sheets("Acceuil").Range("W34").value
    
    ' Changer la couleur de fond de la cellule active en blanc et la couleur du texte en noir
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    
    ' Définir la police de la cellule active en Arial Narrow, taille 8
    ActiveCell.Font.Name = "Arial Narrow"
    ActiveCell.Font.Size = 8
    
    ' Sélectionner la cellule à droite de la cellule active
    ActiveCell.offset(0, 1).Select
End Sub

Private Sub CommandButton114_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 15 di"
End Sub

Private Sub CommandButton115_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 15"
End Sub

Private Sub CommandButton116_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 20"
End Sub

Private Sub CommandButton117_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 19 di"
End Sub

Private Sub CommandButton118_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 19"
End Sub

Private Sub CommandButton119_Click()
' masquer toutes les lignes de "43:50" en premier lieu. Ensuite, la boucle parcourt chaque ligne et la
'rend visible seulement si elle est vide et que le nombre de lignes visibles est inférieur au nombre demandé.
Call AfficherMasquerLignesDynamiques

End Sub

Private Sub CommandButton120_Click()
    ' Afficher le UserForm pour la mise à jour des lignes
    Call MajLigne
End Sub


Private Sub CommandButton121_Click()
' Verify if the active cell is in the "planning" range of the active sheet
    If Not Intersect(ActiveCell, ActiveSheet.Range("planning")) Is Nothing Then
        ' Write "RHS 6h" into the active cell
        ActiveCell.value = "TV"

        ' Change the background color of the active cell to yellow and the text color to black
        ActiveCell.Interior.Color = RGB(255, 255, 0) ' Yellow
        ActiveCell.Font.Color = RGB(0, 0, 0) ' Black

        ' Select the cell to the right of the active cell
        ActiveCell.offset(0, 1).Select
    End If
End Sub

Private Sub CommandButton124_Click()
Call GenererDatesEtJoursPourTousLesMois

End Sub

Private Sub CommandButton125_Click()
'place une note
Call CreerNoteRemplacement
End Sub

Private Sub CommandButton126_Click()
Call GenererRoulementOptimise
End Sub

Private Sub CommandButton127_Click()
Call ImprimerPage1FeuilleActive

End Sub

Private Sub CommandButton128_Click()
Call CTR_CheckWeekendEligibility

End Sub

Private Sub CommandButton129_Click()
Call CheckDPMonthlyCodes

End Sub

Private Sub CommandButton130_Click()
Call UpdateMonthlySheets_Final_Polished

End Sub

Private Sub CommandButton131_Click()
Call Check_Presence_Infirmiers
End Sub

Private Sub CommandButton28_Click()
Sheets("HORAIRES").Activate
Range("A1:J1").Select
ActiveWindow.Zoom = True
Range("C5").Select
End Sub

Private Sub CommandButton29_Click()
Sheets("CYCLES").Activate
Range("A1:AT1").Select
ActiveWindow.Zoom = True
Range("C2").Select
End Sub

Private Sub CommandButton3_Click()
Sheets("PLANNING").Activate
 Selection.AutoFilter Field:=2
ActiveWindow.Zoom = 70
End Sub

Private Sub CommandButton4_Click()
Sheets("Acceuil").Activate
Unload UserForm1
End Sub

Private Sub CommandButton5_Click()
Sheets("AIDE").Activate
Unload UserForm1
End Sub

Private Sub CommandButton6_Click()
Sheets("Janv").Activate
Unload UserForm1
End Sub

Private Sub CommandButton7_Click()
Sheets("Fev").Activate
Unload UserForm1
End Sub

Private Sub CommandButton8_Click()
Sheets("Mars").Activate
Unload UserForm1
End Sub

Private Sub CommandButton9_Click()
Sheets("04").Activate
Unload UserForm1
End Sub


Private Sub CommandButton23_Click()
Unload UserForm1
End Sub
Private Sub CommandButton33_Click()
Unload UserForm1
End Sub

Private Sub CommandButton31_Click()
Call RAZPlanMens
End Sub

Private Sub CommandButton32_Click()
Call GenererRoulement8SemJusqu31Dec_DynamiqueNuit

End Sub

Private Sub CommandButton34_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Accueil"
End Sub
Private Sub CommandButton36_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Janv"
End Sub

Private Sub CommandButton37_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Fev"
End Sub

Private Sub CommandButton38_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Mars"
End Sub

Private Sub CommandButton39_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Avril"
End Sub

Private Sub CommandButton40_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Mai"
End Sub

Private Sub CommandButton41_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Juin"
End Sub

Private Sub CommandButton42_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Dec"
End Sub

Private Sub CommandButton43_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Nov"
End Sub

Private Sub CommandButton44_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Oct"
End Sub

Private Sub CommandButton45_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Sept"
End Sub

Private Sub CommandButton46_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Aout"
End Sub

Private Sub CommandButton47_Click()
' Appelle la macro publique qui se trouve dans Module_UserActions
    Module_UserActions.NavigateToSheet "Juillet"
End Sub
Private Sub CommandButton48_Click() ' Le bouton pour "6:45 12:45"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "6:45 12:45"
End Sub

Private Sub CommandButton49_Click() ' Le bouton pour "6:45 15:15"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "6:45 15:15"
End Sub
Private Sub CommandButton50_Click() ' Le bouton pour "7 15:30"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "7 15:30"
End Sub

Private Sub CommandButton52_Click() ' Le bouton pour "8 16:30"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "8 16:30"
End Sub

Private Sub CommandButton53_Click() ' Le bouton pour "7 13"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "7 13"
End Sub

Private Sub CommandButton55_Click() ' Le bouton pour "9 15:30"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "9 15:30"
End Sub

Private Sub CommandButton56_Click() ' Le bouton pour "10 20"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "10 20"
End Sub

Private Sub CommandButton57_Click() ' Le bouton pour "13 19"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "13 19"
End Sub

Private Sub CommandButton58_Click() ' Le bouton pour "14 20"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "14 20"
End Sub

Private Sub CommandButton60_Click() ' Le bouton pour "7:15 13:15"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "7:15 13:15"
End Sub
Private Sub CommandButton61_Click() ' Le bouton pour "C 19"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 19"
End Sub

Private Sub CommandButton62_Click() ' Le bouton pour "C 19 di"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 19 di"
End Sub

Private Sub CommandButton63_Click() ' Le bouton pour "C 20"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 20"
End Sub

Private Sub CommandButton64_Click() ' Le bouton pour "C 15"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 15"
End Sub

Private Sub CommandButton65_Click() ' Le bouton pour "C 15 di"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "C 15 di"
End Sub

Private Sub CommandButton66_Click() ' Le bouton pour "19:45 06:45"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "19:45 06:45"
End Sub

Private Sub CommandButton67_Click() ' Le bouton pour "20 7"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "20 7"
End Sub

Private Sub CommandButton69_Click() ' Le bouton pour "20 7"
'Sub bouton_Rhs 8h
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "20 7"
End Sub

Private Sub CommandButton70_Click() ' Le bouton pour "20 7"
'Sub bouton_CSOC
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "20 7"
End Sub

Private Sub CommandButton71_Click() ' Le bouton pour "20 7"
'Sub bouton_CSOC
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "20 7"
End Sub

Private Sub CommandButton72_Click() ' Le bouton pour "7:15 15:45"
'Sub bouton_21()7:15 15:45
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "7:15 15:45"
End Sub

Private Sub CommandButton73_Click() ' Le bouton pour "4/5*"
'Sub bouton_23()"4/5*"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "4/5*"
End Sub

Private Sub CommandButton74_Click() ' Le bouton pour "3/4*"
'Sub bouton_22()"3/4*"
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "3/4*"
End Sub

Private Sub CommandButton75_Click()
    ' Verify if the active cell is in the "planning" range of the active sheet
    If Intersect(ActiveCell, ActiveSheet.Range("planning")) Is Nothing Then Exit Sub
    
    ' Write the value from W32 of "Acceuil" sheet to the active cell
    With ActiveCell
        .value = Sheets("Acceuil").Range("W32").value
        ' Apply the exact color and text style
        .Interior.Color = RGB(0, 112, 192) ' Set background color to match the image
        .Font.Color = RGB(255, 255, 255)  ' Set text color to white
        .Font.Name = "Arial"
        .Font.Size = 12
    End With
    
    ' Select the cell to the right of the active cell
    ActiveCell.offset(0, 1).Select
End Sub

Private Sub CommandButton82_Click()
'Sub bouton_22() 7 11
InsertCodeAndMove "7 11"
End Sub

Private Sub CommandButton83_Click()
'Sub bouton_22() 8 14
InsertCodeAndMove "8 14"
End Sub

Private Sub CommandButton84_Click()
'Sub bouton_22() RHS 6h
InsertCodeAndMove "RHS 6h"
End Sub

Private Sub CommandButton85_Click()
' 'Sub bouton_22() RHS 6h
InsertCodeAndMove "RHS 6h"
End Sub

Private Sub CommandButton86_Click()
' 'Sub bouton_22() RHS 6h
InsertCodeAndMove "RHS 6h"
End Sub

Private Sub CommandButton87_Click()
Call ModulePDFGeneration.Generate_PDF_Jour
End Sub

Private Sub CommandButton88_Click()
Call ModulePDFGeneration.Generate_PDF_Nuit
End Sub

Private Sub CommandButton89_Click()
    Call Mode_Jour
End Sub
Sub AfficherMasquerLignes()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lignes As Range
    Set lignes = ws.Rows("43:44")
    
    If lignes.Hidden = True Then
        lignes.Hidden = False ' Affiche les lignes 43 et 44
    Else
        lignes.Hidden = True ' Masque les lignes 43 et 44
    End If
End Sub
Private Sub CommandButton90_Click()
Call ColorCells2
End Sub

Public Sub CommandButton91_Click()
Call Mode_Nuit
ActiveWindow.Zoom = 70 'Réglage de zoom
End Sub

Public Sub CommandButton92_Click()
Call Mode_Jour

ActiveWindow.Zoom = 70 'Réglage de zoom
End Sub

Private Sub CommandButton93_Click()
    ' Appelle la macro publique qui se trouve dans Module_Planning
    Call UpdateDailyTotals_V2
End Sub


Private Sub CommandButton94_Click()
' Call the PasteToPlanning macro
    PasteToPlanning
End Sub

Private Sub CommandButton96_Click()
' Call colorier cellule verte macro pr hrel
Call ToggleColorierCelluleVertFonce
End Sub

Private Sub CommandButton95_Click()
    FormulaireEntrees.Show
End Sub

Private Sub CommandButton97_Click()
' Call colorier cellule verte macro pr 7 15 30 asbd
Call ToggleColorierCelluleBleuClair
End Sub

Private Sub CommandButton98_Click()
'place un astérix si besoin inf niveau du code
Call ToggleAsterisqueCellule

End Sub

Private Sub CommandButton99_Click()
' On appelle la procédure publique et on lui passe le code
    Module_UserActions.InsertCodeFromUserForm "7:30 16"
End Sub

Private Sub UserForm_Click()
Me.width = 256.55 ' Ajustez la largeur selon vos besoins
End Sub
' --- AJOUTEZ CETTE PROCÉDURE DANS LE CODE DU USERFORM "Menu" ---
Private Sub InsertCodeAndMove(ByVal code As String)
    If Intersect(ActiveCell, Range("planning")) Is Nothing Then Exit Sub
    
    With ActiveCell
        .value = code
        .Interior.Color = vbWhite
        .Font.Color = vbBlack
        On Error Resume Next
        .offset(0, 1).Select
        On Error GoTo 0
    End With
End Sub
