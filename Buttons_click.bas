Attribute VB_Name = "Buttons_click"
Sub Button_6_45_12_45_Click()
'Sub bouton_02()
    If Intersect(ActiveCell, Range("planning")) Is Nothing Then Exit Sub
    ActiveCell = Sheets("Acceuil").Range("W2").value '
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    ActiveCell.Offset(0, 1).Select
End Sub
Sub Button_6_45_15_15_Click()
'Sub bouton_02()
    If Intersect(ActiveCell, Range("planning")) Is Nothing Then Exit Sub
    ActiveCell = Sheets("Acceuil").Range("W3").value '
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    ActiveCell.Offset(0, 1).Select
End Sub
Sub Button_7_13_Click()
'Sub bouton_02()
    If Intersect(ActiveCell, Range("planning")) Is Nothing Then Exit Sub
    ActiveCell = Sheets("Acceuil").Range("W4").value '
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    ActiveCell.Offset(0, 1).Select
End Sub
Sub Button_7_15_13_15_Click()
'Sub bouton_02()
    If Intersect(ActiveCell, Range("planning")) Is Nothing Then Exit Sub
    ActiveCell = Sheets("Acceuil").Range("W5").value '
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    ActiveCell.Offset(0, 1).Select
End Sub
Sub Button_7_15_30_Click()
'Sub bouton_02()
    If Intersect(ActiveCell, Range("planning")) Is Nothing Then Exit Sub
    ActiveCell = Sheets("Acceuil").Range("W6").value '
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    ActiveCell.Offset(0, 1).Select
End Sub
Sub Button_7_15_15_45_Click()
'Sub bouton_02()
    If Intersect(ActiveCell, Range("planning")) Is Nothing Then Exit Sub
    ActiveCell = Sheets("Acceuil").Range("W7").value '
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    ActiveCell.Offset(0, 1).Select
End Sub
Sub Button_8_16_30_Click()
'Sub bouton_02()
    If Intersect(ActiveCell, Range("planning")) Is Nothing Then Exit Sub
    ActiveCell = Sheets("Acceuil").Range("W8").value '
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    ActiveCell.Offset(0, 1).Select
End Sub
Sub Button_9_15_30_Click()
'Sub bouton_02()
    If Intersect(ActiveCell, Range("planning")) Is Nothing Then Exit Sub
    ActiveCell = Sheets("Acceuil").Range("W9").value '
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    ActiveCell.Offset(0, 1).Select
End Sub
Sub Button_10_20_Click()
'Sub bouton_02()
    If Intersect(ActiveCell, Range("planning")) Is Nothing Then Exit Sub
    ActiveCell = Sheets("Acceuil").Range("W10").value '
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    ActiveCell.Offset(0, 1).Select
End Sub
Sub Button_13_19_Click()
'Sub bouton_02()
    If Intersect(ActiveCell, Range("planning")) Is Nothing Then Exit Sub
    ActiveCell = Sheets("Acceuil").Range("W11").value '
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    ActiveCell.Offset(0, 1).Select
End Sub
Sub Button_14_20_Click()
'Sub bouton_02()
    If Intersect(ActiveCell, Range("planning")) Is Nothing Then Exit Sub
    ActiveCell = Sheets("Acceuil").Range("W12").value '
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    ActiveCell.Offset(0, 1).Select
End Sub
Sub Button_C_19_Click()
'Sub bouton_02()
    If Intersect(ActiveCell, Range("planning")) Is Nothing Then Exit Sub
    ActiveCell = Sheets("Acceuil").Range("W13").value '
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    ActiveCell.Offset(0, 1).Select
End Sub
Sub Button_C_19_di_Click()
'Sub bouton_02()
    If Intersect(ActiveCell, Range("planning")) Is Nothing Then Exit Sub
    ActiveCell = Sheets("Acceuil").Range("W14").value '
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    ActiveCell.Offset(0, 1).Select
End Sub
Sub Button_C_20_Click()
'Sub bouton_02()
    If Intersect(ActiveCell, Range("planning")) Is Nothing Then Exit Sub
    ActiveCell = Sheets("Acceuil").Range("W15").value '
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    ActiveCell.Offset(0, 1).Select
End Sub
Sub Button_C_15_Click()
'Sub bouton_02()
    If Intersect(ActiveCell, Range("planning")) Is Nothing Then Exit Sub
    ActiveCell = Sheets("Acceuil").Range("W16").value '
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    ActiveCell.Offset(0, 1).Select
End Sub
Sub Button_C_15_di_Click()
'Sub bouton_02()
    If Intersect(ActiveCell, Range("planning")) Is Nothing Then Exit Sub
    ActiveCell = Sheets("Acceuil").Range("W17").value '
    ActiveCell.Interior.Color = RGB(255, 255, 255)
    ActiveCell.Font.Color = RGB(0, 0, 0)
    ActiveCell.Offset(0, 1).Select
End Sub
Private Sub CommandButton36_Click()
    'Janv
    Sheets("Janv").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
End Sub
Private Sub CommandButton37_Click()
    'Fev
    Sheets("Fev").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
End Sub
Private Sub CommandButton38_Click()
    'Mars
    Sheets("Mars").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
End Sub
Private Sub CommandButton39_Click()
    'Avril
    Sheets("Avril").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
End Sub
Private Sub CommandButton40_Click()
    'Mai
    Sheets("Mai").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
    
End Sub

Private Sub CommandButton41_Click()
    Sheets("Juin").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
    
End Sub
Private Sub CommandButton47_Click()
    'Juillet
    Sheets("Juillet").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
End Sub
Private Sub CommandButton77_Click()
    'Aout
    Sheets("Aout").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
End Sub
Private Sub CommandButton45_Click()
    'Septembre
    Sheets("Sept").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
    
End Sub
Private Sub CommandButton76_Click()
    'Octobre
    Sheets("Oct").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
End Sub

Private Sub CommandButton78_Click()
'Novembre
    Sheets("Nov").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
End Sub

Private Sub CommandButton79_Click()
'Décembre
    Sheets("Dec").Activate
    Range("B6").Select
    ActiveWindow.Zoom = 70
End Sub
Sub AfficherMasquerLignes()
    If ActiveSheet.CheckBoxes("CheckBox1").value = 1 Then
        rows("43:44").EntireRow.Hidden = False ' Affiche les lignes 43 et 44
    Else
        rows("43:44").EntireRow.Hidden = True ' Masque les lignes 43 et 44
    End If
End Sub




