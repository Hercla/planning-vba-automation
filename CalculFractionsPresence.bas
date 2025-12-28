Attribute VB_Name = "CalculFractionsPresence"
Option Explicit

' =========================================================================================
'   MODULE DE CATÉGORISATION AUTOMATIQUE DES HORAIRES - VERSION FINALE DÉFINITIVE
'   Date de dernière mise à jour: 12 juin 2025
'
'   Description:
'   Ce module analyse les codes horaires et applique un ensemble de règles de
'   catégorisation très précises pour remplir les 13 colonnes d'analyse.
'   CORRECTION : Un demi-poste (0.5) requiert au moins 2h de présence dans le créneau.
' =========================================================================================

Sub AutoCategoriserEtColorerHoraires_Final()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Liste")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Erreur : La feuille nommée ""Liste"" n'a pas été trouvée.", vbCritical, "Feuille Manquante"
        Exit Sub
    End If

    Dim i As Long, lastRow As Long
    Dim codeHoraire As String
    Dim heures As Variant
    Dim data As Variant, results As Variant
    
    Dim valMatin As Double, valAM As Double, valSoir As Double, valNuit As Double
    Dim valP0645 As Long, valP7H8H As Long, valP8H1630 As Long
    Dim valC15 As Long, valC20 As Long, valC20E As Long, valC19 As Long
    Dim valN1945 As Double, valN20H7 As Double

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If lastRow < 2 Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    data = ws.Range("A2:A" & lastRow).value
    ReDim results(1 To UBound(data), 1 To 13)

    For i = 1 To UBound(data)
        codeHoraire = Trim(CStr(data(i, 1)))

        valMatin = 0: valAM = 0: valSoir = 0: valNuit = 0
        valP0645 = 0: valP7H8H = 0: valP8H1630 = 0
        valC15 = 0: valC20 = 0: valC20E = 0: valC19 = 0
        valN1945 = 0: valN20H7 = 0

        ' --- BLOC DE CONTRÔLE PRINCIPAL ---
        Dim isLeaveCode As Boolean
        isLeaveCode = False

        ' Vérification prioritaire si c'est un code de congé
        If UCase(codeHoraire) Like "F *" Or UCase(codeHoraire) Like "R *" Then
            isLeaveCode = True
        Else
            ' Liste des autres codes à ignorer
            Select Case UCase(codeHoraire)
                Case "WE", "ANC", "CA", "CEP", "CP", "CS", "CSS", "CTR", "DÉCÈS", "DÉMÉNAG", "DP", "EL", "EM", "FP", "GRÈVE", "PAT", "PREAVIS", "RCT", "RHS", "RV", "VJ", "C SOC", "FOR", "FSH", "MAL", "PETIT CHOM", "CRIC", "STAFF N", "RF", "H++"
                    isLeaveCode = True
            End Select
        End If
        
        ' Si ce n'est PAS un code de congé, alors on calcule
        If Not isLeaveCode And codeHoraire <> "" Then
            heures = ExtraireHeures(codeHoraire)
            
            If IsArray(heures) Then
                Dim j As Long, hDeb As Double, hFin As Double
                For j = LBound(heures) To UBound(heures) - 1 Step 2
                    hDeb = heures(j)
                    hFin = heures(j + 1)
                    If hFin <= hDeb Then hFin = hFin + 24
                    
                    ' --- *** NOUVELLE LOGIQUE: 2 HEURES MINIMUM POUR UN DEMI-POSTE *** ---
                    Dim overlap As Double
                    
                    ' Matin (Fenêtre de calcul 7h-12h)
                    overlap = Application.Max(0, Application.Min(hFin, 12) - Application.Max(hDeb, 7))
                    If (hDeb <= 8 And hFin >= 12) Then ' Règle du poste complet reste prioritaire
                        valMatin = 1
                    ElseIf overlap >= 2 Then ' Il faut au moins 2h de présence dans la fenêtre pour 0.5
                        valMatin = Application.Max(valMatin, 0.5)
                    End If
                    
                    ' Après-midi (Fenêtre de calcul 12h-17h)
                    overlap = Application.Max(0, Application.Min(hFin, 17) - Application.Max(hDeb, 12))
                    If (hDeb <= 13 And hFin >= 16.5) Then ' Règle du poste complet reste prioritaire
                        valAM = 1
                    ElseIf overlap >= 2 Then ' Il faut au moins 2h de présence dans la fenêtre pour 0.5
                        valAM = Application.Max(valAM, 0.5)
                    End If
                    
                    ' Soir (Fenêtre de calcul 17h-20.25h)
                    overlap = Application.Max(0, Application.Min(hFin, 20.25) - Application.Max(hDeb, 17))
                    If (hDeb < 17.5 And hFin >= 19) Then ' Règle du poste complet reste prioritaire
                        valSoir = 1
                    ElseIf overlap >= 2 Then ' Il faut au moins 2h de présence dans la fenêtre pour 0.5
                        valSoir = Application.Max(valSoir, 0.5)
                    End If
                    
                    ' Nuit
                    If hDeb >= 20 Or hFin > 24 Then
                        valNuit = 1
                    End If
                    
                    ' Présences spécifiques basées sur les heures
                    If hDeb = 6.75 Then valP0645 = 1
                    If hDeb >= 6.75 And hDeb < 8 Then valP7H8H = 1
                    If hDeb >= 8 And hDeb < 9 And hFin >= 16.5 Then valP8H1630 = 1
                Next j
            End If

            ' LOGIQUE BASÉE SUR LES CODES SPÉCIFIQUES DE TRAVAIL
            Select Case UCase(codeHoraire)
                Case "C 15", "C 15 SA", "C 15 DI", "16:30 20:15", "8:30 12:45 16:30 20:15": valC15 = 1
                Case "C 20", "8:30 12:30 16 20": valC20 = 1
                Case "C 20 E": valC20E = 1
                Case "C 19", "C 19 SA", "C 19 DI": valC19 = 1
                Case "19:45 6:45": valN1945 = 1: valNuit = 1
                Case "20 7": valN20H7 = 1: valNuit = 1
                Case "20 24": valN20H7 = 0.5: valNuit = 1
            End Select
            
            ' CORRECTIONS MANUELLES
            If UCase(codeHoraire) = "13:30 17:30" Then valSoir = 0
            If UCase(codeHoraire) = "8 18" Then valMatin = 1: valAM = 0.5: valSoir = 0.5
            If UCase(codeHoraire) = "9 18" Then valMatin = 0.5: valAM = 1: valSoir = 0
            If UCase(codeHoraire) = "6:45 20:30" Then valMatin = 1: valAM = 1: valSoir = 1
            
            ' Ré-appliquer les valeurs pour les codes "C"
            Select Case UCase(codeHoraire)
                Case "C 19", "C 19 SA", "C 19 DI", "C 20", "C 20 E", "C 15", "C 15 SA", "C 15 DI"
                    valMatin = 1: valAM = 0: valSoir = 1
            End Select
            
            ' Si c'est un poste de Nuit, on force la soirée à 0
            If valNuit = 1 And (UCase(codeHoraire) = "19:45 6:45" Or UCase(codeHoraire) = "20 7" Or UCase(codeHoraire) = "20 24") Then
                valSoir = 0
            End If
        End If

        results(i, 1) = valMatin: results(i, 2) = valAM: results(i, 3) = valSoir: results(i, 4) = valNuit
        results(i, 5) = valP0645: results(i, 6) = valP7H8H: results(i, 7) = valP8H1630
        results(i, 8) = valC15: results(i, 9) = valC20: results(i, 10) = valC20E: results(i, 11) = valC19
        results(i, 12) = valN1945: results(i, 13) = valN20H7
    Next i

    ws.Range("C2:O" & lastRow).value = results
    ColorationOptimisee ws, lastRow
    AjouterLegendeHoraires ws

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "Auto-catégorisation complète terminée !", vbInformation
End Sub


' =========================================================================================
'   FONCTIONS ET PROCÉDURES UTILITAIRES (DÉPENDANCES - INCHANGÉES)
' =========================================================================================
Private Function ExtraireHeures(code As String) As Variant
    On Error GoTo GestionErreur
    Dim rawParts() As String, cleanParts() As String, part As Variant
    Dim numCleanParts As Long, i As Long, result() As Double
    code = Replace(code, "-", " "): code = Application.WorksheetFunction.Trim(code)
    If code = "" Then GoTo GestionErreur
    rawParts = Split(code, " ")
    ReDim cleanParts(LBound(rawParts) To UBound(rawParts))
    numCleanParts = 0
    For Each part In rawParts
        If part <> "" And IsNumeric(Left(part, 1)) Then
            cleanParts(numCleanParts) = part
            numCleanParts = numCleanParts + 1
        End If
    Next part
    If numCleanParts = 0 Or numCleanParts Mod 2 <> 0 Then GoTo GestionErreur
    ReDim Preserve cleanParts(0 To numCleanParts - 1)
    ReDim result(1 To numCleanParts)
    For i = LBound(cleanParts) To UBound(cleanParts)
        result(i + 1) = ConvertTimeToDecimal(cleanParts(i))
    Next i
    ExtraireHeures = result
    Exit Function
GestionErreur:
    ExtraireHeures = False
End Function

Private Function ConvertTimeToDecimal(timeString As String) As Double
    Dim cleanString As String, timeParts() As String, i As Long
    cleanString = ""
    For i = 1 To Len(timeString)
        Dim char As String: char = Mid(timeString, i, 1)
        If IsNumeric(char) Or char = ":" Or char = "." Or char = "," Then
            cleanString = cleanString & char
        Else
            Exit For
        End If
    Next i
    cleanString = Replace(cleanString, ",", ".")
    If InStr(cleanString, ":") > 0 Then
        timeParts = Split(cleanString, ":")
        ConvertTimeToDecimal = val(timeParts(0)) + (val(timeParts(1)) / 60)
    Else
        ConvertTimeToDecimal = val(cleanString)
    End If
End Function

Sub ColorationOptimisee(ws As Worksheet, lastRow As Long)
    Dim i As Long, col As Integer, val As Variant, cell As Range
    Dim colors(1 To 4, 1 To 2) As Long
    colors(1, 1) = RGB(255, 255, 153): colors(1, 2) = RGB(255, 255, 204)
    colors(2, 1) = RGB(255, 204, 153): colors(2, 2) = RGB(255, 229, 204)
    colors(3, 1) = RGB(153, 204, 255): colors(3, 2) = RGB(204, 229, 255)
    colors(4, 1) = RGB(204, 153, 255): colors(4, 2) = RGB(229, 204, 255)
    ws.Range("C2:F" & lastRow).Interior.ColorIndex = xlNone
    For col = 3 To 6
        For i = 2 To lastRow
            Set cell = ws.Cells(i, col)
            val = cell.value
            If IsNumeric(val) And val > 0 Then
                If val = 1 Then
                    cell.Interior.Color = colors(col - 2, 1)
                ElseIf val = 0.5 Then
                    cell.Interior.Color = colors(col - 2, 2)
                End If
            End If
        Next i
    Next col
End Sub

Sub AjouterLegendeHoraires(ws As Worksheet)
    ws.Range("H:I").Clear
    ws.Range("R:T").Clear
    ws.Range("R1").value = "Légende : Couleurs = présence sur le créneau"
    Dim startCell As Range
    Set startCell = ws.Range("S2")
    startCell.value = "Légende des couleurs": startCell.Font.Bold = True
    startCell.offset(1, 0).value = "Matin (Poste)": startCell.offset(1, 1).Interior.Color = RGB(255, 255, 153)
    startCell.offset(2, 0).value = "Matin (Demi)": startCell.offset(2, 1).Interior.Color = RGB(255, 255, 204)
    startCell.offset(3, 0).value = "Après-midi (Poste)": startCell.offset(3, 1).Interior.Color = RGB(255, 204, 153)
    startCell.offset(4, 0).value = "Après-midi (Demi)": startCell.offset(4, 1).Interior.Color = RGB(255, 229, 204)
    startCell.offset(5, 0).value = "Soir (Poste)": startCell.offset(5, 1).Interior.Color = RGB(153, 204, 255)
    startCell.offset(6, 0).value = "Soir (Demi)": startCell.offset(6, 1).Interior.Color = RGB(204, 229, 255)
    startCell.offset(7, 0).value = "Nuit": startCell.offset(7, 1).Interior.Color = RGB(204, 153, 255)
    ws.Range("S:T").Columns.AutoFit
End Sub
