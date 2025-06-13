Attribute VB_Name = "AutoHoraire"
Option Explicit

' =========================================================================================
'   MODULE DE CATÉGORISATION AUTOMATIQUE DES HORAIRES - VERSION FINALE
'   Date de dernière mise à jour: 12 juin 2025
'
'   Description:
'   Ce module analyse les codes horaires.
'   CORRECTION : Ignore correctement tous les codes de congé, y compris ceux
'   commençant par "F" et "R", pour qu'ils résultent en zéro.
' =========================================================================================

' =========================================================================================
'   PROCÉDURE PRINCIPALE : LANCE LA CATÉGORISATION ET LA COLORATION
' =========================================================================================
Sub AutoCategoriserEtColorerHoraires()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Liste")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Erreur : La feuille nommée ""Liste"" n'a pas été trouvée.", vbCritical, "Feuille Manquante"
        Exit Sub
    End If

    Dim i As Long, lastRow As Long
    Dim cellValue As Variant
    Dim codeHoraire As String
    Dim heures As Variant
    Dim valMatin As Double, valAM As Double, valSoir As Double, valNuit As Double
    Dim skipCodes As Variant
    Dim data As Variant, results As Variant
    Dim codeHandled As Boolean

    ' --- LISTE COMPLÈTE DES CODES À IGNORER ---
    skipCodes = Array("WE", "FP", "CEP", "CP", "DP", "ANC", "CA", "CTR", "EL", "C SOC", _
                      "FOR", "FSH", "MAL", "PETIT CHOM", "CSS", "DÉCÈS", "EM", "PAT", _
                      "PREAVIS", "VJ", "RCT", "RHS", "RV", "DÉMÉNAG", "GRÈVE", "F", "R", _
                      "RC", "RTT", "C", "CONG", "CONGE", "CRIC", "STAFF N", "RF", "H++")

    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
    If lastRow < 2 Then Exit Sub

    ' --- OPTIMISATION DE LA PERFORMANCE ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    data = ws.Range("A2:A" & lastRow).value
    ReDim results(1 To UBound(data), 1 To 4)

    ' --- BOUCLE PRINCIPALE DE TRAITEMENT ---
    For i = 1 To UBound(data)
        cellValue = data(i, 1)
        If IsNull(cellValue) Or IsError(cellValue) Then
            codeHoraire = ""
        Else
            codeHoraire = CStr(Trim(cellValue))
        End If

        valMatin = 0: valAM = 0: valSoir = 0: valNuit = 0
        codeHandled = False

        ' --- GESTION DES CODES COUPÉS SPÉCIAUX (lettres) ---
        Select Case UCase(codeHoraire)
            Case "C 19", "C 19 SA", "C 19 DI", "C 20", "C 20 E", "C 15", "C 15 SA", "C 15 DI"
                valMatin = 1: valSoir = 1: codeHandled = True
        End Select
        
        If codeHandled Then
            ' Le code a été géré
        ' --- *** CORRECTION MAJEURE ICI *** ---
        ' On vérifie si le code doit être ignoré (congé, F, R, etc.) AVANT d'essayer de le calculer.
        ElseIf codeHoraire = "" Or IsInArray(UCase(codeHoraire), skipCodes) Or EstCodeJourFerieOuRecup(codeHoraire) Then
            ' Code à ignorer, toutes les valeurs restent à 0
        Else
            ' --- EXTRACTION DES HORAIRES NUMÉRIQUES (uniquement pour les postes de travail) ---
            heures = ExtraireHeures(codeHoraire)
            If IsArray(heures) Then
                Dim j As Long
                For j = LBound(heures) To UBound(heures) - 1 Step 2
                    Dim hDeb As Double, hFin As Double
                    hDeb = heures(j)
                    hFin = heures(j + 1)
                    If hFin <= hDeb Then hFin = hFin + 24

                    ' LOGIQUE DE CATÉGORISATION
                    If valMatin < 1 Then
                        If (hDeb <= 9 And hFin >= 12.5) Or (hDeb <= 8 And hFin >= 11) Then valMatin = 1
                        ElseIf hDeb < 12 And hFin > 7 Then valMatin = Application.Max(valMatin, 0.5)
                    End If
                    If valAM < 1 Then
                        If (hDeb <= 12.5 And hFin >= 17) Or (hDeb <= 13.5 And hFin >= 18) Or (hDeb <= 13 And hFin >= 16.5) Then valAM = 1
                        ElseIf hDeb < 17.5 And hFin > 12 Then valAM = Application.Max(valAM, 0.5)
                    End If
                    If valSoir < 1 Then
                         If (hDeb <= 17 And hFin >= 20.25) Or hDeb >= 18 Then valSoir = 1
                        ElseIf hFin > 17 Then valSoir = Application.Max(valSoir, 0.5)
                    End If
                    If codeHoraire = "13:30 17:30" Then valSoir = 0
                    If hFin > 21 Or hDeb < 6 Then valNuit = 1
                Next j
            End If
        End If
        
        results(i, 1) = valMatin
        results(i, 2) = valAM
        results(i, 3) = valSoir
        results(i, 4) = valNuit
    Next i

    ' --- ÉCRITURE ET FINALISATION ---
    ws.Range("C2:F" & lastRow).value = results
    ColorationOptimisee ws, lastRow
    AjouterLegendeHoraires ws

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "Auto-catégorisation et coloration terminées !", vbInformation
End Sub


' =========================================================================================
'   FONCTIONS ET PROCÉDURES UTILITAIRES (DÉPENDANCES)
' =========================================================================================

' -----------------------------------------------------------------------------------------
'   DÉTERMINE SI UN CODE COMMENCE PAR "F " OU "R "
' -----------------------------------------------------------------------------------------
Private Function EstCodeJourFerieOuRecup(code As String) As Boolean
    code = UCase(Trim(code))
    If code Like "F *" Or code Like "R *" Then
        EstCodeJourFerieOuRecup = True
    Else
        EstCodeJourFerieOuRecup = False
    End If
End Function

' -----------------------------------------------------------------------------------------
'   VÉRIFIE SI UN ÉLÉMENT EXISTE DANS UN TABLEAU (ARRAY)
' -----------------------------------------------------------------------------------------
Public Function IsInArray(stringToFind As String, arr As Variant) As Boolean
    Dim element As Variant
    On Error Resume Next
    For Each element In arr
        If UCase(element) = UCase(stringToFind) Then
            IsInArray = True
            Exit Function
        End If
    Next element
    IsInArray = False
End Function

' -----------------------------------------------------------------------------------------
'   HELPER : CONVERTIT UNE CHAÎNE (ex: "16:30") EN HEURE DÉCIMALE (16.5)
' -----------------------------------------------------------------------------------------
Private Function ConvertTimeToDecimal(timeString As String) As Double
    Dim cleanString As String
    Dim timeParts() As String
    Dim i As Long
    cleanString = ""

    For i = 1 To Len(timeString)
        Dim char As String
        char = Mid(timeString, i, 1)
        If IsNumeric(char) Or char = ":" Or char = "." Or char = "," Then
            cleanString = cleanString & char
        Else
            Exit For
        End If
    Next i

    cleanString = Replace(cleanString, ",", ".")

    If InStr(cleanString, ":") > 0 Then
        timeParts = Split(cleanString, ":")
        ConvertTimeToDecimal = Val(timeParts(0)) + (Val(timeParts(1)) / 60)
    Else
        ConvertTimeToDecimal = Val(cleanString)
    End If
End Function

' -----------------------------------------------------------------------------------------
'   EXTRAIT LES HEURES D'UNE CHAÎNE EN TABLEAU NUMÉRIQUE
' -----------------------------------------------------------------------------------------
Public Function ExtraireHeures(code As String) As Variant
    On Error GoTo GestionErreur
    
    Dim rawParts() As String
    Dim cleanParts() As String
    Dim part As Variant
    Dim numCleanParts As Long
    Dim i As Long
    Dim result() As Double

    code = Replace(code, "-", " ")
    code = Application.WorksheetFunction.Trim(code)
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

    If numCleanParts = 0 Or numCleanParts Mod 2 <> 0 Then
        GoTo GestionErreur
    End If
    
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

' -----------------------------------------------------------------------------------------
'   APPLIQUE LES COULEURS DE FOND AUX CELLULES
' -----------------------------------------------------------------------------------------
Sub ColorationOptimisee(ws As Worksheet, lastRow As Long)
    Dim i As Long, col As Integer
    Dim val As Variant
    Dim cell As Range
    Dim colors(1 To 4, 1 To 2) As Long
    colors(1, 1) = RGB(255, 255, 153): colors(1, 2) = RGB(255, 255, 204) ' Matin
    colors(2, 1) = RGB(255, 204, 153): colors(2, 2) = RGB(255, 229, 204) ' Après-midi
    colors(3, 1) = RGB(153, 204, 255): colors(3, 2) = RGB(204, 229, 255) ' Soir
    colors(4, 1) = RGB(204, 153, 255): colors(4, 2) = RGB(229, 204, 255) ' Nuit
    
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

' -----------------------------------------------------------------------------------------
'   AJOUTE UNE LÉGENDE DES COULEURS
' -----------------------------------------------------------------------------------------
Public Sub AjouterLegendeHoraires(ws As Worksheet)
    Dim startCell As Range
    Set startCell = ws.Range("H2")
    
    startCell.Resize(12, 2).Clear
    startCell.value = "Légende des couleurs"
    startCell.Font.Bold = True
    
    startCell.Offset(2, 0).value = "Matin (Poste)": startCell.Offset(2, 1).Interior.Color = RGB(255, 255, 153)
    startCell.Offset(3, 0).value = "Matin (Demi)": startCell.Offset(3, 1).Interior.Color = RGB(255, 255, 204)
    startCell.Offset(4, 0).value = "Après-midi (Poste)": startCell.Offset(4, 1).Interior.Color = RGB(255, 204, 153)
    startCell.Offset(5, 0).value = "Après-midi (Demi)": startCell.Offset(5, 1).Interior.Color = RGB(255, 229, 204)
    startCell.Offset(6, 0).value = "Soir (Poste)": startCell.Offset(6, 1).Interior.Color = RGB(153, 204, 255)
    startCell.Offset(7, 0).value = "Soir (Demi)": startCell.Offset(7, 1).Interior.Color = RGB(204, 229, 255)
    startCell.Offset(8, 0).value = "Nuit": startCell.Offset(8, 1).Interior.Color = RGB(204, 153, 255)
    
    startCell.EntireColumn.AutoFit
End Sub

