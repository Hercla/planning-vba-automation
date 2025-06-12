Attribute VB_Name = "AutoHoraire"
Option Explicit

Private Function EstCodeJourFerieOuRecup(code As String) As Boolean
    code = UCase(Trim(code))
    If code Like "F *" Or code Like "R *" Then
        EstCodeJourFerieOuRecup = True
    Else
        EstCodeJourFerieOuRecup = False
    End If
End Function

Sub AutoCategoriserEtColorerHoraires()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Liste")
    
    Dim i As Long, lastRow As Long
    Dim codeHoraire As String
    Dim heures As Variant
    Dim valMatin As Double, valAM As Double, valSoir As Double, valNuit As Double
    Dim skipCodes As Variant
    Dim data As Variant, results As Variant
    ' Liste complète des codes à ignorer
    skipCodes = Array("FP", "CEP", "CP", "DP", "ANC", "CA", "CTR", "EL", "C SOC", "FOR", "FSH", "MAL", _
                      "PETIT CHOM", "CSS", "DÉCÈS", "EM", "PAT", "PREAVIS", "VJ", "RCT", "RHS", "RV", _
                      "DÉMÉNAG", "GRÈVE", "F", "R", "RC", "RTT", "C", "CONG", "CONGE", _
                      "CRIC", "STAFF N", "RF", "H++")
    
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
    If lastRow < 2 Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Lire tous les codes horaires d'un coup
    data = ws.Range("A2:A" & lastRow).value
    ReDim results(1 To UBound(data), 1 To 4)
    
    For i = 1 To UBound(data)
        codeHoraire = Trim(data(i, 1))
        valMatin = 0: valAM = 0: valSoir = 0: valNuit = 0
        
        If codeHoraire = "" Or IsInArray(UCase(codeHoraire), skipCodes) Or EstCodeJourFerieOuRecup(codeHoraire) Then
            ' Rien à faire, tout reste à 0/"" (pas de cotation, pas de couleur)
        Else
            heures = ExtraireHeures(codeHoraire)
            Dim j As Long
            For j = LBound(heures) To UBound(heures) - 1 Step 2
                Dim hDeb As Double, hFin As Double
                hDeb = heures(j)
                hFin = heures(j + 1)
                If hFin < hDeb Then hFin = hFin + 24
                
                If hDeb < 12 And hFin > 6.75 Then
                    If hDeb <= 8 And hFin >= 12 Then
                        valMatin = 1
                    Else
                        valMatin = Application.Max(valMatin, 0.5)
                    End If
                End If
                If hDeb < 16.5 And hFin > 12 Then
                    If hDeb <= 12 And hFin >= 16.5 Then
                        valAM = 1
                    Else
                        valAM = Application.Max(valAM, 0.5)
                    End If
                End If
                ' Un poste est considéré "Soir" seulement s'il commence avant 19h
                If hDeb < 19 And hFin > 15.5 Then
                    If hDeb <= 16 And hFin >= 20 Then
                        valSoir = 1
                    Else
                        valSoir = Application.Max(valSoir, 0.5)
                    End If
                End If
                If (hDeb >= 19 Or hFin > 24 Or hFin <= 7) Or (hDeb >= 20 Or hDeb < 7) Then
                    valNuit = 1
                End If
            Next j
        End If
        results(i, 1) = valMatin
        results(i, 2) = valAM
        results(i, 3) = valSoir
        results(i, 4) = valNuit
    Next i
    
    ' Écriture en une seule opération
    ws.Range("C2:F" & lastRow).value = results
    
    ' Coloration optimisée
    ColorationOptimisee ws, lastRow
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "Auto-catégorisation et coloration terminées !", vbInformation
    AjouterLegendeHoraires
End Sub

Sub ColorationOptimisee(ws As Worksheet, lastRow As Long)
    Dim i As Long
    Dim val As Variant
    Dim colors(1 To 4, 1 To 2) As Long
    colors(1, 1) = RGB(255, 255, 153): colors(1, 2) = RGB(255, 255, 204) ' Matin
    colors(2, 1) = RGB(255, 204, 153): colors(2, 2) = RGB(255, 229, 204) ' AM
    colors(3, 1) = RGB(153, 204, 255): colors(3, 2) = RGB(204, 229, 255) ' Soir
    colors(4, 1) = RGB(204, 153, 255): colors(4, 2) = RGB(229, 204, 255) ' Nuit
    
    Dim col As Integer, cell As Range
    For col = 3 To 6
        For i = 2 To lastRow
            Set cell = ws.Cells(i, col)
            val = cell.value
            If val = 1 Then
                cell.Interior.Color = colors(col - 2, 1)
            ElseIf val = 0.5 Then
                cell.Interior.Color = colors(col - 2, 2)
            Else
                cell.Interior.ColorIndex = xlNone
            End If
        Next i
    Next col
End Sub

Sub AjouterLegendeHoraires()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Liste")
    Dim ligneLegende As Long
    ligneLegende = 1 ' Ligne 1 pour la légende

    ' Texte des légendes
    ws.Cells(ligneLegende, "C").value = "Matin"
    ws.Cells(ligneLegende, "D").value = "Après-midi"
    ws.Cells(ligneLegende, "E").value = "Soir"
    ws.Cells(ligneLegende, "F").value = "Nuit"

    ' Couleurs de la légende (mêmes que la macro principale)
    ws.Cells(ligneLegende, "C").Interior.Color = RGB(255, 255, 153)      ' Jaune Matin
    ws.Cells(ligneLegende, "D").Interior.Color = RGB(255, 204, 153)      ' Orange AM
    ws.Cells(ligneLegende, "E").Interior.Color = RGB(153, 204, 255)      ' Bleu Soir
    ws.Cells(ligneLegende, "F").Interior.Color = RGB(204, 153, 255)      ' Violet Nuit

    ws.Range("C1:F1").Font.Bold = True
    ws.Range("C1:F1").HorizontalAlignment = xlCenter

    ' Optionnel : Ajoute une info-bulle
    ws.Cells(ligneLegende, "G").value = "Légende : Couleurs = présence sur le créneau"
    ws.Cells(ligneLegende, "G").Font.Italic = True
End Sub

Function IsInArray(val As String, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If val = arr(i) Then IsInArray = True: Exit Function
    Next i
    IsInArray = False
End Function

Function ExtraireHeures(code As String) As Variant
    Dim t As Variant, h() As Double, i As Long, nb As Long
    code = Trim(code)
    If code = "" Then
        ReDim h(0 To 1)
        h(0) = 0
        h(1) = 0
        ExtraireHeures = h
        Exit Function
    End If

    t = Split(code, " ")
    nb = UBound(t) - LBound(t) + 1

    ' Nettoyage des entrées vides éventuelles
    Dim tempList As Collection
    Set tempList = New Collection
    For i = LBound(t) To UBound(t)
        If Trim(t(i)) <> "" Then tempList.Add t(i)
    Next i
    If tempList.Count = 0 Then
        ReDim h(0 To 1)
        h(0) = 0
        h(1) = 0
        ExtraireHeures = h
        Exit Function
    End If

    nb = tempList.Count
    If nb = 1 Then
        ReDim h(0 To 1)
        h(0) = ConvertirHeureTexte(tempList(1))
        h(1) = h(0)
    ElseIf nb Mod 2 <> 0 Then
        ReDim h(0 To nb)
        For i = 1 To nb - 1
            h(i - 1) = ConvertirHeureTexte(tempList(i))
        Next i
        h(nb - 1) = ConvertirHeureTexte(tempList(nb))
        h(nb) = h(nb - 1)
    Else
        ReDim h(0 To nb - 1)
        For i = 1 To nb
            h(i - 1) = ConvertirHeureTexte(tempList(i))
        Next i
    End If
    ExtraireHeures = h
End Function

Function ConvertirHeureTexte(h As String) As Double
    Dim hh As Double, mm As Double
    If InStr(h, ":") > 0 Then
        hh = val(Split(h, ":")(0))
        mm = val(Split(h, ":")(1))
    Else
        hh = val(h)
        mm = 0
    End If
    ConvertirHeureTexte = hh + mm / 60
End Function
