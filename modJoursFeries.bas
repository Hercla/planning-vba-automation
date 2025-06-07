Attribute VB_Name = "modJoursFeries"
Option Explicit

' Détecte si le code horaire est un jour férié ou un récup (ex : F 1-1, R 1-1, etc.)
Function IsJourFerieOuRecup(code As String) As Boolean
    code = UCase(Trim(code))
    If code Like "F *" Or code Like "R *" Then
        IsJourFerieOuRecup = True
    Else
        IsJourFerieOuRecup = False
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
    skipCodes = Array("CA", "MAL", "EM", "CP", "CSS", "F", "R", "RC", "RTT", "C", "CONGÉ", "CONGE")
    
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
    Application.ScreenUpdating = False
    
    For i = 2 To lastRow
        codeHoraire = Trim(ws.Cells(i, "A").value)
        ' Ignore les lignes vides, codes de congé, fériés, récup
        If codeHoraire = "" Or IsInArray(UCase(codeHoraire), skipCodes) Or IsJourFerieOuRecup(codeHoraire) Then
            ws.Cells(i, "C").Resize(1, 4).value = ""
            EffaceCouleurs ws, i
            GoTo NextLine
        End If
        
        heures = ExtraireHeures(codeHoraire)
        valMatin = 0: valAM = 0: valSoir = 0: valNuit = 0
        
        Dim j As Long
        For j = LBound(heures) To UBound(heures) - 1 Step 2
            Dim hDeb As Double, hFin As Double
            hDeb = heures(j)
            hFin = heures(j + 1)
            If hFin < hDeb Then hFin = hFin + 24 ' Gestion nuit traversant minuit
            
            ' --- Matin (C) : 6h45 à 12h ---
            If hDeb < 12 And hFin > 6.75 Then
                If hDeb <= 8 And hFin >= 12 Then
                    valMatin = 1
                Else
                    valMatin = Application.Max(valMatin, 0.5)
                End If
            End If
            ' --- Après-midi (D) : 12h à 16h30 ---
            If hDeb < 16.5 And hFin > 12 Then
                If hDeb <= 12 And hFin >= 16.5 Then
                    valAM = 1
                Else
                    valAM = Application.Max(valAM, 0.5)
                End If
            End If
            ' --- Soir (E) : 15h30 à 20h15 ---
            If hDeb < 20.25 And hFin > 15.5 Then
                If hDeb <= 16 And hFin >= 20 Then
                    valSoir = 1
                Else
                    valSoir = Application.Max(valSoir, 0.5)
                End If
            End If
            ' --- Nuit (F) : >=19h ou traverse minuit jusqu'à 7h ---
            If (hDeb >= 19 Or hFin > 24 Or hFin <= 7) Or (hDeb >= 20 Or hDeb < 7) Then
                valNuit = 1
            End If
        Next j
        
        ws.Cells(i, "C").value = valMatin
        ws.Cells(i, "D").value = valAM
        ws.Cells(i, "E").value = valSoir
        ws.Cells(i, "F").value = valNuit
        
        ColorCellule ws.Cells(i, "C"), valMatin, RGB(255, 255, 153), RGB(255, 255, 204)
        ColorCellule ws.Cells(i, "D"), valAM, RGB(255, 204, 153), RGB(255, 229, 204)
        ColorCellule ws.Cells(i, "E"), valSoir, RGB(153, 204, 255), RGB(204, 229, 255)
        ColorCellule ws.Cells(i, "F"), valNuit, RGB(204, 153, 255), RGB(229, 204, 255)
NextLine:
    Next i
    Application.ScreenUpdating = True
    MsgBox "Auto-catégorisation et coloration terminées !", vbInformation
    AjouterLegendeHoraires
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

Sub EffaceCouleurs(ws As Worksheet, ligne As Long)
    Dim col As Integer
    For col = 3 To 6
        ws.Cells(ligne, col).Interior.ColorIndex = xlNone
    Next col
End Sub

Sub ColorCellule(cell As Range, val As Double, col1 As Long, col05 As Long)
    If val = 1 Then
        cell.Interior.Color = col1
    ElseIf val = 0.5 Then
        cell.Interior.Color = col05
    Else
        cell.Interior.ColorIndex = xlNone
    End If
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

Sub AjouterLegendeHor()
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
