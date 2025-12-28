Attribute VB_Name = "Module_Configuration"
Option Explicit

'MISE À JOUR DE LA CONFIGURATION
'Objectif : À lancer une fois par an, ou si vous ajoutez/modifiez des codes.
'Génère les codes pour les jours fériés de l'année en cours dans A2:A21.
'Calcule la valeur en heures décimales (colonne R) et en format HH:MM (colonne S) pour tous les codes de la feuille.
'Nom de la macro à lancer : MettreAJourConfigurationCodes

' =====================================================================================
' CONSTANTES DU PROJET
' =====================================================================================
Private Const SHEET_NAME As String = "Config_Codes"
Private Const COL_CODE As String = "A"
Private Const COL_TYPE_CODE As String = "C"
Private Const COL_HEURE_DECIMALE As String = "R"
Private Const COL_HEURE_NORMALE As String = "S"
Private Const COL_FORCE_DECIMALE As String = "T"
Private Const COL_FORCE_NORMALE As String = "U"

' Noms des plages nommées (peut rester pour compatibilité avec d'anciens codes)
Private Const NAMED_RANGE_CONGES As String = "ListeCodesConges"
Private Const NAMED_RANGE_COUPES8H As String = "ListeCodesCoupes8h"


' =====================================================================================
'   PROCÉDURE PRINCIPALE À LANCER
' =====================================================================================
Public Sub MettreAJourConfigurationCodes()
    Dim ws As Worksheet
    
    ' --- Étape 1: Initialisation de la feuille ---
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    If ws Is Nothing Then
        MsgBox "Erreur: La feuille '" & SHEET_NAME & "' n'a pas été trouvée.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' --- Étape 2: Mise à jour des jours fériés dans la colonne A ---
    MettreAJourFeries ws
    
    ' --- Étape 3: Calcul et remplissage des heures pour tous les codes ---
    CalculerHeuresPourTousLesCodes ws
    
    MsgBox "Les jours fériés ont été mis à jour et toutes les heures ont été calculées avec succès.", vbInformation
End Sub


' =====================================================================================
'   SOUS-TÂCHES ET FONCTIONS UTILITAIRES
' =====================================================================================

' Met à jour les codes pour les jours fériés dans la colonne A.
Private Sub MettreAJourFeries(ByVal ws As Worksheet)
    Dim annee As Integer
    Dim jourFeries(1 To 10) As Date
    Dim i As Long
    Dim data(1 To 20, 1 To 1) As Variant
    
    annee = Year(Date)
    
    jourFeries(1) = DateSerial(annee, 1, 1): jourFeries(2) = CalculerPaques(annee) + 1
    jourFeries(3) = DateSerial(annee, 5, 1): jourFeries(4) = CalculerPaques(annee) + 39
    jourFeries(5) = CalculerPaques(annee) + 50: jourFeries(6) = DateSerial(annee, 7, 21)
    jourFeries(7) = DateSerial(annee, 8, 15): jourFeries(8) = DateSerial(annee, 11, 1)
    jourFeries(9) = DateSerial(annee, 11, 11): jourFeries(10) = DateSerial(annee, 12, 25)
    
    For i = 1 To 10
        data(i, 1) = "F " & Format(jourFeries(i), "d-m")
        data(i + 10, 1) = "R " & Format(jourFeries(i), "d-m")
    Next i
    
    ws.Range(COL_CODE & "2").Resize(20, 1).value = data
End Sub

' Calcule les heures pour tous les codes de la feuille.
Private Sub CalculerHeuresPourTousLesCodes(ByVal ws As Worksheet)
    Dim lastRow As Long, i As Long
    Dim code As String, typeCode As String
    Dim heureDigitale As Double, heureNormale As String
    Dim pieces() As String
    
    lastRow = ws.Cells(ws.Rows.Count, COL_CODE).End(xlUp).row

    For i = 2 To lastRow
        code = Trim(ws.Cells(i, COL_CODE).value)
        typeCode = ws.Cells(i, COL_TYPE_CODE).value
        heureDigitale = 0: heureNormale = ""
        ws.Cells(i, COL_CODE).Interior.Color = xlNone

        ' Cas prioritaire : Valeur manuelle forcée
        If ws.Cells(i, COL_FORCE_DECIMALE).value <> "" Then
            heureDigitale = val(ws.Cells(i, COL_FORCE_DECIMALE).value)
            If ws.Cells(i, COL_FORCE_NORMALE).value <> "" Then
                heureNormale = CStr(ws.Cells(i, COL_FORCE_NORMALE).value)
            Else
                heureNormale = FormatHeure(heureDigitale)
            End If
        
        ' Cas spécial pour les codes temps partiel ou sans solde
        ElseIf typeCode = "SansSolde" Then
            heureDigitale = 0
            
        ' Cas des Fériés et Récupérations
        ElseIf Left(code, 2) = "F " Then
            heureDigitale = 7.6
        ElseIf Left(code, 2) = "R " Then
            heureDigitale = 8

        ' Cas des codes horaires
        ElseIf InStr(code, " ") > 0 And IsNumeric(Left(code, 1)) Then
            pieces = Split(code, " ")
            Select Case UBound(pieces)
                Case 1 ' Format "debut fin"
                    heureDigitale = CalculerDuree(pieces(0), pieces(1))
                    If heureDigitale >= 8 Then heureDigitale = heureDigitale - 0.5
                Case 3 ' Format "debut1 fin1 debut2 fin2"
                    heureDigitale = CalculerDuree(pieces(0), pieces(1)) + CalculerDuree(pieces(2), pieces(3))
                Case Else
                    heureDigitale = 0
            End Select
        
        ' Cas des codes simples (CA, MAL, etc.)
        Else
            ' Logique par défaut : 7.6h pour les absences, 8h pour les récup, etc.
             Select Case typeCode
                Case "Congé", "Maladie", "Férié", "Famille", "Exceptionnel"
                    heureDigitale = 7.6
                Case "Recup"
                    heureDigitale = 8
                Case Else
                    heureDigitale = 7.6 ' Valeur par défaut pour les autres codes
            End Select
        End If

        heureNormale = FormatHeure(heureDigitale)
        ws.Cells(i, COL_HEURE_DECIMALE).value = heureDigitale
        ws.Cells(i, COL_HEURE_NORMALE).value = heureNormale
    Next i
End Sub

Private Function CalculerPaques(ByVal annee As Integer) As Date
    Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer, f As Integer, g As Integer, h As Integer, i As Integer, k As Integer, l As Integer, m As Integer, mois As Integer, jour As Integer
    a = annee Mod 19: b = annee \ 100: c = annee Mod 100: d = b \ 4: e = b Mod 4: f = (b + 8) \ 25: g = (b - f + 1) \ 3: h = (19 * a + b - d - g + 15) Mod 30: i = c \ 4: k = c Mod 4: l = (32 + 2 * e + 2 * i - h - k) Mod 7: m = (a + 11 * h + 22 * l) \ 451: mois = (h + l - 7 * m + 114) \ 31: jour = ((h + l - 7 * m + 114) Mod 31) + 1
    CalculerPaques = DateSerial(annee, mois, jour)
End Function

Private Function FormatHeure(ByVal heuresDecimales As Double) As String
    Dim heures As Long, minutes As Long
    heures = Int(heuresDecimales)
    minutes = Round((heuresDecimales - heures) * 60, 0)
    FormatHeure = Format(heures, "0") & ":" & Format(minutes, "00")
End Function

Private Function CalculerDuree(ByVal heureDebut As String, ByVal heureFin As String) As Double
    Dim debutDecimal As Double: debutDecimal = ConvertirTexteEnHeure(heureDebut)
    Dim finDecimal As Double: finDecimal = ConvertirTexteEnHeure(heureFin)
    If finDecimal < debutDecimal Then finDecimal = finDecimal + 24
    CalculerDuree = finDecimal - debutDecimal
End Function

Private Function ConvertirTexteEnHeure(ByVal texteHeure As String) As Double
    Dim t As Variant
    texteHeure = Replace(Trim(texteHeure), ",", ".")
    If InStr(texteHeure, ":") > 0 Then
        t = Split(texteHeure, ":")
        ConvertirTexteEnHeure = val(t(0)) + val(t(1)) / 60
    Else
        ConvertirTexteEnHeure = val(texteHeure)
    End If
End Function
