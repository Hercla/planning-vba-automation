Attribute VB_Name = "AnalyseEtRemplacementPlanning"
Option Explicit

'-------------------------------
' FONCTIONS UTILITAIRES
'-------------------------------

Function IsJourFerieOuRecup(code As String) As Boolean
    Dim joursFeries As Variant
    joursFeries = Array("F 1-1", "F 8-5", "F 14-7", "F 15-8", "F 1-11", "F 11-11", "F 25-12", "R 8-5", "R 1-1")
    IsJourFerieOuRecup = IsInArray(code, joursFeries)
End Function

Function IsInArray(val As String, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If StrComp(val, arr(i), vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

Function ContientCodePlanningJourArr(planningArr As Variant, jourCol As Long, codeList As Variant) As Boolean
    Dim row As Long, code As String, i As Long
    
    If UBound(planningArr, 1) < 1 Or UBound(planningArr, 2) < jourCol Then
        ContientCodePlanningJourArr = False
        Exit Function
    End If

    For row = 1 To UBound(planningArr, 1)
        code = Trim(planningArr(row, jourCol))
        If code <> "" Then
            For i = LBound(codeList) To UBound(codeList)
                If (Left(codeList(i), 5) = "6:45" And Left(code, 5) = "6:45") Or _
                   (StrComp(code, codeList(i), vbTextCompare) = 0) Then
                    ContientCodePlanningJourArr = True
                    Exit Function
                End If
            Next i
        End If
    Next row
    ContientCodePlanningJourArr = False
End Function

Function EstUnOngletDeMois(nomFeuille As String) As Boolean
    EstUnOngletDeMois = (nomFeuille Like "Janv*" Or nomFeuille Like "Fev*" Or nomFeuille Like "Mars*" Or _
                         nomFeuille Like "Avril*" Or nomFeuille Like "Mai*" Or nomFeuille Like "Juin*" Or _
                         nomFeuille Like "Juillet*" Or nomFeuille Like "Aout*" Or nomFeuille Like "Sept*" Or _
                         nomFeuille Like "Oct*" Or nomFeuille Like "Nov*" Or nomFeuille Like "Dec*")
End Function

'-------------------------------
' MACRO PRINCIPALE
'-------------------------------

Sub AnalyseEtRemplacementPlanningUltraOptimise()
    Dim ws As Worksheet
    Dim LdebFractions As Long: LdebFractions = 64
    Dim LfinFractions As Long: LfinFractions = 70
    Dim colDeb As Long: colDeb = 2
    Dim groupesExclusifs As Variant
    Dim codesSuggestion As Variant
    Dim choixUtilisateur As VbMsgBoxResult

    groupesExclusifs = Array( _
        Array("6:45 15:15", "6:45 12:45"), _
        Array("C 15", "C 15 di"), _
        Array("C 20", "C 20 E"), _
        Array("C 19", "C 19 di"), _
        Array("7:15 15:45", "7:30 16", "8:30 16:30", "8 14", "8:30 14") _
    )

    codesSuggestion = Array( _
        Array("6:45 15:15", "6:45 12:45"), _
        Array("7 15:30"), _
        Array("8 16:30", "8:30 16:30"), _
        Array("C 15", "C 15 di"), _
        Array("C 20", "C 20 E"), _
        Array("C 20 E", "C 20"), _
        Array("C 19", "C 19 di"), _
        Array("19:45 6:45"), _
        Array("20 7"), _
        Array("20 7", "19:45 6:45") _
    )

    choixUtilisateur = MsgBox("Analyser uniquement l'onglet actif (" & ActiveSheet.Name & ") ?" & vbCrLf & _
                              "Oui pour actif, Non pour tous.", vbYesNoCancel + vbQuestion, "Choix de l'analyse")

    If choixUtilisateur = vbCancel Then
        MsgBox "Opération annulée.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    If choixUtilisateur = vbYes Then
        Set ws = ActiveSheet
        If EstUnOngletDeMois(ws.Name) Then
            TraiterUneFeuilleDeMois ws, LdebFractions, LfinFractions, colDeb, groupesExclusifs, codesSuggestion
            MsgBox "Analyse terminée : " & ws.Name, vbInformation
        Else
            MsgBox "Onglet invalide.", vbExclamation
        End If
    Else
        For Each ws In Worksheets
            If EstUnOngletDeMois(ws.Name) Then
                TraiterUneFeuilleDeMois ws, LdebFractions, LfinFractions, colDeb, groupesExclusifs, codesSuggestion
            End If
        Next ws
        MsgBox "Analyse globale terminée.", vbInformation
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


