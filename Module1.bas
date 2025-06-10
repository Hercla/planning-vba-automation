Attribute VB_Name = "Module1"
Option Explicit

' --- MODULE LEVEL CONSTANTS FOR ROW NUMBERS ---
Public Enum RowIdx
    rowDebutPlanningPersonnel = 6
    rowFinPlanningPersonnel = 30
    rowAideSoignantC19 = 24
    rowRemplacementDebutJour = 40
    rowRemplacementFinJour = 44
    rowRemplacementDebutNuit = 46
    rowRemplacementFinNuit = 47
End Enum

Public Const LIGNE_DEBUT_PLANNING_PERSONNEL As Long = RowIdx.rowDebutPlanningPersonnel
Public Const LIGNE_FIN_PLANNING_PERSONNEL As Long = RowIdx.rowFinPlanningPersonnel
Public Const LIGNE_AIDE_SOIGNANT_C19_PLANNING As Long = RowIdx.rowAideSoignantC19
Public Const LIGNE_REMPLACEMENT_DEBUT_JOUR As Long = RowIdx.rowRemplacementDebutJour
Public Const LIGNE_REMPLACEMENT_FIN_JOUR As Long = RowIdx.rowRemplacementFinJour
Public Const NB_REMPLACEMENT_JOUR_LIGNES As Long = LIGNE_REMPLACEMENT_FIN_JOUR - LIGNE_REMPLACEMENT_DEBUT_JOUR + 1
Public Const LIGNE_REMPLACEMENT_DEBUT_NUIT As Long = RowIdx.rowRemplacementDebutNuit
Public Const LIGNE_REMPLACEMENT_FIN_NUIT As Long = RowIdx.rowRemplacementFinNuit
Public Const NB_REMPLACEMENT_NUIT_LIGNES As Long = LIGNE_REMPLACEMENT_FIN_NUIT - LIGNE_REMPLACEMENT_DEBUT_NUIT + 1

' ============================
' CONSTANTES INDICES (utilises pour codesSuggestionPM et autres arrays)
' ============================
Public Enum SuggestionIndex
    SUGG_645 = 0
    SUGG_7_1530
    SUGG_7_1130
    SUGG_7_13
    SUGG_8_1630
    SUGG_C15_GRP
    SUGG_C20_CODE
    SUGG_C20E_CODE
    SUGG_C19_CODE
    SUGG_12_30_16_30
    SUGG_NUIT1
    SUGG_NUIT2
End Enum

' Module-level string constants for target array data.
' Each string represents a week (Mon-Sun) separated by semicolons.
' Each day's data is two values (non-holiday, holiday) separated by commas.
Const TARGET_MATIN_DATA As String = "7,5;7,5;7,5;7,5;7,5;5,5;5,5"
Const TARGET_PM_DATA As String = "4,2;3,2;3,2;4,2;4,2;2,2;2,2"
Const TARGET_SOIR_DATA As String = "3,3;3,3;3,3;3,3;3,3;3,3;3,3"

' ============================
' FONCTIONS UTILITAIRES
' ============================


Function ContientCodeDuGroupe(planningArr As Variant, rempContextArr As Variant, jourCol As Long, groupCodes As Variant) As Boolean
    Dim code As Variant
    ContientCodeDuGroupe = False ' Initialize
    If Not IsArray(groupCodes) Then Exit Function
    If LBound(groupCodes) > UBound(groupCodes) Then Exit Function ' Handle empty groupCodes array

    For Each code In groupCodes
        If ModuleUtils.CodeDejaPresent(planningArr, rempContextArr, jourCol, CStr(code), True) Then
            ContientCodeDuGroupe = True
            Exit Function
        End If
    Next code
End Function

Function EstUnOngletDeMois(nomFeuille As String) As Boolean
    EstUnOngletDeMois = ( _
        nomFeuille Like "Janv*" Or nomFeuille Like "Fev*" Or nomFeuille Like "Mars*" Or _
        nomFeuille Like "Avril*" Or nomFeuille Like "Mai*" Or nomFeuille Like "Juin*" Or _
        nomFeuille Like "Juillet*" Or nomFeuille Like "Aout*" Or nomFeuille Like "Sept*" Or _
        nomFeuille Like "Oct*" Or nomFeuille Like "Nov*" Or nomFeuille Like "Dec*" _
    )
End Function

' Helper sub to count frequencies of codes in a single array and update the dictionary
Private Sub CountFrequenciesInSingleArray(arrToCount As Variant, ByVal col As Long, ByRef freq As Object)
    Dim r As Long, cellVal As String
    If IsArray(arrToCount) Then
        If LBound(arrToCount, 1) <= UBound(arrToCount, 1) And LBound(arrToCount, 2) <= UBound(arrToCount, 2) Then
            If col >= LBound(arrToCount, 2) And col <= UBound(arrToCount, 2) Then
                For r = LBound(arrToCount, 1) To UBound(arrToCount, 1)
                    On Error Resume Next
                    cellVal = Trim(CStr(arrToCount(r, col)))
                    On Error GoTo 0
                    If freq.Exists(cellVal) Then freq(cellVal) = freq(cellVal) + 1
                Next r
            End If
        End If
    End If
End Sub

Function ChoisirCodePertinent(codesPossibles As Variant, planningArr As Variant, rempArr As Variant, col As Long) As String
    Dim code As Variant, freq As Object
    Set freq = CreateObject("Scripting.Dictionary")

    ' Ensure codesPossibles is a usable array
    If Not IsArray(codesPossibles) Then
        If IsMissing(codesPossibles) Or IsEmpty(codesPossibles) Or IsNull(codesPossibles) Or codesPossibles = "" Then
            ChoisirCodePertinent = ""
            Exit Function
        Else
            codesPossibles = Array(codesPossibles) ' Convert single item to array
        End If
    ElseIf LBound(codesPossibles) > UBound(codesPossibles) Then
        ChoisirCodePertinent = "" ' Empty array provided
        Exit Function
    End If
    
    ' Initialize frequency for all possible codes
    For Each code In codesPossibles
        freq(CStr(code)) = 0
    Next code
    
    ' Count frequencies using the helper sub
    Call CountFrequenciesInSingleArray(planningArr, col, freq)
    Call CountFrequenciesInSingleArray(rempArr, col, freq)

    ' First, try to pick a code that is not present at all (frequency 0)
    For Each code In codesPossibles
        If freq(CStr(code)) = 0 Then ChoisirCodePertinent = CStr(code): Exit Function
    Next code
    
    ' If all codes are present at least once, pick the one with the minimum frequency
    Dim minFreq As Long: minFreq = -1 ' Initialize to ensure first code's freq is picked
    Dim bestCode As String: bestCode = ""
    
    If LBound(codesPossibles) <= UBound(codesPossibles) Then
        bestCode = CStr(codesPossibles(LBound(codesPossibles))) ' Initialize with the first possible code
        minFreq = freq(bestCode)
    Else
        ChoisirCodePertinent = "" ' Should not be reached if initial checks are correct
        Exit Function
    End If

    For Each code In codesPossibles
        If freq(CStr(code)) < minFreq Then
            minFreq = freq(CStr(code))
            bestCode = CStr(code)
        End If
    Next code
    ChoisirCodePertinent = bestCode
End Function

' Cette fonction a été remplacée par la version Private Sub MettreAJourCompteursMAS plus bas dans le code
' qui gère correctement tous les cas de codes coupés et les règles métier associées

Function CreateTargetArray(dataString As String) As Variant
    ' Helper to create a 2D array from a string representation of targets.
    ' dataString format: "val1,val2;val3,val4;..." where each val is for (non-holiday, holiday)
    Dim rows As Variant, cols As Variant
    Dim r As Long, c As Long
    Dim tempArr As Variant

    rows = Split(dataString, ";")
    If UBound(rows) < LBound(rows) Then Exit Function

    ' Redimension array: 0 to 6 for days (Mon-Sun), 0 to 1 for holiday (0=No, 1=Yes)
    ReDim tempArr(0 To UBound(rows) - LBound(rows), 0 To 1)

    For r = LBound(rows) To UBound(rows)
        cols = Split(rows(r), ",")
        If UBound(cols) >= LBound(cols) + 1 Then ' Expecting two values (non-holiday, holiday)
            On Error Resume Next ' Use On Error Resume Next carefully for conversion
            tempArr(r - LBound(rows), 0) = CLng(cols(LBound(cols)))
            tempArr(r - LBound(rows), 1) = CLng(cols(LBound(cols) + 1))
            If Err.Number <> 0 Then ' If conversion fails, set to 0 and clear error
                tempArr(r - LBound(rows), 0) = 0
                tempArr(r - LBound(rows), 1) = 0
                Err.Clear
            End If
            On Error GoTo 0 ' Reset error handling
        Else ' Handle malformed data: default to 0
            tempArr(r - LBound(rows), 0) = 0
            tempArr(r - LBound(rows), 1) = 0
        End If
    Next r
    CreateTargetArray = tempArr
End Function

Private Sub ActualiserManquesValeurs(ByRef manqueMatin As Long, ByRef manquePM As Long, ByRef manqueSoir As Long, _
                                   ByVal targetMatin As Long, ByVal targetPM As Long, ByVal targetSoir As Long, _
                                   ByVal actualMatin As Long, ByVal actualPM As Long, ByVal actualSoir As Long)
    manqueMatin = targetMatin - actualMatin: If manqueMatin < 0 Then manqueMatin = 0
    manquePM = targetPM - actualPM: If manquePM < 0 Then manquePM = 0
    manqueSoir = targetSoir - actualSoir: If manqueSoir < 0 Then manqueSoir = 0
End Sub

' Initializes the static target staffing arrays.
Private Sub InitialiserTableauxCiblesStatiques(ByRef arrTargetMatin As Variant, ByRef arrTargetPM As Variant, ByRef arrTargetSoir As Variant)
    If Not IsEmpty(arrTargetMatin) Then Exit Sub ' Already initialized

    arrTargetMatin = CreateTargetArray(TARGET_MATIN_DATA)
    arrTargetPM = CreateTargetArray(TARGET_PM_DATA)
    arrTargetSoir = CreateTargetArray(TARGET_SOIR_DATA)
End Sub

' Tries to place a code if valid, updates counters and needs. Returns True if placed.
Private Function TryPlaceCodeIfValid(ByVal codeToPlace As String, _
                                     ByRef rempArrayToUpdate As Variant, ByVal targetRowInRemp As Long, ByVal targetCol As Long, _
                                     planningArr As Variant, _
                                     ByRef currentMatin As Long, ByRef currentPM As Long, ByRef currentSoir As Long, _
                                     ByVal absoluteRowForMAS As Long, ByRef currentPresence7_8h As Long, _
                                     ByRef neededMatin As Long, ByRef neededPM As Long, ByRef neededSoir As Long, _
                                     targetMatinVal As Long, targetPMVal As Long, targetSoirVal As Long, _
                                     Optional exactMatchCheck As Boolean = True) As Boolean
    TryPlaceCodeIfValid = False
    If codeToPlace = "" Then Exit Function

    If Not ModuleUtils.CodeDejaPresent(planningArr, rempArrayToUpdate, targetCol, codeToPlace, exactMatchCheck) Then
        rempArrayToUpdate(targetRowInRemp, targetCol) = codeToPlace
        ' Appel avec les paramètres dans le bon ordre pour la nouvelle signature
        Call MettreAJourCompteursMAS(codeToPlace, currentMatin, currentPM, currentSoir, absoluteRowForMAS, currentPresence7_8h)
        Call ActualiserManquesValeurs(neededMatin, neededPM, neededSoir, targetMatinVal, targetPMVal, targetSoirVal, currentMatin, currentPM, currentSoir)
        TryPlaceCodeIfValid = True
    End If
End Function

' Fonction spéciale pour traiter le cas du 5 septembre
Private Function EstLe5Septembre(ws As Worksheet, col As Long, colDeb As Long) As Boolean
    Dim dateCell As Variant
    EstLe5Septembre = False
    
    On Error Resume Next
    dateCell = ws.Cells(4, col + colDeb - 1).value
    If Err.Number = 0 And IsDate(dateCell) Then
        If Day(dateCell) = 5 And Month(dateCell) = 9 Then
            EstLe5Septembre = True
        End If
    End If
    On Error GoTo 0
End Function

' Fonction pour déterminer si un jour est vendredi, samedi ou férié
Private Function EstJourVendrediSamediOuFerie(ByVal jourSemaine As Long, ByVal codeFerie As Boolean) As Boolean
    EstJourVendrediSamediOuFerie = (jourSemaine = 5 Or jourSemaine = 6 Or codeFerie) ' Friday (5), Saturday (6)
End Function

' Fonction pour détecter les conditions spéciales (jeudi, vendredi, lundi avec fraction 5 2 2)
Private Function EstJourSpecialAvecFraction522(ByVal jourSemaine As Long, ByVal actualMatin As Long, ByVal actualPM As Long, ByVal actualSoir As Long) As Boolean
    EstJourSpecialAvecFraction522 = False
    ' Check if it's a Monday (1), Thursday (4) or Friday (5)
    If jourSemaine = 1 Or jourSemaine = 4 Or jourSemaine = 5 Then
        ' Check if current fractions match 5 2 2 (This is the actual count, not target)
        If actualMatin = 5 And actualPM = 2 And actualSoir = 2 Then
            EstJourSpecialAvecFraction522 = True
        End If
    End If
End Function

' Fonction pour compter le nombre de codes manquants dans le planning principal (empty slots)
Private Function CompterCodesManquants(planningArr As Variant, col As Long) As Long
    Dim nbCodesPresents As Long, r As Long
    
    If IsArray(planningArr) Then
        If col >= LBound(planningArr, 2) And col <= UBound(planningArr, 2) Then
            For r = LBound(planningArr, 1) To UBound(planningArr, 1)
                If Trim(CStr(planningArr(r, col))) <> "" Then
                    nbCodesPresents = nbCodesPresents + 1
                End If
            Next r
        End If
    End If
    
    ' Determine the number of potential slots to fill (up to 5 for now)
    ' This logic assumes a maximum of 5 general positions to fill if they are empty.
    CompterCodesManquants = Application.Min(5, UBound(planningArr, 1) - nbCodesPresents)
End Function

' Fonction pour mettre à jour les compteurs matin, après-midi et soir en fonction du code placé
Sub MettreAJourCompteursMAS(ByVal code As String, _
                       ByRef currentMatin As Long, ByRef currentPM As Long, ByRef currentSoir As Long, _
                       ByVal rowForMAS As Long, ByRef currentPresence7_8h As Long)
    ' Traitement spécial pour certains codes coupés
    If code = "C 19" Then ' 7h-14h30 et 17h30-20h15
        currentMatin = currentMatin + 1
        currentSoir = currentSoir + 1
        ' Ne pas incrémenter currentPM car ce code ne couvre pas l'après-midi
    ElseIf code = "C 20" Then ' 7h-14h et 17h30-21h15
        currentMatin = currentMatin + 1
        currentSoir = currentSoir + 1
        ' Ne pas incrémenter currentPM car ce code ne couvre pas l'après-midi
    ElseIf code = "C 20 E" Then ' 7h-14h et 17h30-21h15 (variante)
        currentMatin = currentMatin + 1
        currentSoir = currentSoir + 1
        ' Ne pas incrémenter currentPM car ce code ne couvre pas l'après-midi
    ElseIf code = "C 15" Then ' 7h-14h et 14h-21h
        currentMatin = currentMatin + 1
        currentPM = currentPM + 1
        currentSoir = currentSoir + 1
    ElseIf code = "8:30 12:45 16:30 20:15" Then ' Code spécial
        currentMatin = currentMatin + 1
        currentPM = currentPM + 1
        currentSoir = currentSoir + 1
    Else
        ' Pour les autres codes, on vérifie la présence matin, après-midi et soir
        ' en fonction de la position dans le planning (rowForMAS)
        If rowForMAS >= LIGNE_DEBUT_PLANNING_PERSONNEL Then
            ' Codes standards dans le planning principal
            Select Case code
                Case "M" ' Matin uniquement
                    currentMatin = currentMatin + 1
                    currentPresence7_8h = currentPresence7_8h + 1
                Case "S" ' Soir uniquement
                    currentSoir = currentSoir + 1
                Case "A" ' Après-midi uniquement
                    currentPM = currentPM + 1
                Case "J" ' Journée complète (matin + après-midi)
                    currentMatin = currentMatin + 1
                    currentPM = currentPM + 1
                    currentPresence7_8h = currentPresence7_8h + 1
                Case "MS" ' Matin + Soir
                    currentMatin = currentMatin + 1
                    currentSoir = currentSoir + 1
                    currentPresence7_8h = currentPresence7_8h + 1
                Case "AS" ' Après-midi + Soir
                    currentPM = currentPM + 1
                    currentSoir = currentSoir + 1
                Case "MAS", "JAS" ' Matin + Après-midi + Soir ou Journée + Après-midi + Soir
                    currentMatin = currentMatin + 1
                    currentPM = currentPM + 1
                    currentSoir = currentSoir + 1
                    currentPresence7_8h = currentPresence7_8h + 1
            End Select
        End If
    End If
End Sub

' --- NOTE IMPORTANTE ---
' La fonction Private CodeDejaPresent a été supprimée ici car elle était dupliquée.
' Utiliser plutôt la version publique définie aux lignes 86-98 qui utilise la fonction utilitaire
' CheckSingleArrayForCode pour une meilleure modularité.
'
' Cette suppression résout l'erreur de compilation "Nom ambigu détecté: CodeDejaPresent".

' Fonction pour compter les occurrences des codes spécifiques dans les remplacements
Private Sub CompterCodesSpecifiques(rempArr As Variant, col As Long, _
                                  ByRef nbC19 As Long, ByRef nbC20 As Long, ByRef nbC20E As Long, ByRef nbC15 As Long)
    Dim r As Long
    
    nbC19 = 0: nbC20 = 0: nbC20E = 0: nbC15 = 0
    
    If IsArray(rempArr) Then
        For r = LBound(rempArr, 1) To UBound(rempArr, 1)
            If col >= LBound(rempArr, 2) And col <= UBound(rempArr, 2) Then
                Select Case Trim(CStr(rempArr(r, col)))
                    Case "C 19": nbC19 = nbC19 + 1
                    Case "C 20": nbC20 = nbC20 + 1
                    Case "C 20 E": nbC20E = nbC20E + 1
                    Case "C 15": nbC15 = nbC15 + 1
                End Select
            End If
        Next r
    End If
End Sub

' Fonction pour déterminer le code coupé approprié en fonction des codes déjà présents
' Refactored to take counts as parameters instead of re-scanning arrays.
Private Function DeterminerCodeCoupe(ByVal nbC19Actuels As Long, ByVal nbC20Actuels As Long, _
                                   ByVal nbC20EActuels As Long, ByVal nbC15Actuels As Long, _
                                   ByVal estJourSpecial As Boolean) As String
    Dim codeToSuggest As String
    
    If estJourSpecial Then
        ' For Friday, Saturday, Holidays, we want a specific rotation of C19, C20, C20E
        If nbC19Actuels = 0 Then
            codeToSuggest = "C 19"
        ElseIf nbC20Actuels = 0 Then
            codeToSuggest = "C 20"
        ElseIf nbC20EActuels = 0 Then
            codeToSuggest = "C 20 E"
        Else
            ' If all are present, choose the least frequent among them
            If nbC19Actuels <= nbC20Actuels And nbC19Actuels <= nbC20EActuels Then
                codeToSuggest = "C 19"
            ElseIf nbC20Actuels <= nbC19Actuels And nbC20Actuels <= nbC20EActuels Then
                codeToSuggest = "C 20"
            Else
                codeToSuggest = "C 20 E"
            End If
        End If
    Else
        ' For other days, standard rotation
        If nbC19Actuels = 0 Then
            codeToSuggest = "C 19"
        ElseIf nbC20Actuels = 0 Then
            codeToSuggest = "C 20"
        ElseIf nbC20EActuels = 0 Then
            codeToSuggest = "C 20 E"
        ElseIf nbC15Actuels = 0 Then
            codeToSuggest = "C 15"
        Else
            ' If all are present, choose the least frequent
            If nbC19Actuels <= nbC20Actuels And nbC19Actuels <= nbC20EActuels And nbC19Actuels <= nbC15Actuels Then
                codeToSuggest = "C 19"
            ElseIf nbC20Actuels <= nbC19Actuels And nbC20Actuels <= nbC20EActuels And nbC20Actuels <= nbC15Actuels Then
                codeToSuggest = "C 20"
            ElseIf nbC20EActuels <= nbC19Actuels And nbC20EActuels <= nbC20Actuels And nbC20EActuels <= nbC15Actuels Then
                codeToSuggest = "C 20 E"
            Else
                codeToSuggest = "C 15"
            End If
        End If
    End If
    DeterminerCodeCoupe = codeToSuggest
End Function
Private Sub LireDonneesFeuille(ws As Worksheet, LdebFractions As Long, LfinFractions As Long, colDeb As Long, nbJours As Long, _
                              planningArr As Variant, fractionsArr As Variant, rempJourArr As Variant, rempNuitArr As Variant, _
                              dateArr As Variant, ferieArr As Variant)
    On Error Resume Next
    nbJours = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column - (colDeb - 1)
    If Err.Number <> 0 Then nbJours = 0
    On Error GoTo 0
    Dim daysInMonth As Long
    On Error Resume Next
    daysInMonth = Day(DateSerial(Year(ws.Cells(1, 1).Value), Month(ws.Cells(1, 1).Value) + 1, 0))
    If Err.Number <> 0 Then daysInMonth = 31
    On Error GoTo 0
    If nbJours > daysInMonth Then nbJours = daysInMonth
    planningArr = ws.Range(ws.Cells(LIGNE_DEBUT_PLANNING_PERSONNEL, colDeb), ws.Cells(LIGNE_FIN_PLANNING_PERSONNEL, colDeb + nbJours - 1)).Value2
    If LdebFractions > 0 And LfinFractions >= LdebFractions Then
        fractionsArr = ws.Range(ws.Cells(LdebFractions, colDeb), ws.Cells(LfinFractions, colDeb + nbJours - 1)).Value2
    End If
    rempJourArr = ws.Range(ws.Cells(LIGNE_REMPLACEMENT_DEBUT_JOUR, colDeb), ws.Cells(LIGNE_REMPLACEMENT_FIN_JOUR, colDeb + nbJours - 1)).Value2
    rempNuitArr = ws.Range(ws.Cells(LIGNE_REMPLACEMENT_DEBUT_NUIT, colDeb), ws.Cells(LIGNE_REMPLACEMENT_FIN_NUIT, colDeb + nbJours - 1)).Value2
    dateArr = ws.Range(ws.Cells(4, colDeb), ws.Cells(4, colDeb + nbJours - 1)).Value2
    ferieArr = ws.Range(ws.Cells(5, colDeb), ws.Cells(5, colDeb + nbJours - 1)).Value2
End Sub

Sub TraiterUneFeuilleDeMois(ws As Worksheet, _
                            LdebFractions As Long, LfinFractions As Long, _
                            colDeb As Long, _
                            codesSuggestion As Variant)
    Dim col As Long, jourSemaine As Long, i As Long, l As Long
    Dim dateJour As Date, codeFerie As Boolean
    Dim nbJours As Long
    Dim planningArr As Variant, fractionsArr As Variant
    Dim rempJourArr As Variant, rempNuitArr As Variant
    Dim dateArr As Variant, ferieArr As Variant
    Dim newlyPlaced_presence7_8h As Long ' Counter for newly placed codes affecting 7-8h presence

    Call LireDonneesFeuille(ws, LdebFractions, LfinFractions, colDeb, nbJours, planningArr, fractionsArr, rempJourArr, rempNuitArr, dateArr, ferieArr)
    If nbJours <= 0 Then Exit Sub

    ' Ensure replacement arrays are correctly dimensioned if they were read from empty ranges
    If Not IsArray(rempJourArr) Then
        ReDim rempJourArr(1 To NB_REMPLACEMENT_JOUR_LIGNES, 1 To nbJours)
    Else
        ' Vérifier les dimensions du tableau rempJourArr et les corriger si nécessaire
        If LBound(rempJourArr, 1) <> 1 Or UBound(rempJourArr, 1) <> NB_REMPLACEMENT_JOUR_LIGNES Or _
           LBound(rempJourArr, 2) <> 1 Or UBound(rempJourArr, 2) <> nbJours Then
            ' Créer un tableau temporaire avec les bonnes dimensions
            Dim tempArr As Variant
            ReDim tempArr(1 To NB_REMPLACEMENT_JOUR_LIGNES, 1 To nbJours)
            
            ' Copier les données existantes si possible
            Dim j As Long ' i est déjà déclaré au niveau de la fonction
            For i = 1 To NB_REMPLACEMENT_JOUR_LIGNES
                For j = 1 To nbJours
                    If i <= UBound(rempJourArr, 1) - LBound(rempJourArr, 1) + 1 And _
                       j <= UBound(rempJourArr, 2) - LBound(rempJourArr, 2) + 1 Then
                        tempArr(i, j) = rempJourArr(LBound(rempJourArr, 1) + i - 1, LBound(rempJourArr, 2) + j - 1)
                    End If
                Next j
            Next i
            
            ' Remplacer rempJourArr par le tableau temporaire
            rempJourArr = tempArr
        End If
    End If
    
    If Not IsArray(rempNuitArr) Then ReDim rempNuitArr(1 To NB_REMPLACEMENT_NUIT_LIGNES, 1 To nbJours)

    ' Initialize static target staffing arrays (done only once per full macro run)
    Static arrTargetMatin As Variant, arrTargetPM As Variant, arrTargetSoir As Variant
    Call InitialiserTableauxCiblesStatiques(arrTargetMatin, arrTargetPM, arrTargetSoir)

    ' Load initial staff counts from rows 60-62
    Dim initialEffectifsMatin As Variant, initialEffectifsPM As Variant, initialEffectifsSoir As Variant
    initialEffectifsMatin = ws.Range(ws.Cells(60, colDeb), ws.Cells(60, colDeb + nbJours - 1)).Value2
    initialEffectifsPM = ws.Range(ws.Cells(61, colDeb), ws.Cells(61, colDeb + nbJours - 1)).Value2
    initialEffectifsSoir = ws.Range(ws.Cells(62, colDeb), ws.Cells(62, colDeb + nbJours - 1)).Value2

    ' --- Main loop through each day (column) ---
    For col = 1 To nbJours ' Array column index (1 to nbJours)
        newlyPlaced_presence7_8h = 0 ' Reset for the current day
        
        ' Traitement spécial pour le 5 septembre
        If EstLe5Septembre(ws, col, colDeb) Then
            ' Force les codes spécifiques pour le 5 septembre
            rempJourArr(1, col) = "7 11:30"  ' Ligne 40
            rempJourArr(2, col) = "7 15:30"  ' Ligne 41
            rempJourArr(3, col) = "C 20 E"   ' Ligne 42
            
            ' Mise à jour des compteurs pour ces codes
            ' Utilisation de nouveaux noms pour éviter le conflit avec actualMatin/PM/Soir qui sont déclarés plus bas
            Dim cibleMatinSpeciale As Long: cibleMatinSpeciale = 7  ' Fraction cible pour le matin
            Dim ciblePMSpeciale As Long: ciblePMSpeciale = 4     ' Fraction cible pour l'après-midi
            Dim cibleSoirSpeciale As Long: cibleSoirSpeciale = 3    ' Fraction cible pour le soir
            
            ' Les valeurs seront utilisées pour le jour suivant via GoTo
            
            ' Passer au jour suivant
            GoTo JourSuivant
        End If

        ' Determine day of week and if it's a holiday
        If IsDate(dateArr(1, col)) Then
            dateJour = CDate(dateArr(1, col))
            jourSemaine = Weekday(dateJour, vbMonday) ' Monday = 1, ..., Sunday = 7
        Else
            jourSemaine = ((ws.Cells(4, col + colDeb - 1).Column - colDeb) Mod 7) + 1 ' Fallback
            Debug.Print "Warning: Invalid date in sheet " & ws.Name & ", column " & (col + colDeb - 1) & ". Using fallback for day of week."
        End If
        codeFerie = ModuleUtils.IsJourFerieOuRecup(CStr(ferieArr(1, col)))
        
        ' Get target staffing for the day
        Dim targetMatin As Long, targetPM As Long, targetSoir As Long
        targetMatin = arrTargetMatin(jourSemaine - 1, IIf(codeFerie, 1, 0)) ' Arrays are 0-indexed for day of week
        targetPM = arrTargetPM(jourSemaine - 1, IIf(codeFerie, 1, 0))
        targetSoir = arrTargetSoir(jourSemaine - 1, IIf(codeFerie, 1, 0))

        ' Get actual initial staffing for the day
        Dim actualMatin As Long, actualPM As Long, actualSoir As Long
        On Error Resume Next ' Handle non-numeric values in initial staffing rows
        actualMatin = IIf(IsNumeric(initialEffectifsMatin(1, col)), CLng(initialEffectifsMatin(1, col)), 0)
        actualPM = IIf(IsNumeric(initialEffectifsPM(1, col)), CLng(initialEffectifsPM(1, col)), 0)
        actualSoir = IIf(IsNumeric(initialEffectifsSoir(1, col)), CLng(initialEffectifsSoir(1, col)), 0)
        On Error GoTo 0

        ' Calculate initial needs
        Dim manqueMatin As Long, manquePM As Long, manqueSoir As Long
        Call ActualiserManquesValeurs(manqueMatin, manquePM, manqueSoir, targetMatin, targetPM, targetSoir, actualMatin, actualPM, actualSoir)

        ' --- Loop through available day replacement slots ---
        For l = 1 To NB_REMPLACEMENT_JOUR_LIGNES ' Corresponds to row index in rempJourArr
            If Trim(CStr(rempJourArr(l, col))) = "" Then ' If slot is empty
                Dim codePlaceCeTour As Boolean: codePlaceCeTour = False
                Dim codeATenter As String
                Dim ligneAbsolue As Long: ligneAbsolue = LIGNE_REMPLACEMENT_DEBUT_JOUR + l - 1
                
                ' --- GENERAL PRIORITY LOGIC (Example: Morning/PM combined need) ---
                If Not codePlaceCeTour And manqueMatin > 0 And manquePM > 0 Then
                    codeATenter = codesSuggestion(SUGG_7_1530)(0) ' e.g., "7 15:30"
                    If TryPlaceCodeIfValid(codeATenter, rempJourArr, l, col, planningArr, _
                        actualMatin, actualPM, actualSoir, ligneAbsolue, newlyPlaced_presence7_8h, _
                        manqueMatin, manquePM, manqueSoir, targetMatin, targetPM, targetSoir) Then
                        codePlaceCeTour = True
                    End If
                End If
                
                ' --- SPECIFIC EVENING LOGIC (FRIDAY/SATURDAY/HOLIDAY) ---
                Dim estJourSpecialSoir As Boolean
                estJourSpecialSoir = (jourSemaine = 5 Or jourSemaine = 6 Or codeFerie) ' Friday, Saturday, or Holiday

                If Not codePlaceCeTour And estJourSpecialSoir And manqueSoir > 0 Then
                    Dim c19CodePourJourSpecial As String: c19CodePourJourSpecial = IIf(codeFerie Or jourSemaine = 7, "C 19 di", "C 19") ' Sunday or holiday -> C 19 di

                    Dim c19EstPresentCeJour As Boolean: c19EstPresentCeJour = False
                    Dim c19EstRoleAS As Boolean: c19EstRoleAS = False ' Is the C19 an Aide-Soignant?
                    Dim rScan As Long, valScan As String
                    
                    ' 1. Check C19 in main planning (planningArr)
                    If IsArray(planningArr) Then
                        For rScan = LBound(planningArr, 1) To UBound(planningArr, 1)
                            valScan = Trim(CStr(planningArr(rScan, col)))
                            If StrComp(valScan, c19CodePourJourSpecial, vbTextCompare) = 0 Then
                                c19EstPresentCeJour = True
                                ' Check if this C19 is on the designated AS line
                                If (LIGNE_DEBUT_PLANNING_PERSONNEL + rScan - 1) = LIGNE_AIDE_SOIGNANT_C19_PLANNING Then c19EstRoleAS = True
                                Exit For
                            End If
                        Next rScan
                    End If

                    ' 2. If not in main planning, check in ALREADY PLACED day replacements for this day (rows above current)
                    If Not c19EstPresentCeJour Then
                        For rScan = 1 To l - 1 ' Only check rows already processed in rempJourArr for this day
                            valScan = Trim(CStr(rempJourArr(rScan, col)))
                            If StrComp(valScan, c19CodePourJourSpecial, vbTextCompare) = 0 Then
                                c19EstPresentCeJour = True
                                c19EstRoleAS = True ' Rule: C19 placed in replacement is considered AS for this logic block
                                Exit For
                            End If
                        Next rScan
                    End If
                    
                    ' Count existing C20 / C20E in main planning and already placed replacements
                    Dim nbC19Actuels As Long, nbC20EActuels As Long, nbC15Actuels As Long
                    Dim nbC20Actuels As Long
                    ' Utiliser la fonction CompterCodesSpecifiques pour compter les occurrences des codes coupés
                    Call CompterCodesSpecifiques(planningArr, col, nbC19Actuels, nbC20Actuels, nbC20EActuels, nbC15Actuels)
                    
                    ' Compter également dans les codes déjà placés dans rempJourArr jusqu'à la ligne actuelle
                    Dim tempRempArr As Variant
                    ReDim tempRempArr(1 To l - 1, 1 To 1) ' Créer un tableau temporaire pour les lignes déjà traitées
                    
                    Dim tempCol As Long  ' rScan est déjà déclaré à la ligne 608
                    tempCol = 1 ' Une seule colonne dans le tableau temporaire
                    
                    For rScan = 1 To l - 1
                        If rScan <= UBound(rempJourArr, 1) And col <= UBound(rempJourArr, 2) Then
                            tempRempArr(rScan, tempCol) = rempJourArr(rScan, col)
                        End If
                    Next rScan
                    
                    Dim nbC19Temp As Long, nbC20Temp As Long, nbC20ETemp As Long, nbC15Temp As Long
                    Call CompterCodesSpecifiques(tempRempArr, tempCol, nbC19Temp, nbC20Temp, nbC20ETemp, nbC15Temp)
                    
                    ' Additionner les compteurs
                    nbC19Actuels = nbC19Actuels + nbC19Temp
                    nbC20Actuels = nbC20Actuels + nbC20Temp
                    nbC20EActuels = nbC20EActuels + nbC20ETemp
                    nbC15Actuels = nbC15Actuels + nbC15Temp

                    ' Déterminer si c'est un jour spécial (vendredi, samedi ou férié)
                    Dim estJourSpecial As Boolean
                    estJourSpecial = EstJourVendrediSamediOuFerie(jourSemaine, codeFerie)
                    
                    ' A. Si C19 n'est pas présent et qu'il manque du personnel le soir
                    If Not codePlaceCeTour And Not c19EstPresentCeJour And manqueSoir > 0 Then
                        ' Utiliser le code C19 approprié pour ce jour
                        If TryPlaceCodeIfValid(c19CodePourJourSpecial, rempJourArr, l, col, planningArr, _
                            actualMatin, actualPM, actualSoir, ligneAbsolue, newlyPlaced_presence7_8h, _
                            manqueMatin, manquePM, manqueSoir, targetMatin, targetPM, targetSoir) Then
                            codePlaceCeTour = True
                            c19EstPresentCeJour = True ' Mettre à jour le statut car il est maintenant placé
                            c19EstRoleAS = True      ' Supposer que le C19 placé est AS pour la logique suivante
                        End If
                    End If

                    ' B. Si C19 est présent (ou vient d'être placé), essayer de placer C20/C20E
                    If Not codePlaceCeTour And c19EstPresentCeJour And manqueSoir > 0 Then
                        ' Utiliser la fonction DeterminerCodeCoupe pour choisir le code approprié
                        codeATenter = ""
                        
                        If c19EstRoleAS Then ' Si C19 est AS (ou considéré comme AS s'il est placé en remplacement)
                            ' Utiliser la fonction DeterminerCodeCoupe pour déterminer le code approprié
                            codeATenter = DeterminerCodeCoupe(nbC19Actuels, nbC20Actuels, nbC20EActuels, nbC15Actuels, estJourSpecial)
                        Else ' C19 est IDE
                            ' Pour les IDE, on autorise jusqu'à deux C20
                            If nbC20Actuels < 2 Then codeATenter = "C 20" ' Autoriser jusqu'à deux C20 si C19 est IDE
                        End If

                        If codeATenter <> "" Then
                            If TryPlaceCodeIfValid(codeATenter, rempJourArr, l, col, planningArr, _
                                actualMatin, actualPM, actualSoir, ligneAbsolue, newlyPlaced_presence7_8h, _
                                manqueMatin, manquePM, manqueSoir, targetMatin, targetPM, targetSoir) Then
                                codePlaceCeTour = True
                            End If
                        End If
                    End If
                
                ' --- ELSE (NOT a special evening day) AND manqueSoir > 0 for OTHER DAYS ---
                ElseIf Not codePlaceCeTour And manqueSoir > 0 Then ' Not Fri/Sat/Holiday, but evening still needed
                    ' Compter les occurrences des codes coupés pour les jours normaux
                    Dim nbC19NormalJour As Long, nbC20NormalJour As Long, nbC20ENormalJour As Long, nbC15NormalJour As Long
                    
                    ' Réutiliser les compteurs déjà calculés plus haut
                    nbC19NormalJour = nbC19Actuels
                    nbC20NormalJour = nbC20Actuels
                    nbC20ENormalJour = nbC20EActuels
                    nbC15NormalJour = nbC15Actuels
                    
                    ' Déterminer le code approprié pour les jours normaux
                    Dim estJourNormal As Boolean: estJourNormal = False ' Pas un jour spécial
                    Dim codeSoirGen As String
                    
                    ' Utiliser la fonction DeterminerCodeCoupe pour choisir le code approprié
                    codeSoirGen = DeterminerCodeCoupe(nbC19NormalJour, nbC20NormalJour, nbC20ENormalJour, nbC15NormalJour, estJourNormal)
                    
                    ' Ajuster le code C19 pour les dimanches et jours fériés
                    If codeSoirGen = "C 19" And (jourSemaine = 7 Or codeFerie) Then
                        codeSoirGen = "C 19 di" ' Utiliser C19 di pour les dimanches et jours fériés
                    End If
                    
                    ' Tenter de placer le code sélectionné
                    If codeSoirGen <> "" Then
                        If TryPlaceCodeIfValid(codeSoirGen, rempJourArr, l, col, planningArr, _
                            actualMatin, actualPM, actualSoir, ligneAbsolue, newlyPlaced_presence7_8h, _
                            manqueMatin, manquePM, manqueSoir, targetMatin, targetPM, targetSoir) Then
                            codePlaceCeTour = True
                        End If
                    End If
                End If

                ' --- LOGIC for Morning need ONLY --- (If no evening code was placed and morning is needed)
                If Not codePlaceCeTour And manqueMatin > 0 Then
                    codeATenter = codesSuggestion(SUGG_7_13)(0) ' "7 13"
                    If TryPlaceCodeIfValid(codeATenter, rempJourArr, l, col, planningArr, _
                        actualMatin, actualPM, actualSoir, ligneAbsolue, newlyPlaced_presence7_8h, _
                        manqueMatin, manquePM, manqueSoir, targetMatin, targetPM, targetSoir) Then
                        codePlaceCeTour = True
                    Else
                        codeATenter = codesSuggestion(SUGG_7_1130)(0) ' "7 11:30"
                        If TryPlaceCodeIfValid(codeATenter, rempJourArr, l, col, planningArr, _
                            actualMatin, actualPM, actualSoir, ligneAbsolue, newlyPlaced_presence7_8h, _
                            manqueMatin, manquePM, manqueSoir, targetMatin, targetPM, targetSoir) Then
                            codePlaceCeTour = True
                        End If
                        ' Add more alternatives for Matin if needed (e.g., SUGG_645, SUGG_8_1630 if they also cover Matin)
                    End If
                End If
                
                ' --- LOGIC for PM need ONLY --- (If no morning/evening code was placed and PM is needed)
                If Not codePlaceCeTour And manquePM > 0 Then
                     codeATenter = codesSuggestion(SUGG_12_30_16_30)(0) ' "12:30 16:30"
                     If TryPlaceCodeIfValid(codeATenter, rempJourArr, l, col, planningArr, _
                        actualMatin, actualPM, actualSoir, ligneAbsolue, newlyPlaced_presence7_8h, _
                        manqueMatin, manquePM, manqueSoir, targetMatin, targetPM, targetSoir) Then
                        codePlaceCeTour = True
                    End If
                    ' Add more alternatives for PM if needed
                End If

            End If ' End If slot is empty
        Next l ' Next day replacement slot

        ' --- Night shift management ---
        Dim nuitCodesProposes As Variant
        If codeFerie Or jourSemaine = 5 Or jourSemaine = 6 Then ' Friday, Saturday, Holiday
            nuitCodesProposes = Array(codesSuggestion(SUGG_NUIT1)(0), codesSuggestion(SUGG_NUIT2)(0)) ' e.g., "19:45 6:45", "20 7"
        Else
            nuitCodesProposes = Array(codesSuggestion(SUGG_NUIT2)(0), codesSuggestion(SUGG_NUIT2)(0)) ' e.g., "20 7", "20 7" (typically same for normal days)
        End If

        Dim iNuitSlot As Long
        For iNuitSlot = 1 To NB_REMPLACEMENT_NUIT_LIGNES ' Iterate through night replacement slots
            If UBound(rempNuitArr, 1) >= iNuitSlot And col <= UBound(rempNuitArr, 2) Then ' Check array bounds
                If Trim(CStr(rempNuitArr(iNuitSlot, col))) = "" Then ' If slot is empty
                    Dim codeNuitAPlacer As String: codeNuitAPlacer = ""
                    Dim codesPourChoixNuit As Variant

                    If iNuitSlot = 1 Then ' First night slot
                        codesPourChoixNuit = nuitCodesProposes
                        ' Ensure codesPourChoixNuit is always an array for ChoisirCodePertinent
                        ' This handles the case where nuitCodesProposes might be a single-element array from codesSuggestion
                        If Not IsArray(codesPourChoixNuit(LBound(codesPourChoixNuit))) And UBound(codesPourChoixNuit) = LBound(codesPourChoixNuit) Then
                             codesPourChoixNuit = Array(codesPourChoixNuit(LBound(codesPourChoixNuit)))
                        End If
                        codeNuitAPlacer = ChoisirCodePertinent(codesPourChoixNuit, planningArr, rempNuitArr, col)
                    Else ' Second night slot (or subsequent)
                        If CStr(nuitCodesProposes(LBound(nuitCodesProposes))) <> CStr(nuitCodesProposes(UBound(nuitCodesProposes))) Then
                            If Trim(CStr(rempNuitArr(1, col))) = CStr(nuitCodesProposes(LBound(nuitCodesProposes))) Then
                                codeNuitAPlacer = CStr(nuitCodesProposes(UBound(nuitCodesProposes)))
                            ElseIf Trim(CStr(rempNuitArr(1, col))) = CStr(nuitCodesProposes(UBound(nuitCodesProposes))) Then
                                codeNuitAPlacer = CStr(nuitCodesProposes(LBound(nuitCodesProposes)))
                            Else
                                ' If first slot filled with something else or empty, try preferred then other
                                If Not ModuleUtils.CodeDejaPresent(planningArr, rempNuitArr, col, CStr(nuitCodesProposes(LBound(nuitCodesProposes))), True) Then
                                    codeNuitAPlacer = CStr(nuitCodesProposes(LBound(nuitCodesProposes)))
                                ElseIf Not ModuleUtils.CodeDejaPresent(planningArr, rempNuitArr, col, CStr(nuitCodesProposes(UBound(nuitCodesProposes))), True) Then
                                     codeNuitAPlacer = CStr(nuitCodesProposes(UBound(nuitCodesProposes)))
                            codeNuitAPlacer = CStr(nuitCodesProposes(LBound(nuitCodesProposes)))
                        ElseIf Not ModuleUtils.CodeDejaPresent(planningArr, rempNuitArr, col, CStr(nuitCodesProposes(UBound(nuitCodesProposes))), True) Then
                             codeNuitAPlacer = CStr(nuitCodesProposes(UBound(nuitCodesProposes)))
                        End If
                    End If
                Else ' Both proposed codes are the same
                    codeNuitAPlacer = CStr(nuitCodesProposes(LBound(nuitCodesProposes)))
                End If
            End If
            
            If codeNuitAPlacer <> "" And Not ModuleUtils.CodeDejaPresent(planningArr, rempNuitArr, col, codeNuitAPlacer, True) Then
                rempNuitArr(iNuitSlot, col) = codeNuitAPlacer
            End If
        End If
    End If
Next iNuitSlot
JourSuivant:
Next col ' Next day

' Write the modified replacement arrays back to the worksheet once
ws.Range(ws.Cells(LIGNE_REMPLACEMENT_DEBUT_JOUR, colDeb), ws.Cells(LIGNE_REMPLACEMENT_FIN_JOUR, colDeb + nbJours - 1)).Value2 = rempJourArr
ws.Range(ws.Cells(LIGNE_REMPLACEMENT_DEBUT_NUIT, colDeb), ws.Cells(LIGNE_REMPLACEMENT_FIN_NUIT, colDeb + nbJours - 1)).Value2 = rempNuitArr
End Sub

Sub AnalyseEtRemplacementPlanningUltraOptimise()
    Dim ws As Worksheet
    Dim LdebFractions As Long, LfinFractions As Long
    Dim colDeb As Long
    Dim choixUtilisateur As VbMsgBoxResult
    Dim codesSuggestion As Variant

    ' --- USER CONFIGURABLE PARAMETERS ---
    colDeb = 2 ' Starting column for planning data (e.g., 2 for 'B', 3 for 'C')
    LdebFractions = 0 ' Start row for "fractions" data, 0 if not used
    LfinFractions = 0 ' End row for "fractions" data, 0 if not used
    ' --- END USER CONFIGURABLE PARAMETERS ---

    ' Initialize suggestion codes (arrays of codes for each suggestion type)
    codesSuggestion = Array( _
        Array("6:45 15:15"), Array("7 15:30"), Array("7 11:30"), Array("7 13"), Array("8 16:30"), _
        Array("C 15", "C 15 bis"), Array("C 20"), Array("C 20 E"), Array("C 19"), _
        Array("12:30 16:30"), Array("19:45 6:45"), Array("20 7"))

    ' User prompt for processing scope
    choixUtilisateur = MsgBox("Voulez-vous analyser uniquement l'onglet actif (" & ActiveSheet.Name & ") ?" & vbCrLf & _
                              vbCrLf & "Cliquez sur 'Oui' pour l'onglet actif." & vbCrLf & _
                              "Cliquez sur 'Non' pour analyser tous les onglets de mois." & vbCrLf & _
                              "Cliquez sur 'Annuler' pour vider les lignes de remplacement de l'onglet actif.", _
                              vbYesNoCancel + vbQuestion, "Choix de l'analyse")
    
    If colDeb <= 0 Then
        MsgBox "La colonne de début (colDeb = " & colDeb & ") n'est pas valide. Opération annulée.", vbCritical
        Exit Sub
    End If

    If choixUtilisateur = vbCancel Then
        If MsgBox("Voulez-vous VRAIMENT effacer les lignes de remplacement (Lignes " & LIGNE_REMPLACEMENT_DEBUT_JOUR & "-" & LIGNE_REMPLACEMENT_FIN_JOUR & " et " & LIGNE_REMPLACEMENT_DEBUT_NUIT & "-" & LIGNE_REMPLACEMENT_FIN_NUIT & ") de l'onglet '" & ActiveSheet.Name & "'?", _
                  vbYesNo + vbExclamation, "Confirmation Effacement") = vbYes Then
            Set ws = ActiveSheet
            If EstUnOngletDeMois(ws.Name) Then
                Application.ScreenUpdating = False
                Dim lastColData As Long
                On Error Resume Next ' Handle case where row 4 might be empty
                lastColData = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column
                If Err.Number <> 0 Or lastColData < colDeb Then lastColData = colDeb - 1 ' If error or no data beyond colDeb, set to avoid error on Range
                On Error GoTo 0

                If lastColData >= colDeb Then
                    ws.Range(ws.Cells(LIGNE_REMPLACEMENT_DEBUT_JOUR, colDeb), ws.Cells(LIGNE_REMPLACEMENT_FIN_JOUR, lastColData)).ClearContents
                    ws.Range(ws.Cells(LIGNE_REMPLACEMENT_DEBUT_NUIT, colDeb), ws.Cells(LIGNE_REMPLACEMENT_FIN_NUIT, lastColData)).ClearContents
                End If
                Application.ScreenUpdating = True
                MsgBox "Lignes de remplacement effacées pour l'onglet '" & ws.Name & "'.", vbInformation
            Else
                MsgBox "L'onglet actif (" & ws.Name & ") n'est pas un onglet de mois valide. Effacement non effectué.", vbExclamation
            End If
        Else
            MsgBox "Opération annulée par l'utilisateur.", vbInformation
        End If
        Exit Sub
    End If

    ' Performance optimizations
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler_Main ' Centralized error handling for the main process

    If choixUtilisateur = vbYes Then ' Process active sheet only
        Set ws = ActiveSheet
        If EstUnOngletDeMois(ws.Name) Then
            Debug.Print "Traitement onglet actif: " & ws.Name
            Call TraiterUneFeuilleDeMois(ws, LdebFractions, LfinFractions, colDeb, codesSuggestion)
            MsgBox "Analyse et remplacements pour l'onglet '" & ws.Name & "' terminés !", vbInformation
        Else
            MsgBox "L'onglet actif (" & ws.Name & ") n'est pas un onglet de mois valide. Opération non effectuée.", vbExclamation
        End If
    Else ' Process all month sheets
        For Each ws In ThisWorkbook.Worksheets
            If EstUnOngletDeMois(ws.Name) Then
                 Debug.Print "Traitement onglet: " & ws.Name
                Call TraiterUneFeuilleDeMois(ws, LdebFractions, LfinFractions, colDeb, codesSuggestion)
            End If
        Next ws
        MsgBox "Analyse et remplacements pour tous les mois terminés !", vbInformation
    End If

CleanExit_Main:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Debug.Print "Fin du script."
    Exit Sub

ErrorHandler_Main:
    MsgBox "Erreur d'exécution N° " & Err.Number & ":" & vbCrLf & Err.Description & vbCrLf & "Source: " & Err.Source, vbCritical, "Erreur VBA"
    Resume CleanExit_Main ' Go to cleanup steps
End Sub

