Attribute VB_Name = "Module_Configuration"
Option Explicit

' --- CONSTANTES DE CONFIGURATION ---
Private Const LISTE_SHEET_NAME As String = "Liste"
Private Const COL_A_HEADER As String = "CodeComplet"
Private Const COL_B_HEADER As String = "CodeCongéFormulaire_Base"
Private Const NAMED_RANGE_B As String = "ListeCongesStandards"

Sub UpdateListeSheet_PreserveExisting()
    Dim wsListe As Worksheet
    Dim dictExistingCodesA As Object ' Pour vérifier les codes existants en Col A
    Dim colorDict As Object          ' Pour appliquer la couleur aux nouveaux codes
    
    ' ****** MODIFICATION TYPE VARIABLE ******
    Dim standardSelectableCodes() As Variant ' Déclaré explicitement comme tableau
    Dim tempArray As Variant                 ' Variable temporaire pour Array()
    Dim k As Long                            ' Compteur pour copie
    ' ****** FIN MODIFICATION ******

    ' Listes de référence (pas pour écriture directe en A cette fois)
    Dim holidayCodes As Variant, recupCodes As Variant
    Dim workCodes As Variant, structureCodes As Variant
    Dim codeItem As Variant
    Dim i As Long, lrA As Long, lrB As Long
    Dim userResponse As VbMsgBoxResult
    Dim startTime As Single: startTime = Timer
    Dim addedCodesCount As Long: addedCodesCount = 0

    ' Optimisation d'Excel
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    ' --- 1. Obtenir la Feuille "Liste" ---
    On Error Resume Next
    Set wsListe = ThisWorkbook.Worksheets(LISTE_SHEET_NAME)
    On Error GoTo 0
    If wsListe Is Nothing Then
        MsgBox "La feuille '" & LISTE_SHEET_NAME & "' est introuvable." & vbCrLf & _
               "Veuillez créer la feuille et la structurer (au moins Colonne A) avant de lancer cette macro.", vbCritical, "Feuille Manquante"
        GoTo Cleanup
    End If

    ' --- 2. Définir les Listes de Codes "Cibles" pour la colonne B ---
    ' ****** MODIFICATION ASSIGNATION ******
    tempArray = Array("ANC", "CA", "EL", "CTR", "CS", "FOR", "MAL", "CSS", "EM", "PAT", "PREAVIS", "VJ", "FP", "CEP", "CP", "DP", "FSH", "PETIT CHOM", "Décès", "Déménag", "Grève") ' Assigner à la variable temporaire

    ' Redimensionner et copier vers le tableau explicite
    If IsArray(tempArray) Then ' Vérification de sécurité
        ReDim standardSelectableCodes(LBound(tempArray) To UBound(tempArray))
        For k = LBound(tempArray) To UBound(tempArray)
            standardSelectableCodes(k) = tempArray(k)
        Next k
        Erase tempArray ' Libérer la mémoire de la variable temporaire
    Else
        ' Gérer le cas où Array() échouerait (très improbable)
        MsgBox "Erreur critique: Impossible de créer le tableau initial des codes.", vbCritical
        GoTo Cleanup
    End If
    ' ****** FIN MODIFICATION ******

    ' Trier le tableau maintenant explicitement déclaré
    QuickSortStringArray standardSelectableCodes ' Appel de la fonction de tri

    ' Listes pour référence (à adapter à vos codes exacts)
    holidayCodes = Array("F 1-1", "F 22-4", "F 1-5", "F 30-5", "F 10-6", "F 21-7", "F 15-8", "F 1-11", "F 11-11", "F 25-12", "F 9-4", "F 21-5", "F 28-5")
    recupCodes = Array("R 1-1", "R 22-4", "R 1-5", "R 30-5", "R 10-6", "R 21-7", "R 15-8", "R 1-11", "R 11-11", "R 25-12", "R 9-4", "R 21-5", "R 28-5")
    workCodes = Array("8 12", "8 12:30", "8:30 12", "8:30 13", "9 13", "9:30 18", "9 16:30", "10 13", "10 14", "10:30 17", _
                      "11 30 20", "11 16:30", "10 19", "8 20", "11:30 20", "13 18", "13 16.30", "13 20:15", _
                      "14 18", "14 20", "15 19", "15 19 sa", "15:30 19", "15 20", "16 20", "12 16:30", "18 21:00", "6:45 12:45", "M", "S", "N")
    structureCodes = Array("WE", "/")
    
    ' --- 3. Lire les codes existants en colonne A ---
    Set dictExistingCodesA = CreateObject("Scripting.Dictionary")
    dictExistingCodesA.CompareMode = vbTextCompare
    lrA = wsListe.Cells(wsListe.rows.Count, "A").End(xlUp).row
    If lrA < 1 Then lrA = 1

    Dim startRowA As Long
    If UCase(Trim(CStr(wsListe.Cells(1, "A").value))) = UCase(COL_A_HEADER) Then
        startRowA = 2
    Else
        startRowA = 1
    End If

    For i = startRowA To lrA
        Dim currentCodeRead As String ' Utiliser un nom différent pour éviter conflit potentiel
        currentCodeRead = Trim(CStr(wsListe.Cells(i, "A").value))
        If currentCodeRead <> "" Then
            If Not dictExistingCodesA.Exists(currentCodeRead) Then
                dictExistingCodesA.Add currentCodeRead, i
            End If
        End If
    Next i

    ' --- 4. Définir la palette de couleurs pour les nouveaux codes ---
    Set colorDict = CreateObject("Scripting.Dictionary")
    colorDict.Add "StandardLeave", RGB(204, 255, 204)
    colorDict.Add "SpecialLeave", RGB(204, 229, 255)
    colorDict.Add "EventLeave", RGB(255, 255, 204)
    colorDict.Add "Holiday", RGB(255, 217, 102)
    colorDict.Add "Recup", RGB(221, 217, 255)
    colorDict.Add "NonWork", RGB(242, 242, 242)
    colorDict.Add "Maladie", RGB(255, 204, 153)
    colorDict.Add "Formation", RGB(157, 195, 230)
    colorDict.Add "SansSolde", RGB(255, 230, 230)
    colorDict.Add "Problem", RGB(255, 150, 150)
    colorDict.Add "WorkCode", xlNone
    colorDict.Add "Default", RGB(217, 217, 217)

    ' --- 5. Ajouter les codes génériques manquants en colonne A ---
    lrA = wsListe.Cells(wsListe.rows.Count, "A").End(xlUp).row
    For Each codeItem In standardSelectableCodes ' Boucle sur le tableau explicite
        Dim codeStr As String
        codeStr = CStr(codeItem)
        If Not dictExistingCodesA.Exists(codeStr) Then
            lrA = lrA + 1
            wsListe.Cells(lrA, "A").value = codeStr
            wsListe.Cells(lrA, "A").Interior.Color = GetCodeColorOptimized(codeStr, colorDict, workCodes) ' Utilise la fonction couleur affinée
            dictExistingCodesA.Add codeStr, lrA
            addedCodesCount = addedCodesCount + 1
            Debug.Print "Code ajouté en A " & lrA & ": " & codeStr
        End If
    Next codeItem

    ' --- 6. Nettoyer et remplir la colonne B ---
    lrB = wsListe.Cells(wsListe.rows.Count, "B").End(xlUp).row
    If lrB < 1 Then lrB = 1
    If lrB >= 2 Then
        With wsListe.Range("B2:B" & lrB)
            .ClearContents
            .Interior.ColorIndex = xlNone
            .Font.ColorIndex = xlAutomatic
            .Font.Bold = False
        End With
    End If
    With wsListe.Cells(1, "B")
        If UCase(Trim(CStr(.value))) <> UCase(COL_B_HEADER) Then
            .value = COL_B_HEADER
        End If
        .Font.Bold = True
    End With
    lrB = 1
    For Each codeItem In standardSelectableCodes ' Boucle sur le tableau explicite et trié
        lrB = lrB + 1
        Dim cellB As Range
        Set cellB = wsListe.Cells(lrB, "B")
        cellB.value = CStr(codeItem)
        ' Appliquer la couleur correspondante
        Dim bgColorB As Variant ' Variable locale pour la couleur
        bgColorB = GetCodeColorOptimized(CStr(codeItem), colorDict, workCodes)
         If Not IsEmpty(bgColorB) Then
             If bgColorB <> xlNone Then
                 cellB.Interior.Color = bgColorB
             Else
                 cellB.Interior.ColorIndex = xlNone
             End If
         End If
    Next codeItem

    ' --- 7. Créer/mettre à jour la plage nommée ---
    If lrB > 1 Then
        Dim rangeAddress As String
        rangeAddress = "=" & wsListe.Name & "!$B$2:$B$" & lrB
        On Error Resume Next
        ThisWorkbook.Names(NAMED_RANGE_B).Delete
        On Error GoTo 0
        ThisWorkbook.Names.Add Name:=NAMED_RANGE_B, RefersTo:=rangeAddress
    Else
        Debug.Print "Avertissement: Aucun code n'a été écrit en colonne B, la plage nommée '" & NAMED_RANGE_B & "' n'a pas été créée/mise à jour."
    End If

    ' --- 8. Finalisation ---
    wsListe.Columns("A:B").AutoFit
    MsgBox "Configuration de la feuille '" & LISTE_SHEET_NAME & "' terminée." & vbCrLf & vbCrLf & _
           addedCodesCount & " code(s) générique(s) ajouté(s) en colonne A (si manquants)." & vbCrLf & _
           "Colonne B mise à jour avec " & lrB - 1 & " code(s) sélectionnable(s) et couleurs appliquées." & vbCrLf & _
           "Plage nommée '" & NAMED_RANGE_B & "' créée/mise à jour." & vbCrLf & vbCrLf & _
           "Traitement terminé en " & Format(Timer - startTime, "0.00") & " secondes.", _
           vbInformation, "Configuration Terminée"

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Set wsListe = Nothing
    Set dictExistingCodesA = Nothing
    Set colorDict = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Une erreur est survenue lors de la configuration de la feuille Liste :" & vbCrLf & vbCrLf & _
           "Erreur N°: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical, "Erreur Configuration Liste"
    Resume Cleanup
End Sub

' --- Fonction pour obtenir la couleur d'un code (AFFINÉE) ---
Private Function GetCodeColorOptimized(ByVal code As String, dictColor As Object, workCodes As Variant) As Variant
    Dim defaultColor As Variant
    Dim bgColor As Variant
    defaultColor = dictColor("Default") ' Couleur par défaut si non catégorisé
    bgColor = defaultColor              ' Initialiser avec défaut

    On Error Resume Next ' Au cas où une clé manque dans dictColor

    ' Gérer les cas spécifiques en premier
    If Left(code, 1) = "F" And InStr(code, "-") > 0 Then
        bgColor = dictColor("Holiday")
    ElseIf Left(code, 1) = "R" And InStr(code, "-") > 0 Then
        bgColor = dictColor("RecupJF") ' Catégorie spécifique pour R x-x
    ElseIf code = "WE" Or code = "/" Then
        bgColor = dictColor("NonWork")
    ElseIf IsInArray(code, workCodes) Then
        bgColor = dictColor("WorkCode")
    Else
        ' Essayer de trouver une correspondance directe avec les clés génériques
        Select Case code
            Case "CL", "ANC", "EL", "VJ"
                bgColor = dictColor("StandardLeave")
            Case "FP", "CEP", "CP", "AAIR", "AFC", "BUS", "FSH", "CS", "RHS", "RV" ' Ajout RHS, RV
                bgColor = dictColor("SpecialLeave")
            Case "PETITCHOM", "Décès", "Déménag"
                bgColor = dictColor("EventLeave")
            Case "CTR", "R.AFC", "RCT" ' Ajout RCT
                 bgColor = dictColor("Recup") ' Même couleur que R x-x
            Case "MAL", "MAT", "PAT", "EM" ' Ajout EM
                bgColor = dictColor("Maladie")
            Case "FOR" ' FOR était seul dans la catégorie
                bgColor = dictColor("Formation")
            Case "CSS", "PREAVIS" ' CP géré avec SpecialLeave ou ici? Décidé ici.
                bgColor = dictColor("SansSolde")
            Case "Grève"
                bgColor = dictColor("Problem")
            Case "DP" ' Code DP spécifique
                bgColor = xlNone ' Pas de couleur pour DP selon l'image
            Case Else
                ' Si aucune correspondance, reste la couleur par défaut
                 'Debug.Print "Warning: Code '" & code & "' non catégorisé, utilisant Default."
        End Select
    End If

    If Err.Number <> 0 Then bgColor = defaultColor: Err.Clear ' Fallback
    GetCodeColorOptimized = bgColor
    On Error GoTo 0
End Function


' --- Fonction IsInArray (Inchangée) ---
Private Function IsInArray(itemToFind As Variant, arr As Variant) As Boolean
' Vérifie si un élément existe dans un tableau (1D) - insensible à la casse pour les strings
    Dim element As Variant
    IsInArray = False
    If Not IsArray(arr) Then Exit Function

    On Error Resume Next ' Au cas où le tableau contient des erreurs
    For Each element In arr
        If VarType(itemToFind) = vbString And VarType(element) = vbString Then
            If StrComp(itemToFind, element, vbTextCompare) = 0 Then
                IsInArray = True
                Exit Function
            End If
        Else
             If itemToFind = element Then
                IsInArray = True
                Exit Function
            End If
        End If
    Next element
    On Error GoTo 0
End Function

' --- QuickSort pour trier un tableau de Variants (contenant des Strings) ---
' La signature attend bien un tableau ByRef arr() As Variant
Sub QuickSortStringArray(ByRef arr() As Variant, Optional LB As Long = -1, Optional UB As Long = -1)
    Dim p As Variant
    Dim l As Long, r As Long
    Dim i As Long, j As Long

    If Not IsArray(arr) Then Exit Sub ' Vérification si c'est un tableau

    ' Gestion des bornes par défaut ou invalides
    If LB = -1 Then LB = LBound(arr)
    If UB = -1 Then UB = UBound(arr)
    If LB >= UB Then Exit Sub ' Tableau vide, à un élément ou bornes invalides

    i = LB
    j = UB
    ' Choose pivot (middle element) - convert to string for comparison
    On Error Resume Next ' Handle potential non-string variant
    p = CStr(arr((LB + UB) \ 2))
    If Err.Number <> 0 Then p = "": Err.Clear ' Fallback pivot
    On Error GoTo 0

    Do While i <= j
        ' Find element on left that should be on right (case-insensitive)
        While StrComp(CStr(arr(i)), p, vbTextCompare) < 0 And i < UB
            i = i + 1
        Wend
         If i > UB And StrComp(CStr(arr(UB)), p, vbTextCompare) < 0 Then i = UB

        ' Find element on right that should be on left (case-insensitive)
        While StrComp(CStr(arr(j)), p, vbTextCompare) > 0 And j > LB
            j = j - 1
        Wend
        If j < LB And StrComp(CStr(arr(LB)), p, vbTextCompare) > 0 Then j = LB

        If i <= j Then
            ' Swap elements
            Dim Temp As Variant
            Temp = arr(i)
            arr(i) = arr(j)
            arr(j) = Temp
            ' Move indices
            i = i + 1
            j = j - 1
        End If
    Loop

    ' Recursive calls for sub-arrays
    If LB < j Then QuickSortStringArray arr, LB, j
    If i < UB Then QuickSortStringArray arr, i, UB
End Sub

