Attribute VB_Name = "ModulePersonnelUpdates"
Option Explicit

' --- Constants ---
Private Const PERSONNEL_SHEET_NAME As String = "Personnel"
' --- Adapté à tes captures d'écran ---
Private Const NAME_LAST_COL As Long = 2       ' Colonne B (Nom)
Private Const NAME_FIRST_COL As Long = 3      ' Colonne C (Prénom)
Private Const START_DATA_ROW As Long = 2      ' Données commencent ligne 2
Private Const START_POS_COL As Long = 27      ' Colonne AA (Position Janv)
Private Const START_PCT_COL As Long = 28      ' Colonne AB (Pourcentage Janv)
' --- Fin des adaptations principales ---
Private Const START_READ_COL As Long = NAME_LAST_COL ' Colonne la plus à gauche à lire (B=2)
Private Const NUM_MONTHS As Long = 12
Private Const TARGET_WRITE_COL As Long = 1     ' Colonne A sur feuilles cibles
Private Const MIN_TARGET_ROW As Long = 6      ' Ligne cible minimale valide

Sub UpdateMonthlySheets_Optimized()
    Dim wsPersonnel As Worksheet
    Dim lastRow As Long
    Dim i As Long, m As Long ' Loop counters (Long is better than Integer)
    Dim personnelData As Variant
    Dim dictMonthSheets As Object ' Dictionary(Of String, Worksheet)
    Dim dictEquivSheets As Object ' Dictionary(Of String, Worksheet)
    Dim monthSheetNames As Variant
    Dim equivSheetNames As Variant
    Dim errorCount As Long
    Dim startTime As Double

    startTime = Timer ' Start timing
    On Error GoTo ErrorHandler

    ' --- 1. Initial Setup & Validation ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Get Personnel Sheet
    On Error Resume Next
    Set wsPersonnel = ThisWorkbook.Sheets(PERSONNEL_SHEET_NAME)
    On Error GoTo ErrorHandler ' Restore proper error handling
    If wsPersonnel Is Nothing Then
        MsgBox "Feuille '" & PERSONNEL_SHEET_NAME & "' introuvable.", vbCritical, "Erreur"
        GoTo Cleanup
    End If

    ' Find Last Row (using a reliable column like LastName or FirstName)
    lastRow = wsPersonnel.Cells(wsPersonnel.rows.Count, NAME_LAST_COL).End(xlUp).row ' Utilise la colonne Nom (B)
    If lastRow < START_DATA_ROW Then
        MsgBox "Aucune donnée trouvée sur la feuille '" & PERSONNEL_SHEET_NAME & "'.", vbInformation, "Terminé"
        GoTo Cleanup
    End If

    ' --- 2. Load Target Sheet References ---
    monthSheetNames = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", "Juillet", "Aout", "Sept", "Oct", "Nov", "Dec")
    equivSheetNames = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12") ' Use strings if sheet names are strings

    Set dictMonthSheets = CreateObject("Scripting.Dictionary") ' Initialize dictionaries
    Set dictEquivSheets = CreateObject("Scripting.Dictionary")
    If Not LoadTargetSheets(monthSheetNames, dictMonthSheets) Then GoTo Cleanup
    If Not LoadTargetSheets(equivSheetNames, dictEquivSheets) Then GoTo Cleanup

    ' --- 3. Read Personnel Data into Array ---
    ' Determine the range to read: From the earliest needed column (LastName) to the last Percentage Col
    Dim endDataCol As Long
    endDataCol = START_PCT_COL + (NUM_MONTHS - 1) * 2 ' Last percentage column index

    On Error Resume Next ' Handle potential errors reading large range
    ' Lire à partir de la colonne Nom (B) jusqu'à la dernière colonne Pourcentage
    personnelData = wsPersonnel.Range(wsPersonnel.Cells(START_DATA_ROW, START_READ_COL), wsPersonnel.Cells(lastRow, endDataCol)).Value2
    If Err.Number <> 0 Then
         MsgBox "Erreur lors de la lecture des données depuis '" & PERSONNEL_SHEET_NAME & "':" & vbCrLf & Err.Description, vbCritical
         On Error GoTo ErrorHandler ' Back to main handler
         GoTo Cleanup ' Exit gracefully
    End If
    On Error GoTo ErrorHandler ' Restore main handler

    ' Safety check - reading a multi-cell range should always yield an array
    If Not IsArray(personnelData) Then
        MsgBox "Erreur critique: Les données lues depuis la feuille Personnel n'ont pas pu être interprétées comme un tableau (cas inattendu).", vbCritical
        GoTo Cleanup
    End If

    ' --- 4. Process Data in Memory ---
    Dim lastName As String, firstName As String, employeeFullName As String ' Changed order for clarity
    Dim targetRow As Variant ' Use Variant to handle potential errors or non-numeric values
    Dim percentageValue As Variant
    Dim lastNameIndexInArray As Long, firstNameIndexInArray As Long
    Dim posColIndexInArray As Long, pctColIndexInArray As Long
    Dim monthlyWs As Worksheet, equivWs As Worksheet
    Dim maxTargetRow As Long

    ' Calculate fixed indices within the array (relative to START_READ_COL)
    lastNameIndexInArray = NAME_LAST_COL - START_READ_COL + 1 ' B(2) - B(2) + 1 = 1
    firstNameIndexInArray = NAME_FIRST_COL - START_READ_COL + 1 ' C(3) - B(2) + 1 = 2

    errorCount = 0
    For i = 1 To UBound(personnelData, 1) ' Loop through ROWS (Employees) in the array (Range.Value2 gives 1-based array)
        lastName = Trim(CStr(personnelData(i, lastNameIndexInArray)))   ' Get LastName from Array Col 1
        firstName = Trim(CStr(personnelData(i, firstNameIndexInArray))) ' Get FirstName from Array Col 2

        If lastName = "" Or firstName = "" Then ' Check using the correct variables now
            Debug.Print "Avertissement: Nom/Prénom manquant à la ligne " & (i + START_DATA_ROW - 1) & " de la feuille Personnel."
            errorCount = errorCount + 1
            GoTo NextEmployee ' Skip this employee if name is incomplete
        End If

        ' Consistent Name Format
        employeeFullName = lastName & "_" & firstName

        ' Loop through MONTHS
        For m = 0 To NUM_MONTHS - 1 ' monthSheetNames is 0-based
            ' Calculate column indices WITHIN the personnelData array (relative to START_READ_COL)
            posColIndexInArray = (START_POS_COL - START_READ_COL + 1) + m * 2 ' AA(27) - B(2) + 1 = 26 (+m*2)
            pctColIndexInArray = (START_PCT_COL - START_READ_COL + 1) + m * 2 ' AB(28) - B(2) + 1 = 27 (+m*2)

            ' Boundary check for array indices (robustness)
            If posColIndexInArray > UBound(personnelData, 2) Or pctColIndexInArray > UBound(personnelData, 2) Then
                 Debug.Print "Erreur interne: Indice de colonne calculé hors limites pour " & employeeFullName & " (Ligne " & i & "), mois " & (m + 1)
                 errorCount = errorCount + 1
                 GoTo NextMonth ' Skip to next month for this employee
            End If

            percentageValue = personnelData(i, pctColIndexInArray)

            ' Check if percentage exists for this month
            If Not IsEmpty(percentageValue) And CStr(percentageValue) <> "" Then ' Check IsEmpty and non-blank string
                targetRow = personnelData(i, posColIndexInArray)

                ' Validate Target Row
                If IsNumeric(targetRow) Then
                    targetRow = CLng(targetRow) ' Convert to Long for comparison

                    ' Get target sheets for this month
                    Set monthlyWs = Nothing
                    Set equivWs = Nothing
                    If dictMonthSheets.Exists(monthSheetNames(m)) Then Set monthlyWs = dictMonthSheets(monthSheetNames(m))
                    If dictEquivSheets.Exists(equivSheetNames(m)) Then Set equivWs = dictEquivSheets(equivSheetNames(m))

                    Dim isRowValid As Boolean
                    isRowValid = (targetRow >= MIN_TARGET_ROW)
                    ' Optional: Add check against actual Rows.Count if needed
                    ' If isRowValid And Not monthlyWs Is Nothing Then
                    '     If targetRow > monthlyWs.Rows.Count Then isRowValid = False
                    ' End If

                    If isRowValid Then
                        ' Write to Monthly Sheet (if exists)
                        If Not monthlyWs Is Nothing Then
                             On Error Resume Next ' Handle potential write errors
                             monthlyWs.Cells(targetRow, TARGET_WRITE_COL).value = employeeFullName
                             If Err.Number <> 0 Then Debug.Print "Erreur écriture: " & monthSheetNames(m) & ", Ligne " & targetRow & ", Emp: " & employeeFullName & ", Err: " & Err.Description: Err.Clear: errorCount = errorCount + 1
                             On Error GoTo ErrorHandler
                        End If

                        ' Write to Equivalent Sheet (if exists)
                        If Not equivWs Is Nothing Then
                             On Error Resume Next ' Handle potential write errors
                             equivWs.Cells(targetRow, TARGET_WRITE_COL).value = employeeFullName
                              If Err.Number <> 0 Then Debug.Print "Erreur écriture: " & equivSheetNames(m) & ", Ligne " & targetRow & ", Emp: " & employeeFullName & ", Err: " & Err.Description: Err.Clear: errorCount = errorCount + 1
                             On Error GoTo ErrorHandler
                        End If
                    Else
                        Debug.Print "Avertissement: Ligne cible invalide (" & personnelData(i, posColIndexInArray) & ") pour " & employeeFullName & " Mois: " & monthSheetNames(m) & " (Ligne Personnel: " & (i + START_DATA_ROW - 1) & ")"
                        errorCount = errorCount + 1
                    End If
                Else
                    Debug.Print "Avertissement: Ligne cible non numérique (" & personnelData(i, posColIndexInArray) & ") pour " & employeeFullName & " Mois: " & monthSheetNames(m) & " (Ligne Personnel: " & (i + START_DATA_ROW - 1) & ")"
                    errorCount = errorCount + 1
                End If
            End If ' End check for percentage value
NextMonth:
        Next m ' Next Month
NextEmployee:
    Next i ' Next Employee

    ' --- 5. Finalization ---
    Dim finishMsg As String
    finishMsg = "Mise à jour terminée en " & Format(Timer - startTime, "0.00") & " secondes."
    If errorCount > 0 Then
        finishMsg = finishMsg & vbCrLf & vbCrLf & errorCount & " avertissement(s)/erreur(s) rencontré(s)." & vbCrLf & _
                    "Vérifiez la fenêtre Exécution (Ctrl+G) pour les détails."
        MsgBox finishMsg, vbExclamation, "Terminé avec Avertissements"
    Else
        MsgBox finishMsg, vbInformation, "Terminé"
    End If

Cleanup:
    ' Restore Application Settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ' Release Objects
    Set wsPersonnel = Nothing
    Set dictMonthSheets = Nothing
    Set dictEquivSheets = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Une erreur d'exécution est survenue:" & vbCrLf & vbCrLf & _
           "Erreur N°: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical, "Erreur d'Exécution"
    Resume Cleanup ' Go to cleanup section to restore settings
End Sub

Private Function LoadTargetSheets(sheetNameKeys As Variant, ByRef dictSheets As Object) As Boolean
    ' Loads worksheet objects into the provided dictionary. Returns False if critical error occurs.
    Dim ws As Worksheet
    Dim key As Variant
    Dim sheetFound As Boolean
    Dim missingSheets As String
    Dim firstMissing As Boolean: firstMissing = True

    ' Ensure the dictionary is initialized (passed ByRef)
    If dictSheets Is Nothing Then Set dictSheets = CreateObject("Scripting.Dictionary")
    dictSheets.CompareMode = vbTextCompare ' Case-insensitive keys

    missingSheets = ""
    For Each key In sheetNameKeys
        sheetFound = False
        On Error Resume Next ' Check if sheet exists
        Set ws = ThisWorkbook.Sheets(CStr(key))
        On Error GoTo 0 ' Turn off resume next immediately

        If Not ws Is Nothing Then
            If Not dictSheets.Exists(CStr(key)) Then ' Avoid adding duplicates if called multiple times
                dictSheets.Add CStr(key), ws
            End If
            sheetFound = True
            Set ws = Nothing ' Release temporary variable
        Else
            Debug.Print "Avertissement: La feuille cible '" & CStr(key) & "' n'a pas été trouvée et sera ignorée."
            If Not firstMissing Then missingSheets = missingSheets & ", "
            missingSheets = missingSheets & "'" & CStr(key) & "'"
            firstMissing = False
        End If
    Next key

    ' Decide if it's a critical failure (no sheets found at all)
    If dictSheets.Count = 0 Then
         MsgBox "Aucune des feuilles cibles attendues n'a été trouvée. L'opération ne peut pas continuer." & vbCrLf & _
                "Feuilles Manquantes: " & Join(sheetNameKeys, ", "), vbCritical, "Erreur Critique" ' Changed message slightly
         LoadTargetSheets = False
    Else
        If Len(missingSheets) > 0 Then
             Debug.Print "Info: Certaines feuilles cibles étaient manquantes: " & missingSheets
        End If
        LoadTargetSheets = True ' Succeeded, even if some sheets were missing (logged as warnings)
    End If
End Function


