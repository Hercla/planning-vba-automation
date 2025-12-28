Attribute VB_Name = "ModuleVerificationCTR"
Sub CTR_CheckWeekendEligibility()
    On Error GoTo ErrorHandler

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsCurrent As Worksheet, wsPrev As Worksheet, wbPrev As Workbook
    Dim baseName As String, monthDate As Date, prevMonthDate As Date
    Dim prevSheetName As String, shiftType As String, prevYearFilePath As String
    Dim startRow As Long, lastRow As Long, headerRow As Long, startCol As Long, endCol As Long
    Dim header As Variant, sanitizedHeader() As String, headerCount As Long
    Dim row As Long, j As Long
    Dim data As Variant
    Dim employeesWithoutWeekend As String
    Dim wsConfig As Worksheet, configCol As Long
    Dim validShifts As Object

    Set wsCurrent = ActiveSheet
    baseName = wsCurrent.Name
    baseName = Replace(baseName, " nuit", "", , , vbTextCompare)
    baseName = Replace(baseName, " jour", "", , , vbTextCompare)

    ' Tente de trouver la feuille de configuration avec le nom correct
    On Error Resume Next
    ' ===================================================================
    ' ----- LIGNE CORRIGÉE -----
    ' ===================================================================
    Set wsConfig = ThisWorkbook.Sheets("Configuration_CTR_CheckWeek")
    ' ===================================================================

    On Error GoTo ErrorHandler ' Réactive le gestionnaire d'erreurs principal

    If wsConfig Is Nothing Then
        MsgBox "La feuille 'Configuration_CTR_CheckWeek' est introuvable.", vbCritical, "Erreur de configuration"
        GoTo Cleanup
    End If

    ' --- Détection du type de planning (jour/nuit) ---
    shiftType = ""
    Dim startRowJour As Long, startRowNuit As Long
    startRowJour = wsConfig.Cells(2, 2).value
    startRowNuit = wsConfig.Cells(2, 3).value

    If wsCurrent.Rows(startRowJour).Hidden = False Then
        shiftType = "jour"
    ElseIf wsCurrent.Rows(startRowNuit).Hidden = False Then
        shiftType = "nuit"
    End If
    
    If shiftType = "" Then
        If InStr(1, wsCurrent.Name, "nuit", vbTextCompare) > 0 Then
            shiftType = "nuit"
        ElseIf InStr(1, wsCurrent.Name, "jour", vbTextCompare) > 0 Then
            shiftType = "jour"
        End If
    End If

    If shiftType = "" Then
        MsgBox "Impossible de déterminer si le planning est de type Jour ou Nuit.", vbExclamation, "Vérification CTR"
        GoTo Cleanup
    End If

    ' --- Configuration des variables depuis la feuille 'Configuration' ---
    If shiftType = "jour" Then configCol = 2 Else configCol = 3
    startRow = wsConfig.Cells(2, configCol).value
    lastRow = wsConfig.Cells(3, configCol).value
    headerRow = wsConfig.Cells(4, configCol).value
    startCol = wsConfig.Cells(5, configCol).value
    endCol = wsConfig.Cells(6, configCol).value

    ' --- Chargement des codes de shift valides ---
    Set validShifts = CreateObject("Scripting.Dictionary")
    validShifts.CompareMode = vbTextCompare
    Dim shiftRange As Range, cell As Range
    Set shiftRange = wsConfig.Range("E2", wsConfig.Cells(wsConfig.Rows.Count, "E").End(xlUp))
    For Each cell In shiftRange
        If Trim(cell.value) <> "" Then validShifts(Trim(cell.value)) = 1
    Next cell

    ' --- Recherche de la feuille du mois précédent ---
    monthDate = GetMonthDateFromName(baseName)
    If monthDate = CDate(0) Then GoTo Cleanup
    prevMonthDate = DateAdd("m", -1, monthDate)
    prevSheetName = MonthToSheetName(prevMonthDate) & " " & shiftType

    On Error Resume Next
    Set wsPrev = Nothing
    If Month(monthDate) = 1 Then
        prevYearFilePath = ThisWorkbook.Path & "\Planning_" & Year(prevMonthDate) & ".xlsm"
        If Dir(prevYearFilePath) <> "" Then
            Set wbPrev = Workbooks.Open(prevYearFilePath, ReadOnly:=True)
            If Not wbPrev Is Nothing Then Set wsPrev = wbPrev.Sheets(prevSheetName)
            If wsPrev Is Nothing Then Set wsPrev = wbPrev.Sheets(MonthToSheetName(prevMonthDate))
        End If
    Else
        Set wsPrev = ThisWorkbook.Sheets(prevSheetName)
        If wsPrev Is Nothing Then Set wsPrev = ThisWorkbook.Sheets(MonthToSheetName(prevMonthDate))
    End If
    On Error GoTo ErrorHandler ' Réactive le gestionnaire d'erreurs principal

    If wsPrev Is Nothing Then
        MsgBox "Feuille du mois précédent introuvable : '" & prevSheetName & "' ou '" & MonthToSheetName(prevMonthDate) & "'", vbCritical, "Feuille manquante"
        GoTo Cleanup
    End If

    If lastRow < startRow Then GoTo DisplayResult

    ' --- Traitement des en-têtes (avec gestion du cas d'une seule colonne) ---
    header = wsPrev.Range(wsPrev.Cells(headerRow, startCol), wsPrev.Cells(headerRow, endCol)).value
    If Not IsArray(header) Then
        headerCount = 1
        ReDim sanitizedHeader(1 To 1)
        sanitizedHeader(1) = LCase(Trim(CStr(header)))
    Else
        headerCount = UBound(header, 2)
        ReDim sanitizedHeader(1 To headerCount)
        For j = 1 To headerCount
            sanitizedHeader(j) = LCase(Trim(CStr(header(1, j))))
        Next j
    End If
    
    ' --- Vérification des employés ---
    employeesWithoutWeekend = ""
    For row = startRow To lastRow
        data = wsPrev.Range(wsPrev.Cells(row, startCol), wsPrev.Cells(row, endCol)).value
        ' Si les données ne sont pas un tableau (une seule cellule), un WE complet est impossible
        If IsArray(data) Then
            If Not HasWorkedCompleteWeekend(data, sanitizedHeader, validShifts) Then
                employeesWithoutWeekend = employeesWithoutWeekend & wsPrev.Cells(row, 1).value & vbNewLine
            End If
        Else
            employeesWithoutWeekend = employeesWithoutWeekend & wsPrev.Cells(row, 1).value & vbNewLine
        End If
    Next row

DisplayResult:
    If employeesWithoutWeekend <> "" Then
        MsgBox "Les employés suivants n'ont pas presté de week-end complet le mois précédent et ne peuvent pas recevoir de code CTR :" & _
               vbNewLine & employeesWithoutWeekend, vbExclamation, "Vérification CTR"
    Else
        MsgBox "Tous les employés de l'équipe '" & shiftType & "' sont éligibles pour un code CTR ce mois-ci.", _
               vbInformation, "Vérification CTR"
    End If

    GoTo Cleanup ' Sortie normale

ErrorHandler:
    ' Affiche uniquement les erreurs VRAIMENT inattendues
    MsgBox "Erreur inattendue: " & Err.Description & " (Code: " & Err.Number & ")", vbCritical, "Erreur d'exécution"

Cleanup:
    ' Nettoyage final, s'exécute dans tous les cas
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    If Not wbPrev Is Nothing Then wbPrev.Close SaveChanges:=False
End Sub
' Vos fonctions HasWorkedCompleteWeekend et IsWeekendShift restent ici, inchangées

'
' Vérifie si une ligne de données contient un Samedi travaillé suivi d'un Dimanche travaillé.
'
Private Function HasWorkedCompleteWeekend(ByVal dataRow As Variant, _
                                          ByRef headers() As String, _
                                          ByVal shifts As Object) As Boolean
    Dim i As Long
    For i = 1 To UBound(headers) - 1
        ' Vérifie si on a une paire sam/dim
        If headers(i) = "sam" And headers(i + 1) = "dim" Then
            ' Vérifie si les DEUX jours ont un code de travail valide
            If IsWeekendShift(dataRow(1, i), shifts) And _
               IsWeekendShift(dataRow(1, i + 1), shifts) Then
                HasWorkedCompleteWeekend = True
                Exit Function ' C'est bon, on a trouvé. Inutile de chercher plus loin.
            End If
        End If
    Next i
    ' Si on arrive ici, aucun WE complet n'a été trouvé. La fonction retourne False par défaut.
End Function

'
' Détermine si une valeur de cellule correspond à un code de travail valide.
'
Private Function IsWeekendShift(ByVal cellValue As Variant, shiftDict As Object) As Boolean
    Dim val As String
    val = Trim(CStr(cellValue))
    If val = "" Then Exit Function
    IsWeekendShift = shiftDict.Exists(val)
End Function



