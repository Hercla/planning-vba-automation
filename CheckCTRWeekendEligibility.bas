Attribute VB_Name = "CTRWeekendCheck"
Option Explicit
'
' CheckCTRWeekendEligibility verifies that each employee has worked at least one
' Saturday and one Sunday during the previous month. Valid weekend shifts are
' read from the 'Configuration' sheet so the list can be updated without code changes.

Sub CTR_CheckWeekendEligibility()
    On Error GoTo Cleanup

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsCurrent As Worksheet
    Dim wsPrev As Worksheet
    Dim wbPrev As Workbook ' Workbook for previous year when needed
    Dim baseName As String
    Dim monthDate As Date
    Dim prevMonthDate As Date
    Dim prevSheetName As String
    Dim shiftType As String
    Dim prevYearFilePath As String ' Path to previous year's file
    Dim startRow As Long, lastRow As Long
    Dim headerRow As Long, startCol As Long, endCol As Long
    Dim header As Variant
    Dim row As Long, j As Long
    Dim saturdayWorked As Boolean
    Dim sundayWorked As Boolean
    Dim data As Variant
    Dim employeesWithoutWeekend As String
    Dim wsConfig As Worksheet
    Dim configCol As Long
    Dim validShifts As Object
    Dim startRowJour As Long, startRowNuit As Long

    Set wsCurrent = ActiveSheet

    baseName = wsCurrent.Name
    baseName = Replace(baseName, " nuit", "", , , vbTextCompare)
    baseName = Replace(baseName, " jour", "", , , vbTextCompare)

    ' --- Charger les paramètres depuis la feuille Configuration ---
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Configuration")
    On Error GoTo Cleanup
    If wsConfig Is Nothing Then
        MsgBox "La feuille 'Configuration' est introuvable.", vbCritical, "Vérification CTR"
        GoTo Cleanup
    End If

    startRowJour = wsConfig.Cells(2, 2).Value
    startRowNuit = wsConfig.Cells(2, 3).Value

    shiftType = ""
    If wsCurrent.Rows(startRowJour).Hidden = False Then
        shiftType = "jour"
    ElseIf wsCurrent.Rows(startRowNuit).Hidden = False Then
        shiftType = "nuit"
    End If

    ' Fallback: infer from sheet name if visibility detection fails
    If shiftType = "" Then
        If InStr(1, wsCurrent.Name, "nuit", vbTextCompare) > 0 Then
            shiftType = "nuit"
        ElseIf InStr(1, wsCurrent.Name, "jour", vbTextCompare) > 0 Then
            shiftType = "jour"
        End If
    End If

    If shiftType = "" Then
        MsgBox "Cette macro doit être lancée depuis un planning jour ou nuit.", _
               vbExclamation, "Vérification CTR"
        GoTo Cleanup
    End If

    If shiftType = "jour" Then
        configCol = 2 ' Colonne B
    Else
        configCol = 3 ' Colonne C
    End If

    startRow = wsConfig.Cells(2, configCol).Value
    lastRow = wsConfig.Cells(3, configCol).Value
    headerRow = wsConfig.Cells(4, configCol).Value
    startCol = wsConfig.Cells(5, configCol).Value
    endCol = wsConfig.Cells(6, configCol).Value

    Set validShifts = CreateObject("Scripting.Dictionary")
    validShifts.CompareMode = vbTextCompare
    Dim shiftRange As Range, cell As Range
    Set shiftRange = wsConfig.Range("E2", wsConfig.Cells(wsConfig.Rows.Count, "E").End(xlUp))
    For Each cell In shiftRange
        If Trim(cell.Value) <> "" Then validShifts(Trim(cell.Value)) = 1
    Next cell

    monthDate = GetMonthDateFromName(baseName)
    If monthDate = CDate(0) Then
        MsgBox "Nom de feuille non reconnu : " & wsCurrent.Name, vbExclamation, "Vérification CTR"
        GoTo Cleanup
    End If

    prevMonthDate = DateAdd("m", -1, monthDate)
    prevSheetName = MonthToSheetName(prevMonthDate) & " " & shiftType

    ' --- Determine the worksheet containing the previous month's data ---
    ' The sheet might be named either "<Month> <shiftType>" or simply "<Month>".
    On Error Resume Next
    If Month(monthDate) = 1 Then
        ' January: previous month is in last year's file
        prevYearFilePath = ThisWorkbook.Path & "\Planning_" & Year(prevMonthDate) & ".xlsm"
        If Dir(prevYearFilePath) <> "" Then
            Set wbPrev = Workbooks.Open(prevYearFilePath, ReadOnly:=True)
            Set wsPrev = wbPrev.Sheets(prevSheetName)
            If wsPrev Is Nothing Then
                Set wsPrev = wbPrev.Sheets(MonthToSheetName(prevMonthDate))
                If Not wsPrev Is Nothing Then prevSheetName = wsPrev.Name
            End If
        Else
            MsgBox "Le fichier du mois précédent est introuvable." & vbNewLine & _
                   "Chemin recherché : " & prevYearFilePath, vbCritical, "Vérification CTR"
            GoTo Cleanup
        End If
    Else
        ' Normal case: the previous sheet is expected in this workbook
        Set wsPrev = ThisWorkbook.Sheets(prevSheetName)
        If wsPrev Is Nothing Then
            Set wsPrev = ThisWorkbook.Sheets(MonthToSheetName(prevMonthDate))
            If Not wsPrev Is Nothing Then prevSheetName = wsPrev.Name
        End If
    End If
    On Error GoTo Cleanup
    If wsPrev Is Nothing Then
        MsgBox "Feuille du mois précédent introuvable : '" & prevSheetName & "'", vbCritical, "Vérification CTR"
        GoTo Cleanup
    End If

    If lastRow < startRow Then GoTo Display

    header = wsPrev.Range(wsPrev.Cells(headerRow, startCol), wsPrev.Cells(headerRow, endCol)).Value
    employeesWithoutWeekend = ""

    For row = startRow To lastRow
        data = wsPrev.Range(wsPrev.Cells(row, startCol), wsPrev.Cells(row, endCol)).Value
        saturdayWorked = False
        sundayWorked = False
        For j = 1 To UBound(header, 2)
            If LCase(Trim(header(1, j))) = "sam" Then
                If IsWeekendShift(data(1, j), validShifts) Then saturdayWorked = True
            ElseIf LCase(Trim(header(1, j))) = "dim" Then
                If IsWeekendShift(data(1, j), validShifts) Then sundayWorked = True
            End If
        Next j
        If Not (saturdayWorked And sundayWorked) Then
            employeesWithoutWeekend = employeesWithoutWeekend & wsPrev.Cells(row, 1).Value & vbNewLine
        End If
    Next row

Display:
    If employeesWithoutWeekend <> "" Then
        MsgBox "Les employés suivants n'ont pas presté de week-end complet le mois précédent et ne peuvent pas recevoir de code CTR :" & _
               vbNewLine & employeesWithoutWeekend, vbExclamation, "Vérification CTR"
    Else
        MsgBox "Tous les employés de l'équipe '" & shiftType & "' sont éligibles pour un code CTR ce mois-ci.", _
               vbInformation, "Vérification CTR"
    End If

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    If Not wbPrev Is Nothing Then
        wbPrev.Close SaveChanges:=False
    End If
    If Err.Number <> 0 Then
        MsgBox "Erreur : " & Err.Description, vbCritical, "Vérification CTR"
    End If
End Sub

' Determine if a cell value indicates the employee worked a shift.
' Weekend work codes are matched case-insensitively after trimming spaces.
Private Function IsWeekendShift(ByVal cellValue As Variant, shiftDict As Object) As Boolean
    Dim val As String
    val = Trim(CStr(cellValue))
    If val = "" Then Exit Function

    IsWeekendShift = shiftDict.Exists(val)
End Function

