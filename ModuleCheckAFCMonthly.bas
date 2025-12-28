Attribute VB_Name = "ModuleCheckAFCMonthly"
Option Explicit
' Vérifie que certains employés ont le nombre requis de codes "AFC"
' dans la feuille de planning actuellement ouverte.
Sub CheckAFCMonthlyCodes()
    On Error GoTo ErrorHandler

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet
    Dim wsConfig As Worksheet
    Dim shiftType As String
    Dim startRow As Long, lastRow As Long
    Dim startCol As Long, endCol As Long
    Dim row As Long, col As Long
    Dim employeeName As String, countAFC As Long
    Dim configCol As Long
    Dim expectedCounts As Object
    Dim report As String

    Set ws = ActiveSheet
    
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
    Dim startRowJour As Long, startRowNuit As Long
    startRowJour = wsConfig.Cells(2, 2).value
    startRowNuit = wsConfig.Cells(2, 3).value

    shiftType = ""

    If ws.Rows(startRowJour).Hidden = False Then
        shiftType = "jour"
    ElseIf ws.Rows(startRowNuit).Hidden = False Then
        shiftType = "nuit"
    End If
    
    If shiftType = "" Then
        If InStr(1, ws.Name, "nuit", vbTextCompare) > 0 Then
            shiftType = "nuit"
        ElseIf InStr(1, ws.Name, "jour", vbTextCompare) > 0 Then
            shiftType = "jour"
        End If
    End If

    If shiftType = "" Then
        MsgBox "Impossible de déterminer si le planning est de type Jour ou Nuit." & vbNewLine & vbNewLine & _
               "Vérifiez que les lignes des employés (jour/nuit) sont correctement affichées/masquées " & _
               "OU que le nom de l'onglet contient 'jour' ou 'nuit'.", _
               vbExclamation, "Vérification AFC"
        GoTo Cleanup
    End If

    ' --- Configuration des variables ---
    If shiftType = "jour" Then configCol = 2 Else configCol = 3
    startRow = wsConfig.Cells(2, configCol).value
    lastRow = wsConfig.Cells(3, configCol).value
    startCol = wsConfig.Cells(5, configCol).value
    endCol = wsConfig.Cells(6, configCol).value

    ' --- Chargement des employés et des comptes attendus ---
    Set expectedCounts = CreateObject("Scripting.Dictionary")
    
    Dim lastConfigRow As Long
    Dim i As Long
    Dim configEmployeeName As String
    Dim configExpectedCount As Long
    Dim configShiftType As String
    
    lastConfigRow = wsConfig.Cells(wsConfig.Rows.Count, "G").End(xlUp).row
    
    For i = 2 To lastConfigRow
        configEmployeeName = LCase(Trim(CStr(wsConfig.Cells(i, "G").value)))
        configExpectedCount = wsConfig.Cells(i, "H").value
        configShiftType = LCase(Trim(CStr(wsConfig.Cells(i, "I").value)))
        
        If configShiftType = shiftType And configEmployeeName <> "" Then
            If Not expectedCounts.Exists(configEmployeeName) Then
                expectedCounts.Add configEmployeeName, configExpectedCount
            End If
        End If
    Next i
    
    If expectedCounts.Count = 0 Then
        MsgBox "Aucun employé à vérifier n'a été trouvé dans la configuration pour l'équipe de " & shiftType & ".", vbInformation, "Vérification AFC"
        GoTo Cleanup
    End If

    ' --- Vérification des codes AFC sur la feuille de planning ---
    report = ""
    For row = startRow To lastRow
        If Not IsEmpty(ws.Cells(row, 1).value) Then
            employeeName = LCase(Trim(CStr(ws.Cells(row, 1).value)))
            
            If expectedCounts.Exists(employeeName) Then
                countAFC = 0
                For col = startCol To endCol
                    If UCase(Trim(CStr(ws.Cells(row, col).value))) = "AFC" Then
                        countAFC = countAFC + 1
                    End If
                Next col
                
                If countAFC <> expectedCounts(employeeName) Then
                    report = report & ws.Cells(row, 1).value & " : " & countAFC & _
                             " AFC (attendu " & expectedCounts(employeeName) & ")" & vbNewLine
                End If
            End If
        End If
    Next row

    ' --- Affichage du rapport final ---
    If report <> "" Then
        MsgBox "Vérification AFC - écarts détectés pour l'équipe de " & shiftType & ":" & vbNewLine & vbNewLine & report, vbExclamation, "Rapport AFC"
    Else
        MsgBox "Tous les employés ciblés de l'équipe de " & shiftType & " possèdent le nombre requis de codes AFC.", _
               vbInformation, "Vérification AFC"
    End If
    
    GoTo Cleanup ' Sortie normale

ErrorHandler:
    MsgBox "Erreur inattendue: " & Err.Description & " (Code: " & Err.Number & ")", vbCritical, "Erreur d'exécution"

Cleanup:
    Set ws = Nothing
    Set wsConfig = Nothing
    Set expectedCounts = Nothing
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

