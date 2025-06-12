Attribute VB_Name = "CalculateShiftsModule"
' =========================================================================================
'   MACRO DE CALCUL DE PLANNING - VERSION FINALE SIMPLIFIÉE
'   Date:   12 juin 2025
'   Description: Calcule les totaux pour la feuille active. La feuille "Liste" est la
'                SEULE source de vérité pour les calculs. La double vérification
'                a été supprimée pour plus de clarté et de flexibilité.
' =========================================================================================
Sub CalculateShiftsForActiveSheet()

    ' --- Configuration des Plages ---
    Const DayRangeAddress As String = "B6:AF25"
    Const NightRangeAddress As String = "B31:AF38"
    Const ReplacementRangeAddress As String = "B40:AF58"

    ' --- Lignes de base ---
    Const DayBaseRow As Long = 6
    Const NightBaseRow As Long = 31
    Const ReplacementBaseRow As Long = 40

    ' --- Déclarations ---
    Dim ws As Worksheet
    Dim shiftDict As Object ' Le seul dictionnaire dont nous avons besoin
    Dim shiftCode As String, cleanShiftCode As String
    Dim shiftInfo As Variant
    Dim daySchedule As Variant, nightSchedule As Variant, replacementSchedule As Variant
    Dim dayIdx As Long, rowIdx As Long
    Dim excludeCell As Boolean
    Dim scheduleArray As Variant, baseRow As Long
    Dim schedIdx As Long

    ' Arrays pour les totaux
    Dim dayTotalsMatin() As Long, dayTotalsApresMidi() As Long, dayTotalsSoir() As Long
    Dim dayTotalsPresence6h45() As Long, dayTotalsPresence7h8h() As Long, dayTotalsPresence8h16h30() As Long
    Dim dayTotalsPresenceC15() As Long, dayTotalsPresenceC20() As Long, dayTotalsPresenceC20E() As Long, dayTotalsPresenceC19() As Long
    Dim dayTotalsPresence1945() As Long, dayTotalsPresence207() As Long
    Dim dayTotalsTotalNuit() As Long

    ' --- Début de l'exécution ---
    Set ws = ActiveSheet

    ' Vérification que la feuille active est bien une feuille de planning
    If Not IsMonthSheet(ws.Name) Then
        MsgBox "Opération annulée." & vbCrLf & _
               "Veuillez lancer cette macro depuis une feuille de planning mensuel valide (ex: Oct, Nov, etc.).", _
               vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' --- ÉTAPE 1: Création du dictionnaire depuis la feuille "Liste" ---
    Set shiftDict = CreateShiftDictionaryFromSheet()
    If shiftDict Is Nothing Then GoTo CleanExit_Error

    ' --- ÉTAPE 2: Traitement de la feuille active ---
    ReDim dayTotalsMatin(1 To 31)
    ReDim dayTotalsApresMidi(1 To 31)
    ReDim dayTotalsSoir(1 To 31)
    ReDim dayTotalsPresence6h45(1 To 31)
    ReDim dayTotalsPresence7h8h(1 To 31)
    ReDim dayTotalsPresence8h16h30(1 To 31)
    ReDim dayTotalsPresenceC15(1 To 31)
    ReDim dayTotalsPresenceC20(1 To 31)
    ReDim dayTotalsPresenceC20E(1 To 31)
    ReDim dayTotalsPresenceC19(1 To 31)
    ReDim dayTotalsPresence1945(1 To 31)
    ReDim dayTotalsPresence207(1 To 31)
    ReDim dayTotalsTotalNuit(1 To 31)

    daySchedule = ReadRangeToArray(ws, DayRangeAddress)
    nightSchedule = ReadRangeToArray(ws, NightRangeAddress)
    replacementSchedule = ReadRangeToArray(ws, ReplacementRangeAddress)

    Dim arrSchedules As Variant: arrSchedules = Array(daySchedule, nightSchedule, replacementSchedule)
    Dim arrBaseRows As Variant: arrBaseRows = Array(DayBaseRow, NightBaseRow, ReplacementBaseRow)

    ' Boucle sur chaque jour
    For dayIdx = 1 To 31
        ' Boucle sur les 3 plages (Jour, Nuit, Remplacement)
        For schedIdx = 0 To 2
            scheduleArray = arrSchedules(schedIdx)
            baseRow = arrBaseRows(schedIdx)
            If IsArray(scheduleArray) Then
                ' Boucle sur chaque ligne de la plage
                For rowIdx = 1 To UBound(scheduleArray, 1)
                    shiftCode = Trim(CStr(scheduleArray(rowIdx, dayIdx)))

                    If shiftCode <> "" Then
                        excludeCell = False
                        ' Règle d'exclusion pour certains codes s'ils sont colorés
                        If shiftCode = "7 15:30" Or shiftCode = "6:45 15:15" Then
                            With ws.Cells(baseRow + rowIdx - 1, dayIdx + 1).DisplayFormat.Interior
                                If .ColorIndex <> xlNone Then
                                    excludeCell = True
                                End If
                            End With
                        End If

                        ' --- LOGIQUE DE CALCUL SIMPLE ET PROPRE ---
                        If schedIdx = 0 Then ' Totaux Matin/AM/Soir uniquement pour la plage JOUR
                            If shiftDict.Exists(shiftCode) Then
                                shiftInfo = shiftDict(shiftCode)
                                If shiftInfo(0) And Not excludeCell Then dayTotalsMatin(dayIdx) = dayTotalsMatin(dayIdx) + 1
                                If shiftInfo(1) And Not excludeCell Then dayTotalsApresMidi(dayIdx) = dayTotalsApresMidi(dayIdx) + 1
                                If shiftInfo(2) And Not excludeCell Then dayTotalsSoir(dayIdx) = dayTotalsSoir(dayIdx) + 1
                            End If
                        End If

                        cleanShiftCode = Replace(shiftCode, " ", "")
                        UpdatePresenceTotals cleanShiftCode, dayIdx, excludeCell, schedIdx, _
                            dayTotalsPresence6h45, dayTotalsPresence7h8h, dayTotalsPresence8h16h30, _
                            dayTotalsPresenceC15, dayTotalsPresenceC20, dayTotalsPresenceC20E, dayTotalsPresenceC19, _
                            dayTotalsPresence1945, dayTotalsPresence207
                    End If
                Next rowIdx
            End If
        Next schedIdx
    Next dayIdx

    For dayIdx = 1 To 31
        dayTotalsTotalNuit(dayIdx) = dayTotalsPresence1945(dayIdx) + dayTotalsPresence207(dayIdx)
    Next dayIdx

    WriteResultsToSheet ws, dayTotalsMatin, dayTotalsApresMidi, dayTotalsSoir, _
        dayTotalsPresence6h45, dayTotalsPresence7h8h, dayTotalsPresence8h16h30, _
        dayTotalsPresenceC15, dayTotalsPresenceC20, dayTotalsPresenceC20E, dayTotalsPresenceC19, _
        dayTotalsPresence1945, dayTotalsPresence207, dayTotalsTotalNuit

CleanExit_Success:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Calculs pour la feuille '" & ws.Name & "' terminés avec succès !", vbInformation
    Exit Sub

CleanExit_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Erreur VBA #" & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit_Error
End Sub

' --- Fonctions de support ---

Private Function IsMonthSheet(sheetName As String) As Boolean
    Select Case LCase(Trim(sheetName))
        Case "janv", "fev", "mars", "avril", "mai", "juin", "juillet", "aout", "sept", "oct", "nov", "dec"
            IsMonthSheet = True
        Case Else
            IsMonthSheet = False
    End Select
End Function

Private Function CreateShiftDictionaryFromSheet() As Object
    Dim wsListe As Worksheet
    Dim lastRow As Long
    Dim dataArr As Variant
    Dim i As Long
    Dim dict As Object

    On Error Resume Next
    Set wsListe = ThisWorkbook.Sheets("Liste")
    On Error GoTo 0
    If wsListe Is Nothing Then
        MsgBox "Feuille 'Liste' introuvable.", vbCritical
        Set CreateShiftDictionaryFromSheet = Nothing
        Exit Function
    End If

    lastRow = wsListe.Cells(wsListe.Rows.Count, 1).End(xlUp).Row
    dataArr = wsListe.Range("A2:G" & lastRow).Value
    Set dict = CreateObject("Scripting.Dictionary")

    If IsArray(dataArr) Then
        For i = LBound(dataArr, 1) To UBound(dataArr, 1)
            Dim code As String
            code = Trim(CStr(dataArr(i, 1)))
            If code <> "" Then
                Dim arr As Variant
                arr = Array( _
                    (Not IsError(dataArr(i, 4))) And CBool(dataArr(i, 4) > 0), _
                    (Not IsError(dataArr(i, 5))) And CBool(dataArr(i, 5) > 0), _
                    (Not IsError(dataArr(i, 6))) And CBool(dataArr(i, 6) > 0), _
                    (Not IsError(dataArr(i, 7))) And CBool(dataArr(i, 7) > 0))
                If Not dict.Exists(code) Then dict.Add code, arr
            End If
        Next i
    End If

    Set CreateShiftDictionaryFromSheet = dict
End Function

Private Function ReadRangeToArray(ws As Worksheet, rangeAddr As String) As Variant
    Dim tempArray As Variant
    On Error Resume Next
    tempArray = ws.Range(rangeAddr).Value
    If Err.Number <> 0 Then
        ReadRangeToArray = Empty
        Err.Clear
    ElseIf Not IsArray(tempArray) Then
        Dim arr(1 To 1, 1 To 1) As Variant
        arr(1, 1) = tempArray
        ReadRangeToArray = arr
    Else
        ReadRangeToArray = tempArray
    End If
End Function

Private Sub UpdatePresenceTotals(code As String, day As Long, exclude As Boolean, schedIdx As Long, _
        T64() As Long, T65() As Long, T66() As Long, T67() As Long, T68() As Long, T69() As Long, T70() As Long, _
        T71() As Long, T72() As Long)
    If schedIdx = 0 Then ' Plage Jour
        Select Case code
            Case "6:4515:15", "6:4512:45"
                If Not exclude Then T64(day) = 1: T65(day) = T65(day) + 1
            Case "6:4512:14", "713", "711", "711:30"
                T65(day) = T65(day) + 1
            Case "715:30"
                If Not exclude Then T65(day) = T65(day) + 1
            Case "7:3016"
                T65(day) = T65(day) + 1: T66(day) = 1
            Case "1016:30", "8:3016:30"
                T66(day) = 1
            Case "C15", "16:3020:15", "8:3012:4516:3020:15"
                T67(day) = 1
            Case "C20"
                T68(day) = 1
            Case "C20E"
                T69(day) = 1
            Case "C19", "C19di"
                T65(day) = T65(day) + 1: T70(day) = 1
            Case "1519", "15:3019"
                T70(day) = 1
        End Select
    ElseIf schedIdx = 1 Then ' Plage Nuit
        Select Case code
            Case "19:456:45"
                If Not exclude Then T71(day) = T71(day) + 1
            Case "207"
                If Not exclude Then T72(day) = T72(day) + 1
        End Select
    End If
End Sub

Private Sub WriteResultsToSheet(ws As Worksheet, T60() As Long, T61() As Long, T62() As Long, _
    T64() As Long, T65() As Long, T66() As Long, T67() As Long, T68() As Long, T69() As Long, T70() As Long, _
    T71() As Long, T72() As Long, T73() As Long)

    On Error Resume Next
    ws.Range("B60:AF60").Value = T60
    ws.Range("B61:AF61").Value = T61
    ws.Range("B62:AF62").Value = T62
    ws.Range("B64:AF64").Value = T64
    ws.Range("B65:AF65").Value = T65
    ws.Range("B66:AF66").Value = T66
    ws.Range("B67:AF67").Value = T67
    ws.Range("B68:AF68").Value = T68
    ws.Range("B69:AF69").Value = T69
    ws.Range("B70:AF70").Value = T70
    ws.Range("B71:AF71").Value = T71
    ws.Range("B72:AF72").Value = T72
    ws.Range("B73:AF73").Value = T73
    On Error GoTo 0
End Sub

