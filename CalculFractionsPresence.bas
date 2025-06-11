Attribute VB_Name = "CalculFractionsPresence"
Sub CalculateAllShiftsAllSheetsOptimized_Combined_V8_Hybrid()

' --- Enumerations for presence lines ---
Private Enum PresenceLine
    PL_6h45 = 64
    PL_7h8h = 65
    PL_8h16h30 = 66
    PL_C15 = 67
    PL_C20 = 68
    PL_C20E = 69
    PL_C19 = 70
    PL_1945 = 71
    PL_207 = 72
End Enum

Private Enum PresenceAction
    ActionSet = 1
    ActionAdd = 2
End Enum

 ' --- Configuration des Plages ---
    Const DayRangeAddress As String = "B6:AF25"
    Const NightRangeAddress As String = "B31:AF38"
    Const ReplacementRangeAddress As String = "B40:AF58"

    ' --- Base Rows ---
    Const DayBaseRow As Long = 6
    Const NightBaseRow As Long = 31
    Const ReplacementBaseRow As Long = 40

    Dim codePresenceDay As Object, codePresenceNight As Object
    ' --- Couleurs à exclure ---
    Const YellowColor As Long = 65535
    Const BlueColor As Long = 16711680

    ' --- !!! Listes Codes Autorisés (VÉRIFIÉES ET MISES À JOUR) !!! ---
    ' Ces codes ne seront comptés (pour L60-L62) que s'ils viennent de la PLAGE JOUR (B6:AF25)
    Const AllowedCodesMatin As String = "|6:45 15:15|7 15:30|6:45 12:45|7:30 16|8 16:30|8:30 16:30|C 19|C 19 di|C 20 E|C 15|C 15 di|8:30 12:45 16:30 20:15|7 13|7:15 13:15|8 14|8:30 14|7 11:30|7:15 15:45|C 20|10 19|8 16|10 16:30|7 16:30|"
    Const AllowedCodesAM As String = "|6:45 15:15|7 15:30|7:30 16|10 16:30|8:30 16:30|C 15|16:30 20:15|8:30 12:45 16:30 20:15|15 19|15:30 19|8 16:30|7 16:30|8 16|12:30 16:30|7:15 15:45|12 16|13 19|10 19|14 20|"
    Const AllowedCodesSoir As String = "|C 15|16:30 20:15|8:30 12:45 16:30 20:15|C 20|C 20 E|C 19|15 19|15:30 19|C 19 di|C 15 di|16 20|14 20|13 19|10 19|"
    ' --- FIN LISTES CODES AUTORISÉS ---


    ' --- Déclarations ---
    Dim wsListe As Worksheet, ws As Worksheet
    Dim listLR As Long, i As Long
    Dim listDataRange As Range
    Dim shiftData As Variant
    Dim shiftDict As Object ' Scripting.Dictionary
    Dim shiftCode As String, cleanShiftCode As String
    Dim shiftAssignments As Variant ' Array [M, AM, S, N] from Liste
    Dim daySchedule As Variant, nightSchedule As Variant, replacementSchedule As Variant
    Dim dayIdx As Long, rowIdx As Long, wsCol As Long, wsRow As Long
    Dim shiftInfo As Variant
    Dim cellColor As Long, excludeCell As Boolean
    Dim scheduleArray As Variant, baseRow As Long

    ' Arrays pour totaux
    Dim dayTotalsMatin() As Long, dayTotalsApresMidi() As Long, dayTotalsSoir() As Long
    Dim dayTotalsPresence6h45() As Long, dayTotalsPresence7h8h() As Long, dayTotalsPresence8h16h30() As Long
    Dim dayTotalsPresenceC15() As Long, dayTotalsPresenceC20() As Long, dayTotalsPresenceC20E() As Long, dayTotalsPresenceC19() As Long
    Dim dayTotalsPresence1945() As Long, dayTotalsPresence207() As Long
    Dim dayTotalsTotalNuit() As Long ' Tableau pour Ligne 73

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' --- Initialisation Dictionnaire depuis "Liste" ---
    On Error Resume Next
    Set wsListe = ThisWorkbook.Sheets("Liste")
    If wsListe Is Nothing Then MsgBox "Feuille 'Liste' introuvable.", vbCritical: GoTo CleanExit_Error
    On Error GoTo ErrorHandler
    Set shiftDict = CreateObject("Scripting.Dictionary")
    listLR = wsListe.Cells(wsListe.rows.Count, "A").End(xlUp).row
    If listLR >= 2 Then
        Set listDataRange = wsListe.Range("A2:G" & listLR)
        shiftData = listDataRange.value
        If IsArray(shiftData) Then
            For i = 1 To UBound(shiftData, 1)
                shiftCode = Trim(CStr(shiftData(i, 1)))
                If shiftCode <> "" Then
                    shiftAssignments = Array( _
                        (Not IsError(shiftData(i, 4))) And CBool(shiftData(i, 4) > 0), _
                        (Not IsError(shiftData(i, 5))) And CBool(shiftData(i, 5) > 0), _
                        (Not IsError(shiftData(i, 6))) And CBool(shiftData(i, 6) > 0), _
                        (Not IsError(shiftData(i, 7))) And CBool(shiftData(i, 7) > 0))
                    If Not shiftDict.Exists(shiftCode) Then shiftDict.Add shiftCode, shiftAssignments
                End If
            Next i
        End If
    End If
    ' --- Fin Initialisation Dictionnaire ---

    ' --- Traitement Feuilles Mensuelles ---
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
                    Set codePresenceDay = CreatePresenceCodeMapDay()
                    Set codePresenceNight = CreatePresenceCodeMapNight()
             Select Case True ' Vérifier nom feuille
                Case ws.Name Like "Janv*", ws.Name Like "Fev*", ws.Name Like "Mars*", _
                     ws.Name Like "Avril*", ws.Name Like "Mai*", ws.Name Like "Juin*", _
                     ws.Name Like "Juillet*", ws.Name Like "Aout*", ws.Name Like "Sept*", _
                     ws.Name Like "Oct*", ws.Name Like "Nov*", ws.Name Like "Dec*", _
                     ws.Name Like "JanvB", ws.Name Like "FevB"

                    ' Réinitialiser totaux
                    ReDim dayTotalsMatin(1 To 31): ReDim dayTotalsApresMidi(1 To 31): ReDim dayTotalsSoir(1 To 31)
                    ReDim dayTotalsPresence6h45(1 To 31): ReDim dayTotalsPresence7h8h(1 To 31): ReDim dayTotalsPresence8h16h30(1 To 31)
                    ReDim dayTotalsPresenceC15(1 To 31): ReDim dayTotalsPresenceC20(1 To 31): ReDim dayTotalsPresenceC20E(1 To 31): ReDim dayTotalsPresenceC19(1 To 31)
                    ReDim dayTotalsPresence1945(1 To 31): ReDim dayTotalsPresence207(1 To 31)
                    ReDim dayTotalsTotalNuit(1 To 31) ' ReDim Tableau Ligne 73

                    ' Lire planning
                    daySchedule = ReadRangeToArray(ws, DayRangeAddress)
                    nightSchedule = ReadRangeToArray(ws, NightRangeAddress)
                    replacementSchedule = ReadRangeToArray(ws, ReplacementRangeAddress)

                    ' Boucles Jour / Plages / Lignes
                    For dayIdx = 1 To 31
                        wsCol = dayIdx + 1
                        Dim arrSchedules As Variant: arrSchedules = Array(daySchedule, nightSchedule, replacementSchedule)
                        Dim arrBaseRows As Variant: arrBaseRows = Array(DayBaseRow, NightBaseRow, ReplacementBaseRow)
                        Dim schedIdx As Long

                        ' Boucle sur les 3 plages : 0=Jour, 1=Nuit, 2=Remplacement
                        For schedIdx = LBound(arrSchedules) To UBound(arrSchedules)
                            scheduleArray = arrSchedules(schedIdx): baseRow = arrBaseRows(schedIdx)
                            If IsArray(scheduleArray) Then
                                For rowIdx = 1 To UBound(scheduleArray, 1)
                                    If Not IsError(scheduleArray(rowIdx, dayIdx)) Then
                                        shiftCode = Trim(CStr(scheduleArray(rowIdx, dayIdx)))
                                    Else: shiftCode = ""
                                    End If

                                    If shiftCode <> "" Then
                                        cleanShiftCode = Replace(shiftCode, " ", "")
                                        excludeCell = False
                                        ' Logique exclusion couleur (inchangée)
                                        If shiftCode = "7 15:30" Then
                                            wsRow = baseRow + rowIdx - 1: On Error Resume Next
                                            cellColor = ws.Cells(wsRow, wsCol).Interior.Color
                                            If Err.Number <> 0 Then cellColor = 0: Err.Clear
                                            On Error GoTo ErrorHandler
                                            If cellColor = YellowColor Or cellColor = BlueColor Then excludeCell = True
                                        End If

                                        ' --- CALCULS UNIQUEMENT SI PLAGE JOUR (schedIdx = 0) ---
                                        If schedIdx = 0 Then
                                            ' --- LOGIQUE HYBRIDE pour L60, L61, L62 ---
                                            If shiftDict.Exists(shiftCode) Then
                                                shiftInfo = shiftDict(shiftCode)
                                                ' Calcul L60
                                                If shiftInfo(0) Then
                                                    If InStr(1, AllowedCodesMatin, "|" & shiftCode & "|", vbTextCompare) > 0 Then
                                                        If Not (shiftCode = "7 15:30" And excludeCell) Then
                                                            dayTotalsMatin(dayIdx) = dayTotalsMatin(dayIdx) + 1
                                                        End If
                                                    End If
                                                End If
                                                ' Calcul L61
                                                If shiftInfo(1) Then
                                                    If InStr(1, AllowedCodesAM, "|" & shiftCode & "|", vbTextCompare) > 0 Then
                                                         If Not (shiftCode = "7 15:30" And excludeCell) Then
                                                            dayTotalsApresMidi(dayIdx) = dayTotalsApresMidi(dayIdx) + 1
                                                        End If
                                                    End If
                                                End If
                                                ' Calcul L62
                                                If shiftInfo(2) Then
                                                    If InStr(1, AllowedCodesSoir, "|" & shiftCode & "|", vbTextCompare) > 0 Then
                                                        dayTotalsSoir(dayIdx) = dayTotalsSoir(dayIdx) + 1
                                                    End If
                                                End If
                                            End If ' Fin if shiftDict.Exists

                                            If codePresenceDay.Exists(cleanShiftCode) Then
                                                If Not (cleanShiftCode = "715:30" And excludeCell) Then
                                                    Dim act As Variant
                                                    For Each act In codePresenceDay(cleanShiftCode)
                                                        ApplyPresenceAction act(0), act(1), dayIdx, _
                                                            dayTotalsPresence6h45, dayTotalsPresence7h8h, dayTotalsPresence8h16h30, _
                                                            dayTotalsPresenceC15, dayTotalsPresenceC20, dayTotalsPresenceC20E, _
                                                            dayTotalsPresenceC19, dayTotalsPresence1945, dayTotalsPresence207
                                                    Next act
                                                End If
                                            End If
                                            If codePresenceNight.Exists(cleanShiftCode) Then
                                                Dim act As Variant
                                                For Each act In codePresenceNight(cleanShiftCode)
                                                    ApplyPresenceAction act(0), act(1), dayIdx, _
                                                        dayTotalsPresence6h45, dayTotalsPresence7h8h, dayTotalsPresence8h16h30, _
                                                        dayTotalsPresenceC15, dayTotalsPresenceC20, dayTotalsPresenceC20E, _
                                                        dayTotalsPresenceC19, dayTotalsPresence1945, dayTotalsPresence207
                                                Next act
                                            End If
                                            ' --- FIN CALCULS PLAGE NUIT ---'
                                                    dayTotalsPresence1945(dayIdx) = dayTotalsPresence1945(dayIdx) + 1
                                                Case "207"
                                                    dayTotalsPresence207(dayIdx) = dayTotalsPresence207(dayIdx) + 1
                                            End Select
                                        ' --- FIN CALCULS PLAGE NUIT ---

                                        ' End If ' Implicite : rien à faire pour schedIdx = 2 (Remplacement)
                                        End If ' Fin de la condition principale sur schedIdx

                                    End If ' End If shiftCode <> ""
                                Next rowIdx
                            End If ' End If IsArray
                        Next schedIdx ' Prochaine Plage
                    Next dayIdx ' Fin boucle jours

                    ' *** Calculer le total pour la ligne 73 ***
                    For dayIdx = 1 To 31
                        dayTotalsTotalNuit(dayIdx) = dayTotalsPresence1945(dayIdx) + dayTotalsPresence207(dayIdx)
                    Next dayIdx

                    ' --- Écriture résultats (SANS Ligne 63, AVEC Ligne 73) ---
                    On Error Resume Next
                    ws.Range("B60:AF60").value = dayTotalsMatin
                    ws.Range("B61:AF61").value = dayTotalsApresMidi
                    ws.Range("B62:AF62").value = dayTotalsSoir
                    ' Ligne 63 ignorée
                    ws.Range("B64:AF64").value = dayTotalsPresence6h45
                    ws.Range("B65:AF65").value = dayTotalsPresence7h8h
                    ws.Range("B66:AF66").value = dayTotalsPresence8h16h30
                    ws.Range("B67:AF67").value = dayTotalsPresenceC15
                    ws.Range("B68:AF68").value = dayTotalsPresenceC20
                    ws.Range("B69:AF69").value = dayTotalsPresenceC20E
                    ws.Range("B70:AF70").value = dayTotalsPresenceC19
                    ws.Range("B71:AF71").value = dayTotalsPresence1945
                    ws.Range("B72:AF72").value = dayTotalsPresence207
                    ws.Range("B73:AF73").value = dayTotalsTotalNuit ' *** Écriture Ligne 73 ***
                    If Err.Number <> 0 Then
                         MsgBox "Avertissement: Écriture résultats échouée sur '" & ws.Name & "'.", vbExclamation: Err.Clear
                    End If
                    On Error GoTo ErrorHandler

            End Select ' Fin Select Case nom feuille
        End If ' End If ws.Visible
    Next ws ' Prochaine feuille

CleanExit_Success:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Calculs (Hybride, Filtres Plages Stricts + L73) terminés !", vbInformation
    Exit Sub

CleanExit_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Erreur VBA #" & Err.Number & ": " & Err.Description & vbCrLf & _
           "Procédure: CalculateAllShiftsAllSheetsOptimized_Combined_V8_Hybrid_StrictRanges_Final_V2", vbCritical
    Resume CleanExit_Error
End Sub


' --- Fonction ReadRangeToArray (CORRIGÉE) ---
Function ReadRangeToArray(ws As Worksheet, rangeAddr As String) As Variant
    Dim tempArray As Variant

    On Error Resume Next ' Gérer l'erreur si la plage n'existe pas ou autre problème
    tempArray = ws.Range(rangeAddr).value
    If Err.Number <> 0 Then ' Si une erreur s'est produite lors de la lecture
        ReadRangeToArray = Empty ' Retourner Empty
        Err.Clear             ' Effacer l'erreur
' --- Utility to add presence codes ---
Private Sub AddPresenceCode(dict As Object, code As String, ParamArray actions() As Variant)
    Dim c As Collection, i As Long
    Set c = New Collection
    For i = LBound(actions) To UBound(actions)
        c.Add actions(i)
    Next i
    dict.Add Replace(code, " ", ""), c
End Sub

Private Function CreatePresenceCodeMapDay() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    AddPresenceCode d, "6:4515:15", Array(PL_6h45, ActionSet), Array(PL_7h8h, ActionAdd)
    AddPresenceCode d, "6:4512:45", Array(PL_6h45, ActionSet), Array(PL_7h8h, ActionAdd)
    AddPresenceCode d, "6:4512:14", Array(PL_7h8h, ActionAdd)
    AddPresenceCode d, "713", Array(PL_7h8h, ActionAdd)
    AddPresenceCode d, "711", Array(PL_7h8h, ActionAdd)
    AddPresenceCode d, "711:30", Array(PL_7h8h, ActionAdd)
    AddPresenceCode d, "715:30", Array(PL_7h8h, ActionAdd)
    AddPresenceCode d, "7:3016", Array(PL_7h8h, ActionAdd), Array(PL_8h16h30, ActionSet)
    AddPresenceCode d, "1016:30", Array(PL_8h16h30, ActionSet)
    AddPresenceCode d, "8:3016:30", Array(PL_8h16h30, ActionSet)
    AddPresenceCode d, "C15", Array(PL_C15, ActionSet)
    AddPresenceCode d, "16:3020:15", Array(PL_C15, ActionSet)
    AddPresenceCode d, "8:3012:4516:3020:15", Array(PL_C15, ActionSet)
    AddPresenceCode d, "C20", Array(PL_C20, ActionSet)
    AddPresenceCode d, "C20E", Array(PL_C20E, ActionSet)
    AddPresenceCode d, "C19", Array(PL_7h8h, ActionAdd), Array(PL_C19, ActionSet)
    AddPresenceCode d, "C19di", Array(PL_7h8h, ActionAdd), Array(PL_C19, ActionSet)
    AddPresenceCode d, "1519", Array(PL_C19, ActionSet)
    AddPresenceCode d, "15:3019", Array(PL_C19, ActionSet)
    Set CreatePresenceCodeMapDay = d
End Function

Private Function CreatePresenceCodeMapNight() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    AddPresenceCode d, "19:456:45", Array(PL_1945, ActionAdd)
    AddPresenceCode d, "207", Array(PL_207, ActionAdd)
    Set CreatePresenceCodeMapNight = d
End Function

Private Sub ApplyPresenceAction(lineId As PresenceLine, act As PresenceAction, dayIdx As Long, _
        dayTotalsPresence6h45() As Long, dayTotalsPresence7h8h() As Long, dayTotalsPresence8h16h30() As Long, _
        dayTotalsPresenceC15() As Long, dayTotalsPresenceC20() As Long, dayTotalsPresenceC20E() As Long, _
        dayTotalsPresenceC19() As Long, dayTotalsPresence1945() As Long, dayTotalsPresence207() As Long)

    Select Case lineId
        Case PL_6h45
            If act = ActionSet Then
                dayTotalsPresence6h45(dayIdx) = 1
            Else
                dayTotalsPresence6h45(dayIdx) = dayTotalsPresence6h45(dayIdx) + 1
            End If
        Case PL_7h8h
            If act = ActionSet Then
                dayTotalsPresence7h8h(dayIdx) = 1
            Else
                dayTotalsPresence7h8h(dayIdx) = dayTotalsPresence7h8h(dayIdx) + 1
            End If
        Case PL_8h16h30
            If act = ActionSet Then
                dayTotalsPresence8h16h30(dayIdx) = 1
            Else
                dayTotalsPresence8h16h30(dayIdx) = dayTotalsPresence8h16h30(dayIdx) + 1
            End If
        Case PL_C15
            If act = ActionSet Then
                dayTotalsPresenceC15(dayIdx) = 1
            Else
                dayTotalsPresenceC15(dayIdx) = dayTotalsPresenceC15(dayIdx) + 1
            End If
        Case PL_C20
            If act = ActionSet Then
                dayTotalsPresenceC20(dayIdx) = 1
            Else
                dayTotalsPresenceC20(dayIdx) = dayTotalsPresenceC20(dayIdx) + 1
            End If
        Case PL_C20E
            If act = ActionSet Then
                dayTotalsPresenceC20E(dayIdx) = 1
            Else
                dayTotalsPresenceC20E(dayIdx) = dayTotalsPresenceC20E(dayIdx) + 1
            End If
        Case PL_C19
            If act = ActionSet Then
                dayTotalsPresenceC19(dayIdx) = 1
            Else
                dayTotalsPresenceC19(dayIdx) = dayTotalsPresenceC19(dayIdx) + 1
            End If
        Case PL_1945
            dayTotalsPresence1945(dayIdx) = dayTotalsPresence1945(dayIdx) + 1
        Case PL_207
            dayTotalsPresence207(dayIdx) = dayTotalsPresence207(dayIdx) + 1
    End Select
End Sub

        On Error GoTo 0       ' Rétablir la gestion d'erreur normale
        Exit Function         ' Sortir de la fonction
    End If
    On Error GoTo 0           ' Rétablir la gestion d'erreur normale si aucune erreur initiale

    ' Analyser le contenu de tempArray
    If IsEmpty(tempArray) Then
        ReadRangeToArray = Empty ' La plage est vide

    ' --- ATTENTION A CETTE LIGNE : ElseIf sans espace ---
    ElseIf Not IsArray(tempArray) Then ' La plage contient une seule valeur

        ' Mettre cette valeur unique dans un tableau 1x1 pour la cohérence
        Dim singleCellArray(1 To 1, 1 To 1) As Variant
        singleCellArray(1, 1) = tempArray
        ReadRangeToArray = singleCellArray
    Else ' La plage contenait plusieurs cellules, tempArray est déjà un tableau
        ReadRangeToArray = tempArray
    End If

End Function




