Attribute VB_Name = "avecnuit"
Sub CalculateAllShiftsAllSheetsOptimized_Final_WithNight_Modified()
    Dim wsListe As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim shiftData As Variant
    Dim shiftDict As Object
    Dim dayTotalsMatin() As Long
    Dim dayTotalsApresMidi() As Long
    Dim dayTotalsSoir() As Long
    Dim dayTotalsNuit() As Long
    Dim dayTotalsSpecial() As Long
    Dim dayTotalsSpecial2() As Long
    Dim dayTotalsFraction() As Long
    Dim dayTotalsFractionC20E() As Long
    Dim dayTotalsFractionC19() As Long
    Dim dayTotalsPresence645() As Long
    Dim dayTotalsPresence81630() As Long
    Dim dayTotalsC15() As Long
    Dim dayTotalsC20() As Long
    Dim shiftSchedule As Variant
    Dim shiftCode As String
    Dim shiftInfo As Variant
    Dim rowIdx As Long, colIdx As Long
    Dim shiftAssignments As Variant
    Dim i As Long

    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Définir la feuille "Liste" et initialiser le dictionnaire
    Set wsListe = ThisWorkbook.Sheets("Liste")
    Set shiftDict = CreateObject("Scripting.Dictionary")
    
    ' Charger les données de quart depuis la feuille "Liste"
    lastRow = wsListe.Cells(wsListe.rows.Count, 1).End(xlUp).row
    shiftData = wsListe.Range("A2:G" & lastRow).value
    
    ' Remplir le dictionnaire avec les assignations des quarts
    For i = 1 To UBound(shiftData, 1)
        shiftCode = Trim(CStr(shiftData(i, 1)))
        If shiftCode <> "" Then
            shiftAssignments = Array(shiftData(i, 4) > 0, shiftData(i, 5) > 0, shiftData(i, 6) > 0, shiftData(i, 7) > 0)
            shiftDict(shiftCode) = shiftAssignments
        End If
    Next i
    
    ' Parcourir les feuilles du classeur
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Janv*" Or ws.Name Like "Fev" Or ws.Name Like "Mars" Or _
           ws.Name Like "Avril" Or ws.Name Like "Mai" Or ws.Name Like "Juin" Or _
           ws.Name Like "Juillet" Or ws.Name Like "Aout" Or ws.Name Like "Sept" Or _
           ws.Name Like "Oct" Or ws.Name Like "Nov" Or ws.Name Like "Dec" Then
            
            ' Réinitialiser les totaux journaliers pour 31 jours
            ReDim dayTotalsMatin(1 To 31)
            ReDim dayTotalsApresMidi(1 To 31)
            ReDim dayTotalsSoir(1 To 31)
            ReDim dayTotalsNuit(1 To 31)
            ReDim dayTotalsSpecial(1 To 31)
            ReDim dayTotalsSpecial2(1 To 31)
            ReDim dayTotalsFraction(1 To 31)
            ReDim dayTotalsFractionC20E(1 To 31)
            ReDim dayTotalsFractionC19(1 To 31)
            ReDim dayTotalsPresence645(1 To 31)
            ReDim dayTotalsPresence81630(1 To 31)
            ReDim dayTotalsC15(1 To 31)
            ReDim dayTotalsC20(1 To 31)
            
            ' Charger le planning des quarts pour la feuille actuelle
            shiftSchedule = ws.Range(ws.Cells(6, 2), ws.Cells(26, 32)).value
            
            ' Parcourir le planning des quarts
            For rowIdx = 1 To UBound(shiftSchedule, 1)
                For colIdx = 1 To UBound(shiftSchedule, 2)
                    shiftCode = Trim(CStr(shiftSchedule(rowIdx, colIdx)))
                    
                    If shiftDict.Exists(shiftCode) Then
                        shiftInfo = shiftDict(shiftCode)
                        If shiftInfo(0) Then dayTotalsMatin(colIdx) = dayTotalsMatin(colIdx) + 1
                        If shiftInfo(1) Then dayTotalsApresMidi(colIdx) = dayTotalsApresMidi(colIdx) + 1
                        If shiftInfo(2) Then dayTotalsSoir(colIdx) = dayTotalsSoir(colIdx) + 1
                    End If
                    
                    ' Calcul des fractions
                    If shiftCode Like "7*" Or shiftCode Like "6*" Then
                        dayTotalsFraction(colIdx) = dayTotalsFraction(colIdx) + 1
                    End If
                    
                    ' Comptabiliser "C 20" distinctement de "C 20 E"
                    If shiftCode = "C 20" Then
                        dayTotalsC20(colIdx) = dayTotalsC20(colIdx) + 1
                    End If
                    
                    If shiftCode = "C 20 E" Then
                        dayTotalsFractionC20E(colIdx) = dayTotalsFractionC20E(colIdx) + 1
                    End If
                    
                    If shiftCode = "C 19" Or shiftCode = "15 19" Or shiftCode = "15:30 19" Or shiftCode = "C 19 di" Then
                        dayTotalsFractionC19(colIdx) = dayTotalsFractionC19(colIdx) + 1
                    End If
                    
                    ' Modifier la logique pour les horaires spéciaux
                    If shiftCode = "6:45 15:15" Or shiftCode = "6:45 12:45" Then
                        dayTotalsSpecial(colIdx) = dayTotalsSpecial(colIdx) + 1
                    End If
                    
                    ' Comptabiliser la présence sur 8-16/16:30
                    If shiftCode = "8 16:30" Or shiftCode = "8:30 16" Or shiftCode = "8:30 16:30" Or shiftCode = "8 16" Then
                        dayTotalsPresence81630(colIdx) = dayTotalsPresence81630(colIdx) + 1
                    End If
                    
                    If shiftCode = "C 15" Or shiftCode = "16:30 20:15" Then
                        dayTotalsC15(colIdx) = dayTotalsC15(colIdx) + 1
                    End If
                    
                Next colIdx
            Next rowIdx
            
            ' Écrire les totaux sur la feuille en évitant les chevauchements :
            ws.Range("B60:AF60").value = dayTotalsMatin
            ws.Range("B61:AF61").value = dayTotalsApresMidi
            ws.Range("B62:AF62").value = dayTotalsSoir
            
            ws.Range("B106:AF106").value = dayTotalsSpecial
            ws.Range("B107:AF107").value = dayTotalsSpecial2
            ws.Range("B109:AF109").value = dayTotalsPresence645
            ws.Range("B110:AF110").value = dayTotalsPresence81630
            ws.Range("B111:AF111").value = dayTotalsFractionC20E
            ws.Range("B112:AF112").value = dayTotalsFractionC19
            ws.Range("B113:AF113").value = dayTotalsC15
            ws.Range("B114:AF114").value = dayTotalsC20
            
        End If
    Next ws
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Calculs terminés !"
    
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Une erreur est survenue : " & Err.Description
End Sub


