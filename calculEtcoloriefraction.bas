Attribute VB_Name = "calculEtcoloriefraction"
Option Explicit

Sub CalculateAndColorAllSheets_Combined()
    Dim wsListe As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim shiftData As Variant
    Dim shiftDict As Object
    Dim shiftSchedule As Variant
    Dim shiftCode As String
    Dim shiftInfo As Variant
    Dim shiftAssignments As Variant

    Dim rowIdx As Long, colIdx As Long
    Dim i As Long

    ' Tableaux de totaux journaliers (pour 31 jours)
    Dim dayTotalsMatin() As Long
    Dim dayTotalsApresMidi() As Long
    Dim dayTotalsSoir() As Long
    Dim dayTotalsNuit() As Long

    ' === Définir les codes pour chaque shift ===
    Dim codesMatin As Variant, codesApresMidi As Variant, codesSoir As Variant, codesNuit As Variant
    codesMatin = Array( _
        "7 15:30", "6:45 15:15", "6:45 12:45", "7 13", "7 11:30", "7:15 15:45", _
        "C 19", "C 19 di", "C 15", "C 15 di", "8:30 12:45 16:30 20:15", _
        "C 20 E", "8 11:30", "8 16:30", "C 20", "8:30 14", "8:30 16:30", "7:30 16" _
    )
    codesApresMidi = Array( _
        "7 15:30", "6:45 15:15", "8 14", _
        "8:30 14:30", "8 16:30", "8:30 16:30", "7:15 15:45", "13 19", "8:30 14", "7:30 16" _
    )
    codesSoir = Array( _
        "C 15", "C 19", "C 20 E", "13 19", "16 20", "16:30 20:15", _
        "C 20", "8:30 12:45 16:30 20:15", "C 15 di", "C 19 di" _
    )
    codesNuit = Array( _
        "19:45 6:45", "20 7" _
    )

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 1) Chargement de la feuille "Liste" et constitution du dictionnaire
    Set wsListe = ThisWorkbook.Sheets("Liste")
    Set shiftDict = CreateObject("Scripting.Dictionary")

    lastRow = wsListe.Cells(wsListe.rows.Count, 1).End(xlUp).row
    shiftData = wsListe.Range("A2:G" & lastRow).value

    For i = 1 To UBound(shiftData, 1)
        shiftCode = Trim(CStr(shiftData(i, 1)))
        If shiftCode <> "" Then
            ' shiftAssignments = { Matin, Aprem, Soir, Nuit }
            shiftAssignments = Array( _
                (shiftData(i, 4) > 0), _
                (shiftData(i, 5) > 0), _
                (shiftData(i, 6) > 0), _
                (shiftData(i, 7) > 0) _
            )
            shiftDict(shiftCode) = shiftAssignments
        End If
    Next i

    ' 2) Parcours des feuilles concernées
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Janv*" Or ws.Name Like "Fev*" Or ws.Name Like "Mars*" Or _
           ws.Name Like "Avril*" Or ws.Name Like "Mai*" Or ws.Name Like "Juin*" Or _
           ws.Name Like "Juillet*" Or ws.Name Like "Aout*" Or ws.Name Like "Sept*" Or _
           ws.Name Like "Oct*" Or ws.Name Like "Nov*" Or ws.Name Like "Dec*" Or _
           ws.Name Like "JanvB" Or ws.Name Like "FevB" Then

            ' Initialiser les tableaux pour 31 jours
            ReDim dayTotalsMatin(1 To 31)
            ReDim dayTotalsApresMidi(1 To 31)
            ReDim dayTotalsSoir(1 To 31)
            ReDim dayTotalsNuit(1 To 31)

            ' Lire l'ensemble du planning des quarts
            shiftSchedule = ws.Range(ws.Cells(6, 2), ws.Cells(26, 32)).value

            ' Boucler à travers chaque cellule du planning des quarts
            For rowIdx = 1 To UBound(shiftSchedule, 1)
                For colIdx = 1 To UBound(shiftSchedule, 2)
                    Dim code As String, codeMatin As Variant, codeAprem As Variant, codeSoir As Variant, codeNuit As Variant
                    code = Trim(CStr(shiftSchedule(rowIdx, colIdx)))
                    If code <> "" Then
                        ' --- Ajout : ignorer si la cellule est surlignée jaune vif ou bleu clair ---
                        Dim cellColor As Long
                        cellColor = ws.Cells(rowIdx + 5, colIdx + 1).Interior.Color ' +5 car la plage commence à la ligne 6, +1 car colonne B=2
                        If cellColor = RGB(255, 255, 0) Or cellColor = RGB(204, 255, 255) Then
                            ' Jaune vif ou bleu clair : on ignore cette cellule
                            GoTo NextCell
                        End If
                        ' --- Fin ajout ---

                        ' Matin
                        For Each codeMatin In codesMatin
                            If StrComp(code, codeMatin, vbTextCompare) = 0 Then
                                dayTotalsMatin(colIdx) = dayTotalsMatin(colIdx) + 1
                                Exit For
                            End If
                        Next codeMatin
                        ' Après-midi
                        For Each codeAprem In codesApresMidi
                            If StrComp(code, codeAprem, vbTextCompare) = 0 Then
                                dayTotalsApresMidi(colIdx) = dayTotalsApresMidi(colIdx) + 1
                                Exit For
                            End If
                        Next codeAprem
                        ' Soir
                        For Each codeSoir In codesSoir
                            If StrComp(code, codeSoir, vbTextCompare) = 0 Then
                                dayTotalsSoir(colIdx) = dayTotalsSoir(colIdx) + 1
                                Exit For
                            End If
                        Next codeSoir
                        ' Nuit
                        For Each codeNuit In codesNuit
                            If StrComp(code, codeNuit, vbTextCompare) = 0 Then
                                dayTotalsNuit(colIdx) = dayTotalsNuit(colIdx) + 1
                                Exit For
                            End If
                        Next codeNuit
                    End If
NextCell:
                Next colIdx
            Next rowIdx

            ' Écriture des totaux dans la feuille
            ws.Range("B60:AF60").value = dayTotalsMatin
            ws.Range("B61:AF61").value = dayTotalsApresMidi
            ws.Range("B62:AF62").value = dayTotalsSoir
            ws.Range("B63:AF63").value = dayTotalsNuit
        End If
    Next ws

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Calculs et coloration terminés !"
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Une erreur est survenue : " & Err.Description
End Sub

