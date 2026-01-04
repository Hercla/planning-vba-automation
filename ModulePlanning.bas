' ExportedAt: 2026-01-04 17:02:16 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "ModulePlanning"
Option Explicit

'================================================================================================
' MODULE :          Module_Planning (Ultimate Production Version)
' DESCRIPTION :     Version V2 avec nom unique pour éviter les conflits.
'================================================================================================

' ... (Les constantes restent les mêmes, je ne les répète pas pour abréger ici, gardez celles d'avant ou copiez le bloc complet ci-dessous) ...

' --- CONSTANTES PLANNING (JOUR) ---
Private Const START_ROW As Long = 6
Private Const END_ROW As Long = 26
Private Const START_COL As Long = 3
Private Const END_COL As Long = 33

' --- CONSTANTES LIGNES TOTAUX ---
Private Const TOTAL_ROW_MATIN As Long = 60
Private Const TOTAL_ROW_APRESMIDI As Long = 61
Private Const TOTAL_ROW_SOIR As Long = 62
Private Const PRESENCE_ROW_P06H45 As Long = 64
Private Const PRESENCE_ROW_P07H8H As Long = 65
Private Const PRESENCE_ROW_P8H1630 As Long = 66
Private Const PRESENCE_ROW_C15 As Long = 67
Private Const PRESENCE_ROW_C20 As Long = 68
Private Const PRESENCE_ROW_C20E As Long = 69
Private Const PRESENCE_ROW_C19 As Long = 70

' --- CONSTANTES NUITS ---
Private Const NIGHT_SHIFT_START_ROW As Long = 31
Private Const NIGHT_SHIFT_END_ROW As Long = 38
Private Const NIGHT_CODE_1 As String = "19:45 6:45"
Private Const NIGHT_CODE_2 As String = "20 7"
Private Const PRESENCE_ROW_NIGHT_1 As Long = 71
Private Const PRESENCE_ROW_NIGHT_2 As Long = 72
Private Const TOTAL_ROW_NUIT As Long = 73

' --- CONFIGURATION PERSONNEL ---
Private Const PERSONNEL_SHEET_NAME As String = "Personnel"
Private Const PERSONNEL_COL_NOM As Long = 2
Private Const PERSONNEL_COL_PRENOM As Long = 3
Private Const PERSONNEL_COL_FONCTION As Long = 5

' --- VARIABLES GLOBALES ---
Private ignoreIfYellowOrBlue As Object
Private cfaPeople As Object
Private codeCache As Object

'================================================================================================
'   PROCÉDURE PRINCIPALE RENOMMÉE (POUR ÉVITER LE CONFLIT)
'================================================================================================
Public Sub UpdateDailyTotals_V2()
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    ' Sécurité Onglet
    If ws.Name = PERSONNEL_SHEET_NAME Or InStr(1, ws.Name, "Config", vbTextCompare) > 0 Then
        MsgBox "Stop : Impossible de lancer depuis '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If

    Dim oldCalc As XlCalculation
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Initialisation
    InitCFAList
    InitIgnoreDicts
    InitCodeCache
    
    Dim nbRows As Long: nbRows = END_ROW - START_ROW + 1
    Dim nbCols As Long: nbCols = END_COL - START_COL + 1

    Dim schedule As Variant
    schedule = ws.Range(ws.Cells(START_ROW, START_COL), ws.Cells(END_ROW, END_COL)).value

    Dim names As Variant
    names = ws.Range(ws.Cells(START_ROW, 1), ws.Cells(END_ROW, 1)).value

    Dim colIndex As Long, rowIndex As Long
    Dim rawCode As String, cleanCode As String
    Dim personRaw As String, personKey As String
    Dim totals(1 To 10) As Double
    Dim cell As Range
    Dim codeInfo As clsCodeInfo
    Dim nightVals As Variant
    Dim nVal As String
    Dim countNight1 As Double, countNight2 As Double
    Dim k As Long
    
    Dim targetNight1 As String: targetNight1 = NormalizeString(NIGHT_CODE_1)
    Dim targetNight2 As String: targetNight2 = NormalizeString(NIGHT_CODE_2)

    For colIndex = 1 To nbCols
        Dim i As Long: For i = 1 To 10: totals(i) = 0: Next i
        countNight1 = 0: countNight2 = 0

        ' BOUCLE JOUR
        For rowIndex = 1 To nbRows
            personRaw = CStr(names(rowIndex, 1))
            personKey = NormalizePersonKey(personRaw)
            
            If cfaPeople.Exists(personKey) Then
                ' CFA exclu
            Else
                rawCode = CStr(schedule(rowIndex, colIndex))
                cleanCode = NormalizeString(rawCode)
                
                If cleanCode <> "" Then
                    Set cell = ws.Cells(START_ROW + rowIndex - 1, START_COL + colIndex - 1)
                    If Not ShouldBeIgnored(cell, personKey, cleanCode) Then
                        Set codeInfo = GetCachedCodeInfo(cleanCode)
                        If codeInfo.code <> "INCONNU" Then
                            totals(1) = totals(1) + codeInfo.Fractions(1)
                            totals(2) = totals(2) + codeInfo.Fractions(2)
                            totals(3) = totals(3) + codeInfo.Fractions(3)
                            totals(4) = totals(4) + codeInfo.Fractions(5)
                            totals(5) = totals(5) + codeInfo.Fractions(6)
                            totals(6) = totals(6) + codeInfo.Fractions(7)
                            totals(7) = totals(7) + codeInfo.Fractions(8)
                            totals(8) = totals(8) + codeInfo.Fractions(9)
                            totals(9) = totals(9) + codeInfo.Fractions(10)
                            totals(10) = totals(10) + codeInfo.Fractions(11)
                        End If
                    End If
                End If
            End If
        Next rowIndex

        ' BOUCLE NUITS
        nightVals = ws.Range(ws.Cells(NIGHT_SHIFT_START_ROW, START_COL + colIndex - 1), _
                             ws.Cells(NIGHT_SHIFT_END_ROW, START_COL + colIndex - 1)).value
        For k = 1 To UBound(nightVals, 1)
            nVal = NormalizeString(CStr(nightVals(k, 1)))
            If nVal = targetNight1 Then
                countNight1 = countNight1 + 1
            ElseIf nVal = targetNight2 Then
                countNight2 = countNight2 + 1
            End If
        Next k

        WriteTotalsToSheet ws, START_COL + colIndex - 1, totals, countNight1, countNight2
    Next colIndex

    Set codeCache = Nothing
    Set cfaPeople = Nothing
    Set ignoreIfYellowOrBlue = Nothing

    Application.ScreenUpdating = True
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    MsgBox "Mise à jour V2 terminée !", vbInformation
End Sub

'================================================================================================
'   FONCTIONS SUPPORT (COPIÉES-COLLÉES DE LA VERSION PRÉCÉDENTE)
'================================================================================================

Private Function NormalizePersonKey(ByVal s As String) As String
    If Len(s) = 0 Then Exit Function
    s = Replace(s, Chr(160), " ")
    s = UCase$(Trim$(s))
    s = Replace(s, "-", "_")
    s = Replace(s, " ", "_")
    Do While InStr(s, "__") > 0
        s = Replace(s, "__", "_")
    Loop
    NormalizePersonKey = s
End Function

Private Function NormalizeString(ByVal s As String) As String
    If Len(s) = 0 Then Exit Function
    s = Replace(s, Chr(160), " ")
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeString = s
End Function

Private Sub InitCFAList()
    Set cfaPeople = CreateObject("Scripting.Dictionary")
    cfaPeople.CompareMode = vbTextCompare
    Dim wsP As Worksheet
    On Error Resume Next
    Set wsP = ThisWorkbook.Worksheets(PERSONNEL_SHEET_NAME)
    On Error GoTo 0
    If wsP Is Nothing Then Exit Sub
    Dim lastRow As Long
    lastRow = wsP.Cells(wsP.Rows.Count, PERSONNEL_COL_NOM).End(xlUp).row
    If lastRow < 2 Then Exit Sub
    Dim arr As Variant, i As Long, nom As String, prenom As String, func As String, key As String
    arr = wsP.Range("B2:E" & lastRow).value
    For i = 1 To UBound(arr, 1)
        func = NormalizeString(CStr(arr(i, 4)))
        If UCase(func) = "CFA" Then
            nom = CStr(arr(i, 1))
            prenom = CStr(arr(i, 2))
            key = NormalizePersonKey(nom & "_" & prenom)
            cfaPeople(key) = True
            key = NormalizePersonKey(nom & " " & prenom)
            cfaPeople(key) = True
        End If
    Next i
End Sub

Private Sub InitCodeCache()
    Set codeCache = CreateObject("Scripting.Dictionary")
    codeCache.CompareMode = vbTextCompare
End Sub

Private Function GetCachedCodeInfo(ByVal code As String) As clsCodeInfo
    If Not codeCache.Exists(code) Then
        Set codeCache(code) = GetCodeInfo(code)
    End If
    Set GetCachedCodeInfo = codeCache(code)
End Function

Private Sub InitIgnoreDicts()
    Set ignoreIfYellowOrBlue = CreateObject("Scripting.Dictionary")
    ignoreIfYellowOrBlue.CompareMode = vbTextCompare
    ignoreIfYellowOrBlue("BOURGEOIS_AURORE|7 15:30") = True
    ignoreIfYellowOrBlue("BOURGEOIS_AURORE|6:45 15:15") = True
    ignoreIfYellowOrBlue("DIALLO_MAMADOU|7 15:30") = True
    ignoreIfYellowOrBlue("DIALLO_MAMADOU|6:45 15:15") = True
    ignoreIfYellowOrBlue("DELA VEGA_EDELYN|7 15:30") = True
    ignoreIfYellowOrBlue("DELA VEGA_EDELYN|6:45 15:15") = True
End Sub

Private Function ShouldBeIgnored(ByVal cell As Range, ByVal normPersonKey As String, ByVal code As String) As Boolean
    Dim key As String: key = normPersonKey & "|" & code
    If Not ignoreIfYellowOrBlue.Exists(key) Then
        ShouldBeIgnored = False
        Exit Function
    End If
    If IsYellow(cell) Or IsLightBlue(cell) Then
        ShouldBeIgnored = True
    Else
        ShouldBeIgnored = False
    End If
End Function

Private Function IsYellow(c As Range) As Boolean
    IsYellow = (c.Interior.Color = vbYellow) Or (c.Interior.ColorIndex = 6)
End Function

Private Function IsLightBlue(c As Range) As Boolean
    On Error Resume Next
    Dim themec As Long: themec = c.Interior.ThemeColor
    Dim tint As Double: tint = c.Interior.TintAndShade
    Dim idx As Long: idx = c.Interior.ColorIndex
    Dim rgbv As Long: rgbv = c.Interior.Color
    IsLightBlue = (themec = xlThemeColorAccent1 And tint > 0) _
                  Or (idx = 37 Or idx = 34 Or idx = 41) _
                  Or (rgbv = RGB(221, 235, 247) Or rgbv = RGB(204, 232, 255) Or rgbv = RGB(198, 239, 255))
End Function

Private Sub WriteTotalsToSheet(ByVal ws As Worksheet, ByVal col As Long, ByRef totals() As Double, ByVal n1 As Double, ByVal n2 As Double)
    ws.Cells(TOTAL_ROW_MATIN, col).value = IIf(totals(1) > 0, totals(1), "")
    ws.Cells(TOTAL_ROW_APRESMIDI, col).value = IIf(totals(2) > 0, totals(2), "")
    ws.Cells(TOTAL_ROW_SOIR, col).value = IIf(totals(3) > 0, totals(3), "")
    ws.Cells(PRESENCE_ROW_P06H45, col).value = IIf(totals(4) > 0, totals(4), "")
    ws.Cells(PRESENCE_ROW_P07H8H, col).value = IIf(totals(5) > 0, totals(5), "")
    ws.Cells(PRESENCE_ROW_P8H1630, col).value = IIf(totals(6) > 0, totals(6), "")
    ws.Cells(PRESENCE_ROW_C15, col).value = IIf(totals(7) > 0, totals(7), "")
    ws.Cells(PRESENCE_ROW_C20, col).value = IIf(totals(8) > 0, totals(8), "")
    ws.Cells(PRESENCE_ROW_C20E, col).value = IIf(totals(9) > 0, totals(9), "")
    ws.Cells(PRESENCE_ROW_C19, col).value = IIf(totals(10) > 0, totals(10), "")
    ws.Cells(PRESENCE_ROW_NIGHT_1, col).value = IIf(n1 > 0, n1, "")
    ws.Cells(PRESENCE_ROW_NIGHT_2, col).value = IIf(n2 > 0, n2, "")
    ws.Cells(TOTAL_ROW_NUIT, col).value = IIf((n1 + n2) > 0, n1 + n2, "")
End Sub

Sub Afficher_cacher_menu()
    If UserForm1.Visible Then
        UserForm1.Hide
        Exit Sub
    End If
    With UserForm1
        .height = 509.25
        .width = 201.75
        .StartUpPosition = 0
        .Left = Application.Left + Application.width - .width - 25
        .Top = Application.Top + 50
        .Show vbModeless
    End With
End Sub

Public Sub InitOngletRoulement()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Roulements")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    With ws
        .Activate
        .Columns("B").ColumnWidth = 20.95
        .Columns("C:BG").ColumnWidth = 4.95
        ActiveWindow.Zoom = 50
        .Cells(1, 1).Select
    End With
    Application.ScreenUpdating = True
End Sub

