Attribute VB_Name = "ModuleGestionAbsences"
Option Explicit

' --- Constants ---
Private Const WS_ABSENCES As String = "Absences"
Private Const WS_SUMMARY As String = "Résumé Absences"
Private Const ABS_CA As String = "CA"
Private Const ABS_CSOC As String = "C SOC"
Private Const ABS_PC As String = "PETIT CHOM"
Private Const ABS_MAL As String = "MAL"

' Remplace les accents et met en majuscules pour faciliter les comparaisons
Private Function NormalizeText(text As String) As String
    text = UCase(text)
    text = Replace(text, "À", "A")
    text = Replace(text, "Â", "A")
    text = Replace(text, "Ä", "A")
    text = Replace(text, "Æ", "AE")
    text = Replace(text, "Ç", "C")
    text = Replace(text, "É", "E")
    text = Replace(text, "È", "E")
    text = Replace(text, "Ê", "E")
    text = Replace(text, "Ë", "E")
    text = Replace(text, "Î", "I")
    text = Replace(text, "Ï", "I")
    text = Replace(text, "Ô", "O")
    text = Replace(text, "Ö", "O")
    text = Replace(text, "Ù", "U")
    text = Replace(text, "Û", "U")
    text = Replace(text, "Ü", "U")
    text = Replace(text, "à", "A")
    text = Replace(text, "â", "A")
    text = Replace(text, "ä", "A")
    text = Replace(text, "ç", "C")
    text = Replace(text, "é", "E")
    text = Replace(text, "è", "E")
    text = Replace(text, "ê", "E")
    text = Replace(text, "ë", "E")
    text = Replace(text, "î", "I")
    text = Replace(text, "ï", "I")
    text = Replace(text, "ô", "O")
    text = Replace(text, "ö", "O")
    text = Replace(text, "ù", "U")
    text = Replace(text, "û", "U")
    text = Replace(text, "ü", "U")
    NormalizeText = text
End Function


' ===================================================================
'               *** BLOC DE CORRECTION INTÉGRÉ ICI ***
'   Ce bloc déclare les variables pour les numéros de colonne,
'   ce qui corrige l'erreur "Variable non définie".
' ===================================================================

' Indices des colonnes de l'onglet Absences via une Énumération
Public Enum AbsColumns
    COL_NOM = 1
    COL_TYPE
    COL_DEBUT
    COL_FIN
    COL_JOURS
    COL_COMMENT
End Enum

' (Optionnel) Alias en minuscules pour compatibilité avec d'anciens modules
' Le code actuel utilise les membres de l'Enum directement (ex: COL_NOM)
' Public Const colNom As Long = AbsColumns.COL_NOM
' Public Const colType As Long = AbsColumns.COL_TYPE
' Public Const colDebut As Long = AbsColumns.COL_DEBUT
' Public Const colFin As Long = AbsColumns.COL_FIN
' Public Const colJours As Long = AbsColumns.COL_JOURS
' Public Const colComment As Long = AbsColumns.COL_COMMENT

' Ajoute une nouvelle ligne d'absence (congé ou maladie)
Sub AjouterAbsence()
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim nomEmp As String
    Dim typ As String
    Dim dDebut As Variant
    Dim dFin As Variant
    Dim jours As Double

    Set ws = GetOrCreateAbsenceSheet()

    Dim prefix As Variant
    prefix = Application.InputBox("Premières lettres du nom ou prénom :", _
                                  "Nouvelle absence", Type:=2)
    If prefix = False Or prefix = "" Then Exit Sub
    nomEmp = SelectEmployeeByPrefix(CStr(prefix), GetEmployeeList())
    If nomEmp = "" Then Exit Sub
    typ = Application.InputBox("Type d'absence (" & ABS_CA & " / " & ABS_CSOC & " / " & ABS_PC & " / " & ABS_MAL & ") :", _
                               "Nouvelle absence", ABS_CA, Type:=2)
    If typ = False Then Exit Sub
    typ = UCase(Trim(CStr(typ)))
    If typ <> ABS_CA And typ <> ABS_CSOC And typ <> ABS_PC And typ <> ABS_MAL Then
        MsgBox "Type d'absence invalide", vbExclamation
        Exit Sub
    End If
    dDebut = Application.InputBox("Date de début :", "Nouvelle absence", Type:=1)
    If dDebut = False Then Exit Sub
    dFin = Application.InputBox("Date de fin :", "Nouvelle absence", Type:=1)
    If dFin = False Then Exit Sub
    If dFin < dDebut Then
        MsgBox "La date de fin doit être supérieure ou égale à la date de début.", vbExclamation
        Exit Sub
    End If
    jours = dFin - dDebut + 1

    nextRow = ws.Cells(ws.Rows.Count, COL_NOM).End(xlUp).Row + 1
    ws.Cells(nextRow, COL_NOM).Value = nomEmp
    ws.Cells(nextRow, COL_TYPE).Value = UCase(typ)
    ws.Cells(nextRow, COL_DEBUT).Value = CDate(dDebut)
    ws.Cells(nextRow, COL_FIN).Value = CDate(dFin)
    ws.Cells(nextRow, COL_JOURS).Value = jours

    MsgBox "Absence enregistrée", vbInformation
End Sub

' Génère ou met à jour le résumé des absences par employé
Sub RafraichirResumeAbsences()
    Dim wsAbs As Worksheet
    Dim wsRes As Worksheet
    Dim dict As Object
    Dim lastRow As Long
    Dim i As Long
    Dim nom As String
    Dim t As String
    Dim jours As Double
    Dim arr

    On Error Resume Next
    Set wsAbs = ThisWorkbook.Sheets(WS_ABSENCES)
    On Error GoTo 0
    If wsAbs Is Nothing Then
        MsgBox "La feuille '" & WS_ABSENCES & "' est introuvable.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wsRes = ThisWorkbook.Sheets(WS_SUMMARY)
    On Error GoTo 0
    If wsRes Is Nothing Then
        Set wsRes = ThisWorkbook.Sheets.Add(After:=wsAbs)
        wsRes.Name = WS_SUMMARY
    Else
        wsRes.Cells.Clear
    End If
    wsRes.Range("A1:E1").Value = Array("Employé", ABS_CA, "C Soc", ABS_PC, ABS_MAL)

    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = wsAbs.Cells(wsAbs.Rows.Count, COL_NOM).End(xlUp).Row
    For i = 2 To lastRow
        nom = wsAbs.Cells(i, COL_NOM).Value
        t = UCase(wsAbs.Cells(i, COL_TYPE).Value)
        jours = wsAbs.Cells(i, COL_JOURS).Value
        If Not dict.Exists(nom) Then
            dict.Add nom, Array(0#, 0#, 0#, 0#)
        End If
        Select Case t
            Case ABS_CA
                dict(nom)(0) = dict(nom)(0) + jours
            Case ABS_CSOC
                dict(nom)(1) = dict(nom)(1) + jours
            Case ABS_PC
                dict(nom)(2) = dict(nom)(2) + jours
            Case ABS_MAL
                dict(nom)(3) = dict(nom)(3) + jours
        End Select
    Next i

    arr = dict.Keys
    For i = 0 To dict.Count - 1
        wsRes.Cells(i + 2, 1).Value = arr(i)
        wsRes.Cells(i + 2, 2).Resize(1, 4).Value = dict(arr(i))
    Next i
    wsRes.Columns("A:E").AutoFit
    MsgBox "Résumé mis à jour", vbInformation
End Sub

' Renvoie la feuille des absences, en la créant si besoin
Private Function GetOrCreateAbsenceSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_ABSENCES)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = WS_ABSENCES
        ws.Range("A1:F1").Value = Array("Nom", "Type", "Début", "Fin", "Jours", "Commentaire")
    End If
    Set GetOrCreateAbsenceSheet = ws
End Function

' Trouve la ligne contenant "Nom" dans la colonne A du planning
Private Function GetNomHeaderRow(ws As Worksheet) As Long
    Dim r As Variant
    On Error Resume Next
    r = Application.Match("Nom", ws.Columns(1), 0)
    On Error GoTo 0
    If IsError(r) Then
        GetNomHeaderRow = 0
    Else
        GetNomHeaderRow = CLng(r)
    End If
End Function

' Identifie la ligne o les dates commencent dans le planning
Private Function GetDateHeaderRow(ws As Worksheet) As Long
    Dim i As Long, j As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To 10
        For j = 1 To lastCol
            If IsDate(ws.Cells(i, j).Value) Then
                GetDateHeaderRow = i
                Exit Function
            End If
        Next j
    Next i
    GetDateHeaderRow = 0
End Function

' Liste les noms d'employé depuis la feuille "Planning" en détectant
' automatiquement la ligne d'entête "Nom"
Private Function GetEmployeeList() As Variant
    Dim ws As Worksheet, lastRow As Long, dict As Object, i As Long
    Dim startRow As Long, hdrRow As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Planning")
    On Error GoTo 0
    If ws Is Nothing Then
        GetEmployeeList = Array()
        Exit Function
    End If
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    hdrRow = Application.Match("Nom", ws.Columns(1), 0)
    If IsError(hdrRow) Then
        startRow = 6
    Else
        startRow = hdrRow + 1
    End If
    Set dict = CreateObject("Scripting.Dictionary")
    For i = startRow To lastRow
        Dim nm As String
        nm = Trim(CStr(ws.Cells(i, 1).Value))
        If nm <> "" Then
            If Not dict.Exists(nm) Then dict.Add nm, 1
        End If
    Next i
    If dict.Count = 0 Then
        GetEmployeeList = Array()
    Else
        GetEmployeeList = dict.Keys
    End If
End Function

' Sélectionne un employé en comparant un préfixe sur le nom ou le prénom
Private Function SelectEmployeeByPrefix(prefix As String, listNoms As Variant) As String
    Dim matches As Object, nm As Variant
    Dim arr() As String, i As Long
    Dim normPrefix As String, nmNorm As String, found As Boolean
    normPrefix = NormalizeText(prefix)
    Set matches = CreateObject("Scripting.Dictionary")
    For Each nm In listNoms
        nmNorm = NormalizeText(CStr(nm))
        found = False
        arr = Split(nmNorm)
        For i = LBound(arr) To UBound(arr)
            If Left(arr(i), Len(normPrefix)) = normPrefix Then
                found = True
                Exit For
            End If
        Next i
        If Not found Then
            If InStr(1, nmNorm, normPrefix, vbBinaryCompare) > 0 Then found = True
        End If
        If found Then
            If Not matches.Exists(nm) Then matches.Add nm, 1
        End If
    Next nm
    If matches.Count = 0 Then
        MsgBox "Aucun employé trouvé pour " & prefix, vbExclamation
    ElseIf matches.Count = 1 Then
        SelectEmployeeByPrefix = matches.Keys()(0)
    Else
        SelectEmployeeByPrefix = Application.InputBox("Sélectionnez le nom :" & _
                                         vbCrLf & Join(matches.Keys(), vbCrLf), _
                                         "Choix employé", Type:=2)
        If SelectEmployeeByPrefix = False Then SelectEmployeeByPrefix = ""
    End If
End Function

' Dtermine si un code existant doit tre remplac par "MAL"
Private Function ShouldMarkMaladie(ByVal code As String) As Boolean
    Dim c As String
    c = UCase(Trim(CStr(code)))
    If c = "" Then Exit Function
    If c = "WE" Or c = "/" Then Exit Function
    If Left(c, 1) = "R" Then Exit Function
    If Left(c, 3) = "3/4" Or Left(c, 3) = "4/5" Then Exit Function
    ShouldMarkMaladie = True
End Function

' Encode une priode de maladie dans le planning et enregistre dans l'onglet Absences
' Les lignes Nom et dates sont dtectes automatiquement.
Sub EncoderMaladiePlanning()
    Dim noms As Variant, choix As Variant, dDeb As Variant, dFin As Variant
    Dim wsPlan As Worksheet, wsAbs As Worksheet
    Dim i As Long, col As Long, rowEmp As Long, nMarked As Long

    noms = GetEmployeeList()
    If UBound(noms) < LBound(noms) Then
        MsgBox "Aucun employ trouv dans la feuille Planning.", vbExclamation
        Exit Sub
    End If

    Dim prefix As Variant
    prefix = Application.InputBox("Premières lettres du nom ou prénom :", _
                                  "Encoder Maladie", Type:=2)
    If prefix = False Or prefix = "" Then Exit Sub
    choix = SelectEmployeeByPrefix(CStr(prefix), noms)
    If choix = "" Then Exit Sub

    dDeb = Application.InputBox("Date de dbut :", "Encoder Maladie", Type:=1)
    If dDeb = False Then Exit Sub
    dFin = Application.InputBox("Date de fin :", "Encoder Maladie", Type:=1)
    If dFin = False Then Exit Sub
    If dFin < dDeb Then
        MsgBox "La date de fin doit tre suprieure  la date de dbut", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wsPlan = ThisWorkbook.Sheets("Planning")
    Set wsAbs = GetOrCreateAbsenceSheet()
    On Error GoTo 0
    If wsPlan Is Nothing Then
        MsgBox "Feuille 'Planning' introuvable", vbExclamation
        Exit Sub
    End If

    rowEmp = -1
    Dim lastRow As Long, startRow As Long, hdrRow As Long
    lastRow = wsPlan.Cells(wsPlan.Rows.Count, 1).End(xlUp).Row
    hdrRow = GetNomHeaderRow(wsPlan)
    If hdrRow = 0 Then
        startRow = 6
    Else
        startRow = hdrRow + 1
    End If
    For i = startRow To lastRow
        If Trim(CStr(wsPlan.Cells(i, 1).Value)) = choix Then
            rowEmp = i
            Exit For
        End If
    Next i
    If rowEmp = -1 Then
        MsgBox "Employ non trouv dans le planning", vbExclamation
        Exit Sub
    End If

    nMarked = 0
    Dim lastCol As Long, dateRow As Long
    dateRow = GetDateHeaderRow(wsPlan)
    If dateRow = 0 Then dateRow = 4
    lastCol = wsPlan.Cells(dateRow, wsPlan.Columns.Count).End(xlToLeft).Column
    For col = 1 To lastCol
        If IsDate(wsPlan.Cells(dateRow, col).Value) Then
            If wsPlan.Cells(dateRow, col).Value >= CDate(dDeb) And _
               wsPlan.Cells(dateRow, col).Value <= CDate(dFin) Then
                If ShouldMarkMaladie(wsPlan.Cells(rowEmp, col).Value) Then
                    wsPlan.Cells(rowEmp, col).Value = "MAL"
                    nMarked = nMarked + 1
                End If
            End If
        End If
    Next col

    If nMarked > 0 Then
        Dim nr As Long
        nr = wsAbs.Cells(wsAbs.Rows.Count, COL_NOM).End(xlUp).Row + 1
        wsAbs.Cells(nr, COL_NOM).Value = choix
        wsAbs.Cells(nr, COL_TYPE).Value = ABS_MAL
        wsAbs.Cells(nr, COL_DEBUT).Value = CDate(dDeb)
        wsAbs.Cells(nr, COL_FIN).Value = CDate(dFin)
        wsAbs.Cells(nr, COL_JOURS).Value = nMarked
    End If

    MsgBox nMarked & " jour(s) marques comme MAL", vbInformation
End Sub

' Cree/rafrachit une vue des jours maladie par priode 1,3,6 et 12 mois
Sub GenererVueMaladies()
    Dim wsAbs As Worksheet, wsVue As Worksheet
    Dim d1 As Date, d3 As Date, d6 As Date, d12 As Date
    Dim dict As Object, arr, i As Long, nom As String
    Dim dDeb As Date, dFin As Date

    Set wsAbs = GetOrCreateAbsenceSheet()
    On Error Resume Next
    Set wsVue = ThisWorkbook.Sheets("Onglet Absence")
    On Error GoTo 0
    If wsVue Is Nothing Then
        Set wsVue = ThisWorkbook.Sheets.Add(After:=wsAbs)
        wsVue.Name = "Onglet Absence"
    Else
        wsVue.Cells.Clear
    End If

    d1 = DateAdd("m", -1, Date)
    d3 = DateAdd("m", -3, Date)
    d6 = DateAdd("m", -6, Date)
    d12 = DateAdd("m", -12, Date)

    Set dict = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long
    lastRow = wsAbs.Cells(wsAbs.Rows.Count, COL_NOM).End(xlUp).Row
    For i = 2 To lastRow
        If UCase(wsAbs.Cells(i, COL_TYPE).Value) = ABS_MAL Then
            nom = wsAbs.Cells(i, COL_NOM).Value
            dDeb = wsAbs.Cells(i, COL_DEBUT).Value
            dFin = wsAbs.Cells(i, COL_FIN).Value
            If dFin >= d12 Then
                If Not dict.Exists(nom) Then dict.Add nom, Array(0, 0, 0, 0)
                If dFin >= d1 Then dict(nom)(0) = dict(nom)(0) + _
                        Application.Max(0, Application.Min(dFin, Date) - _
                        Application.Max(dDeb, d1) + 1)
                If dFin >= d3 Then dict(nom)(1) = dict(nom)(1) + _
                        Application.Max(0, Application.Min(dFin, Date) - _
                        Application.Max(dDeb, d3) + 1)
                If dFin >= d6 Then dict(nom)(2) = dict(nom)(2) + _
                        Application.Max(0, Application.Min(dFin, Date) - _
                        Application.Max(dDeb, d6) + 1)
                dict(nom)(3) = dict(nom)(3) + _
                        Application.Max(0, Application.Min(dFin, Date) - _
                        Application.Max(dDeb, d12) + 1)
            End If
        End If
    Next i

    wsVue.Range("A1:E1").Value = Array("Employ", "1 mois", "3 mois", "6 mois", "12 mois")
    arr = dict.Keys
    For i = 0 To dict.Count - 1
        wsVue.Cells(i + 2, 1).Value = arr(i)
        wsVue.Cells(i + 2, 2).Resize(1, 4).Value = dict(arr(i))
    Next i
    wsVue.Columns("A:E").AutoFit
    MsgBox "Vue absences mise jour", vbInformation
End Sub
