Attribute VB_Name = "check_infirmiers_presence"
Option Explicit
'================================================================================
' MODULE: check_infirmiers_presence
' Purpose: Vérifie qu'il y a suffisamment d'infirmier·e·s programmés dans un
'          planning mensuel.
'================================================================================

' Entrée : demande à l’utilisateur l’équipe à vérifier (Jour/Nuit) puis lance
' l’analyse pour la feuille active.
Public Sub Check_Presence_Infirmiers()
    Dim team As String
    ' Demande à l'utilisateur de spécifier l'équipe à contrôler
    team = Trim(InputBox("Saisissez l'équipe à vérifier (Jour ou Nuit) :", _
        "Vérification des infirmier·e·s", "Jour"))

    If team = "" Then Exit Sub
    team = UCase(team)

    If team <> "JOUR" And team <> "NUIT" Then
        MsgBox "Valeur d'équipe non reconnue. Veuillez saisir 'Jour' ou 'Nuit'.", _
            vbCritical, "Equipe inconnue"
        Exit Sub
    End If

    ' Vérifie que la feuille active est bien une feuille de calcul
    If Not TypeOf ActiveSheet Is Worksheet Then
        MsgBox "Veuillez sélectionner une feuille de planning avant de lancer la vérification.", _
            vbExclamation, "Aucune feuille sélectionnée"
        Exit Sub
    End If

    ' Lance l'analyse sur la feuille active
    Call CheckPresenceForTeam(ActiveSheet, team)
End Sub

' Fonction privée : réalise l'ensemble de l'analyse pour une feuille et une
' équipe données. Affiche un message synthétique à l'issue.
Private Sub CheckPresenceForTeam(ByVal ws As Worksheet, ByVal teamName As String)
    Dim headerRow As Long
    Dim firstRow As Long
    Dim colStart As Long
    Dim lastCol As Long
    Dim lastRow As Long
    Dim dayCount As Long
    Dim i As Long, j As Long

    ' --- Déterminer la ligne d'en-têtes et la première ligne du personnel ---
    headerRow = 4
    firstRow = headerRow + 2 ' donc 6

    ' --- Déterminer la colonne de début (premier jour) et la dernière colonne ---
    colStart = 0
    lastCol = 0
    For j = 1 To ws.Columns.Count
        If IsNumeric(ws.Cells(headerRow, j).value) Then
            colStart = j
            Exit For
        End If
    Next j
    If colStart = 0 Then colStart = 3

    For j = colStart To ws.Columns.Count
        If IsNumeric(ws.Cells(headerRow, j).value) Then
            lastCol = j
        Else
            Exit For
        End If
    Next j
    If lastCol = 0 Then lastCol = colStart + 30
    dayCount = lastCol - colStart + 1

    ' --- Déterminer la dernière ligne contenant un nom ---
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow < firstRow Then
        MsgBox "Aucune donnée de personnel trouvée sur cette feuille.", vbExclamation
        Exit Sub
    End If

    ' --- Charger la table des codes horaires (Config_Codes) ---
    Dim dictCodes As Object
    Set dictCodes = CreateObject("Scripting.Dictionary")
    dictCodes.CompareMode = vbTextCompare
    Dim wsCodes As Worksheet
    On Error Resume Next
    Set wsCodes = ThisWorkbook.Sheets("Config_Codes")
    On Error GoTo 0
    If wsCodes Is Nothing Then
        MsgBox "La feuille 'Config_Codes' est introuvable.", vbCritical
        Exit Sub
    End If

    Dim colCode As Long, colMatin As Long, colApres As Long, colSoir As Long, colNuit As Long
    colCode = 0: colMatin = 0: colApres = 0: colSoir = 0: colNuit = 0
    For j = 1 To wsCodes.Columns.Count
        Dim head As String
        head = Trim(UCase(wsCodes.Cells(1, j).value))
        Select Case head
            Case "CODE": colCode = j
            Case "MATIN": colMatin = j
            Case "APRÈS-MIDI", "APRES-MIDI", "APRES_MIDI": colApres = j
            Case "SOIR": colSoir = j
            Case "NUIT": colNuit = j
        End Select
    Next j
    If colCode = 0 Or colMatin = 0 Or colApres = 0 Or colSoir = 0 Then
        MsgBox "Les colonnes 'Code', 'Matin', 'Après-midi' et 'Soir' sont requises dans Config_Codes.", vbCritical
        Exit Sub
    End If

    Dim lastRowCodes As Long
    lastRowCodes = wsCodes.Cells(wsCodes.Rows.Count, colCode).End(xlUp).row
    For i = 2 To lastRowCodes
        Dim codeKey As String
        codeKey = Trim(CStr(wsCodes.Cells(i, colCode).value))
        If codeKey <> "" Then
            Dim presence(1 To 4) As Boolean
            presence(1) = (val(wsCodes.Cells(i, colMatin).value) > 0)
            presence(2) = (val(wsCodes.Cells(i, colApres).value) > 0)
            presence(3) = (val(wsCodes.Cells(i, colSoir).value) > 0)
            If colNuit > 0 Then presence(4) = (val(wsCodes.Cells(i, colNuit).value) > 0) Else presence(4) = False
            dictCodes(codeKey) = presence
        End If
    Next i

    ' --- Charger les informations du personnel ---
    Dim wsPers As Worksheet
    Set wsPers = Nothing
    On Error Resume Next
    Set wsPers = ThisWorkbook.Sheets("Personnel")
    On Error GoTo 0
    If wsPers Is Nothing Then
        MsgBox "La feuille 'Personnel' est introuvable.", vbCritical
        Exit Sub
    End If

    Dim colNom As Long, colPrenom As Long, colFonction As Long, colEquipe As Long
    colNom = 0: colPrenom = 0: colFonction = 0: colEquipe = 0
    For j = 1 To wsPers.Columns.Count
        Dim h As String
        h = Trim(UCase(wsPers.Cells(1, j).value))
        Select Case h
            Case "NOM": colNom = j
            Case "PRÉNOM", "PRENOM": colPrenom = j
            Case "FONCTION": colFonction = j
            Case "ÉQUIPE", "EQUIPE": colEquipe = j
        End Select
    Next j
    If colNom = 0 Or colFonction = 0 Or colEquipe = 0 Then
        MsgBox "Les colonnes Nom, Fonction et Équipe sont requises dans la feuille Personnel.", vbCritical
        Exit Sub
    End If

    Dim lastRowPers As Long
    lastRowPers = wsPers.Cells(wsPers.Rows.Count, colNom).End(xlUp).row
    Dim staffCount As Long
    staffCount = lastRowPers - 1
    Dim staffFonction() As String, staffEquipe() As String, staffFullName() As String
    ReDim staffFonction(1 To staffCount)
    ReDim staffEquipe(1 To staffCount)
    ReDim staffFullName(1 To staffCount)
    For i = 1 To staffCount
        Dim idxRow As Long
        idxRow = i + 1
        staffFonction(i) = UCase(Trim(CStr(wsPers.Cells(idxRow, colFonction).value)))
        staffEquipe(i) = UCase(Trim(CStr(wsPers.Cells(idxRow, colEquipe).value)))
        Dim nom As String, prenom As String
        nom = Trim(CStr(wsPers.Cells(idxRow, colNom).value))
        prenom = ""
        If colPrenom > 0 Then prenom = Trim(CStr(wsPers.Cells(idxRow, colPrenom).value))
        If nom <> "" And prenom <> "" Then
            staffFullName(i) = Replace(nom, " ", "_") & "_" & Replace(prenom, " ", "_")
        Else
            staffFullName(i) = Replace(nom, " ", "_")
        End If
    Next i

    Dim dictNameIndex As Object
    Set dictNameIndex = CreateObject("Scripting.Dictionary")
    dictNameIndex.CompareMode = vbTextCompare
    For i = 1 To staffCount
        If staffFullName(i) <> "" Then
            If Not dictNameIndex.Exists(staffFullName(i)) Then dictNameIndex(staffFullName(i)) = i
        End If
    Next i

    ' --- Préparation des compteurs par jour ---
    Dim countsMorning() As Long, countsAfternoon() As Long, countsEvening() As Long, countsNight() As Long
    ReDim countsMorning(1 To dayCount)
    ReDim countsAfternoon(1 To dayCount)
    ReDim countsEvening(1 To dayCount)
    ReDim countsNight(1 To dayCount)
    Dim dayNumbers() As Variant
    ReDim dayNumbers(1 To dayCount)
    For j = 1 To dayCount
        dayNumbers(j) = ws.Cells(headerRow, colStart + j - 1).value
    Next j
    Const COULEUR_A_IGNORER As Long = 15849925

    ' --- Parcours des lignes du personnel ---
    For i = firstRow To lastRow
        Dim fullName As String
        fullName = Trim(CStr(ws.Cells(i, 1).value))
        
        Dim lookupKey As String
        lookupKey = Replace(fullName, ", ", "_")
        lookupKey = Replace(lookupKey, ",", "_")
        lookupKey = Replace(lookupKey, " ", "_")

        Dim persIndex As Long
        persIndex = 0
        If lookupKey <> "" Then
            If dictNameIndex.Exists(lookupKey) Then
                persIndex = dictNameIndex(lookupKey)
            End If
        End If

        If persIndex = 0 And fullName <> "" Then
            Dim tentative As Long
            tentative = i - firstRow + 1
            If tentative >= 1 And tentative <= staffCount Then
                persIndex = tentative
            End If
        End If

        If persIndex >= 1 And persIndex <= staffCount Then
            If staffEquipe(persIndex) = teamName Then
                
                ' On vérifie si la fonction contient "INF" OU si la fonction est exactement "IC"
                If InStr(1, staffFonction(persIndex), "INF", vbTextCompare) > 0 Or staffFonction(persIndex) = "IC" Then

                    For j = 1 To dayCount
                        Dim c As Long
                        c = colStart + j - 1
                        If ws.Cells(i, c).Interior.Color <> COULEUR_A_IGNORER Then
                            Dim codeVal As String
                            codeVal = Trim(CStr(ws.Cells(i, c).value))
                            If codeVal <> "" Then
                                Dim presenceArr As Variant
                                Dim known As Boolean
                                known = False
                                If dictCodes.Exists(codeVal) Then
                                    presenceArr = dictCodes(codeVal)
                                    known = True
                                End If
                                If Not known Then
                                    ReDim presenceArr(1 To 4)
                                    presenceArr(1) = True
                                    presenceArr(2) = True
                                    presenceArr(3) = True
                                    presenceArr(4) = True
                                End If
                                If teamName = "JOUR" Then
                                    If presenceArr(1) Then countsMorning(j) = countsMorning(j) + 1
                                    If presenceArr(2) Then countsAfternoon(j) = countsAfternoon(j) + 1
                                    If presenceArr(3) Then countsEvening(j) = countsEvening(j) + 1
                                Else
                                    If presenceArr(4) Or presenceArr(1) Or presenceArr(2) Or presenceArr(3) Then
                                        countsNight(j) = countsNight(j) + 1
                                    End If
                                End If
                            End If
                        End If
                    Next j
                End If
            End If
        End If
    Next i

    ' --- Détermination des jours en infraction ---
    Dim issuesMorning As String, issuesAfternoon As String, issuesEvening As String, issuesNight As String
    issuesMorning = "": issuesAfternoon = "": issuesEvening = "": issuesNight = ""
    For j = 1 To dayCount
        If teamName = "JOUR" Then
            If countsMorning(j) < 2 Then issuesMorning = issuesMorning & IIf(issuesMorning <> "", ", ", "") & dayNumbers(j)
            If countsAfternoon(j) < 2 Then issuesAfternoon = issuesAfternoon & IIf(issuesAfternoon <> "", ", ", "") & dayNumbers(j)
            If countsEvening(j) < 2 Then issuesEvening = issuesEvening & IIf(issuesEvening <> "", ", ", "") & dayNumbers(j)
        Else
            If countsNight(j) < 1 Then issuesNight = issuesNight & IIf(issuesNight <> "", ", ", "") & dayNumbers(j)
        End If
    Next j

    ' --- Préparer et afficher le message final ---
    Dim msg As String
    If teamName = "JOUR" Then
        If issuesMorning = "" And issuesAfternoon = "" And issuesEvening = "" Then
            msg = "Aucun problème : toutes les journées du planning contiennent au moins deux infirmier·e·s pour chaque fraction."
        Else
            msg = "Jours présentant un effectif insuffisant d'infirmier·e·s (minimum 2 requis) :" & vbCrLf & vbCrLf
            If issuesMorning <> "" Then msg = msg & "Matin : " & issuesMorning & vbCrLf
            If issuesAfternoon <> "" Then msg = msg & "Après-midi : " & issuesAfternoon & vbCrLf
            If issuesEvening <> "" Then msg = msg & "Soir : " & issuesEvening & vbCrLf
        End If
    Else
        If issuesNight = "" Then
            msg = "Aucun problème : toutes les journées du planning contiennent au moins un infirmier·e pour l'équipe Nuit."
        Else
            msg = "Jours sans infirmier·e pour l'équipe Nuit (minimum 1 requis) : " & issuesNight
        End If
    End If
    MsgBox msg, vbInformation, "Résultat de la vérification"
End Sub

