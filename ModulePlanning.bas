Attribute VB_Name = "ModulePlanning"
' --- Module: ModulePlanning ---

' Fonction pour obtenir l'index du mois à partir du nom de l'onglet
Function GetMonthIndex(sheetName As String) As Long
    Dim monthNames As Variant
    monthNames = Array("Janv", "JanvB", "Fev", "Mars", "Avril", "Mai", "Juin", "Juil", "Aout", "Août", "Sept", "Oct", "Nov", "Dec")
    Dim idx As Variant

    idx = Application.Match(sheetName, monthNames, 0)
    If Not IsError(idx) Then
        GetMonthIndex = idx
    ElseIf IsNumeric(sheetName) Then
        If CLng(sheetName) >= 1 And CLng(sheetName) <= 12 Then
            GetMonthIndex = CLng(sheetName)
        Else
            GetMonthIndex = 0
        End If
    Else
        GetMonthIndex = 0
    End If
End Function

' Fonction pour obtenir le numéro de colonne à partir de l'en-tête
Function GetColumnIndex(headerRow As Range, headerName As String) As Long
    Dim idx As Variant
    idx = Application.Match(headerName, headerRow, 0)
    If Not IsError(idx) Then
        GetColumnIndex = idx
    Else
        GetColumnIndex = 0
    End If
End Function

' Fonction pour trouver le nom complet dans la feuille Personnel
Function FindFullName(personnelSheet As Worksheet, fullNameToFind As String) As Range
    Dim combinedNames As Variant
    Dim fullNames() As String
    Dim lastRow As Long, i As Long

    lastRow = personnelSheet.Cells(personnelSheet.Rows.Count, "B").End(xlUp).row
    combinedNames = personnelSheet.Range("B2:C" & lastRow).value

    ReDim fullNames(1 To UBound(combinedNames, 1))
    For i = 1 To UBound(combinedNames, 1)
        fullNames(i) = Trim(combinedNames(i, 1) & " " & combinedNames(i, 2))
    Next i

    For i = 1 To UBound(fullNames)
        If StrComp(fullNames(i), Trim(fullNameToFind), vbTextCompare) = 0 Then
            Set FindFullName = personnelSheet.Cells(i + 1, "B")
            Exit Function
        End If
    Next i
End Function

' Sub principale pour la mise à jour de la ligne
Public Sub MajLigne()
    Dim ws As Worksheet, personnelSheet As Worksheet
    Dim selectedRow As Long, selectedName As String
    Dim currentSheetName As String, monthIndex As Long
    Dim personnelMonth As String, monthCol As Long
    Dim foundCell As Range, positionValue As Variant
    Dim lastUsedRow As Long

    ' Nom de l'onglet actuel
    currentSheetName = ActiveSheet.Name

    ' Obtenir l'index du mois
    monthIndex = GetMonthIndex(currentSheetName)

    ' Vérifier si l'onglet est valide pour la mise à jour
    If monthIndex = 0 Then
        MsgBox "Cet onglet (" & currentSheetName & ") n'est pas configuré pour la mise à jour.", vbExclamation
        Exit Sub
    End If

    ' Définir la liste des mois pour la correspondance
    Dim personnelMonthArray As Variant
    personnelMonthArray = Array("Janv", "JanvB", "Fev", "Mars", "Avril", "Mai", "Juin", "Juil", "Aout", "Aoû", "Sept", "Oct", "Nov", "Dec")

    ' Vérifier que l'index est valide dans le tableau
    If monthIndex < 1 Or monthIndex > UBound(personnelMonthArray) + 1 Then
        MsgBox "Index de mois invalide pour l'onglet sélectionné.", vbExclamation
        Exit Sub
    End If

    personnelMonth = personnelMonthArray(monthIndex - 1)
    Set personnelSheet = ThisWorkbook.Sheets("Personnel")

    ' Ouvrir le formulaire pour sélectionner un nom
    With UserForm3
        .Show
        selectedName = .Tag
    End With
    Unload UserForm3

    ' Vérifier qu'un nom a été sélectionné
    If Trim(selectedName) = "" Then
        MsgBox "Aucun nom sélectionné.", vbExclamation
        Exit Sub
    End If

    ' Obtenir le numéro de colonne pour le mois
    monthCol = GetColumnIndex(personnelSheet.Rows(1), personnelMonth & " Position")
    If monthCol = 0 Then
        MsgBox "La colonne pour '" & personnelMonth & " Position' n'a pas été trouvée.", vbExclamation
        Exit Sub
    End If

    ' Trouver le nom complet dans la feuille Personnel
    Set foundCell = FindFullName(personnelSheet, selectedName)
    If foundCell Is Nothing Then
        MsgBox "Nom non trouvé dans la feuille Personnel.", vbExclamation
        Exit Sub
    End If

    ' Obtenir la valeur de la position
    positionValue = personnelSheet.Cells(foundCell.row, monthCol).value
    If Not IsNumeric(positionValue) Then
        MsgBox "La position pour " & personnelMonth & " à la ligne " & foundCell.row & " n'est pas numérique.", vbExclamation
        Exit Sub
    End If
    selectedRow = CLng(positionValue)

    ' Vérifier que la ligne sélectionnée peut être utilisée
    If selectedRow = 5 Or (selectedRow >= 109 And selectedRow <= 112) Then
        MsgBox "La ligne sélectionnée (" & selectedRow & ") ne peut pas être utilisée car elle doit rester masquée.", vbExclamation
        Exit Sub
    End If

    ' Manipuler les lignes visibles
    Set ws = ActiveSheet
    Application.ScreenUpdating = False
    ws.Rows.Hidden = False ' Afficher toutes les lignes

    ' Masquer les lignes spécifiques
    ws.Rows(5).Hidden = True
    ws.Rows("109:112").Hidden = True

    ' Trouver la dernière ligne utilisée dans la feuille active
    lastUsedRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' Masquer les lignes non sélectionnées
    If selectedRow > 6 Then ws.Range("6:" & selectedRow - 1).EntireRow.Hidden = True
    If selectedRow < lastUsedRow Then ws.Range(selectedRow + 1 & ":" & lastUsedRow).EntireRow.Hidden = True

    ' Toujours afficher les lignes d'en-tête
    ws.Rows("1:4").Hidden = False

    Application.ScreenUpdating = True
    MsgBox "Mise à jour terminée.", vbInformation
End Sub


' Additional helper functions such as GetMonthIndex, GetColumnIndex, and FindFullName would need to be correctly defined for this code to work effectively.


' --- AfficherMasquerLignes1 Macro ---
Sub AfficherMasquerLignes1()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    With ws.Rows("43:44")
        .Hidden = Not .Hidden
    End With
End Sub

' --- AfficherMasquerLignes_Fractions Macro ---
Sub AfficherMasquerLignes_Fractions()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    With ws.Rows("102:108")
        .Hidden = Not .Hidden
    End With
End Sub

' --- AfficherMasquerLignes Macro ---
Sub AfficherMasquerLignes()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    With ws.Rows("43:50")
        .Hidden = Not .Hidden
    End With
End Sub

' --- RemplirDatesSub ---
Sub RemplirDatesSub(ws As Worksheet)
    Dim monthNumber As Integer, yearNumber As Integer
    Dim startDate As Date, endDate As Date, currentDate As Date
    Dim i As Long, holidays As Collection, holidayDate As Variant

    monthNumber = val(ws.Range("F3").value)
    yearNumber = Year(Date)

    If monthNumber >= 1 And monthNumber <= 12 Then
        startDate = DateSerial(yearNumber, monthNumber, 1)
        endDate = DateSerial(yearNumber, monthNumber + 1, 0)

        Set holidays = New Collection
        Dim cell As Range
        For Each cell In ws.Range("H2:H100")
            If IsDate(cell.value) Then holidays.Add cell.value
        Next cell

        Application.ScreenUpdating = False
        i = 7
        For currentDate = startDate To endDate
            ws.Cells(i, 2).value = currentDate
            ws.Cells(i, 2).NumberFormat = "dd-mm"
            ws.Cells(i, 2).HorizontalAlignment = xlCenter

            Select Case Weekday(currentDate, vbMonday)
                Case 1: ws.Cells(i, 1).value = "Lu"
                Case 2: ws.Cells(i, 1).value = "Ma"
                Case 3: ws.Cells(i, 1).value = "Me"
                Case 4: ws.Cells(i, 1).value = "Je"
                Case 5: ws.Cells(i, 1).value = "Ve"
                Case 6: ws.Cells(i, 1).value = "Sa"
                Case 7: ws.Cells(i, 1).value = "Di"
            End Select

            ws.Cells(i, 1).HorizontalAlignment = xlCenter
            ws.Cells(i, 1).Font.Bold = True

            Dim IsWeekend As Boolean, IsHoliday As Boolean
            IsWeekend = (Weekday(currentDate, vbMonday) >= 6)
            IsHoliday = False

            For Each holidayDate In holidays
                If holidayDate = currentDate Then
                    IsHoliday = True
                    Exit For
                End If
            Next holidayDate

            If IsWeekend Or IsHoliday Then
                ws.Range(ws.Cells(i, 1), ws.Cells(i, 2)).Interior.Color = RGB(200, 200, 200)
                ws.Range(ws.Cells(i, 1), ws.Cells(i, 2)).Font.Color = RGB(255, 0, 0)
            End If

            i = i + 1
        Next currentDate
        Application.ScreenUpdating = True
    Else
        MsgBox "La valeur en F3 n'est pas un numéro de mois valide."
    End If
End Sub

' --- PDF_JOUR Macro ---
Sub PDF_JOUR()
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:= _
        "Chemin\Vers\Votre\Fichier.pdf", Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
End Sub

' --- SendMailWithAttachment ---
Sub SendMailWithAttachment()
    Dim objOutlook As Object, objMail As Object
    Dim strPath As String, strFile As String
    Dim fs As Object, folder As Object, file As Object
    Dim dtmLast As Date, response As Integer

    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0)

    strPath = "C:\Chemin\Vers\Votre\Dossier\"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set folder = fs.GetFolder(strPath)
    dtmLast = DateSerial(1900, 1, 1)

    For Each file In folder.Files
        If LCase(fs.GetExtensionName(file.Name)) = "pdf" Then
            If file.DateLastModified > dtmLast Then
                dtmLast = file.DateLastModified
                strFile = file.Path
            End If
        End If
    Next file

    With objMail
        .To = "destinataire@example.com"
        response = MsgBox("Mettre un destinataire en CC?", vbYesNo, "Choix CC")
        If response = vbYes Then
            .CC = "cc@example.com"
            .BCC = "bcc1@example.com; bcc2@example.com"
        Else
            .CC = "autrecc@example.com"
            .BCC = "bcc1@example.com; bcc2@example.com"
        End If

        .Subject = "Demande de Remplacement"
        .Body = "Bonjour," & vbCrLf & vbCrLf & _
                "Veuillez trouver ci-joint le document pour les remplacements nécessaires." & vbCrLf & vbCrLf & _
                "Cordialement," & vbCrLf & "Votre Nom"

        If strFile <> "" Then .Attachments.Add strFile
        .Display
    End With

    Set objMail = Nothing
    Set objOutlook = Nothing
    Set file = Nothing
    Set folder = Nothing
    Set fs = Nothing
End Sub

' --- ToggleColorierCelluleVertFonce ---
Sub ToggleColorierCelluleVertFonce()
    If Not Application.ActiveCell Is Nothing Then
        With Application.ActiveCell
            If .Interior.Color = RGB(0, 100, 0) Then
                .Interior.pattern = xlNone
                .Font.Color = RGB(0, 0, 0)
            Else
                .Interior.Color = RGB(0, 100, 0)
                .Font.Color = RGB(255, 255, 255)
            End If
        End With
    End If
End Sub

' --- ToggleColorierCelluleBleuClair ---
Sub ToggleColorierCelluleBleuClair()
    If Not Application.ActiveCell Is Nothing Then
        With Application.ActiveCell.Interior
            If .Color = RGB(173, 216, 230) Then
                .pattern = xlNone
            Else
                .Color = RGB(173, 216, 230)
            End If
        End With
    End If
End Sub

' --- ToggleAsterisqueCellule ---
Sub ToggleAsterisqueCellule()
    If Not Application.ActiveCell Is Nothing Then
        With Application.ActiveCell
            If Right(.value, 1) = "*" Then
                .value = Left(.value, Len(.value) - 1)
            Else
                .value = .value & "*"
            End If
        End With
    End If
End Sub

Sub InsererCodeCTR()
    On Error GoTo Cleanup

    ' Désactiver les événements pour optimiser la performance
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim response As VbMsgBoxResult
    Dim employeeName As String
    Dim a4Value As String
    Dim moisNom As String
    Dim annee As Integer
    Dim moisNumero As Integer
    Dim moisPrecedent As Integer
    Dim parts() As String
    Dim moisDict As Object
    Dim anneePrecedente As Integer
    Dim currentYear As Integer

    ' Initialiser le Dictionary pour les mois
    Set moisDict = CreateObject("Scripting.Dictionary")
    moisDict.CompareMode = vbTextCompare ' Insensible à la casse
    moisDict.Add "janv", 1
    moisDict.Add "février", 2
    moisDict.Add "mars", 3
    moisDict.Add "avril", 4
    moisDict.Add "mai", 5
    moisDict.Add "juin", 6
    moisDict.Add "juillet", 7
    moisDict.Add "août", 8
    moisDict.Add "septembre", 9
    moisDict.Add "octobre", 10
    moisDict.Add "novembre", 11
    moisDict.Add "décembre", 12

    ' Lire et traiter la valeur de la cellule A4
    a4Value = Trim(Range("A4").value)

    ' Ajouter un espace entre le texte du mois et l'année
    Dim i As Integer
    For i = 1 To Len(a4Value)
        If IsNumeric(Mid(a4Value, i, 1)) Then
            a4Value = Trim(Left(a4Value, i - 1)) & " " & Mid(a4Value, i)
            Exit For
        End If
    Next i

    parts = Split(a4Value, " ")

    ' Vérifier que la cellule A4 contient bien un mois et une année
    If UBound(parts) <> 1 Then
        MsgBox "La cellule A4 ne contient pas une date valide (ex: 'mars 2024').", vbExclamation, "Erreur de Date"
        GoTo Cleanup
    End If

    moisNom = Trim(LCase(parts(0))) ' Nom du mois
    annee = val(parts(1)) ' Année extraite

    If annee = 0 Then
        MsgBox "L'année extraite de la cellule A4 n'est pas valide.", vbExclamation, "Erreur de Date"
        GoTo Cleanup
    End If

    ' Obtenir le numéro du mois en utilisant le Dictionary
    If moisDict.Exists(moisNom) Then
        moisNumero = moisDict(moisNom)
    Else
        MsgBox "Le mois extrait de la cellule A4 n'est pas valide.", vbExclamation, "Erreur de Date"
        GoTo Cleanup
    End If

    ' Calculer le mois précédent
    moisPrecedent = IIf(moisNumero = 1, 12, moisNumero - 1)

    ' Si c'est Janvier, vérifier les conditions sur Décembre de l'année précédente
    If moisNumero = 1 Then
        anneePrecedente = annee - 1

        ' Demander confirmation si le week-end de décembre de l'année précédente a bien été vérifié
        response = MsgBox("Avez-vous vérifié que l'employé a bien presté un week-end en Décembre " & anneePrecedente & " ?", vbYesNo + vbQuestion, "Vérification de Décembre")

        If response = vbYes Then
            ' Insérer le code CTR12
            If Not Application.ActiveCell Is Nothing Then
                With Application.ActiveCell
                    .value = "CTR12"
                    .Interior.Color = RGB(255, 165, 0)
                End With
            Else
                MsgBox "Aucune cellule active sélectionnée.", vbExclamation, "Erreur"
            End If
        Else
            MsgBox "Vous devez vérifier le week-end presté en Décembre " & anneePrecedente & " avant d'insérer le CTR.", vbExclamation, "Vérification non effectuée"
            GoTo Cleanup
        End If

    ' Pour les autres mois (Février à Décembre)
    Else
        If Not Application.ActiveCell Is Nothing Then
            With Application.ActiveCell
                .value = "CTR" & moisPrecedent
                .Interior.Color = RGB(255, 165, 0)
            End With
        Else
            MsgBox "Aucune cellule active sélectionnée.", vbExclamation, "Erreur"
        End If
    End If

Cleanup:
    ' Réactiver les paramètres d'application
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' Gérer les erreurs
    If Err.Number <> 0 Then
        MsgBox "Une erreur est survenue: " & Err.Description, vbCritical, "Erreur"
    End If
End Sub


Sub Afficher_cacher_menu()
    With UserForm1
        ' Définir la taille fixe du UserForm ici
        .height = 509.25 ' La hauteur en points
        .width = 201.75  ' La largeur en points
        
        ' Positionnement manuel du UserForm
        .StartUpPosition = 0 ' Manuel
        
        ' Calculer la position pour que le UserForm s'ouvre en haut à droite
        ' Ajuster pour un affichage à 70% de la largeur de l'écran
        Dim screenWidth As Double
        Dim screenHeight As Double
        screenWidth = Application.width * 0.7
        screenHeight = Application.height

        .Left = Application.Left + screenWidth - .width
        .Top = Application.Top

        ' Afficher ou masquer le UserForm
        If .Visible = False Then
            .Show
        Else
            .Hide
        End If
    End With
End Sub
' === (OPTIONNEL) À placer dans un module standard pour initialiser l'affichage une seule fois ===
Public Sub InitOngletRoulement()
    With Sheets("Roulements")
        .Columns("B").ColumnWidth = 20.95
        .Columns("C:BG").ColumnWidth = 4.95
        .Parent.Windows(1).Zoom = 50
    End With
End Sub
