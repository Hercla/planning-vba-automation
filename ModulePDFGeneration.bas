Attribute VB_Name = "ModulePDFGeneration"
Option Explicit

' --- CONSTANTES DE CHEMINS RELATIFS ET NOMS DE DOSSIERS ---
' Chemin RELATIF depuis la racine OneDrive vers le dossier parent des PDF
Private Const RELATIVE_PDF_PARENT_PATH As String = "IC.1 Admin & Team\1. Hot Topics\"
' Noms des dossiers spécifiques pour les PDF (doivent correspondre aux noms DANS OneDrive)
Private Const PDF_FOLDER_JOUR As String = "HORAIRE PDF TEAM JOUR"
Private Const PDF_FOLDER_NUIT As String = "HORAIRE PDF TEAM NUIT"
' Noms des SOUS-DOSSIERS d'archive DANS les dossiers PDF_FOLDER_JOUR/NUIT
Private Const ARCHIVE_SUBFOLDER_JOUR As String = "Archive_Jour" ' Nom du sous-dossier d'archive pour l'équipe JOUR
Private Const ARCHIVE_SUBFOLDER_NUIT As String = "Archive_Nuit" ' Nom du sous-dossier d'archive pour l'équipe NUIT

' --- CONSTANTE POUR WHATSAPP ---
                    If yearStr = "" Then y = Year(tempDate)
                Else
                    m = 0 ' Month not recognized
                End If
            End If
            On Error GoTo 0
    End Select

    If m = 0 Then
        Exit Function ' Return CDate(0) as month was not recognized
    End If

    If yearStr <> "" Then
        y = CInt(yearStr)
    ElseIf y <> 0 Then ' y might have been set if monthStr was like "Avril 2024" but not split correctly.
        ' This condition is mostly for the case where y was extracted from DateValue(monthStr)
    ElseIf yearVal <> 0 Then
        y = yearVal
    Else
        y = Year(Date) ' Default to current year
    End If

    GetMonthDateFromName = DateSerial(y, m, 1)
End Function

' --- Archive les PDF du mois précédent dans leurs sous-dossiers d'archive respectifs ---
Sub ArchivePreviousMonthPDFs()
    Dim datePreviousMonth As Date
    Dim previousMonthName As String
    Dim baseOneDrivePath As String
    Dim fullParentPath As String
    Dim teamFolderPathJour As String, teamFolderPathNuit As String
    Dim archiveFolderPathJour As String, archiveFolderPathNuit As String
    Dim sourceFileJour As String, sourceFileNuit As String
    Dim destinationFileJour As String, destinationFileNuit As String

    On Error GoTo ArchiveErrorHandler

    ' 1. Calculer le nom du mois précédent
    datePreviousMonth = DateAdd("m", -1, Date)
    previousMonthName = Format(datePreviousMonth, "mmmm") ' vbUseSystemDayOfWeek is not needed for "mmmm"
    previousMonthName = UCase(Left(previousMonthName, 1)) & Mid(previousMonthName, 2)
    Debug.Print "Archivage: Recherche des fichiers du mois précédent à archiver : " & previousMonthName

    ' 2. Construire les chemins
    baseOneDrivePath = FindUserOneDriveBasePath()
    If baseOneDrivePath = "" Then
        Debug.Print "Archivage ignoré: Chemin OneDrive de base non trouvé."
        Exit Sub
    End If

    fullParentPath = baseOneDrivePath & RELATIVE_PDF_PARENT_PATH
    If Right(fullParentPath, 1) <> "\" Then fullParentPath = fullParentPath & "\"

    teamFolderPathJour = fullParentPath & PDF_FOLDER_JOUR & "\"
    teamFolderPathNuit = fullParentPath & PDF_FOLDER_NUIT & "\"

    archiveFolderPathJour = teamFolderPathJour & ARCHIVE_SUBFOLDER_JOUR & "\"
    archiveFolderPathNuit = teamFolderPathNuit & ARCHIVE_SUBFOLDER_NUIT & "\"

    sourceFileJour = teamFolderPathJour & "Horaire_" & previousMonthName & "_Jour.pdf"
    sourceFileNuit = teamFolderPathNuit & "Horaire_" & previousMonthName & "_Nuit.pdf"

    destinationFileJour = archiveFolderPathJour & "Horaire_" & previousMonthName & "_Jour.pdf"
    destinationFileNuit = archiveFolderPathNuit & "Horaire_" & previousMonthName & "_Nuit.pdf"

    ' 3. Créer les dossiers d'archive s'ils n'existent pas
    On Error Resume Next ' Ignorer l'erreur si le dossier existe déjà
    If Dir(archiveFolderPathJour, vbDirectory) = "" Then
        MkDir archiveFolderPathJour
        Debug.Print "Archivage: Dossier '" & archiveFolderPathJour & "' créé."
    End If
    If Dir(archiveFolderPathNuit, vbDirectory) = "" Then
        MkDir archiveFolderPathNuit
        Debug.Print "Archivage: Dossier '" & archiveFolderPathNuit & "' créé."
    End If
    On Error GoTo ArchiveErrorHandler ' Rétablir la gestion d'erreur normale

    ' 4. Déplacer les fichiers s'ils existent (Source -> Archive)
    On Error Resume Next ' Ignorer l'erreur si le fichier source n'existe pas

    If Dir(sourceFileJour) <> "" Then
        Name sourceFileJour As destinationFileJour ' Déplace le fichier
        Debug.Print "Archivage : Déplacé " & sourceFileJour & " vers " & destinationFileJour
    Else
        Debug.Print "Archivage : Fichier Jour du mois précédent non trouvé dans le dossier principal : " & sourceFileJour
    End If

    If Dir(sourceFileNuit) <> "" Then
        Name sourceFileNuit As destinationFileNuit ' Déplace le fichier
        Debug.Print "Archivage : Déplacé " & sourceFileNuit & " vers " & destinationFileNuit
    Else
        Debug.Print "Archivage : Fichier Nuit du mois précédent non trouvé dans le dossier principal : " & sourceFileNuit
    End If

    On Error GoTo 0 ' Rétablir la gestion d'erreur par défaut
    Exit Sub

ArchiveErrorHandler:
    MsgBox "Une erreur est survenue lors de l'archivage des PDF du mois précédent : " & Err.Description, vbExclamation
    On Error GoTo 0
End Sub


' --- Nettoie les PDF DANS LES ARCHIVES datant de trois mois (ex: au 1er Juillet, supprime ceux d'Avril des archives) ---
Sub CleanupArchivedPDFs()
    Dim dateForCleanupTarget As Date
    Dim monthToCleanupName As String
    Dim baseOneDrivePath As String
    Dim folderPathJour As String, folderPathNuit As String
    Dim archiveFolderPathJour As String, archiveFolderPathNuit As String
    Dim fileToDeleteJour As String, fileToDeleteNuit As String
    Dim fullParentPath As String

    On Error GoTo CleanupErrorHandler

    ' 1. Calculer le nom du mois à nettoyer (Mois actuel - 3 mois)
    dateForCleanupTarget = DateAdd("m", -3, Date)
    monthToCleanupName = Format(dateForCleanupTarget, "mmmm") ' vbUseSystemDayOfWeek is not needed for "mmmm"
    monthToCleanupName = UCase(Left(monthToCleanupName, 1)) & Mid(monthToCleanupName, 2)
    Debug.Print "Nettoyage des archives : Recherche des fichiers (datant de 3 mois) : " & monthToCleanupName

    ' 2. Construire les chemins complets des fichiers potentiels à supprimer DANS LES ARCHIVES
    baseOneDrivePath = FindUserOneDriveBasePath()
    If baseOneDrivePath = "" Then
        Debug.Print "Nettoyage des archives ignoré: Chemin OneDrive de base non trouvé."
        Exit Sub
    End If

    fullParentPath = baseOneDrivePath & RELATIVE_PDF_PARENT_PATH
    If Right(fullParentPath, 1) <> "\" Then fullParentPath = fullParentPath & "\"

    folderPathJour = fullParentPath & PDF_FOLDER_JOUR & "\"
    folderPathNuit = fullParentPath & PDF_FOLDER_NUIT & "\"

    archiveFolderPathJour = folderPathJour & ARCHIVE_SUBFOLDER_JOUR & "\"
    archiveFolderPathNuit = folderPathNuit & ARCHIVE_SUBFOLDER_NUIT & "\"

    fileToDeleteJour = archiveFolderPathJour & "Horaire_" & monthToCleanupName & "_Jour.pdf"
    fileToDeleteNuit = archiveFolderPathNuit & "Horaire_" & monthToCleanupName & "_Nuit.pdf"

    ' 3. Supprimer les fichiers s'ils existent DANS LES ARCHIVES (ignorer erreurs si dossier archive n'existe pas encore)
    On Error Resume Next

    If Dir(archiveFolderPathJour, vbDirectory) <> "" Then ' Vérifie si le dossier archive existe
        If Dir(fileToDeleteJour) <> "" Then
            Kill fileToDeleteJour
            Debug.Print "Nettoyage des archives : Supprimé " & fileToDeleteJour
        Else
            Debug.Print "Nettoyage des archives : Fichier Jour (datant de 3 mois) non trouvé dans l'archive : " & fileToDeleteJour
        End If
    Else
        Debug.Print "Nettoyage des archives : Dossier archive Jour non trouvé: " & archiveFolderPathJour
    End If


    If Dir(archiveFolderPathNuit, vbDirectory) <> "" Then ' Vérifie si le dossier archive existe
        If Dir(fileToDeleteNuit) <> "" Then
            Kill fileToDeleteNuit
            Debug.Print "Nettoyage des archives : Supprimé " & fileToDeleteNuit
        Else
            Debug.Print "Nettoyage des archives : Fichier Nuit (datant de 3 mois) non trouvé dans l'archive : " & fileToDeleteNuit
        End If
    Else
        Debug.Print "Nettoyage des archives : Dossier archive Nuit non trouvé: " & archiveFolderPathNuit
    End If

    On Error GoTo 0
    Exit Sub

CleanupErrorHandler:
    MsgBox "Une erreur est survenue lors du nettoyage des anciens PDF archivés : " & Err.Description, vbExclamation
    On Error GoTo 0
End Sub


' --- Subroutine Principale d'Export PDF (MODIFIED) ---
Sub ExportHorairePDF(ws As Worksheet, equipe As String)
    Dim pdfTeamFolderPath As String ' Dossier principal de l'équipe (Jour/Nuit)
    Dim targetPdfFolderPath As String ' Dossier de destination final (principal ou archive)
    Dim archiveSubFolder As String
    Dim pdfFileName As String
    Dim userOneDrivePath As String
    Dim fullParentPath As String
    Dim monthName As String
    Dim printRangeAddress As String
    Dim success As Boolean
    Dim isPastMonth As Boolean
    Dim sheetMonthDate As Date
    Dim currentMonthStartDate As Date

    On Error GoTo ErrorHandler_Export

    ' 1. Valider l'équipe
    If UCase(equipe) <> "JOUR" And UCase(equipe) <> "NUIT" Then
        MsgBox "Type d'équipe non valide spécifié ('" & equipe & "').", vbCritical, "Erreur d'Export"
        Exit Sub
    End If

    ' 2. Obtenir le chemin de base OneDrive utilisateur
    userOneDrivePath = FindUserOneDriveBasePath()
    If userOneDrivePath = "" Then
        MsgBox "Impossible de trouver un dossier OneDrive valide pour les utilisateurs configurés." & vbCrLf & _
               "Veuillez vérifier la configuration dans la fonction 'FindUserOneDriveBasePath'.", vbCritical, "Erreur de Chemin Utilisateur"
        Exit Sub
    End If

    ' 3. Construire le chemin du dossier PARENT des PDF
    fullParentPath = userOneDrivePath & RELATIVE_PDF_PARENT_PATH
    If Right(fullParentPath, 1) <> "\" Then fullParentPath = fullParentPath & "\"

    ' 4. Obtenir le nom du mois depuis la feuille et déterminer si c'est un mois passé
    monthName = ws.Name
    If Trim(monthName) = "" Then
        MsgBox "Le nom de la feuille (utilisé comme nom de mois) est vide. Veuillez nommer la feuille correctement.", vbCritical, "Erreur Nom de Mois"
        Exit Sub
    End If

    ' --- MODIFICATION START: Utiliser GetMonthDateFromName pour interpréter le nom de la feuille ---
    sheetMonthDate = GetMonthDateFromName(monthName)

    If sheetMonthDate = CDate(0) Then ' CDate(0) est retourné par GetMonthDateFromName si le nom n'est pas valide
        MsgBox "Le nom de la feuille '" & monthName & "' ne peut pas être interprété comme un mois/date valide." & vbCrLf & _
               "Veuillez utiliser un nom de mois (ex: 'Avril') ou un format reconnaissable (ex: 'Avril 2024').", vbCritical, "Erreur Nom de Mois"
        Exit Sub
    End If
    ' --- MODIFICATION END ---

    currentMonthStartDate = DateSerial(Year(Date), Month(Date), 1) ' Premier jour du mois en cours

    ' Un mois est "passé" s'il est strictement antérieur au mois en cours.
    ' Si la feuille est pour le mois en cours, ce n'est PAS "passé" pour cette logique.
    isPastMonth = (sheetMonthDate < currentMonthStartDate)

    Debug.Print "Export PDF: Feuille '" & monthName & "', Date feuille interprétée: " & Format(sheetMonthDate, "dd-mmm-yyyy") & ", Début mois actuel: " & Format(currentMonthStartDate, "dd-mmm-yyyy") & ", Est un mois passé: " & isPastMonth

    ' 5. Construire le chemin du dossier PDF SPÉCIFIQUE de l'équipe (Jour ou Nuit)
    If UCase(equipe) = "JOUR" Then
        pdfTeamFolderPath = fullParentPath & PDF_FOLDER_JOUR & "\"
        archiveSubFolder = ARCHIVE_SUBFOLDER_JOUR
    Else ' NUIT
        pdfTeamFolderPath = fullParentPath & PDF_FOLDER_NUIT & "\"
        archiveSubFolder = ARCHIVE_SUBFOLDER_NUIT
    End If

    ' Déterminer le dossier de destination final
    If isPastMonth Then
        targetPdfFolderPath = pdfTeamFolderPath & archiveSubFolder & "\"
        Debug.Print "Export PDF: Cible pour mois passé: " & targetPdfFolderPath
    Else
        targetPdfFolderPath = pdfTeamFolderPath
        Debug.Print "Export PDF: Cible pour mois actuel/futur: " & targetPdfFolderPath
    End If

    ' Vérifier si le dossier de destination (principal ou archive) existe, sinon le créer
    If Dir(targetPdfFolderPath, vbDirectory) = "" Then
        Debug.Print "Export PDF: Dossier cible '" & targetPdfFolderPath & "' non trouvé, tentative de création."
        On Error Resume Next
        MkDir targetPdfFolderPath
        If Err.Number <> 0 Then
             MsgBox "Impossible de créer ou d'accéder au dossier de destination PDF '" & targetPdfFolderPath & "'. Erreur: " & Err.Description, vbCritical, "Erreur de Dossier"
             Exit Sub
        End If
        On Error GoTo ErrorHandler_Export ' Rétablir la gestion d'erreur pour la suite
        Debug.Print "Export PDF: Dossier cible '" & targetPdfFolderPath & "' créé."
    End If


    ' 6. Construire le nom de fichier FIXE par mois
    ' Utiliser le nom du mois tel qu'interprété et formaté par le système pour cohérence
    ' (par exemple, si la feuille est "Aout", mais le système génère "Août", utiliser "Août")
    Dim formattedSheetMonthName As String
    formattedSheetMonthName = Format(sheetMonthDate, "mmmm")
    formattedSheetMonthName = UCase(Left(formattedSheetMonthName, 1)) & Mid(formattedSheetMonthName, 2)
    
    pdfFileName = "Horaire_" & formattedSheetMonthName & "_" & equipe & ".pdf"


    ' 7. Définir la zone d'impression
    printRangeAddress = "$A$1:$AF$104"

    ' 8. Exporter le PDF (Écrase le fichier existant dans le dossier cible)
    success = False
    With ws
        .PageSetup.PrintArea = printRangeAddress
        .PageSetup.PrintComments = xlPrintInPlace
        .PageSetup.Orientation = xlLandscape
        .PageSetup.Zoom = False
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = False

        On Error Resume Next ' Erreur spécifique pour l'exportation
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=targetPdfFolderPath & pdfFileName, _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        If Err.Number = 0 Then
            success = True
        Else
            MsgBox "Échec de l'exportation PDF pour '" & pdfFileName & "'." & vbCrLf & vbCrLf & _
                   "Vérifiez si le fichier n'est pas déjà ouvert ou si le chemin est correct." & vbCrLf & _
                   "Dossier cible : " & targetPdfFolderPath & vbCrLf & _
                   "Erreur: " & Err.Description, vbCritical, "Erreur Export PDF"
        End If
        On Error GoTo ErrorHandler_Export ' Rétablir la gestion d'erreur générale
    End With

    ' 9. Message de succès et Préparation WhatsApp
    If success Then
        MsgBox "Le fichier PDF '" & pdfFileName & "' a été généré/mis à jour avec succès dans :" & vbCrLf & targetPdfFolderPath, vbInformation, "Export PDF Terminé"

        Dim whatsappNumber As String
        Dim messageText As String
        Dim encodedMessage As String
        Dim whatsappLink As String

        whatsappNumber = MY_WHATSAPP_NUMBER

        If whatsappNumber <> "" Then
            ' Utiliser formattedSheetMonthName pour le message WhatsApp pour la cohérence
            messageText = "L'horaire PDF de " & formattedSheetMonthName & " (" & equipe & ") a été mis à jour."
            If isPastMonth Then messageText = messageText & " (Archivé)"

            On Error Resume Next
            encodedMessage = Application.EncodeURL(messageText)
            If Err.Number <> 0 Then
                encodedMessage = Replace(messageText, " ", "%20")
                encodedMessage = Replace(encodedMessage, "(", "%28")
                encodedMessage = Replace(encodedMessage, ")", "%29")
                Err.Clear
            End If
            On Error GoTo ErrorHandler_Export ' Rétablir

            whatsappLink = "https://wa.me/" & whatsappNumber & "?text=" & encodedMessage

            Dim openLinkConfirmation As VbMsgBoxResult
            openLinkConfirmation = MsgBox("Le PDF a été généré." & vbCrLf & vbCrLf & _
                                         "Voulez-vous ouvrir WhatsApp maintenant ?" & vbCrLf & _
                                         "(Un message sera pré-rempli dans une discussion avec vous-même. " & _
                                         "Vous pourrez ensuite le TRANSFÉRER au groupe HORAIRES)", _
                                         vbYesNo + vbQuestion, "Ouvrir WhatsApp ?")

            If openLinkConfirmation = vbYes Then
                On Error Resume Next
                ThisWorkbook.FollowHyperlink whatsappLink
                If Err.Number <> 0 Then
                    MsgBox "Impossible d'ouvrir le lien WhatsApp automatiquement." & vbCrLf & _
                           "Vous pouvez copier/coller ce lien dans votre navigateur :" & vbCrLf & vbCrLf & whatsappLink, _
                           vbExclamation, "Erreur Ouverture Lien"
                    Err.Clear
                End If
                On Error GoTo ErrorHandler_Export ' Rétablir
            End If
        Else
             Debug.Print "Numéro WhatsApp principal (MY_WHATSAPP_NUMBER) non configuré. Notification non préparée."
        End If
    End If

    Exit Sub

ErrorHandler_Export:
    MsgBox "Une erreur inattendue est survenue dans ExportHorairePDF : " & Err.Description, vbCritical, "Erreur Inattendue"
End Sub

' --- Procédures d'Appel Utilisateur ---
Sub Generate_PDF_Jour()
    ' 1. Archiver les PDF du mois précédent (Jour ET Nuit) si présents dans les dossiers principaux
    ArchivePreviousMonthPDFs

    ' 2. Nettoyer les PDF (datant de 3 mois) DANS LES ARCHIVES (Jour ET Nuit)
    CleanupArchivedPDFs

    ' 3. Exporter le PDF du mois en cours (ou passé) pour l'équipe Jour sur la feuille active
    If TypeName(ActiveSheet) = "Worksheet" Then
        ExportHorairePDF ActiveSheet, "Jour"
    Else
        MsgBox "Veuillez sélectionner une feuille de calcul valide avant de lancer la génération du PDF Jour.", vbExclamation
    End If
End Sub

Sub Generate_PDF_Nuit()
    ' 1. Archiver les PDF du mois précédent (Jour ET Nuit) si présents dans les dossiers principaux
    ArchivePreviousMonthPDFs

    ' 2. Nettoyer les PDF (datant de 3 mois) DANS LES ARCHIVES (Jour ET Nuit)
    CleanupArchivedPDFs

    ' 3. Exporter le PDF du mois en cours (ou passé) pour l'équipe Nuit sur la feuille active
    If TypeName(ActiveSheet) = "Worksheet" Then
        ExportHorairePDF ActiveSheet, "Nuit"
    Else
        MsgBox "Veuillez sélectionner une feuille de calcul valide avant de lancer la génération du PDF Nuit.", vbExclamation
    End If
End Sub

