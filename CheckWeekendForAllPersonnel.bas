Attribute VB_Name = "CheckWeekendForAllPersonnel"
Option Explicit

Sub CheckWeekendForAllEmployees()
    On Error GoTo Cleanup

    ' Désactiver les événements pour optimiser la performance
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Déclarations des variables
    Dim wsPersonnel As Worksheet
    Dim wsMonthly As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim employeeRow As Long
    Dim startCol As Long
    Dim endCol As Long
    Dim header As Variant
    Dim data As Variant
    Dim saturdayWorked As Boolean
    Dim sundayWorked As Boolean
    Dim employeesWithoutWeekend As String
    Dim employeeFullName As String
    Dim foundCell As Range
    Dim j As Long
    
    ' Assigner la feuille Personnel et la feuille de mois actuel (Septembre par exemple)
    Set wsPersonnel = ThisWorkbook.Sheets("Personnel")
    Set wsMonthly = ThisWorkbook.Sheets("Sept")

    ' Définir les colonnes de début et de fin de la plage de jours de la semaine dans la feuille mensuelle
    startCol = 2 ' B correspond à 2
    endCol = 32  ' AF correspond à 32

    ' Trouver la dernière ligne utilisée dans la feuille Personnel (dynamique)
    lastRow = wsPersonnel.Cells(wsPersonnel.rows.Count, "B").End(xlUp).row

    ' Initialiser une variable pour stocker les employés qui n'ont pas travaillé le week-end
    employeesWithoutWeekend = ""

    ' Charger les en-têtes de la feuille mensuelle dans un tableau
    header = wsMonthly.Range(wsMonthly.Cells(3, startCol), wsMonthly.Cells(3, endCol)).value

    ' Parcourir chaque employé dans la feuille Personnel (en commençant à la ligne 2)
    For row = 2 To lastRow
        ' Récupérer le nom complet de l'employé (nom + prénom)
        employeeFullName = wsPersonnel.Cells(row, 2).value & "_" & wsPersonnel.Cells(row, 3).value

        ' Rechercher la ligne correspondante dans la feuille de mois (par exemple Septembre)
        Set foundCell = wsMonthly.Columns(1).Find(What:=employeeFullName, LookIn:=xlValues, LookAt:=xlWhole)

        ' Si l'employé est trouvé dans la feuille de mois
        If Not foundCell Is Nothing Then
            employeeRow = foundCell.row

            ' Charger les données de l'employé dans un tableau
            data = wsMonthly.Range(wsMonthly.Cells(employeeRow, startCol), wsMonthly.Cells(employeeRow, endCol)).value

            ' Initialiser les indicateurs pour le travail du samedi et du dimanche
            saturdayWorked = False
            sundayWorked = False

            ' Parcourir les en-têtes pour identifier "sam" et "dim" et vérifier le travail
            For j = 1 To UBound(header, 2)
                ' Vérification de la colonne samedi ("sam")
                If LCase(Trim(header(1, j))) = "sam" Then
                    If Not IsEmpty(data(1, j)) And IsNumeric(data(1, j)) And data(1, j) > 0 Then
                        saturdayWorked = True
                    End If
                ' Vérification de la colonne dimanche ("dim")
                ElseIf LCase(Trim(header(1, j))) = "dim" Then
                    If Not IsEmpty(data(1, j)) And IsNumeric(data(1, j)) And data(1, j) > 0 Then
                        sundayWorked = True
                    End If
                End If
            Next j

            ' Si l'employé n'a pas travaillé à la fois samedi et dimanche
            If Not (saturdayWorked And sundayWorked) Then
                employeesWithoutWeekend = employeesWithoutWeekend & employeeFullName & vbNewLine
            End If
        End If
    Next row

    ' Afficher les employés qui n'ont pas travaillé le week-end complet
    If employeesWithoutWeekend <> "" Then
        MsgBox "Les employés suivants n'ont pas presté un week-end complet :" & vbNewLine & employeesWithoutWeekend, vbExclamation, "Vérification de week-end"
    Else
        MsgBox "Tous les employés ont travaillé un week-end complet.", vbInformation, "Vérification de week-end"
    End If

Cleanup:
    ' Réactiver les événements
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' Gérer les erreurs
    If Err.Number <> 0 Then
        MsgBox "Une erreur est survenue: " & Err.Description, vbCritical, "Erreur"
    End If
End Sub

