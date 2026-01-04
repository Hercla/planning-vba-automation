' ExportedAt: 2026-01-04 17:02:15 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "GenerateurCalendrier"
Option Explicit

'===================================================================================
' MODULE :      GenerateurCalendrier (Version corrigée colonne C)
' DESCRIPTION : Remplit les jours/semaine aux bonnes colonnes,
'               efface l'ancien en-tête, et masque les colonnes inutiles.
'               Aligné sur ton layout réel : jours du mois commencent en colonne C.
'===================================================================================

Public Sub GenererDatesEtJoursPourTousLesMois()
    Dim feuilleActuelle As Worksheet
    Dim dateJour As Date
    Dim annee As Long, indexMois As Integer
    Dim moisFrancais As Variant, nomsJoursFrancais As Variant
    Dim wd As Integer, totalJours As Integer
    Dim jourFeries As Collection
    Dim i As Long, col As Long
    
    ' --- constantes liées à ta mise en page ---
    Const FIRST_DAY_COL As Long = 3      ' Colonne C = 3
    Const LAST_DAY_COL As Long = 33      ' Colonne AG = 33
    Const ROW_JOUR_SEMAINE As Long = 3   ' Ligne avec "Lun Mar Mer ..."
    Const ROW_NUMERO_JOUR As Long = 4    ' Ligne avec "1 2 3 ..."
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    moisFrancais = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", _
                         "Juil", "Aout", "Sept", "Oct", "Nov", "Dec")
    nomsJoursFrancais = Array("Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim")
    
    ' Récupérer l'année dans Feuil_Config!B2
    On Error Resume Next
    annee = Sheets("Feuil_Config").Range("B2").value
    On Error GoTo 0
    
    If annee < 1900 Or annee > 2100 Then
        MsgBox "Année non valide dans 'Feuil_Config'!B2.", vbCritical
        GoTo Cleanup
    End If
    
    ' Construire la liste des jours fériés pour cette année
    Set jourFeries = New Collection
    On Error Resume Next
    jourFeries.Add DateSerial(annee, 1, 1), CStr(DateSerial(annee, 1, 1))         ' Nouvel An
    jourFeries.Add CalculerPaques(annee) + 1, CStr(CalculerPaques(annee) + 1)     ' Lundi de Pâques
    jourFeries.Add DateSerial(annee, 5, 1), CStr(DateSerial(annee, 5, 1))         ' Fête du travail
    jourFeries.Add CalculerPaques(annee) + 39, CStr(CalculerPaques(annee) + 39)   ' Ascension
    jourFeries.Add CalculerPaques(annee) + 50, CStr(CalculerPaques(annee) + 50)   ' Pentecôte (lundi)
    jourFeries.Add DateSerial(annee, 7, 21), CStr(DateSerial(annee, 7, 21))       ' Fête nationale (BE)
    jourFeries.Add DateSerial(annee, 8, 15), CStr(DateSerial(annee, 8, 15))       ' Assomption
    jourFeries.Add DateSerial(annee, 11, 1), CStr(DateSerial(annee, 11, 1))       ' Toussaint
    jourFeries.Add DateSerial(annee, 11, 11), CStr(DateSerial(annee, 11, 11))     ' Armistice
    jourFeries.Add DateSerial(annee, 12, 25), CStr(DateSerial(annee, 12, 25))     ' Noël
    On Error GoTo 0
    
    ' Boucle sur les 12 mois
    For indexMois = 1 To 12
        
        ' Essaye d'attraper la feuille correspondant au mois (Janv, Fev, ...)
        Set feuilleActuelle = Nothing
        On Error Resume Next
        Set feuilleActuelle = Sheets(moisFrancais(indexMois - 1))
        On Error GoTo 0
        
        If Not feuilleActuelle Is Nothing Then
        
            ' 1. Petit rappel d'année en haut (tu peux adapter la cellule si besoin)
            With feuilleActuelle.Range("B1")
                .value = annee
                .Font.Bold = True
            End With
            
            ' 2. On réaffiche toutes les colonnes C:AG au cas où elles étaient masquées
            feuilleActuelle.Columns("C:AG").Hidden = False
            
            ' 3. On nettoie l'ancien en-tête (jours + numéros + couleurs)
            feuilleActuelle.Range( _
                feuilleActuelle.Cells(ROW_JOUR_SEMAINE, FIRST_DAY_COL), _
                feuilleActuelle.Cells(ROW_NUMERO_JOUR, LAST_DAY_COL) _
            ).ClearContents
            feuilleActuelle.Range( _
                feuilleActuelle.Cells(ROW_JOUR_SEMAINE, FIRST_DAY_COL), _
                feuilleActuelle.Cells(ROW_NUMERO_JOUR, LAST_DAY_COL) _
            ).Interior.Color = xlNone
            
            ' 4. Calcul du nombre de jours du mois
            totalJours = Day(DateSerial(annee, indexMois + 1, 0))
            
            ' 5. On prépare un tableau 2 lignes :
            '    Ligne 1 = nom du jour (Lun, Mar, ...)
            '    Ligne 2 = numéro du jour (1,2,3...)
            Dim arrHeaders() As Variant
            ReDim arrHeaders(1 To 2, 1 To totalJours)
            
            For i = 1 To totalJours
                dateJour = DateSerial(annee, indexMois, i)
                
                ' Weekday(..., vbMonday) renvoie :
                '   1 = Lundi, 2 = Mardi, ..., 7 = Dimanche
                wd = Weekday(dateJour, vbMonday)
                
                arrHeaders(1, i) = nomsJoursFrancais(wd - 1) ' "Lun", "Mar", etc.
                arrHeaders(2, i) = i                         ' 1,2,3,...
            Next i
            
            ' 6. On écrit ces valeurs dans la feuille à partir de C3
            feuilleActuelle.Cells(ROW_JOUR_SEMAINE, FIRST_DAY_COL) _
                .Resize(2, totalJours).value = arrHeaders
            
            ' 7. On recolorie chaque colonne jour par jour
            For col = FIRST_DAY_COL To FIRST_DAY_COL + totalJours - 1
                
                ' On relit la date correspondante via le numéro du jour écrit en ligne 4
                dateJour = DateSerial(annee, indexMois, _
                                      feuilleActuelle.Cells(ROW_NUMERO_JOUR, col).value)
                wd = Weekday(dateJour, vbMonday)
                
                If wd >= 6 Or EstJourFerie(dateJour, jourFeries) Then
                    ' Samedi / Dimanche / férié ? rouge
                    feuilleActuelle.Range( _
                        feuilleActuelle.Cells(ROW_JOUR_SEMAINE, col), _
                        feuilleActuelle.Cells(ROW_NUMERO_JOUR, col) _
                    ).Interior.Color = RGB(255, 0, 0)
                Else
                    ' Jour ouvré ? bleu clair
                    feuilleActuelle.Range( _
                        feuilleActuelle.Cells(ROW_JOUR_SEMAINE, col), _
                        feuilleActuelle.Cells(ROW_NUMERO_JOUR, col) _
                    ).Interior.Color = RGB(204, 229, 255)
                End If
            Next col
            
            ' 8. Masquer les colonnes après le dernier jour du mois
            '    Exemple : février ? masquer du (C + 28) jusqu'à AG
            If totalJours < 31 Then
                feuilleActuelle.Range( _
                    feuilleActuelle.Cells(1, FIRST_DAY_COL + totalJours), _
                    feuilleActuelle.Cells(1, LAST_DAY_COL) _
                ).EntireColumn.Hidden = True
            End If
        
        End If
    Next indexMois
    
    MsgBox "Calendriers générés pour l'année " & annee & " (alignement colonne C corrigé).", vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub


'------------------ FONCTIONS UTILITAIRES ------------------

Private Function CalculerPaques(annee As Long) As Date
    Dim a As Integer, b As Integer, c As Integer
    Dim d As Integer, e As Integer, f As Integer
    Dim g As Integer, h As Integer, i As Integer
    Dim k As Integer, l As Integer, m As Integer
    Dim mois As Integer, jour As Integer
    
    a = annee Mod 19
    b = annee \ 100
    c = annee Mod 100
    d = b \ 4
    e = b Mod 4
    f = (b + 8) \ 25
    g = (b - f + 1) \ 3
    h = (19 * a + b - d - g + 15) Mod 30
    i = c \ 4
    k = c Mod 4
    l = (32 + 2 * e + 2 * i - h - k) Mod 7
    m = (a + 11 * h + 22 * l) \ 451
    mois = (h + l - 7 * m + 114) \ 31
    jour = ((h + l - 7 * m + 114) Mod 31) + 1
    
    CalculerPaques = DateSerial(annee, mois, jour)
End Function

Private Function EstJourFerie(d As Date, feries As Collection) As Boolean
    On Error Resume Next
    Dim tmp As Variant
    tmp = feries(CStr(d))
    If Err.Number = 0 Then
        EstJourFerie = True
    Else
        EstJourFerie = False
    End If
    Err.Clear
    On Error GoTo 0
End Function


