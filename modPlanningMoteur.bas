Attribute VB_Name = "modPlanningMoteur"
Option Explicit
' =========================================================================
' SECTION 1: CONSTANTES GLOBALES POUR LES LIGNES DU PLANNING
' =========================================================================
Public Const LIGNE_DEBUT_PLANNING_PERSONNEL As Long = 6
Public Const LIGNE_FIN_PLANNING_PERSONNEL As Long = 30  ' Ajuste selon ta feuille réelle
Public Const LIGNE_REMPLACEMENT_DEBUT_JOUR As Long = 40
Public Const LIGNE_REMPLACEMENT_FIN_JOUR As Long = 41
Public Const NB_REMPLACEMENT_JOUR_LIGNES As Long = LIGNE_REMPLACEMENT_FIN_JOUR - LIGNE_REMPLACEMENT_DEBUT_JOUR + 1
Public Const LIGNE_REMPLACEMENT_DEBUT_NUIT As Long = 46
Public Const LIGNE_REMPLACEMENT_FIN_NUIT As Long = 47
Public Const NB_REMPLACEMENT_NUIT_LIGNES As Long = LIGNE_REMPLACEMENT_FIN_NUIT - LIGNE_REMPLACEMENT_DEBUT_NUIT + 1


Sub RemplacementPlanning_FullDynamiqueUltraOptimisee(ws As Worksheet, colDeb As Long)
    Dim nbJours As Long, planningArr As Variant, rempJourArr As Variant, rempNuitArr As Variant
    Dim dateArr As Variant, ferieArr As Variant
    Dim jourCol As Long, slotLigne As Long, iRule As Long, iCode As Long
    Dim effectifMatin As Long, effectifAM As Long, effectifSoir As Long
    Dim manqueMatin As Long, manqueAM As Long, manqueSoir As Long
    Dim normJour As NormeJour, impactCode As ImpactCodeSuggestion
    Dim codePlaced As Boolean, codeCand As String, groupeExclu As Variant

    nbJours = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column - (colDeb - 1)
    If nbJours < 1 Then Exit Sub

    planningArr = ws.Range(ws.Cells(LIGNE_DEBUT_PLANNING_PERSONNEL, colDeb), ws.Cells(LIGNE_FIN_PLANNING_PERSONNEL, colDeb + nbJours - 1)).Value2
    rempJourArr = ws.Range(ws.Cells(LIGNE_REMPLACEMENT_DEBUT_JOUR, colDeb), ws.Cells(LIGNE_REMPLACEMENT_FIN_JOUR, colDeb + nbJours - 1)).Value2
    rempNuitArr = ws.Range(ws.Cells(LIGNE_REMPLACEMENT_DEBUT_NUIT, colDeb), ws.Cells(LIGNE_REMPLACEMENT_FIN_NUIT, colDeb + nbJours - 1)).Value2
    dateArr = ws.Range(ws.Cells(4, colDeb), ws.Cells(4, colDeb + nbJours - 1)).Value2
    ferieArr = ws.Range(ws.Cells(5, colDeb), ws.Cells(5, colDeb + nbJours - 1)).Value2

    Dim col As Long
    For col = 1 To nbJours
        Dim jourSemaine As Long, isFerie As Boolean
        If IsDate(dateArr(1, col)) Then
            jourSemaine = Weekday(CDate(dateArr(1, col)), vbMonday)
        Else
            jourSemaine = ((ws.Cells(4, col + colDeb - 1).Column - colDeb) Mod 7) + 1
        End If
        isFerie = EstCodeJourFerieOuRecup(CStr(ferieArr(1, col)))
        normJour = GetNormesPourJour(jourSemaine, isFerie)
        effectifMatin = 0: effectifAM = 0: effectifSoir = 0
        Dim r As Long, cellVal As String
        For r = LBound(planningArr, 1) To UBound(planningArr, 1)
            cellVal = Trim(CStr(planningArr(r, col)))
            impactCode = GetImpactDuCode(cellVal)
            effectifMatin = effectifMatin + impactCode.AjouteMatin
            effectifAM = effectifAM + impactCode.AjouteAM
            effectifSoir = effectifSoir + impactCode.AjouteSoir
        Next r
        For r = LBound(rempJourArr, 1) To UBound(rempJourArr, 1)
            cellVal = Trim(CStr(rempJourArr(r, col)))
            impactCode = GetImpactDuCode(cellVal)
            effectifMatin = effectifMatin + impactCode.AjouteMatin
            effectifAM = effectifAM + impactCode.AjouteAM
            effectifSoir = effectifSoir + impactCode.AjouteSoir
        Next r
        manqueMatin = normJour.Matin - effectifMatin: If manqueMatin < 0 Then manqueMatin = 0
        manqueAM = normJour.AM - effectifAM: If manqueAM < 0 Then manqueAM = 0
        manqueSoir = normJour.soir - effectifSoir: If manqueSoir < 0 Then manqueSoir = 0

        ' === BOUCLE REMPLACEMENTS JOUR ===
        For slotLigne = 1 To NB_REMPLACEMENT_JOUR_LIGNES
            If Trim(CStr(rempJourArr(slotLigne, col))) = "" Then
                codePlaced = False
                For iRule = GetLBoundReglesComblement To GetUBoundReglesComblement
                    Dim regle As RegleComblement
                    regle = GetRegleComblementByIndex(iRule)
                    If regle.NomRegle <> "" Then
                        If EvaluerConditionManque(manqueMatin, regle.ManqueMatinOp, regle.ManqueMatinVal) And _
                           EvaluerConditionManque(manqueAM, regle.ManqueAMOp, regle.ManqueAMVal) And _
                           EvaluerConditionManque(manqueSoir, regle.ManqueSoirOp, regle.ManqueSoirVal) Then
                            For iCode = LBound(regle.CodesCandidats) To UBound(regle.CodesCandidats)
                                codeCand = regle.CodesCandidats(iCode)
                                groupeExclu = GetGroupeExclusivitePourCode(codeCand)
                                If Not IsEmpty(groupeExclu) Then
                                    Dim excluPresent As Boolean: excluPresent = False
                                    Dim excluCode As Variant
                                    For Each excluCode In groupeExclu
                                        If CodeDejaPresent(planningArr, rempJourArr, col, excluCode, True) Then
                                            excluPresent = True: Exit For
                                        End If
                                    Next excluCode
                                    If excluPresent Then GoTo NextCodeCand_Jour
                                End If
                                If Not CodeDejaPresent(planningArr, rempJourArr, col, codeCand, True) Then
                                    rempJourArr(slotLigne, col) = codeCand
                                    impactCode = GetImpactDuCode(codeCand)
                                    manqueMatin = manqueMatin - impactCode.AjouteMatin: If manqueMatin < 0 Then manqueMatin = 0
                                    manqueAM = manqueAM - impactCode.AjouteAM: If manqueAM < 0 Then manqueAM = 0
                                    manqueSoir = manqueSoir - impactCode.AjouteSoir: If manqueSoir < 0 Then manqueSoir = 0
                                    codePlaced = True
                                    Exit For
                                End If
NextCodeCand_Jour:
                            Next iCode
                        End If
                    End If
                    If codePlaced Then Exit For
                Next iRule
            End If
        Next slotLigne

        ' === BOUCLE REMPLACEMENTS NUIT (entièrement paramétrable) ===
        For slotLigne = 1 To NB_REMPLACEMENT_NUIT_LIGNES
            If Trim(CStr(rempNuitArr(slotLigne, col))) = "" Then
                codePlaced = False
                For iRule = GetLBoundReglesComblementNuit To GetUBoundReglesComblementNuit
                    Dim regleNuit As RegleComblementNuit
                    regleNuit = GetRegleComblementNuitByIndex(iRule)
                    If regleNuit.NomRegle <> "" Then
                        For iCode = LBound(regleNuit.CodesCandidats) To UBound(regleNuit.CodesCandidats)
                            codeCand = regleNuit.CodesCandidats(iCode)
                            If Not CodeDejaPresent(planningArr, rempNuitArr, col, codeCand, True) Then
                                rempNuitArr(slotLigne, col) = codeCand
                                codePlaced = True
                                Exit For
                            End If
                        Next iCode
                    End If
                    If codePlaced Then Exit For
                Next iRule
            End If
        Next slotLigne
    Next col

    ' --- ÉCRITURE UNIQUE ARRAYS ---
    ws.Range(ws.Cells(LIGNE_REMPLACEMENT_DEBUT_JOUR, colDeb), ws.Cells(LIGNE_REMPLACEMENT_FIN_JOUR, colDeb + nbJours - 1)).Value2 = rempJourArr
    ws.Range(ws.Cells(LIGNE_REMPLACEMENT_DEBUT_NUIT, colDeb), ws.Cells(LIGNE_REMPLACEMENT_FIN_NUIT, colDeb + nbJours - 1)).Value2 = rempNuitArr
End Sub

' --- Wrapper d’exécution silencieuse sur la feuille active ---
Sub LancerRemplacement_FullDynamique()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim colDeb As Long: colDeb = 2 ' à ajuster si besoin (colonne B)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Call RemplacementPlanning_FullDynamiqueUltraOptimisee(ws, colDeb)
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

