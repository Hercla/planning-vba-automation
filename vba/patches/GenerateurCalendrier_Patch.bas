Attribute VB_Name = "GenerateurCalendrier_Patch"
'===============================================================================
' GENERATEURCALENDRIER_PATCH - Config-Driven Calendar Patch
'===============================================================================
' This file contains PATCHES to apply to GenerateurCalendrier.bas
' Replace the hardcoded constants with these config-driven versions
'===============================================================================
' REQUIRED: Module_Config must be imported first
'===============================================================================

'-------------------------------------------------------------------------------
' PATCH 1: Replace hardcoded layout constants
'-------------------------------------------------------------------------------
' BEFORE (delete these lines):
'   Const FIRST_DAY_COL As Long = 3
'   Const LAST_DAY_COL As Long = 33
'   Const ROW_JOUR_SEMAINE As Long = 3
'   Const ROW_NUMERO_JOUR As Long = 4
'
' AFTER (add these at the start of GenererDatesEtJoursPourTousLesMois):
'
'   Dim FIRST_DAY_COL As Long
'   Dim LAST_DAY_COL As Long
'   Dim ROW_JOUR_SEMAINE As Long
'   Dim ROW_NUMERO_JOUR As Long
'
'   FIRST_DAY_COL = CfgLong("PLN_FirstDayCol")
'   LAST_DAY_COL = CfgLong("PLN_LastDayCol")
'   ROW_JOUR_SEMAINE = CfgLong("PLN_Row_DayNames")
'   ROW_NUMERO_JOUR = CfgLong("PLN_Row_DayNumbers")
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
' PATCH 2: Use CFG_Year instead of hardcoded year
'-------------------------------------------------------------------------------
' BEFORE:
'   annee = Year(Date)
'   ' or
'   annee = 2026
'
' AFTER:
'   annee = CfgLong("CFG_Year")
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
' PATCH 3: Ensure header rows stay visible after generation
'-------------------------------------------------------------------------------
' ADD at the end of the generation loop (after writing dates):
'
'   ' Ensure header rows are always visible
'   feuilleActuelle.Rows(CfgText("VIEW_HeaderRows_Keep")).Hidden = False
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
' EXAMPLE: Complete patched procedure header
'-------------------------------------------------------------------------------
Public Sub GenererDatesEtJoursPourTousLesMois_ConfigDriven()
    ' Configuration variables (read from tblCFG)
    Dim FIRST_DAY_COL As Long
    Dim LAST_DAY_COL As Long
    Dim ROW_JOUR_SEMAINE As Long
    Dim ROW_NUMERO_JOUR As Long
    Dim annee As Long
    
    ' Read configuration
    FIRST_DAY_COL = CfgLong("PLN_FirstDayCol")
    LAST_DAY_COL = CfgLong("PLN_LastDayCol")
    ROW_JOUR_SEMAINE = CfgLong("PLN_Row_DayNames")
    ROW_NUMERO_JOUR = CfgLong("PLN_Row_DayNumbers")
    annee = CfgLong("CFG_Year")
    
    ' ... rest of your existing logic ...
    ' Replace hardcoded values with these variables
    
    ' At the end of each sheet processing:
    ' feuilleActuelle.Rows(CfgText("VIEW_HeaderRows_Keep")).Hidden = False
End Sub

'-------------------------------------------------------------------------------
' PATCH 4: MettreAJourFeries - Use CFG_Year
'-------------------------------------------------------------------------------
' In Module_Configuration.bas or wherever MettreAJourFeries is defined:
'
' BEFORE:
'   annee = Year(Date)
'
' AFTER:
'   annee = CfgLong("CFG_Year")
'-------------------------------------------------------------------------------
