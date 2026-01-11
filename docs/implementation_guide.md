# Guide d'implémentation - Architecture Config-Driven

## Fichiers livrés

| Fichier | Description |
|---------|-------------|
| `tblCFG.csv` | Table de configuration (31 clés) |
| `Module_Config.bas` | Module VBA de lecture config |
| `ModuleModes_ConfigDriven.bas` | Mode Jour/Nuit config-driven |
| `GenerateurCalendrier_Patch.bas` | Instructions de patch calendrier |

---

## Étape 1 : Importer tblCFG dans Feuil_Config

1. Ouvrir `Planning_2026.xlsm`
2. Aller sur la feuille `Feuil_Config`
3. Données > À partir d'un fichier texte/CSV > Sélectionner `tblCFG.csv`
4. Sélectionner la plage importée (A1:D32)
5. Insertion > Tableau
6. Cocher "Mon tableau comporte des en-têtes"
7. Cliquer OK
8. Dans l'onglet "Création de tableau", renommer en `tblCFG`

> **Alternative rapide** : Ouvrir le CSV dans Excel, copier les cellules, coller dans Feuil_Config en A1, puis convertir en tableau.

---

## Étape 2 : Importer Module_Config

1. Ouvrir l'éditeur VBA (Alt+F11)
2. Fichier > Importer un fichier
3. Sélectionner `Module_Config.bas`
4. Vérifier que le module apparaît dans l'arborescence

### Test rapide
Dans la fenêtre Exécution (Ctrl+G) :
```vba
?CfgLong("CFG_Year")
' Résultat attendu: 2026

?CfgText("SHEET_ConfigCalendar")
' Résultat attendu: Config_Calendrier

?CfgBool("VIEW_HideColumnB")
' Résultat attendu: True
```

---

## Étape 3 : Remplacer ModuleModes

### Option A : Remplacement complet
1. Dans VBA, clic droit sur `ModuleModes` > Exporter
2. Renommer en `ModuleModes_BACKUP.bas`
3. Supprimer `ModuleModes` du projet
4. Importer `ModuleModes_ConfigDriven.bas`
5. Renommer le module en `ModuleModes`

### Option B : Modification manuelle
Appliquer les changements décrits dans le module :
- Remplacer les `Array(...)` par `CfgListLong(...)`
- Remplacer les ranges en dur par `CfgText(...)`
- Ajouter les appels `CfgLong("VIEW_Zoom")` etc.

---

## Étape 4 : Patcher GenerateurCalendrier

Ouvrir `GenerateurCalendrier.bas` et appliquer ces modifications :

### Dans `GenererDatesEtJoursPourTousLesMois`

```vba
' SUPPRIMER ces constantes si présentes :
' Const FIRST_DAY_COL As Long = 3
' Const LAST_DAY_COL As Long = 33

' AJOUTER au début de la procédure :
Dim FIRST_DAY_COL As Long
Dim LAST_DAY_COL As Long
Dim ROW_JOUR_SEMAINE As Long
Dim ROW_NUMERO_JOUR As Long

FIRST_DAY_COL = CfgLong("PLN_FirstDayCol")
LAST_DAY_COL = CfgLong("PLN_LastDayCol")
ROW_JOUR_SEMAINE = CfgLong("PLN_Row_DayNames")
ROW_NUMERO_JOUR = CfgLong("PLN_Row_DayNumbers")

' REMPLACER :
annee = Year(Date)  ' ou annee = 2026
' PAR :
annee = CfgLong("CFG_Year")
```

---

## Étape 5 : Patcher Config_Codes (année)

Dans le module où se trouve `MettreAJourFeries` :

```vba
' REMPLACER :
annee = Year(Date)
' PAR :
annee = CfgLong("CFG_Year")
```

---

## Vérification

### Test Mode Jour/Nuit
1. Aller sur une feuille mensuelle (ex: Janv)
2. Exécuter `Mode_Jour`
3. Vérifier : lignes 3-4 visibles, zoom 70%
4. Modifier `VIEW_Zoom` à 80 dans tblCFG
5. Réexécuter `Mode_Jour`
6. ✅ Zoom maintenant à 80% **sans toucher au code**

### Test Calendrier
1. Modifier `CFG_Year` à 2027 dans tblCFG
2. Exécuter `GenererDatesEtJoursPourTousLesMois`
3. ✅ Les dates affichent 2027

---

## Avantages obtenus

| Aspect | Avant | Après |
|--------|-------|-------|
| Maintenance | Modifier VBA | Modifier Excel |
| Risque régression | Élevé | Quasi nul |
| Lisibilité | Technique | Métier |
| Scalabilité | Faible | Excellente |
