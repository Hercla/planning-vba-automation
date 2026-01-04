' ExportedAt: 2026-01-04 17:02:16 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "modMajDate"
Sub Maj_date()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim replaceInfo As Variant
    Dim i As Long
    Dim cellRange As Range
    
    ' Spécifiez le nom de la feuille de calcul
    Set ws = ThisWorkbook.Worksheets("NomDeLaFeuille") ' Remplacez par le nom réel de votre feuille
    
    ' Tableau des informations de remplacement : chaque sous-tableau contient la plage, le texte à remplacer et le texte de remplacement
    replaceInfo = Array( _
        Array("F46:F59", "F d-m", "F j-m"), _
        Array("K46:K59", "F d-m", "F j-m"), _
        Array("G46:G59", "R d-m", "R j-m"), _
        Array("L46:L59", "R d-m", "R j-m") _
    )
    
    ' Boucle à travers chaque ensemble de remplacement et appliquer la fonction Replace
    For i = LBound(replaceInfo) To UBound(replaceInfo)
        ' Essayez de définir la plage
        On Error Resume Next
        Set cellRange = ws.Range(replaceInfo(i)(0))
        On Error GoTo ErrorHandler
        
        ' Vérifiez si la plage existe
        If Not cellRange Is Nothing Then
            ' Effectuer le remplacement
            cellRange.Replace What:=replaceInfo(i)(1), Replacement:=replaceInfo(i)(2), _
                LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        Else
            MsgBox "La plage " & replaceInfo(i)(0) & " n'existe pas sur la feuille.", vbExclamation, "Erreur"
        End If
    Next i
    
    ' Sauvegarder le classeur
    ThisWorkbook.Save
    
    MsgBox "Remplacements effectués avec succès et le classeur a été sauvegardé.", vbInformation, "Succès"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Une erreur s'est produite : " & Err.Description, vbExclamation, "Erreur"
End Sub


