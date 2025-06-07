Attribute VB_Name = "ModNotes"
Option Explicit ' Toujours une bonne idée en haut du module

Sub AddOrReplaceComment_Optimized()
    Dim targetCell As Range
    Dim personName As String
    Dim commentContent As String
    Dim fullCommentText As String
    Dim userResponse As VbMsgBoxResult

    ' --- Validation de la sélection ---
    If TypeName(Selection) <> "Range" Then
        MsgBox "Veuillez sélectionner une cellule.", vbExclamation, "Sélection Invalide"
        Exit Sub
    End If

    ' S'assurer qu'une SEULE cellule est sélectionnée
    If Selection.Cells.CountLarge > 1 Then
        MsgBox "Veuillez sélectionner une seule cellule.", vbExclamation, "Sélection Multiple"
        Exit Sub
    End If

    Set targetCell = Selection ' La sélection est valide (une seule cellule)

    ' --- Obtention des informations ---
    personName = InputBox("Entrez le nom de la personne :", "Ajouter/Modifier Note")
    ' Vérifier si l'utilisateur a annulé ou n'a rien saisi
    If StrPtr(personName) = 0 Then ' L'utilisateur a cliqué sur Annuler
        Exit Sub ' Sortir sans message
    ElseIf Len(Trim(personName)) = 0 Then ' L'utilisateur a cliqué OK mais n'a rien écrit (ou que des espaces)
        MsgBox "Le nom de la personne ne peut pas être vide.", vbExclamation, "Nom Requis"
        Exit Sub
    End If
    personName = Trim(personName) ' Nettoyer les espaces

    commentContent = InputBox("Entrez le contenu de la note pour " & personName & " :", "Ajouter/Modifier Note")
     If StrPtr(commentContent) = 0 Then ' L'utilisateur a cliqué sur Annuler
        Exit Sub ' Sortir sans message
    End If
    ' Pas besoin de vérifier si vide, une note vide est peut-être acceptable ? Sinon, ajouter:
    ' If Len(Trim(commentContent)) = 0 Then ...

    ' --- Gestion du commentaire ---
    On Error GoTo ErrorHandler ' Gérer les erreurs potentielles (feuille protégée...)

    ' Vérifier s'il y a déjà un commentaire et demander confirmation si oui
    If Not targetCell.Comment Is Nothing Then
        userResponse = MsgBox("Cette cellule a déjà une note :" & vbCrLf & vbCrLf & _
                              targetCell.Comment.Text & vbCrLf & vbCrLf & _
                              "Voulez-vous la remplacer ?", _
                              vbYesNo + vbQuestion, "Remplacer Note ?")
        If userResponse = vbNo Then Exit Sub ' L'utilisateur ne veut pas remplacer
        targetCell.Comment.Delete ' Supprimer l'ancien commentaire
    End If

    ' Assembler et ajouter le nouveau commentaire
    fullCommentText = personName & ":" & vbCrLf & commentContent ' Met le nom sur une ligne, le contenu après
    targetCell.AddComment Text:=fullCommentText
    ' Optionnel: définir l'auteur (si nécessaire)
    ' targetCell.Comment.Shape.TextFrame.Characters.Font.Name = "Tahoma" ' Exemple de formatage police

    ' --- Affichage et dimensionnement ---
    With targetCell.Comment.Shape
        .TextFrame.AutoSize = True ' Doit être avant de rendre visible pour un meilleur calcul initial
        .Visible = True
        ' Optionnel: Ajuster la taille minimale/maximale si AutoSize ne suffit pas
        ' If .Width < 100 Then .Width = 100 ' Largeur minimale
        ' If .Height > 300 Then .Height = 300 ' Hauteur maximale
    End With

    ' Optionnel: Resélectionner la cellule si l'interaction avec la note l'a désélectionnée
    targetCell.Select

    On Error GoTo 0 ' Désactiver le gestionnaire d'erreur
    Exit Sub

ErrorHandler:
    MsgBox "Impossible de modifier la note." & vbCrLf & _
           "Vérifiez si la feuille est protégée." & vbCrLf & vbCrLf & _
           "Erreur: " & Err.Description, vbCritical, "Erreur Note"
End Sub

