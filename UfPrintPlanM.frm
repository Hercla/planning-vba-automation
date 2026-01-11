' ExportedAt: 2026-01-11 14:14:03 | Workbook: Planning_2026.xlsm
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UfPrintPlanM 
   Caption         =   "Imprimer planning"
   ClientHeight    =   1407
   ClientLeft      =   -119
   ClientTop       =   -399
   ClientWidth     =   455
   OleObjectBlob   =   "UfPrintPlanM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UfPrintPlanM"
Attribute VB_Base = "0{BCA1D722-A852-432F-8FE6-B8EB251C0CB2}{367C4046-1861-477B-A73D-286D0C817AD7}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Sub CBtAnnule_Click()
UfPrintPlanM.Hide

End Sub

Sub CBtImprime_Click()
Dim ZoneImprim

If UfPrintPlanM.OBtJanvier = True Then
    Call PrintPlanMois(ZoneImprim:="$B$19:$G$53")
ElseIf UfPrintPlanM.OBtFevrier = True Then
    Call PrintPlanMois(ZoneImprim:="$I$19:$N$53")
ElseIf UfPrintPlanM.OBtMars = True Then
    Call PrintPlanMois(ZoneImprim:="$P$19:$U$53")
ElseIf UfPrintPlanM.OBtAvril = True Then
    Call PrintPlanMois(ZoneImprim:="$W$19:$AB$53")
ElseIf UfPrintPlanM.OBtMai = True Then
    Call PrintPlanMois(ZoneImprim:="$AD$19:$AI$53")
ElseIf UfPrintPlanM.OBtJuin = True Then
    Call PrintPlanMois(ZoneImprim:="$AK$19:$AP$53")
ElseIf UfPrintPlanM.OBtJuillet = True Then
    Call PrintPlanMois(ZoneImprim:="$AR$19:$AW$53")
ElseIf UfPrintPlanM.OBtAout = True Then
    Call PrintPlanMois(ZoneImprim:="&AY$19:$BD$53")
ElseIf UfPrintPlanM.OBtSeptembre = True Then
    Call PrintPlanMois(ZoneImprim:="$BF$19:$BK$53")
ElseIf UfPrintPlanM.OBtOctobre = True Then
    Call PrintPlanMois(ZoneImprim:="^BM$19:$BBR$53")
ElseIf UfPrintPlanM.OBtNovembre = True Then
    Call PrintPlanMois(ZoneImprim:="$BT$19:$BY$53")
ElseIf UfPrintPlanM.OBtDecembre = True Then
    Call PrintPlanMois(ZoneImprim:="$CA$19:$CF$53")
End If
UfPrintPlanM.Hide

End Sub
Sub UserForm_Activate()
UfPrintPlanM.LbTxNom = ActiveSheet.Range("D2") & " " & ActiveSheet.Range("D1")
End Sub

