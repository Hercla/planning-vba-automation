' ExportedAt: 2026-01-11 14:14:03 | Workbook: Planning_2026.xlsm
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ManageWorkTimeForm 
   Caption         =   "UserForm3"
   ClientHeight    =   3437
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4802
   OleObjectBlob   =   "ManageWorkTimeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ManageWorkTimeForm"
Attribute VB_Base = "0{1EB40802-28DD-4EE0-B914-43775B32EE7D}{6124330E-79DF-4A3C-A00B-42E01AD12C73}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub btnStartDate_Click()
    ToggleDatePicker Me, Me.Controls("txtStartDate")
End Sub

Private Sub btnEndDate_Click()
    ToggleDatePicker Me, Me.Controls("txtEndDate")
End Sub

Private Sub cmbNom_Change()
    Dim ws As Worksheet
    Dim i As Long
    Set ws = ThisWorkbook.Sheets("Personnel")
    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        If ws.Cells(i, 2).value = Me.Controls("cmbNom").value Then
            Me.Controls("cmbPrenom").value = ws.Cells(i, 3).value
            Exit For
        End If
    Next i
End Sub

Private Sub cmbPrenom_Change()
    Dim ws As Worksheet
    Dim i As Long
    Set ws = ThisWorkbook.Sheets("Personnel")
    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        If ws.Cells(i, 3).value = Me.Controls("cmbPrenom").value Then
            Me.Controls("cmbNom").value = ws.Cells(i, 2).value
            Exit For
        End If
    Next i
End Sub

Private Sub UserForm_Initialize()
    FillEmployeeComboBoxes
End Sub

