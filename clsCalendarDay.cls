' ExportedAt: 2026-01-11 14:14:03 | Workbook: Planning_2026.xlsm
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalendarDay"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public WithEvents lblDay As MSForms.label
Attribute lblDay.VB_VarHelpID = -1
Public ParentForm As Object

Private Sub lblDay_Click()
    Dim frm As Object
    Set frm = ParentForm
    Dim selectedDate As Date
    selectedDate = lblDay.Tag
    Dim dateControl As Control
    Set dateControl = frm.Controls(frm.Tag)
    DayClick frm, selectedDate, dateControl
End Sub

