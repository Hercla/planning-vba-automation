Attribute VB_Name = "ModuleModes_ConfigDriven"
Option Explicit

'===============================================================================
' MODULEMODES_CONFIGDRIVEN - Config-Driven Mode Jour/Nuit
'===============================================================================
' Patched version of ModuleModes using tblCFG configuration
' All hardcoded row numbers, zoom levels, and column references are now
' read from Feuil_Config.tblCFG
'===============================================================================
' REQUIRED: Module_Config must be imported first
'===============================================================================

Public Enum ViewMode
    ViewJour = 1
    ViewNuit = 2
End Enum

'-------------------------------------------------------------------------------
' Mode_Jour - Activate Day view mode
'-------------------------------------------------------------------------------
Public Sub Mode_Jour()
    AdjustView ViewJour
End Sub

'-------------------------------------------------------------------------------
' Mode_Nuit - Activate Night view mode
'-------------------------------------------------------------------------------
Public Sub Mode_Nuit()
    AdjustView ViewNuit
End Sub

'-------------------------------------------------------------------------------
' AdjustView - Main view adjustment routine (config-driven)
'-------------------------------------------------------------------------------
Private Sub AdjustView(mode As ViewMode)
    Dim ws As Worksheet
    Dim dynamicRows As Variant
    Dim rowsToHide As Variant
    Dim i As Long
    Dim blockParts() As String
    Dim startRow As Long, endRow As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo CleanUp
    
    Set ws = ActiveSheet
    
    ' 1) Show all rows first
    ws.Rows.Hidden = False
    
    ' 2) Get config for current mode
    If mode = ViewJour Then
        dynamicRows = CfgListLong("VIEW_Jour_DynamicRows")
        rowsToHide = CfgListText("VIEW_Jour_HideBlocks")
    Else
        dynamicRows = CfgListLong("VIEW_Nuit_DynamicRows")
        rowsToHide = CfgListText("VIEW_Nuit_HideBlocks")
    End If
    
    ' 3) Hide specified blocks
    If IsArray(rowsToHide) Then
        For i = LBound(rowsToHide) To UBound(rowsToHide)
            If Len(rowsToHide(i)) > 0 Then
                ' Parse "startRow:endRow" format
                blockParts = Split(rowsToHide(i), ":")
                If UBound(blockParts) >= 1 Then
                    startRow = CLng(blockParts(0))
                    endRow = CLng(blockParts(1))
                    ws.Rows(startRow & ":" & endRow).Hidden = True
                ElseIf UBound(blockParts) = 0 Then
                    ' Single row
                    ws.Rows(CLng(blockParts(0))).Hidden = True
                End If
            End If
        Next i
    End If
    
    ' 4) Auto-hide empty name rows in dynamic range
    AutoHideRowsBasedOnName ws
    
    ' 5) Ensure header rows are ALWAYS visible (from config)
    ws.Rows(CfgText("VIEW_HeaderRows_Keep")).Hidden = False
    
    ' 6) Apply column visibility settings
    If CfgBool("VIEW_HideColumnB") Then
        ws.Columns("B").Hidden = True
    Else
        ws.Columns("B").Hidden = False
    End If
    
    ' 7) Hide menu columns
    If CfgExists("VIEW_MenuCols") Then
        ws.Columns(CfgText("VIEW_MenuCols")).Hidden = True
    End If
    
    ' 8) Set zoom level
    ActiveWindow.Zoom = CfgLong("VIEW_Zoom")
    
    ' 9) Scroll to top-left
    Application.Goto ws.Range("A1"), Scroll:=True

CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    If Err.Number <> 0 Then
        MsgBox "Erreur Mode " & IIf(mode = ViewJour, "Jour", "Nuit") & ": " & Err.Description, vbCritical
    End If
End Sub

'-------------------------------------------------------------------------------
' AutoHideRowsBasedOnName - Hide rows where name column is empty (config-driven)
'-------------------------------------------------------------------------------
Private Sub AutoHideRowsBasedOnName(ws As Worksheet)
    Dim firstRow As Long, lastRow As Long, colName As String
    Dim cell As Range
    Dim checkRange As Range
    
    ' Read configuration
    firstRow = CfgLong("VIEW_AutoHide_FirstRow")
    lastRow = CfgLong("VIEW_AutoHide_LastRow")
    colName = CfgText("VIEW_NameCol_A")
    
    Set checkRange = ws.Range(colName & firstRow & ":" & colName & lastRow)
    
    For Each cell In checkRange
        ' Only process visible rows
        If cell.EntireRow.Hidden = False Then
            If Len(Trim(CStr(cell.Value))) = 0 Then
                cell.EntireRow.Hidden = True
            End If
        End If
    Next cell
End Sub

'-------------------------------------------------------------------------------
' ToggleMode - Switch between Jour and Nuit modes
'-------------------------------------------------------------------------------
Public Sub ToggleMode()
    Static currentMode As ViewMode
    
    If currentMode = ViewJour Then
        currentMode = ViewNuit
    Else
        currentMode = ViewJour
    End If
    
    AdjustView currentMode
End Sub

'-------------------------------------------------------------------------------
' ResetAllRows - Show all rows (useful for debugging/reset)
'-------------------------------------------------------------------------------
Public Sub ResetAllRows()
    Application.ScreenUpdating = False
    ActiveSheet.Rows.Hidden = False
    ActiveSheet.Columns.Hidden = False
    Application.ScreenUpdating = True
End Sub
