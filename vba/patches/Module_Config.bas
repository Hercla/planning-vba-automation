Attribute VB_Name = "Module_Config"
Option Explicit

'===============================================================================
' MODULE_CONFIG - Configuration Reader for Planning 2026
'===============================================================================
' Centralized configuration management reading from tblCFG table in Feuil_Config
' All hardcoded values should be migrated to tblCFG and read via these functions
'===============================================================================

Private Const CONFIG_SHEET As String = "Feuil_Config"
Private Const CONFIG_TABLE As String = "tblCFG"

'-------------------------------------------------------------------------------
' CfgValue - Generic value reader (returns Variant)
'-------------------------------------------------------------------------------
Public Function CfgValue(ByVal key As String) As Variant
    Dim ws As Worksheet, lo As ListObject
    Dim r As Range
    
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    Set lo = ws.ListObjects(CONFIG_TABLE)
    
    ' Updated to use "Cle" instead of "Key" to match French table
    Set r = lo.ListColumns("Cle").DataBodyRange.Find(What:=key, LookAt:=xlWhole, LookIn:=xlValues)
    If r Is Nothing Then
        Err.Raise vbObjectError + 100, "CfgValue", "Clé config introuvable: " & key
    End If
    
    CfgValue = r.Offset(0, 1).Value ' Column "Value" (Valeur) is next to "Key" (Cle)
    Exit Function
    
ErrHandler:
    Err.Raise vbObjectError + 101, "CfgValue", "Erreur lecture config [" & key & "]: " & Err.Description
End Function

'-------------------------------------------------------------------------------
' CfgText - String reader
'-------------------------------------------------------------------------------
Public Function CfgText(ByVal key As String) As String
    CfgText = CStr(CfgValue(key))
End Function

'-------------------------------------------------------------------------------
' CfgLong - Long integer reader
'-------------------------------------------------------------------------------
Public Function CfgLong(ByVal key As String) As Long
    CfgLong = CLng(CfgValue(key))
End Function

'-------------------------------------------------------------------------------
' CfgDouble - Double reader (for decimals)
'-------------------------------------------------------------------------------
Public Function CfgDouble(ByVal key As String) As Double
    CfgDouble = CDbl(CfgValue(key))
End Function

'-------------------------------------------------------------------------------
' CfgBool - Boolean reader (TRUE/VRAI/1 = True, else False)
'-------------------------------------------------------------------------------
Public Function CfgBool(ByVal key As String) As Boolean
    Dim v As Variant
    v = CfgValue(key)
    CfgBool = (UCase$(Trim$(CStr(v))) = "TRUE" Or _
               UCase$(Trim$(CStr(v))) = "VRAI" Or _
               Trim$(CStr(v)) = "1")
End Function

'-------------------------------------------------------------------------------
' CfgSheet - Returns Worksheet object from sheet name stored in config
'-------------------------------------------------------------------------------
Public Function CfgSheet(ByVal key As String) As Worksheet
    Set CfgSheet = ThisWorkbook.Worksheets(CfgText(key))
End Function

'-------------------------------------------------------------------------------
' CfgListLong - Returns array of Long from comma-separated string
' Example: "6,7,8,9,10" -> Array(6, 7, 8, 9, 10)
'-------------------------------------------------------------------------------
Public Function CfgListLong(ByVal key As String) As Variant
    Dim parts() As String, result() As Long
    Dim i As Long, cleanVal As String
    
    cleanVal = Replace(CfgText(key), " ", "")
    If Len(cleanVal) = 0 Then
        CfgListLong = Array()
        Exit Function
    End If
    
    parts = Split(cleanVal, ",")
    ReDim result(LBound(parts) To UBound(parts))
    
    For i = LBound(parts) To UBound(parts)
        result(i) = CLng(parts(i))
    Next i
    
    CfgListLong = result
End Function

'-------------------------------------------------------------------------------
' CfgListText - Returns array of String from delimited string
' Example: "5:28;39:45" with sep=";" -> Array("5:28", "39:45")
'-------------------------------------------------------------------------------
Public Function CfgListText(ByVal key As String, Optional ByVal sep As String = ";") As Variant
    Dim val As String
    val = CfgText(key)
    If Len(val) = 0 Then
        CfgListText = Array()
    Else
        CfgListText = Split(val, sep)
    End If
End Function

'-------------------------------------------------------------------------------
' CfgExists - Check if a key exists in config (without raising error)
'-------------------------------------------------------------------------------
Public Function CfgExists(ByVal key As String) As Boolean
    Dim ws As Worksheet, lo As ListObject
    Dim r As Range
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    Set lo = ws.ListObjects(CONFIG_TABLE)
    Set r = lo.ListColumns("Cle").DataBodyRange.Find(What:=key, LookAt:=xlWhole, LookIn:=xlValues)
    On Error GoTo 0
    
    CfgExists = Not (r Is Nothing)
End Function

'-------------------------------------------------------------------------------
' CfgValueDefault - Returns value or default if key not found
'-------------------------------------------------------------------------------
Public Function CfgValueDefault(ByVal key As String, ByVal defaultVal As Variant) As Variant
    If CfgExists(key) Then
        CfgValueDefault = CfgValue(key)
    Else
        CfgValueDefault = defaultVal
    End If
End Function
