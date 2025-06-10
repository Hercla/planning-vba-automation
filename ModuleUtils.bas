Attribute VB_Name = "ModuleUtils"
Option Explicit

' Central utility functions used across the project

Public Function IsJourFerieOuRecup(code As String) As Boolean
    Dim joursFeries As Variant
    joursFeries = Array("F 1-1", "F 8-5", "F 14-7", "F 15-8", "F 1-11", "F 11-11", _
                       "F 25-12", "R 8-5", "R 1-1", "ASC", "PENT", "L PENT", "L PAQ")
    IsJourFerieOuRecup = IsInArray(code, joursFeries)
End Function

Public Function CodeDejaPresent(planningArr As Variant, rempArr As Variant, jourCol As Long, codeToCheck As String, Optional exactMatch As Boolean = False) As Boolean
    CodeDejaPresent = False

    If CheckSingleArrayForCode(planningArr, jourCol, codeToCheck, exactMatch) Then
        CodeDejaPresent = True
        Exit Function
    End If

    If CheckSingleArrayForCode(rempArr, jourCol, codeToCheck, exactMatch) Then
        CodeDejaPresent = True
        Exit Function
    End If
End Function

Private Function CheckSingleArrayForCode(arrToCheck As Variant, jourCol As Long, codeToCheck As String, exactMatch As Boolean) As Boolean
    Dim r As Long, cellVal As String
    CheckSingleArrayForCode = False
    If IsArray(arrToCheck) Then
        If LBound(arrToCheck, 1) <= UBound(arrToCheck, 1) And LBound(arrToCheck, 2) <= UBound(arrToCheck, 2) Then
            If jourCol >= LBound(arrToCheck, 2) And jourCol <= UBound(arrToCheck, 2) Then
                For r = LBound(arrToCheck, 1) To UBound(arrToCheck, 1)
                    On Error Resume Next
                    cellVal = Trim(CStr(arrToCheck(r, jourCol)))
                    On Error GoTo 0

                    If exactMatch Then
                        If StrComp(cellVal, codeToCheck, vbTextCompare) = 0 Then
                            CheckSingleArrayForCode = True
                            Exit Function
                        End If
                    Else
                        If InStr(1, cellVal, codeToCheck, vbTextCompare) > 0 Then
                            CheckSingleArrayForCode = True
                            Exit Function
                        End If
                    End If
                Next r
            End If
        End If
    End If
End Function

Private Function IsInArray(val As String, arr As Variant) As Boolean
    Dim i As Long
    IsInArray = False
    If Not IsArray(arr) Then Exit Function
    If LBound(arr) > UBound(arr) Then Exit Function

    For i = LBound(arr) To UBound(arr)
        If StrComp(val, CStr(arr(i)), vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next i
End Function

