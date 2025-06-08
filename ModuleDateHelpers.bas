Attribute VB_Name = "ModuleDateHelpers"
Option Explicit

' Convert a date value to the month sheet name used in this workbook
Public Function MonthToSheetName(d As Date) As String
    Dim arr As Variant
    arr = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", "Juil", "Aout", "Sept", "Oct", "Nov", "Dec")
    MonthToSheetName = arr(Month(d) - 1)
End Function

' Parse a month name (optionally with year) and return the first day of that month.
' Returns CDate(0) if the name cannot be interpreted.
Public Function GetMonthDateFromName(monthNameInput As String, Optional yearVal As Integer = 0) As Date
    Dim monthStr As String
    Dim yearStr As String
    Dim m As Integer
    Dim y As Integer
    Dim parts() As String
    Dim tempMonthName As String

    tempMonthName = Trim(monthNameInput)
    GetMonthDateFromName = CDate(0)

    parts = Split(tempMonthName, " ")
    If UBound(parts) = 0 Then
        monthStr = tempMonthName
        yearStr = ""
    ElseIf UBound(parts) = 1 Then
        monthStr = parts(0)
        If IsNumeric(parts(1)) And Len(parts(1)) = 4 Then
            yearStr = parts(1)
        Else
            monthStr = tempMonthName
            yearStr = ""
        End If
    Else
        monthStr = tempMonthName
        yearStr = ""
    End If

    Select Case LCase(monthStr)
        Case "janvier", "janv"
            m = 1
        Case "février", "fevrier", "févr", "fevr"
            m = 2
        Case "mars"
            m = 3
        Case "avril", "avr"
            m = 4
        Case "mai"
            m = 5
        Case "juin"
            m = 6
        Case "juillet", "juil"
            m = 7
        Case "août", "aout", "aoû", "aou"
            m = 8
        Case "septembre", "sept"
            m = 9
        Case "octobre", "oct"
            m = 10
        Case "novembre", "nov"
            m = 11
        Case "décembre", "decembre", "déc", "dec"
            m = 12
        Case Else
            On Error Resume Next
            Dim tempDate As Date
            tempDate = DateValue("1 " & monthStr & " " & Year(Date))
            If Err.Number = 0 Then
                m = Month(tempDate)
            Else
                Err.Clear
                tempDate = DateValue("1 " & monthStr)
                If Err.Number = 0 Then
                    m = Month(tempDate)
                    If yearStr = "" Then y = Year(tempDate)
                Else
                    m = 0
                End If
            End If
            On Error GoTo 0
    End Select

    If m = 0 Then Exit Function

    If yearStr <> "" Then
        y = CInt(yearStr)
    ElseIf y <> 0 Then
        ' year already extracted
    ElseIf yearVal <> 0 Then
        y = yearVal
    Else
        y = Year(Date)
    End If

    GetMonthDateFromName = DateSerial(y, m, 1)
End Function
