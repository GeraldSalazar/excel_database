Attribute VB_Name = "BirthDateValidation"
Option Explicit
Public Sub checkBirthDate(birthDate As String)
    If Len(birthDate) <> 10 Or _
    Mid(birthDate, 3, 1) <> "/" Or _
    Mid(birthDate, 6, 1) <> "/" Then
        Err.Raise InputErrors.BirthDateError, "Birth date input", "Please enter a valid date. The format is: mm/dd/yyyy"
        Exit Sub
    Else
        If Not checkDaysValid(birthDate) Then
            Err.Raise InputErrors.BirthDateError, "Birth date input", "Please enter a valid day"
            Exit Sub
        ElseIf Not checkMonthsValid(birthDate) Then
            Err.Raise InputErrors.BirthDateError, "Birth date input", "Please enter a valid month"
            Exit Sub
        ElseIf Not checkYearsValid(birthDate) Then
            Err.Raise InputErrors.BirthDateError, "Birth date input", "Please enter a valid year"
            Exit Sub
        Else
            Exit Sub
        End If
    End If
End Sub
Private Function checkYearsValid(dateText As String) As Boolean
    Dim years As String
    years = Right(dateText, 4)
    If Not IsNumeric(years) Or _
    years >= Year(Date) Then
        checkYearsValid = False
    Else
        checkYearsValid = True
    End If
End Function
Private Function checkMonthsValid(dateText As String) As Boolean
    Dim months As String
    months = Left(dateText, 2)
    If Not IsNumeric(months) Then
        checkMonthsValid = False
        Exit Function
    End If
    If Not (months > 0 And months <= 12) Then
        checkMonthsValid = False
        Exit Function
    Else
        checkMonthsValid = True
    End If
End Function
Private Function checkDaysValid(dateText As String) As Boolean
    Dim days As String
    days = Mid(dateText, 4, 2)
    If Not IsNumeric(days) Then
        checkDaysValid = False
        Exit Function
    End If
    If Not (days > 0 And days <= 31) Then
        checkDaysValid = False
        Exit Function
    Else
        checkDaysValid = True
    End If
End Function

