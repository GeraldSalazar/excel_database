Attribute VB_Name = "EmailValidation"
Option Explicit

Sub checkEmail(strEmail As String)
    Dim emailValidation As Boolean
    Dim finalCom As String
    
    emailValidation = ValidateEmailAddress(strEmail)
    If Not emailValidation Then
        Err.Raise InputErrors.EmailError, "Email input", "Please provide a valid email. Example: aa@aa.com"
    End If
    
    finalCom = Right(strEmail, 4)
    If Not finalCom = ".com" Then
        Err.Raise InputErrors.EmailError, "Email input", "The email should end with .com"
    End If
End Sub

Private Function ValidateEmailAddress(strEmailAddress As String) As Boolean

    On Error GoTo Quit
    
    Dim objRegEx As Object
    Dim blnIsValidEmail As Boolean
    
    Set objRegEx = CreateObject("Vbscript.Regexp")
    objRegEx.IgnoreCase = True
    objRegEx.Global = True
    objRegEx.Pattern = "^((\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*)\s*[;]{0,1}\s*)+$"

    blnIsValidEmail = objRegEx.Test(strEmailAddress)
    ValidateEmailAddress = blnIsValidEmail

    Exit Function
Quit:
    Set objRegEx = Nothing
    If Err.Number <> 0 Then
        Err.Raise InputErrors.OtherError, "Validate Email", "Regular Expression error"
    End If
End Function

