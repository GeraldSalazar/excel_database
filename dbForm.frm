VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dbForm 
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22905
   OleObjectBlob   =   "dbForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dbForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' Force explicit variable declaration.
Private Sub addRecordBtn_Click()
    Call checkBlanks
End Sub
Private Sub checkBlanks()
    On Error GoTo ErrorHandler
    
    If (nameInput.value = "") Then
        Err.Raise InputErrors.BlankField, "Name input", "Please enter a name"
        Exit Sub
    End If
    If (birthInput.value = "") Then
        Err.Raise InputErrors.BlankField, "Birth date input", "Please enter a birth date"
        Exit Sub
    End If
    
    Call formatValidation
    
ErrorHandler:
    Call ErrorHandler
    
End Sub

Private Sub formatValidation()
    
    Dim birthDate As String
    Dim email As String
    
    
    birthDate = birthInput.Text
    checkBirthDate (birthDate) 'Go to birthDateValidation Module
    
    email = emailInput.Text
    If (email <> "") Then
        checkEmail (email)
    End If
    
    Call addRecordTodbSheet
    
    
End Sub






'Button to confirm the update of a particular record
Private Sub confirmEditBtn_Click()
    Dim iConfirmation As VbMsgBoxResult
    iConfirmation = MsgBox("Edit the row selected?", vbQuestion + vbYesNo, "Database")
    
    If iConfirmation = vbYes Then
        Call updateRecord
        toggleEditionBtns
    End If
End Sub

Private Sub deleteRecord_Click()
    On Error GoTo ErrorHandler
    Call deleteRecordFromDb
    Exit Sub
ErrorHandler:
    Call ErrorHandler
End Sub
'Button to display all the data in the list box
Private Sub displayDataBtn_Click()

    updateDataDisplay
End Sub
'Button to fill the form inputs with existing record in order to edit it
Private Sub editRecordBtn_Click()
    On Error GoTo ErrorHandler
    Call fillInputFields
    toggleEditionBtns
    Exit Sub
ErrorHandler:
    Call ErrorHandler
End Sub



Private Sub exportByNameBtn_Click()
    On Error GoTo ErrorHandler
    Dim fromExport As String
    Dim toExport As String
    Dim nameQuery As String
    nameQuery = dbForm.nameImportInput.value
    
    
    If nameQuery <> "" Then
        exportDatabyName nameQuery
        Exit Sub
    Else
        Err.Raise InputErrors.ExportError, "Export Range Inputs", "Please enter a numerical range"
    End If
    
    
ErrorHandler:
    Call ErrorHandler
    
End Sub

Private Sub exportByRangeBtn_Click()
    On Error GoTo ErrorHandler
    Dim fromExport As String
    Dim toExport As String
    
    fromExport = dbForm.fromInput.value
    toExport = dbForm.toInputExport.value
    
    If (IsNumeric(fromExport) And IsNumeric(toExport)) Then
        ExportDataToNewFile CLng(toExport), CLng(fromExport)
    Else
        Err.Raise InputErrors.ExportError, "Export Range Inputs", "Please enter a numerical range"
    End If
    Exit Sub
    
ErrorHandler:
    Call ErrorHandler
    
End Sub



Private Sub searchBtn_Click()
    Dim searchQuery As String
    searchQuery = searchInput.Text
    If (searchQuery = "") Then
        updateDataDisplay
    Else
        Call searchDataByQuery(searchQuery)
    End If
    
    
    searchInput.Text = ""
    
End Sub

Private Sub UserForm_Initialize()

    updateDataDisplay
    
End Sub

Private Sub toggleEditionBtns()
    confirmEditBtn.Enabled = Not confirmEditBtn.Enabled
    addRecordBtn.Enabled = Not confirmEditBtn.Enabled
End Sub

