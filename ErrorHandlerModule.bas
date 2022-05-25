Attribute VB_Name = "ErrorHandlerModule"
Option Explicit
Enum InputErrors
    BlankField = 514
    BirthDateError = 515
    EmailError = 516
    SearchError = 517
    NoRecordSelected = 518
    ExportError = 519
    OtherError = 520
End Enum

Sub ErrorHandler()
    Select Case Err.Number
    Case InputErrors.BlankField:
        GoTo BlankFieldHandler
    Case InputErrors.BirthDateError:
        GoTo BirthDateHandler
    Case InputErrors.EmailError:
        GoTo EmailHandler
    Case InputErrors.SearchError:
        GoTo SearchHandler
    Case InputErrors.NoRecordSelected:
        GoTo NoRecordSelectedHandler
    Case InputErrors.ExportError:
        GoTo ExportErrorHandler
    Case Else:
        GoTo OtherError
        
    End Select
    
BlankFieldHandler:
    MsgBox (Err.Description & ". Source: " & Err.Source)
    'Input with no text
    Err.Clear
    Exit Sub
BirthDateHandler:
    MsgBox (Err.Description & ". Source: " & Err.Source)
    'Data not in mm/dd/yyyy format
    Err.Clear
    Exit Sub
EmailHandler:
    MsgBox (Err.Description & ". Source: " & Err.Source)
    'Incorrect Email
    Err.Clear
    Exit Sub
SearchHandler:
    MsgBox (Err.Description & ". Source: " & Err.Source)
    'Search Error
    Err.Clear
    Exit Sub
NoRecordSelectedHandler:
    MsgBox (Err.Description & ". Source: " & Err.Source)
    'Search Error
    Err.Clear
    Exit Sub
ExportErrorHandler:
    MsgBox (Err.Description & ". Source: " & Err.Source)
    'Search Error
    Err.Clear
    Exit Sub
OtherError:
    Debug.Print "Unknown error"
    'Unknown
    Err.Clear
    Exit Sub
    
End Sub
    


Function GetErrorMsg(no As Long)
    Select Case no
        Case BlankField:
            GetErrorMsg = "Blank field"
        Case CustomErr2:
            GetErrorMsg = "This is CustomErr2"
    End Select
End Function
