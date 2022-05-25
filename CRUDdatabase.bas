Attribute VB_Name = "CRUDdatabase"
Option Explicit
Sub addRecordTodbSheet()
    Dim rowOffset As Integer
    rowOffset = 2
    
    Dim lastCell As Range
    Dim lastRow As Integer
    lastRow = getLastRow() + rowOffset
    Set lastCell = Sheets("dbSheet").Range("A" & lastRow).End(xlUp).Offset(1, 0)
    
    lastCell.Offset(0, 0).value = getNextCode()
    lastCell.Offset(0, 1).value = dbForm.nameInput.Text
    lastCell.Offset(0, 2).value = dbForm.birthInput.Text
    lastCell.Offset(0, 3).value = dbForm.emailInput.Text
    lastCell.Offset(0, 4).value = dbForm.addrInput.Text
    
    updateDataDisplay
    resetInputFields
End Sub
Function getNextCode() As Long
    
    Dim lastRow As String
    Dim codeValue As String
    Dim cell As String
    Dim codeColum As String
    
    lastRow = getLastRow()
    codeColum = "A"
    cell = codeColum & lastRow
    
    If (lastRow = "1") Then
        codeValue = "0"
    Else
        codeValue = Sheet1.Range(cell).value
    End If
        

    getNextCode = codeValue + 1
    
End Function

Public Sub updateDataDisplay()
    Dim lastRowNumber As Long
    lastRowNumber = getLastRow() + 1
    With dbForm.dbContent
        .RowSource = "dbSheet!A2:E" & lastRowNumber
        .ColumnCount = 5
        .ColumnHeads = True
    End With
End Sub

Sub deleteRecordFromDb()
    Dim rowNumber As Long
    
    rowNumber = recordSelectedRowNum()
    If rowNumber = -1 Then
        Err.Raise InputErrors.NoRecordSelected, "List box", "Can't Delete. No record selected"
        Exit Sub
    End If
    Sheets("dbSheet").rows(rowNumber).Select
    Selection.Delete
    updateDataDisplay
    
End Sub


Private Sub resetInputFields()
    dbForm.nameInput.Text = ""
    dbForm.birthInput.Text = ""
    dbForm.emailInput.Text = ""
    dbForm.addrInput.Text = ""
End Sub

Public Function getLastRow() As Long
    getLastRow = Sheet1.Cells(rows.Count, "A").End(xlUp).Row
End Function


Sub fillInputFields()
    Dim rowNumber As Long
    
    rowNumber = recordSelectedRowNum()
    If rowNumber = -1 Then
        Err.Raise InputErrors.NoRecordSelected, "List box", "Can't edit. No record selected"
    End If
    
    'If record was selected, fill the form inputs with its data
    dbForm.nameInput.value = Sheets("dbSheet").Cells(rowNumber, 2).value
    dbForm.birthInput.value = Sheets("dbSheet").Cells(rowNumber, 3).value
    dbForm.emailInput.value = Sheets("dbSheet").Cells(rowNumber, 4).value
    dbForm.addrInput.value = Sheets("dbSheet").Cells(rowNumber, 5).value
    
        
        
    
End Sub


Function recordSelectedRowNum() As Long
    Dim rowOffset As Integer
    rowOffset = 2
    
    Dim i As Long
    
    
    For i = 0 To getLastRow() - 1
            If dbForm.dbContent.Selected(i) Then
                Sheets("dbSheet").rows(i + 1).Select
                recordSelectedRowNum = i + rowOffset
                Exit Function
            End If
    Next i
    recordSelectedRowNum = -1
End Function

Sub updateRecord()
    Dim rowNumer As Long
    
    rowNumer = recordSelectedRowNum()
    If rowNumer = -1 Then
        Err.Raise InputErrors.NoRecordSelected, "List box", "Can't edit. No record selected"
    End If
    
    Sheets("dbSheet").Cells(rowNumer, 2).value = dbForm.nameInput.value
    Sheets("dbSheet").Cells(rowNumer, 3).value = dbForm.birthInput.value
    Sheets("dbSheet").Cells(rowNumer, 4).value = dbForm.emailInput.value
    Sheets("dbSheet").Cells(rowNumer, 5).value = dbForm.addrInput.value
    
    resetInputFields
    
End Sub

