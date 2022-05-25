Attribute VB_Name = "SearchData"
Option Explicit

Sub searchDataByQuery(searchQuery As String)
    
    Worksheets("searchResultsSheet").UsedRange.Offset(1).ClearContents
    If IsNumeric(searchQuery) Then
        searchByCode (searchQuery)
    Else
        searchByName (searchQuery)
    End If
    
    displayResults
End Sub

Private Sub searchByCode(code As String)
    Dim rowOffset As Integer
    rowOffset = 2
    Dim totalRows As Long
    Dim x As Long
    Dim recordCode As String
    totalRows = Sheets("dbSheet").Range("A" & rows.Count).End(xlUp).Row + rowOffset
    For x = 2 To totalRows
        recordCode = Sheets("dbSheet").Cells(x, 1).value
        If recordCode = dbForm.searchInput.value Then
            Sheets("dbSheet").Range("A" & x & ":E" & x).Copy Destination:=Sheets("searchResultsSheet").Range("A2:E2")
            
            End If
    Next x
End Sub


Private Sub searchByName(keyWord As String)
    Dim keyWordL As Long
    Dim data As String
    Dim totalRows As Long
    totalRows = Sheets("dbSheet").Range("A" & rows.Count).End(xlUp).Row
    
    Dim i As Long
    Dim x As Long
    Dim index As Long
    index = 2
    For i = 2 To totalRows
        For x = 1 To Len(Sheet1.Cells(i, 2))
            keyWordL = Len(keyWord)
            data = Mid(Sheets("dbSheet").Cells(i, 2), x, keyWordL)
            
            If LCase(data) = keyWord Then
                Sheets("dbSheet").Range("A" & i + 1 & ":E" & i + 1).Copy Destination:=Sheets("searchResultsSheet").Range("A" & index & ":E" & index)
            End If
        Next x
        index = index + 1
    Next i
    
End Sub

Private Sub displayResults()
    Dim lastRowNumber As Long
    lastRowNumber = Sheets("searchResultsSheet").Cells(rows.Count, "A").End(xlUp).Row + 1
    With dbForm.dbContent
        .RowSource = "searchResultsSheet!A2:E" & lastRowNumber
        .ColumnCount = 5
        .ColumnHeads = True
    End With
End Sub
