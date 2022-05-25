Attribute VB_Name = "ExportData"
Option Explicit

Sub ExportDataToNewFile(topRange As Long, bottonRange As Long)
    Dim wbImport As Workbook, wbExport As Workbook
    Dim wsImport As Worksheet, wsExport As Worksheet
    Dim relativePath As String
    '~~> Source/Input Workbook
    Set wbImport = ThisWorkbook
    '~~> Set the sheet to copy
    Set wsImport = wbImport.Sheets("exportSheet")

    '~~> Destination/Output Workbook
    Set wbExport = Workbooks.Add

    With wbExport
        
        Set wsExport = wbExport.Sheets("Sheet1")
        'Save the file
        
        relativePath = ThisWorkbook.Path & Application.PathSeparator & "exportedData.xls"
        .SaveAs Filename:=relativePath

        '~~> Copy the range
        wsImport.Range("A1:E" & topRange).Copy
 
        wsExport.Range("A1:E" & topRange).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    End With
End Sub

Sub exportDatabyName(queryName As String)
    Dim keyWordL As Long
    Dim data As String
    Dim totalRows As Long
    totalRows = Sheets("dbSheet").Range("A" & rows.Count).End(xlUp).Row
    
    Dim i As Long
    Dim x As Long
    Dim initials As String
    For i = 1 To totalRows
        initials = Left(Sheets("dbSheet").Cells(i, 2).value, 2)
        If initials = queryName Then
            Sheets("dbSheet").Range("A" & i + 1 & ":E" & i + 1).Copy Destination:=Sheets("exportSheet").Range("A" & i + 1 & ":E" & i + 1)
        End If
    Next i
    ExportDataToNewFile i, 0
    
End Sub

