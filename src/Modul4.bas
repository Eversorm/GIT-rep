Attribute VB_Name = "Modul4"
'helpers used in different sheets to find row and column of the overall sheet or row and column of a table.


Public Function findColumnByName(Name As Variant, SpreadSheet As String, iHeaderRow As Integer) As Variant
' small helper function to find the column index corresponding to a header. This is used to ensure the correct column is used, even if someone made adjustments to the Spreadsheet!

    lastUsedColumn = Worksheets(SpreadSheet).Cells(iHeaderRow, Worksheets(SpreadSheet).Columns.Count).End(xlToLeft).Column
    
    Dim ColumnsToFind
    
    iNumberOfHeaders = UBound(Name) + 1
    If iNumberOfHeaders = 1 Then
        ReDim ColumnsToFind(iNumberOfHeaders)
        ColumnsToFind(1) = 0
    Else
        ReDim ColumnsToFind(iNumberOfHeaders)
        For i = 1 To iNumberOfHeaders
            ColumnsToFind(i) = 0
        Next i
    End If
    
    For col = 1 To lastUsedColumn
        CellValue = Worksheets(SpreadSheet).Cells(iHeaderRow, col).Value
        For i = 1 To iNumberOfHeaders
            If CellValue = Name(i - 1) Then
               ColumnsToFind(i) = col
            End If
        Next i
    Next col
    
    iHeaderNotFound = 0
    For i = 1 To iNumberOfHeaders
        If ColumnsToFind(i) = 0 And Name(i - 1) <> "" Then
            iHeaderNotFound = i - 1
        End If
    Next i
    
    If iHeaderNotFound <> 0 Then
        String1 = "The column with the header "
        String2 = " was not found in spreadsheet "
        ErrorString = String1 & Name(iHeaderNotFound) & String2 & SpreadSheet
        MsgBox (ErrorString)
    End If
    
    findColumnByName = ColumnsToFind

End Function

Public Function findRowByName(Name As Variant, SpreadSheet As String, columnIndex As Variant) As Variant
' small helper function to find the row index corresponding to a header. This is used to ensure the correct column is used, even if someone made adjustments to the Spreadsheet!

    findRowByName = 0
    lastUsedRow = Worksheets(SpreadSheet).Cells.SpecialCells(xlCellTypeLastCell).row
    
    Dim RowsToFind
    
    iNumberOfNames = UBound(Name) + 1
    If iNumberOfNames = 1 Then
        ReDim RowsToFind(iNumberOfNames)
        RowsToFind(1) = 0
    Else
        ReDim RowsToFind(iNumberOfNames)
        For i = 1 To iNumberOfNames
            RowsToFind(i) = 0
        Next i
    End If
    
    For row = 1 To lastUsedRow
        CellValue = Worksheets(SpreadSheet).Cells(row, columnIndex).Value
        For i = 1 To iNumberOfNames
            If CellValue = Name(i - 1) Then
                RowsToFind(i) = row
            End If
        Next i
        
        If iNumberOfNames = 1 Then
            If RowsToFind(1) <> 0 Then
                iRowNotFound = -1
                Exit For
            End If
        Else
            iRowNotFound = -1
            For i = 1 To iNumberOfNames
                If RowsToFind(i) = 0 Then
                    iRowNotFound = i - 1
                End If
            Next i
            If iRowNotFound = -1 Then
                Exit For
            End If
        End If
    Next row
    
    
    If iRowNotFound <> -1 Then
        String1 = "The row with the name "
        String2 = " was not found in spreadsheet "
        ErrorString = String1 & Name(iRowNotFound) & String2 & SpreadSheet
        MsgBox (ErrorString)
    End If

    findRowByName = RowsToFind
End Function


Public Function GetTableColumnIndex(ByRef lo As ListObject, ByVal headerName As String) As Long
    ' Gives relative index of a table (1, 2, 3...) else gives 0
    On Error Resume Next
    GetTableColumnIndex = Application.Match(headerName, lo.HeaderRowRange, 0)
    On Error GoTo 0
End Function

Public Function GetTableRowIndex(ByRef lo As ListObject, ByVal searchName As String, ByVal columnIndex As Long) As Long
    ' search a row value in a specific column, else it gives 0
    If columnIndex <= 0 Then GetTableRowIndex = 0: Exit Function
    
    On Error Resume Next
    GetTableRowIndex = Application.Match(searchName, lo.DataBodyRange.Columns(columnIndex), 0)
    On Error GoTo 0
End Function
