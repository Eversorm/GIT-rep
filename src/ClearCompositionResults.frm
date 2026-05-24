VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ClearCompositionResults 
   Caption         =   "Clear Results"
   ClientHeight    =   1590
   ClientLeft      =   -330
   ClientTop       =   -1260
   ClientWidth     =   1575
   OleObjectBlob   =   "ClearCompositionResults.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ClearCompositionResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub NoButton_Click()
    Unload ClearCompositionResults
End Sub

Private Sub YesButton_Click()

    Dim iHeaderRow As Integer
    Dim iHeaderRow2 As Integer
    
    iHeaderRow = 6
    iHeaderRow2 = 92
    
    If ClearCompositionResults.ClearResultSelection.Value = "All" Then
        ' Add code to clear the results
        With ThisWorkbook.Worksheets("Output Data")
            
            Set ChtRng = .Range(.Cells(iHeaderRow + 1, 3), .Cells(iHeaderRow + 68, 12))
            ChtRng.ClearContents
            
            Set ChtRng = .Range(.Cells(iHeaderRow + 71, 3), .Cells(iHeaderRow + 84, 12))
            ChtRng.ClearContents
            
            Set ChtRng = .Range(.Cells(iHeaderRow2 + 1, 3), .Cells(iHeaderRow2 + 1002, 12))
            ChtRng.ClearContents

        End With
        
        Set ChtObj = ThisWorkbook.Worksheets("Graphical Output").ChartObjects("ESM System Mass")
        For Each Series In ChtObj.Chart.SeriesCollection
            Series.Delete
        Next
            
        Set ChtObj = ThisWorkbook.Worksheets("Graphical Output").ChartObjects("ECLSS ESM Composition")
        For Each Series In ChtObj.Chart.SeriesCollection
            Series.Delete
        Next
    Else
        
        Dim i As Integer
        Dim bFoundRow As Integer
        
        bFoundRow = False
        
        For i = 1 To 10
            If Worksheets("Output Data").Cells(7, 2 + i).Value = ClearCompositionResults.ClearResultSelection.Value Then
                bFoundRow = True
                Exit For
            End If
        Next i
        
        If bFoundRow = False Then
            MsgBox ("Selected Result was not found in the output data sheet!")
            Exit Sub
        End If
        
        Dim iRowToDelete As Integer
        iRowToDelete = i + 2
        
        With ThisWorkbook.Worksheets("Output Data")
        
            Set ChtRng = .Range(.Cells(iHeaderRow + 1, iRowToDelete), .Cells(iHeaderRow + 68, iRowToDelete))
            ChtRng.ClearContents
            
            Set ChtRng = .Range(.Cells(iHeaderRow + 71, iRowToDelete), .Cells(iHeaderRow + 84, iRowToDelete))
            ChtRng.ClearContents
            
            Set ChtRng = .Range(.Cells(iHeaderRow2 + 1, iRowToDelete), .Cells(iHeaderRow2 + 1002, iRowToDelete))
            ChtRng.ClearContents
'
'            Set ChtRng = .Range(.Cells(iHeaderRow + 2, iRowToDelete), .Cells(iHeaderRow + 56, iRowToDelete))
'            ChtRng.ClearContents
'
'            Set ChtRng = .Range(.Cells(iHeaderRow + 1, iRowToDelete), .Cells(iHeaderRow + 1, iRowToDelete))
'            ChtRng.ClearContents
'
'            Set ChtRng = .Range(.Cells(iHeaderRow + 59, iRowToDelete), .Cells(iHeaderRow + 72, iRowToDelete))
'            ChtRng.ClearContents
'
'            Set ChtRng = .Range(.Cells(iHeaderRow2 + 2, iRowToDelete), .Cells(iHeaderRow2 + 1002, iRowToDelete))
'            ChtRng.ClearContents
'
'            Set ChtRng = .Range(.Cells(iHeaderRow2 + 1, iRowToDelete), .Cells(iHeaderRow2 + 1, iRowToDelete))
'            ChtRng.ClearContents
        End With
        
        Set ChtObj = ThisWorkbook.Worksheets("Graphical Output").ChartObjects("ESM System Mass")
        For Each Series In ChtObj.Chart.SeriesCollection
            If Series.Name = ClearCompositionResults.ClearResultSelection.Value Then
                Series.Delete
            End If
        Next
        
        ThisWorkbook.Worksheets("Graphical Output").createESMColumnPlot
        
    End If
    
    Unload ClearCompositionResults
    
End Sub


Private Sub UserForm_Initialize()

Dim i As Integer
i = 1
ClearCompositionResults.ClearResultSelection.AddItem "All"

ClearCompositionResults.ClearResultSelection.Value = "All"

For i = 1 To 10
    If Not IsEmpty(Worksheets("Output Data").Cells(7, 2 + i).Value) Then
        ClearCompositionResults.ClearResultSelection.AddItem Worksheets("Output Data").Cells(7, 2 + i).Value
    End If
Next i


End Sub

