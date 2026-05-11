VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub color()

Dim cell As Range

For Each cell In Worksheets("Graphical Output").Range("D1:Z100")
    cell.Interior.ColorIndex = 2
Next

End Sub

Sub createESMColumnPlot()
    
Dim bFirstSeries As Boolean
Dim iResult As Integer
Dim i As Integer
Dim iResultsHeaderRow As Integer
Dim ChtObj
Dim ChtRng2
Dim ChtRngX2
Dim Ser

iResultsHeaderRow = 6

' set the ESM per Subsystem plot
Set ChtObj = ThisWorkbook.Worksheets("Graphical Output").ChartObjects("ECLSS ESM Composition")
' First clear the old data
For Each Ser In ChtObj.Chart.SeriesCollection
    Ser.Delete
Next

With ThisWorkbook.Worksheets("Output Data")
For i = 1 To 7
    ' add a new series to chart
    Set Ser = ChtObj.Chart.SeriesCollection.NewSeries
    Ser.Name = .Cells(iResultsHeaderRow + 33 + i, 1)
    bFirstSeries = True
    
    ' Check all non empty results and add them to plot, this cannot be done as the solution above, as it is ordered completly different within excel for this chart
    For iResult = 1 To 10
        If .Cells(iResultsHeaderRow + 1, 2 + iResult).Value <> "" Then
            If bFirstSeries Then
                Set ChtRng2 = .Range(.Cells(iResultsHeaderRow + 33 + i, 2 + iResult), .Cells(iResultsHeaderRow + 33 + i, 2 + iResult))
                Set ChtRngX2 = .Range(.Cells(iResultsHeaderRow + 1, 2 + iResult), .Cells(iResultsHeaderRow + 1, 2 + iResult))
                bFirstSeries = False
            Else
                Set ChtRng2 = Union(ChtRng2, .Range(.Cells(iResultsHeaderRow + 33 + i, 2 + iResult), .Cells(iResultsHeaderRow + 33 + i, 2 + iResult)))
                Set ChtRngX2 = Union(ChtRngX2, .Range(.Cells(iResultsHeaderRow + 1, 2 + iResult), .Cells(iResultsHeaderRow + 1, 2 + iResult)))
            End If
        End If
    Next iResult
    
    ' set the source data of the new series
    ' "=" & ChtRng2.Address(False, False, xlA1, xlExternal)
    Ser.Values = ChtRng2
    Ser.XValues = ChtRngX2
Next i
End With
End Sub
