VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadCompositionResults 
   Caption         =   "Load Results"
   ClientHeight    =   2925
   ClientLeft      =   -15
   ClientTop       =   -270
   ClientWidth     =   3750
   OleObjectBlob   =   "LoadCompositionResults.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoadCompositionResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_Click()
    
    Unload LoadCompositionResults
    
End Sub

Private Sub LoadResults_Click()

' Load results from the corresponding slot in the output data sheet, this basically copies the stored system config and executes the calculations:
Dim iResultsHeaderRow As Integer
Dim iResultsColumn As Integer
Dim Columns
iResultsHeaderRow = 6
Columns = findColumnByName(Array(Results.Value), "Output Data", iResultsHeaderRow + 1)
iResultsColumn = Columns(1)

Dim miSystemRows
miSystemRows = findRowByName(Array("Humidity Control System:", "CO2 Removal System:", "Electrolysis System:", "CO2 Reduction System:", "Urine Processing System:", "WW Filtration System:", "Brine Processing System:", "O2 Storage:", "H2 Storage:", "Plant Growth System:", "Atmosphere Monitoring:", "Trace Contaminant Control:", "PBR:", "CH4 Processing System:"), "Output Data", 1)
 
With ThisWorkbook.Worksheets("ECLSS Composition")
    .HumidityControlList.Value = Worksheets("Output Data").Cells(miSystemRows(1), iResultsColumn).Value
    .CO2_RemovalList.Value = Worksheets("Output Data").Cells(miSystemRows(2), iResultsColumn).Value
    .ElectrolysisList.Value = Worksheets("Output Data").Cells(miSystemRows(3), iResultsColumn).Value
    .CO2ReductionList.Value = Worksheets("Output Data").Cells(miSystemRows(4), iResultsColumn).Value
    .UrineProcessingList.Value = Worksheets("Output Data").Cells(miSystemRows(5), iResultsColumn).Value
    .WWFiltrationList.Value = Worksheets("Output Data").Cells(miSystemRows(6), iResultsColumn).Value
    .BrineProcessingList.Value = Worksheets("Output Data").Cells(miSystemRows(7), iResultsColumn).Value
    .O2StorageList.Value = Worksheets("Output Data").Cells(miSystemRows(8), iResultsColumn).Value
    .H2StorageList.Value = Worksheets("Output Data").Cells(miSystemRows(9), iResultsColumn).Value
    .PlantGrowthChamberList.Value = Worksheets("Output Data").Cells(miSystemRows(10), iResultsColumn).Value
    .AtmosphereMonitoringList.Value = Worksheets("Output Data").Cells(miSystemRows(11), iResultsColumn).Value
    .TraceContaminantControlList.Value = Worksheets("Output Data").Cells(miSystemRows(12), iResultsColumn).Value
    .PBRList.Value = Worksheets("Output Data").Cells(miSystemRows(13), iResultsColumn).Value
    .MethaneProcessingList.Value = Worksheets("Output Data").Cells(miSystemRows(14), iResultsColumn).Value
    
End With

ThisWorkbook.Worksheets("ECLSS Composition").CalculateECLSS

Unload LoadCompositionResults
    
End Sub

Private Sub UserForm_Initialize()

Dim i As Integer

For i = 1 To 10

    If IsEmpty(Worksheets("Output Data").Cells(7, 2 + i).Value) Then
        Exit For
    End If
    LoadCompositionResults.Results.AddItem Worksheets("Output Data").Cells(7, 2 + i).Value
Next i
End Sub

