VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsPlants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' this sub gets the plants properties from the table plants
Public Sub GetPlantProperties( _
    ByRef PlantNames As Variant, _
    ByRef PlantCarbonContent As Variant, _
    ByRef PlantO2Production As Variant, _
    ByRef PlantCO2Consumption As Variant, _
    ByRef PlantWaterUptake As Variant, _
    ByRef PlantEdibleBiomassGrowth As Variant, _
    ByRef PlantInedibleBiomassGrowth As Variant, _
    ByRef PlantEdibleBiomassWaterContent As Variant, _
    ByRef PlantEnergyContent As Variant)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Plants and PGC")

    ' Reference the named table directly
    Dim lo As ListObject
    Set lo = ws.ListObjects("Plants")

    ' Search header row using InStr to find each column dynamically
    Dim colName As Long, colCarbon As Long, colO2 As Long
    Dim colCO2 As Long, colWater As Long, colEdible As Long
    Dim colInedible As Long, colWaterContent As Long, colEnergy As Long

    Dim hCell As Range
    For Each hCell In lo.HeaderRowRange
        Dim h As String
        h = hCell.Value
        If InStr(1, h, "PlantName", vbTextCompare) > 0 Then colName = hCell.Column
        If InStr(1, h, "Carbon", vbTextCompare) > 0 Then colCarbon = hCell.Column
        If InStr(1, h, "Production", vbTextCompare) > 0 Then colO2 = hCell.Column  'InStr(1, h, "O2", vbTextCompare) > 0 And
        If InStr(1, h, "Consumption", vbTextCompare) > 0 Then colCO2 = hCell.Column     'InStr(1, h, "CO2", vbTextCompare) > 0 And
        If InStr(1, h, "Uptake", vbTextCompare) > 0 Then colWater = hCell.Column   'InStr(1, h, "Water", vbTextCompare) > 0 And
        If InStr(1, h, "EdibleGrowth", vbTextCompare) > 0 Then colEdible = hCell.Column   'And InStr(1, h, "Growth", vbTextCompare) > 0 _
   And InStr(1, h, "Inedible", vbTextCompare) = 0
        If InStr(1, h, "InedibleGrowth", vbTextCompare) > 0 Then colInedible = hCell.Column   'And InStr(1, h, "Growth", vbTextCompare) > 0
        If InStr(1, h, "Water", vbTextCompare) > 0 And InStr(1, h, "Content", vbTextCompare) > 0 Then colWaterContent = hCell.Column
        If InStr(1, h, "Energy", vbTextCompare) > 0 Then colEnergy = hCell.Column
    Next hCell

    ' Validate that all required columns were found
    If colName = 0 Or colO2 = 0 Or colCO2 = 0 Or colCarbon = 0 Or _
       colWater = 0 Or colEdible = 0 Or colInedible = 0 Or colWaterContent = 0 Or colEnergy = 0 Then
        MsgBox "GetPlantProperties: one or more columns not found in table 'Plants'. Check header names."
        Exit Sub
    End If

    ' Size arrays from the table data body — no hardcoded row count
    Dim nPlants As Long
    nPlants = lo.DataBodyRange.Rows.Count - 1

    ReDim PlantNames(0 To nPlants - 1)
    ReDim PlantCarbonContent(0 To nPlants - 1)
    ReDim PlantO2Production(0 To nPlants - 1)
    ReDim PlantCO2Consumption(0 To nPlants - 1)
    ReDim PlantWaterUptake(0 To nPlants - 1)
    ReDim PlantEdibleBiomassGrowth(0 To nPlants - 1)
    ReDim PlantInedibleBiomassGrowth(0 To nPlants - 1)
    ReDim PlantEdibleBiomassWaterContent(0 To nPlants - 1)
    ReDim PlantEnergyContent(0 To nPlants - 1)

    ' Populate arrays row by row from the table data body
    Dim iPlant As Long
    For iPlant = 0 To nPlants - 1
        Dim r As Long
        r = lo.DataBodyRange.Rows(iPlant + 2).row
        PlantNames(iPlant) = ws.Cells(r, colName).Value
        PlantCarbonContent(iPlant) = ws.Cells(r, colCarbon).Value
        PlantO2Production(iPlant) = ws.Cells(r, colO2).Value
        PlantCO2Consumption(iPlant) = ws.Cells(r, colCO2).Value
        PlantWaterUptake(iPlant) = ws.Cells(r, colWater).Value
        PlantEdibleBiomassGrowth(iPlant) = ws.Cells(r, colEdible).Value
        PlantInedibleBiomassGrowth(iPlant) = ws.Cells(r, colInedible).Value
        PlantEdibleBiomassWaterContent(iPlant) = ws.Cells(r, colWaterContent).Value
        PlantEnergyContent(iPlant) = ws.Cells(r, colEnergy).Value
    Next iPlant

End Sub


'function to get the areas for every plant in every specific PGC from the table, this is called in ECLSS composition code
Public Function GetPGCAreasByName(ByVal pgcName As String, ByRef PlantNames As Variant) As Variant

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Plants and PGC")
    Dim lo As ListObject
    Set lo = ws.ListObjects("PGCConfigs")

    ' 1. find PGC row
    Dim pgcRowIndex As Long
    pgcRowIndex = GetTableRowIndex(lo, pgcName, 1)
    If pgcRowIndex = 0 Then
        MsgBox "The PGC configuration '" & pgcName & "' was not found!", vbExclamation
    End If

    ' 2. initialize array
    Dim areas() As Variant
    ReDim areas(LBound(PlantNames) To UBound(PlantNames))

    ' 3. fill array with the helpers function to find row and column
    Dim iPlant As Long
    Dim colIndex As Long
    
    For iPlant = LBound(PlantNames) To UBound(PlantNames)
        If pgcRowIndex = 0 Then
            areas(iPlant) = 0
        Else
            ' call helper function for the column and get the area value
            colIndex = GetTableColumnIndex(lo, CStr(PlantNames(iPlant)))
            If colIndex > 0 Then
                areas(iPlant) = lo.DataBodyRange.Cells(pgcRowIndex, colIndex).Value
            Else
                areas(iPlant) = 0
            End If
        End If
    Next iPlant

    GetPGCAreasByName = areas
End Function


Sub plantsguide()

    'Worksheets("Plants and PGC").Protect UserInterfaceOnly:=True
    UG_Plants.Width = 550
    UG_Plants.Height = 500
    UG_Plants.Show

End Sub


