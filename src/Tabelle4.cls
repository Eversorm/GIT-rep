VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' basic design of ECLSS composition:
' based on the partial sketch
' Technologies selectied from drop down menu instead of checkboxes --> no matter if more techs are added IF can remain
' add technologies based on technology sheet, as long as they are assemblies and the relevant parameters for that technology type are defined (e.g. efficiency for waste water)


Sub CreateDropDownLists()
' This sub creates the technology drop down lists


Dim na As Integer
Dim iTechnologyRow As Integer
Dim iHeader As Integer
Dim sSystemAbbreviation As String
Dim sSystemName As String
Dim sTabStopString As String
Dim iListRow As Integer
Dim iShape As Integer

Dim TechnologyFunctionHeader
Dim miTechnologyColumn
Dim miTechnologyColumnFunction
Dim csValidSystemsHumidity
Dim csValidSystemsCO2Removal
Dim csValidSystemsCO2Reduction
Dim csValidSystemsO2Storage
Dim csValidSystemsH2Storage
Dim csValidSystemsElectrolysis
Dim csValidSystemsUrineProcessing
Dim csValidSystemsWWFiltration
Dim csValidSystemsBrineProcessing
Dim csValidSystemsISRU
Dim csValidSystemsPBR
Dim csValidSystemsMethaneProcessing

na = Worksheets("Technology").Range("A1:A1000").Cells.SpecialCells(xlCellTypeConstants).Count + 2

TechnologyFunctionHeader = Array("function1", "function2", "function3", "function4", "function5", "function6", "function7", "function8", "function9", "function10")
miTechnologyColumn = findColumnByName(Array("Name", "Abbreviation", "Level", "Efficiency"), "Technology", 1)
miTechnologyColumnFunction = findColumnByName(TechnologyFunctionHeader, "Technology", 1)

' Remove old entries:
With ThisWorkbook.Worksheets("ECLSS Composition")

' Unfortunatly I did not find an option to reference these elements using an array of names....
If .HumidityControlList.ListCount > 0 Then
    For iListRow = 1 To Worksheets("ECLSS Composition").HumidityControlList.ListCount
        .HumidityControlList.RemoveItem (0)
    Next iListRow
End If
If .CO2_RemovalList.ListCount > 0 Then
    For iListRow = 1 To Worksheets("ECLSS Composition").CO2_RemovalList.ListCount
        .CO2_RemovalList.RemoveItem (0)
    Next iListRow
End If
If .CO2ReductionList.ListCount > 0 Then
    For iListRow = 1 To Worksheets("ECLSS Composition").CO2ReductionList.ListCount
        .CO2ReductionList.RemoveItem (0)
    Next iListRow
End If
If .O2StorageList.ListCount > 0 Then
    For iListRow = 1 To Worksheets("ECLSS Composition").O2StorageList.ListCount
        .O2StorageList.RemoveItem (0)
    Next iListRow
End If
If .H2StorageList.ListCount > 0 Then
    For iListRow = 1 To Worksheets("ECLSS Composition").H2StorageList.ListCount
        .H2StorageList.RemoveItem (0)
    Next iListRow
End If
If .ElectrolysisList.ListCount > 0 Then
    For iListRow = 1 To Worksheets("ECLSS Composition").ElectrolysisList.ListCount
        .ElectrolysisList.RemoveItem (0)
    Next iListRow
End If
If .UrineProcessingList.ListCount > 0 Then
    For iListRow = 1 To Worksheets("ECLSS Composition").UrineProcessingList.ListCount
        .UrineProcessingList.RemoveItem (0)
    Next iListRow
End If
If .WWFiltrationList.ListCount > 0 Then
    For iListRow = 1 To Worksheets("ECLSS Composition").WWFiltrationList.ListCount
        .WWFiltrationList.RemoveItem (0)
    Next iListRow
End If
If .ISRUList.ListCount > 0 Then
    For iListRow = 1 To Worksheets("ECLSS Composition").ISRUList.ListCount
        .ISRUList.RemoveItem (0)
    Next iListRow
End If
If .BrineProcessingList.ListCount > 0 Then
    For iListRow = 1 To Worksheets("ECLSS Composition").BrineProcessingList.ListCount
        .BrineProcessingList.RemoveItem (0)
    Next iListRow
End If
If .PlantGrowthChamberList.ListCount > 0 Then
    For iListRow = 1 To Worksheets("ECLSS Composition").PlantGrowthChamberList.ListCount
        .PlantGrowthChamberList.RemoveItem (0)
    Next iListRow
End If
If .AtmosphereMonitoringList.ListCount > 0 Then
    For iListRow = 1 To Worksheets("ECLSS Composition").AtmosphereMonitoringList.ListCount
        .AtmosphereMonitoringList.RemoveItem (0)
    Next iListRow
End If
If .TraceContaminantControlList.ListCount > 0 Then
    For iListRow = 1 To Worksheets("ECLSS Composition").TraceContaminantControlList.ListCount
        .TraceContaminantControlList.RemoveItem (0)
    Next iListRow
End If
If .PBRList.ListCount > 0 Then
    For iListRow = 1 To Worksheets("ECLSS Composition").PBRList.ListCount
        .PBRList.RemoveItem (0)
    Next iListRow
End If
If .MethaneProcessingList.ListCount > 0 Then
    For iListRow = 1 To Worksheets("ECLSS Composition").MethaneProcessingList.ListCount
        .MethaneProcessingList.RemoveItem (0)
    Next iListRow
End If

' Add option none to optional subsystems:
.CO2ReductionList.AddItem ("None")
.WWFiltrationList.AddItem ("None")
.UrineProcessingList.AddItem ("None")
.BrineProcessingList.AddItem ("None")
.ElectrolysisList.AddItem ("None")
.ISRUList.AddItem ("None")
.PlantGrowthChamberList.AddItem ("None")
.H2StorageList.AddItem ("None")
.PBRList.AddItem ("None")
.MethaneProcessingList.AddItem ("None")
                
' List of valid system for the composition calculation, if you want to add a system, provide the required calculations and add the abbreviation to this list
csValidSystemsHumidity = Array("CAMRAS", "CCAA", "2BMS")
csValidSystemsCO2Removal = Array("LiOH", "METOX", "CAMRAS", "2BMS", "4BMS/CDRA", "SAWD", "CDRILS")
csValidSystemsCO2Reduction = Array("SR/SCRA", "BR")
csValidSystemsPBR = Array("PBR")
csValidSystemsMethaneProcessing = Array("PPU", "LPU")
' csValidSystemsO2Storage = Array("CS", "HPS")
' csValidSystemsH2Storage = Array("CS", "HPS")
' For Electrolysis, nothing else must be defined, since the basic reactions are always the same!
' For urine, brine and water treatment, the efficiencies of water recovery must be provided, so for those we do not check these arrays, but if the user provided efficiencies
'Dim csValidSystemsWWFiltration
'Dim csValidSystemsBrineProcessing
'Dim csValidSystemsISRU

For iTechnologyRow = 3 To na
    For iHeader = 1 To UBound(miTechnologyColumnFunction)
        ' We check all functions of the current technology and use the maximum required scaling factor for any of the functions it handles (e.g. the CCAA handles humidity and heat, and the larger of these two difference will be used to scale it)
        If Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumn(3)).Value = "assembly" Then
        
            sSystemAbbreviation = Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumn(2)).Value
            sSystemName = Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumn(1)).Value
            
            ' We want a nicely indented list of the full system names behind the abbreviations, so we have to check how many tab stops we need
'            If Len(sSystemAbbreviation) <= 7 Then
'                sTabStopString = vbTab & vbTab & vbTab & vbTab
'            Else
'                sTabStopString = vbTab & vbTab & vbTab
'            End If
'                    .CO2_RemovalList.AddItem (sSystemAbbreviation & sTabStopString & "(" & sSystemName & ")")
            
            ' "CO2 Removal", "CO2 Reduction", "Process Wastewater a)" , "Process Wastewater c)", "Process Wastewater d)", "Remove or Add Moisture", "Electrolyse Water"
            If Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value = "CO2 Removal" Then
                If IsNumeric(Application.Match(sSystemAbbreviation, csValidSystemsCO2Removal, 0)) Then
                    .CO2_RemovalList.AddItem (sSystemAbbreviation)
                End If
                    
            ElseIf Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value = "CO2 Reduction" Then
                If IsNumeric(Application.Match(sSystemAbbreviation, csValidSystemsCO2Reduction, 0)) Then
                    .CO2ReductionList.AddItem (sSystemAbbreviation)
                End If
                
            ElseIf Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value = "PBR" Then
                If IsNumeric(Application.Match(sSystemAbbreviation, csValidSystemsPBR, 0)) Then
                    .PBRList.AddItem (sSystemAbbreviation)
                End If
                
            ElseIf Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value = "Methane Processing" Then
                If IsNumeric(Application.Match(sSystemAbbreviation, csValidSystemsMethaneProcessing, 0)) Then
                    .MethaneProcessingList.AddItem (sSystemAbbreviation)
                End If
                
                
            ElseIf Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value = "Remove or Add Moisture" Then
                If IsNumeric(Application.Match(sSystemAbbreviation, csValidSystemsHumidity, 0)) Then
                    .HumidityControlList.AddItem (sSystemAbbreviation)
                End If
                
            ElseIf Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value = "Process Wastewater a)" Then
                If Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumn(4)).Value <> "" Then
                    .WWFiltrationList.AddItem (sSystemAbbreviation)
                End If
                
            ElseIf Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value = "Process Wastewater c)" Then
                If Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumn(4)).Value <> "" Then
                    .UrineProcessingList.AddItem (sSystemAbbreviation)
                End If
                
            ElseIf Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value = "Process Wastewater d)" Then
                If Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumn(4)).Value <> "" Then
                    .BrineProcessingList.AddItem (sSystemAbbreviation)
                End If
                
            ElseIf Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value = "Electrolyse Water" Then
                
                .ElectrolysisList.AddItem (sSystemAbbreviation)
                
            ElseIf Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value = "Store Oxygen" Then
                
                .O2StorageList.AddItem (sSystemAbbreviation)
                
            ElseIf Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value = "Store Inert Gas" Then
                
                .H2StorageList.AddItem (sSystemAbbreviation)
               
            ElseIf Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value = "Produce Food a)" Then
                If Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumn(4)).Value <> "" Then
                    .PlantGrowthChamberList.AddItem (sSystemAbbreviation)
                End If
                
            ElseIf Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value = "Detect Hazardous Atmosphere" Then
                .AtmosphereMonitoringList.AddItem (sSystemAbbreviation)
                
            ElseIf Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value = "Remove Gaseous Atmospheric Contaminants" Then
                .TraceContaminantControlList.AddItem (sSystemAbbreviation)
                
            End If
        End If
    Next iHeader
Next iTechnologyRow

End With

End Sub

Function CheckSelectedECLSSComposition() As Boolean
' This sub is used to check the selected ECLSS compositions for errors. For example if one system is capabale of providing multiple functions it must be used for all these functions
' (e.g. CAMRAS at humidity Control and CO2 Removal  "2BMS"

CheckSelectedECLSSComposition = True

' Check if schedule is up to date
If Worksheets("User Interface").Range("Schedule_Number").Value = 2 Then
    If Worksheets("Schedule and Metabolic").Range("Schedule_was_Checked").Value <> 1 Then
        MsgBox ("The user schedule has not been checked since the last change, please go to the sheet 'Schedule and Metabolic' and click on 'Check Schedule'")
        CheckSelectedECLSSComposition = False
        Exit Function
    ElseIf Worksheets("Schedule and Metabolic").Range("Metabolic_Values_up_to_date").Value <> 1 Then
        MsgBox ("The user schedule metabolic values have not been updated since the last change, please go to the sheet 'Schedule and Metabolic' and click on 'Calculate Metabolics'")
        CheckSelectedECLSSComposition = False
        Exit Function
    End If
End If

With ThisWorkbook.Worksheets("ECLSS Composition")
If InStr(1, .HumidityControlList.Value, "CAMRAS", vbBinaryCompare) Or InStr(1, .CO2_RemovalList.Value, "CAMRAS", vbBinaryCompare) Or InStr(1, .HumidityControlList.Value, "2BMS", vbBinaryCompare) Or InStr(1, .CO2_RemovalList.Value, "2BMS", vbBinaryCompare) Then
    ' if 2BMS or CAMRAS are used for either humidity control or CO2 removal, it must also be selected for the other option, otherwise throw error
    If InStr(1, .HumidityControlList.Value, "CAMRAS", vbBinaryCompare) And InStr(1, .CO2_RemovalList.Value, "CAMRAS", vbBinaryCompare) Then
    ElseIf InStr(1, .HumidityControlList.Value, "2BMS", vbBinaryCompare) And InStr(1, .CO2_RemovalList.Value, "2BMS", vbBinaryCompare) Then
    Else
        MsgBox ("If CAMRAS or 2BMS are used for either humidity control or CO2 removal, the same technology must be selected for both functions!")
        CheckSelectedECLSSComposition = False
        Exit Function
    End If
End If

If InStr(1, .CO2_RemovalList.Value, "LiOH", vbBinaryCompare) Then
     If InStr(1, .CO2ReductionList.Value, "None", vbBinaryCompare) Then
     Else
        ' If LiOH is used no CO2 reduction is possible, because LiOH does not produce CO2 but rather LiCO3
        MsgBox ("If LiOH is used for CO2 removal no CO2 reduction can be selected, because LiOH produces LiCO3 as output, not CO2!")
        CheckSelectedECLSSComposition = False
     End If
End If

If InStr(1, .PlantGrowthChamberList.Value, "Veggie", vbBinaryCompare) Or InStr(1, .PlantGrowthChamberList.Value, "ISPR PGC", vbBinaryCompare) Then
    If .PGC_NumberOfUnits <> "" Then
    Else
        MsgBox ("For Veggie and ISPR PGC please specify how many of these units shall be used, as they are not scaled per crew member!")
        CheckSelectedECLSSComposition = False
        Exit Function
    End If
End If


If InStr(1, .UrineProcessingList.Value, "None", vbBinaryCompare) Then
    If InStr(1, .BrineProcessingList.Value, "None", vbBinaryCompare) Then
    Else
        MsgBox ("Brine Processing can only be selected if a urine processing technology is selected as well!")
        CheckSelectedECLSSComposition = False
        Exit Function
    End If
End If

' Check crew metabolic values:
Dim CrewUrineProduction As Double
Dim CrewFecesProduction As Double
Dim CrewCO2Production As Double
Dim CrewH2OProduction As Double
Dim CrewO2Consumption As Double
Dim CrewFoodConsumption As Double
Dim CrewH2OinFoodConsumption As Double
Dim CrewH2OConsumption As Double
Dim CrewHygieneH2OConsumption As Double
Dim Crew_Size As Integer

Crew_Size = Worksheets("User Interface").Range("Crew_Size").Value

If Worksheets("User Interface").Range("Schedule_Number").Value = 1 Then
    CrewUrineProduction = Worksheets("Schedule and Metabolic").Range("SUM_Urine_Production").Value
    CrewFecesProduction = Worksheets("Schedule and Metabolic").Range("SUM_Feces_Production").Value
    CrewCO2Production = Worksheets("Schedule and Metabolic").Range("SUM_CO2_Production").Value
    CrewH2OProduction = Worksheets("Schedule and Metabolic").Range("SUM_Sweat_Production").Value
    CrewO2Consumption = Worksheets("Schedule and Metabolic").Range("SUM_O2_Consumption").Value
    CrewFoodConsumption = Worksheets("Schedule and Metabolic").Range("SUM_Food_Need").Value
    CrewH2OConsumption = Worksheets("Schedule and Metabolic").Range("SUM_Potable_Water_Need").Value
    CrewHygieneH2OConsumption = Worksheets("Schedule and Metabolic").Range("SUM_Hygiene_Water_Need").Value
Else
    CrewUrineProduction = Worksheets("Schedule and Metabolic").Range("SUM_Urine_Production_2").Value
    CrewFecesProduction = Worksheets("Schedule and Metabolic").Range("SUM_Feces_Production_2").Value
    CrewCO2Production = Worksheets("Schedule and Metabolic").Range("SUM_CO2_Production_2").Value
    CrewH2OProduction = Worksheets("Schedule and Metabolic").Range("SUM_Sweat_Production_2").Value
    CrewO2Consumption = Worksheets("Schedule and Metabolic").Range("SUM_O2_Consumption_2").Value
    CrewFoodConsumption = Worksheets("Schedule and Metabolic").Range("SUM_Food_Need_2").Value
    CrewH2OConsumption = Worksheets("Schedule and Metabolic").Range("SUM_Potable_Water_Need_2").Value
    CrewHygieneH2OConsumption = Worksheets("Schedule and Metabolic").Range("SUM_Hygiene_Water_Need_2").Value
End If
CrewH2OinFoodConsumption = CrewFoodConsumption * 0.7 / 1.51

' Divide value with crew to get value per CM
CrewUrineProduction = CrewUrineProduction / Crew_Size + Worksheets("Schedule and Metabolic").Range("SUM_Urine_Solid_Waste").Value
.Range("CrewFecesValue").Value = CrewFecesProduction / Crew_Size
.Range("CO2_ProdCrewValue").Value = CrewCO2Production / Crew_Size
.Range("H2O_ProdCrewValue").Value = CrewH2OProduction / Crew_Size
.Range("O2_ConCrewValue").Value = CrewO2Consumption / Crew_Size
.Range("FoodConCrewValue").Value = CrewFoodConsumption / Crew_Size
.Range("H2O_FoodConCrewValue").Value = CrewH2OinFoodConsumption / Crew_Size
.Range("H2O_ConCrewValue").Value = CrewH2OConsumption / Crew_Size
.Range("HygieneH2O_ConCrewValue").Value = CrewHygieneH2OConsumption / Crew_Size
.Range("HygieneH2O_ProdCrewValue").Value = CrewHygieneH2OConsumption / Crew_Size

' Urine depends in selected option:
If InStr(1, .UrineProcessingList.Value, "None", vbBinaryCompare) Then
    .Range("CrewUrineWithoutProcessing").Value = CrewUrineProduction
Else
    .Range("CrewUrineWithProcessing").Value = CrewUrineProduction
End If


' add leakage
Dim LeakValues

LeakValues = calculateLeakage
.Range("O2_LeakageValue").Value = LeakValues(1)
.Range("N2_LeakageValue").Value = LeakValues(2)


'Check if String inputs are Numbers (PGR & PBR) and in correct Range)
If Not InStr(1, .PlantGrowthChamberList.Value, "None", vbBinaryCompare) Then
    If Not IsNumeric(PGC_NumberOfUnits.Value) Then
        MsgBox ("Please enter a valid Number in 'Number of PGR Units'")
        CheckSelectedECLSSComposition = False
        Exit Function
    End If
End If

If Not InStr(1, .PBRList.Value, "None", vbBinaryCompare) Then
    If Not IsNumeric(PBR_CO2Percentage.Value) Then
        MsgBox ("Please enter a valid Number (between 0 and 100) in 'Percentage of CO2 going to PBR'")
        CheckSelectedECLSSComposition = False
        Exit Function
    End If
    If Not (PBR_CO2Percentage.Value >= 0 And PBR_CO2Percentage.Value <= 100) Then
        MsgBox ("Please enter a valid Number (between 0 and 100) in 'Percentage of CO2 going to PBR'")
        CheckSelectedECLSSComposition = False
        Exit Function
    End If
End If


' Comment of all lists, currently no further limitations are known
'        .HumidityControlList.Value
'        .CO2_RemovalList.Value
'        .CO2ReductionList.Value
'        .O2StorageList.Value
'        .H2StorageList.Value
'        .ElectrolysisList.Value
'        .UrineProcessingList.Value
'        .WWFiltrationList.Value
'        .ISRUList.Value
'        .BrineProcessingList.Value
'        .AtmosphereMonitoringList.Value
'        .TraceContaminantControlList.Value
End With
End Function


Sub CalculateECLSS()
' This sub is called to actually calculate the kg/day and CM values for the selected ECLSS
Dim bValidSystem As Boolean

bValidSystem = CheckSelectedECLSSComposition
If bValidSystem = False Then
    Exit Sub
End If

' now get the efficiency column from the technology sheet

Dim miTechnologyColumn
miTechnologyColumn = findColumnByName(Array("Abbreviation", "Efficiency", "Resupply Mass"), "Technology", 1)

With ThisWorkbook.Worksheets("ECLSS Composition")
' The calculation always follows the same scheme, but depends on the used technologies for some cases, while others (like electrolysis) are independent of the used technology

' first check if plants are used, if that is the case the remaining calculations must be adjusted, so we first calculate the plants

Dim O2_fromStorage As Double
Dim CO2_toRemoval  As Double
Dim H2O_toHumidityControl  As Double

'-------------------------------------------------------------------------------------------------------------------------------------------------------------



'-------------------------------------------------------------------------------------------------------------------------------------------------------------
' initialize crew urine production:

Dim UrineSolids As Double
Dim fUrineWaterToPlants As Double
Dim CrewUrineProduction As Double
Dim Crew_Size As Integer

Crew_Size = Worksheets("User Interface").Range("Crew_Size").Value

UrineSolids = Worksheets("Schedule and Metabolic").Range("SUM_Urine_Solid_Waste").Value

If Worksheets("User Interface").Range("Schedule_Number").Value = 1 Then
    CrewUrineProduction = Worksheets("Schedule and Metabolic").Range("SUM_Urine_Production").Value
Else
    CrewUrineProduction = Worksheets("Schedule and Metabolic").Range("SUM_Urine_Production_2").Value
End If
' Divide value with crew to get value per CM
CrewUrineProduction = CrewUrineProduction / Crew_Size + UrineSolids

.Range("CrewUrineWithProcessing").Value = CrewUrineProduction


If InStr(1, .PlantGrowthChamberList.Value, "None", vbBinaryCompare) Then
' if no PGC is used we do not have to adjust anything, the values interfacing with the cabin are set to the base values from the crew
    O2_fromStorage = .Range("O2_ConCrewValue").Value + .Range("O2_LeakageValue").Value
    CO2_toRemoval = .Range("CO2_ProdCrewValue").Value
    H2O_toHumidityControl = .Range("H2O_ProdCrewValue").Value
    
    .Range("O2_toIncinerationLabel").Font.color = vbWhite
    .Range("O2_toIncinerationValue").Font.color = vbWhite
    .Range("O2_toIncinerationValue").Value = 0
    .Range("CO2_FromWasteToCabinValue").Font.color = vbWhite
    .Range("CO2_FromWasteToCabinValue").Value = 0
    .Shapes("CO2_FromWasteToCabin1").Visible = False
    .Shapes("CO2_FromWasteToCabin2").Visible = False
Else
    '-------------------------------------------------------------------------
    ' Plant growth chamber calculation
    ' All these variables are used as arrays and contain the information in the same order as the BVAD tables from top to bottom, except for mushrooms, as no values for it are provided
    Dim PlantNames
    Dim PlantO2Production
    Dim PlantCO2Consumption
    Dim PlantWaterUptake
    Dim PlantEdibleBiomassGrowth
    Dim PlantInedibleBiomassGrowth
    Dim PlantEdibleBiomassWaterContent
    Dim PlantEnergyContent
    Dim PGCGrowthAreas As Variant
    Dim TotalGrowthArea As Double
    Dim iIndex As Integer
    Dim iI As Integer
    Dim iNumberOfUnits As Integer
    Dim CrewSize As Integer
    Dim CO2ToPlants As Double
    Dim O2FromPlants As Double
    Dim baseTotalGrowthArea As Double
    Dim InedibleBiomass As Double
    Dim CarbonContentBiomass As Double
    Dim PlantCarbonContent
    Dim AlgeaCarbonContent
    
    ' See BVAD table 4.99 for O2, CO2 and H2O interfaces per m^2 and day for different plants. Of course these values are normally dynamic
    ' See BVAD table 4.100 for growth rates of edible and inedible per m^2 and day
    ' values for chufa, cucumber and strawberry are from "An optimized 4-day diet meal plan for ‘Lunar Palace 1’", Yuming Fu et al. DOI 10.1002/jsfa.9234
    ' but the reported lunar palace values are missing the corresponding O2, CO2, H2O and inedible biomass values and therefore complemented by BVAD data of similar species
    ' Since chufa is a Monocotyledons like wheat, the values for wheat are used for Chufa
    ' cucumbers uses snap bean data due to the lack of a closer resemblance
    

    
    'Get plant data from Plants table in same sheet
    Call Tabelle12.GetPlantProperties(PlantNames, PlantCarbonContent, PlantO2Production, PlantCO2Consumption, _
    PlantWaterUptake, PlantEdibleBiomassGrowth, PlantInedibleBiomassGrowth, PlantEdibleBiomassWaterContent, PlantEnergyContent)
    
    ReDim PGCGrowthAreas(LBound(PlantNames) To UBound(PlantNames))
    ' Initialize all areas to 0 as fallback baseline
    For iI = LBound(PlantNames) To UBound(PlantNames)
        PGCGrowthAreas(iI) = 0
    Next iI
       
    CrewSize = Worksheets("User Interface").Range("Crew_Size").Value
    
    Dim miTechnologyPlantRow
    miTechnologyPlantRow = findRowByName(Array(.PlantGrowthChamberList.Value), "Technology", miTechnologyColumn(1))
    TotalGrowthArea = Worksheets("Technology").Cells(miTechnologyPlantRow(1), miTechnologyColumn(2)).Value
    iNumberOfUnits = CInt(.PGC_NumberOfUnits.Value)
    
    Dim matchResult As Variant
    Dim pgcType As String
    pgcType = .PlantGrowthChamberList.Value
    
    

' Process specific predefined PGC setups or call dynamic database lookup
    If InStr(1, pgcType, "Veggie", vbBinaryCompare) > 0 Then
        ' Veggie Configuration: 100% Lettuce
        matchResult = Application.Match("Lettuce", PlantNames, 0)
        If Not IsError(matchResult) Then
            PGCGrowthAreas(matchResult - 1 + LBound(PlantNames)) = iNumberOfUnits * TotalGrowthArea
        Else
            MsgBox "Error: Lettuce column missing from plant configuration data.", vbCritical
            Exit Sub
        End If
        ' _
        iIndex = Application.Match("Lettuce", PlantNames, False) _
        iNumberOfUnits = CInt(.PGC_NumberOfUnits.Value)
        ' Here we divide with the crew size because the units are seen as absolut, so for the values to be per CM we have to divide with the crew size
        'PGCGrowthAreas(iIndex) = iNumberOfUnits * TotalGrowthArea
    ElseIf InStr(1, .PlantGrowthChamberList.Value, "ISPR PGC", vbBinaryCompare) Then
        ' here 50% lettuce and 50% tomato is used
        iIndex = Application.Match("Lettuce", PlantNames, False) - 1
        'iNumberOfUnits = CInt(.PGC_NumberOfUnits.Value)
        PGCGrowthAreas(iIndex) = iNumberOfUnits * 0.5 * TotalGrowthArea
        ' Here we divide with the crew size because the units are seen as absolut, so for the values to be per CM we have to divide with the crew size
        iIndex = Application.Match("Tomato", PlantNames, False) - 1
        PGCGrowthAreas(iIndex) = iNumberOfUnits * 0.5 * TotalGrowthArea
        
    Else
        ' all others configurations getting the datas from the table in the Plants sheet
        PGCGrowthAreas = Tabelle12.GetPGCAreasByName(.PlantGrowthChamberList.Value, PlantNames)
    End If

' ElseIf InStr(1, .PlantGrowthChamberList.Value, "ISPR PGC", vbBinaryCompare) Then
'     ' here 50% lettuce and 50% tomato is used
'     iIndex = Application.Match("Lettuce", PlantNames, False)
'     iNumberOfUnits = CInt(.PGC_NumberOfUnits.Value)
'     ' Here we divide with the crew size because the units are seen as absolut, so for the values to be per CM we have to divide with the crew size
'     iIndex = Application.Match("Tomato", PlantNames, False)
'     PGCGrowthAreas(iIndex) = iNumberOfUnits * 0.5 * TotalGrowthArea
'
' ElseIf InStr(1, .PlantGrowthChamberList.Value, "Scaled Part ELS PGC", vbBinaryCompare) Then
'     PGCGrowthAreas = Array(0.086938, 1.4094, 0, 0.19232, 3.0823, 1.0959, 0.42152, 0, 0.81932, 0, 0.54797, 0, 0, 0, 0.17651, 0, 1.4437, 0, 9.168, 3.1851, 25.499, 4.252, 0, 0)
'
' ElseIf InStr(1, .PlantGrowthChamberList.Value, "Part ELS PGC", vbBinaryCompare) Then
'     PGCGrowthAreas = Array(0.033, 0.535, 0, 0.073, 1.17, 0.416, 0.16, 0, 0.311, 0, 0.208, 0, 0, 0, 0.067, 0, 0.548, 0, 3.48, 1.209, 9.679, 1.614, 0, 0)
'
' ElseIf InStr(1, .PlantGrowthChamberList.Value, "Lunar Palace PGC", vbBinaryCompare) Then
'     PGCGrowthAreas = Array(0, 1.9981, 0, 0, 0, 0.39961, 4.3478, 0, 0, 0, 0, 0, 0, 0, 1.5185, 4.4757, 0, 0.79922, 0, 0, 31.969, 0, 7.9922, 0.047953)
'
' ElseIf InStr(1, .PlantGrowthChamberList.Value, "ELS PGC", vbBinaryCompare) Then
'     PGCGrowthAreas = Array(0, 1.3476, 0, 0, 4.8423, 0.69392, 0.14331, 0, 0, 12.149, 0, 0.41233, 0, 5.2245, 0, 116.73, 1.5965, 0, 3.7336, 4.1283, 10.653, 2.4991, 0, 0)
'
' ElseIf InStr(1, .PlantGrowthChamberList.Value, "Case 1", vbBinaryCompare) Then
'     PGCGrowthAreas = Array(0, 0, 10.551, 0, 0, 0, 30, 0, 0, 4.1116, 14.812, 0, 0, 0, 25.589, 0, 30, 0, 0, 0, 0, 0, 0, 0)
'
' ElseIf InStr(1, .PlantGrowthChamberList.Value, "Case 2", vbBinaryCompare) Then
'     PGCGrowthAreas = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 3.487, 17.004, 0, 0, 0, 32.128, 0, 54.243, 0, 0, 0, 0, 0, 0, 0)
'
' ElseIf InStr(1, .PlantGrowthChamberList.Value, "Case 3", vbBinaryCompare) Then
'     PGCGrowthAreas = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 30, 9.881, 0, 0)
'
' ElseIf InStr(1, .PlantGrowthChamberList.Value, "Case 4", vbBinaryCompare) Then
'     PGCGrowthAreas = Array(0, 0, 0, 0, 10, 0, 0, 0, 0, 0, 0, 0, 0, 7.0389, 10, 0, 0, 0, 10, 0, 10, 10, 0, 0)
'
' Else
'     ' All other chambers are assumed to use the optimized Exploration PGC growth areas
'     PGCGrowthAreas = Array(0, 0, 0, 0, 10, 0, 2, 0, 0, 10, 0, 0, 0, ```2, 0, 0, 0, 0, 2, 5, 18, 8, 0, 0)


    ' calculate the CO2 consumption for one m^2 of this plant composition
    
    
    baseTotalGrowthArea = WorksheetFunction.Sum(PGCGrowthAreas)
    

    If baseTotalGrowthArea = 0 Then
        MsgBox "PGC growth area is zero. Check PGCConfigs table or selected PGC name."
        Exit Sub
    End If

    CO2ToPlants = 0
    O2FromPlants = 0
    CarbonContentBiomass = 0
    For iI = 0 To UBound(PlantO2Production)
        CO2ToPlants = CO2ToPlants + ((PGCGrowthAreas(iI) / baseTotalGrowthArea) * PlantCO2Consumption(iI)) / 1000
        O2FromPlants = O2FromPlants + ((PGCGrowthAreas(iI) / baseTotalGrowthArea) * PlantO2Production(iI)) / 1000
        
        InedibleBiomass = 0.1 * ((PGCGrowthAreas(iI) / baseTotalGrowthArea) * PlantInedibleBiomassGrowth(iI)) / 1000
        CarbonContentBiomass = CarbonContentBiomass + InedibleBiomass * PlantCarbonContent(iI)
    Next iI
        
        AlgeaCarbonContent = 0.48
        CarbonContentBiomass = CarbonContentBiomass + .Range("InedibleBiomassPBRValue").Value * AlgeaCarbonContent
    
    If CO2ToPlants * baseTotalGrowthArea > .Range("CO2_ProdCrewValue").Value Then
    
        If .OptionIncinerateInedible.Value = True Then
            Dim fRequiredCO2
            Dim fAvailableCarbon
            ' If inedible plant mass can be incinerated, incinerate the amount required to provide the missing CO2
            ' Base reaction assumed is C6H12O6 + 6 O2 -> 6 CO2 + 6 H2O
            ' We assume no water is produced from this but the required O2 must be provided
            fRequiredCO2 = CO2ToPlants * baseTotalGrowthArea - .Range("CO2_ProdCrewValue").Value
            fAvailableCarbon = CarbonContentBiomass * baseTotalGrowthArea
            
            ' if not sufficient O2 is available from the plants, the missing O2 is taken from storage / electrolysis!
            ' Therefore this calculation is not required here!
'            Dim fAvailableO2
'            Dim fMissingCO2
'            Dim fMissingO2
'            Dim fProduceableCO2
'            fAvailableO2 = O2FromPlants * baseTotalGrowthArea - .Range("O2_ConCrewValue").Value
'            fProduceableCO2 = (fAvailableO2 / MolarMassO2) * MolarMassCO2
'
'            fMissingCO2 = fRequiredCO2 - fProduceableCO2
'            ' take missing O2 from storage:
'            If fMissingCO2 > 0 Then
'                fMissingO2 = (fMissingCO2 / MolarMassCO2) * MolarMassO2
'            Else
'                fMissingO2 = 0
'            End
            
            Dim fProduceableCO2
            Dim fMissingCarbon
            fMissingCarbon = fRequiredCO2 / MolarMassCO2 - fAvailableCarbon / MolarMassC
            
            Do While fMissingCarbon > 0.00001
            ' in this case the inedible biomass does not contain the required carbon for CO2 production and the growth area is adjusted
            
                fProduceableCO2 = (fAvailableCarbon / MolarMassC) * MolarMassCO2
                TotalGrowthArea = (fProduceableCO2 + .Range("CO2_ProdCrewValue").Value) / CO2ToPlants
                
                fRequiredCO2 = CO2ToPlants * TotalGrowthArea - .Range("CO2_ProdCrewValue").Value
                fAvailableCarbon = CarbonContentBiomass * TotalGrowthArea
            
                fMissingCarbon = fRequiredCO2 / MolarMassCO2 - fAvailableCarbon / MolarMassC
            Loop
                
            .Range("O2_toIncinerationLabel").Font.color = .Range("O2Legend").Font.color
            .Range("O2_toIncinerationValue").Font.color = .Range("O2Legend").Font.color
            .Range("O2_toIncinerationValue").Value = (fRequiredCO2 / MolarMassCO2) * MolarMassO2
            .Range("CO2_FromWasteToCabinValue").Font.color = .Range("CO2Legend").Font.color
            .Range("CO2_FromWasteToCabinValue").Value = fRequiredCO2
            
            .Shapes("CO2_FromWasteToCabin1").Visible = True
            .Shapes("CO2_FromWasteToCabin2").Visible = True
            
        Else
            ' if the plant chamber would consume more CO2 than the crew produces, reduce the plant chamber area
            TotalGrowthArea = .Range("CO2_ProdCrewValue").Value / CO2ToPlants
            
            .Range("O2_toIncinerationLabel").Font.color = vbWhite
            .Range("O2_toIncinerationValue").Font.color = vbWhite
            .Range("O2_toIncinerationValue").Value = 0
            .Range("CO2_FromWasteToCabinValue").Font.color = vbWhite
            .Range("CO2_FromWasteToCabinValue").Value = 0
            .Shapes("CO2_FromWasteToCabin1").Visible = False
            .Shapes("CO2_FromWasteToCabin2").Visible = False
        End If
    Else
        TotalGrowthArea = baseTotalGrowthArea
        
        .Range("O2_toIncinerationLabel").Font.color = vbWhite
        .Range("O2_toIncinerationValue").Font.color = vbWhite
        .Range("O2_toIncinerationValue").Value = 0
        .Range("CO2_FromWasteToCabinValue").Font.color = vbWhite
        .Range("CO2_FromWasteToCabinValue").Value = 0
        .Shapes("CO2_FromWasteToCabin1").Visible = False
        .Shapes("CO2_FromWasteToCabin2").Visible = False
    End If
    
    For iI = 0 To UBound(PGCGrowthAreas)
        PGCGrowthAreas(iI) = (PGCGrowthAreas(iI) / baseTotalGrowthArea) * TotalGrowthArea
    Next iI
    
    ' store the total area in a cell on the sheet, to be used for the final ESM calculation of plant chambers
    .Range("CurrentPGCArea").Value = TotalGrowthArea
    
    Dim H2OtoPlants As Double
    Dim EdibleBiomass As Double
    Dim EdibleBiomassWaterContent As Double
    Dim EdibleBiomassEnergyContent As Double
    
    O2FromPlants = 0
    CO2ToPlants = 0
    H2OtoPlants = 0
    EdibleBiomass = 0
    InedibleBiomass = 0
    EdibleBiomassWaterContent = 0
    EdibleBiomassEnergyContent = 0
    For iI = 0 To UBound(PlantO2Production)
        O2FromPlants = O2FromPlants + (PGCGrowthAreas(iI) * PlantO2Production(iI)) / 1000  ' since BVAD values are in g divide with 1000 to get kg
        CO2ToPlants = CO2ToPlants + (PGCGrowthAreas(iI) * PlantCO2Consumption(iI)) / 1000
        H2OtoPlants = H2OtoPlants + (PGCGrowthAreas(iI) * PlantWaterUptake(iI)) ' only value that is provided in kg/m^2 and day
        EdibleBiomass = EdibleBiomass + (PGCGrowthAreas(iI) * PlantEdibleBiomassGrowth(iI)) / 1000
        InedibleBiomass = InedibleBiomass + (PGCGrowthAreas(iI) * PlantInedibleBiomassGrowth(iI)) / 1000
        EdibleBiomassWaterContent = EdibleBiomassWaterContent + (PlantEdibleBiomassWaterContent(iI) / 100) * (PGCGrowthAreas(iI) * PlantEdibleBiomassGrowth(iI)) / 1000
        EdibleBiomassEnergyContent = EdibleBiomassEnergyContent + PlantEnergyContent(iI) * (PGCGrowthAreas(iI) * PlantEdibleBiomassGrowth(iI)) / 1000
    Next iI
    .Range("O2fromPlantsValue").Value = O2FromPlants - .Range("O2_toIncinerationValue").Value
    .Range("CO2toPlantsValue").Value = CO2ToPlants
    .Range("EdibleFromPlantsValue").Value = EdibleBiomass
    .Range("Waste_FromPlantsValue").Value = InedibleBiomass
    .Range("EdibleFromPlantsEnergyValue").Value = EdibleBiomassEnergyContent
    .Range("EdibleFromPlantsWaterValue").Value = EdibleBiomassWaterContent
    .Range("PlantWaterContentToCrew").Value = EdibleBiomassWaterContent
    
    ' Nutrient usage is calculated based on the values from BVAD table 4.101, the nutrients in the stock solution are provided in the following table, which amounts to 1.166 g/L,
    ' and using the max stock consumption this results in 0.034 L * 1.166 g/l = 0.039 g / g. The acid is used for pH control, which means it is bascially additional water in the mass balance with 0.0618 g/g
    Dim NutrientConsumption As Double
    NutrientConsumption = (EdibleBiomass + InedibleBiomass) * (0.039 + 0.0618)
    ' The H2O output is calculated as the mass balance of the plants:
    .Range("H2OfromPlantsValue").Value = CO2ToPlants + H2OtoPlants + NutrientConsumption - O2FromPlants - EdibleBiomass - InedibleBiomass
    .Range("NutrientsToPlantsValue").Value = NutrientConsumption
    ' inedible biomass has 90% water content see BVAD
    .Range("H2O_InediblePlantsValue").Value = InedibleBiomass * 0.9
    
    If InStr(1, .WWFiltrationList.Value, "None", vbBinaryCompare) Then
        .Range("H2O_InediblePlantsToFiltrationValue").Value = 0
        .Range("H2O_InediblePlantsToFiltrationValue").Font.color = vbWhite
        .Shapes("H2O_InediblePlantsToFiltration").Visible = False
    Else
        ' only send water from inedible plant mass to filtration if some filtration option is used
        ' Assume that 80% of the total water content can be recovered. If no incineration is used, assume 80% is recovered
        If .OptionIncinerateInedible.Value = True Then
            .Range("H2O_InediblePlantsToFiltrationValue").Value = InedibleBiomass * 0.9 * 0.98
        Else
            .Range("H2O_InediblePlantsToFiltrationValue").Value = InedibleBiomass * 0.9 * 0.8
        End If
        
        .Range("H2O_InediblePlantsToFiltrationValue").Font.color = .Range("H2OLegend").Font.color
        .Shapes("H2O_InediblePlantsToFiltration").Visible = True
    End If
    
    ' Adjust the stored food values to match the energy content of the plants stored food is assumed to have 12 MJ / 1.51 kg of energy content
    .Range("FoodConCrewValue").Value = ((.Range("FoodConCrewValue").Value * (12000 / 1.51)) - EdibleBiomassEnergyContent) / (12000 / 1.51)
    ' Stored food water content is 0.7 kg/1.51 kg of food
    .Range("H2O_FoodConCrewValue").Value = .Range("FoodConCrewValue").Value * 0.7 / 1.51
    ' Now we have to adjust the potable water consumption to match the new food water intake
    ' First we check if the water intake would fall below 1 kg, if that is the case we cannot further reduce it:
    Dim fNewWaterConsumption As Double
    fNewWaterConsumption = .Range("H2O_ConCrewValue").Value - (.Range("H2O_FoodConCrewValue").Value + EdibleBiomassWaterContent - 0.7)
    If fNewWaterConsumption < 1 Then
        ' if more water is consumed than initially considered, we have to increase the urine production of the crew accordingly because the overall water
        ' intake becomes higher:
        Dim fAdditionalUrine As Double
        fAdditionalUrine = 1 - fNewWaterConsumption
        .Range("CrewUrineWithProcessing").Value = fAdditionalUrine
    
        fNewWaterConsumption = 1
    End If
    .Range("H2O_ConCrewValue").Value = fNewWaterConsumption
    
    O2_fromStorage = .Range("O2_ConCrewValue").Value - O2FromPlants + .Range("O2_LeakageValue").Value + .Range("O2_toIncinerationValue").Value
    CO2_toRemoval = .Range("CO2_ProdCrewValue").Value - (CO2ToPlants - .Range("CO2_FromWasteToCabinValue").Value)
    H2O_toHumidityControl = .Range("H2O_ProdCrewValue").Value + .Range("H2OfromPlantsValue").Value
    
End If

' set the O2 flow that must come from storage and the CO2 that must be removed by removal
O2_fromStorage = Round(O2_fromStorage, 5)
CO2_toRemoval = Round(CO2_toRemoval, 5)
.Range("O2_FromStorageValue").Value = O2_fromStorage
.Range("CO2_FromRemovalValue").Value = CO2_toRemoval

Dim miTechnologyRow
Dim Efficiency As Double

'-------------------------------------------------------------------------------------------------------------------------------------------------------------
' CO2 Removal:
If InStr(1, .CO2_RemovalList.Value, "LiOH", vbBinaryCompare) Then
    ' 2 LiOH + CO2 -> Li2CO3 + H2O
    
    .Range("CO2_RemovalToWasteValue1").Value = (CO2_toRemoval / MolarMassCO2) * MolarMassLi2CO3
    .Range("CO2_RemovalToWasteValue2").Value = (CO2_toRemoval / MolarMassCO2) * MolarMassLi2CO3
    
    .Range("H2O_FromCO2RemovalValue").Value = (CO2_toRemoval / MolarMassCO2) * MolarMassH2O
    
    .Range("CO2_FromRemovalValue").Value = 0
    
ElseIf InStr(1, .CO2_RemovalList.Value, "SAWD", vbBinaryCompare) Then
    ' Solid Amine Water desorption requires water to desorb the CO2, and that water is also released as humidity to the cabin
    ' See "Development Status of the Advanced Closed Loop System ACLS", Klaus Bockstahler et. al, 2015, Figure 4
    ' 3.7 kg of H2O per 3 kg of CO2 required
    
    If InStr(1, .HumidityControlList.Value, "2BMS", vbBinaryCompare) Or InStr(1, .HumidityControlList.Value, "CAMRAS", vbBinaryCompare) Then
        ' in this case use potable water, as humidity control does not produce liquid water
        ' note this should usually not occur, as both of these technologies must also be selected as co2 removal, but for future Gaseous
        ' humiditiy control systems
        
        .Shapes("H2O_ToCO2Removal").Visible = True
        .Range("H2O_ToCO2RemovalValue").Value = (CO2_toRemoval / 3) * 3.7
        .Range("H2O_ToCO2RemovalValue").Font.color = .Range("H2OLegend").Font.color
        
        .Shapes("H2O_FromHumControlToCO2Removal").Visible = False
        .Range("H2O_FromHumControlToCO2RemovalValue").Value = 0
        .Range("H2O_FromHumControlToCO2RemovalValue").Font.color = vbWhite
    Else
        .Shapes("H2O_ToCO2Removal").Visible = False
        .Range("H2O_ToCO2RemovalValue").Value = 0
        .Range("H2O_ToCO2RemovalValue").Font.color = vbWhite
        
        .Shapes("H2O_FromHumControlToCO2Removal").Visible = True
        .Range("H2O_FromHumControlToCO2RemovalValue").Value = (CO2_toRemoval / 3) * 3.7
        .Range("H2O_FromHumControlToCO2RemovalValue").Font.color = .Range("H2OLegend").Font.color
        
    End If
    
    .Range("H2O_FromCO2RemovalValue").Value = (CO2_toRemoval / 3) * 3.7
    
End If

H2O_toHumidityControl = H2O_toHumidityControl + .Range("H2O_FromCO2RemovalValue").Value - .Range("H2O_FromHumControlToCO2RemovalValue").Value

'-------------------------------------------------------------------------------------------------------------------------------------------------------------
' PBR:
If InStr(1, .PBRList.Value, "None", vbBinaryCompare) Then
    
    .Range("CO2ToPBRValue").Value = 0
    .Range("WaterToPBRValue").Value = 0
    .Range("WaterFromPBRValue").Value = 0
    .Range("O2FromPBRValue").Value = 0
    .Range("InedibleBiomassPBRValue").Value = 0
    .Range("FoodFromPBRValue").Value = 0
    .Range("NutrientsToPBRValue").Value = 0
    
   Else
   
    .Range("CO2ToPBRValue").Value = (.Range("CO2_FromRemovalValue").Value * PBR_CO2Percentage.Value * 0.01)
        ' PBR reaction is photosynthesis: 6 CO2 + 6 H2O -> 6 O2 + C6H12O6
    .Range("WaterToPBRValue").Value = .Range("CO2ToPBRValue").Value * (MolarMassH2O / MolarMassCO2) * 10
    .Range("WaterFromPBRValue").Value = .Range("CO2ToPBRValue").Value * (MolarMassH2O / MolarMassCO2) * 9
    .Range("O2FromPBRValue").Value = .Range("CO2ToPBRValue").Value * (MolarMassO2 / MolarMassCO2)
    
    .Range("InedibleBiomassPBRValue").Value = .Range("CO2ToPBRValue").Value * (0.1 / 1.85)
    .Range("FoodFromPBRValue").Value = .Range("CO2ToPBRValue").Value * (0.9 / 1.85)
    .Range("NutrientsToPBRValue").Value = .Range("CO2ToPBRValue").Value * (0.1 / 1.85)
    
End If

'
'H2O_toHumidityControl = H2O_toHumidityControl + .Range("H2O_FromCO2RemovalValue").Value - .Range("H2O_FromHumControlToCO2RemovalValue").Value



'-------------------------------------------------------------------------------------------------------------------------------------------------------------
' Urine Processing:

Dim fUrineWaterToProcessing As Double
Dim UrineSolidsToProc As Double
Dim fNutrientsReplaceableByUrine As Double

.Range("CrewUrineToPlantsValue").Font.color = .Range("GreyWaterLegend").Font.color
.Range("UrineToPlantsValue").Font.color = .Range("GreyWaterLegend").Font.color
.Range("UrineToPlantsLabel").Font.color = .Range("GreyWaterLegend").Font.color
.Shapes("CrewUrineToPlants").Visible = True


If .PGCurineTreatmentOption.Value = True Then

    If .Range("CrewUrineWithProcessing").Value / H2OtoPlants < 0.25 Then
        .Range("CrewUrineToPlantsValue").Value = .Range("CrewUrineWithProcessing").Value
        .Range("UrineToPlantsValue") = .Range("CrewUrineWithProcessing").Value
        .Range("CrewUrineWithProcessing").Value = 0
        
    Else
        .Range("CrewUrineWithProcessing").Value = .Range("CrewUrineWithProcessing").Value - 0.25 * H2OtoPlants
        .Range("CrewUrineToPlantsValue").Value = 0.25 * H2OtoPlants
        .Range("UrineToPlantsValue") = 0.25 * H2OtoPlants
    End If

    fUrineWaterToPlants = .Range("CrewUrineToPlantsValue").Value - ((.Range("CrewUrineToPlantsValue").Value / CrewUrineProduction) * UrineSolids)

    ' reduce plant nutrient demand
    fNutrientsReplaceableByUrine = 1 * .Range("CrewUrineToPlantsValue").Value
    If .Range("NutrientsToPlantsValue").Value < fNutrientsReplaceableByUrine Then
        .Range("NutrientsToPlantsValue").Value = 0
    Else
        .Range("NutrientsToPlantsValue").Value = .Range("NutrientsToPlantsValue").Value - 0.5 * .Range("CrewUrineToPlantsValue").Value
    End If
Else

fUrineWaterToPlants = 0
.Range("UrineToPlantsValue").Value = 0
.Range("CrewUrineToPlantsValue").Font.color = vbWhite
.Range("UrineToPlantsValue").Font.color = vbWhite
.Range("UrineToPlantsLabel").Font.color = vbWhite
.Shapes("CrewUrineToPlants").Visible = False

End If

UrineSolidsToProc = ((.Range("CrewUrineWithProcessing").Value / CrewUrineProduction) * UrineSolids)
        
If InStr(1, .UrineProcessingList.Value, "None", vbBinaryCompare) Then
    .Range("CrewUrineWithProcessing").Value = 0
Else
    miTechnologyRow = findRowByName(Array(.UrineProcessingList.Value), "Technology", miTechnologyColumn(1))
    Efficiency = Worksheets("Technology").Cells(miTechnologyRow(1), miTechnologyColumn(2)).Value
    
    fUrineWaterToProcessing = .Range("CrewUrineWithProcessing").Value - UrineSolidsToProc
    
    .Range("UrineProcessingGreyWater").Value = fUrineWaterToProcessing * Efficiency / 100
    .Range("BrineProdValue").Value = fUrineWaterToProcessing * (1 - Efficiency / 100) + UrineSolidsToProc
    
    
    If InStr(1, .WWFiltrationList.Value, "None", vbBinaryCompare) Then
        ' not a really useful case, but theoretically possible
        .Shapes("GreyWaterFromUrine").Visible = False
        .Shapes("GreyWaterFromUrineToWaste").Visible = True
        .Range("UrineProcessingGreyWaterToWasteValue").Value = .Range("UrineProcessingGreyWater").Value
        .Range("UrineProcessingGreyWaterToWasteValue").Font.color = .Range("GreyWaterLegend").Font.color
        .Range("UrineProcessingGreyWater").Value = 0
        .Range("UrineProcessingGreyWater").Font.color = vbWhite
        
    Else
        .Shapes("GreyWaterFromUrine").Visible = True
        .Shapes("GreyWaterFromUrineToWaste").Visible = False
        .Range("UrineProcessingGreyWaterToWasteValue").Value = 0
        .Range("UrineProcessingGreyWaterToWasteValue").Font.color = vbWhite
        .Range("UrineProcessingGreyWater").Font.color = .Range("GreyWaterLegend").Font.color
    End If
End If

'-------------------------------------------------------------------------------------------------------------------------------------------------------------
' Brine Processing:
If InStr(1, .BrineProcessingList.Value, "None", vbBinaryCompare) Then

Else
    miTechnologyRow = findRowByName(Array(.BrineProcessingList.Value), "Technology", miTechnologyColumn(1))
    Efficiency = Worksheets("Technology").Cells(miTechnologyRow(1), miTechnologyColumn(2)).Value
    
    .Range("BrinProcH2OtoCabinValue").Value = (.Range("BrineProdValue").Value - UrineSolidsToProc) * Efficiency / 100
    .Range("BrineProcSolidWasteValue").Value = (.Range("BrineProdValue").Value - UrineSolidsToProc) * (1 - Efficiency / 100) + UrineSolidsToProc
    
    H2O_toHumidityControl = H2O_toHumidityControl + .Range("BrinProcH2OtoCabinValue").Value
    
End If

'-------------------------------------------------------------------------------------------------------------------------------------------------------------
' Plant Humidity:
' must be calculated after CO2 removal, as that can produce additional water as humidity

If InStr(1, .PlantGrowthChamberList.Value, "None", vbBinaryCompare) Then
Else
    If InStr(1, .CO2_RemovalList.Value, "2BMS", vbBinaryCompare) Or InStr(1, .CO2_RemovalList.Value, "CAMRAS", vbBinaryCompare) Then
        
        .Shapes("HumControlToWaste").Visible = False
        .Shapes("HumControlToWWFiltration").Visible = False
        .Shapes("HumControlToPlants").Visible = False
        .Range("HumControlToPlantsValue").Font.color = vbWhite
        .Range("HumControlToPlantsValue").Value = 0
        .Shapes("H2O_ConPlants").Visible = True
        .Range("H2OtoPlantsValue").Font.color = .Range("H2OLegend").Font.color
        .Range("H2OtoPlantsValue").Value = H2OtoPlants
    Else
        If H2O_toHumidityControl < H2OtoPlants Then
            ' in this case all of humidity control water goes directly to the plants, so we hide the original path, set the value for this to a number but the color to white, as it would double itself if we put it here
            .Shapes("HumControlToWaste").Visible = False
            .Shapes("HumControlToWWFiltration").Visible = False
            
            .Shapes("HumControlToPlants").Visible = True
            .Range("HumControlToPlantsValue").Font.color = vbWhite
            .Range("HumControlToPlantsValue").Value = H2O_toHumidityControl
            
            .Shapes("H2O_ConPlants").Visible = True
            .Range("H2OtoPlantsValue").Font.color = .Range("H2OLegend").Font.color
            .Range("H2OtoPlantsValue").Value = H2OtoPlants - H2O_toHumidityControl
        Else
            If InStr(1, .WWFiltrationList.Value, "None", vbBinaryCompare) Then
                .Shapes("HumControlToWaste").Visible = True
                .Shapes("HumControlToWWFiltration").Visible = False
            Else
                .Shapes("HumControlToWaste").Visible = False
                .Shapes("HumControlToWWFiltration").Visible = True
            End If
            .Shapes("HumControlToPlants").Visible = True
            .Range("HumControlToPlantsValue").Font.color = .Range("H2OLegend").Font.color
            .Range("HumControlToPlantsValue").Value = H2OtoPlants
            
            .Shapes("H2O_ConPlants").Visible = False
            .Range("H2OtoPlantsValue").Font.color = vbWhite
            .Range("H2OtoPlantsValue").Value = 0
        
        End If
    End If
    ' Reduce plant water demand if urine is used:
    If .Range("H2OtoPlantsValue").Value > fUrineWaterToPlants Then
        .Range("H2OtoPlantsValue").Value = .Range("H2OtoPlantsValue").Value - fUrineWaterToPlants
    Else
        .Range("HumControlToPlantsValue").Value = .Range("HumControlToPlantsValue").Value - (fUrineWaterToPlants - .Range("H2OtoPlantsValue").Value)
    End If
    
' This code could be used to assume veggie or the ISPR pgc can only work with potable water
'    If InStr(1, .PlantGrowthChamberList.Value, "Veggie", vbBinaryCompare) Or InStr(1, .PlantGrowthChamberList.Value, "ISPR PGC", vbBinaryCompare) _
'       Or InStr(1, .HumidityControlList.Value, "2BMS", vbBinaryCompare) Or InStr(1, .HumidityControlList.Value, "CAMRAS", vbBinaryCompare) Then
'
'        If InStr(1, .WWFiltrationList.Value, "None", vbBinaryCompare) Then
'            .Shapes("WWFiltrationToWaste").Visible = True
'            .Shapes("HumControlToWWFiltration").Visible = False
'        Else
'            .Shapes("WWFiltrationToWaste").Visible = False
'            .Shapes("HumControlToWWFiltration").Visible = True
'        End If
'
'        .Shapes("HumControlToPlants").Visible = False
'        .Range("HumControlToPlantsValue").Font.color = vbWhite
'        .Range("HumControlToPlantsValue").Value = 0
'
'        .Shapes("H2O_ConPlants").Visible = True
'        .Range("H2OtoPlantsValue").Font.color = .Range("H2OLegend").Font.color
'        .Range("H2OtoPlantsValue").Value = H2OtoPlants
'    Else
'    End If
End If



'-------------------------------------------------------------------------------------------------------------------------------------------------------------
' humidity control
If InStr(1, .HumidityControlList.Value, "CAMRAS", vbBinaryCompare) Or InStr(1, .HumidityControlList.Value, "2BMS", vbBinaryCompare) Then
    ' For this case humidity is vented:
    .Range("HumControlToVentValue").Value = H2O_toHumidityControl
    
    .Shapes("HumControlToWaste").Visible = False
    .Shapes("HumControlToWWFiltration").Visible = False
    .Range("HumControlToWasteValue").Font.color = vbWhite
    .Range("HumControlToWWFiltrationValue").Font.color = vbWhite
    .Range("HumControlToWasteValue").Value = 0
    .Range("HumControlToWWFiltrationValue").Value = 0
    
Else
    ' For this case humidity is either filtered back to potable or put into waste storage
    If InStr(1, .WWFiltrationList.Value, "None", vbBinaryCompare) Then
        .Range("HumControlToWasteValue").Value = H2O_toHumidityControl
    Else
        .Range("HumControlToWWFiltrationValue").Value = H2O_toHumidityControl
    End If
End If

'-------------------------------------------------------------------------------------------------------------------------------------------------------------
' electrolysis
Dim O2_fromELY
If InStr(1, .ElectrolysisList.Value, "None", vbBinaryCompare) Then

Else
If InStr(1, .PBRList.Value, "None", vbBinaryCompare) Then
    O2_fromELY = O2_fromStorage - .Range("O2_FromCO2ReductionValue").Value
Else
    O2_fromELY = O2_fromStorage - .Range("O2_FromCO2ReductionValue").Value - .Range("O2FromPBRValue").Value
End If
    .Range("O2_FromElectrolysisValue").Value = O2_fromELY
    ' 2 H2O -> 2 H2 + O2
    .Range("H2_FromElectrolysisValue").Value = (O2_fromELY / MolarMassO2) * MolarMassH2 * 2
    .Range("H2O_ToElectrolysisValue").Value = (O2_fromELY / MolarMassO2) * MolarMassH2O * 2
End If


'-------------------------------------------------------------------------------------------------------------------------------------------------------------
' CO2 Reduction
Dim H2fromElectrolysis As Double
Dim CO2fromRemoval As Double
Dim H2toVent As Double
Dim CO2toVent As Double

H2fromElectrolysis = .Range("H2_FromElectrolysisValue").Value

If InStr(1, .PBRList.Value, "None", vbBinaryCompare) Then
CO2fromRemoval = .Range("CO2_FromRemovalValue").Value
Else
CO2fromRemoval = .Range("CO2_FromRemovalValue").Value - .Range("CO2ToPBRValue").Value
End If

If InStr(1, .CO2ReductionList.Value, "None", vbBinaryCompare) Then
    ' in this case all the CO2 and H2 must be vented
    H2toVent = H2fromElectrolysis
    CO2toVent = CO2fromRemoval
Else

    ' get efficiency of CO2 conversion:
    miTechnologyRow = findRowByName(Array(.CO2ReductionList.Value), "Technology", miTechnologyColumn(1))
    Efficiency = Worksheets("Technology").Cells(miTechnologyRow(1), miTechnologyColumn(2)).Value

    .Range("O2_FromCO2ReductionValue").Value = 0
    
    If InStr(1, .CO2ReductionList.Value, "SR/SCRA", vbBinaryCompare) Then
        If .MethaneProcessingList.Value = "None" Then
        ' Sabatier reaction: CO2 + 4 H2 -> CH4 + 2 H2O
            If (H2fromElectrolysis / MolarMassH2) / 4 > (CO2fromRemoval / MolarMassCO2) Or .H2StorageList.Value <> "None" Then
            ' too much hydrogen, CO2 is limiting
            .Range("CO2_ToReductionValue").Value = CO2fromRemoval
            .Range("H2O_FromCO2ReductionValue").Value = 2 * (CO2fromRemoval / MolarMassCO2) * MolarMassH2O * Efficiency / 100
            .Range("CO2_FromReductionToVentValue").Value = (CO2fromRemoval / MolarMassCO2) * MolarMassCH4 * Efficiency / 100
            .Range("H2_ToCO2ReductionValue").Value = 4 * (CO2fromRemoval / MolarMassCO2) * MolarMassH2
            ' add the not converted masses to the vent flows:
            .Range("CO2_FromReductionToVentValue").Value = .Range("CO2_FromReductionToVentValue").Value + (CO2fromRemoval + .Range("H2_ToCO2ReductionValue").Value) * (1 - (Efficiency / 100))
            
            CO2toVent = 0
            H2toVent = H2fromElectrolysis - .Range("H2_ToCO2ReductionValue").Value
            
            Else
            ' too much CO2, H2 is limiting
            .Range("H2_ToCO2ReductionValue").Value = H2fromElectrolysis
            .Range("H2O_FromCO2ReductionValue").Value = 2 / 4 * (H2fromElectrolysis / MolarMassH2) * MolarMassH2O * Efficiency / 100
            .Range("CO2_FromReductionToVentValue").Value = 1 / 4 * (H2fromElectrolysis / MolarMassH2) * MolarMassCH4 * Efficiency / 100
            .Range("CO2_ToReductionValue").Value = 1 / 4 * (H2fromElectrolysis / MolarMassH2) * MolarMassCO2
            ' add the not converted masses to the vent flows:
            .Range("CO2_FromReductionToVentValue").Value = .Range("CO2_FromReductionToVentValue").Value + (H2fromElectrolysis + .Range("CO2_ToReductionValue").Value) * (1 - (Efficiency / 100))
            
            CO2toVent = CO2fromRemoval - .Range("CO2_ToReductionValue").Value
            H2toVent = 0
        End If
        ElseIf .MethaneProcessingList.Value <> "None" Then
            If (H2fromElectrolysis / MolarMassH2) / 4 > (CO2fromRemoval / MolarMassCO2) Then
            ' too much hydrogen, CO2 is limiting
            MsgBox ("Already enough H2, CH4 Processing deactivated")
            .MethaneProcessingList.Value = "None"
            .Range("CO2_ToReductionValue").Value = CO2fromRemoval
            .Range("H2O_FromCO2ReductionValue").Value = 2 * (CO2fromRemoval / MolarMassCO2) * MolarMassH2O * Efficiency / 100
            .Range("CO2_FromReductionToVentValue").Value = (CO2fromRemoval / MolarMassCO2) * MolarMassCH4 * Efficiency / 100
            .Range("H2_ToCO2ReductionValue").Value = 4 * (CO2fromRemoval / MolarMassCO2) * MolarMassH2
            ' add the not converted masses to the vent flows:
            .Range("CO2_FromReductionToVentValue").Value = .Range("CO2_FromReductionToVentValue").Value + (CO2fromRemoval + .Range("H2_ToCO2ReductionValue").Value) * (1 - (Efficiency / 100))
            CO2toVent = 0
            H2toVent = H2fromElectrolysis - .Range("H2_ToCO2ReductionValue").Value
            
            Else
            If .MethaneProcessingList.Value = "LPU" Then
                'H2 Loop closed, CH4 vented
                .Range("CO2_ToReductionValue").Value = CO2fromRemoval
                .Range("H2O_FromCO2ReductionValue").Value = 2 * (CO2fromRemoval / MolarMassCO2) * MolarMassH2O * Efficiency / 100
                .Range("H2_ToCO2ReductionValue").Value = 4 * (CO2fromRemoval / MolarMassCO2) * MolarMassH2
                .Range("H2_FromMethaneProcessingValue").Value = .Range("H2_ToCO2ReductionValue").Value - .Range("H2_FromElectrolysisValue").Value
                .Range("CH4_FromReductionToProcessingValue").Value = 0.5 * (.Range("H2_FromMethaneProcessingValue").Value / MolarMassH2) * MolarMassCH4
                .Range("CO2_FromReductionToVentValue").Value = (CO2fromRemoval / MolarMassCO2) * MolarMassCH4 * Efficiency / 100 - .Range("CH4_FromReductionToProcessingValue")
                .Range("FromMethaneProcessingToWasteValue").Value = (.Range("CH4_FromReductionToProcessingValue").Value / MolarMassCH4) * MolarMassC
                        CO2toVent = 0
                        H2toVent = 0
                
            ElseIf .MethaneProcessingList.Value = "PPU" Then
                'H2 Loop not closed, H2 from Storage
                .Range("CO2_ToReductionValue").Value = CO2fromRemoval
                .Range("H2O_FromCO2ReductionValue").Value = 2 * (CO2fromRemoval / MolarMassCO2) * MolarMassH2O * Efficiency / 100
                .Range("H2_ToCO2ReductionValue").Value = 4 * (CO2fromRemoval / MolarMassCO2) * MolarMassH2
                .Range("CH4_FromReductionToProcessingValue").Value = (CO2fromRemoval / MolarMassCO2) * MolarMassCH4 * Efficiency / 100
                .Range("H2_FromMethaneProcessingValue").Value = 1.5 * (.Range("CH4_FromReductionToProcessingValue").Value / MolarMassCH4) * MolarMassH2
                .Range("CO2_FromReductionToVentValue").Value = 0
                .Range("FromMethaneProcessingToWasteValue").Value = 0
                .Range("FromMethaneProcessingToVentValue").Value = 0.5 * (.Range("CH4_FromReductionToProcessingValue").Value / MolarMassCH4) * MolarMassC2H2
                        CO2toVent = 0
                        H2toVent = 0
            End If
        End If
            
    End If
        
        
        
        
        
        
    ElseIf InStr(1, .CO2ReductionList.Value, "BR", vbBinaryCompare) Then
        Dim TotalInFlow As Double
        
        ' Bosch reaction: CO2 + 2 H2 -> C + 2 H2O
        ' Check which reactant is abundant
        If (H2fromElectrolysis / MolarMassH2) / 2 > (CO2fromRemoval / MolarMassCO2) Or .H2StorageList.Value <> "None" Then
            ' too much hydrogen, CO2 is limiting
            .Range("CO2_ToReductionValue").Value = CO2fromRemoval
            .Range("H2O_FromCO2ReductionValue").Value = 2 * (CO2fromRemoval / MolarMassCO2) * MolarMassH2O * Efficiency / 100
            .Range("CO2_FromReductionToWasteValue").Value = (CO2fromRemoval / MolarMassCO2) * MolarMassC * Efficiency / 100
            .Range("H2_ToCO2ReductionValue").Value = 2 * (CO2fromRemoval / MolarMassCO2) * MolarMassH2
            
            TotalInFlow = (CO2fromRemoval + .Range("H2_ToCO2ReductionValue").Value)
            CO2toVent = 0
            H2toVent = H2fromElectrolysis - .Range("H2_ToCO2ReductionValue").Value
        Else
            ' too much CO2, H2 is limiting
            .Range("H2_ToCO2ReductionValue").Value = H2fromElectrolysis
            .Range("H2O_FromCO2ReductionValue").Value = 1 * (H2fromElectrolysis / MolarMassH2) * MolarMassH2O * Efficiency / 100
            .Range("CO2_FromReductionToWasteValue").Value = 1 * (H2fromElectrolysis / MolarMassH2) * MolarMassC * Efficiency / 100
            .Range("CO2_ToReductionValue").Value = 1 * (H2fromElectrolysis / MolarMassH2) * MolarMassCO2
            
            TotalInFlow = .Range("CO2_ToReductionValue").Value + H2fromElectrolysis
            CO2toVent = CO2fromRemoval - .Range("CO2_ToReductionValue").Value
            H2toVent = 0
        End If
        
        ' add the not converted masses to the vent flows:
        If Efficiency <> 100 Then
            .Shapes("CO2_FromReductionToVent").Visible = True
            .Range("CO2_FromReductionToVentLabel").Font.color = .Range("MiscLegend").Font.color
            .Range("CO2_FromReductionToVentValue").Font.color = .Range("MiscLegend").Font.color
            .Range("CO2_FromReductionToVentLabel").Value = "H2 + CO2"
            .Range("CO2_FromReductionToVentValue").Value = TotalInFlow * (1 - (Efficiency / 100))
        Else
            .Shapes("CO2_FromReductionToVent").Visible = False
            .Range("CO2_FromReductionToVentLabel").Font.color = vbWhite
            .Range("CO2_FromReductionToVentValue").Font.color = vbWhite
            .Range("CO2_FromReductionToVentValue").Value = 0
        End If
    End If
    
End If


H2toVent = Round(H2toVent, 5)
CO2toVent = Round(CO2toVent, 5)

If H2toVent > 0 Then
    .Shapes("H2_ToVent").Visible = True
    .Range("H2_ToVentValue").Value = H2toVent
    .Range("H2_ToVentValue").Font.color = .Range("H2Legend").Font.color
Else
    .Shapes("H2_ToVent").Visible = False
    .Range("H2_ToVentValue").Value = 0
    .Range("H2_ToVentValue").Font.color = vbWhite
End If
If CO2toVent > 0 Then
    .Shapes("CO2ToVent").Visible = True
    .Range("CO2_ToVentValue").Value = CO2toVent
    .Range("CO2_ToVentValue").Font.color = .Range("CO2Legend").Font.color
Else
    .Shapes("CO2ToVent").Visible = False
    .Range("CO2_ToVentValue").Value = 0
    .Range("CO2_ToVentValue").Font.color = vbWhite
End If




'-------------------------------------------------------------------------------------------------------------------------------------------------------------
' Waste Water Filtration
If InStr(1, .WWFiltrationList.Value, "None", vbBinaryCompare) Then

Else
    miTechnologyRow = findRowByName(Array(.WWFiltrationList.Value), "Technology", miTechnologyColumn(1))
    Efficiency = Worksheets("Technology").Cells(miTechnologyRow(1), miTechnologyColumn(2)).Value
    Dim fTotalWaterInputWWFiltration As Double
    
    fTotalWaterInputWWFiltration = .Range("HumControlToWWFiltrationValue").Value - .Range("HumControlToPlantsValue").Value + .Range("UrineProcessingGreyWater").Value + .Range("HygieneH2O_ProdCrewValue").Value + .Range("H2O_InediblePlantsToFiltrationValue").Value
    
    .Range("H2O_FromWWFiltValue").Value = fTotalWaterInputWWFiltration * Efficiency / 100
    
    If Efficiency <> 100 Then
        .Shapes("WWFiltrationToWaste").Visible = True
        .Range("WWFiltrationToWasteValue").Value = fTotalWaterInputWWFiltration * (1 - (Efficiency / 100))
        .Range("WWFiltrationToWasteValue").Font.color = .Range("MiscLegend").Font.color
        .Range("WWFiltrationToWasteLabel").Font.color = .Range("MiscLegend").Font.color
    Else
        .Shapes("WWFiltrationToWaste").Visible = False
        .Range("WWFiltrationToWasteValue").Value = 0
        .Range("WWFiltrationToWasteValue").Font.color = vbWhite
        .Range("WWFiltrationToWasteLabel").Font.color = vbWhite
    End If
    
End If

' Add water balance value to sheet:
Dim H2O_MassBalance As Double
H2O_MassBalance = .Range("H2O_FromWWFiltValue").Value + .Range("H2O_FromCO2ReductionValue").Value - .Range("WaterToPBRValue").Value + .Range("WaterFromPBRValue").Value _
                            - .Range("H2O_ToElectrolysisValue").Value - .Range("H2O_ToCO2ReductionValue").Value - .Range("HygieneH2O_ConCrewValue").Value - .Range("H2O_ConCrewValue").Value - .Range("H2OtoPlantsValue").Value
                     
H2O_MassBalance = Round(H2O_MassBalance, 2)

.WaterBalance.Caption = "Water Balance: " & CStr(H2O_MassBalance)

' now perform the dynamic analysis values for this config (mostly the values are independent of config
DynamicAnalysis

End With
End Sub


Sub DynamicAnalysis()

Dim CrewSize As Integer
Dim Schedule As Integer
Dim PressurizedVolume As Double

CrewSize = Worksheets("User Interface").Range("Crew_Size").Value
PressurizedVolume = Worksheets("User Interface").Range("Pressurized_volume").Value
Schedule = Worksheets("User Interface").Range("Schedule_Number").Value


Dim O2_Con As Double 'O2 Consumption in [kg/d]
Dim CO2_Pro As Double 'CO2 Production [kg/d]
Dim urine As Double 'urine Production in [kg/d]
Dim potable As Double 'potable water Need in [kg/d]
Dim hygiene As Double 'hygiene water Need in [kg/d]
Dim feces As Double 'feces Production in [kg/d]
Dim sweat As Double 'sweat and vpr Production in [kg/d]
Dim peak_sweat As Double 'sweat and vpr Production in [kg/h]
Dim heat As Double 'heat Production in [kJ/d]
Dim peak_heat As Double 'heat Production in [kJ/h]
Dim food As Double 'food consumption in [kg/d]
Dim H2O_Con As Double 'Total amount of H2O consumed (potable+hygiene+rehydration)
Dim H2O_Pro As Double 'Total amount of Waste water produced (urine+hygiene+sweat)

Dim iScheduleHeaderRow As Integer
Dim iScheduleRows As Integer
iScheduleRows = 48

If Schedule = 1 Then

    O2_Con = Worksheets("Schedule and Metabolic").Range("SUM_O2_Consumption").Value
    CO2_Pro = Worksheets("Schedule and Metabolic").Range("SUM_CO2_Production").Value
    urine = Worksheets("Schedule and Metabolic").Range("SUM_Urine_Production").Value
    hygiene = Worksheets("Schedule and Metabolic").Range("SUM_Hygiene_Water_Need").Value
    feces = Worksheets("Schedule and Metabolic").Range("SUM_Feces_Production").Value
    sweat = Worksheets("Schedule and Metabolic").Range("SUM_Sweat_Production").Value
    peak_sweat = WorksheetFunction.Max(Worksheets("Schedule and Metabolic").Range("sweat").Value)
    heat = Worksheets("Schedule and Metabolic").Range("SUM_Heat_Load").Value
    peak_heat = WorksheetFunction.Max(Worksheets("Schedule and Metabolic").Range("heat").Value)
    potable = Worksheets("Schedule and Metabolic").Range("SUM_Potable_Water_Need").Value
    food = Worksheets("Schedule and Metabolic").Range("SUM_Food_Need").Value

    iScheduleHeaderRow = 2

ElseIf Schedule = 2 Then

    O2_Con = Worksheets("Schedule and Metabolic").Range("SUM_O2_Consumption_2").Value
    CO2_Pro = Worksheets("Schedule and Metabolic").Range("SUM_CO2_Production_2").Value
    urine = Worksheets("Schedule and Metabolic").Range("SUM_Urine_Production_2").Value
    hygiene = Worksheets("Schedule and Metabolic").Range("SUM_Hygiene_Water_Need_2").Value
    feces = Worksheets("Schedule and Metabolic").Range("SUM_Feces_Production_2").Value
    sweat = Worksheets("Schedule and Metabolic").Range("SUM_Sweat_Production_2").Value
    peak_sweat = WorksheetFunction.Max(Worksheets("Schedule and Metabolic").Range("Sweat_Schedule2").Value)
    heat = Worksheets("Schedule and Metabolic").Range("SUM_Heat_Load_2").Value
    peak_heat = WorksheetFunction.Max(Worksheets("Schedule and Metabolic").Range("Heat_Schedule2").Value)
    potable = Worksheets("Schedule and Metabolic").Range("SUM_Potable_Water_Need_2").Value
    food = Worksheets("Schedule and Metabolic").Range("SUM_Food_Need_2").Value
    
    'Check if user defined schedule is up to date
    If Worksheets("Schedule and Metabolic").Range("Metabolic_Values_up_to_date").Value = 0 Then
        MsgBox ("User defined schedule is not up to date! Adjust schedule before proceeding!")
        Exit Sub
    End If
     
    iScheduleHeaderRow = 128
End If
H2O_Con = (potable + hygiene)
H2O_Pro = (hygiene + sweat + urine)

Dim HumidityFromECLSS As Double

With ThisWorkbook.Worksheets("ECLSS Composition")
    HumidityFromECLSS = (.Range("H2OfromPlantsValue").Value + .Range("BrinProcH2OtoCabinValue").Value) / 24
End With


Dim miScheduleColumns
Dim miOutputColumns
miScheduleColumns = findColumnByName(Array("O2", "CO2", "urine", "feces", "potable", "hygiene", "food", "sweat", "heat"), "Schedule and Metabolic", iScheduleHeaderRow)

Dim CurrentCO2 As Double
Dim CurrentO2 As Double
Dim CurrentHumidity As Double
Dim CurrentTemperature As Double

CurrentCO2 = 0.4 ' kPA
CurrentO2 = 21 ' kPa
CurrentHumidity = 0.5
CurrentTemperature = 293 ' K

Dim HeatCapacityAir As Double
Dim CabinMassAir As Double
Dim AntoineA As Double
Dim AntoineB As Double
Dim AntoineC As Double


HeatCapacityAir = 1005 ' kJ/kg K

CabinMassAir = (10 ^ 5 * PressurizedVolume) / ((IdealGasConst / MolarMassAir) * CurrentTemperature)

' Antoine equation from nist chemistry webbook using the values between 255.9 and 373 K
AntoineA = 4.6543
AntoineB = 1435.264
AntoineC = -64.848

Dim CurrentCO2Mass As Double
Dim CurrentO2Mass As Double
Dim CurrentHumidityMass As Double
Dim CurrentVaporPressure As Double
Dim CurrentHeat As Double

CurrentCO2Mass = (CurrentCO2 * 1000 * PressurizedVolume) / ((IdealGasConst / MolarMassCO2) * CurrentTemperature)
CurrentO2Mass = (CurrentO2 * 1000 * PressurizedVolume) / ((IdealGasConst / MolarMassO2) * CurrentTemperature)

CurrentVaporPressure = 10 ^ (AntoineA - (AntoineB / (CurrentTemperature + AntoineC)))
CurrentHumidityMass = (CurrentHumidity * CurrentVaporPressure * 1000 * PressurizedVolume) / ((IdealGasConst / MolarMassH2O) * CurrentTemperature)

CurrentHeat = CabinMassAir * HeatCapacityAir * CurrentTemperature

Dim iOutputHeaderRow As Integer
Dim i As Integer
iOutputHeaderRow = 5

miOutputColumns = findColumnByName(Array("CO2_Prod", "CO2_pp", "O2_Con", "O2_pp", "humidity", "urine", "WW", "Hygiene", "Potable", "Rel. Hum.", "Heat", "Temp."), "Output Data", iOutputHeaderRow)

iOutputHeaderRow = iOutputHeaderRow + 1

For i = 1 To iScheduleRows
    ' Add kg/h values
    Worksheets("Output Data").Cells(iOutputHeaderRow + i, miOutputColumns(1)).Value = Worksheets("Schedule and Metabolic").Cells(iScheduleHeaderRow + i, miScheduleColumns(2)).Value ' CO2
    Worksheets("Output Data").Cells(iOutputHeaderRow + i, miOutputColumns(3)).Value = Worksheets("Schedule and Metabolic").Cells(iScheduleHeaderRow + i, miScheduleColumns(1)).Value ' O2
    Worksheets("Output Data").Cells(iOutputHeaderRow + i, miOutputColumns(6)).Value = Worksheets("Schedule and Metabolic").Cells(iScheduleHeaderRow + i, miScheduleColumns(3)).Value ' Urine
    Worksheets("Output Data").Cells(iOutputHeaderRow + i, miOutputColumns(8)).Value = Worksheets("Schedule and Metabolic").Cells(iScheduleHeaderRow + i, miScheduleColumns(6)).Value ' Hygiene
    Worksheets("Output Data").Cells(iOutputHeaderRow + i, miOutputColumns(9)).Value = Worksheets("Schedule and Metabolic").Cells(iScheduleHeaderRow + i, miScheduleColumns(5)).Value ' Potable
    Worksheets("Output Data").Cells(iOutputHeaderRow + i, miOutputColumns(7)).Value = Worksheets("Schedule and Metabolic").Cells(iScheduleHeaderRow + i, miScheduleColumns(6)).Value + _
                                                                                      Worksheets("Schedule and Metabolic").Cells(iScheduleHeaderRow + i, miScheduleColumns(8)).Value + _
                                                                                      HumidityFromECLSS ' Waste Water
    Worksheets("Output Data").Cells(iOutputHeaderRow + i, miOutputColumns(5)).Value = Worksheets("Schedule and Metabolic").Cells(iScheduleHeaderRow + i, miScheduleColumns(8)).Value + _
                                                                                      HumidityFromECLSS ' Humidity
    Worksheets("Output Data").Cells(iOutputHeaderRow + i, miOutputColumns(11)).Value = Worksheets("Schedule and Metabolic").Cells(iScheduleHeaderRow + i, miScheduleColumns(9)).Value ' Heat
    
Next i

Dim MeanCO2 As Double
Dim MeanO2 As Double
Dim MeanHumidity As Double
Dim MeanHeat As Double

With Worksheets("Output Data")

    MeanCO2 = WorksheetFunction.Average(.Range(.Cells(iOutputHeaderRow + 1, miOutputColumns(1)), .Cells(iOutputHeaderRow + iScheduleRows, miOutputColumns(1))))
    MeanO2 = WorksheetFunction.Average(.Range(.Cells(iOutputHeaderRow + 1, miOutputColumns(3)), .Cells(iOutputHeaderRow + iScheduleRows, miOutputColumns(3))))
    MeanHumidity = WorksheetFunction.Average(.Range(.Cells(iOutputHeaderRow + 1, miOutputColumns(5)), .Cells(iOutputHeaderRow + iScheduleRows, miOutputColumns(5))))
    MeanHeat = WorksheetFunction.Average(.Range(.Cells(iOutputHeaderRow + 1, miOutputColumns(11)), .Cells(iOutputHeaderRow + iScheduleRows, miOutputColumns(11))))

Dim CurrentCO2Change As Double
Dim CurrentO2Change As Double
Dim CurrentHumidityChange As Double
Dim CurrentHeatChange As Double
Dim CurrentMinimumHumidityMass As Double
Dim CurrentMaximumHumidityMass As Double

For i = 1 To iScheduleRows

    If i > 1 Then
        CurrentCO2Change = .Cells(iOutputHeaderRow + i - 1, miOutputColumns(1)).Value
        CurrentO2Change = .Cells(iOutputHeaderRow + i - 1, miOutputColumns(3)).Value
        CurrentHumidityChange = (.Cells(iOutputHeaderRow + i - 1, miOutputColumns(5)).Value - MeanHumidity) * 0.5
        CurrentHeatChange = .Cells(iOutputHeaderRow + i - 1, miOutputColumns(11)).Value
        
        CurrentHeat = CurrentHeat + (CurrentHeatChange - MeanHeat) * 0.5
        CurrentTemperature = CurrentHeat / (CabinMassAir * HeatCapacityAir)
        
        CurrentCO2Mass = CurrentCO2Mass + (CurrentCO2Change - MeanCO2) * 0.5
        CurrentO2Mass = CurrentO2Mass + (CurrentO2Change - MeanO2) * 0.5
        
        CurrentMinimumHumidityMass = (0.25 * CurrentVaporPressure * 1000 * PressurizedVolume) / ((IdealGasConst / MolarMassH2O) * CurrentTemperature)
        CurrentMaximumHumidityMass = (0.75 * CurrentVaporPressure * 1000 * PressurizedVolume) / ((IdealGasConst / MolarMassH2O) * CurrentTemperature)
        If (CurrentHumidityMass + CurrentHumidityChange) < 0 Then
            ' Minimum humidity is 0.25
            CurrentHumidityMass = CurrentMinimumHumidityMass
        Else
            CurrentHumidityMass = CurrentHumidityMass + CurrentHumidityChange
            If CurrentHumidityMass > CurrentMaximumHumidityMass Then
                CurrentHumidityMass = CurrentMaximumHumidityMass
            End If
        End If
        
        CurrentCO2 = (CurrentCO2Mass * ((IdealGasConst / MolarMassCO2) * CurrentTemperature)) / PressurizedVolume / 1000
        CurrentO2 = (CurrentO2Mass * ((IdealGasConst / MolarMassO2) * CurrentTemperature)) / PressurizedVolume / 1000
        
        CurrentVaporPressure = 10 ^ (AntoineA - (AntoineB / (CurrentTemperature + AntoineC)))
        CurrentHumidity = (((CurrentHumidityMass * ((IdealGasConst / MolarMassH2O) * CurrentTemperature)) / PressurizedVolume) / 1000) / CurrentVaporPressure
    End If

    ' Calculate dynamic cabin values
    .Cells(iOutputHeaderRow + i, miOutputColumns(2)).Value = CurrentCO2 ' CO2 PP
    .Cells(iOutputHeaderRow + i, miOutputColumns(4)).Value = CurrentO2 ' O2 PP
    .Cells(iOutputHeaderRow + i, miOutputColumns(10)).Value = CurrentHumidity ' Relative Humidity
    .Cells(iOutputHeaderRow + i, miOutputColumns(12)).Value = CurrentTemperature ' temperature

Next i

End With

End Sub

Sub AddESMResultsToPosition()
    
' just to be sure, calculate the ECLSS before storing values
    CalculateECLSS
    AddCompositionResults.Width = 300
    AddCompositionResults.Height = 250
    AddCompositionResults.Show
    
End Sub

Private Sub AdditionalSystemConfigButton_Click()

    AdditionalSystemConfiguration.Width = 500
    AdditionalSystemConfiguration.Height = 550
    AdditionalSystemConfiguration.Show
    
End Sub

Private Sub BrineProcessingList_Change()

With ThisWorkbook.Worksheets("ECLSS Composition")
If InStr(1, .BrineProcessingList.Value, "None", vbBinaryCompare) Then

    .Shapes("BrineProcToCabin").Visible = False
    .Range("BrinProcH2OtoCabinValue").Font.color = vbWhite
    .Range("BrinProcH2OtoCabinValue").Value = 0
    
    .Shapes("BrineSolidWaste").Visible = False
    .Range("BrineProcSolidWasteValue").Font.color = vbWhite
    .Range("BrineProcSolidWasteLabel").Font.color = vbWhite
    .Range("BrineProcSolidWasteValue").Value = 0

Else

    .Shapes("BrineProcToCabin").Visible = True
    .Range("BrinProcH2OtoCabinValue").Font.color = .Range("H2OLegend").Font.color
    
    .Shapes("BrineSolidWaste").Visible = True
    .Range("BrineProcSolidWasteValue").Font.color = .Range("MiscLegend").Font.color
    .Range("BrineProcSolidWasteLabel").Font.color = .Range("MiscLegend").Font.color

End If
End With
End Sub

Sub ClearResults()

    ClearCompositionResults.Width = 250
    ClearCompositionResults.Height = 200
    ClearCompositionResults.Show
    
End Sub

Private Sub CO2ReductionList_Change()

With ThisWorkbook.Worksheets("ECLSS Composition")
' Show or hide the connections if the system is activated/deactivated
If InStr(1, .CO2ReductionList.Value, "None", vbBinaryCompare) Then

    .Shapes("CO2_ToReduction").Visible = False
    .Range("CO2_ToReductionValue").Font.color = vbWhite
    .Range("CO2_ToReductionValue").Value = 0
    
    .Shapes("CO2_FromReductionToVent").Visible = False
    .Range("CO2_FromReductionToVentLabel").Font.color = vbWhite
    .Range("CO2_FromReductionToVentValue").Font.color = vbWhite
    .Range("CO2_FromReductionToVentValue").Value = 0
    
    .Shapes("CO2_FromReductionToWaste").Visible = False
    .Range("CO2_FromReductionToWasteLabel").Font.color = vbWhite
    .Range("CO2_FromReductionToWasteValue").Font.color = vbWhite
    .Range("CO2_FromReductionToWasteValue").Value = 0
    
    .Shapes("H2O_FromCO2Reduction").Visible = False
    .Range("H2O_FromCO2ReductionValue").Font.color = vbWhite
    .Range("H2O_FromCO2ReductionValue").Value = 0
    
    .Shapes("H2O_ToCO2Reduction").Visible = False
    .Range("H2O_ToCO2ReductionValue").Font.color = vbWhite
    .Range("H2O_ToCO2ReductionValue").Value = 0
    
    .Shapes("O2_FromCo2Reduction").Visible = False
    .Range("O2_FromCO2ReductionValue").Font.color = vbWhite
    .Range("O2_FromCO2ReductionValue").Value = 0
    
    .Shapes("H2_ToCO2Reduction").Visible = False
    .Range("H2_ToCO2ReductionValue").Font.color = vbWhite
    .Range("H2_ToCO2ReductionValue").Value = 0
    
    .Shapes("CH4_FromReductionToProcessing").Visible = False
'    .Range("CH4_FromReductionToProcessingValue").Font.color = vbWhite
'    .Range("CH4_FromReductionToProcessingValue").Value = 0
    .MethaneProcessingList.Value = "None"

    
Else

    .Shapes("CO2_ToReduction").Visible = True
    .Range("CO2_ToReductionValue").Font.color = .Range("CO2Legend").Font.color
    
    .Shapes("CO2_FromReductionToVent").Visible = True
    .Range("CO2_FromReductionToVentLabel").Font.color = .Range("MiscLegend").Font.color
    .Range("CO2_FromReductionToVentValue").Font.color = .Range("MiscLegend").Font.color
    
    .Shapes("CO2_FromReductionToWaste").Visible = True
    .Range("CO2_FromReductionToWasteLabel").Font.color = .Range("MiscLegend").Font.color
    .Range("CO2_FromReductionToWasteValue").Font.color = .Range("MiscLegend").Font.color
    
    .Shapes("H2O_FromCO2Reduction").Visible = True
    .Range("H2O_FromCO2ReductionValue").Font.color = .Range("H2OLegend").Font.color
    
    .Shapes("H2O_ToCO2Reduction").Visible = True
    .Range("H2O_ToCO2ReductionValue").Font.color = .Range("H2OLegend").Font.color
    
    .Shapes("O2_FromCo2Reduction").Visible = True
    .Range("O2_FromCO2ReductionValue").Font.color = .Range("O2Legend").Font.color
    
    .Shapes("H2_ToCO2Reduction").Visible = True
    .Range("H2_ToCO2ReductionValue").Font.color = .Range("H2Legend").Font.color

' Hide components based on the selected system
    
    If InStr(1, .CO2ReductionList.Value, "PBR", vbBinaryCompare) Then
    
        .Shapes("H2_ToCO2Reduction").Visible = False
        .Range("H2_ToCO2ReductionValue").Font.color = vbWhite
        .Range("H2_ToCO2ReductionValue").Value = 0
        
        .Shapes("CO2_FromReductionToVent").Visible = False
        .Range("CO2_FromReductionToVentLabel").Font.color = vbWhite
        .Range("CO2_FromReductionToVentValue").Font.color = vbWhite
        .Range("CO2_FromReductionToVentValue").Value = 0
        
        .Range("CO2_FromReductionToWasteLabel").Value = "Algae"
        
    ElseIf InStr(1, .CO2ReductionList.Value, "SR/SCRA", vbBinaryCompare) Then
        
        .Shapes("O2_FromCo2Reduction").Visible = False
        .Range("O2_FromCO2ReductionValue").Font.color = vbWhite
        .Range("O2_FromCO2ReductionValue").Value = 0
        
        .Shapes("H2O_ToCO2Reduction").Visible = False
        .Range("H2O_ToCO2ReductionValue").Font.color = vbWhite
        .Range("H2O_ToCO2ReductionValue").Value = 0
        
        .Shapes("CO2_FromReductionToWaste").Visible = False
        .Range("CO2_FromReductionToWasteLabel").Font.color = vbWhite
        .Range("CO2_FromReductionToWasteValue").Font.color = vbWhite
        .Range("CO2_FromReductionToWasteValue").Value = 0
        
        .Range("CO2_FromReductionToVentLabel").Value = "CH4"
        
    ElseIf InStr(1, .CO2ReductionList.Value, "BR", vbBinaryCompare) Then
        
        .Shapes("O2_FromCo2Reduction").Visible = False
        .Range("O2_FromCO2ReductionValue").Font.color = vbWhite
        .Range("O2_FromCO2ReductionValue").Value = 0
        
        .Shapes("H2O_ToCO2Reduction").Visible = False
        .Range("H2O_ToCO2ReductionValue").Font.color = vbWhite
        .Range("H2O_ToCO2ReductionValue").Value = 0
        
        .Shapes("CO2_FromReductionToVent").Visible = False
        .Range("CO2_FromReductionToVentLabel").Font.color = vbWhite
        .Range("CO2_FromReductionToVentValue").Font.color = vbWhite
        .Range("CO2_FromReductionToVentValue").Value = 0
        
        .Shapes("CH4_FromReductionToProcessing").Visible = False
'       .Range("CH4_FromReductionToProcessingValue").Font.color = vbWhite
'       .Range("CH4_FromReductionToProcessingValue").Value = 0
        .MethaneProcessingList.Value = "None"

        
        .Range("CO2_FromReductionToWasteLabel").Value = "C"
        
    End If

End If
End With
End Sub


Sub ECLSSCompUserGuid()

    Worksheets("ECLSS Composition").Protect UserInterfaceOnly:=True
    UG_ECLSSComposition_1.Width = 500
    UG_ECLSSComposition_1.Height = 550
    UG_ECLSSComposition_1.Show

End Sub

Private Sub ElectrolysisList_Change()

With ThisWorkbook.Worksheets("ECLSS Composition")
If InStr(1, .ElectrolysisList.Value, "None", vbBinaryCompare) Then

    .Shapes("H2O_ToElectrolysis").Visible = False
    .Range("H2O_ToElectrolysisValue").Font.color = vbWhite
    .Range("H2O_ToElectrolysisValue").Value = 0

    .Shapes("H2_FromElectrolysis").Visible = False
    .Range("H2_FromElectrolysisValue").Font.color = vbWhite
    .Range("H2_FromElectrolysisValue").Value = 0

    .Shapes("O2_FromElectrolysis").Visible = False
    .Range("O2_FromElectrolysisValue").Font.color = vbWhite
    .Range("O2_FromElectrolysisValue").Value = 0

Else

    .Shapes("H2O_ToElectrolysis").Visible = True
    .Range("H2O_ToElectrolysisValue").Font.color = .Range("H2OLegend").Font.color

    .Shapes("H2_FromElectrolysis").Visible = True
    .Range("H2_FromElectrolysisValue").Font.color = .Range("H2Legend").Font.color
    
    .Shapes("O2_FromElectrolysis").Visible = True
    .Range("O2_FromElectrolysisValue").Font.color = .Range("O2Legend").Font.color

End If
End With
End Sub

Private Sub CO2_RemovalList_Change()
' if CAMRAS or 2BMS is select, select the same system for the CO2 removal!
With ThisWorkbook.Worksheets("ECLSS Composition")

.Shapes("H2O_ToCO2Removal").Visible = False
.Range("H2O_ToCO2RemovalValue").Font.color = vbWhite
.Range("H2O_ToCO2RemovalValue").Value = 0
.Range("H2O_FromCO2RemovalValue").Font.color = vbWhite
.Range("H2O_FromCO2RemovalValue").Value = 0

.Range("CO2_FromRemovalValue").Font.color = .Range("CO2Legend").Font.color

.Shapes("CO2_RemovalToWaste").Visible = False
.Range("CO2_RemovalToWasteLabel1").Font.color = vbWhite
.Range("CO2_RemovalToWasteLabel2").Font.color = vbWhite
.Range("CO2_RemovalToWasteValue1").Font.color = vbWhite
.Range("CO2_RemovalToWasteValue2").Font.color = vbWhite
.Range("CO2_RemovalToWasteValue1").Value = 0
.Range("CO2_RemovalToWasteValue2").Value = 0

.Shapes("H2O_FromHumControlToCO2Removal").Visible = False
.Range("H2O_FromHumControlToCO2RemovalValue").Font.color = vbWhite
.Range("H2O_FromHumControlToCO2RemovalValue").Value = 0


If InStr(1, .CO2_RemovalList.Value, "CAMRAS", vbBinaryCompare) Or InStr(1, .CO2_RemovalList.Value, "2BMS", vbBinaryCompare) Then
    If .CO2_RemovalList.Value <> .HumidityControlList.Value Then
        .HumidityControlList.Value = .CO2_RemovalList.Value
    End If
ElseIf InStr(1, .CO2_RemovalList.Value, "LiOH", vbBinaryCompare) Then

    .Shapes("CO2_RemovalToWaste").Visible = True
    .Range("CO2_RemovalToWasteLabel1").Font.color = .Range("MiscLegend").Font.color
    .Range("CO2_RemovalToWasteLabel2").Font.color = .Range("MIscLegend").Font.color
    .Range("CO2_RemovalToWasteValue1").Font.color = .Range("MIscLegend").Font.color
    .Range("CO2_RemovalToWasteValue2").Font.color = .Range("MIscLegend").Font.color
    
    .Range("H2O_FromCO2RemovalValue").Font.color = .Range("H2OLegend").Font.color
    
    .Range("CO2_FromRemovalValue").Font.color = vbWhite
    .Range("CO2_FromRemovalValue").Value = 0
    
    
ElseIf InStr(1, .CO2_RemovalList.Value, "SAWD", vbBinaryCompare) Then

    .Shapes("H2O_ToCO2Removal").Visible = True
    .Range("H2O_ToCO2RemovalValue").Font.color = .Range("H2OLegend").Font.color
    .Range("H2O_FromCO2RemovalValue").Font.color = .Range("H2OLegend").Font.color
    
End If
End With
End Sub

Sub ExecuteCalculations()

    CalculateECLSS
    
End Sub

Private Sub HumidityControlList_Change()


With ThisWorkbook.Worksheets("ECLSS Composition")

' if CAMRAS or 2BMS is select, select the same system for the CO2 removal!
If InStr(1, .HumidityControlList.Value, "CAMRAS", vbBinaryCompare) Or InStr(1, .HumidityControlList.Value, "2BMS", vbBinaryCompare) Then
    If .CO2_RemovalList.Value <> .HumidityControlList.Value Then
        .CO2_RemovalList.Value = .HumidityControlList.Value
    End If
    
    ' These two system also vent the humidity so we show the corresponding objects
    .Shapes("HumControlToVent").Visible = True
    .Range("HumControlToVentValue").Font.color = .Range("H2OLegend").Font.color
    
    .Shapes("HumControlToWaste").Visible = False
    .Range("HumControlToWasteValue").Font.color = vbWhite
    .Range("HumControlToWasteValue").Value = 0
    
    .Shapes("HumControlToWWFiltration").Visible = False
    .Range("HumControlToWWFiltrationValue").Font.color = vbWhite
    .Range("HumControlToWWFiltrationValue").Value = 0
Else
    .Shapes("HumControlToVent").Visible = False
    .Range("HumControlToVentValue").Font.color = vbWhite
    .Range("HumControlToVentValue").Value = 0
    
    If InStr(1, .WWFiltrationList.Value, "None", vbBinaryCompare) Then
        .Shapes("HumControlToWaste").Visible = True
        .Range("HumControlToWasteValue").Font.color = .Range("H2OLegend").Font.color
    Else
        .Shapes("HumControlToWWFiltration").Visible = True
        .Range("HumControlToWWFiltrationValue").Font.color = .Range("H2OLegend").Font.color
    End If
End If
End With

End Sub

Sub InitializeDropDownButton()

    CreateDropDownLists
    
End Sub

Sub LoadResults()

    LoadCompositionResults.Width = 280
    LoadCompositionResults.Height = 190
    LoadCompositionResults.Show

End Sub




Private Sub OptionIncinerateInedible_Click()

End Sub

Private Sub PGC_NumberOfUnits_Change()

End Sub

Private Sub PlantGrowthChamberList_Change()
' Check if the user selected veggie or ISPR PGC and if that is the case show the number of units IF, if not hide it!

With ThisWorkbook.Worksheets("ECLSS Composition")
If InStr(1, .PlantGrowthChamberList.Value, "Veggie", vbBinaryCompare) Or InStr(1, .PlantGrowthChamberList.Value, "ISPR PGC", vbBinaryCompare) Then
    .PGC_NumberOfUnits.Visible = True
    .PGC_LabelNumberOfUnits.Visible = True
Else
    .PGC_NumberOfUnits.Visible = False
    .PGC_LabelNumberOfUnits.Visible = False
End If

' Show hide connections if plant chamber is enabled/disabled
If InStr(1, .PlantGrowthChamberList.Value, "None", vbBinaryCompare) Then
    If InStr(1, .PBRList.Value, "No", vbBinaryCompare) Then
    .Shapes("Nutrients").Visible = False
    ' remove option to incinerate waste biomass and set parameter to false
    .Shapes("WasteStorage").TextFrame.VerticalAlignment = xlVAlignCenter
    .OptionIncinerateInedible.Value = False
    .OptionIncinerateInedible.Visible = False
    End If
    
    .Shapes("CO2_ToPlants").Visible = False
    .Range("CO2toPlantsValue").Font.color = vbWhite
    .Range("CO2toPlantsValue").Value = 0
    
    .Shapes("H2O_FromPlants").Visible = False
    .Range("H2OfromPlantsValue").Font.color = vbWhite
    .Range("H2OfromPlantsValue").Value = 0
    .Shapes("H2O_FromPlants").Visible = False
    .Range("H2OfromPlantsValue").Font.color = vbWhite
    .Range("H2OfromPlantsValue").Value = 0

    .Shapes("O2_FromPlants").Visible = False
    .Range("O2fromPlantsValue").Font.color = vbWhite
    .Range("O2fromPlantsValue").Value = 0

    .Shapes("H2O_ConPlants").Visible = False
    .Range("H2OtoPlantsValue").Font.color = vbWhite
    .Range("H2OtoPlantsValue").Value = 0
    
    .Shapes("NutrientsToPlants").Visible = False
    .Range("NutrientsToPlantsLabel").Font.color = vbWhite
    .Range("NutrientsToPlantsValue").Font.color = vbWhite
    .Range("NutrientsToPlantsValue").Value = 0

    .Shapes("EdibleFromPlants").Visible = False
    .Range("EdibleFromPlantsLabel").Font.color = vbWhite
    .Range("EdibleFromPlantsValue").Font.color = vbWhite
    .Range("EdibleFromPlantsValue").Value = 0
    
    .Shapes("Waste_FromPlants").Visible = False
    .Range("Waste_FromPlantsLabel").Font.color = vbWhite
    .Range("Waste_FromPlantsValue").Font.color = vbWhite
    .Range("Waste_FromPlantsValue").Value = 0
    
    .Range("EdibleFromPlantsEnergyLabel").Font.color = vbWhite
    .Range("EdibleFromPlantsEnergyValue").Font.color = vbWhite
    .Range("EdibleFromPlantsEnergyValue").Value = 0
    
    .Range("EdibleFromPlantsWaterLabel").Font.color = vbWhite
    .Range("EdibleFromPlantsWaterValue").Font.color = vbWhite
    .Range("EdibleFromPlantsWaterValue").Value = 0
    
    .Range("PlantWaterContentToCrewLabel").Font.color = vbWhite
    .Range("PlantWaterContentToCrewLabel2").Font.color = vbWhite
    .Range("PlantWaterContentToCrew").Font.color = vbWhite
    .Range("PlantWaterContentToCrew").Value = 0
    
    .Shapes("HumControlToPlants").Visible = False
    .Range("HumControlToPlantsValue").Font.color = vbWhite
    .Range("HumControlToPlantsValue").Value = 0
    
    .Range("EdibleFromPlantsValue2").Font.color = vbWhite
    .Range("EdibleFromPlantsLabel2").Font.color = vbWhite
    
    .Range("H2O_InediblePlantsLabel").Font.color = vbWhite
    .Shapes("H2O_InediblePlantsToFiltration").Visible = False
    .Range("H2O_InediblePlantsValue").Value = 0
    .Range("H2O_InediblePlantsValue").Font.color = vbWhite
    .Range("H2O_InediblePlantsToFiltrationValue").Value = 0
    .Range("H2O_InediblePlantsToFiltrationValue").Font.color = vbWhite
    
    If InStr(1, .HumidityControlList.Value, "2BMS", vbBinaryCompare) Or InStr(1, .HumidityControlList.Value, "CAMRAS", vbBinaryCompare) Then
    Else
        If InStr(1, .WWFiltrationList.Value, "None", vbBinaryCompare) Then
            .Shapes("HumControlToWaste").Visible = True
        Else
            .Shapes("HumControlToWWFiltration").Visible = True
        End If
    End If

    
    .Shapes("CO2_FromWasteToCabin1").Visible = False
    .Shapes("CO2_FromWasteToCabin2").Visible = False
    .Range("CO2_FromWasteToCabinValue").Value = 0
    .Range("CO2_FromWasteToCabinValue").Font.color = vbWhite
    
    .PGCurineTreatmentOption.Value = False
    
Else
    .Shapes("Nutrients").Visible = True
    
    .Shapes("CO2_ToPlants").Visible = True
    .Range("CO2toPlantsValue").Font.color = .Range("CO2Legend").Font.color
    
    .Shapes("H2O_FromPlants").Visible = True
    .Range("H2OfromPlantsValue").Font.color = .Range("H2OLegend").Font.color

    .Shapes("O2_FromPlants").Visible = True
    .Range("O2fromPlantsValue").Font.color = .Range("O2Legend").Font.color

    .Shapes("H2O_ConPlants").Visible = True
    .Range("H2OtoPlantsValue").Font.color = .Range("H2OLegend").Font.color
    
    .Range("PlantWaterContentToCrew").Font.color = .Range("H2OLegend").Font.color
    .Range("PlantWaterContentToCrewLabel").Font.color = .Range("H2OLegend").Font.color
    .Range("PlantWaterContentToCrewLabel2").Font.color = .Range("H2OLegend").Font.color
    
    .Range("EdibleFromPlantsWaterValue").Font.color = .Range("H2OLegend").Font.color
    .Range("EdibleFromPlantsWaterLabel").Font.color = .Range("H2OLegend").Font.color
    
    .Range("H2O_InediblePlantsLabel").Font.color = .Range("MiscLegend").Font.color
    .Range("H2O_InediblePlantsValue").Value = 0
    .Range("H2O_InediblePlantsValue").Font.color = .Range("H2OLegend").Font.color
    
    .Range("H2O_InediblePlantsLabel").Font.color = .Range("MiscLegend").Font.color
    .Range("EdibleFromPlantsEnergyValue").Font.color = .Range("MiscLegend").Font.color
    .Range("EdibleFromPlantsEnergyLabel").Font.color = .Range("MiscLegend").Font.color
    
    .Shapes("NutrientsToPlants").Visible = True
    .Range("NutrientsToPlantsLabel").Font.color = .Range("MiscLegend").Font.color
    .Range("NutrientsToPlantsValue").Font.color = .Range("MiscLegend").Font.color

    .Shapes("EdibleFromPlants").Visible = True
    .Range("EdibleFromPlantsLabel").Font.color = .Range("MiscLegend").Font.color
    .Range("EdibleFromPlantsValue").Font.color = .Range("MiscLegend").Font.color
    
    .Shapes("Waste_FromPlants").Visible = True
    .Range("Waste_FromPlantsLabel").Font.color = .Range("MiscLegend").Font.color
    .Range("Waste_FromPlantsValue").Font.color = .Range("MiscLegend").Font.color
    
    .Range("EdibleFromPlantsValue2").Font.color = .Range("MiscLegend").Font.color
    .Range("EdibleFromPlantsLabel2").Font.color = .Range("MiscLegend").Font.color
    
    ' add option to incinerate waste biomass and set parameter to true
    .Shapes("WasteStorage").TextFrame.VerticalAlignment = xlVAlignTop
    .OptionIncinerateInedible.Value = True
    .OptionIncinerateInedible.Visible = True
    
    .Shapes("CO2_FromWasteToCabin1").Visible = True
    .Shapes("CO2_FromWasteToCabin2").Visible = True
    .Range("CO2_FromWasteToCabinValue").Value = 0
    .Range("CO2_FromWasteToCabinValue").Font.color = .Range("CO2Legend").Font.color
    
End If
End With
End Sub

Private Sub UrineProcessingList_Change()
' here check if no urine processing is used, and if that is the case hide the connections and set values to 0
' if SCWO is used change the connection to produce potable water, since this is the only Urine Processing tech that can recover potable water directly!

Dim CrewUrineProduction As Double

If Worksheets("User Interface").Range("Schedule_Number").Value = 1 Then
    CrewUrineProduction = Worksheets("Schedule and Metabolic").Range("SUM_Urine_Production").Value
Else
    CrewUrineProduction = Worksheets("Schedule and Metabolic").Range("SUM_Urine_Production_2").Value
End If
' Divide value with crew to get value per CM
CrewUrineProduction = CrewUrineProduction / Worksheets("User Interface").Range("Crew_Size").Value

With ThisWorkbook.Worksheets("ECLSS Composition")

If InStr(1, .UrineProcessingList.Value, "None", vbBinaryCompare) Then

.Shapes("Urine_ProdCrew").Visible = False
.Range("CrewUrineWithProcessing").Font.color = vbWhite
.Range("CrewUrineWithProcessing").Value = 0

.Shapes("GreyWaterFromUrine").Visible = False
.Range("UrineProcessingGreyWater").Font.color = vbWhite
.Range("UrineProcessingGreyWater").Value = 0

.Shapes("Brine_Prod").Visible = False
.Range("BrineProdValue").Font.color = vbWhite
.Range("BrineProdValue").Value = 0

.Range("CrewUrineWithoutProcessing").Font.color = .Range("MiscLegend").Font.color
.Range("CrewUrineWithoutProcessingLabel").Font.color = .Range("MiscLegend").Font.color
.Range("CrewUrineWithoutProcessing").Value = CrewUrineProduction

Else

.Shapes("Urine_ProdCrew").Visible = True
.Range("CrewUrineWithProcessing").Font.color = .Range("GreyWaterLegend").Font.color
.Range("CrewUrineWithProcessing").Value = CrewUrineProduction

.Shapes("GreyWaterFromUrine").Visible = True
.Range("UrineProcessingGreyWater").Font.color = .Range("GreyWaterLegend").Font.color

.Shapes("Brine_Prod").Visible = True
.Range("BrineProdValue").Font.color = .Range("GreyWaterLegend").Font.color

.Range("CrewUrineWithoutProcessing").Font.color = vbWhite
.Range("CrewUrineWithoutProcessingLabel").Font.color = vbWhite
.Range("CrewUrineWithoutProcessing").Value = 0


    If InStr(1, .UrineProcessingList.Value, "SCWO", vbBinaryCompare) Then
        .Shapes("H2O_FromUrineProcessing").Visible = True
        .Range("UrineProcesingPotableWater").Font.color = .Range("H2OLegend").Font.color
        .Shapes("GreyWaterFromUrine").Visible = False
        .Range("UrineProcessingGreyWater").Font.color = vbWhite
    Else
        .Shapes("H2O_FromUrineProcessing").Visible = False
        .Range("UrineProcesingPotableWater").Font.color = vbWhite
        .Shapes("GreyWaterFromUrine").Visible = True
        .Range("UrineProcessingGreyWater").Font.color = .Range("GreyWaterLegend").Font.color
    End If
End If
End With
    
End Sub

Private Sub WWFiltrationList_Change()

Dim CrewHygieneH2OConsumption As Double

' get crew hygiene water consumption
If Worksheets("User Interface").Range("Schedule_Number").Value = 1 Then
    CrewHygieneH2OConsumption = Worksheets("Schedule and Metabolic").Range("SUM_Hygiene_Water_Need").Value
Else
    CrewHygieneH2OConsumption = Worksheets("Schedule and Metabolic").Range("SUM_Hygiene_Water_Need_2").Value
End If

With ThisWorkbook.Worksheets("ECLSS Composition")

If InStr(1, .WWFiltrationList.Value, "None", vbBinaryCompare) Then

    .Shapes("HygieneH2O_ProdCrew").Visible = False
    .Range("HygieneH2O_ProdCrewValue").Font.color = vbWhite
    .Range("HygieneH2O_ProdCrewValue").Value = 0
    
    .Shapes("HumControlToWWFiltration").Visible = False
    .Range("HumControlToWWFiltrationValue").Font.color = vbWhite
    .Range("HumControlToWWFiltrationValue").Value = 0
    
    .Shapes("H2O_FromWWFilt").Visible = False
    .Range("H2O_FromWWFiltValue").Font.color = vbWhite
    .Range("H2O_FromWWFiltValue").Value = 0
    
    .Shapes("HumControlToWaste").Visible = True
    .Range("HumControlToWasteValue").Font.color = .Range("H2OLegend").Font.color
    
    .Shapes("HumControlToWaste").Visible = True
    .Range("HygieneH2O_ProdCrewWithoutFiltrationLabel").Font.color = .Range("MiscLegend").Font.color
    .Range("HygieneH2O_ProdCrewWithoutFiltrationValue").Font.color = .Range("MiscLegend").Font.color
    .Range("HygieneH2O_ProdCrewWithoutFiltrationValue").Value = CrewHygieneH2OConsumption / Worksheets("User Interface").Range("Crew_Size").Value
    
Else


    .Shapes("HygieneH2O_ProdCrew").Visible = True
    .Range("HygieneH2O_ProdCrewValue").Font.color = .Range("H2OLegend").Font.color
    .Range("HygieneH2O_ProdCrewValue").Value = CrewHygieneH2OConsumption / Worksheets("User Interface").Range("Crew_Size").Value
    
    .Shapes("HumControlToWWFiltration").Visible = True
    .Range("HumControlToWWFiltrationValue").Font.color = .Range("H2OLegend").Font.color
    
    .Shapes("H2O_FromWWFilt").Visible = True
    .Range("H2O_FromWWFiltValue").Font.color = .Range("H2OLegend").Font.color
    
    .Shapes("HumControlToWaste").Visible = False
    .Range("HumControlToWasteValue").Font.color = vbWhite
    .Range("HumControlToWasteValue").Value = 0
    
    .Range("HygieneH2O_ProdCrewWithoutFiltrationLabel").Font.color = vbWhite
    .Range("HygieneH2O_ProdCrewWithoutFiltrationValue").Font.color = vbWhite
    .Range("HygieneH2O_ProdCrewWithoutFiltrationValue").Value = 0

End If
End With
End Sub

Private Sub PBRList_Change()

With ThisWorkbook.Worksheets("ECLSS Composition")

If InStr(1, .PBRList.Value, "None", vbBinaryCompare) Then

    .Shapes("CO2ToPBR").Visible = False
    .Range("CO2ToPBRValue").Font.color = vbWhite
    .Range("CO2ToPBRValue").Value = 0

    .Shapes("O2FromPBR").Visible = False
    .Range("O2FromPBRValue").Font.color = vbWhite
    .Range("O2FromPBRValue").Value = "0"

    .Shapes("FoodFromPBR").Visible = False
    .Range("FoodFromPBRValue").Font.color = vbWhite
    .Range("FoodFromPBRValue").Value = 0

    .Shapes("InedibleBiomassPBR").Visible = False
    .Range("InedibleBiomassPBRValue").Font.color = vbWhite
    .Range("InedibleBiomassPBRValue").Value = 0
    
    .Shapes("NutrientsToPBR").Visible = False
    .Range("NutrientsToPBRValue").Font.color = vbWhite
    .Range("NutrientsToPBRValue").Value = 0
    
    .Shapes("WaterToPBR").Visible = False
    .Range("WaterToPBRValue").Font.color = vbWhite
    .Range("WaterToPBRValue").Value = 0
    
    .Shapes("WaterFromPBR").Visible = False
    .Range("WaterFromPBRValue").Font.color = vbWhite
    .Range("WaterFromPBRValue").Value = 0
    
    If InStr(1, .PlantGrowthChamberList.Value, "None", vbBinaryCompare) Then
    .Shapes("Nutrients").Visible = False
    ' remove option to incinerate waste biomass and set parameter to false
    .Shapes("WasteStorage").TextFrame.VerticalAlignment = xlVAlignCenter
    .OptionIncinerateInedible.Value = False
    .OptionIncinerateInedible.Visible = False
    End If

Else


    .Shapes("CO2ToPBR").Visible = True
    .Range("CO2ToPBRValue").Font.color = .Range("CO2Legend").Font.color
    .Range("CO2ToPBRValue").Value = 0 'CrewHygieneH2OConsumption / Worksheets("User Interface").Range("Crew_Size").Value

    .Shapes("O2FromPBR").Visible = True
    .Range("O2FromPBRValue").Font.color = .Range("O2Legend").Font.color
    .Range("O2FromPBRValue").Value = 0

    .Shapes("FoodFromPBR").Visible = True
    .Range("FoodFromPBRValue").Font.color = .Range("MiscLegend").Font.color
    .Range("FoodFromPBRValue").Value = 0

    .Shapes("InedibleBiomassPBR").Visible = True
    .Range("InedibleBiomassPBRValue").Font.color = .Range("MiscLegend").Font.color
    .Range("InedibleBiomassPBRValue").Value = 0
    
    .Shapes("NutrientsToPBR").Visible = True
    .Range("NutrientsToPBRValue").Font.color = .Range("MiscLegend").Font.color
    .Range("NutrientsToPBRValue").Value = 0
    
    .Shapes("WaterToPBR").Visible = True
    .Range("WaterToPBRValue").Font.color = .Range("WaterLegend").Font.color
    .Range("WaterToPBRValue").Value = 0
    
    .Shapes("WaterFromPBR").Visible = True
    .Range("WaterFromPBRValue").Font.color = .Range("WaterLegend").Font.color
    .Range("WaterFromPBRValue").Value = 0
    
    .Shapes("Nutrients").Visible = True
    
    .Shapes("WasteStorage").TextFrame.VerticalAlignment = xlVAlignTop
    .OptionIncinerateInedible.Value = True
    .OptionIncinerateInedible.Visible = True

End If

End With
End Sub

Private Sub H2StorageList_Change()

With ThisWorkbook.Worksheets("ECLSS Composition")

    If InStr(1, .H2StorageList.Value, "None", vbBinaryCompare) Then
        If InStr(1, .MethaneProcessingList.Value, "PPU", vbBinaryCompare) Then
            MsgBox ("PPU can not close H2 loop. H2-Storage is reactivated")
            .H2StorageList.Value = "HPS"
        End If
    End If
End With
End Sub

Private Sub MethaneProcessingList_Change()

With ThisWorkbook.Worksheets("ECLSS Composition")

If InStr(1, .MethaneProcessingList.Value, "None", vbBinaryCompare) Then

    .Shapes("FromMethaneProcessingToVent").Visible = False
    .Range("FromMethaneProcessingToVentValue").Font.color = vbWhite
    .Range("FromMethaneProcessingToVentValue").Value = 0
    .Range("FromMethaneProcessingToVentlabel").Font.color = vbWhite
    
    .Shapes("FromMethaneProcessingToWaste").Visible = False
    .Range("FromMethaneProcessingToWasteValue").Font.color = vbWhite
    .Range("FromMethaneProcessingToWasteValue").Value = 0
    .Range("FromMethaneProcessingToWasteLabel").Font.color = vbWhite

    .Shapes("H2_FromMethaneProcessing").Visible = False
    .Range("H2_FromMethaneProcessingValue").Font.color = vbWhite
    .Range("H2_FromMethaneProcessingValue").Value = 0

    .Shapes("CH4_FromReductionToProcessing").Visible = False
    .Range("CH4_FromReductionToProcessingValue").Font.color = vbWhite
    .Range("CH4_FromReductionToProcessingValue").Value = 0

    If InStr(1, .CO2ReductionList.Value, "SR/SCRA", vbBinaryCompare) Then
    .Shapes("CO2_FromReductionToVent").Visible = True
    .Range("CO2_FromReductionToVentValue").Font.color = .Range("MiscLegend").Font.color
    .Range("CO2_FromReductionToVentValue").Value = 0
    .Range("CO2_FromReductionToVentLabel").Font.color = .Range("MiscLegend").Font.color
    .Range("CO2_FromReductionToVentLabel").Value = "CH4"
    End If
    
    

Else

    If InStr(1, .CO2ReductionList.Value, "SR/SCRA", vbBinaryCompare) Then
    
    If InStr(1, .MethaneProcessingList.Value, "PPU", vbBinaryCompare) Then
    
        .Shapes("FromMethaneProcessingToVent").Visible = True
        .Range("FromMethaneProcessingToVentValue").Font.color = .Range("MiscLegend").Font.color
        .Range("FromMethaneProcessingToVentValue").Value = 0
        .Range("FromMethaneProcessingToVentlabel").Font.color = .Range("MiscLegend").Font.color
    
        .Shapes("FromMethaneProcessingToWaste").Visible = False
        .Range("FromMethaneProcessingToWasteValue").Font.color = vbWhite
        .Range("FromMethaneProcessingToWasteValue").Value = 0
        .Range("FromMethaneProcessingToWasteLabel").Font.color = vbWhite
        
        If InStr(1, .H2StorageList.Value, "None", vbBinaryCompare) Then
            MsgBox ("PPU can not close H2-loop. H2-Storage is activated")
            H2StorageList.Value = "HPS"
        End If
        
        .Range("CO2_FromReductionToVentValue").Font.color = vbWhite
        .Shapes("CO2_FromReductionToVent").Visible = False
        .Range("CO2_FromReductionToVentLabel").Font.color = vbWhite
        .Range("CO2_FromReductionToVentLabel").Value = ""
    
        
    ElseIf InStr(1, .MethaneProcessingList.Value, "LPU", vbBinaryCompare) Then
    
        .Shapes("FromMethaneProcessingToWaste").Visible = True
        .Range("FromMethaneProcessingToWasteValue").Font.color = .Range("MiscLegend").Font.color
        .Range("FromMethaneProcessingToWasteValue").Value = 0
        .Range("FromMethaneProcessingToWasteLabel").Font.color = .Range("MiscLegend").Font.color
    
        .Shapes("FromMethaneProcessingToVent").Visible = False
        .Range("FromMethaneProcessingToVentValue").Font.color = vbWhite
        .Range("FromMethaneProcessingToVentValue").Value = 0
        .Range("FromMethaneProcessingToVentlabel").Font.color = vbWhite
        
        .Range("CO2_FromReductionToVentValue").Font.color = .Range("MiscLegend").Font.color
        .Shapes("CO2_FromReductionToVent").Visible = True
        .Range("CO2_FromReductionToVentLabel").Font.color = .Range("MiscLegend").Font.color
        .Range("CO2_FromReductionToVentLabel").Value = "CH4"
    
    End If
    
    

    .Shapes("H2_FromMethaneProcessing").Visible = True
    .Range("H2_FromMethaneProcessingValue").Font.color = .Range("H2Legend").Font.color
    .Range("H2_FromMethaneProcessingValue").Value = 0

    .Shapes("CH4_FromReductionToProcessing").Visible = True
    .Range("CH4_FromReductionToProcessingValue").Font.color = .Range("MiscLegend").Font.color
    .Range("CH4_FromReductionToProcessingValue").Value = 0
    
    Else
    
    MsgBox ("No CH4 available from CO2-Reduction")
    .MethaneProcessingList.Value = "None"
    
    End If
    
End If

End With
End Sub
