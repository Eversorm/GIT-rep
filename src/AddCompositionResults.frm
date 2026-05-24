VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddCompositionResults 
   Caption         =   "Add Results"
   ClientHeight    =   3015
   ClientLeft      =   -1200
   ClientTop       =   -5880
   ClientWidth     =   2250
   OleObjectBlob   =   "AddCompositionResults.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddCompositionResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddResults_Click()

    ' Add code here to add the values to the ESM
    ' Check if dynamic results should also be overwritten
    Dim iResultsHeaderRow As Integer
    Dim iResultsHeaderRow2 As Integer
    Dim ResultsColumn
    Dim CrewSize As Integer
    Dim MissionDuration As Double
    Dim PressurizedVolume As Double
    Dim Schedule As Integer
    Dim i As Integer
    
    iResultsHeaderRow = 6
    iResultsHeaderRow2 = 92
    ResultsColumn = findColumnByName(Array(Results.Value), "Output Data", iResultsHeaderRow)
    ResultsColumn = ResultsColumn(1)
    
    If ResultsColumn = 0 Then
        ' in this case the result is provided with a user name:
        ResultsColumn = findColumnByName(Array(Results.Value), "Output Data", iResultsHeaderRow + 1)
        ResultsColumn = ResultsColumn(1)
    End If
    
    CrewSize = Worksheets("User Interface").Range("Crew_Size").Value
    MissionDuration = Worksheets("User Interface").Range("Mission_Duration").Value
    PressurizedVolume = Worksheets("User Interface").Range("Pressurized_volume").Value
    Schedule = Worksheets("User Interface").Range("Schedule_Number").Value

    With ThisWorkbook.Worksheets("Output Data")
        If UserNameForResults.Value <> "" Then
            .Cells(iResultsHeaderRow + 1, ResultsColumn).Value = UserNameForResults.Value
            .Cells(iResultsHeaderRow2 + 1, ResultsColumn).Value = UserNameForResults.Value
        Else
        ' if no user name is provided and the value is currently empty, we call it "results x"
            If IsEmpty(.Cells(iResultsHeaderRow + 1, ResultsColumn).Value) Then
                .Cells(iResultsHeaderRow + 1, ResultsColumn).Value = Results.Value
                .Cells(iResultsHeaderRow2 + 1, ResultsColumn).Value = Results.Value
            End If
        End If
    
        '------------------------------------------------------------------------------------------------------------------------------------------------
        'ESM Calculation for the selected ECLSS composition
        Dim ESM_LocationF As Double
        Dim ESM_VolumeF As Double
        Dim ESM_PowerF As Double
        Dim ESM_CoolingF As Double
        Dim ESM_Crew_TimeF As Double
        Dim FixESM As Double
        Dim ScalingESM As Double
        
        ESM_LocationF = Worksheets("User Interface").Range("ESM_LocationF").Value
        ESM_VolumeF = Worksheets("User Interface").Range("ESM_VolumeF").Value
        ' Since power and cooling factors are provided per kW, we divide them with 1000 since the technology values are W
        ESM_PowerF = Worksheets("User Interface").Range("ESM_PowerF").Value / 1000
        ESM_CoolingF = Worksheets("User Interface").Range("ESM_CoolingF").Value / 1000
        ESM_Crew_TimeF = Worksheets("User Interface").Range("ESM_Crew_TimeF").Value
        
        Dim ACS_ESM As Double
        Dim AR_ESM As Double
        Dim WRM_ESM As Double
        Dim THC_ESM As Double
        Dim PGC_ESM As Double
        Dim Waste_ESM As Double
        Dim Food_ESM As Double
        Dim Cloth_ESM As Double
            
        ACS_ESM = 0
        AR_ESM = 0
        WRM_ESM = 0
        THC_ESM = 0
        PGC_ESM = 0
        Waste_ESM = 0
        Food_ESM = 0
        Cloth_ESM = 0
        
        Dim TotalSystemMass As Double
        Dim TotalSystemVolume As Double
        Dim TotalSystemCooling As Double
        Dim TotalSystemPower As Double
        Dim TotalSystemResupply As Double
        Dim TotalSystemResupplyVolume As Double
        Dim TotalSystemMaintenance As Double
        Dim TotalSpareMass As Double
        Dim TotalSpareVolume As Double
        
        TotalSystemMass = 0
        TotalSystemVolume = 0
        TotalSystemCooling = 0
        TotalSystemPower = 0
        TotalSystemResupply = 0
        TotalSystemResupplyVolume = 0
        TotalSystemMaintenance = 0
        TotalSpareMass = 0
        TotalSpareVolume = 0
        
        
        ' Calculate storage systems
        Dim O2_MassBalance As Double
        Dim H2_MassBalance As Double
        Dim H2O_MassBalance As Double
        Dim Waste_MassBalance As Double
        Dim Food_MassBalance As Double
        Dim Nutrients_MassBalance As Double
        Dim TankValues
        Dim ReliabilityValues
        Dim SystemMass As Double
        
        Dim miResupplyRows
        miResupplyRows = findRowByName(Array("Clothes/misc.", "O2 res.", "H2 res.", "N2 res.", "H2O res.", "Food res.", "Nutrients res.", "PGC res.", "PBR res.", "CH4 proc. res."), "Output Data", 1)
        
        With ThisWorkbook.Worksheets("ECLSS Composition")
        
            O2_MassBalance = .Range("O2_FromElectrolysisValue").Value + .Range("O2_FromCO2ReductionValue").Value + .Range("O2FromPBRValue").Value - .Range("O2_FromStorageValue").Value
            
            H2_MassBalance = .Range("H2_FromElectrolysisValue").Value - .Range("H2_ToCO2ReductionValue").Value - .Range("H2_ToVentValue").Value + .Range("H2_FromMethaneProcessingValue").Value
            
            H2O_MassBalance = .Range("H2O_FromWWFiltValue").Value + .Range("H2O_FromCO2ReductionValue").Value - .Range("WaterToPBRValue").Value + .Range("WaterFromPBRValue").Value _
                            - .Range("H2O_ToElectrolysisValue").Value - .Range("H2O_ToCO2ReductionValue").Value - .Range("HygieneH2O_ConCrewValue").Value - .Range("H2O_ConCrewValue").Value - .Range("H2OtoPlantsValue").Value
                            
            Waste_MassBalance = .Range("CrewFecesValue").Value + .Range("CrewUrineWithoutProcessing").Value + .Range("UrineProcessingGreyWaterToWasteValue").Value + .Range("HygieneH2O_ProdCrewWithoutFiltrationValue").Value + .Range("BrineProcSolidWasteValue").Value + .Range("InedibleBiomassPBRValue").Value _
                              + .Range("CO2_FromReductionToWasteValue").Value + .Range("Waste_FromPlantsValue").Value + .Range("FoodPackagingValue").Value + .Range("WWFiltrationToWasteValue").Value - .Range("H2O_InediblePlantsToFiltrationValue").Value + .Range("FromMethaneProcessingToWasteValue").Value
            
            ' plants are always supplied to the human in the same amount as they are produced, therefore not considered here
            Food_MassBalance = -.Range("FoodConCrewValue").Value - .Range("FoodPackagingValue").Value + .Range("FoodFromPBRValue").Value
            Nutrients_MassBalance = -.Range("NutrientsToPlantsValue").Value - .Range("NutrientsToPBRValue").Value
            
            O2_MassBalance = O2_MassBalance * CrewSize
            H2_MassBalance = H2_MassBalance * CrewSize
            H2O_MassBalance = H2O_MassBalance * CrewSize
            Waste_MassBalance = Waste_MassBalance * CrewSize
            Food_MassBalance = Food_MassBalance * CrewSize
            Nutrients_MassBalance = Nutrients_MassBalance * CrewSize
            
            Worksheets("Output Data").Cells(miResupplyRows(2), ResultsColumn) = -O2_MassBalance
            Worksheets("Output Data").Cells(miResupplyRows(3), ResultsColumn) = -H2_MassBalance
            Worksheets("Output Data").Cells(miResupplyRows(4), ResultsColumn) = .Range("N2_LeakageValue").Value
            If H2O_MassBalance < 0 Then
                Worksheets("Output Data").Cells(miResupplyRows(5), ResultsColumn) = -H2O_MassBalance
            Else
                Worksheets("Output Data").Cells(miResupplyRows(5), ResultsColumn) = 0
            End If
            Worksheets("Output Data").Cells(miResupplyRows(6), ResultsColumn) = -Food_MassBalance
            Worksheets("Output Data").Cells(miResupplyRows(7), ResultsColumn) = -Nutrients_MassBalance
            
            O2_MassBalance = Abs(O2_MassBalance)
            H2_MassBalance = Abs(H2_MassBalance)
            Waste_MassBalance = Abs(Waste_MassBalance)
            Food_MassBalance = Abs(Food_MassBalance)
            Nutrients_MassBalance = Abs(Nutrients_MassBalance)
            
            ' Values are calculated per day, since we want to calculate break even points for the systems. Therefore tanks are added to the resupply values
            If InStr(1, .O2StorageList.Value, "CS", vbBinaryCompare) Then
                TankValues = calculateCryogenicTank(O2_MassBalance, "O2")
            ElseIf InStr(1, .O2StorageList.Value, "HPS", vbBinaryCompare) Then
                TankValues = calculateHighPressureTank(O2_MassBalance, "O2")
            End If
            TotalSystemResupply = TotalSystemResupply + TankValues(1) + O2_MassBalance
            TotalSystemResupplyVolume = TotalSystemResupplyVolume + TankValues(2)
            
            SystemMass = TankValues(1)
            ReliabilityValues = calcsparesORUmass(.O2StorageList.Value, SystemMass)
            TotalSystemMass = TotalSystemMass + ReliabilityValues(1)
            TotalSystemVolume = TotalSystemVolume + ReliabilityValues(2)
            
            ' Add oxygen storage mass to air revitalization ESM
            AR_ESM = AR_ESM + (TankValues(1) + TankValues(2) * ESM_VolumeF + O2_MassBalance) * MissionDuration + ReliabilityValues(1) + ReliabilityValues(2) * ESM_VolumeF
            
            If H2_MassBalance <> 0 Then
                If InStr(1, .H2StorageList.Value, "None", vbBinaryCompare) Then
                Else
                    If InStr(1, .H2StorageList.Value, "CS", vbBinaryCompare) Then
                        TankValues = calculateCryogenicTank(H2_MassBalance, "H2")
                    ElseIf InStr(1, .H2StorageList.Value, "HPS", vbBinaryCompare) Then
                        TankValues = calculateHighPressureTank(H2_MassBalance, "H2")
                    End If
                    TotalSystemResupply = TotalSystemResupply + TankValues(1) + H2_MassBalance
                    TotalSystemResupplyVolume = TotalSystemResupplyVolume + TankValues(2)
                    
                    SystemMass = TankValues(1)
                    ReliabilityValues = calcsparesORUmass(.H2StorageList.Value, SystemMass)
                    TotalSystemMass = TotalSystemMass + ReliabilityValues(1)
                    TotalSystemVolume = TotalSystemVolume + ReliabilityValues(2)
                    
                    ' Add hydrogen storage mass to air revitalization ESM
                    AR_ESM = AR_ESM + (TankValues(1) + TankValues(2) * ESM_VolumeF + H2_MassBalance) * MissionDuration + ReliabilityValues(1) + ReliabilityValues(2) * ESM_VolumeF
                    
                End If
            End If
            
            ' Only add water storage to esm if water is consumed. For surplus water no storage is assumed, as it could be vented and should not hamper a system since it is an advantage
            If H2O_MassBalance < 0 Then
                TankValues = calculateWaterTank(Abs(H2O_MassBalance))
                Dim WRM_Resupply As Double
                
                If H2O_MassBalance < 0 Then
                    WRM_Resupply = TankValues(1) + Abs(H2O_MassBalance)
                Else
                    WRM_Resupply = 0
                    TankValues(1) = 0
                    TankValues(2) = 0
                End If
                TotalSystemResupply = TotalSystemResupply + WRM_Resupply
                TotalSystemResupplyVolume = TotalSystemResupplyVolume + TankValues(2)
                SystemMass = TankValues(1)
                
                ReliabilityValues = calcsparesORUmass("Water Tank", SystemMass)
                TotalSystemMass = TotalSystemMass + ReliabilityValues(1)
                TotalSystemVolume = TotalSystemVolume + ReliabilityValues(2)
                ' Add water storage mass to water recovery ESM
                WRM_ESM = WRM_ESM + (WRM_Resupply + TankValues(2) * ESM_VolumeF) * MissionDuration + ReliabilityValues(1) + ReliabilityValues(2) * ESM_VolumeF
            End If
            
            ' Assume waste, food and nutrients require a similar storage mass and volume as water
            TankValues = calculateWaterTank(Food_MassBalance)
            TotalSystemResupply = TotalSystemResupply + TankValues(1) + Food_MassBalance
            TotalSystemResupplyVolume = TotalSystemResupplyVolume + TankValues(2)
            SystemMass = TankValues(1)
            ReliabilityValues = calcsparesORUmass("Water Tank", SystemMass)
            TotalSystemMass = TotalSystemMass + ReliabilityValues(1)
            TotalSystemVolume = TotalSystemVolume + ReliabilityValues(2)
            ' Add food storage mass to air revitalization ESM
            Food_ESM = Food_ESM + (TankValues(1) + Food_MassBalance + TankValues(2) * ESM_VolumeF) * MissionDuration + ReliabilityValues(1) + ReliabilityValues(2) * ESM_VolumeF
            
            ' Add clothes storage to ESM, see BVAD Table 4.49
            Dim ClothesMass As Double
            Dim ClothesVolume As Double
            ClothesMass = CrewSize * 0.343
            ClothesVolume = CrewSize * 0.0013
            TotalSystemResupply = TotalSystemResupply + ClothesMass
            TotalSystemResupplyVolume = TotalSystemResupplyVolume + ClothesVolume
            Cloth_ESM = Cloth_ESM + (ClothesMass + ClothesVolume * ESM_VolumeF) * MissionDuration
            Worksheets("Output Data").Cells(miResupplyRows(1), ResultsColumn) = ClothesMass
            
            If CheckBoxIncludeWasteStorage.Value Then
                TankValues = calculateWaterTank(Waste_MassBalance)
                TotalSystemResupply = TotalSystemResupply + TankValues(1)
                TotalSystemResupplyVolume = TotalSystemResupplyVolume + TankValues(2)
                SystemMass = TankValues(1)
                ReliabilityValues = calcsparesORUmass("Water Tank", SystemMass)
                TotalSystemMass = TotalSystemMass + ReliabilityValues(1)
                TotalSystemVolume = TotalSystemVolume + ReliabilityValues(2)
                ' Add waste storage mass to air revitalization ESM
                Waste_ESM = Waste_ESM + (TankValues(1) + TankValues(2) * ESM_VolumeF) * MissionDuration + ReliabilityValues(1) + ReliabilityValues(2) * ESM_VolumeF
            
            Else
                Waste_ESM = 0
            End If
            
            TankValues = calculateWaterTank(Nutrients_MassBalance)
            TotalSystemResupply = TotalSystemResupply + TankValues(1) + Nutrients_MassBalance
            TotalSystemResupplyVolume = TotalSystemResupplyVolume + TankValues(2)
            SystemMass = TankValues(1)
            If SystemMass <> 0 Then
                ReliabilityValues = calcsparesORUmass("Water Tank", SystemMass)
                TotalSystemMass = TotalSystemMass + ReliabilityValues(1)
                TotalSystemVolume = TotalSystemVolume + ReliabilityValues(2)
            Else
                ReliabilityValues(1) = 0
                ReliabilityValues(2) = 0
            End If
            ' Add nutrient storage mass to air revitalization ESM
            PGC_ESM = PGC_ESM + (TankValues(1) + Nutrients_MassBalance + TankValues(2) * ESM_VolumeF) * MissionDuration + ReliabilityValues(1) + ReliabilityValues(2) * ESM_VolumeF
            
        ' Initilize crew scaling values:
            Dim CrewValues(11) As Double
            If Schedule = 1 Then
                peak_sweat = WorksheetFunction.Max(Worksheets("Schedule and Metabolic").Range("sweat").Value)
                heat = Worksheets("Schedule and Metabolic").Range("SUM_Heat_Load").Value
                peak_heat = WorksheetFunction.Max(Worksheets("Schedule and Metabolic").Range("heat").Value)
            
            ElseIf Schedule = 2 Then
                peak_sweat = WorksheetFunction.Max(Worksheets("Schedule and Metabolic").Range("Sweat_Schedule2").Value)
                heat = Worksheets("Schedule and Metabolic").Range("SUM_Heat_Load_2").Value
                peak_heat = WorksheetFunction.Max(Worksheets("Schedule and Metabolic").Range("Heat_Schedule2").Value)
            End If
            
            ' For the composition ESM scaling calculation, not the values the crew actually consumes are relevant, but rather the values each of the systems must provide!
            ' O2 consumption
            CrewValues(1) = .Range("O2_FromElectrolysisValue").Value
            ' CO2 pRemoval
            CrewValues(2) = .Range("CO2_FromRemovalValue").Value
            ' urine prodcution
            CrewValues(3) = .Range("CrewUrineWithProcessing").Value
            ' total water to processing prodcution
            CrewValues(4) = .Range("HumControlToWWFiltrationValue").Value - .Range("HumControlToPlantsValue").Value + .Range("UrineProcessingGreyWater").Value + .Range("HygieneH2O_ProdCrewValue").Value + .Range("H2O_InediblePlantsToFiltrationValue").Value
            ' feces prodcution
            CrewValues(5) = .Range("CrewFecesValue").Value
            ' sweat prodcution
            CrewValues(6) = .Range("H2O_ProdCrewValue").Value + .Range("H2OfromPlantsValue").Value + .Range("BrinProcH2OtoCabinValue").Value
            ' peak sweat prodcution (per hour)
            CrewValues(7) = peak_sweat / CrewSize + (.Range("H2OfromPlantsValue").Value + .Range("BrinProcH2OtoCabinValue").Value) / 24
            ' heat prodcution
            CrewValues(8) = heat / CrewSize
            ' peak heat prodcution
            CrewValues(9) = peak_heat / CrewSize
            ' potable water consumption
            CrewValues(10) = .Range("H2O_ConCrewValue").Value
            ' food consumption
            CrewValues(11) = .Range("FoodConCrewValue").Value
            
            For i = 1 To 11
                 CrewValues(i) = CrewValues(i) * CrewSize
            Next i
            
            ' Humidity Control
            Dim csSystemAbbreviations(1 To 7) As String
            csSystemAbbreviations(1) = .HumidityControlList.Value
            csSystemAbbreviations(2) = .CO2_RemovalList.Value
            csSystemAbbreviations(3) = .ElectrolysisList.Value
            csSystemAbbreviations(4) = .CO2ReductionList.Value
            csSystemAbbreviations(5) = .UrineProcessingList.Value
            csSystemAbbreviations(6) = .WWFiltrationList.Value
            csSystemAbbreviations(7) = .BrineProcessingList.Value
            
            ' Set the rows in which the selected systems are stored on the output data sheet, note order must match the abbreviation set above:
            Dim miSystemRows
            Dim miResupplyRowsSystems
            
            miSystemRows = findRowByName(Array("Humidity Control System:", "CO2 Removal System:", "Electrolysis System:", "CO2 Reduction System:", "Urine Processing System:", "WW Filtration System:", "Brine Processing System:"), "Output Data", 1)
            
            miResupplyRowsSystems = findRowByName(Array("CO2 removal", "Electrolysis", "CO2 reduct.", "Urine Proc.", "WW filt", "Brine Proc."), "Output Data", 1)
            
            Dim TechnologyValues
            Dim iSystem As Integer
            Dim SystemESM As Double
            
            For iSystem = 1 To UBound(csSystemAbbreviations)
                ' Set the selected system to the output:
                Worksheets("Output Data").Cells(miSystemRows(iSystem), ResultsColumn).Value = csSystemAbbreviations(iSystem)
                
                If csSystemAbbreviations(iSystem) = "None" Then
                    Worksheets("Output Data").Cells(miSystemRows(iSystem) + 1, ResultsColumn).Value = 0
                    GoTo NextIteration
                    If iSystem > 1 Then
                        Worksheets("Output Data").Cells(miResupplyRowsSystems(iSystem - 1), ResultsColumn) = 0
                    End If
                Else
                    If iSystem = 2 And csSystemAbbreviations(iSystem) = .HumidityControlList.Value Then
                        GoTo NextIteration
                    ElseIf iSystem = 4 Then
                        ' For the co2 reduction, adjust the value with the CO2 actually supplied to it:
                        CrewValues(2) = Worksheets("ECLSS Composition").Range("CO2_ToReductionValue").Value * CrewSize
                    End If
                    
                    '"Mass", "Volume", "Heat", "Power", "Resupply Mass", "Maintenance"
                    TechnologyValues = getTechnologyValues(csSystemAbbreviations(iSystem), CrewValues)
            
                    Worksheets("Output Data").Cells(miSystemRows(iSystem) + 1, ResultsColumn).Value = TechnologyValues(4)
                    
                    SystemMass = TechnologyValues(1)
                    
                    TotalSystemMass = TotalSystemMass + SystemMass
                    TotalSystemVolume = TotalSystemVolume + TechnologyValues(2)
                    TotalSystemCooling = TotalSystemCooling + TechnologyValues(3)
                    TotalSystemPower = TotalSystemPower + TechnologyValues(4)
                    TotalSystemResupply = TotalSystemResupply + TechnologyValues(5) / MissionDuration ' resupply is given as total resupply, but we want the kg/d value here
                    TotalSystemMaintenance = TotalSystemMaintenance + TechnologyValues(6)
            
                    If iSystem > 1 Then
                        Worksheets("Output Data").Cells(miResupplyRowsSystems(iSystem - 1), ResultsColumn) = TechnologyValues(5) / MissionDuration
                    End If
            
                    '"Mass for Spares/ORU", "Volume for Spares/ORU", "Reliability without Spares", "Reliability"
                    ReliabilityValues = calcsparesORUmass(csSystemAbbreviations(iSystem), SystemMass)
                
                    TotalSpareMass = TotalSpareMass + ReliabilityValues(1)
                    TotalSpareVolume = TotalSpareVolume + ReliabilityValues(2)
                    
                    SystemESM = TechnologyValues(1) + (TechnologyValues(2) + ReliabilityValues(2)) * ESM_VolumeF + TechnologyValues(3) * ESM_CoolingF + TechnologyValues(4) * ESM_PowerF + TechnologyValues(5) + TotalSystemMaintenance * ESM_Crew_TimeF + ReliabilityValues(1)
                    SystemESM = ESM_LocationF * SystemESM
                    Select Case iSystem    ' Evaluate Number.
                    Case 1
                        THC_ESM = THC_ESM + SystemESM
                    Case 2
                        ACS_ESM = ACS_ESM + SystemESM
                    Case 3 To 4
                        AR_ESM = AR_ESM + SystemESM
                    Case 5 To 7
                        WRM_ESM = WRM_ESM + SystemESM
                    End Select
                End If
NextIteration:
            Next iSystem
            
            ' And the PGC are a special case, as for them the area is most important (or the values are fix and multiplied with the number of units)
            Dim miTechnologyColumn
            Dim miTechnologyRows
            Dim PGC_ResupplyMass As Double
            
            miTechnologyColumn = findColumnByName(Array("Name", "Abbreviation", "Mass", "Volume", "Heat", "Power", "Resupply Mass", "fix Maintenance", "Maintenance per CM", "Efficiency"), "Technology", 1)
            
            If .PlantGrowthChamberList.Value = "None" Then
                PGC_ESM = 0
                miSystemRows = findRowByName(Array("Plant Growth System:"), "Output Data", 1)
                Worksheets("Output Data").Cells(miSystemRows(1), ResultsColumn).Value = .PlantGrowthChamberList.Value
                Worksheets("Output Data").Cells(miSystemRows(1) + 1, ResultsColumn).Value = 0
            Else
                miTechnologyRows = findRowByName(Array(.PlantGrowthChamberList.Value, "IMV", "PCA", .AtmosphereMonitoringList.Value, .TraceContaminantControlList.Value), "Technology", miTechnologyColumn(2))
                iTechnologyRow = miTechnologyRows(1)
                
                miSystemRows = findRowByName(Array("Plant Growth System:"), "Output Data", 1)
                Worksheets("Output Data").Cells(miSystemRows(1), ResultsColumn).Value = .PlantGrowthChamberList.Value
                
                If .PlantGrowthChamberList.Value = "Veggie" Or .PlantGrowthChamberList.Value = "ISPR PGC" Then
                
                    Dim iNumberOfPGC As Integer
                    
                    iNumberOfPGC = CInt(.PGC_NumberOfUnits.Value)
                    With Worksheets("Technology")
                        TotalSystemMass = TotalSystemMass + iNumberOfPGC * CrewSize * .Cells(iTechnologyRow, miTechnologyColumn(3)).Value
                        TotalSystemVolume = TotalSystemVolume + iNumberOfPGC * CrewSize * .Cells(iTechnologyRow, miTechnologyColumn(4)).Value
                        TotalSystemCooling = TotalSystemCooling + iNumberOfPGC * CrewSize * .Cells(iTechnologyRow, miTechnologyColumn(5)).Value
                        TotalSystemPower = TotalSystemPower + iNumberOfPGC * CrewSize * .Cells(iTechnologyRow, miTechnologyColumn(6)).Value
                        TotalSystemResupply = TotalSystemResupply + iNumberOfPGC * CrewSize * .Cells(iTechnologyRow, miTechnologyColumn(7)).Value
                        TotalSystemMaintenance = TotalSystemMaintenance + iNumberOfPGC * CrewSize * .Cells(iTechnologyRow, miTechnologyColumn(8)).Value
                        
                        Worksheets("Output Data").Cells(miSystemRows(1) + 1, ResultsColumn).Value = iNumberOfPGC * CrewSize * .Cells(iTechnologyRow, miTechnologyColumn(6)).Value
                        
                        SystemESM = iNumberOfPGC * CrewSize * (.Cells(iTechnologyRow, miTechnologyColumn(3)).Value + .Cells(iTechnologyRow, miTechnologyColumn(4)).Value * ESM_VolumeF + .Cells(iTechnologyRow, miTechnologyColumn(5)).Value * ESM_CoolingF + .Cells(iTechnologyRow, miTechnologyColumn(6)).Value * ESM_PowerF + .Cells(iTechnologyRow, miTechnologyColumn(7)).Value * MissionDuration + .Cells(iTechnologyRow, miTechnologyColumn(8)).Value * ESM_Crew_TimeF)
                        PGC_ESM = PGC_ESM + SystemESM
                        
                        PGC_ResupplyMass = iNumberOfPGC * CrewSize * .Cells(iTechnologyRow, miTechnologyColumn(7)).Value
                    End With
                Else
                    ' now get the values and scale them with the "efficiency" column value, which is the growth area for PGC
                    Dim ScalingFactorPGC As Double
                    
                    ScalingFactorPGC = CrewSize * Worksheets("ECLSS Composition").Range("CurrentPGCArea").Value / Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumn(10)).Value
                    
                    With Worksheets("Technology")
                        TotalSystemMass = TotalSystemMass + ScalingFactorPGC * .Cells(iTechnologyRow, miTechnologyColumn(3)).Value
                        TotalSystemVolume = TotalSystemVolume + ScalingFactorPGC * .Cells(iTechnologyRow, miTechnologyColumn(4)).Value
                        TotalSystemCooling = TotalSystemCooling + ScalingFactorPGC * .Cells(iTechnologyRow, miTechnologyColumn(5)).Value
                        TotalSystemPower = TotalSystemPower + ScalingFactorPGC * .Cells(iTechnologyRow, miTechnologyColumn(6)).Value
                        TotalSystemResupply = TotalSystemResupply + ScalingFactorPGC * .Cells(iTechnologyRow, miTechnologyColumn(7)).Value
                        TotalSystemMaintenance = TotalSystemMaintenance + ScalingFactorPGC * .Cells(iTechnologyRow, miTechnologyColumn(9)).Value
                        
                        Worksheets("Output Data").Cells(miSystemRows(1) + 1, ResultsColumn).Value = ScalingFactorPGC * .Cells(iTechnologyRow, miTechnologyColumn(6)).Value
                        
                        SystemESM = ScalingFactorPGC * (.Cells(iTechnologyRow, miTechnologyColumn(3)).Value + .Cells(iTechnologyRow, miTechnologyColumn(4)).Value * ESM_VolumeF + .Cells(iTechnologyRow, miTechnologyColumn(5)).Value * ESM_CoolingF + .Cells(iTechnologyRow, miTechnologyColumn(6)).Value * ESM_PowerF + .Cells(iTechnologyRow, miTechnologyColumn(7)).Value * MissionDuration + .Cells(iTechnologyRow, miTechnologyColumn(8)).Value * ESM_Crew_TimeF)
                        PGC_ESM = PGC_ESM + SystemESM
                        
                        PGC_ResupplyMass = ScalingFactorPGC * .Cells(iTechnologyRow, miTechnologyColumn(7)).Value
                    End With
                    
                End If
                
            End If
             Worksheets("Output Data").Cells(miResupplyRows(8), ResultsColumn).Value = PGC_ResupplyMass
                
                            ' And the PBR is a special case
            Dim iPBRScaling As Double
            Dim PBR_ResupplyMass As Double
            
            miTechnologyColumn = findColumnByName(Array("Name", "Abbreviation", "Mass", "Volume", "Heat", "Power", "Resupply Mass", "fix Maintenance", "Maintenance per CM", "Efficiency"), "Technology", 1)
            
            If .PBRList.Value = "None" Then
                PBR_ESM = 0
                miSystemRows = findRowByName(Array("PBR:"), "Output Data", 1)
                Worksheets("Output Data").Cells(miSystemRows(1), ResultsColumn).Value = .PBRList.Value
                Worksheets("Output Data").Cells(miSystemRows(1) + 1, ResultsColumn).Value = 0
                Worksheets("Output Data").Cells(miSystemRows(1) + 2, ResultsColumn).Value = 0
                Worksheets("Output Data").Cells(miSystemRows(1) + 3, ResultsColumn).Value = 0
            Else
                miTechnologyRows = findRowByName(Array(.PBRList.Value), "Technology", miTechnologyColumn(2))
                iTechnologyRow = miTechnologyRows(1)

                miSystemRows = findRowByName(Array("PBR:"), "Output Data", 1)
                Worksheets("Output Data").Cells(miSystemRows(1), ResultsColumn).Value = .PBRList.Value

                    'Data is for all CO2 for Crew of 1 -> Scaling
                    iPBRScaling = (Worksheets("ECLSS Composition").PBR_CO2Percentage.Value / 100) * CrewSize
                    With Worksheets("Technology")
                        TotalSystemMass = TotalSystemMass + iPBRScaling * .Cells(iTechnologyRow, miTechnologyColumn(3)).Value
                        TotalSystemVolume = TotalSystemVolume + iPBRScaling * .Cells(iTechnologyRow, miTechnologyColumn(4)).Value
                        TotalSystemCooling = TotalSystemCooling + iPBRScaling * .Cells(iTechnologyRow, miTechnologyColumn(5)).Value
                        TotalSystemPower = TotalSystemPower + iPBRScaling * .Cells(iTechnologyRow, miTechnologyColumn(6)).Value
                        TotalSystemResupply = TotalSystemResupply + iPBRScaling * .Cells(iTechnologyRow, miTechnologyColumn(7)).Value
                        TotalSystemMaintenance = TotalSystemMaintenance + iPBRScaling * .Cells(iTechnologyRow, miTechnologyColumn(8)).Value

                        Worksheets("Output Data").Cells(miSystemRows(1) + 1, ResultsColumn).Value = iPBRScaling * .Cells(iTechnologyRow, miTechnologyColumn(6)).Value

                        SystemESM = iPBRScaling * (.Cells(iTechnologyRow, miTechnologyColumn(3)).Value + .Cells(iTechnologyRow, miTechnologyColumn(4)).Value * ESM_VolumeF + .Cells(iTechnologyRow, miTechnologyColumn(5)).Value * ESM_CoolingF + .Cells(iTechnologyRow, miTechnologyColumn(6)).Value * ESM_PowerF + .Cells(iTechnologyRow, miTechnologyColumn(7)).Value * MissionDuration + .Cells(iTechnologyRow, miTechnologyColumn(8)).Value * ESM_Crew_TimeF)
                        PBR_ESM = PBR_ESM + SystemESM
                        PBR_ResupplyMass = iPBRScaling * .Cells(iTechnologyRow, miTechnologyColumn(7)).Value
                    End With
                
            End If
            Worksheets("Output Data").Cells(miResupplyRows(9), ResultsColumn).Value = PBR_ResupplyMass
            
            
              ' And the CH4-Processing is a special case (Only scales with CH4)
            Dim iCH4Scaling As Double
            Dim CH4_ResupplyMass As Double
            
            miTechnologyColumn = findColumnByName(Array("Name", "Abbreviation", "Mass", "Volume", "Heat", "Power", "Resupply Mass", "fix Maintenance", "Maintenance per CM", "Efficiency"), "Technology", 1)
            
            If .MethaneProcessingList.Value = "None" Then
                CH4_ESM = 0
                miSystemRows = findRowByName(Array("CH4 Processing System:"), "Output Data", 1)
                Worksheets("Output Data").Cells(miSystemRows(1), ResultsColumn).Value = .MethaneProcessingList.Value
                Worksheets("Output Data").Cells(miSystemRows(1) + 1, ResultsColumn).Value = 0
                Worksheets("Output Data").Cells(miSystemRows(1) + 2, ResultsColumn).Value = 0
            Else
                miTechnologyRows = findRowByName(Array(.MethaneProcessingList.Value), "Technology", miTechnologyColumn(2))
                iTechnologyRow = miTechnologyRows(1)

                miSystemRows = findRowByName(Array("CH4 Processing System:"), "Output Data", 1)
                Worksheets("Output Data").Cells(miSystemRows(1), ResultsColumn).Value = .MethaneProcessingList.Value

                    'Data is for 1 CM -> Scaling
                    iCH4Scaling = CrewSize
                    With Worksheets("Technology")
                        TotalSystemMass = TotalSystemMass + iCH4Scaling * .Cells(iTechnologyRow, miTechnologyColumn(3)).Value
                        TotalSystemVolume = TotalSystemVolume + iCH4Scaling * .Cells(iTechnologyRow, miTechnologyColumn(4)).Value
                        TotalSystemCooling = TotalSystemCooling + iCH4Scaling * .Cells(iTechnologyRow, miTechnologyColumn(5)).Value
                        TotalSystemPower = TotalSystemPower + iCH4Scaling * .Cells(iTechnologyRow, miTechnologyColumn(6)).Value
                        TotalSystemResupply = TotalSystemResupply + iCH4Scaling * .Cells(iTechnologyRow, miTechnologyColumn(7)).Value
                        TotalSystemMaintenance = TotalSystemMaintenance + iCH4Scaling * .Cells(iTechnologyRow, miTechnologyColumn(8)).Value

                        Worksheets("Output Data").Cells(miSystemRows(1) + 1, ResultsColumn).Value = iCH4Scaling * .Cells(iTechnologyRow, miTechnologyColumn(6)).Value

                        SystemESM = iCH4Scaling * (.Cells(iTechnologyRow, miTechnologyColumn(3)).Value + .Cells(iTechnologyRow, miTechnologyColumn(4)).Value * ESM_VolumeF + .Cells(iTechnologyRow, miTechnologyColumn(5)).Value * ESM_CoolingF + .Cells(iTechnologyRow, miTechnologyColumn(6)).Value * ESM_PowerF + .Cells(iTechnologyRow, miTechnologyColumn(7)).Value * MissionDuration + .Cells(iTechnologyRow, miTechnologyColumn(8)).Value * ESM_Crew_TimeF)
                        CH4_ESM = CH4_ESM + SystemESM
                        CH4_ResupplyMass = iCH4Scaling * .Cells(iTechnologyRow, miTechnologyColumn(7)).Value
                    End With
                
            End If
            Worksheets("Output Data").Cells(miResupplyRows(10), ResultsColumn).Value = CH4_ResupplyMass
            
            
       End With
        
        ' Calculate total system ESM
        FixESM = ESM_LocationF * (TotalSystemMass + ESM_VolumeF * TotalSystemVolume + ESM_CoolingF * TotalSystemCooling + ESM_PowerF * TotalSystemPower + TotalSystemResupply + ESM_Crew_TimeF * TotalSystemMaintenance + TotalSpareMass + ESM_VolumeF * TotalSpareVolume)
        ScalingESM = ESM_LocationF * (TotalSystemResupply + ESM_VolumeF * TotalSystemResupplyVolume)
        
        
        Dim iModules As Integer
        
        iModules = Worksheets("User Interface").Range("Amount_of_Modules").Value
        With Worksheets("Technology")
            miTechnologyRows = findRowByName(Array("IMV", "PCA", Worksheets("ECLSS Composition").AtmosphereMonitoringList.Value, Worksheets("ECLSS Composition").TraceContaminantControlList.Value), "Technology", miTechnologyColumn(2))
            
            ' add intermodular ventilation (IMV) per module
            iTechnologyRow = miTechnologyRows(1)
            SystemESM = iModules * (.Cells(iTechnologyRow, miTechnologyColumn(3)).Value + .Cells(iTechnologyRow, miTechnologyColumn(4)).Value * ESM_VolumeF + .Cells(iTechnologyRow, miTechnologyColumn(5)).Value * ESM_CoolingF + .Cells(iTechnologyRow, miTechnologyColumn(6)).Value * ESM_PowerF + .Cells(iTechnologyRow, miTechnologyColumn(7)).Value * MissionDuration + .Cells(iTechnologyRow, miTechnologyColumn(8)).Value * ESM_Crew_TimeF)
            ACS_ESM = ACS_ESM + SystemESM
            
            FixESM = FixESM + SystemESM
            
            ' add pressure control assemblies (PCA) per module
            iTechnologyRow = miTechnologyRows(2)
            SystemESM = iModules * (.Cells(iTechnologyRow, miTechnologyColumn(3)).Value + .Cells(iTechnologyRow, miTechnologyColumn(4)).Value * ESM_VolumeF + .Cells(iTechnologyRow, miTechnologyColumn(5)).Value * ESM_CoolingF + .Cells(iTechnologyRow, miTechnologyColumn(6)).Value * ESM_PowerF + .Cells(iTechnologyRow, miTechnologyColumn(7)).Value * MissionDuration + .Cells(iTechnologyRow, miTechnologyColumn(8)).Value * ESM_Crew_TimeF)
            ACS_ESM = ACS_ESM + SystemESM
            
            FixESM = FixESM + SystemESM
            
            ' add atmosphere monitoring
            iTechnologyRow = miTechnologyRows(3)
            SystemESM = iModules * (.Cells(iTechnologyRow, miTechnologyColumn(3)).Value + .Cells(iTechnologyRow, miTechnologyColumn(4)).Value * ESM_VolumeF + .Cells(iTechnologyRow, miTechnologyColumn(5)).Value * ESM_CoolingF + .Cells(iTechnologyRow, miTechnologyColumn(6)).Value * ESM_PowerF + .Cells(iTechnologyRow, miTechnologyColumn(7)).Value * MissionDuration + .Cells(iTechnologyRow, miTechnologyColumn(8)).Value * ESM_Crew_TimeF)
            ACS_ESM = ACS_ESM + SystemESM
            
            FixESM = FixESM + SystemESM
            
            ' add trace contaminant control
            iTechnologyRow = miTechnologyRows(4)
            SystemESM = iModules * (.Cells(iTechnologyRow, miTechnologyColumn(3)).Value + .Cells(iTechnologyRow, miTechnologyColumn(4)).Value * ESM_VolumeF + .Cells(iTechnologyRow, miTechnologyColumn(5)).Value * ESM_CoolingF + .Cells(iTechnologyRow, miTechnologyColumn(6)).Value * ESM_PowerF + .Cells(iTechnologyRow, miTechnologyColumn(7)).Value * MissionDuration + .Cells(iTechnologyRow, miTechnologyColumn(8)).Value * ESM_Crew_TimeF)
            ACS_ESM = ACS_ESM + SystemESM
            
            FixESM = FixESM + SystemESM
        End With
        
        '------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add data to ESM output data and add other selected systems to output data
        miSystemRows = findRowByName(Array("H2O Storage:", "O2 Storage:", "H2 Storage:", "Atmosphere Monitoring:", "Trace Contaminant Control:", "Schedule Number:", "Crew Size:"), "Output Data", 1)
        
        .Cells(miSystemRows(1), ResultsColumn).Value = "Water Tank"
        .Cells(miSystemRows(2), ResultsColumn).Value = Worksheets("ECLSS Composition").O2StorageList.Value
        .Cells(miSystemRows(3), ResultsColumn).Value = Worksheets("ECLSS Composition").H2StorageList.Value
        .Cells(miSystemRows(4), ResultsColumn).Value = Worksheets("ECLSS Composition").AtmosphereMonitoringList.Value
        .Cells(miSystemRows(5), ResultsColumn).Value = Worksheets("ECLSS Composition").TraceContaminantControlList.Value
        .Cells(miSystemRows(6), ResultsColumn).Value = Worksheets("User Interface").Range("Schedule_Number").Value
        .Cells(miSystemRows(7), ResultsColumn).Value = Worksheets("User Interface").Range("Crew_Size").Value
        
        .Cells(45, ResultsColumn).Value = FixESM + ScalingESM * MissionDuration
        .Cells(46, ResultsColumn).Value = ScalingESM
        .Cells(47, ResultsColumn).Value = ACS_ESM
        .Cells(48, ResultsColumn).Value = AR_ESM
        .Cells(49, ResultsColumn).Value = WRM_ESM
        .Cells(50, ResultsColumn).Value = THC_ESM
        .Cells(51, ResultsColumn).Value = PGC_ESM
        .Cells(52, ResultsColumn).Value = PBR_ESM
        .Cells(53, ResultsColumn).Value = CH4_ESM
        .Cells(54, ResultsColumn).Value = Food_ESM
        .Cells(55, ResultsColumn).Value = Cloth_ESM
        .Cells(56, ResultsColumn).Value = Waste_ESM
        
        .Cells(87, ResultsColumn).Value = TotalSystemMass
        .Cells(88, ResultsColumn).Value = TotalSystemVolume
        .Cells(89, ResultsColumn).Value = TotalSpareMass
        .Cells(90, ResultsColumn).Value = TotalSpareVolume
        
        '------------------------------------------------------------------------------------------------------------------------------------------------
        ' add the kg/d processing values for the subsystems to the output data sheet
        .Cells(10, ResultsColumn).Value = Worksheets("ECLSS Composition").Range("CO2_FromRemovalValue").Value ' CO2 removal CO2
        .Cells(13, ResultsColumn).Value = Worksheets("ECLSS Composition").Range("H2O_FromCO2ReductionValue").Value ' CO2 reduction H2O
        .Cells(14, ResultsColumn).Value = Worksheets("ECLSS Composition").Range("O2_FromCO2ReductionValue").Value ' CO2 reduction O2
        .Cells(17, ResultsColumn).Value = Worksheets("ECLSS Composition").Range("HumControlToWWFiltrationValue").Value + Worksheets("ECLSS Composition").Range("HumControlToWasteValue").Value  ' Humidity Control H2O
        .Cells(20, ResultsColumn).Value = Worksheets("ECLSS Composition").Range("H2O_FromWWFiltValue").Value ' WW proc H2O
        .Cells(23, ResultsColumn).Value = Worksheets("ECLSS Composition").Range("UrineProcessingGreyWater").Value ' Urine proc H2O
        .Cells(26, ResultsColumn).Value = Worksheets("ECLSS Composition").Range("BrinProcH2OtoCabinValue").Value ' Brine proc. H2O
        .Cells(29, ResultsColumn).Value = Worksheets("ECLSS Composition").Range("O2_FromElectrolysisValue").Value ' ELY O2
        .Cells(32, ResultsColumn).Value = Worksheets("ECLSS Composition").Range("EdibleFromPlantsValue").Value ' PGC Edible
        .Cells(35, ResultsColumn).Value = Worksheets("ECLSS Composition").Range("O2FromPBRValue").Value ' PGC Edible
        .Cells(36, ResultsColumn).Value = Worksheets("ECLSS Composition").Range("FoodFromPBRValue").Value ' PGC Edible
        .Cells(39, ResultsColumn).Value = Worksheets("ECLSS Composition").Range("H2_FromMethaneProcessingValue").Value ' PGC Edible
        '------------------------------------------------------------------------------------------------------------------------------------------------
        ' create ESM per mission duration plot for breakeven points:
        
        Dim iRow As Integer
        Dim iEndRow As Integer
        Dim MissionDays As Long
        MissionDays = 0
        Dim iMissionEndRow As Integer
        iMissionEndRow = 0
        For iRow = iResultsHeaderRow2 + 2 To iResultsHeaderRow2 + 1002
            If iRow - (iResultsHeaderRow2 + 2) < iResultsHeaderRow2 + 2 + 101 Then
                MissionDays = MissionDays + 1
            ElseIf iRow - (iResultsHeaderRow2 + 2) < iResultsHeaderRow2 + 2 + 201 Then
                MissionDays = MissionDays + 10
            Else
                MissionDays = MissionDays + 100
            End If
            
            .Cells(iRow, 2).Value = MissionDays
            
            If MissionDays > MissionDuration And iMissionEndRow = 0 Then
                iMissionEndRow = iRow
            End If
            
            .Cells(iRow, ResultsColumn).Value = FixESM + MissionDays * ScalingESM
            
        Next iRow
        iEndRow = iResultsHeaderRow2 + 1002
        
        '------------------------------------------------------------------------------------------------------------------------------------------------
        ' Add data to ESM output plots
        
        ' set the ESM / Mission duration plot
        Set ChtObj = ThisWorkbook.Worksheets("Graphical Output").ChartObjects("ESM System Mass")
        
        ' Set the Range of the Chart's source data
        Set ChtRng = .Range(.Cells(iResultsHeaderRow2 + 2, ResultsColumn), .Cells(iEndRow, ResultsColumn))
        Set ChtRngX = .Range(.Cells(iResultsHeaderRow2 + 2, 2), .Cells(iEndRow, 2))
        
        Dim bSeriesExist As Boolean
        ' check if this series already exists, if so only change the values do not create a new series
        bSeriesExist = False
        For Each Series In ChtObj.Chart.SeriesCollection
            If Series.Name = .Cells(iResultsHeaderRow2 + 1, ResultsColumn) Then
                bSeriesExist = True
                Set Ser = Series
                Exit For
            End If
        Next
        
        If bSeriesExist = False Then
            ' add a new series to chart
            Set Ser = ChtObj.Chart.SeriesCollection.NewSeries
            Ser.Name = .Cells(iResultsHeaderRow2 + 1, ResultsColumn)
        End If
        
        ' set the source data of the new series
        Ser.Values = "=" & ChtRng.Address(False, False, xlA1, xlExternal)
        Ser.XValues = "=" & ChtRngX.Address(False, False, xlA1, xlExternal)
        
        ' Set the X Axis limit to the currently selected mission duration
        ChtObj.Chart.Axes(xlCategory).MinimumScale = 0
        ChtObj.Chart.Axes(xlCategory).MaximumScale = MissionDuration
        ' Set the Y Axis limit to the maximum from all results:
        ChtObj.Chart.Axes(xlValue).MinimumScale = 0
        ChtObj.Chart.Axes(xlValue).MaximumScale = WorksheetFunction.Max(.Range(.Cells(iMissionEndRow, 3), .Cells(iMissionEndRow, 12)))
        
        ' since it is possible that the user renamed a previously exiting series, we loop through the series in the chart and check if they still exist. If not they are deleted
        For Each Series In ChtObj.Chart.SeriesCollection
            If IsError(Application.Match(Series.Name, Array(.Range(.Cells(iResultsHeaderRow2 + 1, 2), .Cells(iResultsHeaderRow2 + 1, 11))), 0)) Then
                Series.Delete
            End If
        Next
        
    End With
        
    ThisWorkbook.Worksheets("Graphical Output").createESMColumnPlot
    
    ThisWorkbook.Worksheets("Output Data").findBreakEvenPoints
    
    Unload AddCompositionResults
    
End Sub

Private Sub Cancel_Click()
    
    Unload AddCompositionResults
    
End Sub

Private Sub UserForm_Initialize()

Dim i As Integer
Dim bFoundFirstEmptyResult As Boolean
bFoundFirstEmptyResult = False
i = 1
AddCompositionResults.Results.Value = Worksheets("Output Data").Cells(6, 2 + i).Value
For i = 1 To 10


    If IsEmpty(Worksheets("Output Data").Cells(7, 2 + i).Value) And bFoundFirstEmptyResult = False Then
        AddCompositionResults.Results.AddItem Worksheets("Output Data").Cells(6, 2 + i).Value
        AddCompositionResults.Results.Value = Worksheets("Output Data").Cells(6, 2 + i).Value
        bFoundFirstEmptyResult = True
    ElseIf bFoundFirstEmptyResult = False Then
        AddCompositionResults.Results.AddItem Worksheets("Output Data").Cells(7, 2 + i).Value
    Else
        AddCompositionResults.Results.AddItem Worksheets("Output Data").Cells(6, 2 + i).Value
    End If
Next i


End Sub
