Attribute VB_Name = "Modul1"
' global variable which stores the conversion of rows from the mission specific technology sheet, to the rows of the technology sheet
Option Explicit

Dim MissionSpecificRowToTechnologyRow(1 To 300) As Integer


' Execution order of these subs is:
' 1. calcsystemprops
' 2. calcSparesAndORUMass
' 3. esmcalc


Sub calcsystemprops()
'Approach: Based on a given system in- or output and the metabolic factors of a human being (according  to NASA BVAD) a rescale factor is claculated
'with which the literature values for mass, power, cooling, volume and resupply of a system can be adjusted to a desired crew size and mission duration
'In addition, there a scaling factor for each system since it is unlikely that system attrributes scale purely linearly over a wide range of crew sizes.
'This factor can be individually adjusted for each system in this module. Since the scaling approach is based on assumptiopns, contigency masses and power
'are added based on the TRL of a system (eg. lower contingency values are added if the system posseses a higher TRL)

'Load attributes from User Interface
Dim Crew_Size As Integer
Dim mission_duration As Double
Dim Schedule As Integer
Dim Margin As Integer
Dim Pressurized_volume As Double

Crew_Size = Worksheets("User Interface").Range("Crew_Size").Value
mission_duration = Worksheets("User Interface").Range("Mission_Duration").Value
Schedule = Worksheets("User Interface").Range("Schedule_Number").Value
Margin = Worksheets("User Interface").Range("Mass_margin").Value
Pressurized_volume = Worksheets("User Interface").Range("Pressurized_Volume").Value '[m3]

' Values based on schedule from LiSTOT
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
     
End If
H2O_Con = (potable + hygiene) * mission_duration
H2O_Pro = (hygiene + sweat + urine) * mission_duration

' Clear the mission specific values
Dim iRows
Dim iColumn
Dim cellRange
iColumn = Worksheets("MissionSpecificTechnologyValues").Columns.Count
iRows = Worksheets("MissionSpecificTechnologyValues").Rows.Count

Set cellRange = Worksheets("MissionSpecificTechnologyValues").Range(Worksheets("MissionSpecificTechnologyValues").Cells(3, 1), Worksheets("MissionSpecificTechnologyValues").Cells(iRows, iColumn))
cellRange.Clear

'---------------------------------------------------------------------------------------------------------------------
' In this section we set some variables to identify the correct columns to use without using static references

Dim TechnologyHeadersToCopy
Dim MissionSpecificHeadersToCopy
Dim miTechnologyColumnToCopy
Dim miMissionSpecificColumnToCopy
Dim iNameColumn As Integer
Dim MissionSpecificHeadersToCalculate
Dim miMissionSpecificColumn

TechnologyHeadersToCopy = Array("Name", "Abbreviation", "Subsystem", "Type", "Uses", "Level")
MissionSpecificHeadersToCopy = Array("Name", "Abbreviation", "Subsystem", "Type", "Uses", "Level")
miTechnologyColumnToCopy = findColumnByName(TechnologyHeadersToCopy, "Technology", 1)
miMissionSpecificColumnToCopy = findColumnByName(MissionSpecificHeadersToCopy, "MissionSpecificTechnologyValues", 1)

iNameColumn = miTechnologyColumnToCopy(1)

MissionSpecificHeadersToCalculate = Array("Mass", "Volume", "Heat", "Power", "Resupply Mass", "Maintenance")
miMissionSpecificColumn = findColumnByName(MissionSpecificHeadersToCalculate, "MissionSpecificTechnologyValues", 1)


Dim ReliabilityHeaders
Dim miMissionSpecificReliabilityColumns
Dim iMassColumn As Integer
Dim iSpareVolumeColumn As Integer
Dim iSpareMassColumn As Integer
Dim iReliabilityNoSparesColumn As Integer
Dim iReliabilityColumn As Integer

ReliabilityHeaders = Array("Mass for Spares/ORU", "Volume for Spares/ORU", "Reliability without Spares", "Reliability")
miMissionSpecificReliabilityColumns = findColumnByName(ReliabilityHeaders, "MissionSpecificTechnologyValues", 1)
iSpareMassColumn = miMissionSpecificReliabilityColumns(1)
iSpareVolumeColumn = miMissionSpecificReliabilityColumns(2)
iReliabilityNoSparesColumn = miMissionSpecificReliabilityColumns(3)
iReliabilityColumn = miMissionSpecificReliabilityColumns(4)

'---------------------------------------------------------------------------------------------------------------------
' In this section the storage systems are calculated, as their scaling is different from the other systems

Dim O2_leakage As Double
Dim N2_leakage As Double
Dim O2_decompression As Double
Dim N2_decompression As Double
Dim O2MassForCrew As Double
Dim TankValues
Dim MassTankForCrew_HPS As Double
Dim VolumeTankForCrew_HPS As Double
Dim MassTankForCrew_CS As Double
Dim VolumeTankForCrew_CS As Double
Dim MassTankForDecompression_HPS As Double
Dim VolumeTankForDecompression_HPS As Double
Dim MassTankForDecompression_CS As Double
Dim VolumeTankForDecompression_CS As Double

TankValues = calculateLeakage
O2_leakage = TankValues(1) * mission_duration
N2_leakage = TankValues(2) * mission_duration

Dim MM_O2 As Double
Dim MM_N2 As Double
Dim r As Double
Dim pp_O2 As Double
Dim pp_N2 As Double

MM_N2 = 28.0134 / 1000
MM_O2 = 31.9988 / 1000
r = 8.314472 '[J/K*mol] 'gas constant
pp_O2 = 21.3 * 10 ^ 3 '[Pa] 'partial pressure O2
pp_N2 = 79.76 * 10 ^ 3 '[Pa] 'partial pressure N2

O2_decompression = (pp_O2 / ((r / MM_O2) * 293)) * Pressurized_volume
N2_decompression = (pp_N2 / ((r / MM_N2) * 293)) * Pressurized_volume

' Now we calculate the required tank masses and volume
O2MassForCrew = (O2_leakage + O2_Con * Crew_Size) * mission_duration

TankValues = calculateHighPressureTank(O2MassForCrew, "O2")
MassTankForCrew_HPS = TankValues(1)
VolumeTankForCrew_HPS = TankValues(2)

TankValues = calculateHighPressureTank((N2_leakage * mission_duration), "N2")
MassTankForCrew_HPS = MassTankForCrew_HPS + TankValues(1)
VolumeTankForCrew_HPS = VolumeTankForCrew_HPS + TankValues(2)

TankValues = calculateCryogenicTank(O2MassForCrew, "O2")
MassTankForCrew_CS = TankValues(1)
VolumeTankForCrew_CS = TankValues(2)

TankValues = calculateCryogenicTank((N2_leakage * mission_duration), "N2")
MassTankForCrew_CS = MassTankForCrew_CS + TankValues(1)
VolumeTankForCrew_CS = VolumeTankForCrew_CS + TankValues(2)

TankValues = calculateHighPressureTank(O2_decompression, "O2")
MassTankForDecompression_HPS = TankValues(1)
VolumeTankForDecompression_HPS = TankValues(2)

TankValues = calculateHighPressureTank(N2_decompression, "N2")
MassTankForDecompression_HPS = MassTankForDecompression_HPS + TankValues(1)
VolumeTankForDecompression_HPS = VolumeTankForDecompression_HPS + TankValues(2)

TankValues = calculateCryogenicTank(O2_decompression, "O2")
MassTankForDecompression_CS = TankValues(1)
VolumeTankForDecompression_CS = TankValues(2)

TankValues = calculateCryogenicTank(N2_decompression, "N2")
MassTankForDecompression_CS = MassTankForDecompression_CS + TankValues(1)
VolumeTankForDecompression_CS = VolumeTankForDecompression_CS + TankValues(2)

' Calculate water tanks
Dim MassWaterTank
Dim VolumeWaterTank
Dim MassWasteWaterTank
Dim VolumeWasteWaterTank

TankValues = calculateWaterTank(H2O_Con)
MassWaterTank = TankValues(1)
VolumeWaterTank = TankValues(2)

TankValues = calculateWaterTank(H2O_Pro)
MassWasteWaterTank = TankValues(1)
VolumeWasteWaterTank = TankValues(2)

' now we write the tank values into the mission specific sheet
Dim TankSystemNames
Dim BaseTankSystemName
Dim TankMasses
Dim TankVolumes
Dim ResupplyMasses
Dim iMissionSpecific As Integer
Dim iTank As Integer
Dim CurrentRow
Dim iCurrentRow As Integer
Dim iHeader As Integer
Dim CrewValues(11) As Double
Dim TechnologyValues
Dim ReliabilityValues
Dim SystemMass As Double
    
TankSystemNames = Array("High Pressure Storage for Crew", "High Pressure Storage for Decompression", "Cryogenic Storage for Crew", "Cryogenic Storage for Decompression", "Potable Water Storage", "Waste Water Storage")
BaseTankSystemName = Array("High Pressure Storage", "High Pressure Storage", "Cryogenic Storage", "Cryogenic Storage", "Water Tank", "Water Tank")
TankMasses = Array(MassTankForCrew_HPS, MassTankForDecompression_HPS, MassTankForCrew_CS, MassTankForDecompression_CS, MassWaterTank, MassWasteWaterTank)
TankVolumes = Array(VolumeTankForCrew_HPS, VolumeTankForDecompression_HPS, VolumeTankForCrew_CS, VolumeTankForDecompression_CS, VolumeWaterTank, VolumeWasteWaterTank)
ResupplyMasses = Array(O2MassForCrew + (N2_leakage * mission_duration), (N2_decompression + O2_decompression), O2MassForCrew + (N2_leakage * mission_duration), (N2_decompression + O2_decompression), H2O_Con, 0)

iMissionSpecific = 2
For iTank = 0 To UBound(TankSystemNames)
    iMissionSpecific = iMissionSpecific + 1
    
    CurrentRow = findRowByName(Array(BaseTankSystemName(iTank)), "Technology", iNameColumn)
    iCurrentRow = CurrentRow(1)
    ' now we copy the relevant entries into the mission specific technology sheet. It is important to note, that the cell index i refers to the sheet "Technolog" and the index iMissionSpecific to the MissionSpecificTechnologyValues
    For iHeader = 1 To UBound(miTechnologyColumnToCopy)
        Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, miMissionSpecificColumnToCopy(iHeader)).Value = Worksheets("Technology").Cells(iCurrentRow, miTechnologyColumnToCopy(iHeader)).Value
    Next iHeader
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, miMissionSpecificColumnToCopy(1)).Value = TankSystemNames(iTank)
    
    ' "Mass", "Volume", "Heat", "Power", "Resupply Mass", "Maintenance"
    ' miMissionSpecificColumn
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, miMissionSpecificColumn(1)).Value = TankMasses(iTank)
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, miMissionSpecificColumn(2)).Value = TankVolumes(iTank)
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, miMissionSpecificColumn(3)).Value = 0
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, miMissionSpecificColumn(4)).Value = 0
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, miMissionSpecificColumn(5)).Value = ResupplyMasses(iTank)
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, miMissionSpecificColumn(6)).Value = 0
    
    MissionSpecificRowToTechnologyRow(iMissionSpecific) = iCurrentRow
    
    SystemMass = TankMasses(iTank)
    
    ReliabilityValues = calcsparesORUmass(iCurrentRow, SystemMass)
    
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, iSpareMassColumn).Value = ReliabilityValues(1)
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, iSpareVolumeColumn).Value = ReliabilityValues(2)
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, iReliabilityNoSparesColumn).Value = ReliabilityValues(3)
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, iReliabilityColumn).Value = ReliabilityValues(4)
    
Next iTank

'---------------------------------------------------------------------------------------------------------------------
' in this section we calculate all other systems
Dim SystemLevelFilter
Dim na As Integer
Dim i As Integer

SystemLevelFilter = Array(Worksheets("User Interface").Range("SystemLevelFilter").Value)
If SystemLevelFilter(0) = "both" Then
    SystemLevelFilter = Array("assembly", "component")
End If

'Get number of entries in technology sheet
na = Worksheets("Technology").Range("A1:A1000").Cells.SpecialCells(xlCellTypeConstants).Count + 1
    
    

' First we reset the global variable
For i = 1 To 300
    MissionSpecificRowToTechnologyRow(1) = 0
Next i
    
' Now calculate the values for the remaining assemblies:
' Important, the index i references the rows in the "Technology" worksheet, while the index iMissionSpecific references the rows in the "MissionSpecificTechnologyValues"
For i = 3 To na

    ' Only scale technologies matching the selected filter!
    
    If IsNumeric(Application.Match(Worksheets("Technology").Cells(i, miTechnologyColumnToCopy(6)).Value, SystemLevelFilter, 0)) Then
        ' do nothing (or actually execute code below but do not skip iteration)
    Else
        GoTo NextIteration
    End If
    
    ' skip tank storage system as they are handled above
    If Worksheets("Technology").Cells(i, miTechnologyColumnToCopy(1)).Value = "High Pressure Storage" Then
        GoTo NextIteration
    ElseIf Worksheets("Technology").Cells(i, miTechnologyColumnToCopy(1)).Value = "Cryogenic Storage" Then
        GoTo NextIteration
    ElseIf Worksheets("Technology").Cells(i, miTechnologyColumnToCopy(1)).Value = "Water tank" Then
        GoTo NextIteration
    End If
    
    iMissionSpecific = iMissionSpecific + 1

    ' First we copy the relevant entries into the mission specific technology sheet. It is important to note, that the cell index i refers to the sheet "Technolog" and the index iMissionSpecific to the MissionSpecificTechnologyValues
    For iHeader = 1 To UBound(miTechnologyColumnToCopy)
        Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, miMissionSpecificColumnToCopy(iHeader)).Value = Worksheets("Technology").Cells(i, miTechnologyColumnToCopy(iHeader)).Value
    Next iHeader
    
    CrewValues(1) = O2_Con
    CrewValues(2) = CO2_Pro
    CrewValues(3) = urine
    CrewValues(4) = hygiene + urine + sweat
    CrewValues(5) = feces
    CrewValues(6) = sweat
    CrewValues(7) = peak_sweat
    CrewValues(8) = heat
    CrewValues(9) = peak_heat
    CrewValues(10) = potable
    CrewValues(11) = food
    
    TechnologyValues = getTechnologyValues(i, CrewValues)
    '-----------------------
    
    ' "Mass", "Volume", "Heat", "Power", "Resupply Mass", "Maintenance"
    ' miMissionSpecificColumn
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, miMissionSpecificColumn(1)).Value = TechnologyValues(1)
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, miMissionSpecificColumn(2)).Value = TechnologyValues(2)
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, miMissionSpecificColumn(3)).Value = TechnologyValues(3)
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, miMissionSpecificColumn(4)).Value = TechnologyValues(4)
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, miMissionSpecificColumn(5)).Value = TechnologyValues(5)
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, miMissionSpecificColumn(6)).Value = TechnologyValues(6)
    
    SystemMass = TechnologyValues(1)
    ReliabilityValues = calcsparesORUmass(i, SystemMass)
    
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, iSpareMassColumn).Value = ReliabilityValues(1)
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, iSpareVolumeColumn).Value = ReliabilityValues(2)
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, iReliabilityNoSparesColumn).Value = ReliabilityValues(3)
    Worksheets("MissionSpecificTechnologyValues").Cells(iMissionSpecific, iReliabilityColumn).Value = ReliabilityValues(4)
    
    MissionSpecificRowToTechnologyRow(iMissionSpecific) = i
    
NextIteration:
Next i

' After calculation all values, we can perform the ESM calculation subroutine
esmcalc "MissionSpecificTechnologyValues"

End Sub

Function calculateLeakage() As Variant

Dim Pressurized_volume As Double
Dim NumberOfModules As Integer
Dim MM_O2 As Double
Dim MM_N2 As Double
Dim MM_Air As Double
Dim r As Double
Dim pp_O2 As Double
Dim pp_N2 As Double
Dim fDensityAir As Double
Dim fSpecificLeakRate As Double
Dim fLeakRatePerVestibular As Double
Dim fTotalLeakRate As Double
Dim O2_leakage As Double
Dim N2_leakage As Double

NumberOfModules = ThisWorkbook.Worksheets("User Interface").Range("Amount_of_Modules").Value
Pressurized_volume = ThisWorkbook.Worksheets("User Interface").Range("Pressurized_Volume").Value '[m3]

' Estimate leakage rates based on the pressurized volumen and number of modules:
' Uses values from P.Plötners Diplomarbeit page 45 for the leakage pr module and adapter between modules. For modules, the leak rate from US Lab (0.002722 kg/day) is divided with the pressurized volume of US Lab (97.71 m^3) to calculate a leak rate per pressurized volume.
' The number of modules adds verstibulares between the modules which have a leakrate of 0.000122 kg/day. The leakrates are converted to m^3 / day to estimate the influence of total pressure on the leakage
' In addition a oxygen partial pressure of 21.3 kPa is assumed regardless of overall atmospheric pressure, currently the N2 pressure is assumed fix at 79.76 as habitat pressure is no trade off parameter in LiSTOT at the moment
MM_Air = 0.028949 '[kg/mol]
MM_N2 = 28.0134 / 1000
MM_O2 = 31.9988 / 1000
r = 8.314472 '[J/K*mol] 'gas constant
pp_O2 = 21.3 * 10 ^ 3 '[Pa] 'partial pressure O2
pp_N2 = 79.76 * 10 ^ 3 '[Pa] 'partial pressure N2

fDensityAir = 101325 / ((r / MM_Air) * 293)
fSpecificLeakRate = (0.002722 / fDensityAir) / 97.71 ' [(m^3/day)/m^3]
fLeakRatePerVestibular = (0.000122 / fDensityAir)
fTotalLeakRate = fSpecificLeakRate * Pressurized_volume + (NumberOfModules - 1) * fLeakRatePerVestibular

O2_leakage = fTotalLeakRate * pp_O2 / ((r / MM_O2) * 293)  '[kg/d]
N2_leakage = fTotalLeakRate * pp_N2 / ((r / MM_N2) * 293)  '[kg/d]

Dim Leakage(1 To 2) As Double

Leakage(1) = O2_leakage
Leakage(2) = N2_leakage

calculateLeakage = Leakage

End Function

Function calculateWaterTank(WaterMass As Double) As Variant
' Subfunction to calculate the mass of the empty water tank and the total volume
' Output is Mass and Volume


' Calculate water tanks
Dim H2O_vol As Double 'needed H2O volume
Dim ShellWallThickness As Double
Dim BladderWallThickness As Double
Dim DensityWater As Double
Dim DensityShell As Double
Dim DensityBladder As Double
Dim InnerRadius As Double
Dim OuterRadius As Double
Dim MassWaterTank As Double
Dim VolumeBladder As Double
Dim VolumeShell As Double
Dim VolumeWaterTank As Double

DensityWater = 997 '[kg/m3] 'density of water
ShellWallThickness = 0.0005
BladderWallThickness = 0.0005
DensityShell = 2710 '[kg/m3]
DensityBladder = 1500 '[kg/m3]

H2O_vol = WaterMass / DensityWater

InnerRadius = ((3 / 4) * (H2O_vol / WorksheetFunction.Pi)) ^ (1 / 3)
OuterRadius = InnerRadius + ShellWallThickness + BladderWallThickness

VolumeWaterTank = (4 / 3) * WorksheetFunction.Pi * OuterRadius ^ 3

VolumeBladder = (4 / 3) * WorksheetFunction.Pi * (InnerRadius + BladderWallThickness) ^ 3
VolumeShell = VolumeWaterTank - VolumeBladder

MassWaterTank = VolumeBladder * DensityBladder + VolumeShell * DensityShell + 2.3

Dim TankValues(1 To 2) As Double
TankValues(1) = MassWaterTank
TankValues(2) = VolumeWaterTank
calculateWaterTank = TankValues

End Function

Function calculateHighPressureTank(ContentMass As Double, Content As String) As Variant
' Subfunction to calculate the mass of high pressure tanks
' Output is Mass and Volume

' define values for cryogenic storage
Dim DensityContent As Double
Dim VolumeContent As Double
Dim TankMassFactor As Double
Dim WallThickness As Double
Dim InnerRadius As Double
Dim OuterRadius As Double
Dim VolumeTank As Double
Dim MassTank As Double

' Densities are from NIST chemistry webbook
'See BVAD table 4.5 for mass factor of O2 and N2, for hydrogen used shuttle tanks to size them
Select Case Content
Case "N2"
    ' Stored at ~700 bar and ~293 K, assumed same max pressure as for H2 tank, see source below
    DensityContent = 498.34
    TankMassFactor = 0.556
Case "O2"
    ' Stored at ~ 200 bar and ~ 293 K
    DensityContent = 279.81
    TankMassFactor = 0.364
Case "H2"
    ' Stored at ~ 700 bar and ~ 293 K
    DensityContent = 39.7
    ' "For example, analysis of 35 MPa and 70 MPa Type IV composite cylinders that use high-strength carbon fiber (CF) reinforcement to reduce cylinder weight for automotive applications
    ' indicates hydrogen system capacities of 5.4% and 4.4% hydrogen by mass and 17.7 and 25.0 g hydrogen per liter volume" "Introduction to hydrogen storage" N.T. Stetson et.al, in Compendium of Hydrogen Energy, 2016
    TankMassFactor = 1 / 0.054
Case Else    ' Other values.
    MsgBox ("Unknown content provided to cryogenic tank calculation")
End Select
' just used to estimate the total tank volume
WallThickness = 0.02

VolumeContent = ContentMass / DensityContent

InnerRadius = ((3 / 4) * (VolumeContent / WorksheetFunction.Pi)) ^ (1 / 3)
OuterRadius = InnerRadius + WallThickness
VolumeTank = (4 / 3) * WorksheetFunction.Pi * OuterRadius ^ 3
MassTank = ContentMass * TankMassFactor

Dim TankValues(1 To 2) As Double
TankValues(1) = MassTank
TankValues(2) = VolumeTank
calculateHighPressureTank = TankValues

End Function

Function calculateCryogenicTank(ContentMass As Double, Content As String) As Variant
' Subfunction to calculate the mass of cryogenic tanks
' Output is Mass and Volume

' define values for cryogenic storage
Dim DensityContent As Double
Dim VolumeContent As Double
Dim TankMassFactor As Double
Dim WallThickness As Double
Dim InnerRadius As Double
Dim OuterRadius As Double
Dim VolumeTank As Double
Dim MassTank As Double


' Densities are from NIST chemistry webbook
'See BVAD table 4.5 for mass factor of O2 and N2, for hydrogen used shuttle tanks to size them
Select Case Content
Case "N2"
    ' Stored at ~36 bar and ~66 K
    DensityContent = 861.87
    TankMassFactor = 0.524
Case "O2"
    ' Stored at ~ 36 bar and ~ 66 K
    DensityContent = 1259.9
    TankMassFactor = 0.429
Case "H2"
    ' Stored at ~ 36 bar and ~ 18 K
    DensityContent = 76.764
    ' Shuttle H2 Tanks had a weight of 98 kg with a content of 41.8 kg see: S. Elitzur, V. Rosenband, A. Gany, Combined energy production and waste management in manned spacecraft utilizing on-demand hydrogen production and fuel cells, Acta Astronautica 128 (2016) 580–583.
    TankMassFactor = 2.34
Case Else    ' Other values.
    MsgBox ("Unknown content provided to cryogenic tank calculation")
End Select
' just used to estimate the total tank volume
WallThickness = 0.02

VolumeContent = ContentMass / DensityContent

InnerRadius = ((3 / 4) * (VolumeContent / WorksheetFunction.Pi)) ^ (1 / 3)
OuterRadius = InnerRadius + WallThickness
VolumeTank = (4 / 3) * WorksheetFunction.Pi * OuterRadius ^ 3
MassTank = ContentMass * TankMassFactor

Dim TankValues(1 To 2) As Double
TankValues(1) = MassTank
TankValues(2) = VolumeTank
calculateCryogenicTank = TankValues

End Function

Function getTechnologyValues(Technology As Variant, CrewValues As Variant) As Variant
' This sub is used to get the values for mass, power, cooling etc from the technology sheet and scale it with the current use case
' Output is order:  "Mass", "Volume", "Heat", "Power", "Resupply Mass", "Maintenance"
'
' The first input can either be an integer containing the row number in the technology sheet of the specific technology or it can be the abbreviation of the technology from the technology sheet
'
' Input for crew values must be for the whole crew, not per CM and is ordered:
' O2_Con, CO2_Pro, urine, TotalWaterToProcessing, feces, sweat, peak_sweat, heat, peak_heat, potable, food

    Dim iTechnologyRow As Integer
    
    If IsNumeric(Technology) Then
        iTechnologyRow = Technology
    ElseIf TypeName(Technology) = "String" Then
        ' In this case the function was provided with an abbreviation of a technology, so we have to look for the row here
        
        Dim na As Integer
        Dim i As Integer
        Dim miTechnologyColumnToCopy
        
        miTechnologyColumnToCopy = findColumnByName(Array("Name", "Abbreviation"), "Technology", 1)
        
        'Get number of entries in technology sheet
        na = Worksheets("Technology").Range("A1:A1000").Cells.SpecialCells(xlCellTypeConstants).Count + 1
        
        ' Now calculate the values for the remaining assemblies:
        ' Important, the index i references the rows in the "Technology" worksheet, while the index iMissionSpecific references the rows in the "MissionSpecificTechnologyValues"
        For i = 3 To na
            If Worksheets("Technology").Cells(i, miTechnologyColumnToCopy(2)).Value = Technology Then
                iTechnologyRow = i
                Exit For
            End If
        Next i
    Else
        MsgBox ("getTechnologyValues subfunction was provided a wrong input")
        Exit Function
    End If
    
    Dim Crew_Size As Integer
    Dim mission_duration As Double
    
    Crew_Size = Worksheets("User Interface").Range("Crew_Size").Value
    mission_duration = Worksheets("User Interface").Range("Mission_Duration").Value
    
    '---------------------------------------------------------------------------------------------------------------------
    ' Define BVAD metabolic values
    'Metabolic Needs
    ' Values from NASA Baseline Values and Assumptions Document (BVAD)
    Dim BVAD_O2_Con As Double 'O2 Consumption in [kg/CM*d]
    Dim BVAD_CO2_Pro As Double 'CO2 Production [kg/CM*d]
    Dim BVAD_urine As Double 'urine Production in [kg/CM*d]
    Dim BVAD_potable As Double 'potable water Need in [kg/CM*d]
    Dim BVAD_hygiene As Double 'hygiene water Need in [kg/CM*d]
    Dim BVAD_feces As Double 'feces Production in [kg/CM*d]
    Dim BVAD_avg_swe_vapr As Double 'average sweat and vpr Production in [kg/CM*d]
    Dim BVAD_peak_swe_vapr As Double 'peak sweat and vpr Production in [kg/CM*h]
    Dim BVAD_heat As Double 'heat Production in [kJ/CM*d]
    Dim BVAD_peak_heat As Double 'heat Production in [kJ/CM*h]
    Dim BVAD_H2O_Con As Double 'Total amount of H2O consumed (potable+hygiene+rehydration)
    Dim BVAD_H2O_Pro As Double 'Total amount of Waste water produced (urine+hygiene+sweat)
    Dim BVAD_food As Double 'Total amount of Waste water produced (urine+hygiene+sweat)
    
    'Metabolic needs assignement
    ' BVAD values from table 3.26 Summary of Nominal Human Metabolic Interface Values except for hygiene water which is based on
    ' Table 4.21 Typical Steady-State Water Usage Rates for Various Missions
    BVAD_O2_Con = 0.816
    BVAD_CO2_Pro = 1.04
    BVAD_urine = 1.6
    BVAD_potable = 2.5
    BVAD_hygiene = 0.7 ' up to 7.32 for mature planetary bases
    BVAD_feces = 0.132
    BVAD_avg_swe_vapr = 1.9
    BVAD_peak_swe_vapr = 0.77 ' from Table 3.22 Crew Induced Metabolic Loads only value per h and not per day
    BVAD_heat = 12000
    BVAD_peak_heat = 2974 ' from Table 3.22 Crew Induced Metabolic Loads only value per h and not per day
    BVAD_H2O_Con = (BVAD_potable + BVAD_hygiene) * mission_duration * Crew_Size
    BVAD_H2O_Pro = (BVAD_hygiene + BVAD_avg_swe_vapr + BVAD_urine) * mission_duration * Crew_Size
    BVAD_food = 1.51

    '---------------------------------------------------------------------------------------------------------------------
    ' Define scaling factors from inputs
    Dim CrewScalingO2 As Double
    Dim CrewScalingCO2 As Double
    Dim CrewScalingUrine As Double
    Dim CrewScalingTotalWater As Double
    Dim CrewScalingFeces As Double
    Dim CrewScalingSweat As Double
    Dim CrewScalingPeakSweat As Double
    Dim CrewScalingHeat As Double
    Dim CrewScalingPeakHeat As Double
    Dim CrewScalingPotable As Double
    Dim CrewScalingFood As Double
    
    Dim miEfficiencyColumn
    Dim fEfficiency
    miEfficiencyColumn = findColumnByName(Array("Efficiency"), "Technology", 1)
    fEfficiency = Worksheets("Technology").Cells(iTechnologyRow, miEfficiencyColumn(1)).Value
    
    ' Now we calculate metabolic scaling factors, not only the crew but also the schedule impact how systems should be scaled (more exercise results in more O2 CO2 loads on the system)
    CrewScalingO2 = CrewValues(1) / BVAD_O2_Con
    CrewScalingCO2 = CrewValues(2) / BVAD_CO2_Pro
    CrewScalingUrine = CrewValues(3) / BVAD_urine
    CrewScalingTotalWater = CrewValues(4) / (BVAD_hygiene + BVAD_avg_swe_vapr + BVAD_urine)
    CrewScalingFeces = CrewValues(5) / BVAD_feces
    If IsNumeric(fEfficiency) = True And IsEmpty(fEfficiency) = False And fEfficiency <> 0 Then
        CrewScalingSweat = WorksheetFunction.Min(CrewValues(6) / BVAD_avg_swe_vapr, (CrewValues(6) * 2257000 / 86400) / fEfficiency)
        CrewScalingPeakSweat = WorksheetFunction.Min(CrewValues(7) / BVAD_peak_swe_vapr, (CrewValues(7) * 2257000 / 3600) / fEfficiency)
    Else
        CrewScalingSweat = CrewValues(6) / BVAD_avg_swe_vapr
        CrewScalingPeakSweat = CrewValues(7) / BVAD_peak_swe_vapr
    End If
    CrewScalingHeat = CrewValues(8) / BVAD_heat
    CrewScalingPeakHeat = CrewValues(9) / BVAD_peak_heat
    CrewScalingPotable = CrewValues(10) / BVAD_potable
    CrewScalingFood = CrewValues(11) / BVAD_food
    
    
    
    'Contigency Factrors
    
    Dim Con1_mass50(3) As Double
    Dim ContingencyMass50(3) As Double
    Dim ContingencyMass500(3) As Double
    Dim ContingencyMass2500(3) As Double
    Dim ContingencyMassOver2500(3) As Double
    Dim ContingencyPower50(3) As Double
    Dim ContingencyPower500(3) As Double
    Dim ContingencyPower2500(3) As Double
    Dim ContingencyPowerOver2500(3) As Double
    
    'Contigency mass and power assignment
    ContingencyMass50(0) = 35 / 100
    ContingencyMass50(1) = 25 / 100
    ContingencyMass50(2) = 3 / 100
    
    ContingencyMass500(0) = 30 / 100
    ContingencyMass500(1) = 20 / 100
    ContingencyMass500(2) = 3 / 100
    
    ContingencyMass2500(0) = 25 / 100
    ContingencyMass2500(1) = 15 / 100
    ContingencyMass2500(2) = 1 / 100
    
    ContingencyMassOver2500(0) = 22 / 100
    ContingencyMassOver2500(1) = 12 / 100
    ContingencyMassOver2500(2) = 0.8 / 100
    
    ContingencyPower50(0) = 75 / 100
    ContingencyPower50(1) = 25 / 100
    ContingencyPower50(2) = 12 / 100
    
    ContingencyPower500(0) = 65 / 100
    ContingencyPower500(1) = 22 / 100
    ContingencyPower500(2) = 12 / 100
    
    ContingencyPower2500(0) = 60 / 100
    ContingencyPower2500(1) = 20 / 100
    ContingencyPower2500(2) = 12 / 100
    
    ContingencyPowerOver2500(0) = 35 / 100
    ContingencyPowerOver2500(1) = 20 / 100
    ContingencyPowerOver2500(2) = 11 / 100
    
    
    ' Scaling is based on crews sized, mission duration and on nominal BVAD flows compared to this crew schedules flows
    ' All values in the technology sheet are per CM, here we assume that this means per Standard BVAD CM
    ' How to find out the correct flowrate to which it should be adjusted? Use type column?
    ' urine processing, CO2 reduction, O2 generation, CO2 removal, filtration (Water), control -> subsystem air-THC
    
    'Scaling factor
    Dim ScalingFactor As Double
    If Crew_Size >= 1 And Crew_Size < 7 Then ScalingFactor = 1
    If Crew_Size >= 7 And Crew_Size < 10 Then ScalingFactor = 0.9
    If Crew_Size >= 10 Then ScalingFactor = 0.8
    

    '---------------------------------------------------------------------------------------------------------------------
    ' In this section we match the functions for technologies to the values with which they scale.
    ' This way the user can simply add new technologies in the technology sheet, as long as ThemeColor functions are defined correctly everything still works
    
    Dim ComparisonArrayO2
    Dim ComparisonArrayCO2
    Dim ComparisonArrayCrew
    Dim ComparisonArrayHeat
    Dim ComparisonArraySweat
    Dim ComparisonArrayTotalWaterLoad
    Dim ComparisonArrayUrine
    Dim ComparisonArrayFecesAndFoodPackaging
    
    ' functions scaling with O2 consumption
    ComparisonArrayO2 = Array("Regenerate Oxygen")
    ' functions scaling with required CO2
    ComparisonArrayCO2 = Array("CO2 Removal", "CO2 Reduction")
    ' functions scaling with crew size directly
    ComparisonArrayCrew = Array("Remove Gaseous Atmospheric Contaminants", "Remove Airborne Microbes")
    ' Functions scaling with heat release into the cabin
    ComparisonArrayHeat = Array("Control Atmospheric Temperature a)", "Control Atmospheric Temperature b)")
    ' Functions Scaling with humidity release
    ComparisonArraySweat = Array("Remove or Add Sensible Heat", "Control Atmospheric Humidity", "Remove or Add Moisture")
    ' Process Wastewater a) is condensate, b is hygiene waste water c is urine
    ComparisonArrayTotalWaterLoad = Array("Process Wastewater a)", "Process Wastewater b)")
    ComparisonArrayUrine = Array("Process Wastewater c)", "Process Wastewater d)")
    ' Functions that scale with solid waste production e.g. from feces and food packaging, but also from concentrated liquid (e.g. brine from urine).
    ' But the technologies to process urine and brine should be fiven the function Process Wastewater c) as well to identify them as scaling with urine amount
    ComparisonArrayFecesAndFoodPackaging = Array("Store Solid and Concentrated Liquid Wastes a)", "Store Solid and Concentrated Liquid Wastes b)", "Process Solid and Concentrated Liquid Wastes")
    
    '---------------------------------------------------------------------------------------------------------------------
    ' Get the column indices
    
    Dim TRLHeader
    Dim iTRLHeader As Integer
    Dim TechnologyHeaderToCalculate
    Dim TechnologyFunctionHeader
    Dim miTechnologyColumn
    Dim miTechnologyColumnFunction
    TRLHeader = findColumnByName(Array("TRL"), "Technology", 1)
    iTRLHeader = TRLHeader(1)
    
    TechnologyHeaderToCalculate = Array("Mass", "Volume", "Heat", "Power", "Resupply Mass", "fix Maintenance", "Maintenance per CM")
    TechnologyFunctionHeader = Array("function1", "function2", "function3", "function4", "function5", "function6", "function7", "function8", "function9", "function10")
    miTechnologyColumn = findColumnByName(TechnologyHeaderToCalculate, "Technology", 1)
    miTechnologyColumnFunction = findColumnByName(TechnologyFunctionHeader, "Technology", 1)

    '---------------------------------------------------------------------------------------------------------------------
    Dim TechScalingFactor As Double
    Dim MassPerCM As Double
    Dim VolumePerCM As Double
    Dim CoolantPerCM As Double
    Dim PowerPerCM As Double
    Dim ResupplyMassPerDayAndCM As Double
    Dim fixMaintenance As Double
    Dim MaintenancePerCM As Double
    Dim AssemblyMass As Double
    Dim AssemblyVolume As Double
    Dim AssemblyCoolant As Double
    Dim AssemblyPower As Double
    Dim AssemblyResupply As Double
    Dim AssemblyMaintenance As Double
    Dim trl As Double
    Dim iTRLIndex As Double
    Dim iHeader As Integer

    TechScalingFactor = -1
    
    ' Now we check the defined functions for this component to decide how to scale it
    For iHeader = 1 To UBound(miTechnologyColumnFunction)
        ' We check all functions of the current technology and use the maximum required scaling factor for any of the functions it handles (e.g. the CCAA handles humidity and heat, and the larger of these two difference will be used to scale it)
        If IsNumeric(Application.Match(Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value, ComparisonArrayO2, 0)) Then
            TechScalingFactor = WorksheetFunction.Max(TechScalingFactor, CrewScalingO2)
            
        ElseIf IsNumeric(Application.Match(Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value, ComparisonArrayCO2, 0)) Then
            TechScalingFactor = WorksheetFunction.Max(TechScalingFactor, CrewScalingCO2)
            
        ElseIf IsNumeric(Application.Match(Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value, ComparisonArrayCrew, 0)) Then
            TechScalingFactor = WorksheetFunction.Max(TechScalingFactor, Crew_Size)
            
        ElseIf IsNumeric(Application.Match(Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value, ComparisonArrayHeat, 0)) Then
            TechScalingFactor = WorksheetFunction.Max(TechScalingFactor, CrewScalingHeat, CrewScalingPeakHeat)
            
        ElseIf IsNumeric(Application.Match(Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value, ComparisonArraySweat, 0)) Then
            TechScalingFactor = WorksheetFunction.Max(TechScalingFactor, CrewScalingSweat, CrewScalingPeakSweat)
            
        ElseIf IsNumeric(Application.Match(Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value, ComparisonArrayTotalWaterLoad, 0)) Then
            TechScalingFactor = WorksheetFunction.Max(TechScalingFactor, CrewScalingTotalWater)
            
        ElseIf IsNumeric(Application.Match(Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value, ComparisonArrayUrine, 0)) Then
            TechScalingFactor = WorksheetFunction.Max(TechScalingFactor, CrewScalingUrine)
            
        ElseIf IsNumeric(Application.Match(Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumnFunction(iHeader)).Value, ComparisonArrayFecesAndFoodPackaging, 0)) Then
            TechScalingFactor = WorksheetFunction.Max(TechScalingFactor, CrewScalingFeces, CrewScalingFood)
        End If
    Next iHeader
    
    If TechScalingFactor = -1 Then
        ' All technologies that do not fall into the above category are assumed as fixed units which do not scale
        TechScalingFactor = 1
    End If
    
    ' Now we get the requires parameters for the calculation from the technology sheet
    MassPerCM = Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumn(1)).Value
    VolumePerCM = Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumn(2)).Value
    CoolantPerCM = Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumn(3)).Value
    PowerPerCM = Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumn(4)).Value
    ResupplyMassPerDayAndCM = Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumn(5)).Value
    fixMaintenance = Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumn(6)).Value
    MaintenancePerCM = Worksheets("Technology").Cells(iTechnologyRow, miTechnologyColumn(7)).Value
    
    ' scale all values to the maximum scaling value using the scaling factor to slightly reduce mass for large crews
    ' Note since the scaling factor is calculated using the whole crew metabolic factors, it basically includes the number of crew members!
    AssemblyMass = MassPerCM * ScalingFactor * TechScalingFactor
    AssemblyVolume = VolumePerCM * ScalingFactor * TechScalingFactor
    AssemblyCoolant = CoolantPerCM * ScalingFactor * TechScalingFactor
    AssemblyPower = PowerPerCM * ScalingFactor * TechScalingFactor
    AssemblyResupply = ResupplyMassPerDayAndCM * TechScalingFactor * mission_duration
    AssemblyMaintenance = fixMaintenance + MaintenancePerCM * TechScalingFactor
    
    trl = Worksheets("Technology").Cells(iTechnologyRow, iTRLHeader).Value

    If trl > 1 And trl < 4 Then
        iTRLIndex = 0
    ElseIf trl >= 4 And trl < 7 Then
        iTRLIndex = 1
    ElseIf trl >= 7 And trl < 10 Then
        iTRLIndex = 2
    End If
    ' Scale mass with contingency factor
    If AssemblyMass < 50 Then
        AssemblyMass = (1 + ContingencyMass50(iTRLIndex)) * AssemblyMass
    ElseIf AssemblyMass > 50 And AssemblyMass < 500 Then
        AssemblyMass = (1 + ContingencyMass500(iTRLIndex)) * AssemblyMass
    ElseIf AssemblyMass > 500 And AssemblyMass < 2500 Then
        AssemblyMass = (1 + ContingencyMass2500(iTRLIndex)) * AssemblyMass
    ElseIf AssemblyMass > 2500 Then
        AssemblyMass = (1 + ContingencyMassOver2500(iTRLIndex)) * AssemblyMass
    End If
    ' scale power with contingency factor
    If AssemblyPower < 50 Then
        AssemblyPower = (1 + ContingencyPower50(iTRLIndex)) * AssemblyPower
    ElseIf AssemblyPower > 50 And AssemblyPower < 500 Then
        AssemblyPower = (1 + ContingencyPower500(iTRLIndex)) * AssemblyPower
    ElseIf AssemblyPower > 500 And AssemblyPower < 2500 Then
        AssemblyPower = (1 + ContingencyPower2500(iTRLIndex)) * AssemblyPower
    ElseIf AssemblyPower > 2500 Then
        AssemblyPower = (1 + ContingencyPowerOver2500(iTRLIndex)) * AssemblyPower
    End If
    
    ' "Mass", "Volume", "Heat", "Power", "Resupply Mass", "Maintenance"
    ' miMissionSpecificColumn
    Dim TechnologyValues(6) As Double
    
    TechnologyValues(1) = AssemblyMass
    TechnologyValues(2) = AssemblyVolume
    TechnologyValues(3) = AssemblyCoolant
    TechnologyValues(4) = AssemblyPower
    TechnologyValues(5) = AssemblyResupply
    TechnologyValues(6) = AssemblyMaintenance
    getTechnologyValues = TechnologyValues
    
End Function



Function calcsparesORUmass(Technology As Variant, SystemMass As Double) As Variant
' This function is used to calculate the spare parts mass and volume, and the reliability of the system including spares
'
' as input it requires the technology row from the technology sheet as integer, or abbreviation name as string and the system mass as double in kg
'
' Outputs are ordered: "Mass for Spares/ORU", "Volume for Spares/ORU", "Reliability without Spares", "Reliability"


    Dim iTechnologyRow As Integer
    
    If IsNumeric(Technology) Then
        iTechnologyRow = Technology
    ElseIf TypeName(Technology) = "String" Then
        ' In this case the function was provided with an abbreviation of a technology, so we have to look for the row here
        
        Dim na As Integer
        Dim i As Integer
        Dim miTechnologyColumnToCopy
        
        miTechnologyColumnToCopy = findColumnByName(Array("Name", "Abbreviation"), "Technology", 1)
        
        'Get number of entries in technology sheet
        na = Worksheets("Technology").Range("A1:A1000").Cells.SpecialCells(xlCellTypeConstants).Count + 1
        
        ' Now calculate the values for the remaining assemblies:
        ' Important, the index i references the rows in the "Technology" worksheet, while the index iMissionSpecific references the rows in the "MissionSpecificTechnologyValues"
        For i = 3 To na
            If Worksheets("Technology").Cells(i, miTechnologyColumnToCopy(2)).Value = Technology Then
                iTechnologyRow = i
                Exit For
            End If
        Next i
    Else
        MsgBox ("calcsparesORUmass subfunction was provided a wrong input")
        Exit Function
    End If
    
    Dim n As Integer
    Dim mission_duration
    Dim doi 'date of installment
    Dim cpoi 'current point in time
    Dim lt 'life time of the respective system
    Dim age 'age of the system
    Dim imprf 'improvement factor for reliability
    Dim ORU_mass 'mass of 1 ORU
    Dim ORU_mass_total 'mass of all ORU
    Dim fp 'failure probability
    Dim noORU 'number of ORUs
    
    Dim bAvailableMTBF As Boolean
    Dim MTBF
    Dim MassOfOneORU
    Dim SystemType
    Dim ReliabilityWithoutSpare As Double
    Dim ReliabilityWithSpares As Double
    
    Dim ReliabilityHeaders
    Dim miReliabilityColumns
    
    Dim Outputs(4) As Double
    
    'Get parameters from mission definition and MCA analysis
    mission_duration = Worksheets("User Interface").Range("Mission_Duration").Value
    noORU = Worksheets("User Interface").Range("Amount_of_ORUs").Value
    
    imprf = 1
    
    '--------------------------------------------------------------------
    ' This step calculates the reliability without spares or ORUS to improve them
    
    'Calculate reliability of systems based on their MTBF. MTBF of an assembly is the seriell addition of each of the assembly's components
    'MTBF_ges=1/(1/MTBF_1+1/MTBF_2+...). The MTBF value needs to be inserted in the technology sheet for the respective system.
    'Poisson distribution is used to calculate rel values
    
    ReliabilityHeaders = Array("MTBF", "Spare Parts", "Mass of 1 ORU", "Uses", "Type")
    miReliabilityColumns = findColumnByName(ReliabilityHeaders, "Technology", 1)
    
    ReliabilityHeaders = Array("Mass", "Volume for Spares/ORU", "Mass for Spares/ORU", "Reliability without Spares", "Reliability")
    
    mission_duration = Worksheets("User Interface").Range("Mission_Duration").Value
    
    imprf = 1 'default at 1. Can be adjusted according to new technology standards
    
    MTBF = Worksheets("Technology").Cells(iTechnologyRow, miReliabilityColumns(1)).Value
    bAvailableMTBF = IsNumeric(MTBF) And Not IsEmpty(MTBF)
    If bAvailableMTBF = True Then
        ReliabilityWithoutSpare = WorksheetFunction.Poisson_Dist(0, (1 / MTBF) * mission_duration * 24, True) * imprf
        Outputs(3) = ReliabilityWithoutSpare
    End If
    
    '--------------------------------------------------------------------
    ' now calculate the ORU mass if the Amount of ORUs is selected as reliability calculation
    If Worksheets("User Interface").Range("F31").Value = "A. of ORUs" Then
        'calculate ORU masses if no ORU mass is provided in technology sheet
        'The factors 0.8, 0.67, etc. are equivalent to those of the respective ECLSS system on the ISS
        'This assumption was made due to the lack of information concerning new ECLSS systems
        
        With ThisWorkbook.Worksheets("Technology")
            MassOfOneORU = .Cells(iTechnologyRow, miReliabilityColumns(3))
    
            If IsNumeric(MassOfOneORU) And Not IsEmpty(MassOfOneORU) Then
                ' in this case a valid ORU mass is provided in the technology sheet and we use that value
            Else
                SystemType = .Cells(iTechnologyRow, miReliabilityColumns(5))
                
                If SystemType = "CO2 removal" Then
                    MassOfOneORU = 0.8 * SystemMass 'mass of 1 ORU
                ElseIf SystemType = "CO2 reduction" Then
                    MassOfOneORU = 0.67 * SystemMass 'mass of 1 ORU
                ElseIf SystemType = "O2 generation" Then
                    MassOfOneORU = 0.59 * SystemMass 'mass of 1 ORU
                ElseIf SystemType = "urine processing" Then
                    MassOfOneORU = 0.49 * SystemMass 'mass of 1 ORU
                ElseIf SystemType = "filtration" Then
                    MassOfOneORU = 0.55 * SystemMass 'mass of 1 ORU
                ElseIf SystemType = "storage" Then
                    MassOfOneORU = 20 'mass of 1 ORU, for storage system a fix mass is used, because the majority of the mass is not really active mass, only valves etc require spares
                ElseIf SystemType <> "CO2 reduction" And SystemType <> "CO2 removal" And SystemType <> "O2 generation" And SystemType <> "urine processing" And SystemType <> "filtration" Then
                    MassOfOneORU = 0.62 * SystemMass 'mass of 1 ORU. 0.62 is the mean value of the other ORU factors
                End If
            End If
        
        End With
        
        Outputs(1) = noORU * MassOfOneORU
        ' volume assuming a stowage factor of 0.004 m^3/kg, which is based on the CCAA hardware mass compare to hardware volume
        Outputs(2) = noORU * MassOfOneORU * 0.004
        
        
        If bAvailableMTBF = True Then
            ReliabilityWithSpares = (WorksheetFunction.Poisson_Dist(noORU, (1 / MTBF) * mission_duration * 24, True) * imprf)
            Outputs(4) = ReliabilityWithSpares
        End If
        
    ElseIf Worksheets("User Interface").Range("F31").Value = "0.001" Then
        
        If bAvailableMTBF = True Then
            While (1 - WorksheetFunction.Poisson_Dist(n, (1 / MTBF) * mission_duration * 24, True) * imprf) > 0.001
            n = n + 1
            Wend
            Outputs(1) = n * Worksheets("Technology").Cells(iTechnologyRow, ReliabilityHeaders(2))
            ' volume assuming a stowage factor of 0.004 m^3/kg, which is based on the CCAA hardware mass compare to hardware volume
            Outputs(2) = n * Worksheets("Technology").Cells(iTechnologyRow, ReliabilityHeaders(2)) * 0.004
            
            ReliabilityWithSpares = (WorksheetFunction.Poisson_Dist(n, (1 / MTBF) * mission_duration * 24, True) * imprf)
            Outputs(4) = ReliabilityWithSpares
            
            n = 0
        End If
        
    End If
    
    calcsparesORUmass = Outputs

End Function

Sub esmcalc(sWorkSheet As String)

'Calculate the ESM

Dim na
Dim i As Integer
Dim miColumns
Dim iColumnMass As Integer
Dim iColumnVolume As Integer
Dim iColumnHeat As Integer
Dim iColumnPower As Integer
Dim iColumnResupply As Integer
Dim iColumnMaintenance As Integer
Dim iColumnSpares As Integer
Dim iColumnORUs As Integer
Dim iColumnVolumeORUs As Integer
Dim iColumnESMNCT As Integer
Dim iColumnESMNoSpares As Integer
Dim iColumnESM As Integer

Dim ESM_LocationF As Double
Dim m_ESM
Dim ESM_VolumeF As Double
Dim v_ESM
Dim ESM_PowerF As Double
Dim p_ESM
Dim ESM_CoolingF As Double
Dim c_ESM
Dim ESM_Crew_TimeF As Double
Dim ct_ESM
Dim l_ESM
Dim SpareMass As Double
Dim SpareVolumeESM As Double

Dim iHeaderRow As Integer

'Load ESM factors

ESM_LocationF = Worksheets("User Interface").Range("ESM_LocationF").Value
ESM_VolumeF = Worksheets("User Interface").Range("ESM_VolumeF").Value
ESM_PowerF = Worksheets("User Interface").Range("ESM_PowerF").Value
ESM_CoolingF = Worksheets("User Interface").Range("ESM_CoolingF").Value
ESM_Crew_TimeF = Worksheets("User Interface").Range("ESM_Crew_TimeF").Value

'Get number of entries in the technology sheet
na = Worksheets(sWorkSheet).Range("A1:A1000").Cells.SpecialCells(xlCellTypeConstants).Count + 1

'Calculate ESM values depending on existing mass, power, volume, cooling, maintenance and resupply values

If sWorkSheet = "MissionSpecificTechnologyValues" Then
    iHeaderRow = 1
    miColumns = findColumnByName(Array("Mass", "Volume", "Heat", "Power", "Resupply Mass", "Maintenance", "Mass for Spares/ORU", "Volume for Spares/ORU", "ESM Without Crew Time", "ESM Without Spares", "ESM"), sWorkSheet, iHeaderRow)
    iColumnMass = miColumns(1)
    iColumnVolume = miColumns(2)
    iColumnHeat = miColumns(3)
    iColumnPower = miColumns(4)
    iColumnResupply = miColumns(5)
    iColumnMaintenance = miColumns(6)
    iColumnSpares = miColumns(7)
    iColumnVolumeORUs = miColumns(8)
    iColumnESMNCT = miColumns(9)
    iColumnESMNoSpares = miColumns(10)
    iColumnESM = miColumns(11)
    
ElseIf sWorkSheet = "MCA ESM" Then
    iHeaderRow = 4
    miColumns = findColumnByName(Array("Mass", "Volume", "Cooling", "Power", "Maintenance", "ESM Without Crew Time", "ESM"), sWorkSheet, iHeaderRow)
    iColumnMass = miColumns(1)
    iColumnVolume = miColumns(2)
    iColumnHeat = miColumns(3)
    iColumnPower = miColumns(4)
    iColumnMaintenance = miColumns(5)
    iColumnESMNCT = miColumns(6)
    iColumnESM = miColumns(7)

End If

    
For i = iHeaderRow + 2 To na

    If Worksheets(sWorkSheet).Cells(i, iColumnMass).Value <> "" Then
        m_ESM = Worksheets(sWorkSheet).Cells(i, iColumnMass).Value
    Else: m_ESM = 0
    End If
    If Worksheets(sWorkSheet).Cells(i, iColumnVolume).Value <> "" Then
        v_ESM = Worksheets(sWorkSheet).Cells(i, iColumnVolume).Value * ESM_VolumeF
    Else: v_ESM = 0
    End If
    If Worksheets(sWorkSheet).Cells(i, iColumnHeat).Value <> "" Then
        c_ESM = Worksheets(sWorkSheet).Cells(i, iColumnHeat).Value * ESM_CoolingF / 1000
    Else: c_ESM = 0
    End If
    If Worksheets(sWorkSheet).Cells(i, iColumnPower).Value <> "" Then
        p_ESM = Worksheets(sWorkSheet).Cells(i, iColumnPower).Value * ESM_PowerF / 1000
    Else: p_ESM = 0
    End If
    
    If Worksheets(sWorkSheet).Cells(i, iColumnMaintenance).Value <> "" Then
        ct_ESM = Worksheets(sWorkSheet).Cells(i, iColumnMaintenance).Value * ESM_Crew_TimeF
    Else: ct_ESM = 0
    End If
    
    l_ESM = 0
    SpareMass = 0
    SpareVolumeESM = 0
    
    If sWorkSheet = "MissionSpecificTechnologyValues" Then
        If Worksheets(sWorkSheet).Cells(i, iColumnResupply).Value <> "" Then
            l_ESM = Worksheets(sWorkSheet).Cells(i, iColumnResupply).Value
        End If
        If Worksheets(sWorkSheet).Cells(i, iColumnSpares).Value <> "" Then
            SpareMass = Worksheets(sWorkSheet).Cells(i, iColumnSpares).Value
        End If
        If Worksheets(sWorkSheet).Cells(i, iColumnVolumeORUs).Value <> "" Then
            SpareVolumeESM = Worksheets(sWorkSheet).Cells(i, iColumnVolumeORUs).Value * ESM_VolumeF
        End If
        
        Worksheets(sWorkSheet).Cells(i, iColumnESMNoSpares).Value = ESM_LocationF * (m_ESM + v_ESM + p_ESM + c_ESM + l_ESM + ct_ESM) 'ESM_NCT + Crewtime
    End If
'Add ESM values for mass, power, volume, cooling, maintenance and resupply to receive ESM_NCT (ESM without crew time) and ESM for each component and assembly

    Worksheets(sWorkSheet).Cells(i, iColumnESMNCT).Value = ESM_LocationF * (m_ESM + v_ESM + p_ESM + c_ESM + l_ESM) 'NCT
    Worksheets(sWorkSheet).Cells(i, iColumnESM).Value = ESM_LocationF * (m_ESM + v_ESM + p_ESM + c_ESM + l_ESM + ct_ESM + SpareMass + SpareVolumeESM) 'ESM_NCT + Crewtime + SpareMass + Spare Volume
Next i

' if we perform this calculation for the MCA ESM sheet, also recalculate the ranks of the alternatives
If sWorkSheet = "MCA ESM" Then
    Tabelle5.CalculateMCAandESMRanks
End If

End Sub


Sub InitializeMCA()
' This sub initializes the MCA values, by selecting only the technologies that match the criteria and copying the respective values into the MCA sheet

'Variable declaration
'Minimum Values

Dim min_reliability As Double
Dim min_trl As Integer

'Maximum Values

Dim max_mass As Double
Dim MAX_Volume As Double
Dim MAX_Power As Double
Dim MAX_Cooling As Double
Dim MAX_Maintenance As Double

'Desired Assembly Function

Dim DES_Function As String

'Variable declaration
Dim iTechnologyRow As Integer
Dim bExcludePartialAlternatives As Boolean

Worksheets("User Interface").Protect UserInterfaceOnly:=True

min_reliability = Worksheets("User Interface").Range("MIN_Reliability").Value
min_trl = Worksheets("User Interface").Range("MIN_TRL").Value

max_mass = Worksheets("User Interface").Range("MAX_Mass").Value
MAX_Volume = Worksheets("User Interface").Range("MAX_Volume").Value
MAX_Power = Worksheets("User Interface").Range("MAX_Power").Value
MAX_Cooling = Worksheets("User Interface").Range("MAX_Cooling").Value
MAX_Maintenance = Worksheets("User Interface").Range("MAX_Maintenance").Value

DES_Function = Worksheets("User Interface").Range("DES_Function").Value

Application.ScreenUpdating = False

'Deletion of previous MCA results

Worksheets("MCA ESM").Protect UserInterfaceOnly:=True
Worksheets("MCA ESM").Range("A6:Z20").ClearContents
Worksheets("MCA ESM").Range("A3").ClearContents

Dim cell

For Each cell In Worksheets("MCA ESM").Range("A6:R20")

    cell.Interior.ColorIndex = 0
    
Next

Worksheets("MCA ESM").Cells(2, 1) = DES_Function

'Check whether required information was inserted

If IsEmpty(Worksheets("User Interface").Range("Crew_Size").Value) = True Or IsEmpty(Worksheets("User Interface").Range("Mission_Duration").Value) = True Or _
   IsEmpty(Worksheets("User Interface").Range("Schedule_Number").Value) = True Or IsEmpty(Worksheets("User Interface").Range("Pressurized_Volume").Value) = True Then
    MsgBox ("Mission parameters have not been initialized")
    Exit Sub
End If

If IsEmpty(Worksheets("User Interface").Range("DES_Function").Value) = True Then
    MsgBox ("Desired Function has not been initialized")
    Exit Sub
End If

If IsEmpty(Worksheets("User Interface").Range("ESM_LocationF").Value) = True Or IsEmpty(Worksheets("User Interface").Range("ESM_VolumeF").Value) = True Or _
   IsEmpty(Worksheets("User Interface").Range("ESM_PowerF").Value) = True Or IsEmpty(Worksheets("User Interface").Range("ESM_CoolingF").Value) = True Or _
   IsEmpty(Worksheets("User Interface").Range("ESM_Crew_TimeF").Value) = True Then
    MsgBox ("ESM Factors have not been initialized")
    Exit Sub
End If

If IsEmpty(Worksheets("User Interface").Range("SystemLevelFilter").Value) = True Then
    MsgBox ("Required filter options have not been initialized")
    Exit Sub
End If

bExcludePartialAlternatives = False
If Worksheets("User Interface").Range("ExcludePartialAlternatives").Value = "Yes" Then
    bExcludePartialAlternatives = True
End If

Worksheets("MCA ESM").Cells(3, 1) = "System level:  " & Worksheets("User Interface").Range("SystemLevelFilter").Value

Dim noe As Integer
noe = Worksheets("MissionSpecificTechnologyValues").Cells(Rows.Count, 1).End(xlUp).row


Dim iHeaderRow As Integer
Dim miMissionSpecificColumns
Dim miAdditionalColumns
Dim miMCAColumns
Dim miFunctionColumns
Dim miTechnologyConfidenceColumns
Dim miMCAConfidenceColumns
Dim iFunction As Integer
Dim iTRLColumn As Integer
Dim iReliabilityColumn As Integer
Dim iUseColumn As Integer
Dim iMassColumn As Integer
Dim iVolumeColumn As Integer
Dim iPowerColumn As Integer
Dim iCoolingColumn As Integer
Dim iMaintenanceColumn As Integer
Dim iSpareMassColumn As Integer
Dim iSpareVolumeColumn As Integer
Dim iRessuplyMassColumn As Integer
Dim bValidSystem As Boolean
Dim iColumn As Integer
Dim iMCARow As Integer
Dim iMCA_MassColumn As Integer
Dim iMCA_VolumeColumn As Integer
Dim i As Integer
Dim p As Integer


iHeaderRow = 1
iMCARow = 6

miFunctionColumns = findColumnByName(Array("function1", "function2", "function3", "function4", "function5", "function6", "function7", "function8", "function9", "function10"), "Technology", iHeaderRow)
miAdditionalColumns = findColumnByName(Array("Mass for Spares/ORU", "Volume for Spares/ORU", "Resupply Mass"), "MissionSpecificTechnologyValues", iHeaderRow)
miTechnologyConfidenceColumns = findColumnByName(Array("TRL", "Mass Confidence Factor", "Volume Confidence Factor", "Power Confidence Factor", "Heat Confidence Factor", "Maintenance Confidence Factor", "MTBF Confidence Factor", "TRL Confidence Factor"), "Technology", 1)
iTRLColumn = miTechnologyConfidenceColumns(1)

' get header columns for the mission specific sheet that are copied to the MCA ESM sheet, note the order of headers here must match the order in the miMCAColumns!
miMissionSpecificColumns = findColumnByName(Array("Name", "Subsystem", "Type", "Uses", "Mass", "Volume", "Power", "Heat", "Maintenance", "Reliability", "ESM Without Crew Time", "ESM"), "MissionSpecificTechnologyValues", iHeaderRow)
iHeaderRow = 4
miMCAColumns = findColumnByName(Array("Name of Assembly", "ECLSS Subsystem", "Type", "Use", "Mass", "Volume", "Power", "Cooling", "Maintenance", "Reliability", "ESM Without Crew Time", "ESM"), "MCA ESM", iHeaderRow)
miMCAConfidenceColumns = findColumnByName(Array("TRL", "Mass CF", "Volume CF", "Power CF", "Cooling CF", "Maintenance CF", "Reliability CF", "TRL CF"), "MCA ESM", iHeaderRow)

iUseColumn = miMissionSpecificColumns(4)
iMassColumn = miMissionSpecificColumns(5)
iVolumeColumn = miMissionSpecificColumns(6)
iPowerColumn = miMissionSpecificColumns(7)
iCoolingColumn = miMissionSpecificColumns(8)
iMaintenanceColumn = miMissionSpecificColumns(9)
iReliabilityColumn = miMissionSpecificColumns(10)

iSpareMassColumn = miAdditionalColumns(1)
iSpareVolumeColumn = miAdditionalColumns(2)
iRessuplyMassColumn = miAdditionalColumns(3)

iMCA_MassColumn = miMCAColumns(5)
iMCA_VolumeColumn = miMCAColumns(6)

If MissionSpecificRowToTechnologyRow(3) = 0 Then
    calcsystemprops
End If

For i = 3 To noe
    
    iTechnologyRow = MissionSpecificRowToTechnologyRow(i)
    
    With ThisWorkbook.Worksheets("MissionSpecificTechnologyValues")
        ' check if the technology is used for the desired function
        bValidSystem = False
        For iFunction = 1 To 10
            If DES_Function = Worksheets("Technology").Cells(iTechnologyRow, miFunctionColumns(iFunction)) Then
                bValidSystem = True
            End If
        Next iFunction
        
        ' Check if the other criteria are fullfilled for this system
        If Worksheets("Technology").Cells(iTechnologyRow, iTRLColumn) >= min_trl And bValidSystem And _
            (.Cells(i, iReliabilityColumn) >= min_reliability Or min_reliability = 0) And _
            (.Cells(i, iMassColumn) <= max_mass Or max_mass = 0) And _
            (.Cells(i, iVolumeColumn) <= MAX_Volume Or MAX_Volume = 0) And _
            (.Cells(i, iPowerColumn) <= MAX_Power Or MAX_Power = 0) And _
            (.Cells(i, iCoolingColumn) <= MAX_Cooling Or MAX_Cooling = 0) And _
            (.Cells(i, iMaintenanceColumn) <= MAX_Maintenance Or MAX_Maintenance = 0) Then
            bValidSystem = True
        Else
            bValidSystem = False
        End If
        
        If bValidSystem Then
        ' check if we want to exclude alternatives with partially missing information, and then check if this is the case here
            If bExcludePartialAlternatives Then
                For iColumn = 1 To 12
                    If IsEmpty(.Cells(i, miMissionSpecificColumns(iColumn)).Value) = True Then
                        bValidSystem = False
                    End If
                Next iColumn
            End If
        End If
        
        If bValidSystem Then
            If .Cells(i, iMassColumn).Value <> 0 Then
                ' Copy the entries from the Mission Specific technology sheet to the MCA sheet, and sum up the mass and volume values
                For iColumn = 1 To 12
                    Worksheets("MCA ESM").Cells(iMCARow, miMCAColumns(iColumn)) = .Cells(i, miMissionSpecificColumns(iColumn)).Value
                Next iColumn
                ' Copy TRL and Confidence Factors:
                For iColumn = 1 To UBound(miTechnologyConfidenceColumns)
                    Worksheets("MCA ESM").Cells(iMCARow, miMCAConfidenceColumns(iColumn)) = Worksheets("Technology").Cells(iTechnologyRow, miTechnologyConfidenceColumns(iColumn)).Value
                Next iColumn
                
                ' Add spare and resupply mass to the Mass value in the MCA sheet
                Worksheets("MCA ESM").Cells(iMCARow, iMCA_MassColumn) = Worksheets("MCA ESM").Cells(iMCARow, iMCA_MassColumn) + .Cells(i, iSpareMassColumn).Value + .Cells(i, iRessuplyMassColumn).Value
                Worksheets("MCA ESM").Cells(iMCARow, iMCA_VolumeColumn) = Worksheets("MCA ESM").Cells(iMCARow, iMCA_VolumeColumn) + .Cells(i, iSpareVolumeColumn).Value
                
                iMCARow = iMCARow + 1
            End If
        End If
    End With
Next i

With ThisWorkbook.Worksheets("User Interface")
p = 0
For i = 1 To 7
    If .Cells(i + 9, "E").Value <> 0 Then p = p + 1
Next
End With

If p = 0 Then
    MsgBox ("All decision weights set to 0!")
    Exit Sub
End If
End Sub
