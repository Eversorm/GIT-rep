VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Sub calcsystemprops()
' Old code which should not be used, define/calculate system properties per BVAD standard crew member directly in the Technology work sheet!
Exit Sub
'Approach: Based on a given system in- or output and the metabolic factors of a human being (according  to NASA BVAD) a rescale factor is claculated
'with which the literature values for mass, power, cooling, volume and resupply of a system can be adjusted to a desired crew size and mission duration
'In addition, there a scaling factor for each system since it is unlikely that system attrributes scale purely linearly over a wide range of crew sizes.
'This factor can be individually adjusted for each system in this module. Since the scaling approach is based on assumptiopns, contigency masses and power
'are added based on the TRL of a system (eg. lower contingency values are added if the system posseses a higher TRL)

'This module includes as of now:

'   Air Evaporation System AES
'   Bosch Reaktor
'   CO2 and Moisture Removal Amine Swing-Bed System CAMRAS
'   CO2 Electrolysis COE
'   Electrochemical depolarization concentration EDC
'   Molecular Sieves 2BMS
'   Molecular sieves 4BMS or Carbon Dioxide Removal Assembly CDRA
'   Multifiltration MF
'   Sabatier Reactor
'   Solid amine water desorption SAWD
'   Solid polymer water electrolysis SPWE or Oxygen generation assembly OGS
'   Static feed water electrolysis SFWE
'   Thermoelectric Integrated Membrane Evaporation TIMES
'   Vapor Compression Distillation VCD
'   Vapor Phase Catalytic Ammonia Removal VPCAR
'   Advanced Carbon-Formation Reactor System ACRS
'   Lithium hydroxide LiOH
'   High Pressure Storage HPS
'   Cryogenic Storage CS
'   Oxygen Candles
'   CO2 Recution Assembly
'   CO2 Removal Assembly
'   Common Cabin Air Assembly (CCAA)
'   Algae Reactor (CO2 Removal and Urine Processing)

'general variable declaration

Dim DES_Function As String
Dim Crew_Size
Dim mission_duration
Dim Schedule
Dim Margin As Integer
Dim r 'gas constant
Dim R_O2
Dim R_N2
Dim O2_leakage 'daily O2 leakage rate
Dim N2_leakage 'daily N2 leakage rate
Dim MM_O2 'Molar masses for relevant fluids and gasses
Dim MM_CO2
Dim MM_N2
Dim MM_H2
Dim MM_H2O
Dim MM_CH4
Dim MM_LiOH
Dim MM_Li2CO3
Dim MM_LiClO4
Dim na
Dim i As Integer
Dim Pressurized_volume
Dim pp_O2 'partial pressure O2
Dim pp_N2 'partial pressure N2
Dim rho_H2O 'H2O density

'Metabolic Needs

Dim O2_Con 'O2 Consumption in [kg/CM*d]
Dim CO2_Pro 'CO2 Production [kg/CM*d]
Dim urine 'urine Production in [kg/CM*d]
Dim potable 'potable water Need in [kg/CM*d]
Dim hygiene 'hygiene water Need in [kg/CM*d]
Dim feces 'feces Production in [kg/CM*d]
Dim swe_vapr 'sweat and vpr Production in [kg/CM*d]
Dim heat 'heat Production in [kg/CM*d]
Dim H2O_Con 'Total amount of H2O consumed (potable+hygiene+rehydration)
Dim H2O_Pro 'Total amount of Waste water produced (urine+hygiene+sweat)

'Load attributes from User Interface and molar masses as well as other fluid related attributes
DES_Function = Worksheets("User Interface").Range("DES_Function").Value
Crew_Size = Worksheets("User Interface").Range("Crew_Size").Value
mission_duration = Worksheets("User Interface").Range("Mission_Duration").Value
Schedule = Worksheets("User Interface").Range("Schedule_Number").Value
Pressurized_volume = Worksheets("User Interface").Range("Pressurized_Volume").Value '[m3]
Margin = Worksheets("User Interface").Range("Mass_margin").Value
MM_O2 = 32 '[g/mol] 'molar mass
MM_CO2 = 44.01 '[g/mol]
MM_H2O = 18.01 '[g/mol]
MM_H2 = 2.01 '[g/mol]
MM_N2 = 28.01 '[g/mol]
MM_CH4 = 16.04 '[g/mol]
MM_LiOH = 23.95 '[g/mol]
MM_Li2CO3 = 73.9 '[g/mol]
MM_LiClO4 = 106.4 '[g/mol]
O2_leakage = 0.042 '[kg/d]
N2_leakage = 0.158 '[kg/d]
r = 8.314472 '[J/K*mol] 'gas constant
pp_O2 = 21.3 * 10 ^ 3 '[Pa] 'partial pressure O2
pp_N2 = 79.76 * 10 ^ 3 '[Pa] 'partial pressure N2
rho_H2O = 997 '[kg/m3] 'density of water
R_O2 = 259.84
R_N2 = 296.8

'Get number of entries in technology sheet
na = Worksheets("Technology").Range("A2:A1000").Cells.SpecialCells(xlCellTypeConstants).Count

'Metabolic needs assignement

If Schedule = 1 Then

    O2_Con = Worksheets("Schedule and Metabolic").Range("SUM_O2_Consumption").Value
    CO2_Pro = Worksheets("Schedule and Metabolic").Range("SUM_CO2_Production").Value
    urine = Worksheets("Schedule and Metabolic").Range("SUM_Urine_Production").Value
    hygiene = Worksheets("Schedule and Metabolic").Range("SUM_Hygiene_Water_Need").Value
    feces = Worksheets("Schedule and Metabolic").Range("SUM_Feces_Production").Value
    swe_vapr = Worksheets("Schedule and Metabolic").Range("SUM_Sweat_Production").Value
    heat = Worksheets("Schedule and Metabolic").Range("SUM_Heat_Load").Value
    H2O_Con = (potable + hygiene) * mission_duration * Crew_Size
    H2O_Pro = (hygiene + urine + swe_vapr) * mission_duration * Crew_Size

ElseIf Schedule = 2 Then

    O2_Con = Worksheets("Schedule and Metabolic").Range("SUM_O2_Consumption_2").Value / Crew_Size
    CO2_Pro = Worksheets("Schedule and Metabolic").Range("SUM_CO2_Production_2").Value / Crew_Size
    urine = Worksheets("Schedule and Metabolic").Range("SUM_Urine_Production_2").Value / Crew_Size
    hygiene = Worksheets("Schedule and Metabolic").Range("SUM_Hygiene_Water_Need_2").Value / Crew_Size
    feces = Worksheets("Schedule and Metabolic").Range("SUM_Feces_Production_2").Value / Crew_Size
    swe_vapr = Worksheets("Schedule and Metabolic").Range("SUM_Sweat_Production_2").Value / Crew_Size
    heat = Worksheets("Schedule and Metabolic").Range("SUM_Heat_Load_2").Value / Crew_Size
    H2O_Con = (potable + hygiene) * mission_duration * Crew_Size
    H2O_Pro = (hygiene + urine + swe_vapr) * mission_duration * Crew_Size
    
    'Check if user defined schedule has been set
    If O2_Con = 0 Then
        MsgBox ("User defined schedule has not been defined! Adjust schedule before proceeding!")
        Exit Sub
    End If
     
End If

'Contigency mass and power assignment

Con1_mass50 = 0
Con2_mass50 = 0
Con3_mass50 = 0
Con1_mass500 = 0
Con2_mass500 = 0
Con3_mass500 = 0
Con1_mass2500 = 0
Con2_mass2500 = 0
Con3_mass2500 = 0
Con1_mass2500_ = 0
Con2_mass2500_ = 0
Con3_mass2500_ = 0

Con1_power50 = 0
Con2_power50 = 0
Con3_power50 = 0
Con1_power500 = 0
Con2_power500 = 0
Con3_power500 = 0
Con1_power2500 = 0
Con2_power2500 = 0
Con3_power2500 = 0
Con1_power2500_ = 0
Con2_power2500_ = 0
Con3_power2500_ = 0
CO2Reduction:

'Assembly consisting of compressor, valves, sensors and heat exchanger. Mass without actual reactor system

Dim m_CO2rd_a
Dim m_sensor_rd
Dim m_valve_rd
Dim m_heatex_rd
Dim m_comp_rd
Dim m_condens_rd

For i = 2 To na

    If Cells(i, 3).Value = "CO2 reduction" And Cells(i, "I").Value <> "" Then

        If Cells(i, 1).Value = "Flow sensor" Then m_sensor_rd = 5 * Cells(i, "I").Value
        If Cells(i, 1).Value = "Valve (check)" Then m_valve_rd = 3 * Cells(i, "I").Value
        If Cells(i, 1).Value = "Compressor" Then m_comp_rd = Cells(i, "I").Value
        If Cells(i, 1).Value = "Heater" Then m_heatex_rd = Cells(i, "I").Value
        If Cells(i, 1).Value = "Condenser/ separator" Then m_condens_rd = Cells(i, "I").Value
        
    End If
    
Next i

For i = 2 To na

    If Cells(i, 1).Value = "CO2 Reduction Assembly" Then
        Cells(i, "I").Value = m_sensor_rd + m_valve_rd + m_heatex_rd + m_condens_rd + m_comp_rd
        m_CO2rd_a = Cells(i, "I").Value
    End If
    
Next i

'Assembly consisting of pump, pump fan controller, blower, valves, sensors and heat exchanger. Mass without actual removal system (e.g. amine beds)

CO2Removal:

Dim m_CO2rm_a
Dim m_sensor_rm
Dim m_checkvalves_rm
Dim m_selectorvalves_rm
Dim m_heatex_rm
Dim m_pump_rm
Dim m_pumpfanc_rm
Dim m_blower_rm

For i = 2 To na

    If Cells(i, 3).Value = "CO2 removal" And Cells(i, "I").Value <> "" Then

        If Cells(i, 1).Value = "Flow sensor" Then m_sensor_rm = 5 * Cells(i, "I").Value 'represents the mass of all needed sensors (including temperature, pressure, etc.)
        If Cells(i, 1).Value = "Valve (check)" Then m_checkvalves_rm = Cells(i, "I").Value
        If Cells(i, 1).Value = "Valve (selector)" Then m_selectorvalves_rm = Cells(i, "I").Value
        If Cells(i, 1).Value = "Heat exchanger" Then m_heatex_rm = Cells(i, "I").Value
        If Cells(i, 1).Value = "Pump" Then m_pump_rm = Cells(i, "I").Value
        If Cells(i, 1).Value = "Blower" Then m_blower_rm = Cells(i, "I").Value
        If Cells(i, 1).Value = "Pump Fan Controller" Then m_pumpfanc_rm = Cells(i, "I").Value
        
    End If
    
Next i

For i = 2 To na

        If Cells(i, 1).Value = "CO2 Removal Assembly" Then
            Cells(i, "I").Value = m_sensor_rm + m_checkvalves_rm + m_heatex_rm + m_pump_rm + m_pumpfanc_rm + m_blower_rm + m_selectorvalves_rm
            m_CO2rm_a = Cells(i, "I").Value
        End If

Next i


CCAA:

'Hard to scale. Scaled according to 1.45 kg/hr condensate processing by 5 CCAA assemblies on ISS. -> 0.29 kg/hr processing rate assumed for one CCAA
' Note this assumption is wrong. The ISS does have 5 CCAAs but not all operate at the same time. Based on the Value from Wieland 1998 CCAA can remove 1 kW latent heat, which corresponds to 1.59 kg/h of humidity removal capacity.
' Of course this seems largely over designed, because a human only produces ~ 1.9 kg/d, so a crew of 6 on average produces 0.475 kg of humidity per hour. However, the CCAA must be designed for peak loads and during exercise one crew member produces up to 0.77 kg/h of humidity.
' so with two people exercising at the same time, 1.54 kg/h of humidity are produced, which can just barely be covered by CCAA. So a good scaling approach for this would be to select the highest humidity load from the schedule and size the CCAA accordingly


Dim SF_CCAA 'Scaling Faktor
Dim RF_CCAA 'Rescale factor for CCAA
Dim massg_CCAA 'given system mass for CCAA
Dim massreg_CCAA 'given resupply mass for CCAA
Dim volg_CCAA 'given volume for CCAA
Dim powg_CCAA 'given power for CCAA
Dim coolg_CCAA ' given cooling for CCAA
Dim Vapr_in 'given condensate input
Dim TRL_CCAA
Dim eff_CCAA

Dim mass_CCAA
Dim massre_CCAA
Dim vol_CCAA
Dim pow_CCAA
Dim cool_CCAA
Dim massc_CCAA
Dim powc_CCAA

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_CCAA = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_CCAA = 0.9
If Crew_Size >= 10 Then SF_CCAA = 0.8

massg_CCAA = 112
massreg_CCAA = 0
volg_CCAA = 0.4
powg_CCAA = 468
coolg_CCAA = 468
Vapr_in = 0.29 * 24 'per day
TRL_CCAA = 9
eff_CCAA = 1

RF_CCAA = swe_vapr * (1 + Margin) / Vapr_in

massc_CCAA = RF_CCAA * massg_CCAA * Crew_Size * SF_CCAA

If TRL_CCAA > 1 And TRL_CCAA < 4 Then

    If massc_CCAA < 50 Then mass_CCAA = (1 + Con1_mass50) * massc_CCAA
    If massc_CCAA > 50 And massc_CCAA < 500 Then mass_CCAA = (1 + Con1_mass500) * massc_CCAA
    If massc_CCAA > 500 And massc_CCAA < 2500 Then mass_CCAA = (1 + Con1_mass2500) * massc_CCAA
    If massc_CCAA > 2500 Then mass_CCAA = (1 + Con1_mass2500_) * massc_CCAA

End If

If TRL_CCAA >= 4 And TRL_CCAA < 7 Then

    If massc_CCAA < 50 Then mass_CCAA = (1 + Con2_mass50) * massc_CCAA
    If massc_CCAA > 50 And massc_CCAA < 500 Then mass_CCAA = (1 + Con2_mass500) * massc_CCAA
    If massc_CCAA > 500 And massc_CCAA < 2500 Then mass_CCAA = (1 + Con2_mass2500) * massc_CCAA
    If massc_CCAA > 2500 Then mass_CCAA = (1 + Con2_mass2500_) * massc_CCAA

End If

If TRL_CCAA >= 7 And TRL_CCAA < 10 Then

    If massc_CCAA < 50 Then mass_CCAA = (1 + Con3_mass50) * massc_CCAA
    If massc_CCAA > 50 And massc_CCAA < 500 Then mass_CCAA = (1 + Con3_mass500) * massc_CCAA
    If massc_CCAA > 500 And massc_CCAA < 2500 Then mass_CCAA = (1 + Con3_mass2500) * massc_CCAA
    If massc_CCAA > 2500 Then mass_CCAA = (1 + Con3_mass2500_) * massc_CCAA

End If

massre_CCAA = RF_CCAA * massreg_CCAA * Crew_Size * SF_CCAA * mission_duration
vol_CCAA = RF_CCAA * volg_CCAA * Crew_Size * SF_CCAA
cool_CCAA = RF_CCAA * coolg_CCAA * Crew_Size * SF_CCAA
powc_CCAA = RF_CCAA * powg_CCAA * Crew_Size * SF_CCAA

If TRL_CCAA > 1 And TRL_CCAA < 4 Then

    If powc_CCAA < 50 Then pow_CCAA = (1 + Con1_power50) * powc_CCAA
    If powc_CCAA > 50 And powc_CCAA < 500 Then pow_CCAA = (1 + Con1_power500) * powc_CCAA
    If powc_CCAA > 500 And powc_CCAA < 2500 Then pow_CCAA = (1 + Con1_power2500) * powc_CCAA
    If powc_CCAA > 2500 Then pow_CCAA = (1 + Con1_power2500_) * powc_CCAA

End If

If TRL_CCAA >= 4 And TRL_CCAA < 7 Then

    If powc_CCAA < 50 Then pow_CCAA = (1 + Con2_power50) * powc_CCAA
    If powc_CCAA > 50 And powc_CCAA < 500 Then pow_CCAA = (1 + Con2_power500) * powc_CCAA
    If powc_CCAA > 500 And powc_CCAA < 2500 Then pow_CCAA = (1 + Con2_power2500) * powc_CCAA
    If powc_CCAA > 2500 Then pow_CCAA = (1 + Con2_power2500_) * powc_CCAA

End If

If TRL_CCAA >= 7 And TRL_CCAA < 10 Then

    If powc_CCAA < 50 Then pow_CCAA = (1 + Con3_power50) * powc_CCAA
    If powc_CCAA > 50 And powc_CCAA < 500 Then pow_CCAA = (1 + Con3_power500) * powc_CCAA
    If powc_CCAA > 500 And powc_CCAA < 2500 Then pow_CCAA = (1 + Con3_power2500) * powc_CCAA
    If powc_CCAA > 2500 Then pow_CCAA = (1 + Con3_power2500_) * powc_CCAA

End If

AES:

'Data given at 100% waterrecovery for urine and brine water (=sweat/vapor) input of 7.64 kg/d

Dim SF_AES 'Scaling Faktor
Dim RF_AES 'Rescale factor for AES
Dim massg_AES 'given system mass for AES
Dim massreg_AES 'given resupply mass for AES
Dim volg_AES 'given volume for AES
Dim powg_AES 'given power for AES
Dim coolg_AES ' given cooling for AES
Dim WW_out 'given waste water output
Dim TRL_AES
Dim eff_AES

Dim mass_AES
Dim massre_AES
Dim vol_AES
Dim pow_AES
Dim cool_AES
Dim massc_AES
Dim powc_AES

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_AES = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_AES = 0.9
If Crew_Size >= 10 Then SF_AES = 0.8

massg_AES = 45.3
massreg_AES = 0.071
volg_AES = 0.15
powg_AES = 578
coolg_AES = 577
WW_out = 7.64
TRL_AES = 5
eff_AES = 1

RF_AES = ((urine + swe_vapr) * (1 + Margin)) / WW_out

massc_AES = RF_AES * massg_AES * Crew_Size * SF_AES

If TRL_AES > 1 And TRL_AES < 4 Then

    If massc_AES < 50 Then mass_AES = (1 + Con1_mass50) * massc_AES
    If massc_AES > 50 And massc_AES < 500 Then mass_AES = (1 + Con1_mass500) * massc_AES
    If massc_AES > 500 And massc_AES < 2500 Then mass_AES = (1 + Con1_mass2500) * massc_AES
    If massc_AES > 2500 Then mass_AES = (1 + Con1_mass2500_) * massc_AES

End If

If TRL_AES >= 4 And TRL_AES < 7 Then

    If massc_AES < 50 Then mass_AES = (1 + Con2_mass50) * massc_AES
    If massc_AES > 50 And massc_AES < 500 Then mass_AES = (1 + Con2_mass500) * massc_AES
    If massc_AES > 500 And massc_AES < 2500 Then mass_AES = (1 + Con2_mass2500) * massc_AES
    If massc_AES > 2500 Then mass_AES = (1 + Con2_mass2500_) * massc_AES

End If

If TRL_AES >= 7 And TRL_AES < 10 Then

    If massc_AES < 50 Then mass_AES = (1 + Con3_mass50) * massc_AES
    If massc_AES > 50 And massc_AES < 500 Then mass_AES = (1 + Con3_mass500) * massc_AES
    If massc_AES > 500 And massc_AES < 2500 Then mass_AES = (1 + Con3_mass2500) * massc_AES
    If massc_AES > 2500 Then mass_AES = (1 + Con3_mass2500_) * massc_AES

End If

massre_AES = RF_AES * massreg_AES * Crew_Size * SF_AES * mission_duration
vol_AES = RF_AES * volg_AES * Crew_Size * SF_AES
cool_AES = RF_AES * coolg_AES * Crew_Size * SF_AES
powc_AES = RF_AES * powg_AES * Crew_Size * SF_AES

If TRL_AES > 1 And TRL_AES < 4 Then

    If powc_AES < 50 Then pow_AES = (1 + Con1_power50) * powc_AES
    If powc_AES > 50 And powc_AES < 500 Then pow_AES = (1 + Con1_power500) * powc_AES
    If powc_AES > 500 And powc_AES < 2500 Then pow_AES = (1 + Con1_power2500) * powc_AES
    If powc_AES > 2500 Then pow_AES = (1 + Con1_power2500_) * powc_AES

End If

If TRL_AES >= 4 And TRL_AES < 7 Then

    If powc_AES < 50 Then pow_AES = (1 + Con2_power50) * powc_AES
    If powc_AES > 50 And powc_AES < 500 Then pow_AES = (1 + Con2_power500) * powc_AES
    If powc_AES > 500 And powc_AES < 2500 Then pow_AES = (1 + Con2_power2500) * powc_AES
    If powc_AES > 2500 Then pow_AES = (1 + Con2_power2500_) * powc_AES

End If

If TRL_AES >= 7 And TRL_AES < 10 Then

    If powc_AES < 50 Then pow_AES = (1 + Con3_power50) * powc_AES
    If powc_AES > 50 And powc_AES < 500 Then pow_AES = (1 + Con3_power500) * powc_AES
    If powc_AES > 500 And powc_AES < 2500 Then pow_AES = (1 + Con3_power2500) * powc_AES
    If powc_AES > 2500 Then pow_AES = (1 + Con3_power2500_) * powc_AES

End If

'_______________________________________________________________________________________________________

Bosch:

'Data given at input data of 4 kg/d CO2 removal. CO2 + 2H2 -> C + 2H2O

Dim SF_Bs 'Scaling Faktor
Dim RF_Bs 'Rescale factor for Bosch
Dim massg_Bs 'given system mass for Bosch
Dim massreg_Bs 'given resupply mass for Bosch
Dim volg_Bs 'given volume for Bosch
Dim powg_Bs 'given power for Bosch
Dim coolg_Bs ' given cooling for Bosch
Dim CO2_in_Bs 'given CO2 input
Dim H2_in_Bs 'needed H2 input
Dim H2O_out_Bs 'produced H2O output
Dim TRL_Bs
Dim eff_Bs

Dim mass_Bs
Dim massre_Bs
Dim vol_Bs
Dim pow_Bs
Dim cool_Bs
Dim massc_Bs
Dim powc_Bs

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_Bs = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_Bs = 0.9
If Crew_Size >= 10 Then SF_Bs = 0.8

eff_Bs = 1
massg_Bs = 102.1
massreg_Bs = 0.028
volg_Bs = 0.3
powg_Bs = 950
coolg_Bs = 313
CO2_in_Bs = 4
TRL_Bs = 6

RF_Bs = (CO2_Pro * (1 + Margin)) / CO2_in_Bs
H2_in_Bs = 2 * MM_H2 / MM_CO2 * CO2_Pro * (1 + Margin) * Crew_Size
H2O_out_Bs = 2 * MM_H2O / MM_CO2 * CO2_Pro * (1 + Margin) * Crew_Size * eff_Bs

massc_Bs = RF_Bs * (massg_Bs + m_CO2rd_a) * Crew_Size * SF_Bs

If TRL_Bs > 1 And TRL_Bs < 4 Then

    If massc_Bs < 50 Then mass_Bs = (1 + Con1_mass50) * massc_Bs
    If massc_Bs > 50 And massc_Bs < 500 Then mass_Bs = (1 + Con1_mass500) * massc_Bs
    If massc_Bs > 500 And massc_Bs < 2500 Then mass_Bs = (1 + Con1_mass2500) * massc_Bs
    If massc_Bs > 2500 Then mass_Bs = (1 + Con1_mass2500_) * massc_Bs

End If

If TRL_Bs >= 4 And TRL_Bs < 7 Then

    If massc_Bs < 50 Then mass_Bs = (1 + Con2_mass50) * massc_Bs
    If massc_Bs > 50 And massc_Bs < 500 Then mass_Bs = (1 + Con2_mass500) * massc_Bs
    If massc_Bs > 500 And massc_Bs < 2500 Then mass_Bs = (1 + Con2_mass2500) * massc_Bs
    If massc_Bs > 2500 Then mass_Bs = (1 + Con2_mass2500_) * massc_Bs

End If

If TRL_Bs >= 7 And TRL_Bs < 10 Then

    If massc_Bs < 50 Then mass_Bs = (1 + Con3_mass50) * massc_Bs
    If massc_Bs > 50 And massc_Bs < 500 Then mass_Bs = (1 + Con3_mass500) * massc_Bs
    If massc_Bs > 500 And massc_Bs < 2500 Then mass_Bs = (1 + Con3_mass2500) * massc_Bs
    If massc_Bs > 2500 Then mass_Bs = (1 + Con3_mass2500_) * massc_Bs

End If

massre_Bs = RF_Bs * massreg_Bs * Crew_Size * SF_Bs * mission_duration
vol_Bs = RF_Bs * volg_Bs * Crew_Size * SF_Bs
cool_Bs = RF_Bs * coolg_Bs * Crew_Size * SF_Bs
powc_Bs = RF_Bs * powg_Bs * Crew_Size * SF_Bs

If TRL_Bs > 1 And TRL_Bs < 4 Then

    If powc_Bs < 50 Then pow_Bs = (1 + Con1_power50) * powc_Bs
    If powc_Bs > 50 And powc_Bs < 500 Then pow_Bs = (1 + Con1_power500) * powc_Bs
    If powc_Bs > 500 And powc_Bs < 2500 Then pow_Bs = (1 + Con1_power2500) * powc_Bs
    If powc_Bs > 2500 Then pow_Bs = (1 + Con1_power2500_) * powc_Bs

End If

If TRL_Bs >= 4 And TRL_Bs < 7 Then

    If powc_Bs < 50 Then pow_Bs = (1 + Con2_power50) * powc_Bs
    If powc_Bs > 50 And powc_Bs < 500 Then pow_Bs = (1 + Con2_power500) * powc_Bs
    If powc_Bs > 500 And powc_Bs < 2500 Then pow_Bs = (1 + Con2_power2500) * powc_Bs
    If powc_Bs > 2500 Then pow_Bs = (1 + Con2_power2500_) * powc_Bs

End If

If TRL_Bs >= 7 And TRL_Bs < 10 Then

    If powc_Bs < 50 Then pow_Bs = (1 + Con3_power50) * powc_Bs
    If powc_Bs > 50 And powc_Bs < 500 Then pow_Bs = (1 + Con3_power500) * powc_Bs
    If powc_Bs > 500 And powc_Bs < 2500 Then pow_Bs = (1 + Con3_power2500) * powc_Bs
    If powc_Bs > 2500 Then pow_Bs = (1 + Con3_power2500_) * powc_Bs

End If

'______________________________________________________________________________________________________________

CAMRAS:

'No precise data given as of right now. Only mass, power, volume is stated

Dim SF_CAMRAS 'Scaling Faktor
Dim RF_CAMRAS 'is currently set to 1 due to lack of further information
Dim massg_CAMRAS 'given system mass for CAMRAS
Dim massreg_CAMRAS 'given resupply mass for CAMRAS
Dim volg_CAMRAS 'given volume for CAMRAS
Dim powg_CAMRAS 'given power for CAMRAS
Dim coolg_CAMRAS ' given cooling for CAMRAS
Dim CO2_in_CAMRAS 'given CO2 input
Dim TRL_CAMRAS

Dim mass_CAMRAS
Dim massre_CAMRAS
Dim vol_CAMRAS
Dim pow_CAMRAS
Dim cool_CAMRAS
Dim massc_CAMRAS
Dim powc_CAMRAS

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_CAMRAS = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_CAMRAS = 0.9
If Crew_Size >= 10 Then SF_CAMRAS = 0.8

RF_CAMRAS = 1
massg_CAMRAS = 7.26
massreg_CAMRAS = 0
volg_CAMRAS = 0.3
powg_CAMRAS = 12
coolg_CAMRAS = 0
TRL_CAMRAS = 9

massc_CAMRAS = RF_CAMRAS * (massg_CAMRAS + m_CO2rm_a) * Crew_Size * SF_CAMRAS

If TRL_CAMRAS > 1 And TRL_CAMRAS < 4 Then

    If massc_CAMRAS < 50 Then mass_CAMRAS = (1 + Con1_mass50) * massc_CAMRAS
    If massc_CAMRAS > 50 And massc_CAMRAS < 500 Then mass_CAMRAS = (1 + Con1_mass500) * massc_CAMRAS
    If massc_CAMRAS > 500 And massc_CAMRAS < 2500 Then mass_CAMRAS = (1 + Con1_mass2500) * massc_CAMRAS
    If massc_CAMRAS > 2500 Then mass_CAMRAS = (1 + Con1_mass2500_) * massc_CAMRAS

End If

If TRL_CAMRAS >= 4 And TRL_CAMRAS < 7 Then

    If massc_CAMRAS < 50 Then mass_CAMRAS = (1 + Con2_mass50) * massc_CAMRAS
    If massc_CAMRAS > 50 And massc_CAMRAS < 500 Then mass_CAMRAS = (1 + Con2_mass500) * massc_CAMRAS
    If massc_CAMRAS > 500 And massc_CAMRAS < 2500 Then mass_CAMRAS = (1 + Con2_mass2500) * massc_CAMRAS
    If massc_CAMRAS > 2500 Then mass_CAMRAS = (1 + Con2_mass2500_) * massc_CAMRAS

End If

If TRL_CAMRAS >= 7 And TRL_CAMRAS < 10 Then

    If massc_CAMRAS < 50 Then mass_CAMRAS = (1 + Con3_mass50) * massc_CAMRAS
    If massc_CAMRAS > 50 And massc_CAMRAS < 500 Then mass_CAMRAS = (1 + Con3_mass500) * massc_CAMRAS
    If massc_CAMRAS > 500 And massc_CAMRAS < 2500 Then mass_CAMRAS = (1 + Con3_mass2500) * massc_CAMRAS
    If massc_CAMRAS > 2500 Then mass_CAMRAS = (1 + Con3_mass2500_) * massc_CAMRAS

End If

massre_CAMRAS = RF_CAMRAS * massreg_CAMRAS * Crew_Size * SF_CAMRAS * mission_duration
vol_CAMRAS = RF_CAMRAS * volg_CAMRAS * Crew_Size * SF_CAMRAS
cool_CAMRAS = RF_CAMRAS * coolg_CAMRAS * Crew_Size * SF_CAMRAS
powc_CAMRAS = RF_CAMRAS * powg_CAMRAS * Crew_Size * SF_CAMRAS

If TRL_CAMRAS > 1 And TRL_CAMRAS < 4 Then

    If powc_CAMRAS < 50 Then pow_CAMRAS = (1 + Con1_power50) * powc_CAMRAS
    If powc_CAMRAS > 50 And powc_CAMRAS < 500 Then pow_CAMRAS = (1 + Con1_power500) * powc_CAMRAS
    If powc_CAMRAS > 500 And powc_CAMRAS < 2500 Then pow_CAMRAS = (1 + Con1_power2500) * powc_CAMRAS
    If powc_CAMRAS > 2500 Then pow_CAMRAS = (1 + Con1_power2500_) * powc_CAMRAS

End If

If TRL_CAMRAS >= 4 And TRL_CAMRAS < 7 Then

    If powc_CAMRAS < 50 Then pow_CAMRAS = (1 + Con2_power50) * powc_CAMRAS
    If powc_CAMRAS > 50 And powc_CAMRAS < 500 Then pow_CAMRAS = (1 + Con2_power500) * powc_CAMRAS
    If powc_CAMRAS > 500 And powc_CAMRAS < 2500 Then pow_CAMRAS = (1 + Con2_power2500) * powc_CAMRAS
    If powc_CAMRAS > 2500 Then pow_CAMRAS = (1 + Con2_power2500_) * powc_CAMRAS

End If

If TRL_CAMRAS >= 7 And TRL_CAMRAS < 10 Then

    If powc_CAMRAS < 50 Then pow_CAMRAS = (1 + Con3_power50) * powc_CAMRAS
    If powc_CAMRAS > 50 And powc_CAMRAS < 500 Then pow_CAMRAS = (1 + Con3_power500) * powc_CAMRAS
    If powc_CAMRAS > 500 And powc_CAMRAS < 2500 Then pow_CAMRAS = (1 + Con3_power2500) * powc_CAMRAS
    If powc_CAMRAS > 2500 Then pow_CAMRAS = (1 + Con3_power2500_) * powc_CAMRAS

End If
    
'____________________________________________________________________________________________________________________________


COE:

'Data is given for 36.8 kg/d CO2, resupply mass is not provided

Dim SF_COE 'Scaling Faktor
Dim RF_COE 'is currently set to 1 due to lack of further information
Dim massg_COE 'given system mass for COE
Dim massreg_COE 'given resupply mass for COE
Dim volg_COE 'given volume for COE
Dim powg_COE 'given power for COE
Dim coolg_COE ' given cooling for COE
Dim CO2_in_COE 'given CO2 input
Dim O2_out_COE 'produced O2 output
Dim eff_COE 'unknown
Dim TRL_COE

Dim mass_COE
Dim massre_COE
Dim vol_COE
Dim pow_COE
Dim cool_COE
Dim massc_COE
Dim powc_COE

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_COE = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_COE = 0.9
If Crew_Size >= 10 Then SF_COE = 0.8

massg_COE = 356
massreg_COE = 0
volg_COE = 0.28
powg_COE = 13300
coolg_COE = 3500
TRL_COE = 5
CO2_in_COE = 36.8

RF_COE = (CO2_Pro * (1 + Margin)) / CO2_in_COE
O2_out_COE = MM_O2 / MM_CO2 * CO2_Pro * Crew_Size * eff_COE
massc_COE = RF_COE * (massg_COE + m_CO2rd_a) * Crew_Size * SF_COE

If TRL_COE > 1 And TRL_COE < 4 Then

    If massc_COE < 50 Then mass_COE = (1 + Con1_mass50) * massc_COE
    If massc_COE > 50 And massc_COE < 500 Then mass_COE = (1 + Con1_mass500) * massc_COE
    If massc_COE > 500 And massc_COE < 2500 Then mass_COE = (1 + Con1_mass2500) * massc_COE
    If massc_COE > 2500 Then mass_COE = (1 + Con1_mass2500_) * massc_COE

End If

If TRL_COE >= 4 And TRL_COE < 7 Then

    If massc_COE < 50 Then mass_COE = (1 + Con2_mass50) * massc_COE
    If massc_COE > 50 And massc_COE < 500 Then mass_COE = (1 + Con2_mass500) * massc_COE
    If massc_COE > 500 And massc_COE < 2500 Then mass_COE = (1 + Con2_mass2500) * massc_COE
    If massc_COE > 2500 Then mass_COE = (1 + Con2_mass2500_) * massc_COE

End If

If TRL_COE >= 7 And TRL_COE < 10 Then

    If massc_COE < 50 Then mass_COE = (1 + Con3_mass50) * massc_COE
    If massc_COE > 50 And massc_COE < 500 Then mass_COE = (1 + Con3_mass500) * massc_COE
    If massc_COE > 500 And massc_COE < 2500 Then mass_COE = (1 + Con3_mass2500) * massc_COE
    If massc_COE > 2500 Then mass_COE = (1 + Con3_mass2500_) * massc_COE

End If

massre_COE = 0 'No information as of right now
vol_COE = RF_COE * volg_COE * Crew_Size * SF_COE
cool_COE = RF_COE * coolg_COE * Crew_Size * SF_COE
powc_COE = RF_COE * powg_COE * Crew_Size * SF_COE

If TRL_COE > 1 And TRL_COE < 4 Then

    If powc_COE < 50 Then pow_COE = (1 + Con1_power50) * powc_COE
    If powc_COE > 50 And powc_COE < 500 Then pow_COE = (1 + Con1_power500) * powc_COE
    If powc_COE > 500 And powc_COE < 2500 Then pow_COE = (1 + Con1_power2500) * powc_COE
    If powc_COE > 2500 Then pow_COE = (1 + Con1_power2500_) * powc_COE

End If

If TRL_COE >= 4 And TRL_COE < 7 Then

    If powc_COE < 50 Then pow_COE = (1 + Con2_power50) * powc_COE
    If powc_COE > 50 And powc_COE < 500 Then pow_COE = (1 + Con2_power500) * powc_COE
    If powc_COE > 500 And powc_COE < 2500 Then pow_COE = (1 + Con2_power2500) * powc_COE
    If powc_COE > 2500 Then pow_COE = (1 + Con2_power2500_) * powc_COE

End If

If TRL_COE >= 7 And TRL_COE < 10 Then

    If powc_COE < 50 Then pow_COE = (1 + Con3_power50) * powc_COE
    If powc_COE > 50 And powc_COE < 500 Then pow_COE = (1 + Con3_power500) * powc_COE
    If powc_COE > 500 And powc_COE < 2500 Then pow_COE = (1 + Con3_power2500) * powc_COE
    If powc_COE > 2500 Then pow_COE = (1 + Con3_power2500_) * powc_COE

End If

'____________________________________________________________________________________________________________


EDC:

'Data given at input data of 4 kg/d CO2 removal, additional input H2, O2, additional output H20

Dim SF_EDC 'Scaling Faktor
Dim RF_EDC 'Rescale factor for Bosch
Dim massg_EDC 'given system mass for Bosch
Dim massreg_EDC 'given resupply mass for Bosch
Dim volg_EDC 'given volume for Bosch
Dim powg_EDC 'given power for Bosch
Dim coolg_EDC ' given cooling for Bosch
Dim CO2_in_EDC 'given CO2 input
Dim TRL_EDC
Dim H2_in_EDC 'needed H2 input
Dim O2_in_EDC 'needed O2 input
Dim H2O_out_EDC 'H20 output
Dim CO2_out_EDC 'CO2 output (not in atmosphere anymore)
Dim eff_EDC 'Unknown

Dim mass_EDC
Dim massre_EDC
Dim vol_EDC
Dim pow_EDC
Dim cool_EDC
Dim massc_EDC
Dim powc_EDC

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_EDC = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_EDC = 0.9
If Crew_Size >= 10 Then SF_EDC = 0.8

eff_EDC = 1
massg_EDC = 44.4
massreg_EDC = 0.012
volg_EDC = 0.07
powg_EDC = 42
coolg_EDC = 336
CO2_in_EDC = 4
TRL_EDC = 6

RF_EDC = (CO2_Pro * (1 + Margin)) / CO2_in_EDC
massc_EDC = RF_EDC * (massg_EDC + m_CO2rm_a) * Crew_Size * SF_EDC
O2_in_EDC = 0.5 * MM_O2 / MM_CO2 * Crew_Size * CO2_Pro
H2_in_EDC = MM_H2 / MM_CO2 * Crew_Size * CO2_Pro
H2O_out_EDC = MM_H2O / MM_CO2 * Crew_Size * CO2_Pro * eff_EDC
CO2_out_EDC = CO2_Pro * Crew_Size * eff_EDC

If TRL_EDC > 1 And TRL_EDC < 4 Then

    If massc_EDC < 50 Then mass_EDC = (1 + Con1_mass50) * massc_EDC
    If massc_EDC > 50 And massc_EDC < 500 Then mass_EDC = (1 + Con1_mass500) * massc_EDC
    If massc_EDC > 500 And massc_EDC < 2500 Then mass_EDC = (1 + Con1_mass2500) * massc_EDC
    If massc_EDC > 2500 Then mass_EDC = (1 + Con1_mass2500_) * massc_EDC

End If

If TRL_EDC >= 4 And TRL_EDC < 7 Then

    If massc_EDC < 50 Then mass_EDC = (1 + Con2_mass50) * massc_EDC
    If massc_EDC > 50 And massc_EDC < 500 Then mass_EDC = (1 + Con2_mass500) * massc_EDC
    If massc_EDC > 500 And massc_EDC < 2500 Then mass_EDC = (1 + Con2_mass2500) * massc_EDC
    If massc_EDC > 2500 Then mass_EDC = (1 + Con2_mass2500_) * massc_EDC

End If

If TRL_EDC >= 7 And TRL_EDC < 10 Then

    If massc_EDC < 50 Then mass_EDC = (1 + Con3_mass50) * massc_EDC
    If massc_EDC > 50 And massc_EDC < 500 Then mass_EDC = (1 + Con3_mass500) * massc_EDC
    If massc_EDC > 500 And massc_EDC < 2500 Then mass_EDC = (1 + Con3_mass2500) * massc_EDC
    If massc_EDC > 2500 Then mass_EDC = (1 + Con3_mass2500_) * massc_EDC

End If

massre_EDC = RF_EDC * massreg_EDC * Crew_Size * SF_EDC * mission_duration
vol_EDC = RF_EDC * volg_EDC * Crew_Size * SF_EDC
cool_EDC = RF_EDC * coolg_EDC * Crew_Size * SF_EDC
powc_EDC = RF_EDC * powg_EDC * Crew_Size * SF_EDC

If TRL_EDC > 1 And TRL_EDC < 4 Then

    If powc_EDC < 50 Then pow_EDC = (1 + Con1_power50) * powc_EDC
    If powc_EDC > 50 And powc_EDC < 500 Then pow_EDC = (1 + Con1_power500) * powc_EDC
    If powc_EDC > 500 And powc_EDC < 2500 Then pow_EDC = (1 + Con1_power2500) * powc_EDC
    If powc_EDC > 2500 Then pow_EDC = (1 + Con1_power2500_) * powc_EDC

End If

If TRL_EDC >= 4 And TRL_EDC < 7 Then

    If powc_EDC < 50 Then pow_EDC = (1 + Con2_power50) * powc_EDC
    If powc_EDC > 50 And powc_EDC < 500 Then pow_EDC = (1 + Con2_power500) * powc_EDC
    If powc_EDC > 500 And powc_EDC < 2500 Then pow_EDC = (1 + Con2_power2500) * powc_EDC
    If powc_EDC > 2500 Then pow_EDC = (1 + Con2_power2500_) * powc_EDC

End If

If TRL_EDC >= 7 And TRL_EDC < 10 Then

    If powc_EDC < 50 Then pow_EDC = (1 + Con3_power50) * powc_EDC
    If powc_EDC > 50 And powc_EDC < 500 Then pow_EDC = (1 + Con3_power500) * powc_EDC
    If powc_EDC > 500 And powc_EDC < 2500 Then pow_EDC = (1 + Con3_power2500) * powc_EDC
    If powc_EDC > 2500 Then pow_EDC = (1 + Con3_power2500_) * powc_EDC

End If

'___________________________________________________________________________________________________________

'2BMS:

'Data given for 3 crew (=3*CO2_Pro*(1+Margin)

Dim SF_2BMS 'Scaling Faktor
Dim RF_2BMS 'Rescale factor for Bosch
Dim massg_2BMS 'given system mass for Bosch
Dim massreg_2BMS 'given resupply mass for Bosch
Dim volg_2BMS 'given volume for Bosch
Dim powg_2BMS 'given power for Bosch
Dim coolg_2BMS ' given cooling for Bosch
Dim CO2_in_2BMS 'given CO2 input
Dim TRL_2BMS

Dim mass_2BMS
Dim massre_2BMS
Dim vol_2BMS
Dim pow_2BMS
Dim cool_2BMS
Dim massc_2BMS
Dim powc_2BMS

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_2BMS = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_2BMS = 0.9
If Crew_Size >= 10 Then SF_2BMS = 0.8

massg_2BMS = 48.1
massreg_2BMS = 0.02 'Assumption: similar to 4BMS
volg_2BMS = 0.26
powg_2BMS = 230
coolg_2BMS = 230
CO2_in_2BMS = 3 * CO2_Pro * (1 + Margin)
TRL_2BMS = 4

RF_2BMS = (CO2_Pro * (1 + Margin)) / CO2_in_2BMS
massc_2BMS = RF_2BMS * (massg_2BMS + m_CO2rm_a) * Crew_Size * SF_2BMS

If TRL_2BMS > 1 And TRL_2BMS < 4 Then

    If massc_2BMS < 50 Then mass_2BMS = (1 + Con1_mass50) * massc_2BMS
    If massc_2BMS > 50 And massc_2BMS < 500 Then mass_2BMS = (1 + Con1_mass500) * massc_2BMS
    If massc_2BMS > 500 And massc_2BMS < 2500 Then mass_2BMS = (1 + Con1_mass2500) * massc_2BMS
    If massc_2BMS > 2500 Then mass_2BMS = (1 + Con1_mass2500_) * massc_2BMS

End If

If TRL_2BMS >= 4 And TRL_2BMS < 7 Then

    If massc_2BMS < 50 Then mass_2BMS = (1 + Con2_mass50) * massc_2BMS
    If massc_2BMS > 50 And massc_2BMS < 500 Then mass_2BMS = (1 + Con2_mass500) * massc_2BMS
    If massc_2BMS > 500 And massc_2BMS < 2500 Then mass_2BMS = (1 + Con2_mass2500) * massc_2BMS
    If massc_2BMS > 2500 Then mass_2BMS = (1 + Con2_mass2500_) * massc_2BMS

End If

If TRL_2BMS >= 7 And TRL_2BMS < 10 Then

    If massc_2BMS < 50 Then mass_2BMS = (1 + Con3_mass50) * massc_2BMS
    If massc_2BMS > 50 And massc_2BMS < 500 Then mass_2BMS = (1 + Con3_mass500) * massc_2BMS
    If massc_2BMS > 500 And massc_2BMS < 2500 Then mass_2BMS = (1 + Con3_mass2500) * massc_2BMS
    If massc_2BMS > 2500 Then mass_2BMS = (1 + Con3_mass2500_) * massc_2BMS

End If

massre_2BMS = RF_2BMS * massreg_2BMS * Crew_Size * SF_2BMS * mission_duration
vol_2BMS = RF_2BMS * volg_2BMS * Crew_Size * SF_2BMS
cool_2BMS = RF_2BMS * coolg_2BMS * Crew_Size * SF_2BMS
powc_2BMS = RF_2BMS * powg_2BMS * Crew_Size * SF_2BMS

If TRL_2BMS > 1 And TRL_2BMS < 4 Then

    If powc_2BMS < 50 Then pow_2BMS = (1 + Con1_power50) * powc_2BMS
    If powc_2BMS > 50 And powc_2BMS < 500 Then pow_2BMS = (1 + Con1_power500) * powc_2BMS
    If powc_2BMS > 500 And powc_2BMS < 2500 Then pow_2BMS = (1 + Con1_power2500) * powc_2BMS
    If powc_2BMS > 2500 Then pow_2BMS = (1 + Con1_power2500_) * powc_2BMS

End If

If TRL_2BMS >= 4 And TRL_2BMS < 7 Then

    If powc_2BMS < 50 Then pow_2BMS = (1 + Con2_power50) * powc_2BMS
    If powc_2BMS > 50 And powc_2BMS < 500 Then pow_2BMS = (1 + Con2_power500) * powc_2BMS
    If powc_2BMS > 500 And powc_2BMS < 2500 Then pow_2BMS = (1 + Con2_power2500) * powc_2BMS
    If powc_2BMS > 2500 Then pow_2BMS = (1 + Con2_power2500_) * powc_2BMS

End If

If TRL_2BMS >= 7 And TRL_2BMS < 10 Then

    If powc_2BMS < 50 Then pow_2BMS = (1 + Con3_power50) * powc_2BMS
    If powc_2BMS > 50 And powc_2BMS < 500 Then pow_2BMS = (1 + Con3_power500) * powc_2BMS
    If powc_2BMS > 500 And powc_2BMS < 2500 Then pow_2BMS = (1 + Con3_power2500) * powc_2BMS
    If powc_2BMS > 2500 Then pow_2BMS = (1 + Con3_power2500_) * powc_2BMS

End If

'___________________________________________________________________________________________________________

'4BMS:

'Data given at input data of 4 kg/d CO2 removal

Dim SF_4BMS 'Scaling Faktor
Dim RF_4BMS 'Rescale factor for Bosch
Dim massg_4BMS 'given system mass for Bosch
Dim massreg_4BMS 'given resupply mass for Bosch
Dim volg_4BMS 'given volume for Bosch
Dim powg_4BMS 'given power for Bosch
Dim coolg_4BMS ' given cooling for Bosch
Dim CO2_in_4BMS 'given CO2 input
Dim TRL_4BMS

Dim mass_4BMS
Dim massre_4BMS
Dim vol_4BMS
Dim pow_4BMS
Dim cool_4BMS
Dim massc_4BMS
Dim powc_4BMS

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_4BMS = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_4BMS = 0.9
If Crew_Size >= 10 Then SF_4BMS = 0.8

massg_4BMS = 87.9
massreg_4BMS = 0.024
volg_4BMS = 0.33
powg_4BMS = 535
coolg_4BMS = 535
CO2_in_4BMS = 3
TRL_4BMS = 9

RF_4BMS = (CO2_Pro * (1 + Margin)) / CO2_in_4BMS
massc_4BMS = RF_4BMS * (massg_4BMS + m_CO2rm_a) * Crew_Size * SF_4BMS

If TRL_4BMS > 1 And TRL_4BMS < 4 Then

    If massc_4BMS < 50 Then mass_4BMS = (1 + Con1_mass50) * massc_4BMS
    If massc_4BMS > 50 And massc_4BMS < 500 Then mass_4BMS = (1 + Con1_mass500) * massc_4BMS
    If massc_4BMS > 500 And massc_4BMS < 2500 Then mass_4BMS = (1 + Con1_mass2500) * massc_4BMS
    If massc_4BMS > 2500 Then mass_4BMS = (1 + Con1_mass2500_) * massc_4BMS

End If

If TRL_4BMS >= 4 And TRL_4BMS < 7 Then

    If massc_4BMS < 50 Then mass_4BMS = (1 + Con2_mass50) * massc_4BMS
    If massc_4BMS > 50 And massc_4BMS < 500 Then mass_4BMS = (1 + Con2_mass500) * massc_4BMS
    If massc_4BMS > 500 And massc_4BMS < 2500 Then mass_4BMS = (1 + Con2_mass2500) * massc_4BMS
    If massc_4BMS > 2500 Then mass_4BMS = (1 + Con2_mass2500_) * massc_4BMS

End If

If TRL_4BMS >= 7 And TRL_4BMS < 10 Then

    If massc_4BMS < 50 Then mass_4BMS = (1 + Con3_mass50) * massc_4BMS
    If massc_4BMS > 50 And massc_4BMS < 500 Then mass_4BMS = (1 + Con3_mass500) * massc_4BMS
    If massc_4BMS > 500 And massc_4BMS < 2500 Then mass_4BMS = (1 + Con3_mass2500) * massc_4BMS
    If massc_4BMS > 2500 Then mass_4BMS = (1 + Con3_mass2500_) * massc_4BMS

End If

massre_4BMS = RF_4BMS * massreg_4BMS * Crew_Size * SF_4BMS * mission_duration
vol_4BMS = RF_4BMS * volg_4BMS * Crew_Size * SF_4BMS
cool_4BMS = RF_4BMS * coolg_4BMS * Crew_Size * SF_4BMS
powc_4BMS = RF_4BMS * powg_4BMS * Crew_Size * SF_4BMS

If TRL_4BMS > 1 And TRL_4BMS < 4 Then

    If powc_4BMS < 50 Then pow_4BMS = (1 + Con1_power50) * powc_4BMS
    If powc_4BMS > 50 And powc_4BMS < 500 Then pow_4BMS = (1 + Con1_power500) * powc_4BMS
    If powc_4BMS > 500 And powc_4BMS < 2500 Then pow_4BMS = (1 + Con1_power2500) * powc_4BMS
    If powc_4BMS > 2500 Then pow_4BMS = (1 + Con1_power2500_) * powc_4BMS

End If

If TRL_4BMS >= 4 And TRL_4BMS < 7 Then

    If powc_4BMS < 50 Then pow_4BMS = (1 + Con2_power50) * powc_4BMS
    If powc_4BMS > 50 And powc_4BMS < 500 Then pow_4BMS = (1 + Con2_power500) * powc_4BMS
    If powc_4BMS > 500 And powc_4BMS < 2500 Then pow_4BMS = (1 + Con2_power2500) * powc_4BMS
    If powc_4BMS > 2500 Then pow_4BMS = (1 + Con2_power2500_) * powc_4BMS

End If

If TRL_4BMS >= 7 And TRL_4BMS < 10 Then

    If powc_4BMS < 50 Then pow_4BMS = (1 + Con3_power50) * powc_4BMS
    If powc_4BMS > 50 And powc_4BMS < 500 Then pow_4BMS = (1 + Con3_power500) * powc_4BMS
    If powc_4BMS > 500 And powc_4BMS < 2500 Then pow_4BMS = (1 + Con3_power2500) * powc_4BMS
    If powc_4BMS > 2500 Then pow_4BMS = (1 + Con3_power2500_) * powc_4BMS

End If

'________________________________________________________________________________________________________________________

MF:

'Data given for waterrecovery for waste water input of 50.9 kg/d
'Waste water (=assumption: hygiene + ww from urine process+potential other component -> unkown yet, set to 0) in and H2O out (useable for hygiene and drinking)

Dim SF_MF 'Scaling Faktor
Dim RF_MF 'Rescale factor for MF
Dim massg_MF 'given system mass for MF
Dim massreg_MF 'given resupply mass for MF
Dim volg_MF 'given volume for MF
Dim powg_MF 'given power for MF
Dim coolg_MF ' given cooling for MF
Dim H2O_out_MF 'given waste water output
Dim ww_other_MF 'other unknown ww component -> 0 as of right now
Dim urine_WW
Dim TRL_MF

Dim mass_MF
Dim massre_MF
Dim vol_MF
Dim pow_MF
Dim cool_MF
Dim massc_MF
Dim powc_MF

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_MF = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_MF = 0.9
If Crew_Size >= 10 Then SF_MF = 0.8

massg_MF = 489.1
massreg_MF = 1.64
volg_MF = 2.26
powg_MF = 1021
coolg_MF = 1021
H2O_out_MF = 50.9
ww_other_MF = 0
TRL_MF = 9

urine_WW = Worksheets("Schedule and Metabolic").Range("SUM_Urine_Production") - Worksheets("Schedule and Metabolic").Range("SUM_Urine_Solid_Waste")
RF_MF = (hygiene + ww_other_MF + urine_WW) * (1 + Margin) / H2O_out_MF
massc_MF = RF_MF * massg_MF * Crew_Size * SF_MF

If TRL_MF > 1 And TRL_MF < 4 Then

    If massc_MF < 50 Then mass_MF = (1 + Con1_mass50) * massc_MF
    If massc_MF > 50 And massc_MF < 500 Then mass_MF = (1 + Con1_mass500) * massc_MF
    If massc_MF > 500 And massc_MF < 2500 Then mass_MF = (1 + Con1_mass2500) * massc_MF
    If massc_MF > 2500 Then mass_MF = (1 + Con1_mass2500_) * massc_MF

End If

If TRL_MF >= 4 And TRL_MF < 7 Then

    If massc_MF < 50 Then mass_MF = (1 + Con2_mass50) * massc_MF
    If massc_MF > 50 And massc_MF < 500 Then mass_MF = (1 + Con2_mass500) * massc_MF
    If massc_MF > 500 And massc_MF < 2500 Then mass_MF = (1 + Con2_mass2500) * massc_MF
    If massc_MF > 2500 Then mass_MF = (1 + Con2_mass2500_) * massc_MF

End If

If TRL_MF >= 7 And TRL_MF < 10 Then

    If massc_MF < 50 Then mass_MF = (1 + Con3_mass50) * massc_MF
    If massc_MF > 50 And massc_MF < 500 Then mass_MF = (1 + Con3_mass500) * massc_MF
    If massc_MF > 500 And massc_MF < 2500 Then mass_MF = (1 + Con3_mass2500) * massc_MF
    If massc_MF > 2500 Then mass_MF = (1 + Con3_mass2500_) * massc_MF

End If

massre_MF = RF_MF * massreg_MF * Crew_Size * SF_MF * mission_duration
vol_MF = RF_MF * volg_MF * Crew_Size * SF_MF
cool_MF = RF_MF * coolg_MF * Crew_Size * SF_MF
powc_MF = RF_MF * powg_MF * Crew_Size * SF_MF

If TRL_MF > 1 And TRL_MF < 4 Then

    If powc_MF < 50 Then pow_MF = (1 + Con1_power50) * powc_MF
    If powc_MF > 50 And powc_MF < 500 Then pow_MF = (1 + Con1_power500) * powc_MF
    If powc_MF > 500 And powc_MF < 2500 Then pow_MF = (1 + Con1_power2500) * powc_MF
    If powc_MF > 2500 Then pow_MF = (1 + Con1_power2500_) * powc_MF

End If

If TRL_MF >= 4 And TRL_MF < 7 Then

    If powc_MF < 50 Then pow_MF = (1 + Con2_power50) * powc_MF
    If powc_MF > 50 And powc_MF < 500 Then pow_MF = (1 + Con2_power500) * powc_MF
    If powc_MF > 500 And powc_MF < 2500 Then pow_MF = (1 + Con2_power2500) * powc_MF
    If powc_MF > 2500 Then pow_MF = (1 + Con2_power2500_) * powc_MF

End If

If TRL_MF >= 7 And TRL_MF < 10 Then

    If powc_MF < 50 Then pow_MF = (1 + Con3_power50) * powc_MF
    If powc_MF > 50 And powc_MF < 500 Then pow_MF = (1 + Con3_power500) * powc_MF
    If powc_MF > 500 And powc_MF < 2500 Then pow_MF = (1 + Con3_power2500) * powc_MF
    If powc_MF > 2500 Then pow_MF = (1 + Con3_power2500_) * powc_MF

End If

'_____________________________________________________________________________________________________

Sabatier:

'Data given for CO2 input of 5 kg/d CO2

Dim SF_SR 'Scaling Faktor
Dim RF_SR 'Rescale factor for SR
Dim massg_SR 'given system mass for SR
Dim massreg_SR 'given resupply mass for SR
Dim volg_SR 'given volume for SR
Dim powg_SR 'given power for SR
Dim coolg_SR ' given cooling for SR
Dim CO2_in_SR 'given CO2 input
Dim H2_in_SR 'needed H2 input
Dim WW_out_SR 'generated waste water output
Dim CH4_out_SR 'generated methane output
Dim eff_SR
Dim TRL_SR

Dim mass_SR
Dim massre_SR
Dim vol_SR
Dim pow_SR
Dim cool_SR
Dim massc_SR
Dim powc_SR

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_SR = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_SR = 0.9
If Crew_Size >= 10 Then SF_SR = 0.8

eff_SR = 0.99
massg_SR = 43
massreg_SR = 0.012
volg_SR = 0.18
powg_SR = 50
coolg_SR = 289
CO2_in_SR = 5
TRL_SR = 9

RF_SR = CO2_Pro * (1 + Margin) / CO2_in_SR
massc_SR = RF_SR * (massg_SR + m_CO2rd_a) * Crew_Size * SF_SR
H2_in_SR = 4 * MM_H2 / MM_CO2 * Crew_Size * CO2_Pro
WW_out_SR = 2 * MM_H2O / MM_CO2 * Crew_Size * CO2_Pro * eff_SR
CH4_out_SR = MM_CH4 / MM_CO2 + Crew_Size * CO2_Pro * eff_SR


If TRL_SR > 1 And TRL_SR < 4 Then

    If massc_SR < 50 Then mass_SR = (1 + Con1_mass50) * massc_SR
    If massc_SR > 50 And massc_SR < 500 Then mass_SR = (1 + Con1_mass500) * massc_SR
    If massc_SR > 500 And massc_SR < 2500 Then mass_SR = (1 + Con1_mass2500) * massc_SR
    If massc_SR > 2500 Then mass_SR = (1 + Con1_mass2500_) * massc_SR

End If

If TRL_SR >= 4 And TRL_SR < 7 Then

    If massc_SR < 50 Then mass_SR = (1 + Con2_mass50) * massc_SR
    If massc_SR > 50 And massc_SR < 500 Then mass_SR = (1 + Con2_mass500) * massc_SR
    If massc_SR > 500 And massc_SR < 2500 Then mass_SR = (1 + Con2_mass2500) * massc_SR
    If massc_SR > 2500 Then mass_SR = (1 + Con2_mass2500_) * massc_SR

End If

If TRL_SR >= 7 And TRL_SR < 10 Then

    If massc_SR < 50 Then mass_SR = (1 + Con3_mass50) * massc_SR
    If massc_SR > 50 And massc_SR < 500 Then mass_SR = (1 + Con3_mass500) * massc_SR
    If massc_SR > 500 And massc_SR < 2500 Then mass_SR = (1 + Con3_mass2500) * massc_SR
    If massc_SR > 2500 Then mass_SR = (1 + Con3_mass2500_) * massc_SR

End If

massre_SR = RF_SR * massreg_SR * Crew_Size * SF_SR * mission_duration
vol_SR = RF_SR * volg_SR * Crew_Size * SF_SR
cool_SR = RF_SR * coolg_SR * Crew_Size * SF_SR
powc_SR = RF_SR * powg_SR * Crew_Size * SF_SR

If TRL_SR > 1 And TRL_SR < 4 Then

    If powc_SR < 50 Then pow_SR = (1 + Con1_power50) * powc_SR
    If powc_SR > 50 And powc_SR < 500 Then pow_SR = (1 + Con1_power500) * powc_SR
    If powc_SR > 500 And powc_SR < 2500 Then pow_SR = (1 + Con1_power2500) * powc_SR
    If powc_SR > 2500 Then pow_SR = (1 + Con1_power2500_) * powc_SR

End If

If TRL_SR >= 4 And TRL_SR < 7 Then

    If powc_SR < 50 Then pow_SR = (1 + Con2_power50) * powc_SR
    If powc_SR > 50 And powc_SR < 500 Then pow_SR = (1 + Con2_power500) * powc_SR
    If powc_SR > 500 And powc_SR < 2500 Then pow_SR = (1 + Con2_power2500) * powc_SR
    If powc_SR > 2500 Then pow_SR = (1 + Con2_power2500_) * powc_SR

End If

If TRL_SR >= 7 And TRL_SR < 10 Then

    If powc_SR < 50 Then pow_SR = (1 + Con3_power50) * powc_SR
    If powc_SR > 50 And powc_SR < 500 Then pow_SR = (1 + Con3_power500) * powc_SR
    If powc_SR > 500 And powc_SR < 2500 Then pow_SR = (1 + Con3_power2500) * powc_SR
    If powc_SR > 2500 Then pow_SR = (1 + Con3_power2500_) * powc_SR

End If

'______________________________________________________________________________________________________

SAWD:

'Data given for CO2 input of 3 kg/d CO2, requires H2O input

Dim SF_SAWD 'Scaling Faktor
Dim RF_SAWD 'Rescale factor for SAWD
Dim massg_SAWD 'given system mass for SAWD
Dim massreg_SAWD 'given resupply mass for SAWD
Dim volg_SAWD 'given volume for SAWD
Dim powg_SAWD 'given power for SAWD
Dim coolg_SAWD ' given cooling for SAWD
Dim CO2_in_SAWD 'given CO2 input
Dim H2O_in_SAWD 'needed H2O input
Dim H2O_out_SAWD 'generated H2O output = H2O input
Dim eff_SAWD 'unknown
Dim TRL_SAWD

Dim mass_SAWD
Dim massre_SAWD
Dim vol_SAWD
Dim pow_SAWD
Dim cool_SAWD
Dim massc_SAWD
Dim powc_SAWD

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_SAWD = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_SAWD = 0.9
If Crew_Size >= 10 Then SF_SAWD = 0.8

massg_SAWD = 51.3
massreg_SAWD = 0.014
volg_SAWD = 0.21
powg_SAWD = 344
coolg_SAWD = 344
CO2_in_SAWD = 3
TRL_SAWD = 8
eff_SAWD = 1

RF_SAWD = CO2_Pro * (1 + Margin) / CO2_in_SAWD
massc_SAWD = RF_SAWD * (massg_SAWD + m_CO2rm_a) * Crew_Size * SF_SAWD
H2O_in_SAWD = MM_H2O / MM_CO2 * Crew_Size * CO2_Pro
H2O_out_SAWD = H2O_in_SAWD

If TRL_SAWD > 1 And TRL_SAWD < 4 Then

    If massc_SAWD < 50 Then mass_SAWD = (1 + Con1_mass50) * massc_SAWD
    If massc_SAWD > 50 And massc_SAWD < 500 Then mass_SAWD = (1 + Con1_mass500) * massc_SAWD
    If massc_SAWD > 500 And massc_SAWD < 2500 Then mass_SAWD = (1 + Con1_mass2500) * massc_SAWD
    If massc_SAWD > 2500 Then mass_SAWD = (1 + Con1_mass2500_) * massc_SAWD

End If

If TRL_SAWD >= 4 And TRL_SAWD < 7 Then

    If massc_SAWD < 50 Then mass_SAWD = (1 + Con2_mass50) * massc_SAWD
    If massc_SAWD > 50 And massc_SAWD < 500 Then mass_SAWD = (1 + Con2_mass500) * massc_SAWD
    If massc_SAWD > 500 And massc_SAWD < 2500 Then mass_SAWD = (1 + Con2_mass2500) * massc_SAWD
    If massc_SAWD > 2500 Then mass_SAWD = (1 + Con2_mass2500_) * massc_SAWD

End If

If TRL_SAWD >= 7 And TRL_SAWD < 10 Then

    If massc_SAWD < 50 Then mass_SAWD = (1 + Con3_mass50) * massc_SAWD
    If massc_SAWD > 50 And massc_SAWD < 500 Then mass_SAWD = (1 + Con3_mass500) * massc_SAWD
    If massc_SAWD > 500 And massc_SAWD < 2500 Then mass_SAWD = (1 + Con3_mass2500) * massc_SAWD
    If massc_SAWD > 2500 Then mass_SAWD = (1 + Con3_mass2500_) * massc_SAWD

End If

massre_SAWD = RF_SAWD * massreg_SAWD * Crew_Size * SF_SAWD * mission_duration
vol_SAWD = RF_SAWD * volg_SAWD * Crew_Size * SF_SAWD
cool_SAWD = RF_SAWD * coolg_SAWD * Crew_Size * SF_SAWD
powc_SAWD = RF_SAWD * powg_SAWD * Crew_Size * SF_SAWD

If TRL_SAWD > 1 And TRL_SAWD < 4 Then

    If powc_SAWD < 50 Then pow_SAWD = (1 + Con1_power50) * powc_SAWD
    If powc_SAWD > 50 And powc_SAWD < 500 Then pow_SAWD = (1 + Con1_power500) * powc_SAWD
    If powc_SAWD > 500 And powc_SAWD < 2500 Then pow_SAWD = (1 + Con1_power2500) * powc_SAWD
    If powc_SAWD > 2500 Then pow_SAWD = (1 + Con1_power2500_) * powc_SAWD

End If

If TRL_SAWD >= 4 And TRL_SAWD < 7 Then

    If powc_SAWD < 50 Then pow_SAWD = (1 + Con2_power50) * powc_SAWD
    If powc_SAWD > 50 And powc_SAWD < 500 Then pow_SAWD = (1 + Con2_power500) * powc_SAWD
    If powc_SAWD > 500 And powc_SAWD < 2500 Then pow_SAWD = (1 + Con2_power2500) * powc_SAWD
    If powc_SAWD > 2500 Then pow_SAWD = (1 + Con2_power2500_) * powc_SAWD

End If

If TRL_SAWD >= 7 And TRL_SAWD < 10 Then

    If powc_SAWD < 50 Then pow_SAWD = (1 + Con3_power50) * powc_SAWD
    If powc_SAWD > 50 And powc_SAWD < 500 Then pow_SAWD = (1 + Con3_power500) * powc_SAWD
    If powc_SAWD > 500 And powc_SAWD < 2500 Then pow_SAWD = (1 + Con3_power2500) * powc_SAWD
    If powc_SAWD > 2500 Then pow_SAWD = (1 + Con3_power2500_) * powc_SAWD

End If

'__________________________________________________________________________________________________________

SPWE:

'Data given for output of 5.3 kg/d O2, requires H2O input, generates H2

Dim SF_SPWE 'Scaling Faktor
Dim RF_SPWE 'Rescale factor for SPWE
Dim massg_SPWE 'given system mass for SPWE
Dim massreg_SPWE 'given resupply mass for SPWE
Dim volg_SPWE 'given volume for SPWE
Dim powg_SPWE 'given power for SPWE
Dim coolg_SPWE ' given cooling for SPWE
Dim H2O_in_SPWE 'needed H2O input
Dim H2_out_SPWE 'generated H2
Dim O2_out_SPWE 'generated O2 output
Dim eff_SPWE 'unknown yet
Dim TRL_SPWE

Dim mass_SPWE
Dim massre_SPWE
Dim vol_SPWE
Dim pow_SPWE
Dim cool_SPWE
Dim massc_SPWE
Dim powc_SPWE

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_SPWE = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_SPWE = 0.9
If Crew_Size >= 10 Then SF_SPWE = 0.8

eff_SPWE = 1
massg_SPWE = 416
massreg_SPWE = 0.114
volg_SPWE = 1.7
powg_SPWE = 1035
coolg_SPWE = 26.5
O2_out_SPWE = 5.3
TRL_SPWE = 9

RF_SPWE = O2_Con * (1 + Margin) / (O2_out_SPWE * eff_SPWE)
massc_SPWE = RF_SPWE * massg_SPWE * Crew_Size * SF_SPWE
H2O_in_SPWE = 2 * MM_H2O / MM_O2 * Crew_Size * O2_Con
H2_out_SPWE = 4 * MM_H2 / MM_O2 * Crew_Size * O2_Con * eff_SPWE

If TRL_SPWE > 1 And TRL_SPWE < 4 Then

    If massc_SPWE < 50 Then mass_SPWE = (1 + Con1_mass50) * massc_SPWE
    If massc_SPWE > 50 And massc_SPWE < 500 Then mass_SPWE = (1 + Con1_mass500) * massc_SPWE
    If massc_SPWE > 500 And massc_SPWE < 2500 Then mass_SPWE = (1 + Con1_mass2500) * massc_SPWE
    If massc_SPWE > 2500 Then mass_SPWE = (1 + Con1_mass2500_) * massc_SPWE

End If

If TRL_SPWE >= 4 And TRL_SPWE < 7 Then

    If massc_SPWE < 50 Then mass_SPWE = (1 + Con2_mass50) * massc_SPWE
    If massc_SPWE > 50 And massc_SPWE < 500 Then mass_SPWE = (1 + Con2_mass500) * massc_SPWE
    If massc_SPWE > 500 And massc_SPWE < 2500 Then mass_SPWE = (1 + Con2_mass2500) * massc_SPWE
    If massc_SPWE > 2500 Then mass_SPWE = (1 + Con2_mass2500_) * massc_SPWE

End If

If TRL_SPWE >= 7 And TRL_SPWE < 10 Then

    If massc_SPWE < 50 Then mass_SPWE = (1 + Con3_mass50) * massc_SPWE
    If massc_SPWE > 50 And massc_SPWE < 500 Then mass_SPWE = (1 + Con3_mass500) * massc_SPWE
    If massc_SPWE > 500 And massc_SPWE < 2500 Then mass_SPWE = (1 + Con3_mass2500) * massc_SPWE
    If massc_SPWE > 2500 Then mass_SPWE = (1 + Con3_mass2500_) * massc_SPWE

End If

massre_SPWE = RF_SPWE * massreg_SPWE * Crew_Size * SF_SPWE * mission_duration
vol_SPWE = RF_SPWE * volg_SPWE * Crew_Size * SF_SPWE
cool_SPWE = RF_SPWE * coolg_SPWE * Crew_Size * SF_SPWE
powc_SPWE = RF_SPWE * powg_SPWE * Crew_Size * SF_SPWE

If TRL_SPWE > 1 And TRL_SPWE < 4 Then

    If powc_SPWE < 50 Then pow_SPWE = (1 + Con1_power50) * powc_SPWE
    If powc_SPWE > 50 And powc_SPWE < 500 Then pow_SPWE = (1 + Con1_power500) * powc_SPWE
    If powc_SPWE > 500 And powc_SPWE < 2500 Then pow_SPWE = (1 + Con1_power2500) * powc_SPWE
    If powc_SPWE > 2500 Then pow_SPWE = (1 + Con1_power2500_) * powc_SPWE

End If

If TRL_SPWE >= 4 And TRL_SPWE < 7 Then

    If powc_SPWE < 50 Then pow_SPWE = (1 + Con2_power50) * powc_SPWE
    If powc_SPWE > 50 And powc_SPWE < 500 Then pow_SPWE = (1 + Con2_power500) * powc_SPWE
    If powc_SPWE > 500 And powc_SPWE < 2500 Then pow_SPWE = (1 + Con2_power2500) * powc_SPWE
    If powc_SPWE > 2500 Then pow_SPWE = (1 + Con2_power2500_) * powc_SPWE

End If

If TRL_SPWE >= 7 And TRL_SPWE < 10 Then

    If powc_SPWE < 50 Then pow_SPWE = (1 + Con3_power50) * powc_SPWE
    If powc_SPWE > 50 And powc_SPWE < 500 Then pow_SPWE = (1 + Con3_power500) * powc_SPWE
    If powc_SPWE > 500 And powc_SPWE < 2500 Then pow_SPWE = (1 + Con3_power2500) * powc_SPWE
    If powc_SPWE > 2500 Then pow_SPWE = (1 + Con3_power2500_) * powc_SPWE

End If

'____________________________________________________________________________________________________________________


SFWE:

'Data given for 4 CM -> output of 4.42 kg/d O2 (4*O2_Con*(1+Margin)), requires H2O input, generates H2

Dim SF_SFWE 'Scaling Faktor
Dim RF_SFWE 'Rescale factor for SFWE
Dim massg_SFWE 'given system mass for SFWE
Dim massreg_SFWE 'given resupply mass for SFWE
Dim volg_SFWE 'given volume for SFWE
Dim powg_SFWE 'given power for SFWE
Dim coolg_SFWE ' given cooling for SFWE
Dim H2O_in_SFWE 'needed H2O input
Dim H2_out_SFWE 'generated H2
Dim O2_out_SFWE 'generated O2 output
Dim TRL_SFWE
Dim eff_SFWE 'unknown

Dim mass_SFWE
Dim massre_SFWE
Dim vol_SFWE
Dim pow_SFWE
Dim cool_SFWE
Dim massc_SFWE
Dim powc_SFWE

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_SFWE = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_SFWE = 0.9
If Crew_Size >= 10 Then SF_SFWE = 0.8

eff_SFWE = 1
massg_SFWE = 113
massreg_SFWE = 0 'Unknown
volg_SFWE = 0.14
powg_SFWE = 1470
coolg_SFWE = 39
O2_out_SFWE = 4 * (1 + Margin) * O2_Con
TRL_SFWE = 8
eff_SFWE = 1 'Unknown yet

RF_SFWE = O2_Con * (1 + Margin) / O2_out_SFWE
massc_SFWE = RF_SFWE * massg_SFWE * Crew_Size * SF_SFWE
H2O_in_SFWE = 2 * MM_H2O / MM_O2 * Crew_Size * O2_Con * (1 / eff_SFWE)
H2_out_SFWE = 4 * MM_H2 / MM_O2 * Crew_Size * O2_Con * eff_SFWE

If TRL_SFWE > 1 And TRL_SFWE < 4 Then

    If massc_SFWE < 50 Then mass_SFWE = (1 + Con1_mass50) * massc_SFWE
    If massc_SFWE > 50 And massc_SFWE < 500 Then mass_SFWE = (1 + Con1_mass500) * massc_SFWE
    If massc_SFWE > 500 And massc_SFWE < 2500 Then mass_SFWE = (1 + Con1_mass2500) * massc_SFWE
    If massc_SFWE > 2500 Then mass_SFWE = (1 + Con1_mass2500_) * massc_SFWE

End If

If TRL_SFWE >= 4 And TRL_SFWE < 7 Then

    If massc_SFWE < 50 Then mass_SFWE = (1 + Con2_mass50) * massc_SFWE
    If massc_SFWE > 50 And massc_SFWE < 500 Then mass_SFWE = (1 + Con2_mass500) * massc_SFWE
    If massc_SFWE > 500 And massc_SFWE < 2500 Then mass_SFWE = (1 + Con2_mass2500) * massc_SFWE
    If massc_SFWE > 2500 Then mass_SFWE = (1 + Con2_mass2500_) * massc_SFWE

End If

If TRL_SFWE >= 7 And TRL_SFWE < 10 Then

    If massc_SFWE < 50 Then mass_SFWE = (1 + Con3_mass50) * massc_SFWE
    If massc_SFWE > 50 And massc_SFWE < 500 Then mass_SFWE = (1 + Con3_mass500) * massc_SFWE
    If massc_SFWE > 500 And massc_SFWE < 2500 Then mass_SFWE = (1 + Con3_mass2500) * massc_SFWE
    If massc_SFWE > 2500 Then mass_SFWE = (1 + Con3_mass2500_) * massc_SFWE

End If

massre_SFWE = RF_SFWE * massreg_SFWE * Crew_Size * SF_SFWE * mission_duration
vol_SFWE = RF_SFWE * volg_SFWE * Crew_Size * SF_SFWE
cool_SFWE = RF_SFWE * coolg_SFWE * Crew_Size * SF_SFWE
powc_SFWE = RF_SFWE * powg_SFWE * Crew_Size * SF_SFWE

If TRL_SFWE > 1 And TRL_SFWE < 4 Then

    If powc_SFWE < 50 Then pow_SFWE = (1 + Con1_power50) * powc_SFWE
    If powc_SFWE > 50 And powc_SFWE < 500 Then pow_SFWE = (1 + Con1_power500) * powc_SFWE
    If powc_SFWE > 500 And powc_SFWE < 2500 Then pow_SFWE = (1 + Con1_power2500) * powc_SFWE
    If powc_SFWE > 2500 Then pow_SFWE = (1 + Con1_power2500_) * powc_SFWE

End If

If TRL_SFWE >= 4 And TRL_SFWE < 7 Then

    If powc_SFWE < 50 Then pow_SFWE = (1 + Con2_power50) * powc_SFWE
    If powc_SFWE > 50 And powc_SFWE < 500 Then pow_SFWE = (1 + Con2_power500) * powc_SFWE
    If powc_SFWE > 500 And powc_SFWE < 2500 Then pow_SFWE = (1 + Con2_power2500) * powc_SFWE
    If powc_SFWE > 2500 Then pow_SFWE = (1 + Con2_power2500_) * powc_SFWE

End If

If TRL_SFWE >= 7 And TRL_SFWE < 10 Then

    If powc_SFWE < 50 Then pow_SFWE = (1 + Con3_power50) * powc_SFWE
    If powc_SFWE > 50 And powc_SFWE < 500 Then pow_SFWE = (1 + Con3_power500) * powc_SFWE
    If powc_SFWE > 500 And powc_SFWE < 2500 Then pow_SFWE = (1 + Con3_power2500) * powc_SFWE
    If powc_SFWE > 2500 Then pow_SFWE = (1 + Con3_power2500_) * powc_SFWE

End If

'_____________________________________________________________________________________________________________

TIMES:

'Data given for output of 22 kg/d WW, input urine, output WW

Dim SF_TIMES 'Scaling Faktor
Dim RF_TIMES 'Rescale factor for TIMES
Dim massg_TIMES 'given system mass for TIMES
Dim massreg_TIMES 'given resupply mass for TIMES
Dim volg_TIMES 'given volume for TIMES
Dim powg_TIMES 'given power for TIMES
Dim coolg_TIMES ' given cooling for TIMES
Dim WW_out_TIMES 'given WW output
Dim eff_TIMES
Dim TRL_TIMES

Dim mass_TIMES
Dim massre_TIMES
Dim vol_TIMES
Dim pow_TIMES
Dim cool_TIMES
Dim massc_TIMES
Dim powc_TIMES

'Assignment
'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_TIMES = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_TIMES = 0.9
If Crew_Size >= 10 Then SF_TIMES = 0.8

eff_TIMES = 0.91
massg_TIMES = 68
massreg_TIMES = 0.05
volg_TIMES = 0.23
powg_TIMES = 170
coolg_TIMES = 170
WW_out_TIMES = 22
TRL_TIMES = 6

RF_TIMES = urine * (1 + Margin) / (WW_out_TIMES * eff_TIMES)
massc_TIMES = RF_TIMES * massg_TIMES * Crew_Size * SF_TIMES

If TRL_TIMES > 1 And TRL_TIMES < 4 Then

    If massc_TIMES < 50 Then mass_TIMES = (1 + Con1_mass50) * massc_TIMES
    If massc_TIMES > 50 And massc_TIMES < 500 Then mass_TIMES = (1 + Con1_mass500) * massc_TIMES
    If massc_TIMES > 500 And massc_TIMES < 2500 Then mass_TIMES = (1 + Con1_mass2500) * massc_TIMES
    If massc_TIMES > 2500 Then mass_TIMES = (1 + Con1_mass2500_) * massc_TIMES

End If

If TRL_TIMES >= 4 And TRL_TIMES < 7 Then

    If massc_TIMES < 50 Then mass_TIMES = (1 + Con2_mass50) * massc_TIMES
    If massc_TIMES > 50 And massc_TIMES < 500 Then mass_TIMES = (1 + Con2_mass500) * massc_TIMES
    If massc_TIMES > 500 And massc_TIMES < 2500 Then mass_TIMES = (1 + Con2_mass2500) * massc_TIMES
    If massc_TIMES > 2500 Then mass_TIMES = (1 + Con2_mass2500_) * massc_TIMES

End If

If TRL_TIMES >= 7 And TRL_TIMES < 10 Then

    If massc_TIMES < 50 Then mass_TIMES = (1 + Con3_mass50) * massc_TIMES
    If massc_TIMES > 50 And massc_TIMES < 500 Then mass_TIMES = (1 + Con3_mass500) * massc_TIMES
    If massc_TIMES > 500 And massc_TIMES < 2500 Then mass_TIMES = (1 + Con3_mass2500) * massc_TIMES
    If massc_TIMES > 2500 Then mass_TIMES = (1 + Con3_mass2500_) * massc_TIMES

End If

massre_TIMES = RF_TIMES * massreg_TIMES * Crew_Size * SF_TIMES * mission_duration
vol_TIMES = RF_TIMES * volg_TIMES * Crew_Size * SF_TIMES
cool_TIMES = RF_TIMES * coolg_TIMES * Crew_Size * SF_TIMES
powc_TIMES = RF_TIMES * powg_TIMES * Crew_Size * SF_TIMES

If TRL_TIMES > 1 And TRL_TIMES < 4 Then

    If powc_TIMES < 50 Then pow_TIMES = (1 + Con1_power50) * powc_TIMES
    If powc_TIMES > 50 And powc_TIMES < 500 Then pow_TIMES = (1 + Con1_power500) * powc_TIMES
    If powc_TIMES > 500 And powc_TIMES < 2500 Then pow_TIMES = (1 + Con1_power2500) * powc_TIMES
    If powc_TIMES > 2500 Then pow_TIMES = (1 + Con1_power2500_) * powc_TIMES

End If

If TRL_TIMES >= 4 And TRL_TIMES < 7 Then

    If powc_TIMES < 50 Then pow_TIMES = (1 + Con2_power50) * powc_TIMES
    If powc_TIMES > 50 And powc_TIMES < 500 Then pow_TIMES = (1 + Con2_power500) * powc_TIMES
    If powc_TIMES > 500 And powc_TIMES < 2500 Then pow_TIMES = (1 + Con2_power2500) * powc_TIMES
    If powc_TIMES > 2500 Then pow_TIMES = (1 + Con2_power2500_) * powc_TIMES

End If

If TRL_TIMES >= 7 And TRL_TIMES < 10 Then

    If powc_TIMES < 50 Then pow_TIMES = (1 + Con3_power50) * powc_TIMES
    If powc_TIMES > 50 And powc_TIMES < 500 Then pow_TIMES = (1 + Con3_power500) * powc_TIMES
    If powc_TIMES > 500 And powc_TIMES < 2500 Then pow_TIMES = (1 + Con3_power2500) * powc_TIMES
    If powc_TIMES > 2500 Then pow_TIMES = (1 + Con3_power2500_) * powc_TIMES

End If

'__________________________________________________________________________________________________________________


VCD:

'Data given for output of 9.09 kg/d WW, input urine, output WW

Dim SF_VCD 'Scaling Faktor
Dim RF_VCD 'Rescale factor for VCD
Dim massg_VCD 'given system mass for VCD
Dim massreg_VCD 'given resupply mass for VCD
Dim volg_VCD 'given volume for VCD
Dim powg_VCD 'given power for VCD
Dim coolg_VCD ' given cooling for VCD
Dim WW_out_VCD 'given WW output
Dim eff_VCD
Dim TRL_VCD

Dim mass_VCD
Dim massre_VCD
Dim vol_VCD
Dim pow_VCD
Dim cool_VCD
Dim massc_VCD
Dim powc_VCD

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_VCD = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_VCD = 0.9
If Crew_Size >= 10 Then SF_VCD = 0.8

eff_VCD = 0.7
massg_VCD = 245.7
massreg_VCD = 0.828
volg_VCD = 0.88
powg_VCD = 1429
coolg_VCD = 429
WW_out_VCD = 9.09
TRL_VCD = 9

RF_VCD = urine * (1 + Margin) / (WW_out_VCD * eff_VCD)
massc_VCD = RF_VCD * massg_VCD * Crew_Size * SF_VCD

If TRL_VCD > 1 And TRL_VCD < 4 Then

    If massc_VCD < 50 Then mass_VCD = (1 + Con1_mass50) * massc_VCD
    If massc_VCD > 50 And massc_VCD < 500 Then mass_VCD = (1 + Con1_mass500) * massc_VCD
    If massc_VCD > 500 And massc_VCD < 2500 Then mass_VCD = (1 + Con1_mass2500) * massc_VCD
    If massc_VCD > 2500 Then mass_VCD = (1 + Con1_mass2500_) * massc_VCD

End If

If TRL_VCD >= 4 And TRL_VCD < 7 Then

    If massc_VCD < 50 Then mass_VCD = (1 + Con2_mass50) * massc_VCD
    If massc_VCD > 50 And massc_VCD < 500 Then mass_VCD = (1 + Con2_mass500) * massc_VCD
    If massc_VCD > 500 And massc_VCD < 2500 Then mass_VCD = (1 + Con2_mass2500) * massc_VCD
    If massc_VCD > 2500 Then mass_VCD = (1 + Con2_mass2500_) * massc_VCD

End If

If TRL_VCD >= 7 And TRL_VCD < 10 Then

    If massc_VCD < 50 Then mass_VCD = (1 + Con3_mass50) * massc_VCD
    If massc_VCD > 50 And massc_VCD < 500 Then mass_VCD = (1 + Con3_mass500) * massc_VCD
    If massc_VCD > 500 And massc_VCD < 2500 Then mass_VCD = (1 + Con3_mass2500) * massc_VCD
    If massc_VCD > 2500 Then mass_VCD = (1 + Con3_mass2500_) * massc_VCD

End If

massre_VCD = RF_VCD * massreg_VCD * Crew_Size * SF_VCD * mission_duration
vol_VCD = RF_VCD * volg_VCD * Crew_Size * SF_VCD
cool_VCD = RF_VCD * coolg_VCD * Crew_Size * SF_VCD
powc_VCD = RF_VCD * powg_VCD * Crew_Size * SF_VCD

If TRL_VCD > 1 And TRL_VCD < 4 Then

    If powc_VCD < 50 Then pow_VCD = (1 + Con1_power50) * powc_VCD
    If powc_VCD > 50 And powc_VCD < 500 Then pow_VCD = (1 + Con1_power500) * powc_VCD
    If powc_VCD > 500 And powc_VCD < 2500 Then pow_VCD = (1 + Con1_power2500) * powc_VCD
    If powc_VCD > 2500 Then pow_VCD = (1 + Con1_power2500_) * powc_VCD

End If

If TRL_VCD >= 4 And TRL_VCD < 7 Then

    If powc_VCD < 50 Then pow_VCD = (1 + Con2_power50) * powc_VCD
    If powc_VCD > 50 And powc_VCD < 500 Then pow_VCD = (1 + Con2_power500) * powc_VCD
    If powc_VCD > 500 And powc_VCD < 2500 Then pow_VCD = (1 + Con2_power2500) * powc_VCD
    If powc_VCD > 2500 Then pow_VCD = (1 + Con2_power2500_) * powc_VCD

End If

If TRL_VCD >= 7 And TRL_VCD < 10 Then

    If powc_VCD < 50 Then pow_VCD = (1 + Con3_power50) * powc_VCD
    If powc_VCD > 50 And powc_VCD < 500 Then pow_VCD = (1 + Con3_power500) * powc_VCD
    If powc_VCD > 500 And powc_VCD < 2500 Then pow_VCD = (1 + Con3_power2500) * powc_VCD
    If powc_VCD > 2500 Then pow_VCD = (1 + Con3_power2500_) * powc_VCD

End If

'_________________________________________________________________________________________________________

VPCAR:

'Data given for output of 50.9 kg/d WW (=hygiene+others) input urine, O2, output N2, CO2, H2O

Dim SF_VPCAR 'Scaling Faktor
Dim RF_VPCAR 'Rescale factor for VPCAR
Dim massg_VPCAR 'given system mass for VPCAR
Dim massreg_VPCAR 'given resupply mass for VPCAR
Dim volg_VPCAR 'given volume for VPCAR
Dim powg_VPCAR 'given power for VPCAR
Dim coolg_VPCAR ' given cooling for VPCAR
Dim WW_in_VPCAR 'given WW output
Dim TRL_VPCAR
Dim eff_VPCAR
Dim ww_other_VPCAR 'other waste water components, unkown yet -> set to 0

Dim mass_VPCAR
Dim massre_VPCAR
Dim vol_VPCAR
Dim pow_VPCAR
Dim cool_VPCAR
Dim massc_VPCAR
Dim powc_VPCAR

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_VPCAR = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_VPCAR = 0.9
If Crew_Size >= 10 Then SF_VPCAR = 0.8

eff_VPCAR = 0.95
massg_VPCAR = 412
massreg_VPCAR = 0.005
volg_VPCAR = 1.57
powg_VPCAR = 2380
coolg_VPCAR = 2380
WW_in_VPCAR = 50.9
ww_other_VPCAR = 0
TRL_VPCAR = 9

RF_VPCAR = (hygiene + ww_other_VPCAR + urine) * (1 + Margin) / WW_in_VPCAR
massc_VPCAR = RF_VPCAR * massg_VPCAR * Crew_Size * SF_VPCAR


If TRL_VPCAR > 1 And TRL_VPCAR < 4 Then

    If massc_VPCAR < 50 Then mass_VPCAR = (1 + Con1_mass50) * massc_VPCAR
    If massc_VPCAR > 50 And massc_VPCAR < 500 Then mass_VPCAR = (1 + Con1_mass500) * massc_VPCAR
    If massc_VPCAR > 500 And massc_VPCAR < 2500 Then mass_VPCAR = (1 + Con1_mass2500) * massc_VPCAR
    If massc_VPCAR > 2500 Then mass_VPCAR = (1 + Con1_mass2500_) * massc_VPCAR

End If

If TRL_VPCAR >= 4 And TRL_VPCAR < 7 Then

    If massc_VPCAR < 50 Then mass_VPCAR = (1 + Con2_mass50) * massc_VPCAR
    If massc_VPCAR > 50 And massc_VPCAR < 500 Then mass_VPCAR = (1 + Con2_mass500) * massc_VPCAR
    If massc_VPCAR > 500 And massc_VPCAR < 2500 Then mass_VPCAR = (1 + Con2_mass2500) * massc_VPCAR
    If massc_VPCAR > 2500 Then mass_VPCAR = (1 + Con2_mass2500_) * massc_VPCAR

End If

If TRL_VPCAR >= 7 And TRL_VPCAR < 10 Then

    If massc_VPCAR < 50 Then mass_VPCAR = (1 + Con3_mass50) * massc_VPCAR
    If massc_VPCAR > 50 And massc_VPCAR < 500 Then mass_VPCAR = (1 + Con3_mass500) * massc_VPCAR
    If massc_VPCAR > 500 And massc_VPCAR < 2500 Then mass_VPCAR = (1 + Con3_mass2500) * massc_VPCAR
    If massc_VPCAR > 2500 Then mass_VPCAR = (1 + Con3_mass2500_) * massc_VPCAR

End If

massre_VPCAR = RF_VPCAR * massreg_VPCAR * Crew_Size * SF_VPCAR * mission_duration
vol_VPCAR = RF_VPCAR * volg_VPCAR * Crew_Size * SF_VPCAR
cool_VPCAR = RF_VPCAR * coolg_VPCAR * Crew_Size * SF_VPCAR
powc_VPCAR = RF_VPCAR * powg_VPCAR * Crew_Size * SF_VPCAR

If TRL_VPCAR > 1 And TRL_VPCAR < 4 Then

    If powc_VPCAR < 50 Then pow_VPCAR = (1 + Con1_power50) * powc_VPCAR
    If powc_VPCAR > 50 And powc_VPCAR < 500 Then pow_VPCAR = (1 + Con1_power500) * powc_VPCAR
    If powc_VPCAR > 500 And powc_VPCAR < 2500 Then pow_VPCAR = (1 + Con1_power2500) * powc_VPCAR
    If powc_VPCAR > 2500 Then pow_VPCAR = (1 + Con1_power2500_) * powc_VPCAR

End If

If TRL_VPCAR >= 4 And TRL_VPCAR < 7 Then

    If powc_VPCAR < 50 Then pow_VPCAR = (1 + Con2_power50) * powc_VPCAR
    If powc_VPCAR > 50 And powc_VPCAR < 500 Then pow_VPCAR = (1 + Con2_power500) * powc_VPCAR
    If powc_VPCAR > 500 And powc_VPCAR < 2500 Then pow_VPCAR = (1 + Con2_power2500) * powc_VPCAR
    If powc_VPCAR > 2500 Then pow_VPCAR = (1 + Con2_power2500_) * powc_VPCAR

End If

If TRL_VPCAR >= 7 And TRL_VPCAR < 10 Then

    If powc_VPCAR < 50 Then pow_VPCAR = (1 + Con3_power50) * powc_VPCAR
    If powc_VPCAR > 50 And powc_VPCAR < 500 Then pow_VPCAR = (1 + Con3_power500) * powc_VPCAR
    If powc_VPCAR > 500 And powc_VPCAR < 2500 Then pow_VPCAR = (1 + Con3_power2500) * powc_VPCAR
    If powc_VPCAR > 2500 Then pow_VPCAR = (1 + Con3_power2500_) * powc_VPCAR

End If

'__________________________________________________________________________________________________________

ACRS:

'Data given for Crew of 3 -> assumption: CO2_in=3*CO2_Pro
'Overall reaction: CO2+2H2->C+2H2O

Dim SF_ACRS 'Scaling Faktor
Dim RF_ACRS 'Rescale factor for ACRS
Dim massg_ACRS 'given system mass for ACRS
Dim massreg_ACRS 'given resupply mass for ACRS
Dim volg_ACRS 'given volume for ACRS
Dim powg_ACRS 'given power for ACRS
Dim coolg_ACRS ' given cooling for ACRS
Dim CO2_in_ACRS 'given CO2 input
Dim H2_in_ACRS 'needed H2 input
Dim H2O_out_ACRS 'generated waste water output
Dim CH4_out_ACRS 'generated methane output
Dim eff_ACRS
Dim TRL_ACRS

Dim mass_ACRS
Dim massre_ACRS
Dim vol_ACRS
Dim pow_ACRS
Dim cool_ACRS
Dim massc_ACRS
Dim powc_ACRS

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_ACRS = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_ACRS = 0.9
If Crew_Size >= 10 Then SF_ACRS = 0.8

eff_ACRS = 0.9
massg_ACRS = 180
massreg_ACRS = 0 'Unknown yet
volg_ACRS = 0.3
powg_ACRS = 400
coolg_ACRS = 150
CO2_in_ACRS = 3 * CO2_Pro
TRL_ACRS = 3

RF_ACRS = CO2_Pro * (1 + Margin) / CO2_in_ACRS
massc_ACRS = RF_ACRS * massg_ACRS * Crew_Size * SF_ACRS
H2_in_ACRS = 2 * MM_H2 / MM_CO2 * Crew_Size * CO2_Pro
H2O_out_ACRS = 2 * MM_H2O / MM_CO2 * Crew_Size * CO2_Pro * eff_ACRS


If TRL_ACRS > 1 And TRL_ACRS < 4 Then

    If massc_ACRS < 50 Then mass_ACRS = (1 + Con1_mass50) * massc_ACRS
    If massc_ACRS > 50 And massc_ACRS < 500 Then mass_ACRS = (1 + Con1_mass500) * massc_ACRS
    If massc_ACRS > 500 And massc_ACRS < 2500 Then mass_ACRS = (1 + Con1_mass2500) * massc_ACRS
    If massc_ACRS > 2500 Then mass_ACRS = (1 + Con1_mass2500_) * massc_ACRS

End If

If TRL_ACRS >= 4 And TRL_ACRS < 7 Then

    If massc_ACRS < 50 Then mass_ACRS = (1 + Con2_mass50) * massc_ACRS
    If massc_ACRS > 50 And massc_ACRS < 500 Then mass_ACRS = (1 + Con2_mass500) * massc_ACRS
    If massc_ACRS > 500 And massc_ACRS < 2500 Then mass_ACRS = (1 + Con2_mass2500) * massc_ACRS
    If massc_ACRS > 2500 Then mass_ACRS = (1 + Con2_mass2500_) * massc_ACRS

End If

If TRL_ACRS >= 7 And TRL_ACRS < 10 Then

    If massc_ACRS < 50 Then mass_ACRS = (1 + Con3_mass50) * massc_ACRS
    If massc_ACRS > 50 And massc_ACRS < 500 Then mass_ACRS = (1 + Con3_mass500) * massc_ACRS
    If massc_ACRS > 500 And massc_ACRS < 2500 Then mass_ACRS = (1 + Con3_mass2500) * massc_ACRS
    If massc_ACRS > 2500 Then mass_ACRS = (1 + Con3_mass2500_) * massc_ACRS

End If

massre_ACRS = RF_ACRS * massreg_ACRS * Crew_Size * SF_ACRS * mission_duration
vol_ACRS = RF_ACRS * volg_ACRS * Crew_Size * SF_ACRS
cool_ACRS = RF_ACRS * coolg_ACRS * Crew_Size * SF_ACRS
powc_ACRS = RF_ACRS * powg_ACRS * Crew_Size * SF_ACRS

If TRL_ACRS > 1 And TRL_ACRS < 4 Then

    If powc_ACRS < 50 Then pow_ACRS = (1 + Con1_power50) * powc_ACRS
    If powc_ACRS > 50 And powc_ACRS < 500 Then pow_ACRS = (1 + Con1_power500) * powc_ACRS
    If powc_ACRS > 500 And powc_ACRS < 2500 Then pow_ACRS = (1 + Con1_power2500) * powc_ACRS
    If powc_ACRS > 2500 Then pow_ACRS = (1 + Con1_power2500_) * powc_ACRS

End If

If TRL_ACRS >= 4 And TRL_ACRS < 7 Then

    If powc_ACRS < 50 Then pow_ACRS = (1 + Con2_power50) * powc_ACRS
    If powc_ACRS > 50 And powc_ACRS < 500 Then pow_ACRS = (1 + Con2_power500) * powc_ACRS
    If powc_ACRS > 500 And powc_ACRS < 2500 Then pow_ACRS = (1 + Con2_power2500) * powc_ACRS
    If powc_ACRS > 2500 Then pow_ACRS = (1 + Con2_power2500_) * powc_ACRS

End If

If TRL_ACRS >= 7 And TRL_ACRS < 10 Then

    If powc_ACRS < 50 Then pow_ACRS = (1 + Con3_power50) * powc_ACRS
    If powc_ACRS > 50 And powc_ACRS < 500 Then pow_ACRS = (1 + Con3_power500) * powc_ACRS
    If powc_ACRS > 500 And powc_ACRS < 2500 Then pow_ACRS = (1 + Con3_power2500) * powc_ACRS
    If powc_ACRS > 2500 Then pow_ACRS = (1 + Con3_power2500_) * powc_ACRS

End If

'_______________________________________________________________________________________________________

LiOH:

'Values taken from SS

Dim SF_LiOH 'Scaling Faktor
Dim RF_LiOH 'Rescale factor for LiOH
Dim massg_hard_LiOH 'given hardware mass for LiOH
Dim massg_cart_LiOH 'given cartridge mass for LiOH
Dim massg_rack_LiOH 'given rack mass for LiOH
Dim massreg_LiOH 'given resupply mass for LiOH
Dim volg_rack_LiOH 'given rack volume for LiOH
Dim powg_LiOH 'given power for LiOH
Dim coolg_LiOH 'given cooling for LiOH
Dim LiOH_needed 'Needed amount of LiOH per kg CO2
Dim LiOH_needed_total 'Total amount of LiOH needed for mission
Dim LiOH_cart_needed 'Total amount of needed cartridges
Dim LiOH_racks_needed 'Total amount of racks needed
Dim H2O_Pro_LiOH 'H2O Produced per day
Dim Li2CO3_Pro 'Produced Li2CO3 by reaction
Dim TRL_LiOH

Dim mass_LiOH
Dim massre_LiOH
Dim vol_LiOH
Dim pow_LiOH
Dim cool_LiOH
Dim massc_LiOH
Dim powc_LiOH

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_LiOH = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_LiOH = 0.9
If Crew_Size >= 10 Then SF_LiOH = 0.8

massg_hard_LiOH = 10.43
massg_cart_LiOH = 0.907
massg_rack_LiOH = 3.63
massreg_LiOH = 0 'all needed cartridges for mission are taken along
volg_rack_LiOH = 0.0311
powg_LiOH = 12
coolg_LiOH = 0
TRL_LiOH = 9

LiOH_needed = 2 * MM_LiOH / MM_CO2
LiOH_needed_total = LiOH_needed * CO2_Pro * Crew_Size * mission_duration
LiOH_cart_needed = LiOH_needed_total / 2.27 '2.27 kg per cartridge
LiOH_racks_needed = LiOH_cart_needed / 27 '27 cartridges per rack

Li2CO3_Pro = MM_Li2CO3 / MM_CO2 * CO2_Pro
H2O_Pro_LiOH = MM_H2O / MM_CO2 * CO2_Pro * Crew_Size

massc_LiOH = LiOH_needed_total + Application.RoundUp(LiOH_cart_needed, 0) * massg_cart_LiOH + Application.RoundUp(LiOH_racks_needed, 0) * massg_rack_LiOH + _
massg_hard_LiOH * SF_LiOH * Crew_Size


If TRL_LiOH > 1 And TRL_LiOH < 4 Then

    If massc_LiOH < 50 Then mass_LiOH = (1 + Con1_mass50) * massc_LiOH
    If massc_LiOH > 50 And massc_LiOH < 500 Then mass_LiOH = (1 + Con1_mass500) * massc_LiOH
    If massc_LiOH > 500 And massc_LiOH < 2500 Then mass_LiOH = (1 + Con1_mass2500) * massc_LiOH
    If massc_LiOH > 2500 Then mass_LiOH = (1 + Con1_mass2500_) * massc_LiOH

End If

If TRL_LiOH >= 4 And TRL_LiOH < 7 Then

    If massc_LiOH < 50 Then mass_LiOH = (1 + Con2_mass50) * massc_LiOH
    If massc_LiOH > 50 And massc_LiOH < 500 Then mass_LiOH = (1 + Con2_mass500) * massc_LiOH
    If massc_LiOH > 500 And massc_LiOH < 2500 Then mass_LiOH = (1 + Con2_mass2500) * massc_LiOH
    If massc_LiOH > 2500 Then mass_LiOH = (1 + Con2_mass2500_) * massc_LiOH

End If

If TRL_LiOH >= 7 And TRL_LiOH < 10 Then

    If massc_LiOH < 50 Then mass_LiOH = (1 + Con3_mass50) * massc_LiOH
    If massc_LiOH > 50 And massc_LiOH < 500 Then mass_LiOH = (1 + Con3_mass500) * massc_LiOH
    If massc_LiOH > 500 And massc_LiOH < 2500 Then mass_LiOH = (1 + Con3_mass2500) * massc_LiOH
    If massc_LiOH > 2500 Then mass_LiOH = (1 + Con3_mass2500_) * massc_LiOH

End If

mass_LiOH = mass_LiOH - LiOH_needed_total
massre_LiOH = LiOH_needed_total 'massreg_LiOH * Crew_size * SF_LiOH * Mission_duration
vol_LiOH = volg_rack_LiOH * Application.RoundUp(LiOH_racks_needed, 0)
cool_LiOH = 0
powc_LiOH = powg_LiOH * Crew_Size * SF_LiOH

If TRL_LiOH > 1 And TRL_LiOH < 4 Then

    If powc_LiOH < 50 Then pow_LiOH = (1 + Con1_power50) * powc_LiOH
    If powc_LiOH > 50 And powc_LiOH < 500 Then pow_LiOH = (1 + Con1_power500) * powc_LiOH
    If powc_LiOH > 500 And powc_LiOH < 2500 Then pow_LiOH = (1 + Con1_power2500) * powc_LiOH
    If powc_LiOH > 2500 Then pow_LiOH = (1 + Con1_power2500_) * powc_LiOH

End If

If TRL_LiOH >= 4 And TRL_LiOH < 7 Then

    If powc_LiOH < 50 Then pow_LiOH = (1 + Con2_power50) * powc_LiOH
    If powc_LiOH > 50 And powc_LiOH < 500 Then pow_LiOH = (1 + Con2_power500) * powc_LiOH
    If powc_LiOH > 500 And powc_LiOH < 2500 Then pow_LiOH = (1 + Con2_power2500) * powc_LiOH
    If powc_LiOH > 2500 Then pow_LiOH = (1 + Con2_power2500_) * powc_LiOH

End If

If TRL_LiOH >= 7 And TRL_LiOH < 10 Then

    If powc_LiOH < 50 Then pow_LiOH = (1 + Con3_power50) * powc_LiOH
    If powc_LiOH > 50 And powc_LiOH < 500 Then pow_LiOH = (1 + Con3_power500) * powc_LiOH
    If powc_LiOH > 500 And powc_LiOH < 2500 Then pow_LiOH = (1 + Con3_power2500) * powc_LiOH
    If powc_LiOH > 2500 Then pow_LiOH = (1 + Con3_power2500_) * powc_LiOH

End If

'______________________________________________________________________________________________________________

HPS:

'Involves losses due to leakage

Dim O2_vol_HPS 'needed O2 volume
Dim N2_vol_HPS 'needed N2 volume
Dim O2_mass 'needed O2 mass
Dim N2_mass 'needed N2 mass
Dim p_tank_HPS 'tank pressure (identical for O2 and N2)
Dim T_tank_HPS 'tank temperature (ambient, identical for O2 and N2)
Dim tank_mass 'needed tank mass
Dim tank_vol 'needed tank volume
Dim tank_minc_O2_HPS 'tank mass increase for O2 storage [kg/kg]
Dim tank_minc_N2_HPS 'tank mass increase for N2 storage [kg/kg]
Dim tfd 'tank filling degree
Dim t_wall_HPS 'tank wall thickness
Dim TRL_HPS

p_tank_HPS = 41.37 * 10 ^ 6
T_tank_HPS = 293.15
tank_minc_O2_HPS = 0.364
tank_minc_N2_HPS = (1.7 - 0.556) / 2 'use mean value
tfd = 0.95
t_wall_HPS = 0.01
TRL_HPS = 9

If DES_Function = "Store Oxygen" Then O2_mass = O2_Con * Crew_Size * mission_duration * (1 + Margin)
If DES_Function = "Recover from Rapid Decompression" Then O2_mass = pp_O2 * Pressurized_volume / (R_O2 * 293.15) + O2_leakage * mission_duration

O2_vol_HPS = (O2_mass * R_O2 * T_tank_HPS) / p_tank_HPS

If DES_Function = "Store Inert Gas" Then N2_mass = N2_leakage * mission_duration * (1 + Margin) + pp_N2 * Pressurized_volume / (R_N2 * 293.15)
If DES_Function = "Recover from Rapid Decompression" Then N2_mass = pp_N2 * Pressurized_volume / (R_N2 * 293.15)

N2_vol_HPS = (N2_mass * R_N2 * T_tank_HPS) / p_tank_HPS

'_________________________________________________________________________________________________


CS:

'Involves losses due to leakage

Dim O2_vol_CS 'needed O2 volume
Dim N2_vol_CS 'needed N2 volume
Dim p_tank_CS 'tank pressure (identical for O2 and N2)
Dim T_tank_CS_N2 'tank temperature N2
Dim T_tank_CS_O2 'tank temperature O2
Dim tank_minc_O2_CS 'tank mass increase for O2 storage [kg/kg]
Dim tank_minc_N2_CS 'tank mass increase for N2 storage [kg/kg]
Dim t_wall_CS 'wall thickness, twice as thick than HPS due to significantly higher insulation requirements
Dim TRL_CS

p_tank_CS = 34.5 * 10 ^ 6
T_tank_CS_N2 = 73.15
T_tank_CS_O2 = 73.15
tank_minc_O2_CS = 0.429
tank_minc_N2_CS = 0.524
t_wall_CS = 0.02
TRL_CS = 9

O2_vol_CS = (O2_mass * R_O2 * T_tank_CS_O2) / p_tank_CS
N2_vol_CS = (N2_mass * R_N2 * T_tank_CS_N2) / p_tank_CS

'_________________________________________________________________________________________________

Oxygen_Candles:

Dim m_LiClO4
Dim nO2c 'number of oxygen candles required
Dim massc_O2c 'calculation mass of O2 candles
Dim mass_O2c 'mass of O2 candles
Dim vol_O2c 'volume of O2 candles
Dim T_O2
Dim T_N2
Dim massre_O2c
Dim pow_O2c
Dim cool_O2c
Dim powc_O2c
Dim TRL_O2c

T_O2 = 293.15
T_N2 = 293.15
m_LiClO4 = MM_LiClO4 / (2 * MM_O2) * O2_Con * Crew_Size * mission_duration
nO2c = Application.RoundUp(m_LiClO4 / 2.2, 0)
massc_O2c = nO2c * 2.7
TRL_O2c = 8

If TRL_O2c > 1 And TRL_O2c < 4 Then

    If massc_O2c < 50 Then mass_O2c = (1 + Con1_mass50) * massc_O2c
    If massc_O2c > 50 And massc_O2c < 500 Then mass_O2c = (1 + Con1_mass500) * massc_O2c
    If massc_O2c > 500 And massc_O2c < 2500 Then mass_O2c = (1 + Con1_mass2500) * massc_O2c
    If massc_O2c > 2500 Then mass_O2c = (1 + Con1_mass2500_) * massc_O2c

End If

If TRL_O2c >= 4 And TRL_O2c < 7 Then

    If massc_O2c < 50 Then mass_O2c = (1 + Con2_mass50) * massc_O2c
    If massc_O2c > 50 And massc_O2c < 500 Then mass_O2c = (1 + Con2_mass500) * massc_O2c
    If massc_O2c > 500 And massc_O2c < 2500 Then mass_O2c = (1 + Con2_mass2500) * massc_O2c
    If massc_O2c > 2500 Then mass_O2c = (1 + Con2_mass2500_) * massc_O2c

End If

If TRL_O2c >= 7 And TRL_O2c < 10 Then

    If massc_O2c < 50 Then mass_O2c = (1 + Con3_mass50) * massc_O2c
    If massc_O2c > 50 And massc_O2c < 500 Then mass_O2c = (1 + Con3_mass500) * massc_O2c
    If massc_O2c > 500 And massc_O2c < 2500 Then mass_O2c = (1 + Con3_mass2500) * massc_O2c
    If massc_O2c > 2500 Then mass_O2c = (1 + Con3_mass2500_) * massc_O2c

End If

massre_O2c = 0  'MM_LiClO4 / (2 * MM_O2) * O2_Con * crew_size * mission_duration
mass_O2c = mass_O2c
vol_O2c = nO2c * 3.6 * 10 ^ (-3)
cool_O2c = 0
powc_O2c = 0

If TRL_O2c > 1 And TRL_O2c < 4 Then

    If powc_O2c < 50 Then pow_O2c = (1 + Con1_power50) * powc_O2c
    If powc_O2c > 50 And powc_O2c < 500 Then pow_O2c = (1 + Con1_power500) * powc_O2c
    If powc_O2c > 500 And powc_O2c < 2500 Then pow_O2c = (1 + Con1_power2500) * powc_O2c
    If powc_O2c > 2500 Then pow_O2c = (1 + Con1_power2500_) * powc_O2c

End If

If TRL_O2c >= 4 And TRL_O2c < 7 Then

    If powc_O2c < 50 Then pow_O2c = (1 + Con2_power50) * powc_O2c
    If powc_O2c > 50 And powc_O2c < 500 Then pow_O2c = (1 + Con2_power500) * powc_O2c
    If powc_O2c > 500 And powc_O2c < 2500 Then pow_O2c = (1 + Con2_power2500) * powc_O2c
    If powc_O2c > 2500 Then pow_O2c = (1 + Con2_power2500_) * powc_O2c

End If

If TRL_O2c >= 7 And TRL_O2c < 10 Then

    If powc_O2c < 50 Then pow_O2c = (1 + Con3_power50) * powc_O2c
    If powc_O2c > 50 And powc_O2c < 500 Then pow_O2c = (1 + Con3_power500) * powc_O2c
    If powc_O2c > 500 And powc_O2c < 2500 Then pow_O2c = (1 + Con3_power2500) * powc_O2c
    If powc_O2c > 2500 Then pow_O2c = (1 + Con3_power2500_) * powc_O2c

End If


'____________________________________________________________________________________________________________

Algalreactor:

Dim massg_alg
Dim massc_alg
Dim mass_alg
Dim volg_alg
Dim vol_alg
Dim powerg_alg
Dim powerc_alg
Dim power_alg
Dim massre_alg
Dim cool_alg
Dim TRL_alg
Dim SF_alg
Dim RF_alg

'Scaling factor
If Crew_Size >= 1 And Crew_Size < 7 Then SF_alg = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_alg = 0.9
If Crew_Size >= 10 Then SF_alg = 0.8


'data given for crew of six
RF_alg = Crew_Size / 6
TRL_alg = 5



For i = 2 To na
    If Cells(i, 1).Value = "Algal Reactor" And Cells(i, 2).Value = "air-AR" Then
    
        massg_alg = 1700
        volg_alg = 1.66
        powerg_alg = 4080
        
        massc_alg = RF_alg * SF_alg * massg_alg
        powerc_alg = RF_alg * SF_alg * powerg_alg
        vol_alg = RF_alg * SF_alg * volg_alg
        
        If TRL_alg > 1 And TRL_alg < 4 Then

            If massc_alg < 50 Then mass_alg = (1 + Con1_mass50) * massc_alg
            If massc_alg > 50 And massc_alg < 500 Then mass_alg = (1 + Con1_mass500) * massc_alg
            If massc_alg > 500 And massc_alg < 2500 Then mass_alg = (1 + Con1_mass2500) * massc_alg
            If massc_alg > 2500 Then mass_alg = (1 + Con1_mass2500_) * massc_alg
        
        End If

        If TRL_alg >= 4 And TRL_alg < 7 Then
        
            If massc_alg < 50 Then mass_alg = (1 + Con2_mass50) * massc_alg
            If massc_alg > 50 And massc_alg < 500 Then mass_alg = (1 + Con2_mass500) * massc_alg
            If massc_alg > 500 And massc_alg < 2500 Then mass_alg = (1 + Con2_mass2500) * massc_alg
            If massc_alg > 2500 Then mass_alg = (1 + Con2_mass2500_) * massc_alg
        
        End If
        
        If TRL_alg >= 7 And TRL_alg < 10 Then
        
            If massc_alg < 50 Then mass_alg = (1 + Con3_mass50) * massc_alg
            If massc_alg > 50 And massc_alg < 500 Then mass_alg = (1 + Con3_mass500) * massc_alg
            If massc_alg > 500 And massc_alg < 2500 Then mass_alg = (1 + Con3_mass2500) * massc_alg
            If massc_alg > 2500 Then mass_alg = (1 + Con3_mass2500_) * massc_alg
        
        End If
        
        If TRL_alg > 1 And TRL_alg < 4 Then

            If powerc_alg < 50 Then power_alg = (1 + Con1_power50) * powerc_alg
            If powerc_alg > 50 And powerc_alg < 500 Then power_alg = (1 + Con1_power500) * powerc_alg
            If powerc_alg > 500 And powerc_alg < 2500 Then power_alg = (1 + Con1_power2500) * powerc_alg
            If powerc_alg > 2500 Then power_alg = (1 + Con1_power2500_) * powerc_alg
        
        End If
        
        If TRL_alg >= 4 And TRL_alg < 7 Then
        
            If powerc_alg < 50 Then power_alg = (1 + Con2_power50) * powerc_alg
            If powerc_alg > 50 And powerc_alg < 500 Then power_alg = (1 + Con2_power500) * powerc_alg
            If powerc_alg > 500 And powerc_alg < 2500 Then power_alg = (1 + Con2_power2500) * powerc_alg
            If powerc_alg > 2500 Then power_alg = (1 + Con2_power2500_) * powerc_alg
        
        End If
        
        If TRL_alg >= 7 And TRL_alg < 10 Then
        
            If powerc_alg < 50 Then power_alg = (1 + Con3_power50) * powerc_alg
            If powerc_alg > 50 And powerc_alg < 500 Then power_alg = (1 + Con3_power500) * powerc_alg
            If powerc_alg > 500 And powerc_alg < 2500 Then power_alg = (1 + Con3_power2500) * powerc_alg
            If powerc_alg > 2500 Then power_alg = (1 + Con3_power2500_) * powerc_alg
        
        End If

        cool_alg = power_alg
        massre_alg = 1300 * mission_duration * 1 / 1000
    
    ElseIf Cells(i, 1).Value = "Algal Reactor" And Cells(i, 2).Value = "water-WRM" Then
    
        massg_alg = 24050
        volg_alg = 21
        powerg_alg = 85700
        
        massc_alg = RF_alg * SF_alg * massg_alg
        powerc_alg = RF_alg * SF_alg * powerg_alg
        vol_alg = RF_alg * SF_alg * volg_alg
        
        If TRL_alg > 1 And TRL_alg < 4 Then

            If massc_alg < 50 Then mass_alg = (1 + Con1_mass50) * massc_alg
            If massc_alg > 50 And massc_alg < 500 Then mass_alg = (1 + Con1_mass500) * massc_alg
            If massc_alg > 500 And massc_alg < 2500 Then mass_alg = (1 + Con1_mass2500) * massc_alg
            If massc_alg > 2500 Then mass_alg = (1 + Con1_mass2500_) * massc_alg
        
        End If

        If TRL_alg >= 4 And TRL_alg < 7 Then
        
            If massc_alg < 50 Then mass_alg = (1 + Con2_mass50) * massc_alg
            If massc_alg > 50 And massc_alg < 500 Then mass_alg = (1 + Con2_mass500) * massc_alg
            If massc_alg > 500 And massc_alg < 2500 Then mass_alg = (1 + Con2_mass2500) * massc_alg
            If massc_alg > 2500 Then mass_alg = (1 + Con2_mass2500_) * massc_alg
        
        End If
        
        If TRL_alg >= 7 And TRL_alg < 10 Then
        
            If massc_alg < 50 Then mass_alg = (1 + Con3_mass50) * massc_alg
            If massc_alg > 50 And massc_alg < 500 Then mass_alg = (1 + Con3_mass500) * massc_alg
            If massc_alg > 500 And massc_alg < 2500 Then mass_alg = (1 + Con3_mass2500) * massc_alg
            If massc_alg > 2500 Then mass_alg = (1 + Con3_mass2500_) * massc_alg
        
        End If
        
        If TRL_alg > 1 And TRL_alg < 4 Then

            If powerc_alg < 50 Then power_alg = (1 + Con1_power50) * powerc_alg
            If powerc_alg > 50 And powerc_alg < 500 Then power_alg = (1 + Con1_power500) * powerc_alg
            If powerc_alg > 500 And powerc_alg < 2500 Then power_alg = (1 + Con1_power2500) * powerc_alg
            If powerc_alg > 2500 Then power_alg = (1 + Con1_power2500_) * powerc_alg
        
        End If
        
        If TRL_alg >= 4 And TRL_alg < 7 Then
        
            If powerc_alg < 50 Then power_alg = (1 + Con2_power50) * powerc_alg
            If powerc_alg > 50 And powerc_alg < 500 Then power_alg = (1 + Con2_power500) * powerc_alg
            If powerc_alg > 500 And powerc_alg < 2500 Then power_alg = (1 + Con2_power2500) * powerc_alg
            If powerc_alg > 2500 Then power_alg = (1 + Con2_power2500_) * powerc_alg
        
        End If
        
        If TRL_alg >= 7 And TRL_alg < 10 Then
        
            If powerc_alg < 50 Then power_alg = (1 + Con3_power50) * powerc_alg
            If powerc_alg > 50 And powerc_alg < 500 Then power_alg = (1 + Con3_power500) * powerc_alg
            If powerc_alg > 500 And powerc_alg < 2500 Then power_alg = (1 + Con3_power2500) * powerc_alg
            If powerc_alg > 2500 Then power_alg = (1 + Con3_power2500_) * powerc_alg
        
        End If

        cool_alg = power_alg
        massre_alg = 0
        
    End If
    
Next i

'_____________________________________________________________________________________________________________________________________
    
        

Water_tank:

Dim H2O_vol_w_wt 'needed H2O volume
Dim H2O_vol_ww_wt 'needed waste H2O volume
Dim t_wall_wt_s 'shell wall thickness
Dim t_wall_wt_b 'bladder wall thickness
Dim rho_s
Dim rho_b
Dim d_wt 'diameter water tank
Dim TRL_wt
Dim pow_w_wt
Dim cool_w_wt
Dim massre_w_wt
Dim pow_ww_wt
Dim cool_ww_wt
Dim massre_ww_wt

t_wall_wt_s = 0.0005
t_wall_wt_b = 0.0005
TRL_wt = 9
rho_s = 2710 '[kg/m3]
rho_b = 1500 '[kg/m3]
H2O_vol_w_wt = H2O_Con / rho_H2O
H2O_vol_ww_wt = H2O_Pro / rho_H2O
pow_w_wt = 0 'unknown
cool_w_wt = 0 'unknown
massre_w_wt = H2O_Con
pow_ww_wt = 0 'unknown
cool_ww_wt = 0 'unknown
massre_ww_wt = 0


'_____________________________________________________________________________________________________________


'Schreiben der Ergebnisse

For i = 2 To na

    If Cells(i, 1).Value = "Air Evaporation System AES" Then
        
        Cells(i, "I").Value = mass_AES
        Cells(i, "L").Value = vol_AES
        Cells(i, "O").Value = cool_AES
        Cells(i, "R").Value = pow_AES
        Cells(i, "Y").Value = massre_AES
        Cells(i, "AA").Value = TRL_AES
    
    End If
    
    If Cells(i, 1).Value = "Common Cabin Air Assembly CCAA" Then
        
        Cells(i, "I").Value = mass_CCAA
        Cells(i, "L").Value = vol_CCAA
        Cells(i, "O").Value = cool_CCAA
        Cells(i, "R").Value = pow_CCAA
        Cells(i, "Y").Value = massre_CCAA
        Cells(i, "AA").Value = TRL_CCAA
    
    End If
    
    If Cells(i, 1).Value = "Bosch Reactor BR" Then
        
        Cells(i, "I").Value = mass_Bs
        Cells(i, "L").Value = vol_Bs
        Cells(i, "O").Value = cool_Bs
        Cells(i, "R").Value = pow_Bs
        Cells(i, "Y").Value = massre_Bs
        Cells(i, "AA").Value = TRL_Bs
    
    End If
    
    If Cells(i, 1).Value = "CO2 and Moisture Removal Amine Swing-Bed System CAMRAS" Then
        
        Cells(i, "I").Value = mass_CAMRAS
        Cells(i, "L").Value = vol_CAMRAS
        Cells(i, "O").Value = cool_CAMRAS
        Cells(i, "R").Value = pow_CAMRAS
        Cells(i, "Y").Value = massre_CAMRAS
        Cells(i, "AA").Value = TRL_CAMRAS
    
    End If
    
    If Cells(i, 1).Value = "CO2 electrolysis COE" Then
        
        Cells(i, "I").Value = mass_COE
        Cells(i, "L").Value = vol_COE
        Cells(i, "O").Value = cool_COE
        Cells(i, "R").Value = pow_COE
        Cells(i, "Y").Value = massre_COE
        Cells(i, "AA").Value = TRL_COE
    
    End If
    
    If Cells(i, 1).Value = "Electrochemical depolarization concentration EDC" Then
        
        Cells(i, "I").Value = mass_EDC
        Cells(i, "L").Value = vol_EDC
        Cells(i, "O").Value = cool_EDC
        Cells(i, "R").Value = pow_EDC
        Cells(i, "Y").Value = massre_EDC
        Cells(i, "AA").Value = TRL_EDC
    
    End If
    
    If Cells(i, 1).Value = "Molecular Sieves 2BMS" Then
        
        Cells(i, "I").Value = mass_2BMS
        Cells(i, "L").Value = vol_2BMS
        Cells(i, "O").Value = cool_2BMS
        Cells(i, "R").Value = pow_2BMS
        Cells(i, "Y").Value = massre_2BMS
        Cells(i, "AA").Value = TRL_2BMS
    
    End If
    
    If Cells(i, 1).Value = "Molecular sieves 4BMS or Carbon Dioxide Removal Assembly CDRA" Then
        
        Cells(i, "I").Value = mass_4BMS
        Cells(i, "L").Value = vol_4BMS
        Cells(i, "O").Value = cool_4BMS
        Cells(i, "R").Value = pow_4BMS
        Cells(i, "Y").Value = massre_4BMS
        Cells(i, "AA").Value = TRL_4BMS
    
    End If
    
    If Cells(i, 1).Value = "Multifiltration MF" Then
        
        Cells(i, "I").Value = mass_MF
        Cells(i, "L").Value = vol_MF
        Cells(i, "O").Value = cool_MF
        Cells(i, "R").Value = pow_MF
        Cells(i, "Y").Value = massre_MF
        Cells(i, "AA").Value = TRL_MF
    
    End If
    
    If Cells(i, 1).Value = "Sabatier Reactor SR" Then
        
        Cells(i, "I").Value = mass_SR
        Cells(i, "L").Value = vol_SR
        Cells(i, "O").Value = cool_SR
        Cells(i, "R").Value = pow_SR
        Cells(i, "Y").Value = massre_SR
        Cells(i, "AA").Value = TRL_SR
    
    End If
    
    If Cells(i, 1).Value = "Solid amine water desorption SAWD" Then
        
        Cells(i, "I").Value = mass_SAWD
        Cells(i, "L").Value = vol_SAWD
        Cells(i, "O").Value = cool_SAWD
        Cells(i, "R").Value = pow_SAWD
        Cells(i, "Y").Value = massre_SAWD
        Cells(i, "AA").Value = TRL_SAWD
    
    End If
    
    If Cells(i, 1).Value = "Solid polymer water electrolysis SPWE or Oxygen generation assembly OGS" _
        Or Cells(i, 1).Value = "Oxygen Generation Assembly OGS or OGA" Then
        
        Cells(i, "I").Value = mass_SPWE
        Cells(i, "L").Value = vol_SPWE
        Cells(i, "O").Value = cool_SPWE
        Cells(i, "R").Value = pow_SPWE
        Cells(i, "Y").Value = massre_SPWE
        Cells(i, "AA").Value = TRL_SPWE
    
    End If
    
    If Cells(i, 1).Value = "Static feed water electrolysis SFWE" Then
        
        Cells(i, "I").Value = mass_SFWE
        Cells(i, "L").Value = vol_SFWE
        Cells(i, "O").Value = cool_SFWE
        Cells(i, "R").Value = pow_SFWE
        Cells(i, "Y").Value = massre_SFWE
        Cells(i, "AA").Value = TRL_SFWE
    
    End If
    
    If Cells(i, 1).Value = "Thermoelectric Integrated Membrane Evaporation TIMES" Then
        
        Cells(i, "I").Value = mass_TIMES
        Cells(i, "L").Value = vol_TIMES
        Cells(i, "O").Value = cool_TIMES
        Cells(i, "R").Value = pow_TIMES
        Cells(i, "Y").Value = massre_TIMES
        Cells(i, "AA").Value = TRL_TIMES
    
    End If
    
    If Cells(i, 1).Value = "Vapor Compression Distillation VCD" Then
        
        Cells(i, "I").Value = mass_VCD
        Cells(i, "L").Value = vol_VCD
        Cells(i, "O").Value = cool_VCD
        Cells(i, "R").Value = pow_VCD
        Cells(i, "Y").Value = massre_VCD
        Cells(i, "AA").Value = TRL_VCD
    
    End If
    
    If Cells(i, 1).Value = "Vapor Phase Catalytic Ammonia Removal VPCAR" Then
        
        Cells(i, "I").Value = mass_VPCAR
        Cells(i, "L").Value = vol_VPCAR
        Cells(i, "O").Value = cool_VPCAR
        Cells(i, "R").Value = pow_VPCAR
        Cells(i, "Y").Value = massre_VPCAR
        Cells(i, "AA").Value = TRL_VPCAR
    
    End If
    
    If Cells(i, 1).Value = "Advanced Carbon-Formation Reactor System ACRS" Then
        
        Cells(i, "I").Value = mass_ACRS
        Cells(i, "L").Value = vol_ACRS
        Cells(i, "O").Value = cool_ACRS
        Cells(i, "R").Value = pow_ACRS
        Cells(i, "Y").Value = massre_ACRS
        Cells(i, "AA").Value = TRL_ACRS
    
    End If
    
    If Cells(i, 1).Value = "Lithium hydroxide LiOH" Then
        
        Cells(i, "I").Value = mass_LiOH
        Cells(i, "L").Value = vol_LiOH
        Cells(i, "O").Value = cool_LiOH
        Cells(i, "R").Value = pow_LiOH
        Cells(i, "Y").Value = massre_LiOH
        Cells(i, "AA").Value = TRL_LiOH
    
    End If
    
    If Cells(i, 1).Value = "Multi Bed Trace Contaminant Control TCCS" Then Cells(i, "Y").Value = 24.1 / 365 * mission_duration
    
    If Cells(i, 1).Value = "High Pressure Storage" Then

        If DES_Function = "Store Oxygen" Then
        
            Cells(i, "I").Value = O2_mass * tank_minc_O2_HPS
            Cells(i, "L").Value = (((O2_vol_HPS * 3 / (4 * WorksheetFunction.Pi)) ^ (1 / 3) + t_wall_HPS) ^ 3) * 4 * WorksheetFunction.Pi / (3 * tfd) + 0.064
            Cells(i, "O").Value = 0
            Cells(i, "R").Value = 9
            Cells(i, "Y").Value = O2_mass 'resupply assumed to be equal to O2 consuption of the crew nad leakage
            Cells(i, "AA").Value = TRL_HPS
         
         ElseIf DES_Function = "Store Inert Gas" Then
            
            Cells(i, "I").Value = N2_mass * tank_minc_N2_HPS
            Cells(i, "L").Value = 4 * WorksheetFunction.Pi / (3 * tfd) * ((N2_vol_HPS * 3 / (4 * WorksheetFunction.Pi)) ^ (1 / 3) + t_wall_HPS) ^ 3 + 0.064
            Cells(i, "O").Value = 0
            Cells(i, "R").Value = 9
            Cells(i, "Y").Value = N2_mass
            Cells(i, "AA").Value = TRL_HPS
            
        ElseIf DES_Function = "Recover from Rapid Decompression" Then
            
            Cells(i, "I").Value = N2_mass * tank_minc_N2_HPS + O2_mass * tank_minc_O2_HPS
            Cells(i, "L").Value = 4 * WorksheetFunction.Pi / (3 * tfd) * (((N2_vol_HPS + O2_vol_HPS) * 3 / (4 * WorksheetFunction.Pi)) ^ (1 / 3) + t_wall_HPS) ^ 3 + 0.064
            Cells(i, "O").Value = 0
            Cells(i, "R").Value = 9
            Cells(i, "Y").Value = N2_mass + O2_mass
            Cells(i, "AA").Value = TRL_HPS
            
        End If
        
    End If
    
    If Cells(i, 1).Value = "Cryogenic Storage" Then

        If DES_Function = "Store Oxygen" Then
        
            Cells(i, "I").Value = O2_mass * tank_minc_O2_CS
            Cells(i, "L").Value = (((O2_vol_CS * 3 / (4 * WorksheetFunction.Pi)) ^ (1 / 3) + t_wall_CS) ^ 3) * 4 * WorksheetFunction.Pi / (3 * tfd) + 0.064
            Cells(i, "O").Value = 0
            Cells(i, "R").Value = 9
            Cells(i, "Y").Value = O2_mass 'resupply assumed to be equal to O2 consuption of the crew nad leakage
            Cells(i, "AA").Value = TRL_CS
         
         ElseIf DES_Function = "Store Inert Gas" Then
            
            Cells(i, "I").Value = N2_mass * tank_minc_N2_CS
            Cells(i, "L").Value = 4 * WorksheetFunction.Pi / (3 * tfd) * ((N2_vol_CS * 3 / (4 * WorksheetFunction.Pi)) ^ (1 / 3) + t_wall_CS) ^ 3 + 0.064
            Cells(i, "O").Value = 0
            Cells(i, "R").Value = 9
            Cells(i, "Y").Value = N2_mass
            Cells(i, "AA").Value = TRL_CS
            
        ElseIf DES_Function = "Recover from Rapid Decompression" Then
            
            Cells(i, "I").Value = N2_mass * tank_minc_N2_CS + O2_mass * tank_minc_O2_CS
            Cells(i, "L").Value = 4 * WorksheetFunction.Pi / (3 * tfd) * (((N2_vol_CS + O2_vol_CS) * 3 / (4 * WorksheetFunction.Pi)) ^ (1 / 3) + t_wall_HPS) ^ 3 + 0.064
            Cells(i, "O").Value = 0
            Cells(i, "R").Value = 9
            Cells(i, "Y").Value = N2_mass + O2_mass
            Cells(i, "AA").Value = TRL_HPS
            
        End If
        
    End If
    
    If Cells(i, 1).Value = "Oxygen Candles LiClO4" Then
        
        Cells(i, "I").Value = mass_O2c
        Cells(i, "L").Value = vol_O2c
        Cells(i, "O").Value = cool_O2c
        Cells(i, "R").Value = pow_O2c
        Cells(i, "Y").Value = massre_O2c
        Cells(i, "AA").Value = TRL_O2c
    
    End If
    
    If Cells(i, 1).Value = "Water tank" Then
        
        If DES_Function = "Store Water" Then
        
            d_wt = 2 * ((3 / 4) * (H2O_vol_w_wt / (0.91 * WorksheetFunction.Pi))) ^ (1 / 3)
        
            Cells(i, "I").Value = WorksheetFunction.Pi * d_wt ^ 2 * rho_s * t_wall_wt_s + WorksheetFunction.Pi * d_wt ^ 2 * rho_b * t_wall_wt_b + 2.3
            Cells(i, "L").Value = H2O_vol_w_wt / 0.91
            Cells(i, "O").Value = cool_w_wt
            Cells(i, "R").Value = pow_w_wt
            Cells(i, "Y").Value = massre_w_wt
            Cells(i, "AA").Value = TRL_wt
    
        End If
        
        If DES_Function = "Store Wastewater" Then
        
            d_wt = 2 * ((3 / 4) * (H2O_vol_ww_wt / (0.91 * WorksheetFunction.Pi))) ^ (1 / 3)
        
            Cells(i, "I").Value = WorksheetFunction.Pi * d_wt ^ 2 * rho_s * t_wall_wt_s + WorksheetFunction.Pi * d_wt ^ 2 * rho_b * t_wall_wt_b + 2.3
            Cells(i, "L").Value = H2O_vol_ww_wt / 0.91
            Cells(i, "O").Value = cool_ww_wt
            Cells(i, "R").Value = pow_ww_wt
            Cells(i, "Y").Value = massre_ww_wt
            Cells(i, "AA").Value = TRL_wt
    
        End If
    
    End If
    
    If Cells(i, 1).Value = "Algae Reactor" Then
    
        If Cells(i, 2).Value = "air-AR" Then
    
            Cells(i, "I").Value = mass_alg
            Cells(i, "L").Value = vol_alg
            Cells(i, "O").Value = cool_alg
            Cells(i, "R").Value = power_alg
            Cells(i, "Y").Value = massre_alg
            Cells(i, "AA").Value = TRL_alg
        
        ElseIf Cells(i, 2).Value = "water-WRM" Then
        
            Cells(i, "I").Value = mass_alg
            Cells(i, "L").Value = vol_alg
            Cells(i, "O").Value = cool_alg
            Cells(i, "R").Value = power_alg
            Cells(i, "Y").Value = massre_alg
            Cells(i, "AA").Value = TRL_alg
            
        End If
        
    End If
    
Next i


End Sub


Sub rescaletech()

'Contigency Faktoren

Dim Con1_mass50
Dim Con2_mass50
Dim Con3_mass50
Dim Con1_mass500
Dim Con2_mass500
Dim Con3_mass500
Dim Con1_mass2500
Dim Con2_mass2500
Dim Con3_mass2500
Dim Con1_mass2500_
Dim Con2_mass2500_
Dim Con3_mass2500_
Dim Con1_power50
Dim Con2_power50
Dim Con3_power50
Dim Con1_power500
Dim Con2_power500
Dim Con3_power500
Dim Con1_power2500
Dim Con2_power2500
Dim Con3_power2500
Dim Con1_power2500_
Dim Con2_power2500_
Dim Con3_power2500_

Con1_mass50 = 35 / 100
Con2_mass50 = 25 / 100
Con3_mass50 = 3 / 100
Con1_mass500 = 30 / 100
Con2_mass500 = 20 / 100
Con3_mass500 = 3 / 100
Con1_mass2500 = 25 / 100
Con2_mass2500 = 15 / 100
Con3_mass2500 = 1 / 100
Con1_mass2500_ = 22 / 100
Con2_mass2500_ = 12 / 100
Con3_mass2500_ = 0.8 / 100

Con1_power50 = 75 / 100
Con2_power50 = 25 / 100
Con3_power50 = 12 / 100
Con1_power500 = 65 / 100
Con2_power500 = 22 / 100
Con3_power500 = 12 / 100
Con1_power2500 = 60 / 100
Con2_power2500 = 20 / 100
Con3_power2500 = 12 / 100
Con1_power2500_ = 35 / 100
Con2_power2500_ = 20 / 100
Con3_power2500_ = 11 / 100

Dim i
Dim na
Dim Crew_Size
Dim mission_duration
Dim Schedule_number
Dim CO2_in
Dim O2_out
Dim MM_O2
Dim MM_CO2
Dim MM_H2O
Dim MM_H2
Dim MM_CH4
Dim m_CO2rd_a

Crew_Size = Worksheets("User Interface").Range("Crew_Size").Value
mission_duration = Worksheets("User Interface").Range("Mission_Duration").Value
Schedule_number = Worksheets("User Interface").Range("Schedule_Number").Value

MM_O2 = 32 '[g/mol]
MM_CO2 = 44.01 '[g/mol]
MM_H2O = 18.01 '[g/mol]
MM_H2 = 2.01 '[g/mol]
MM_CH4 = 16.04 '[g/mol]

CO2_in = Worksheets("ECLSS Composition").Range("O5").Value
O2_out = Worksheets("ECLSS Composition").Range("R18").Value

na = Worksheets("Technology").Range("A2:A1000").Cells.SpecialCells(xlCellTypeConstants).Count

For i = 2 To na

    If Cells(i, 1).Value = "CO2 Reduction Assembly" Then m_CO2rd_a = Cells(i, "I").Value
    
Next i


'Bosch

Dim SF_Bs 'Scaling Faktor
Dim RF_Bs 'Rescale factor for Bosch
Dim massg_Bs 'given system mass for Bosch
Dim massreg_Bs 'given resupply mass for Bosch
Dim volg_Bs 'given volume for Bosch
Dim powg_Bs 'given power for Bosch
Dim coolg_Bs ' given cooling for Bosch
Dim CO2_in_Bs 'given CO2 input
Dim H2_in_Bs 'needed H2 input
Dim H2O_out_Bs 'produced H2O output
Dim TRL_Bs
Dim eff_Bs

Dim mass_Bs
Dim massre_Bs
Dim vol_Bs
Dim pow_Bs
Dim cool_Bs
Dim massc_Bs
Dim powc_Bs

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_Bs = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_Bs = 0.9
If Crew_Size >= 10 Then SF_Bs = 0.8

eff_Bs = 1
massg_Bs = 102.1
massreg_Bs = 0.028
volg_Bs = 0.3
powg_Bs = 950
coolg_Bs = 313
CO2_in_Bs = 4
TRL_Bs = 6

RF_Bs = CO2_in / CO2_in_Bs
H2_in_Bs = 2 * MM_H2 / MM_CO2 * CO2_in * Crew_Size
H2O_out_Bs = 2 * MM_H2O / MM_CO2 * CO2_in * Crew_Size * eff_Bs

massc_Bs = RF_Bs * (massg_Bs + m_CO2rd_a) * Crew_Size * SF_Bs

If TRL_Bs > 1 And TRL_Bs < 4 Then

    If massc_Bs < 50 Then mass_Bs = (1 + Con1_mass50) * massc_Bs
    If massc_Bs > 50 And massc_Bs < 500 Then mass_Bs = (1 + Con1_mass500) * massc_Bs
    If massc_Bs > 500 And massc_Bs < 2500 Then mass_Bs = (1 + Con1_mass2500) * massc_Bs
    If massc_Bs > 2500 Then mass_Bs = (1 + Con1_mass2500_) * massc_Bs

End If

If TRL_Bs >= 4 And TRL_Bs < 7 Then

    If massc_Bs < 50 Then mass_Bs = (1 + Con2_mass50) * massc_Bs
    If massc_Bs > 50 And massc_Bs < 500 Then mass_Bs = (1 + Con2_mass500) * massc_Bs
    If massc_Bs > 500 And massc_Bs < 2500 Then mass_Bs = (1 + Con2_mass2500) * massc_Bs
    If massc_Bs > 2500 Then mass_Bs = (1 + Con2_mass2500_) * massc_Bs

End If

If TRL_Bs >= 7 And TRL_Bs < 10 Then

    If massc_Bs < 50 Then mass_Bs = (1 + Con3_mass50) * massc_Bs
    If massc_Bs > 50 And massc_Bs < 500 Then mass_Bs = (1 + Con3_mass500) * massc_Bs
    If massc_Bs > 500 And massc_Bs < 2500 Then mass_Bs = (1 + Con3_mass2500) * massc_Bs
    If massc_Bs > 2500 Then mass_Bs = (1 + Con3_mass2500_) * massc_Bs

End If

massre_Bs = RF_Bs * massreg_Bs * Crew_Size * SF_Bs * mission_duration
vol_Bs = RF_Bs * volg_Bs * Crew_Size * SF_Bs
cool_Bs = RF_Bs * coolg_Bs * Crew_Size * SF_Bs
powc_Bs = RF_Bs * powg_Bs * Crew_Size * SF_Bs

If TRL_Bs > 1 And TRL_Bs < 4 Then

    If powc_Bs < 50 Then pow_Bs = (1 + Con1_power50) * powc_Bs
    If powc_Bs > 50 And powc_Bs < 500 Then pow_Bs = (1 + Con1_power500) * powc_Bs
    If powc_Bs > 500 And powc_Bs < 2500 Then pow_Bs = (1 + Con1_power2500) * powc_Bs
    If powc_Bs > 2500 Then pow_Bs = (1 + Con1_power2500_) * powc_Bs

End If

If TRL_Bs >= 4 And TRL_Bs < 7 Then

    If powc_Bs < 50 Then pow_Bs = (1 + Con2_power50) * powc_Bs
    If powc_Bs > 50 And powc_Bs < 500 Then pow_Bs = (1 + Con2_power500) * powc_Bs
    If powc_Bs > 500 And powc_Bs < 2500 Then pow_Bs = (1 + Con2_power2500) * powc_Bs
    If powc_Bs > 2500 Then pow_Bs = (1 + Con2_power2500_) * powc_Bs

End If

If TRL_Bs >= 7 And TRL_Bs < 10 Then

    If powc_Bs < 50 Then pow_Bs = (1 + Con3_power50) * powc_Bs
    If powc_Bs > 50 And powc_Bs < 500 Then pow_Bs = (1 + Con3_power500) * powc_Bs
    If powc_Bs > 500 And powc_Bs < 2500 Then pow_Bs = (1 + Con3_power2500) * powc_Bs
    If powc_Bs > 2500 Then pow_Bs = (1 + Con3_power2500_) * powc_Bs

End If

'COE:

'Data is given for 36.8 kg/d CO2, resupply mass is not provided

Dim SF_COE 'Scaling Faktor
Dim RF_COE 'is currently set to 1 due to lack of further information
Dim massg_COE 'given system mass for COE
Dim massreg_COE 'given resupply mass for COE
Dim volg_COE 'given volume for COE
Dim powg_COE 'given power for COE
Dim coolg_COE ' given cooling for COE
Dim CO2_in_COE 'given CO2 input
Dim O2_out_COE 'produced O2 output
Dim eff_COE 'unknown
Dim TRL_COE

Dim mass_COE
Dim massre_COE
Dim vol_COE
Dim pow_COE
Dim cool_COE
Dim massc_COE
Dim powc_COE

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_COE = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_COE = 0.9
If Crew_Size >= 10 Then SF_COE = 0.8

massg_COE = 356
massreg_COE = 0
volg_COE = 0.28
powg_COE = 13300
coolg_COE = 3500
TRL_COE = 4
CO2_in_COE = 36.8

RF_COE = CO2_in / CO2_in_COE
O2_out_COE = MM_O2 / MM_CO2 * CO2_in * Crew_Size * eff_COE
massc_COE = RF_COE * (massg_COE + m_CO2rd_a) * Crew_Size * SF_COE

If TRL_COE > 1 And TRL_COE < 4 Then

    If massc_COE < 50 Then mass_COE = (1 + Con1_mass50) * massc_COE
    If massc_COE > 50 And massc_COE < 500 Then mass_COE = (1 + Con1_mass500) * massc_COE
    If massc_COE > 500 And massc_COE < 2500 Then mass_COE = (1 + Con1_mass2500) * massc_COE
    If massc_COE > 2500 Then mass_COE = (1 + Con1_mass2500_) * massc_COE

End If

If TRL_COE >= 4 And TRL_COE < 7 Then

    If massc_COE < 50 Then mass_COE = (1 + Con2_mass50) * massc_COE
    If massc_COE > 50 And massc_COE < 500 Then mass_COE = (1 + Con2_mass500) * massc_COE
    If massc_COE > 500 And massc_COE < 2500 Then mass_COE = (1 + Con2_mass2500) * massc_COE
    If massc_COE > 2500 Then mass_COE = (1 + Con2_mass2500_) * massc_COE

End If

If TRL_COE >= 7 And TRL_COE < 10 Then

    If massc_COE < 50 Then mass_COE = (1 + Con3_mass50) * massc_COE
    If massc_COE > 50 And massc_COE < 500 Then mass_COE = (1 + Con3_mass500) * massc_COE
    If massc_COE > 500 And massc_COE < 2500 Then mass_COE = (1 + Con3_mass2500) * massc_COE
    If massc_COE > 2500 Then mass_COE = (1 + Con3_mass2500_) * massc_COE

End If

massre_COE = 0 'No information as of right now
vol_COE = RF_COE * volg_COE * Crew_Size * SF_COE
cool_COE = RF_COE * coolg_COE * Crew_Size * SF_COE
powc_COE = RF_COE * powg_COE * Crew_Size * SF_COE

If TRL_COE > 1 And TRL_COE < 4 Then

    If powc_COE < 50 Then pow_COE = (1 + Con1_power50) * powc_COE
    If powc_COE > 50 And powc_COE < 500 Then pow_COE = (1 + Con1_power500) * powc_COE
    If powc_COE > 500 And powc_COE < 2500 Then pow_COE = (1 + Con1_power2500) * powc_COE
    If powc_COE > 2500 Then pow_COE = (1 + Con1_power2500_) * powc_COE

End If

If TRL_COE >= 4 And TRL_COE < 7 Then

    If powc_COE < 50 Then pow_COE = (1 + Con2_power50) * powc_COE
    If powc_COE > 50 And powc_COE < 500 Then pow_COE = (1 + Con2_power500) * powc_COE
    If powc_COE > 500 And powc_COE < 2500 Then pow_COE = (1 + Con2_power2500) * powc_COE
    If powc_COE > 2500 Then pow_COE = (1 + Con2_power2500_) * powc_COE

End If

If TRL_COE >= 7 And TRL_COE < 10 Then

    If powc_COE < 50 Then pow_COE = (1 + Con3_power50) * powc_COE
    If powc_COE > 50 And powc_COE < 500 Then pow_COE = (1 + Con3_power500) * powc_COE
    If powc_COE > 500 And powc_COE < 2500 Then pow_COE = (1 + Con3_power2500) * powc_COE
    If powc_COE > 2500 Then pow_COE = (1 + Con3_power2500_) * powc_COE

End If

'Sabatier

Sabatier:

'Data given for CO2 input of 5 kg/d CO2

Dim SF_SR 'Scaling Faktor
Dim RF_SR 'Rescale factor for SR
Dim massg_SR 'given system mass for SR
Dim massreg_SR 'given resupply mass for SR
Dim volg_SR 'given volume for SR
Dim powg_SR 'given power for SR
Dim coolg_SR ' given cooling for SR
Dim CO2_in_SR 'given CO2 input
Dim H2_in_SR 'needed H2 input
Dim WW_out_SR 'generated waste water output
Dim CH4_out_SR 'generated methane output
Dim eff_SR
Dim TRL_SR

Dim mass_SR
Dim massre_SR
Dim vol_SR
Dim pow_SR
Dim cool_SR
Dim massc_SR
Dim powc_SR

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_SR = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_SR = 0.9
If Crew_Size >= 10 Then SF_SR = 0.8

eff_SR = 0.99
massg_SR = 43
massreg_SR = 0.012
volg_SR = 0.18
powg_SR = 50
coolg_SR = 289
CO2_in_SR = 5
TRL_SR = 9

RF_SR = CO2_in / CO2_in_SR
massc_SR = RF_SR * (massg_SR + m_CO2rd_a) * Crew_Size * SF_SR
H2_in_SR = 4 * MM_H2 / MM_CO2 * Crew_Size * CO2_in
WW_out_SR = 2 * MM_H2O / MM_CO2 * Crew_Size * CO2_in * eff_SR
CH4_out_SR = MM_CH4 / MM_CO2 + Crew_Size * CO2_in * eff_SR


If TRL_SR > 1 And TRL_SR < 4 Then

    If massc_SR < 50 Then mass_SR = (1 + Con1_mass50) * massc_SR
    If massc_SR > 50 And massc_SR < 500 Then mass_SR = (1 + Con1_mass500) * massc_SR
    If massc_SR > 500 And massc_SR < 2500 Then mass_SR = (1 + Con1_mass2500) * massc_SR
    If massc_SR > 2500 Then mass_SR = (1 + Con1_mass2500_) * massc_SR

End If

If TRL_SR >= 4 And TRL_SR < 7 Then

    If massc_SR < 50 Then mass_SR = (1 + Con2_mass50) * massc_SR
    If massc_SR > 50 And massc_SR < 500 Then mass_SR = (1 + Con2_mass500) * massc_SR
    If massc_SR > 500 And massc_SR < 2500 Then mass_SR = (1 + Con2_mass2500) * massc_SR
    If massc_SR > 2500 Then mass_SR = (1 + Con2_mass2500_) * massc_SR

End If

If TRL_SR >= 7 And TRL_SR < 10 Then

    If massc_SR < 50 Then mass_SR = (1 + Con3_mass50) * massc_SR
    If massc_SR > 50 And massc_SR < 500 Then mass_SR = (1 + Con3_mass500) * massc_SR
    If massc_SR > 500 And massc_SR < 2500 Then mass_SR = (1 + Con3_mass2500) * massc_SR
    If massc_SR > 2500 Then mass_SR = (1 + Con3_mass2500_) * massc_SR

End If

massre_SR = RF_SR * massreg_SR * Crew_Size * SF_SR * mission_duration
vol_SR = RF_SR * volg_SR * Crew_Size * SF_SR
cool_SR = RF_SR * coolg_SR * Crew_Size * SF_SR
powc_SR = RF_SR * powg_SR * Crew_Size * SF_SR

If TRL_SR > 1 And TRL_SR < 4 Then

    If powc_SR < 50 Then pow_SR = (1 + Con1_power50) * powc_SR
    If powc_SR > 50 And powc_SR < 500 Then pow_SR = (1 + Con1_power500) * powc_SR
    If powc_SR > 500 And powc_SR < 2500 Then pow_SR = (1 + Con1_power2500) * powc_SR
    If powc_SR > 2500 Then pow_SR = (1 + Con1_power2500_) * powc_SR

End If

If TRL_SR >= 4 And TRL_SR < 7 Then

    If powc_SR < 50 Then pow_SR = (1 + Con2_power50) * powc_SR
    If powc_SR > 50 And powc_SR < 500 Then pow_SR = (1 + Con2_power500) * powc_SR
    If powc_SR > 500 And powc_SR < 2500 Then pow_SR = (1 + Con2_power2500) * powc_SR
    If powc_SR > 2500 Then pow_SR = (1 + Con2_power2500_) * powc_SR

End If

If TRL_SR >= 7 And TRL_SR < 10 Then

    If powc_SR < 50 Then pow_SR = (1 + Con3_power50) * powc_SR
    If powc_SR > 50 And powc_SR < 500 Then pow_SR = (1 + Con3_power500) * powc_SR
    If powc_SR > 500 And powc_SR < 2500 Then pow_SR = (1 + Con3_power2500) * powc_SR
    If powc_SR > 2500 Then pow_SR = (1 + Con3_power2500_) * powc_SR

End If

'MF:

'Data given for waterrecovery for waste water input of 50.9 kg/d
'Waste water (=assumption: hygiene + ww from urine process+potential other component -> unkown yet, set to 0) in and H2O out (useable for hygiene and drinking)

Dim SF_MF 'Scaling Faktor
Dim RF_MF 'Rescale factor for MF
Dim massg_MF 'given system mass for MF
Dim massreg_MF 'given resupply mass for MF
Dim volg_MF 'given volume for MF
Dim powg_MF 'given power for MF
Dim coolg_MF ' given cooling for MF
Dim H2O_out_MF 'given waste water output
Dim ww_MF
Dim TRL_MF

Dim mass_MF
Dim massre_MF
Dim vol_MF
Dim pow_MF
Dim cool_MF
Dim massc_MF
Dim powc_MF

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_MF = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_MF = 0.9
If Crew_Size >= 10 Then SF_MF = 0.8

massg_MF = 489.1
massreg_MF = 1.64
volg_MF = 2.26
powg_MF = 1021
coolg_MF = 1021
H2O_out_MF = 50.9
TRL_MF = 9

ww_MF = Worksheets("ECLSS Composition").Range("H22").Value

RF_MF = ww_MF / H2O_out_MF
massc_MF = RF_MF * massg_MF * Crew_Size * SF_MF

If TRL_MF > 1 And TRL_MF < 4 Then

    If massc_MF < 50 Then mass_MF = (1 + Con1_mass50) * massc_MF
    If massc_MF > 50 And massc_MF < 500 Then mass_MF = (1 + Con1_mass500) * massc_MF
    If massc_MF > 500 And massc_MF < 2500 Then mass_MF = (1 + Con1_mass2500) * massc_MF
    If massc_MF > 2500 Then mass_MF = (1 + Con1_mass2500_) * massc_MF

End If

If TRL_MF >= 4 And TRL_MF < 7 Then

    If massc_MF < 50 Then mass_MF = (1 + Con2_mass50) * massc_MF
    If massc_MF > 50 And massc_MF < 500 Then mass_MF = (1 + Con2_mass500) * massc_MF
    If massc_MF > 500 And massc_MF < 2500 Then mass_MF = (1 + Con2_mass2500) * massc_MF
    If massc_MF > 2500 Then mass_MF = (1 + Con2_mass2500_) * massc_MF

End If

If TRL_MF >= 7 And TRL_MF < 10 Then

    If massc_MF < 50 Then mass_MF = (1 + Con3_mass50) * massc_MF
    If massc_MF > 50 And massc_MF < 500 Then mass_MF = (1 + Con3_mass500) * massc_MF
    If massc_MF > 500 And massc_MF < 2500 Then mass_MF = (1 + Con3_mass2500) * massc_MF
    If massc_MF > 2500 Then mass_MF = (1 + Con3_mass2500_) * massc_MF

End If

massre_MF = RF_MF * massreg_MF * Crew_Size * SF_MF * mission_duration
vol_MF = RF_MF * volg_MF * Crew_Size * SF_MF
cool_MF = RF_MF * coolg_MF * Crew_Size * SF_MF
powc_MF = RF_MF * powg_MF * Crew_Size * SF_MF

If TRL_MF > 1 And TRL_MF < 4 Then

    If powc_MF < 50 Then pow_MF = (1 + Con1_power50) * powc_MF
    If powc_MF > 50 And powc_MF < 500 Then pow_MF = (1 + Con1_power500) * powc_MF
    If powc_MF > 500 And powc_MF < 2500 Then pow_MF = (1 + Con1_power2500) * powc_MF
    If powc_MF > 2500 Then pow_MF = (1 + Con1_power2500_) * powc_MF

End If

If TRL_MF >= 4 And TRL_MF < 7 Then

    If powc_MF < 50 Then pow_MF = (1 + Con2_power50) * powc_MF
    If powc_MF > 50 And powc_MF < 500 Then pow_MF = (1 + Con2_power500) * powc_MF
    If powc_MF > 500 And powc_MF < 2500 Then pow_MF = (1 + Con2_power2500) * powc_MF
    If powc_MF > 2500 Then pow_MF = (1 + Con2_power2500_) * powc_MF

End If

If TRL_MF >= 7 And TRL_MF < 10 Then

    If powc_MF < 50 Then pow_MF = (1 + Con3_power50) * powc_MF
    If powc_MF > 50 And powc_MF < 500 Then pow_MF = (1 + Con3_power500) * powc_MF
    If powc_MF > 500 And powc_MF < 2500 Then pow_MF = (1 + Con3_power2500) * powc_MF
    If powc_MF > 2500 Then pow_MF = (1 + Con3_power2500_) * powc_MF

End If

'SPWE:

'Data given for output of 5.3 kg/d O2, requires H2O input, generates H2

Dim SF_SPWE 'Scaling Faktor
Dim RF_SPWE 'Rescale factor for SPWE
Dim massg_SPWE 'given system mass for SPWE
Dim massreg_SPWE 'given resupply mass for SPWE
Dim volg_SPWE 'given volume for SPWE
Dim powg_SPWE 'given power for SPWE
Dim coolg_SPWE ' given cooling for SPWE
Dim H2O_in_SPWE 'needed H2O input
Dim H2_out_SPWE 'generated H2
Dim O2_out_SPWE 'generated O2 output
Dim eff_SPWE 'unknown yet
Dim TRL_SPWE

Dim mass_SPWE
Dim massre_SPWE
Dim vol_SPWE
Dim pow_SPWE
Dim cool_SPWE
Dim massc_SPWE
Dim powc_SPWE

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_SPWE = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_SPWE = 0.9
If Crew_Size >= 10 Then SF_SPWE = 0.8

eff_SPWE = 1
massg_SPWE = 416
massreg_SPWE = 0.114
volg_SPWE = 1.7
powg_SPWE = 1035
coolg_SPWE = 26.5
O2_out_SPWE = 5.3
TRL_SPWE = 9

RF_SPWE = O2_out / (O2_out_SPWE * eff_SPWE)
massc_SPWE = RF_SPWE * massg_SPWE * Crew_Size * SF_SPWE
H2O_in_SPWE = 2 * MM_H2O / MM_O2 * Crew_Size * O2_out
H2_out_SPWE = 4 * MM_H2 / MM_O2 * Crew_Size * O2_out * eff_SPWE

If TRL_SPWE > 1 And TRL_SPWE < 4 Then

    If massc_SPWE < 50 Then mass_SPWE = (1 + Con1_mass50) * massc_SPWE
    If massc_SPWE > 50 And massc_SPWE < 500 Then mass_SPWE = (1 + Con1_mass500) * massc_SPWE
    If massc_SPWE > 500 And massc_SPWE < 2500 Then mass_SPWE = (1 + Con1_mass2500) * massc_SPWE
    If massc_SPWE > 2500 Then mass_SPWE = (1 + Con1_mass2500_) * massc_SPWE

End If

If TRL_SPWE >= 4 And TRL_SPWE < 7 Then

    If massc_SPWE < 50 Then mass_SPWE = (1 + Con2_mass50) * massc_SPWE
    If massc_SPWE > 50 And massc_SPWE < 500 Then mass_SPWE = (1 + Con2_mass500) * massc_SPWE
    If massc_SPWE > 500 And massc_SPWE < 2500 Then mass_SPWE = (1 + Con2_mass2500) * massc_SPWE
    If massc_SPWE > 2500 Then mass_SPWE = (1 + Con2_mass2500_) * massc_SPWE

End If

If TRL_SPWE >= 7 And TRL_SPWE < 10 Then

    If massc_SPWE < 50 Then mass_SPWE = (1 + Con3_mass50) * massc_SPWE
    If massc_SPWE > 50 And massc_SPWE < 500 Then mass_SPWE = (1 + Con3_mass500) * massc_SPWE
    If massc_SPWE > 500 And massc_SPWE < 2500 Then mass_SPWE = (1 + Con3_mass2500) * massc_SPWE
    If massc_SPWE > 2500 Then mass_SPWE = (1 + Con3_mass2500_) * massc_SPWE

End If

massre_SPWE = RF_SPWE * massreg_SPWE * Crew_Size * SF_SPWE * mission_duration
vol_SPWE = RF_SPWE * volg_SPWE * Crew_Size * SF_SPWE
cool_SPWE = RF_SPWE * coolg_SPWE * Crew_Size * SF_SPWE
powc_SPWE = RF_SPWE * powg_SPWE * Crew_Size * SF_SPWE

If TRL_SPWE > 1 And TRL_SPWE < 4 Then

    If powc_SPWE < 50 Then pow_SPWE = (1 + Con1_power50) * powc_SPWE
    If powc_SPWE > 50 And powc_SPWE < 500 Then pow_SPWE = (1 + Con1_power500) * powc_SPWE
    If powc_SPWE > 500 And powc_SPWE < 2500 Then pow_SPWE = (1 + Con1_power2500) * powc_SPWE
    If powc_SPWE > 2500 Then pow_SPWE = (1 + Con1_power2500_) * powc_SPWE

End If

If TRL_SPWE >= 4 And TRL_SPWE < 7 Then

    If powc_SPWE < 50 Then pow_SPWE = (1 + Con2_power50) * powc_SPWE
    If powc_SPWE > 50 And powc_SPWE < 500 Then pow_SPWE = (1 + Con2_power500) * powc_SPWE
    If powc_SPWE > 500 And powc_SPWE < 2500 Then pow_SPWE = (1 + Con2_power2500) * powc_SPWE
    If powc_SPWE > 2500 Then pow_SPWE = (1 + Con2_power2500_) * powc_SPWE

End If

If TRL_SPWE >= 7 And TRL_SPWE < 10 Then

    If powc_SPWE < 50 Then pow_SPWE = (1 + Con3_power50) * powc_SPWE
    If powc_SPWE > 50 And powc_SPWE < 500 Then pow_SPWE = (1 + Con3_power500) * powc_SPWE
    If powc_SPWE > 500 And powc_SPWE < 2500 Then pow_SPWE = (1 + Con3_power2500) * powc_SPWE
    If powc_SPWE > 2500 Then pow_SPWE = (1 + Con3_power2500_) * powc_SPWE

End If


'SFWE:

'Data given for 4 CM -> output of 3.28 kg/d O2 (4*O2_Con_nominal), requires H2O input, generates H2 (O2_Con nominal=0.82)

Dim SF_SFWE 'Scaling Faktor
Dim RF_SFWE 'Rescale factor for SFWE
Dim massg_SFWE 'given system mass for SFWE
Dim massreg_SFWE 'given resupply mass for SFWE
Dim volg_SFWE 'given volume for SFWE
Dim powg_SFWE 'given power for SFWE
Dim coolg_SFWE ' given cooling for SFWE
Dim H2O_in_SFWE 'needed H2O input
Dim H2_out_SFWE 'generated H2
Dim O2_out_SFWE 'generated O2 output
Dim TRL_SFWE
Dim eff_SFWE 'unknown

Dim mass_SFWE
Dim massre_SFWE
Dim vol_SFWE
Dim pow_SFWE
Dim cool_SFWE
Dim massc_SFWE
Dim powc_SFWE

'Assignment

'Scaling factor

If Crew_Size >= 1 And Crew_Size < 7 Then SF_SFWE = 1
If Crew_Size >= 7 And Crew_Size < 10 Then SF_SFWE = 0.9
If Crew_Size >= 10 Then SF_SFWE = 0.8

eff_SFWE = 1
massg_SFWE = 113
massreg_SFWE = 0 'Unknown
volg_SFWE = 0.14
powg_SFWE = 1470
coolg_SFWE = 39
O2_out_SFWE = 4 * 0.82
TRL_SFWE = 8
eff_SFWE = 1 'Unknown yet

RF_SFWE = O2_out / O2_out_SFWE
massc_SFWE = RF_SFWE * massg_SFWE * Crew_Size * SF_SFWE
H2O_in_SFWE = 2 * MM_H2O / MM_O2 * Crew_Size * O2_out
H2_out_SFWE = 4 * MM_H2 / MM_O2 * Crew_Size * O2_out

If TRL_SFWE > 1 And TRL_SFWE < 4 Then

    If massc_SFWE < 50 Then mass_SFWE = (1 + Con1_mass50) * massc_SFWE
    If massc_SFWE > 50 And massc_SFWE < 500 Then mass_SFWE = (1 + Con1_mass500) * massc_SFWE
    If massc_SFWE > 500 And massc_SFWE < 2500 Then mass_SFWE = (1 + Con1_mass2500) * massc_SFWE
    If massc_SFWE > 2500 Then mass_SFWE = (1 + Con1_mass2500_) * massc_SFWE

End If

If TRL_SFWE >= 4 And TRL_SFWE < 7 Then

    If massc_SFWE < 50 Then mass_SFWE = (1 + Con2_mass50) * massc_SFWE
    If massc_SFWE > 50 And massc_SFWE < 500 Then mass_SFWE = (1 + Con2_mass500) * massc_SFWE
    If massc_SFWE > 500 And massc_SFWE < 2500 Then mass_SFWE = (1 + Con2_mass2500) * massc_SFWE
    If massc_SFWE > 2500 Then mass_SFWE = (1 + Con2_mass2500_) * massc_SFWE

End If

If TRL_SFWE >= 7 And TRL_SFWE < 10 Then

    If massc_SFWE < 50 Then mass_SFWE = (1 + Con3_mass50) * massc_SFWE
    If massc_SFWE > 50 And massc_SFWE < 500 Then mass_SFWE = (1 + Con3_mass500) * massc_SFWE
    If massc_SFWE > 500 And massc_SFWE < 2500 Then mass_SFWE = (1 + Con3_mass2500) * massc_SFWE
    If massc_SFWE > 2500 Then mass_SFWE = (1 + Con3_mass2500_) * massc_SFWE

End If

massre_SFWE = RF_SFWE * massreg_SFWE * Crew_Size * SF_SFWE * mission_duration
vol_SFWE = RF_SFWE * volg_SFWE * Crew_Size * SF_SFWE
cool_SFWE = RF_SFWE * coolg_SFWE * Crew_Size * SF_SFWE
powc_SFWE = RF_SFWE * powg_SFWE * Crew_Size * SF_SFWE

If TRL_SFWE > 1 And TRL_SFWE < 4 Then

    If powc_SFWE < 50 Then pow_SFWE = (1 + Con1_power50) * powc_SFWE
    If powc_SFWE > 50 And powc_SFWE < 500 Then pow_SFWE = (1 + Con1_power500) * powc_SFWE
    If powc_SFWE > 500 And powc_SFWE < 2500 Then pow_SFWE = (1 + Con1_power2500) * powc_SFWE
    If powc_SFWE > 2500 Then pow_SFWE = (1 + Con1_power2500_) * powc_SFWE

End If

If TRL_SFWE >= 4 And TRL_SFWE < 7 Then

    If powc_SFWE < 50 Then pow_SFWE = (1 + Con2_power50) * powc_SFWE
    If powc_SFWE > 50 And powc_SFWE < 500 Then pow_SFWE = (1 + Con2_power500) * powc_SFWE
    If powc_SFWE > 500 And powc_SFWE < 2500 Then pow_SFWE = (1 + Con2_power2500) * powc_SFWE
    If powc_SFWE > 2500 Then pow_SFWE = (1 + Con2_power2500_) * powc_SFWE

End If

If TRL_SFWE >= 7 And TRL_SFWE < 10 Then

    If powc_SFWE < 50 Then pow_SFWE = (1 + Con3_power50) * powc_SFWE
    If powc_SFWE > 50 And powc_SFWE < 500 Then pow_SFWE = (1 + Con3_power500) * powc_SFWE
    If powc_SFWE > 500 And powc_SFWE < 2500 Then pow_SFWE = (1 + Con3_power2500) * powc_SFWE
    If powc_SFWE > 2500 Then pow_SFWE = (1 + Con3_power2500_) * powc_SFWE

End If

'Write results in sheet

For i = 2 To na

    If Cells(i, 1).Value = "Bosch Reactor BR" Then
        
        Cells(i, "I").Value = mass_Bs
        Cells(i, "L").Value = vol_Bs
        Cells(i, "O").Value = cool_Bs
        Cells(i, "R").Value = pow_Bs
        Cells(i, "Y").Value = massre_Bs
        Cells(i, "AA").Value = TRL_Bs
    
    End If
    
    If Cells(i, 1).Value = "CO2 electrolysis COE" Then
        
        Cells(i, "I").Value = mass_COE
        Cells(i, "L").Value = vol_COE
        Cells(i, "O").Value = cool_COE
        Cells(i, "R").Value = pow_COE
        Cells(i, "Y").Value = massre_COE
        Cells(i, "AA").Value = TRL_COE
    
    End If
    
    If Cells(i, 1).Value = "Multifiltration MF" Then
        
        Cells(i, "I").Value = mass_MF
        Cells(i, "L").Value = vol_MF
        Cells(i, "O").Value = cool_MF
        Cells(i, "R").Value = pow_MF
        Cells(i, "Y").Value = massre_MF
        Cells(i, "AA").Value = TRL_MF
    
    End If
    
    If Cells(i, 1).Value = "Sabatier Reactor SR" Then
        
        Cells(i, "I").Value = mass_SR
        Cells(i, "L").Value = vol_SR
        Cells(i, "O").Value = cool_SR
        Cells(i, "R").Value = pow_SR
        Cells(i, "Y").Value = massre_SR
        Cells(i, "AA").Value = TRL_SR
    
    End If
    
    If Cells(i, 1).Value = "Solid polymer water electrolysis SPWE or Oxygen generation assembly OGS" _
        Or Cells(i, 1).Value = "Oxygen Generation Assembly OGS or OGA" Then
        
        Cells(i, "I").Value = mass_SPWE
        Cells(i, "L").Value = vol_SPWE
        Cells(i, "O").Value = cool_SPWE
        Cells(i, "R").Value = pow_SPWE
        Cells(i, "Y").Value = massre_SPWE
        Cells(i, "AA").Value = TRL_SPWE
    
    End If
    
    If Cells(i, 1).Value = "Static feed water electrolysis SFWE" Then
        
        Cells(i, "I").Value = mass_SFWE
        Cells(i, "L").Value = vol_SFWE
        Cells(i, "O").Value = cool_SFWE
        Cells(i, "R").Value = pow_SFWE
        Cells(i, "Y").Value = massre_SFWE
        Cells(i, "AA").Value = TRL_SFWE
    
    End If

Next i

Call Tabelle2.esmcalc
Call Tabelle2.calcsparesORUmass

End Sub
