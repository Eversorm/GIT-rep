Attribute VB_Name = "Constants"
Option Explicit

' "Constants"
' Modul: we need some constants in different parts of the code so we just calculate them for once, also useful to update
' BVAD values without searching them everywhere

' Stoichiometric values for the needed compounds

Public Const MolarMassH2 As Double = 2.01588 / 1000
Public Const MolarMassO2 As Double = 31.9988 / 1000
Public Const MolarMassH2O As Double = 18.0153 / 1000
Public Const MolarMassC As Double = 12.0107 / 1000
Public Const MolarMassCO2 As Double = 44.0095 / 1000
Public Const MolarMassCH4 As Double = 16.0425 / 1000
Public Const MolarMassC6H12O6 As Double = 180.1559 / 1000
Public Const MolarMassLiOH As Double = 23.948 / 1000
Public Const MolarMassLi2CO3 As Double = 73.891 / 1000
Public Const MolarMassC2H2 As Double = 26.038 / 1000


' NASA Baseline Values and Assumptions Document (BVAD) - Metabolic Constants
' All values are expressed in [kg/CM*d] or [kJ/CM*d] unless specified otherwise

Public Const BVAD_O2_Con As Double = 0.816             ' O2 Consumption in [kg/CM*d]
Public Const BVAD_CO2_Pro As Double = 1.04             ' CO2 Production [kg/CM*d]
Public Const BVAD_urine As Double = 1.6                ' urine Production in [kg/CM*d]
Public Const BVAD_potable As Double = 2.5              ' potable water Need in [kg/CM*d]
Public Const BVAD_hygiene As Double = 0.7              ' hygiene water Need in [kg/CM*d]
Public Const BVAD_feces As Double = 0.132              ' feces Production in [kg/CM*d]
Public Const BVAD_avg_swe_vapr As Double = 1.9         ' average sweat and vpr Production in [kg/CM*d]
Public Const BVAD_peak_swe_vapr As Double = 0.77       ' peak sweat and vpr Production in [kg/CM*h]
Public Const BVAD_heat As Double = 12000#              ' heat Production in [kJ/CM*d]
Public Const BVAD_peak_heat As Double = 2974#          ' peak heat Production in [kJ/CM*h]
Public Const BVAD_food As Double = 1.51                ' Reference food mass in [kg/CM*d]
