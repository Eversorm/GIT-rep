VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ResModel 
   Caption         =   "Station Resupply Model"
   ClientHeight    =   5220
   ClientLeft      =   -630
   ClientTop       =   -2535
   ClientWidth     =   4110
   OleObjectBlob   =   "ResModel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ResModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_Cancel_Click()

    Unload ResModel

End Sub

Private Sub CommandButton_Proceed_Click()

'Worksheets("Resupply Approach").Protect UserInterfaceOnly:=True

    If ResModel.TextBox_NoM = "" Then
    
        MsgBox ("The number of missions need to be initialized before proceeding!")
        Exit Sub
        
    End If

    
    If ResModel.TextBox_NoM.Value > 1 Then Worksheets("Resupply Approach").Shapes("Instruction").Visible = True

'Create table

Dim i As Integer

Dim iInitialConsumablesColumn As Integer
Dim iECLSSHeaderRow As Integer
Dim iECLSSConsumablesColumn As Integer
Dim iResupplyConsumablesColumn As Integer
Dim iConsumable As Integer
Dim iSpareMassRow
Dim csInitialConsumables
Dim csResupplyConsumables
Dim miInitialConsumableRows
Dim miECLSSConsumableRows
Dim miResupplyConsumableRows
Dim iECLSSColumn

Dim LowerBound As Integer
Dim UpperBound As Integer
Dim MissionDuration As Double
Dim MissionCrewSize As Integer

Dim MassCrewMember As Double
Dim cb_Crew As Boolean

Dim iPreviousMissions As Integer
Dim iNewMissions As Integer
Dim iLabelColumn As Integer
Dim iMissionRow As Integer
Dim iSumRow As Integer
Dim iLastColumnPrevious As Integer
Dim M_diff As Integer
Dim cell
Dim rng

MassCrewMember = 82 '[kg]

With ThisWorkbook.Worksheets("Resupply Approach").Activate


'Application.ScreenUpdating = False

Worksheets("Resupply Approach").Protect UserInterfaceOnly:=True

On Error Resume Next

If Worksheets("Resupply Approach").CheckBoxes("Check_Box_Crew") = 1 Then
    cb_Crew = 1
Else: cb_Crew = 0
End If


' Column containing the labels of the initial consumables
iInitialConsumablesColumn = 14
iECLSSHeaderRow = 7
iECLSSConsumablesColumn = 1
iResupplyConsumablesColumn = 6

csInitialConsumables = Array("O2 res.", "H2 res.", "N2 res.", "H2O res.", "Food res.", "CO2 removal", "CO2 reduct.", "WW filt", "Urine Proc.", "Brine Proc.", "Electrolysis", "Nutrients res.", "PGC res.", "PBR res.", "CH4 proc. res.", "Clothes/misc.")
csResupplyConsumables = Array("O2 Supply", "H2 Supply", "N2 Supply", "H2O Supply", "Food", "CO2 Removal", "CO2 Reduction", "WW Filtration", "Urine Processing", "Brine Processing", "Electrolysis", "Nutrients Supply", "PGC Supply", "PBR Supply", "CH4 proc. Supply", "Clothes/ Spares", "Crew", "Other Payload")

miInitialConsumableRows = findRowByName(csInitialConsumables, "Output Data", iInitialConsumablesColumn)
miECLSSConsumableRows = findRowByName(csInitialConsumables, "Output Data", iECLSSConsumablesColumn)
miResupplyConsumableRows = findRowByName(csResupplyConsumables, "Resupply Approach", iResupplyConsumablesColumn)

iECLSSColumn = findColumnByName(Array(ResModel.ComboBoxECLSS.Value), "Output Data", iECLSSHeaderRow)
iECLSSColumn = iECLSSColumn(1)

iSpareMassRow = findRowByName(Array("Spare Mass:"), "Output Data", 1)
iSpareMassRow = iSpareMassRow(1)

iPreviousMissions = Range("B7").Value

iNewMissions = ResModel.TextBox_NoM.Value

iLabelColumn = 6
iMissionRow = 8
iSumRow = 31
iLastColumnPrevious = iLabelColumn + iPreviousMissions

'First Mission is identical to mission defined in MCA analysis
 
Dim CargoCapacity As Double
Dim fTotalConsumableMass As Double
Dim fMassToPreposition As Double
Dim fTotalCargoCapacityPrecursor As Double
Dim CrewMass As Double
Dim iPrecursorMissions As Integer
Dim iPrecursor As Integer
Dim AddSpareMass As Double

CargoCapacity = ResModel.TextBox_Mtc_Change.Value

Range("F9:F31").Locked = True

If ResModel.CheckBoxInitAllMission Then
    UpperBound = iNewMissions
Else
    UpperBound = 1
End If
        
MissionDuration = Worksheets("User Interface").Range("Mission_Duration")
MissionCrewSize = Worksheets("User Interface").Range("Crew_Size")

i = 1
Do While i <= UpperBound

    With Worksheets("Resupply Approach")
        
        
        fTotalConsumableMass = 0
        
        For iConsumable = 1 To UBound(miInitialConsumableRows)
            If csResupplyConsumables(iConsumable - 1) = "Clothes/ Spares" Then
                AddSpareMass = Worksheets("Output Data").Cells(iSpareMassRow, iECLSSColumn).Value
            Else
                AddSpareMass = 0
            End If
            fTotalConsumableMass = fTotalConsumableMass + AddSpareMass + Worksheets("Output Data").Cells(miECLSSConsumableRows(iConsumable), iECLSSColumn).Value * MissionDuration * MissionCrewSize / Worksheets("User Interface").Range("Crew_Size")
        Next iConsumable
        
        If cb_Crew Then
            CrewMass = MassCrewMember
        Else
            CrewMass = 0
        End If
        
        fMassToPreposition = 0
        
        If fTotalConsumableMass + (CrewMass * MissionCrewSize) > CargoCapacity Then
            ' in this case we require uncrewed precurser missions to preposition supplies
            fMassToPreposition = (fTotalConsumableMass - (CargoCapacity - CrewMass * MissionCrewSize))
            iPrecursorMissions = Application.WorksheetFunction.RoundUp(fMassToPreposition / CargoCapacity, 0)
            
            fTotalCargoCapacityPrecursor = iPrecursorMissions * CargoCapacity
            
            If fTotalCargoCapacityPrecursor > fTotalConsumableMass Then
                fMassToPreposition = fTotalConsumableMass
            Else
            ' if we have empty space on precursor mission, use it
                fMassToPreposition = fTotalCargoCapacityPrecursor
            End If
            
            iNewMissions = iNewMissions + iPrecursorMissions
            UpperBound = UpperBound + iPrecursorMissions
            
            i = i - 1
            
            For iPrecursor = 1 To iPrecursorMissions
                
                .Cells(iMissionRow + 1, iLabelColumn + i + iPrecursor).Value = 0
                .Cells(iMissionRow + 1, iLabelColumn + i + iPrecursor).Locked = True
                .Cells(iMissionRow + 2, iLabelColumn + i + iPrecursor).Value = 0
                .Cells(iMissionRow + 2, iLabelColumn + i + iPrecursor).Locked = True
                .Cells(iMissionRow + 3, iLabelColumn + i + iPrecursor).Value = CargoCapacity
                .Cells(iMissionRow + 3, iLabelColumn + i + iPrecursor).Locked = True
                
                For iConsumable = 1 To UBound(miInitialConsumableRows)
                
                    If csResupplyConsumables(iConsumable - 1) = "Clothes/ Spares" Then
                        AddSpareMass = Worksheets("Output Data").Cells(iSpareMassRow, iECLSSColumn).Value
                    Else
                        AddSpareMass = 0
                    End If
                    .Cells(miResupplyConsumableRows(iConsumable), iResupplyConsumablesColumn + i + iPrecursor).Value = ((fMassToPreposition / fTotalConsumableMass) / iPrecursorMissions) * (AddSpareMass + Worksheets("Output Data").Cells(miECLSSConsumableRows(iConsumable), iECLSSColumn).Value * MissionDuration * MissionCrewSize / Worksheets("User Interface").Range("Crew_Size"))
                    
                Next iConsumable
                
                .Cells(miResupplyConsumableRows(17), iResupplyConsumablesColumn + i + iPrecursor).Value = 0
                .Cells(miResupplyConsumableRows(18), iResupplyConsumablesColumn + i + iPrecursor).Value = 0
                
                .Cells(iSumRow, iResupplyConsumablesColumn + i + iPrecursor).Value = Application.WorksheetFunction.Sum(.Range(.Cells(miResupplyConsumableRows(1), iResupplyConsumablesColumn + i + iPrecursor), .Cells(miResupplyConsumableRows(15), iResupplyConsumablesColumn + i + iPrecursor)))
            Next iPrecursor
            'shift current mission index
            i = i + iPrecursorMissions + 1
        End If
        
        .Cells(iMissionRow + 1, iLabelColumn + i).Value = MissionDuration
        .Cells(iMissionRow + 1, iLabelColumn + i).Locked = True
        .Cells(iMissionRow + 2, iLabelColumn + i).Value = MissionCrewSize
        .Cells(iMissionRow + 2, iLabelColumn + i).Locked = True
        .Cells(iMissionRow + 3, iLabelColumn + i).Value = CargoCapacity
        .Cells(iMissionRow + 3, iLabelColumn + i).Locked = True
        
        For iConsumable = 1 To UBound(miInitialConsumableRows)
        
            If csResupplyConsumables(iConsumable - 1) = "Clothes/ Spares" Then
                AddSpareMass = Worksheets("Output Data").Cells(iSpareMassRow, iECLSSColumn).Value
            Else
                AddSpareMass = 0
            End If
            .Cells(miResupplyConsumableRows(iConsumable), iResupplyConsumablesColumn + i).Value = ((fTotalConsumableMass - fMassToPreposition) / fTotalConsumableMass) * (AddSpareMass + Worksheets("Output Data").Cells(miECLSSConsumableRows(iConsumable), iECLSSColumn).Value * MissionDuration * MissionCrewSize / Worksheets("User Interface").Range("Crew_Size"))
            
        Next iConsumable
        
        .Cells(miResupplyConsumableRows(17), iResupplyConsumablesColumn + i).Value = MissionCrewSize * CrewMass
        .Cells(miResupplyConsumableRows(18), iResupplyConsumablesColumn + i).Value = 0
        
        .Cells(iSumRow, iResupplyConsumablesColumn + i).Value = Application.WorksheetFunction.Sum(.Range(.Cells(miResupplyConsumableRows(1), iResupplyConsumablesColumn + i), .Cells(miResupplyConsumableRows(15), iResupplyConsumablesColumn + i)))
    End With
    
    i = i + 1
Loop


' Adjust table formatting:

' Set basic parameters
Range("B5").Value = ResModel.ComboBoxECLSS.Value
Range("A5").Font.Bold = True
Range("B5").VerticalAlignment = xlCenter
Range("B5").HorizontalAlignment = xlCenter

Range("B7").Value = iNewMissions
Range("A7").Font.Bold = True
Range("B7").VerticalAlignment = xlCenter
Range("B7").HorizontalAlignment = xlCenter

M_diff = iPreviousMissions - iNewMissions

If M_diff > 0 Then

    Set rng = Range(Cells(7, iLastColumnPrevious - M_diff + 1), Cells(31, iLastColumnPrevious))

    rng.Clear
    
End If


For i = 1 To iNewMissions

    Set rng = Range(Cells(iMissionRow, iLabelColumn + i), Cells(11, iLabelColumn + i))
    With rng
        .Interior.ColorIndex = 50
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Locked = False
    End With
    With Cells(8, iLabelColumn + i)
        .Value = "Mission" & " " & i
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.ColorIndex = 1
        .Locked = True
    End With
    With Cells(12, iLabelColumn + i)
        .Value = "Total amount" & vbNewLine & "[kg]"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.ColorIndex = 2
        .Interior.ColorIndex = 21
        .Locked = True
    End With
    Set rng = Range(Cells(iMissionRow, iLabelColumn + i), Cells(iSumRow, iLabelColumn + i))
    With rng
        .Borders.Weight = xlThin
        .BorderAround ColorIndex:=1, Weight:=xlThick
    End With
    Set rng = Range(Cells(13, iLabelColumn + i), Cells(iSumRow, iLabelColumn + i))
    With rng
        .Interior.ColorIndex = 8
        .NumberFormat = "0.00"
    End With
    Cells(iSumRow, iLabelColumn + i).Borders.Weight = xlThick
    
    Set rng = Range(Cells(iMissionRow, iLabelColumn + i), Cells(iSumRow, iLabelColumn + i))
    rng.Locked = True
    
Next i


Application.ScreenUpdating = True
End With
Unload ResModel

End Sub


Private Sub TextBox_Mtc_Change_Change()

    If IsNumeric(ResModel.TextBox_Mtc_Change.Value) = False Then
    
        MsgBox ("Please enter a valid mass transport capability!")
        ResModel.TextBox_Mtc_Change.Value = ""
        Exit Sub
        
    End If

    If ResModel.TextBox_Mtc_Change.Value = 0 Then
    
        MsgBox ("Please enter a valid mass transport capability!")
        ResModel.TextBox_Mtc_Change.Value = ""
        Exit Sub
        
    End If

End Sub

Private Sub TextBox_NoM_Change()

    If IsNumeric(ResModel.TextBox_NoM) = False Then
    
        MsgBox ("Please enter a valid amount of missions!")
        ResModel.TextBox_NoM.Value = ""
        Exit Sub
        
    End If

    If ResModel.TextBox_NoM.Value <= 0 Then
    
        MsgBox ("Minimal amount of missions is 1!")
        ResModel.TextBox_NoM.Value = ""
        Exit Sub
        
    End If

End Sub


Private Sub UserForm_Initialize()

Dim i As Integer

For i = 1 To 10

    If IsEmpty(Worksheets("Output Data").Cells(7, 2 + i).Value) Then
        Exit For
    End If
    ResModel.ComboBoxECLSS.AddItem Worksheets("Output Data").Cells(7, 2 + i).Value
    If i = 1 Then
        ResModel.ComboBoxECLSS.Value = Worksheets("Output Data").Cells(7, 2 + i).Value
    End If
Next i
End Sub

