VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AdjustSupplyForm 
   Caption         =   "Adjust Supply"
   ClientHeight    =   2955
   ClientLeft      =   -90
   ClientTop       =   -300
   ClientWidth     =   3210
   OleObjectBlob   =   "AdjustSupplyForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AdjustSupplyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonCancel_Click()

    Unload AdjustSupplyForm
    
End Sub

Private Sub CommandButtonProceed_Click()

Dim iResupplyConsumablesColumn As Integer
Dim csResupplyConsumables
Dim miResupplyConsumableRows

Dim LowerBound As Integer
Dim UpperBound As Integer
Dim MissionDuration As Double
Dim MissionCrewSize As Integer

Dim iNewMissions As Integer
Dim iLabelColumn As Integer
Dim iMissionRow As Integer
Dim iSumRow As Integer

Dim i As Integer
Dim iMissionHeaderRow As Integer
Dim miRows
Dim iSourceMissionIndex As Integer
Dim iTargetMissionIndex As Integer
Dim sTemp As String
Dim fPercentageToMove As Double
Dim fMassToMove As Double

Application.ScreenUpdating = False

With ThisWorkbook.Worksheets("Resupply Approach")

iMissionHeaderRow = 8
iLabelColumn = 6
iSumRow = 30

' Initialize variables:
fPercentageToMove = AdjustSupplyForm.TextBoxPercentage.Value
If fPercentageToMove = 0 Then
    Unload AdjustSupplyForm
    Exit Sub
End If

If AdjustSupplyForm.ComboBoxConsumable.Value = "Everything" Then
    csResupplyConsumables = Array("O2 Supply", "H2 Supply", "N2 Supply", "H2O Supply", "Food", "CO2 Removal", "CO2 Reduction", "WW Filtration", "Urine Processing", "Brine Processing", "Electrolysis", "Nutrients Supply", "PGC Supply", "PBR Supply", "CH4 proc. Supply", "Clothes/ Spares")
Else
    csResupplyConsumables = Array(AdjustSupplyForm.ComboBoxConsumable.Value)
End If

miResupplyRows = findRowByName(csResupplyConsumables, "Resupply Approach", iLabelColumn)
    
miRows = findColumnByName(Array(AdjustSupplyForm.ComboBoxFromMission.Value), "Resupply Approach", iMissionHeaderRow)
iSourceMissionIndex = miRows(1)

miRows = findColumnByName(Array(AdjustSupplyForm.ComboBoxToMission.Value), "Resupply Approach", iMissionHeaderRow)
iTargetMissionIndex = miRows(1)


For i = 1 To UBound(miResupplyRows)
    fMassToMove = fPercentageToMove / 100 * .Cells(miResupplyRows(i), iSourceMissionIndex).Value
    .Cells(miResupplyRows(i), iSourceMissionIndex).Value = .Cells(miResupplyRows(i), iSourceMissionIndex).Value - fMassToMove
    .Cells(miResupplyRows(i), iTargetMissionIndex).Value = .Cells(miResupplyRows(i), iTargetMissionIndex).Value + fMassToMove
    
    .Cells(iSumRow, iSourceMissionIndex).Value = .Cells(iSumRow, iSourceMissionIndex).Value - fMassToMove
    .Cells(iSumRow, iTargetMissionIndex).Value = .Cells(iSumRow, iTargetMissionIndex).Value + fMassToMove
Next i


End With

Application.ScreenUpdating = True
Unload AdjustSupplyForm
End Sub



Private Sub TextBoxPercentage_Change()

    If AdjustSupplyForm.TextBoxPercentage.Value <> "" Then
        If Not IsNumeric(AdjustSupplyForm.TextBoxPercentage.Value) Or AdjustSupplyForm.TextBoxPercentage.Value < 0 Or AdjustSupplyForm.TextBoxPercentage.Value > 100 Then
            MsgBox ("Please enter a valid percentage value between 0 and 100!")
            AdjustSupplyForm.TextBoxPercentage.Value = 0
        End If
    End If
End Sub

Private Sub UserForm_Initialize()

Dim iMission As Integer
Dim iTotalMissions As Integer
Dim iConsumable As Integer
Dim iResupplyConsumablesColumn As Integer
Dim csResupplyConsumables
Dim miResupplyConsumableRows

iResupplyConsumablesColumn = 6

csResupplyConsumables = Array("O2 Supply", "H2 Supply", "N2 Supply", "H2O Supply", "Food", "CO2 Removal", "CO2 Reduction", "WW Filtration", "Urine Processing", "Brine Processing", "Electrolysis", "Nutrients Supply", "PGC Supply", "PBR Supply", "CH4 proc. Supply", "Clothes/ Spares", "Crew", "Other Payload")

miResupplyConsumableRows = findRowByName(csResupplyConsumables, "Resupply Approach", iResupplyConsumablesColumn)


AdjustSupplyForm.ComboBoxConsumable.Value = "Everything"
AdjustSupplyForm.ComboBoxConsumable.AddItem "Everything"

AdjustSupplyForm.TextBoxPercentage.Value = "0"

With ThisWorkbook.Worksheets("Resupply Approach")
For iConsumable = 1 To UBound(miResupplyConsumableRows)
    AdjustSupplyForm.ComboBoxConsumable.AddItem .Cells(miResupplyConsumableRows(iConsumable), iResupplyConsumablesColumn).Value
Next iConsumable

iTotalMissions = .Range("B7").Value

For iMission = 1 To iTotalMissions
    AdjustSupplyForm.ComboBoxFromMission.AddItem .Cells(8, iResupplyConsumablesColumn + iMission).Value
    AdjustSupplyForm.ComboBoxToMission.AddItem .Cells(8, iResupplyConsumablesColumn + iMission).Value
Next iMission

End With
End Sub

