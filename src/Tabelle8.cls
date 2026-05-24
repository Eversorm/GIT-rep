VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub showinstruction()

    Shapes("Instruction").Visible = True

End Sub

Sub hideinstruction()

    Shapes("Instruction").Visible = False

End Sub

Sub matrixcheck()




'________________________________________________________________________________________

Dim lastcolumn
Dim lastRow
Dim cell
Dim row
Dim answer As Long
Dim i As Integer
Dim z As Integer
Dim rng As Range
Dim na
Dim n
Dim smbf 'number of missions served by flight

Application.ScreenUpdating = False

row = 33
lastcolumn = Cells(33, Columns.Count).End(xlToLeft).Column
lastRow = Cells(Rows.Count, 1).End(xlUp).row
na = Range("B7").Value
n = lastRow - 31

If Range("C31").Interior.ColorIndex = 2 Then
    MsgBox ("Nothing to check!")
    Exit Sub
End If

    For Each cell In Range(Cells(row, lastcolumn), Cells(lastRow - 1, lastcolumn))
    
        If cell.Value = "" Then
            MsgBox ("Transportation capability for Flight " & cell.row - 30 & " must be inserted before proceeding")
            Exit Sub
        End If
    
        If cell.Interior.ColorIndex <> 4 Then
            MsgBox ("Current cargo mass for Flight " & cell.row - 30 & " is exceeding the tranportation capability")
            Exit Sub
        End If
    
        If cell.Value > 0 Then
            answer = MsgBox("You still have " & cell.Value & " kg cargo tranport capability for Flight " & cell.row - 30 & ". Do you still want to proceed?", vbYesNo + vbQuestion, "Resupply Matrix Check")
            If answer = vbNo Then Exit Sub
        End If
    
    Next

    For Each cell In Range(Cells(lastRow, 3), Cells(lastRow, lastcolumn - 1))
    
        If cell.Interior.ColorIndex <> 4 Then
            MsgBox ("There is mass remaining for Mission ") & cell.Column - 2
            Exit Sub
        End If
    
        If cell.Value < 0 Then
            answer = MsgBox("You carry " & -cell.Value & " kg too much cargo for Mission " & cell.Column - 2 & ". Do you still want to proceed?", vbYesNo + vbQuestion, "Resupply Matrix Check")
            If answer = vbNo Then Exit Sub
        End If
    
    Next

answer = MsgBox("Do you want to plot the resupply dependancy matrix?" & vbNewLine & "Note: If you continue, entries in the mass matrix can no longer be changed.", vbYesNo + vbQuestion, "Resupply Matrix Check")
        If answer = vbNo Then Exit Sub

'Delete previous dependancy matrix

Set rng = Range(Cells(lastRow + 2, 3), Cells(2 * lastRow - 29, na + 4))

    For Each cell In rng
        cell.Clear
        cell.Interior.ColorIndex = 2
    Next
    

'Prevent changes in mass matrix after mass matrix verification

Worksheets("Resupply Approach").Protect UserInterfaceOnly:=True

Set rng = Range(Cells(row, 3), Cells(lastRow - 1, lastcolumn - 1))
    rng.Locked = True

'Plot dependancy matrix

For i = 1 To na

    With Cells(lastRow + 2, 2 + i)
        .Value = "Mission" & " " & i
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.ColorIndex = 1
        .Interior.ColorIndex = 50
        .Locked = True
        .Borders.ColorIndex = 1
        .Borders.Weight = xlThin
    End With

    With Cells(lastRow + 2 + i, 2)
        .Value = "Flight" & " " & i
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.ColorIndex = 1
        .Locked = True
        .Interior.ColorIndex = 50
        .Borders.ColorIndex = 1
        .Borders.Weight = xlThin
    End With

Next i

If n > na Then

    For i = 1 To n - na

         With Cells(lastRow + 2 + na + i, 2)
            .Value = "Flight" & " " & na + i
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.ColorIndex = 1
            .Locked = True
            .Interior.ColorIndex = 50
            .Borders.ColorIndex = 1
            .Borders.Weight = xlThin
        End With
    
    Next

End If
Set rng = Range(Cells(lastRow + 3, na + 3), Cells(2 * lastRow - 29, na + 6))

With rng
        .Borders.ColorIndex = 1
        .Borders.Weight = xlThin
        .Interior.ColorIndex = 50
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Locked = True
End With

Set rng = Range(Cells(lastRow + 3, na + 3), Cells(2 * lastRow - 29, na + 6))

With rng
        .Borders.ColorIndex = 1
        .Borders.Weight = xlThin
        .Interior.ColorIndex = 50
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "0.00"
        .Locked = True
End With

With Cells(lastRow + 2, na + 4)
        .Value = "Dependency" & vbNewLine & "Factor (Sum)"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.ColorIndex = 50
        .Locked = True
        .Borders.ColorIndex = 1
        .Borders.Weight = xlThin
End With

With Cells(lastRow + 2, na + 3)
        .Value = "Number of" & vbNewLine & "Mis. served"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.ColorIndex = 50
        .Locked = True
        .Borders.ColorIndex = 1
        .Borders.Weight = xlThin
End With

Set rng = Range(Cells(lastRow + 3, 3), Cells(2 * lastRow - 29, na + 2))

With rng
        .Borders.ColorIndex = 1
        .Borders.Weight = xlThin
        .Interior.ColorIndex = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "0.00"
End With

With Cells(lastRow + 2, na + 5)
        .Value = "Flight Criti-" & vbNewLine & "cality Index"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.ColorIndex = 50
        .Locked = True
        .Borders.ColorIndex = 1
        .Borders.Weight = xlThin
End With

With Cells(lastRow + 2, na + 6)
        .Value = "Flight Criti-" & vbNewLine & "cality Rank"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.ColorIndex = 50
        .Locked = True
        .Borders.ColorIndex = 1
        .Borders.Weight = xlThin
End With

'Enter dependancy values

For i = 1 To na

    Cells(lastRow + 2 + i, 3).Value = Cells(30 + i, 3).Value / Application.WorksheetFunction.RoundUp(Cells(24, 6 + i).Value, 1)
    'If Cells(lastrow + 2 + i, 3).Value = 0 Then Cells(lastrow + 2 + i, 3).Value = ""
        For z = 1 To na
        
            Cells(lastRow + 2 + i, 2 + z).Value = Cells(30 + i, 2 + z).Value / Application.WorksheetFunction.RoundUp(Cells(24, 6 + z).Value, 1)
            'If Cells(lastrow + 2 + i, 2 + z).Value = 0 Then Cells(lastrow + 2 + i, 2 + z).Value = ""
        Next z
        
Next i

If n > na Then

    For i = 1 To n - na
    
        Cells(lastRow + na + 2 + i, 3).Value = Cells(30 + na + i, 3).Value / Application.WorksheetFunction.RoundUp(Cells(24, 6 + i).Value, 1)
        'If Cells(lastrow + na + 2 + i, 3).Value = 0 Then Cells(lastrow + na + 2 + i, 3).Value = ""
        For z = 1 To na
        
            Cells(lastRow + na + 2 + i, 2 + z).Value = Cells(30 + na + i, 2 + z).Value / Application.WorksheetFunction.RoundUp(Cells(24, 6 + z).Value, 1)
            'If Cells(lastrow + na + 2 + i, 2 + z).Value = 0 Then Cells(lastrow + na + 2 + i, 2 + z).Value = ""
        Next z
        
    Next i
    
End If

For i = 1 To n
    
    smbf = 0
    For Each cell In Range(Cells(lastRow + 2 + i, 3), Cells(lastRow + 2 + i, na + 2))
        If cell.Value <> 0 Then smbf = smbf + 1
    Next cell
    
    Cells(lastRow + 2 + i, na + 3).Value = smbf
    Cells(lastRow + 2 + i, na + 3).NumberFormat = "0"
    Cells(lastRow + 2 + i, na + 4).Value = Application.WorksheetFunction.Sum(Range(Cells(lastRow + 2 + i, 3), Cells(lastRow + 2 + i, na + 2)))
    Cells(lastRow + 2 + i, na + 5).Value = ((Cells(lastRow + 2 + i, na + 3).Value) ^ 2 + (Cells(lastRow + 2 + i, na + 4).Value) ^ 2) ^ (0.5)
    
Next i

Set rng = Range(Cells(lastRow + 3, na + 5), Cells(2 * lastRow - 29, na + 5))

For i = 1 To n

    Cells(lastRow + 2 + i, na + 6).Value = Application.Rank(Cells(lastRow + 2 + i, na + 5), rng)
    Cells(lastRow + 2 + i, na + 6).NumberFormat = "0"
    
Next i

Application.ScreenUpdating = True

End Sub

Sub performresupplyanalysis()
    
    ResModel.Width = 350
    ResModel.Height = 250
    ResModel.Show

End Sub


Sub adjustSupply()
    
    AdjustSupplyForm.Width = 350
    AdjustSupplyForm.Height = 250
    AdjustSupplyForm.Show

End Sub

Sub plotResupplyMatrix()

Dim i As Integer
Dim iResupplyHeaderRow As Integer
Dim iNewMissions As Integer
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
Dim iSumRow As Integer
Dim iHeaderRow As Integer
Dim iLabelColumn As Integer

Application.ScreenUpdating = False

iNewMissions = Range("B7").Value

' Column containing the labels of the initial consumables
iInitialConsumablesColumn = 14
iECLSSHeaderRow = 7
iECLSSConsumablesColumn = 1
iResupplyConsumablesColumn = 6

iResupplyHeaderRow = 12
iHeaderRow = 33
iLabelColumn = 2
iSumRow = 31

csInitialConsumables = Array("O2 res.", "H2 res.", "N2 res.", "H2O res.", "Food res.", "CO2 removal", "CO2 reduct.", "WW filt", "Urine Proc.", "Brine Proc.", "Electrolysis", "Nutrients res.", "PGC res.", "PBR res.", "CH4 proc. res.", "Clothes/misc.")
csResupplyConsumables = Array("O2 Supply", "H2 Supply", "N2 Supply", "H2O Supply", "Food", "CO2 Removal", "CO2 Reduction", "WW Filtration", "Urine Processing", "Brine Processing", "Electrolysis", "Nutrients Supply", "PGC Supply", "PBR Supply", "CH4 proc. Supply", "Clothes/ Spares", "Crew", "Other Payload")

miInitialConsumableRows = findRowByName(csInitialConsumables, "Output Data", iInitialConsumablesColumn)
miECLSSConsumableRows = findRowByName(csInitialConsumables, "Output Data", iECLSSConsumablesColumn)
miResupplyConsumableRows = findRowByName(csResupplyConsumables, "Resupply Approach", iResupplyConsumablesColumn)

iECLSSColumn = findColumnByName(Array(Worksheets("Resupply Approach").Range("B5")), "Output Data", iECLSSHeaderRow)
iECLSSColumn = iECLSSColumn(1)

iSpareMassRow = findRowByName(Array("Spare Mass:"), "Output Data", 1)
iSpareMassRow = iSpareMassRow(1)

Dim iCurrentFlight As Integer
Dim iServicedMission As Integer
Dim iMission As Integer
Dim mfDuration
Dim mfCrew
Dim AddSpareMass As Double
Dim fPreviousTotalTransportedMass As Double
Dim fCurrentTotalMissionNeed As Double
Dim fMatrixEntry As Double
Dim fCurrentMissionRemainingTransportMass As Double

With Worksheets("Resupply Approach")

mfCrew = .Range(.Cells(10, iResupplyConsumablesColumn + 1), .Cells(10, iResupplyConsumablesColumn + iNewMissions)).Value
mfDuration = .Range(.Cells(9, iResupplyConsumablesColumn + 1), .Cells(9, iResupplyConsumablesColumn + iNewMissions)).Value

Dim mfRequiredMass() As Double
ReDim mfRequiredMass(1 To iNewMissions, 1 To UBound(miInitialConsumableRows) + 1)
Dim mfTransportedMass() As Double
ReDim mfTransportedMass(1 To iNewMissions, 1 To UBound(miInitialConsumableRows) + 1)

Dim mfTotalRequiredMass() As Double
ReDim mfTotalRequiredMass(1 To iNewMissions)
Dim mfTotalTransportedMass() As Double
ReDim mfTotalTransportedMass(1 To iNewMissions)

Dim fTotalTransportedMissionMass As Double
Dim fDesiredMass As Double

Dim rng

Set rng = .Range(.Cells(iHeaderRow + 1, iLabelColumn + 1), .Cells(iHeaderRow + iNewMissions, iLabelColumn + iNewMissions))
rng.ClearContents
    
For iMission = 1 To iNewMissions
    ' get the total consumable mass for each mission
    mfTotalRequiredMass(iMission) = 0
    For iConsumable = 1 To UBound(miInitialConsumableRows)
        If csResupplyConsumables(iConsumable - 1) = "Clothes/ Spares" And mfCrew(1, iMission) > 0 Then
            AddSpareMass = Worksheets("Output Data").Cells(iSpareMassRow, iECLSSColumn).Value
        Else
            AddSpareMass = 0
        End If
        mfRequiredMass(iMission, iConsumable) = AddSpareMass + Worksheets("Output Data").Cells(miECLSSConsumableRows(iConsumable), iECLSSColumn).Value * mfDuration(1, iMission) * mfCrew(1, iMission) / Worksheets("User Interface").Range("Crew_Size")
        
        mfTotalRequiredMass(iMission) = mfTotalRequiredMass(iMission) + mfRequiredMass(iMission, iConsumable)
        
        mfTransportedMass(iMission, iConsumable) = .Cells(iResupplyHeaderRow + iConsumable, iResupplyConsumablesColumn + iMission).Value
    Next iConsumable
                                                                                                             ' other payload
    mfRequiredMass(iMission, iConsumable) = mfCrew(1, iMission) * 82 + .Cells(29, iResupplyConsumablesColumn + iMission)
    mfTotalRequiredMass(iMission) = mfTotalRequiredMass(iMission) + mfRequiredMass(iMission, iConsumable)
    
    mfTransportedMass(iMission, iConsumable) = .Cells(iResupplyHeaderRow + iConsumable, iResupplyConsumablesColumn + iMission).Value + .Cells(iResupplyHeaderRow + iConsumable + 1, iResupplyConsumablesColumn + iMission).Value
    
    mfTotalTransportedMass(iMission) = .Cells(iSumRow, iResupplyConsumablesColumn + iMission).Value
Next iMission

For iCurrentFlight = 1 To iNewMissions
    ' The total amount that was already transported in previous missions to this one
    fPreviousTotalTransportedMass = 0
        
    For i = 1 To iCurrentFlight - 1
        For iConsumable = 1 To UBound(miInitialConsumableRows) + 1
            fPreviousTotalTransportedMass = fPreviousTotalTransportedMass + mfTransportedMass(i, iConsumable)
        Next iConsumable
    Next i
    
    fCurrentMissionRemainingTransportMass = mfTotalTransportedMass(iCurrentFlight)
    
    For iServicedMission = 1 To iNewMissions
        fCurrentTotalMissionNeed = 0
        
        ' the total required mass of all missions up to the current considered serviced mission
        fTotalTransportedMissionMass = 0
        
        For iConsumable = 1 To UBound(miInitialConsumableRows) + 1
            ' Check if the mission still required mass of this type
            fDesiredMass = mfRequiredMass(iServicedMission, iConsumable)
            If fDesiredMass > 0 Then
                ' Check if the flight still has the desired mass
                If mfTransportedMass(iCurrentFlight, iConsumable) > 0 Then
                    If mfTransportedMass(iCurrentFlight, iConsumable) > fDesiredMass Then
                        fTotalTransportedMissionMass = fTotalTransportedMissionMass + fDesiredMass
                        mfTransportedMass(iCurrentFlight, iConsumable) = mfTransportedMass(iCurrentFlight, iConsumable) - fDesiredMass
                        mfRequiredMass(iServicedMission, iConsumable) = 0
                    Else
                        fTotalTransportedMissionMass = fTotalTransportedMissionMass + mfTransportedMass(iCurrentFlight, iConsumable)
                        mfRequiredMass(iServicedMission, iConsumable) = mfRequiredMass(iServicedMission, iConsumable) - mfTransportedMass(iCurrentFlight, iConsumable)
                        mfTransportedMass(iCurrentFlight, iConsumable) = 0
                    End If
                End If
            End If
        Next iConsumable
        
        fTotalTransportedMissionMass = Round(fTotalTransportedMissionMass, 10)
        
        If fTotalTransportedMissionMass > 0 Then
        ' Now we can check if the serviced mission need has already been supplied, if that is the case we can use the consumables for a future mission
        ' or if the  serviced mission need must be covered by the current mission, only in this case a new entry will be added to the matrix
        
            .Cells(iHeaderRow + iCurrentFlight, iLabelColumn + iServicedMission).Value = fTotalTransportedMissionMass
            fCurrentMissionRemainingTransportMass = fCurrentMissionRemainingTransportMass - fTotalTransportedMissionMass
        End If
        
    Next iServicedMission
    
    .Cells(iHeaderRow + iCurrentFlight, iLabelColumn + iNewMissions + 1).Value = Round(fCurrentMissionRemainingTransportMass + (.Cells(11, 6 + iCurrentFlight) - mfTotalTransportedMass(iCurrentFlight)), 9)
Next iCurrentFlight

' delete formatting of unused columns:

Dim iPreviousColumns As Integer
Dim M_diff As Integer
Dim iColumn As Integer
Dim iRow As Integer

iPreviousColumns = .Cells(iHeaderRow, Columns.Count).End(xlToLeft).Column

M_diff = iPreviousColumns - (iNewMissions + iLabelColumn + 1)

If M_diff > 0 Then

    Set rng = .Range(.Cells(iHeaderRow, iPreviousColumns - M_diff + 1), .Cells(iHeaderRow + iPreviousColumns, iPreviousColumns))
    rng.Clear
    
    Set rng = .Range(.Cells(iHeaderRow + iPreviousColumns + 1, iLabelColumn), .Cells(iHeaderRow + iPreviousColumns, iPreviousColumns))
    rng.Clear
End If


' set formatting for columns
For iColumn = 1 To iNewMissions
    With .Cells(iHeaderRow, iLabelColumn + iColumn)
        .Value = "Mission" & " " & iColumn
        .Interior.ColorIndex = 50
        
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.ColorIndex = 1
        .Locked = True
    End With
    
    fMatrixEntry = WorksheetFunction.Sum(.Range(.Cells(iHeaderRow + 1, iLabelColumn + iColumn), .Cells(iHeaderRow + iNewMissions, iLabelColumn + iColumn)))
    With .Cells(iHeaderRow + iNewMissions + 1, iLabelColumn + iColumn)
        ' set sum of matrix as value
        .Value = Round(mfTotalRequiredMass(iColumn) - fMatrixEntry, 9)
        .NumberFormat = "0.00"
        If .Value = 0 Then
            .Interior.ColorIndex = 4
        Else
            .Interior.ColorIndex = 3
        End If
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.ColorIndex = 1
        .Borders.Weight = xlThin
        .Locked = True
    End With
Next iColumn

With Cells(iHeaderRow, iLabelColumn + iNewMissions + 1)
    .Value = "Cargo mass Left [kg]"
    .Interior.ColorIndex = 2
    
    .Borders.Weight = xlThin
    .WrapText = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.Bold = True
    .Font.ColorIndex = 1
    .Locked = True
End With


' set formatting for rows
For iRow = 1 To iNewMissions

    With .Cells(iHeaderRow + iRow, iLabelColumn)
        .Value = "Flight" & " " & iRow
        .Interior.ColorIndex = 50
        
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.ColorIndex = 1
        .Locked = True
    End With
    
    ' Set formatting of matrix content
    Set rng = Range(Cells(iHeaderRow + iRow, iLabelColumn + 1), Cells(iHeaderRow + iRow, iLabelColumn + iNewMissions))
    With rng
        .Borders.Weight = xlThin
        .Interior.ColorIndex = 8
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Locked = True
    End With
        
    With Cells(iHeaderRow + iRow, iLabelColumn + iNewMissions + 1)
        .NumberFormat = "0.00"
        If .Value > 0 Then
            .Interior.ColorIndex = 4
        ElseIf .Value = 0 Then
            .Interior.ColorIndex = 2
        Else
            .Interior.ColorIndex = 3
        End If
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.ColorIndex = 1
        .Borders.Weight = xlThin
        .Locked = True
    End With
Next iRow

With Cells(iHeaderRow + iNewMissions + 1, iLabelColumn)
    .Value = "Cargo mass Left [kg]"
    .Interior.ColorIndex = 2
    
    .Borders.Weight = xlThin
    .WrapText = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.Bold = True
    .Font.ColorIndex = 1
    .Locked = True
End With

Worksheets("Resupply Approach").Shapes("Label_Resupply_Matrix").Visible = True

' now plot the flight criticalities
plotFlightCriticalityMatrix

End With
Application.ScreenUpdating = True
End Sub

Sub plotFlightCriticalityMatrix()

Dim iHeaderRow As Integer
Dim iHeaderRowFlightCritMatrix As Integer
Dim iLabelColumn As Integer
Dim iNumberOfMissions As Integer
Dim iFlight As Integer
Dim iMission As Integer
Dim fTotalMissionMass As Double
Dim iTotalMissionsServed As Integer
Dim fCriticalitySum As Double
Dim fFlightCriticality As Double

With Worksheets("Resupply Approach")
iLabelColumn = 2
iNumberOfMissions = .Range("B7").Value
iHeaderRow = 33

iHeaderRowFlightCritMatrix = iHeaderRow + iNumberOfMissions + 3

' Set formatting for flight criticality matrix:

' delete formatting of unused columns:
Dim iLastRow As Integer
Dim iLastColumn
Dim rng

iLastRow = .Cells.Find(What:="*", _
                After:=Range("B1"), _
                LookAt:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).row

iLastColumn = .Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
Worksheets("Resupply Approach").Unprotect

ThisWorkbook.Worksheets("Resupply Approach").ChartObjects("FCP").Delete

Set rng = .Range(.Cells(iHeaderRowFlightCritMatrix - 1, iLabelColumn), .Cells(iLastRow, iLastColumn))
rng.Clear
        
Dim iColumn As Integer
Dim iRow As Integer

' set formatting for columns
iHeaderRow = iHeaderRowFlightCritMatrix
For iColumn = 1 To iNumberOfMissions
    With .Cells(iHeaderRow, iLabelColumn + iColumn)
        .Value = "Mission" & " " & iColumn
        .Interior.ColorIndex = 50
        
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.ColorIndex = 1
        .Locked = True
    End With
Next iColumn

With Cells(iHeaderRow, iLabelColumn + iNumberOfMissions + 1)
    .Interior.ColorIndex = 2
    
    .Borders.Weight = xlThin
    .WrapText = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.Bold = True
    .Font.ColorIndex = 1
    .Locked = True
End With

With Cells(iHeaderRow, iLabelColumn + iNumberOfMissions + 1)
    .Value = "Sum of Flight Criticalities"
    .Interior.ColorIndex = 2
    
    .Borders.Weight = xlThin
    .WrapText = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.Bold = True
    .Font.ColorIndex = 1
    .Locked = True
End With
With Cells(iHeaderRow, iLabelColumn + iNumberOfMissions + 2)
    .Value = "Number of Missions"
    .Interior.ColorIndex = 2
    
    .Borders.Weight = xlThin
    .WrapText = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.Bold = True
    .Font.ColorIndex = 1
    .Locked = True
End With
With Cells(iHeaderRow, iLabelColumn + iNumberOfMissions + 3)
    .Value = "FCI [-]"
    .Interior.ColorIndex = 2
    
    .Borders.Weight = xlThin
    .WrapText = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.Bold = True
    .Font.ColorIndex = 1
    .Locked = True
End With

' set formatting for rows
For iRow = 1 To iNumberOfMissions

    With .Cells(iHeaderRow + iRow, iLabelColumn)
        .Value = "Flight" & " " & iRow
        .Interior.ColorIndex = 50
        
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.ColorIndex = 1
        .Locked = True
    End With
    
    ' Set formatting of matrix content
    Set rng = Range(Cells(iHeaderRow + iRow, iLabelColumn + 1), Cells(iHeaderRow + iRow, iLabelColumn + iNumberOfMissions))
    With rng
        .Borders.Weight = xlThin
        .Interior.ColorIndex = 8
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Locked = True
    End With
        
    With Cells(iHeaderRow + iRow, iLabelColumn + iNumberOfMissions + 1)
        .NumberFormat = "0.00"
        .Interior.ColorIndex = 2
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.ColorIndex = 1
        .Borders.Weight = xlThin
        .Locked = True
    End With
    With Cells(iHeaderRow + iRow, iLabelColumn + iNumberOfMissions + 2)
        .NumberFormat = "0.00"
        .Interior.ColorIndex = 2
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.ColorIndex = 1
        .Borders.Weight = xlThin
        .Locked = True
    End With
    With Cells(iHeaderRow + iRow, iLabelColumn + iNumberOfMissions + 3)
        .NumberFormat = "0.00"
        .Interior.ColorIndex = 2
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.ColorIndex = 1
        .Borders.Weight = xlThin
        .Locked = True
    End With
Next iRow

' Fill matrix with values:

iHeaderRow = 33

' Flight criticality is calculate as mass transported for mission j by flight i compared to total required mass for mission j see equation 4.43 from Andreas Feigls MA
For iFlight = 1 To iNumberOfMissions
    For iMission = 1 To iNumberOfMissions
    
        fTotalMissionMass = WorksheetFunction.Sum(.Range(.Cells(iHeaderRow + 1, iLabelColumn + iMission), .Cells(iHeaderRow + iNumberOfMissions, iLabelColumn + iMission)))
        If fTotalMissionMass > 0 Then
            fFlightCriticality = .Cells(iHeaderRow + iFlight, iLabelColumn + iMission).Value / fTotalMissionMass
            fFlightCriticality = Round(fFlightCriticality, 10)
            If fFlightCriticality > 0 Then
                .Cells(iHeaderRowFlightCritMatrix + iFlight, iLabelColumn + iMission).Value = fFlightCriticality
            End If
        End If
    Next iMission
    
    ' To better compare flights, a scalar flight criticality index for each flight is calculated see equation 4.48 from Andreas Feigls MA
    iTotalMissionsServed = WorksheetFunction.CountIf(.Range(.Cells(iHeaderRowFlightCritMatrix + iFlight, iLabelColumn + 1), .Cells(iHeaderRowFlightCritMatrix + iFlight, iLabelColumn + iMission)), "> 0")
    fCriticalitySum = WorksheetFunction.Sum(.Range(.Cells(iHeaderRowFlightCritMatrix + iFlight, iLabelColumn + 1), .Cells(iHeaderRowFlightCritMatrix + iFlight, iLabelColumn + iMission)))
    
    .Cells(iHeaderRowFlightCritMatrix + iFlight, iLabelColumn + iNumberOfMissions + 1).Value = fCriticalitySum
    .Cells(iHeaderRowFlightCritMatrix + iFlight, iLabelColumn + iNumberOfMissions + 2).Value = iTotalMissionsServed
    .Cells(iHeaderRowFlightCritMatrix + iFlight, iLabelColumn + iNumberOfMissions + 3).Value = (fCriticalitySum ^ 2 + iTotalMissionsServed ^ 2) ^ (1 / 2)
    
Next iFlight

End With

plotflightcriticalitypath

End Sub

Sub plotflightcriticalitypath()

Dim iLabelColumn As Integer
Dim iNumberOfMissions As Integer
Dim iHeaderRow As Integer
Dim iHeaderRowFlightCritMatrix As Integer

Dim iLastRow As Integer
Dim iLastColumn As Integer
Dim rngSumOfCriticalities
Dim rngMissionsServed
Dim rngFlightCriticalityIndex

Dim maxCriticalitySum As Double
Dim maxMissionsServed As Double
Dim frame
Dim Dia_FCP

With ThisWorkbook.Worksheets("Resupply Approach")
iLabelColumn = 2
iNumberOfMissions = .Range("B7").Value
iHeaderRow = 33

iHeaderRowFlightCritMatrix = iHeaderRow + iNumberOfMissions + 3

iLastRow = iHeaderRowFlightCritMatrix + iNumberOfMissions

iLastColumn = iLabelColumn + iNumberOfMissions + 3

If Cells(iHeaderRowFlightCritMatrix, iLastColumn) = "" Then

    MsgBox ("Nothing to plot! Adjust mass resupply matrix before creating chart!")
    Exit Sub

End If

Set rngSumOfCriticalities = Range(Cells(iHeaderRowFlightCritMatrix + 1, iLastColumn - 2), Cells(iLastRow, iLastColumn - 2))
Set rngMissionsServed = Range(Cells(iHeaderRowFlightCritMatrix + 1, iLastColumn - 1), Cells(iLastRow, iLastColumn - 1))
Set rngFlightCriticalityIndex = Range(Cells(iHeaderRowFlightCritMatrix + 1, iLastColumn), Cells(iLastRow, iLastColumn))

    maxCriticalitySum = WorksheetFunction.Max(rngSumOfCriticalities)
    maxMissionsServed = WorksheetFunction.Max(rngMissionsServed)

    Set frame = Worksheets("Resupply Approach").ChartObjects.Add(900, 900, 400, 300)
    Set Dia_FCP = frame.Chart

    Worksheets("Resupply Approach").ChartObjects(1).Top = Cells(iHeaderRowFlightCritMatrix, iLastColumn + 2).Top
    Worksheets("Resupply Approach").ChartObjects(1).Left = Cells(iHeaderRowFlightCritMatrix, iLastColumn + 2).Left
    
    With Dia_FCP
    
        .Parent.Name = "FCP"
        .PlotArea.Border.ColorIndex = 14
        .SetSourceData rngSumOfCriticalities
        .ChartType = xlXYScatter
        .HasLegend = False
        .SeriesCollection(1).XValues = rngMissionsServed
        .SeriesCollection(1).Values = rngSumOfCriticalities
        .SeriesCollection(1).HasDataLabels = True
        .Axes(xlCategory).MinimumScale = 0
        .Axes(xlCategory).MaximumScale = maxMissionsServed + 1
        .Axes(xlCategory).MajorUnit = 1
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Number of Missions Served by Flight"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Sum of Dependency Factors per Flight"
        .Axes(xlValue).MaximumScale = maxCriticalitySum + 1
        .Axes(xlValue).TickLabels.NumberFormat = "0.0"
        .HasTitle = True
        '.Axes(Excel.XlAxisType.xlValue).MajorGridlines.Delete
        .ChartTitle.Text = "Flight Criticality Path for Campaign"

    End With

Dim i As Integer
Dim X() As Integer
ReDim X(0 To maxMissionsServed + 1)
Dim Y() As Integer
ReDim Y(0 To maxMissionsServed + 1)

For i = 0 To maxMissionsServed + 1
    X(i) = i
    Y(i) = i
Next

        With Dia_FCP.SeriesCollection.NewSeries
            .XValues = X
            .Values = Y
            .Trendlines.Add
            .Trendlines(1).Border.Weight = xlThin
            .Trendlines(1).Border.LineStyle = xlDash
            .MarkerStyle = xlMarkerStyleNone
        End With

    For i = 1 To iNumberOfMissions
        Dia_FCP.SeriesCollection(1).Points(i).DataLabel.Text = Cells(iHeaderRowFlightCritMatrix + i, iLabelColumn).Text
    Next i
    
    
End With
End Sub
