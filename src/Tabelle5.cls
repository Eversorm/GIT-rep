VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub ExecuteMCA()
'This sub performs the MCA calculations after the initialization is complete

'Missiondefinition
Dim Crew_Size As Integer
Dim mission_duration As Integer
Dim Schedule_number As Integer
Dim ECLSS_type As String

'Multi-criteria Analysis
'Weighting Coefficients

Dim WeightsMin(1 To 5) As Double
Dim WeightsMax(1 To 2) As Double

Crew_Size = Worksheets("User Interface").Range("Crew_Size").Value
mission_duration = Worksheets("User Interface").Range("Mission_Duration").Value
Schedule_number = Worksheets("User Interface").Range("Schedule_Number").Value
ECLSS_type = Worksheets("User Interface").Range("ECLSS_Type").Value

'Weights
WeightsMin(1) = Worksheets("User Interface").Range("WC_Mass").Value
WeightsMin(2) = Worksheets("User Interface").Range("WC_Volume").Value
WeightsMin(3) = Worksheets("User Interface").Range("WC_Power").Value
WeightsMin(4) = Worksheets("User Interface").Range("WC_Cooling").Value
WeightsMin(5) = Worksheets("User Interface").Range("WC_Maintenance").Value

WeightsMax(1) = Worksheets("User Interface").Range("WC_Reliability").Value
WeightsMax(2) = Worksheets("User Interface").Range("WC_TRL").Value

Dim noe As Integer
Dim iHeaderRow As Integer
Dim iCriteria As Integer
Dim iTotalCriteria As Integer
Dim miColumnsMin
Dim miColumnsCFMin
Dim miColumnsMax
Dim miColumnsCFMax

Dim i As Integer
Dim iMCAResultColumn As Integer
Dim CurrentValue As Double
Dim MaxValue As Double
Dim MinValue As Double
Dim MeritFunction As Double
Dim ConfidenceFactor As Double
Dim CurrentScore As Double

noe = Worksheets("MCA ESM").Cells(Rows.Count, 1).End(xlUp).row

iHeaderRow = 4
miColumnsMin = findColumnByName(Array("MCA Value"), "MCA ESM", iHeaderRow)
iMCAResultColumn = miColumnsMin(1)
' These are criteria where lower values are desirable
miColumnsMin = findColumnByName(Array("Mass", "Volume", "Power", "Cooling", "Maintenance"), "MCA ESM", iHeaderRow)
miColumnsCFMin = findColumnByName(Array("Mass CF", "Volume CF", "Power CF", "Cooling CF", "Maintenance CF", "Reliability CF", "TRL CF"), "MCA ESM", iHeaderRow)
' These are criteria where higher values are desirable
miColumnsMax = findColumnByName(Array("Reliability", "TRL"), "MCA ESM", iHeaderRow)
miColumnsCFMax = findColumnByName(Array("Reliability CF", "TRL CF"), "MCA ESM", iHeaderRow)

' The total number of criteria
iTotalCriteria = UBound(miColumnsMin) + UBound(miColumnsMax)

With ThisWorkbook.Worksheets("MCA ESM")
For i = 6 To noe

    ' For a description of the MCA calculations see chapter 4.7.1 from Andreas Feigls Master Thesis (https://mediatum.ub.tum.de/doc/1503234/1503234.pdf)
    MeritFunction = 0
    For iCriteria = 1 To UBound(miColumnsMin)
        CurrentValue = .Cells(i, miColumnsMin(iCriteria))
        MaxValue = Application.WorksheetFunction.Max(.Range(.Cells(6, miColumnsMin(iCriteria)), .Cells(noe, miColumnsMin(iCriteria))))
        MinValue = Application.WorksheetFunction.Min(.Range(.Cells(6, miColumnsMin(iCriteria)), .Cells(noe, miColumnsMin(iCriteria))))
        ConfidenceFactor = .Cells(i, miColumnsCFMin(iCriteria))
        If MaxValue = MinValue Then
        ' all alternatives have equal score
            CurrentScore = 1
        Else
            CurrentScore = ((CurrentValue ^ 2 - 2 * CurrentValue * MaxValue + MaxValue ^ 2) / (MaxValue - MinValue) ^ 2)
        End If
        
        MeritFunction = MeritFunction + WeightsMin(iCriteria) * CurrentScore * ConfidenceFactor
    Next iCriteria
    
    For iCriteria = 1 To UBound(miColumnsMax)
        CurrentValue = .Cells(i, miColumnsMax(iCriteria))
        MaxValue = Application.WorksheetFunction.Max(.Range(.Cells(6, miColumnsMax(iCriteria)), .Cells(noe, miColumnsMax(iCriteria))))
        MinValue = Application.WorksheetFunction.Min(.Range(.Cells(6, miColumnsMax(iCriteria)), .Cells(noe, miColumnsMax(iCriteria))))
        ConfidenceFactor = .Cells(i, miColumnsCFMax(iCriteria))
        If MaxValue = MinValue Then
        ' all alternatives have equal score
            CurrentScore = 1
        Else
            CurrentScore = ((CurrentValue ^ 2 - CurrentValue * MinValue) / (MaxValue ^ 2 - MaxValue * MinValue))
        End If
        
        MeritFunction = MeritFunction + WeightsMax(iCriteria) * CurrentScore * ConfidenceFactor
    Next iCriteria
    
    ' The resulting MCA value is calculated by dividing it with the total number of criteria, see Equation 4.13 from Andreas Feigls Master Thesis
    .Cells(i, iMCAResultColumn) = MeritFunction / iTotalCriteria
    
Next i
End With

Tabelle5.CalculateMCAandESMRanks

End Sub

Sub CalculateMCAandESMRanks()
With ThisWorkbook.Worksheets("MCA ESM")
' Now calculate the MCA ESM and ESM without crew time ranks
Dim i As Integer
Dim noe As Integer
Dim iHeaderRow As Integer
Dim miRankColumns
Dim iMCAResultColumn As Integer
Dim iMCARankColumn As Integer
Dim iESMNCT_Column As Integer
Dim iESMNCTRank_Column As Integer
Dim iESM_Column As Integer
Dim iESMRank_Column As Integer

iHeaderRow = 4
miRankColumns = findColumnByName(Array("MCA Value"), "MCA ESM", iHeaderRow)
iMCAResultColumn = miRankColumns(1)

miRankColumns = findColumnByName(Array("MCA Rank", "ESM Without Crew Time", "ESM w/o CT Rank", "ESM", "ESM Rank"), "MCA ESM", iHeaderRow)
iMCARankColumn = miRankColumns(1)
iESMNCT_Column = miRankColumns(2)
iESMNCTRank_Column = miRankColumns(3)
iESM_Column = miRankColumns(4)
iESMRank_Column = miRankColumns(5)

noe = Worksheets("MCA ESM").Cells(Rows.Count, 1).End(xlUp).row

For i = 6 To noe
    .Cells(i, iMCARankColumn).Value = Application.WorksheetFunction.Rank(.Cells(i, iMCAResultColumn).Value, .Range(.Cells(6, iMCAResultColumn), .Cells(noe, iMCAResultColumn)))
    .Cells(i, iESMNCTRank_Column).Value = Application.WorksheetFunction.Rank(.Cells(i, iESMNCT_Column).Value, .Range(.Cells(6, iESMNCT_Column), .Cells(noe, iESMNCT_Column)), 1)
    .Cells(i, iESMRank_Column).Value = Application.WorksheetFunction.Rank(.Cells(i, iESM_Column).Value, .Range(.Cells(6, iESM_Column), .Cells(noe, iESM_Column)), 1)
Next i

End With
End Sub


Sub calculateAttributeCriticality()
' checks by how much each attribute must change before the first place of the MCA and ESM analysis changes place

Dim iSign As Integer
Dim iColumnsToAdjust As Integer
Dim csColumnsToAdjust
Dim miColumns
Dim AdjustmentValue As Double
Dim AdjustmentValuePrevious As Double
Dim TotalAdjustmentValue As Double
Dim iRankColumn As Integer
Dim iNameColumn As Integer
Dim iMCAColumn As Integer
Dim iColumn As Integer
Dim iRow As Integer
Dim iRow2 As Integer
Dim iIteration As Integer
Dim iRowToAdjust As Integer
Dim iMaxIterations As Integer
Dim na As Integer
Dim iHeaderRow As Integer
Dim iAnalysis As Integer
Dim sCurrentAnalysis As String
Dim miSign
Dim mlPercentToChange(0 To 1) As Long
Dim csAlternatives
Dim csAnalysis
Dim miRank() As Integer
Dim bValidAlternative As Boolean
Dim iSecondRankRow As Integer

' These variables refer to the "Output Data" Worksheet!
Dim iOutputColumn As Integer
Dim iOutputLabelColumn As Integer
Dim miOutputRows

iOutputLabelColumn = 14
iHeaderRow = 4
iMaxIterations = 300

' get the column containing the system names
miColumns = findColumnByName(Array("Name of Assembly"), "MCA ESM", iHeaderRow)
iNameColumn = miColumns(1)

miColumns = findColumnByName(Array("MCA Value"), "MCA ESM", iHeaderRow)
iMCAColumn = miColumns(1)

' adjust values in 10% steps
AdjustmentValue = 10 / 100

' check whether we should increase or decrease the values in the sensitivity analysis
miSign = Array(1, -1)
csAnalysis = Array("MCA", "ESM")
    
With ThisWorkbook.Worksheets("MCA ESM")
na = .Cells(Rows.Count, 1).End(xlUp).row

If na < iHeaderRow + 3 Then
    ' Cannot calculate criticalities with only one alternative
    Exit Sub
End If

'disable screen updates
Application.ScreenUpdating = False

ReDim csAlternatives(iHeaderRow + 2 To na + 10)
For iRow = iHeaderRow + 2 To na + 10
    csAlternatives(iRow) = .Cells(iRow, iNameColumn).Value
Next iRow

For iAnalysis = 0 To 1
    ' get the rank column we want to compare:
    If csAnalysis(iAnalysis) = "MCA" Then
        miColumns = findColumnByName(Array("MCA Rank"), "MCA ESM", iHeaderRow)
        iRankColumn = miColumns(1)
        csColumnsToAdjust = Array("Mass", "Volume", "Power", "Cooling", "Maintenance", "Reliability", "TRL")
        miColumns = findColumnByName(csColumnsToAdjust, "MCA ESM", iHeaderRow)
        miOutputRows = findRowByName(csColumnsToAdjust, "Output Data", iOutputLabelColumn)
        iOutputColumn = iOutputLabelColumn + 1
        
    ElseIf csAnalysis(iAnalysis) = "ESM" Then
        miColumns = findColumnByName(Array("ESM Rank"), "MCA ESM", iHeaderRow)
        iRankColumn = miColumns(1)
        csColumnsToAdjust = Array("Mass", "Volume", "Power", "Cooling", "Maintenance")
        miColumns = findColumnByName(csColumnsToAdjust, "MCA ESM", iHeaderRow)
        miOutputRows = findRowByName(Array("Mass", "Volume", "Power", "Cooling", "Maintenance", "Reliability", "TRL"), "Output Data", iOutputLabelColumn)
        iOutputColumn = iOutputLabelColumn + 2
        
    End If
    ' clear values
    For iRow = 1 To UBound(miOutputRows)
        Worksheets("Output Data").Cells(miOutputRows(iRow), iOutputColumn).Value = ""
    Next iRow
    
    ' check which system should be adjusted, in this case the system which is currently on the first place
    For iRow = iHeaderRow + 2 To na
        If .Cells(iRow, iRankColumn).Value = 1 Then
            iRowToAdjust = iRow
        ElseIf .Cells(iRow, iRankColumn).Value = 2 Then
            iSecondRankRow = iRow
        End If
    Next iRow
    
    ' check the current ranks before the iteration
    ReDim miRank(iHeaderRow + 2 To na)
    For iRow = iHeaderRow + 2 To na
        miRank(iRow) = .Cells(iRow, iRankColumn).Value
    Next iRow
    ' Now adjust the values, for each property we individually check when the rank change occurs
    For iColumn = 1 To UBound(miColumns)
        mlPercentToChange(0) = 0
        mlPercentToChange(1) = 0
        ' and we also check in which direction the rank change occurs
        For iSign = 0 To 1
            iIteration = 0
            TotalAdjustmentValue = 0
            ' Flag to start an iteration , code will jump back here if iteration has not resulted in rank change yet
Iteration:
            iIteration = iIteration + 1
            For iRow = iHeaderRow + 2 To na
                If .Cells(iRow, iRankColumn).Value = 2 Then
                    iSecondRankRow = iRow
                    Exit For
                End If
            Next iRow
            ' If distance between current first and second rank is large, use large adjustment steps
            AdjustmentValue = (.Cells(iRowToAdjust, iMCAColumn).Value - .Cells(iSecondRankRow, iMCAColumn).Value) * 2
            
            If AdjustmentValue < 0.1 Then
                AdjustmentValue = 0.1
            End If
            
            ' Now adjust the values
            .Cells(iRowToAdjust, miColumns(iColumn)).Value = .Cells(iRowToAdjust, miColumns(iColumn)).Value * (1 + miSign(iSign) * AdjustmentValue)
            TotalAdjustmentValue = TotalAdjustmentValue + AdjustmentValue
            
            ' Limit maximum value for TRL
            If csColumnsToAdjust(iColumn - 1) = "TRL" Then
                If .Cells(iRowToAdjust, miColumns(iColumn)).Value > 9 Then
                    .Cells(iRowToAdjust, miColumns(iColumn)).Value = 9
                End If
            End If
        
            If csAnalysis(iAnalysis) = "MCA" Then
                ' for the sensitivity analysis of the MCA, we adjust the values in the MCA ESM sheet and then call the Tabelle5.ExecuteMCA macro for the MCA calculation!
                 Tabelle5.ExecuteMCA
            ElseIf csAnalysis(iAnalysis) = "ESM" Then
                ' for the ESM sensitvity analysis, we adjust the values in the MCA sheet, and then call the ESM sub but specify that we want to perform the ESM in the MCA sheet
                Modul1.esmcalc "MCA ESM"
            End If
            
            If AdjustmentValuePrevious < (.Cells(iRowToAdjust, iMCAColumn).Value - .Cells(iSecondRankRow, iMCAColumn).Value) * 2 Then
                ' if the change in the MCA difference becomes larger, we are moving in the wrong direction and can skip this iteration
                iIteration = iMaxIterations + 1
            End If
            AdjustmentValuePrevious = (.Cells(iRowToAdjust, iMCAColumn).Value - .Cells(iSecondRankRow, iMCAColumn).Value) * 2
            ' check the current ranks, in case the first place changed, we finish the iteration by going to the EndIteration flag
            If miRank(iRowToAdjust) = .Cells(iRowToAdjust, iRankColumn).Value And iIteration < iMaxIterations + 1 Then
                ' If the first rank has not changed, we go to the iteration flag and adjust the values again
                GoTo Iteration
            Else
            ' note that the iColumn index in the miOutputRows variable here is correct, because the column identifies the properties which are adjusted
                If iIteration > iMaxIterations Then
                    mlPercentToChange(iSign) = 0 ' 0 means no rank change occured at all
                Else
                    mlPercentToChange(iSign) = TotalAdjustmentValue * 100
                End If
            End If
            
            ' reset the values to the default ones for the next property
            Tabelle1.PerformMCA
            For iRow = iHeaderRow + 2 To na + 10
                bValidAlternative = False
                For iRow2 = iHeaderRow + 2 To na + 10
                    If csAlternatives(iRow2) = .Cells(iRow, iNameColumn).Value Then
                        bValidAlternative = True
                    End If
                Next iRow2
                If bValidAlternative = False Then
                     .Cells(iRow, iNameColumn).EntireRow.Delete
                End If
            Next iRow
            Tabelle5.CalculateMCAandESMRanks
            'Set rows height
            Rows("5:20").RowHeight = 28.8
            
            If mlPercentToChange(1) > mlPercentToChange(0) Then
                Exit For
            End If
        Next iSign
    
        ' Write the lower of the values (increase or decrease) into the output data sheet with the corresponding sign
        If mlPercentToChange(0) < mlPercentToChange(1) Then
            Worksheets("Output Data").Cells(miOutputRows(iColumn), iOutputColumn).Value = mlPercentToChange(1)
        Else
            Worksheets("Output Data").Cells(miOutputRows(iColumn), iOutputColumn).Value = mlPercentToChange(0)
        End If
    
    Next iColumn
Next iAnalysis

End With

Application.ScreenUpdating = True

End Sub

Sub sortByESM_NCT()
 sortBy "ESM w/o CT Rank"
End Sub
Sub sortByMCA()
 sortBy "MCA Rank"
End Sub
Sub sortByESM()
 sortBy "ESM Rank"
End Sub

Sub sortBy(sInput As String)

Dim na As Integer
Dim miColumns
Dim iColumn As Integer
Dim lastRow As Long

miColumns = findColumnByName(Array(sInput), "MCA ESM", 4)
iColumn = miColumns(1)

With ThisWorkbook.Worksheets("MCA ESM")
Worksheets("MCA ESM").Protect UserInterfaceOnly:=True

lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
If lastRow < 6 Then lastRow = 6
na = .Range("A6:A" & lastRow).Cells.SpecialCells(xlCellTypeConstants).Count     ' count the number of tech

If na > 1 Then
    
    .Range("A6:Z" & lastRow).Sort Key1:=.Cells(6, iColumn), Order1:=xlAscending, Header:=xlNo
    Rows.EntireRow.Hidden = False
    Rows("5:30").RowHeight = 28.8

End If
End With

End Sub


Sub exclalt()

    Worksheets("MCA ESM").Protect UserInterfaceOnly:=True
    ExcludeAlternative.Width = 500
    ExcludeAlternative.Height = 200
    ExcludeAlternative.Show
    
End Sub
