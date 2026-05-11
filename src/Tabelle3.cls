VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Excel.Range)

Worksheets("Schedule and Metabolic").Unprotect

If Target.Column > 1 And Target.Column < 9 Then
    If Target.row > 59 And Target.row < 110 Then
        ' If something was changed in the schedule set both the check and the calculate marks to 0
        Worksheets("Schedule and Metabolic").Range("Schedule_was_Checked").Value = 0
        Worksheets("Schedule and Metabolic").Range("Metabolic_Values_up_to_date").Value = 0
    End If
    
    If Target.row > 110 Then
        ' If something was changed in metabolics values, set the calculate mark to 0
        Worksheets("Schedule and Metabolic").Range("Metabolic_Values_up_to_date").Value = 0
    End If
End If
End Sub

Sub updateDefaultSchedule()
' This sub can be used to update the default schedule, in case someone changed the Metabolic values or the default schedule.
' Since this is something the user is not supposed to do, there is no button to trigger it but it is triggered if the crew size changes (TO DO)

Application.ScreenUpdating = False

Dim ESM_Crew_TimeF As Double
Dim iHeaderRow As Integer
Dim iHeaderRowSchedule As Integer
Dim sNameColumn As String
Dim MetabolicHeaders
Dim miMetabolicColumn
Dim MetabolicRows
Dim miMetabolicRows
Dim ScheduleHeaders
Dim miScheduleColumns

Dim pos As Integer
Dim posExercise As Integer
Dim posPreExercise As Integer
Dim posPostExercise As Integer
Dim posPotableWater As Integer
Dim posSweat As Integer
Dim posWork As Integer
Dim iMetabolic As Integer
Dim iExerciseTicks As Integer
Dim iBacksteps As Integer
Dim i As Integer
Dim iPostExerciseTicks As Integer

' Define the row in which the metabolic value headers are located
iHeaderRow = 1
' Define the column in which the different schedule events are located
sNameColumn = "N"
' Define the row in which the headers of the default schedule are located
iHeaderRowSchedule = 2

MetabolicHeaders = Array("O2 Consumption", "CO2 Production", "Urine Water Production", "Feces Production", "Potable Water Consumption", "Hygiene Water Consumption", "Food Consumption", "Sweat/Water Vapor Production", "Heat Load")
miMetabolicColumn = findColumnByName(MetabolicHeaders, "Schedule and Metabolic", iHeaderRow)

MetabolicRows = Array("Sleep", "Pre-Sleep", "Post Sleep", "Personal Hygiene", "Work/Planning", "Recreation", "Pre Exercise", "Exercise", "Post Exercise", "Breakfast", "Lunch", "Dinner")
miMetabolicRows = findRowByName(MetabolicRows, "Schedule and Metabolic", sNameColumn)

' Urine production is considered fix for one day and not actually matched to an event like post exercise of which multiple can exist
' Metabolic Water and Water in Food are just required in internal calculation to derive potable water demand

ScheduleHeaders = Array("task", "O2", "CO2", "urine", "feces", "potable", "hygiene", "food", "sweat", "heat")
miScheduleColumns = findColumnByName(ScheduleHeaders, "Schedule and Metabolic", iHeaderRowSchedule)

posExercise = Application.Match("Exercise", MetabolicRows, False)
posPreExercise = Application.Match("Pre Exercise", MetabolicRows, False)
posPostExercise = Application.Match("Post Exercise", MetabolicRows, False)
posPotableWater = Application.Match("Potable Water Consumption", MetabolicHeaders, False)
posSweat = Application.Match("Sweat/Water Vapor Production", MetabolicHeaders, False)
posWork = Application.Match("Work/Planning", MetabolicRows, False)

iPostExerciseTicks = 0

' note for this to work the order of the metabolic values after "task" in the ScheduleHeaders array must be identical to the order of values in the MetabolicHeaders array
' Also the schedule time steps must be 0.5 hours, otherwise the exercise and pre exercise calculation must be fixed
For i = 3 To 50
    If Cells(i, miScheduleColumns(1)).Value = "Exercise" Then
        ' check whether this is the first exercise entry, if so we have to mix the metabolic values with the pre exercise values, otherwise we can use the metabolic values for exercise
        If Cells(i - 1, miScheduleColumns(1)).Value = "Exercise" Then
            For iMetabolic = 1 To UBound(miMetabolicColumn)
                Cells(i, miScheduleColumns(iMetabolic + 1)).Value = Cells(miMetabolicRows(posExercise), miMetabolicColumn(iMetabolic)).Value
            Next iMetabolic
        Else
            For iMetabolic = 1 To UBound(miMetabolicColumn)
                Cells(i, miScheduleColumns(iMetabolic + 1)).Value = (Cells(miMetabolicRows(posExercise), miMetabolicColumn(iMetabolic)).Value + Cells(miMetabolicRows(posPreExercise), miMetabolicColumn(iMetabolic)).Value) / 2
            Next iMetabolic
        End If
        
    Else
        If Cells(i, miScheduleColumns(1)).Value = "Post Exercise" Then
        ' For post exercise, the metabolic values contain the amount of water that was consumed during the total exercise period. However, the value in the metabolic values table is only correct for half an hour of exercise.
        ' Therefore we check here how long the exercise before this block was and sum up the total water demand from that exercise
            iPostExerciseTicks = iPostExerciseTicks + 1
            
            iExerciseTicks = 0
            For iBacksteps = 1 To i
                If Cells(i - iBacksteps, miScheduleColumns(1)).Value = "Exercise" Then
                    iExerciseTicks = iExerciseTicks + 1
                ElseIf Cells(i - iBacksteps, miScheduleColumns(1)).Value = "Post Exercise" Then
                    ' in this case do not exit for loop but also do not increase exercise counter
                Else
                    Exit For
                End If
                    
            Next iBacksteps
            
            ' for only one exercise tick (meaning half an hour) the values in the metabolic table are correct
            pos = Application.Match(Cells(i, miScheduleColumns(1)).Value, MetabolicRows, False)
            For iMetabolic = 1 To UBound(miMetabolicColumn)
                Cells(i, miScheduleColumns(iMetabolic + 1)).Value = Cells(miMetabolicRows(pos), miMetabolicColumn(iMetabolic)).Value
            Next iMetabolic
            If iExerciseTicks > 1 Then
            ' if more exercise is done we have to recalculate the potable water consumption, but the other metabolic values are correct. We always have 15 minutes pre exercise and the remaining exericse plus the sweat from post exercise.
            ' Also we have to consider the corresponding nominal sweat rate, for which water consumption is already planned
                Cells(i, miScheduleColumns(posPotableWater + 1)).Value = Cells(miMetabolicRows(posPreExercise), miMetabolicColumn(posSweat)).Value * 0.25 + Cells(miMetabolicRows(posExercise), miMetabolicColumn(posSweat)).Value * ((iExerciseTicks / 2) - 0.25) + Cells(miMetabolicRows(posPostExercise), miMetabolicColumn(posSweat)).Value - Cells(miMetabolicRows(posWork), miMetabolicColumn(posSweat)).Value * ((iExerciseTicks / 2) + 1)
            End If
            
        Else
            pos = Application.Match(Cells(i, miScheduleColumns(1)).Value, MetabolicRows, False)
            
            For iMetabolic = 1 To UBound(miMetabolicColumn)
                Cells(i, miScheduleColumns(iMetabolic + 1)).Value = Cells(miMetabolicRows(pos), miMetabolicColumn(iMetabolic)).Value
            Next iMetabolic
        End If
    End If

Next i


' Now adjust the schedule to the crew size (also done automatically if the user changes the crew size)
    UpdateScheduleToNumberOfCrew

' Now that we have the total heat output, we can also scale the total food consumption:
Dim TotalFoodRequired As Double
Dim AdditionalWaterToRemove As Double
Dim TotalAdditionalWaterInFood As Double
Dim TotalAdditionalWaterToFeces As Double
Dim TotalAdditionalMetabolicWater As Double
Dim FoodScalingFactor As Double
Dim Crew_Size As Integer

Crew_Size = Worksheets("User Interface").Range("Crew_Size").Value

TotalFoodRequired = (Worksheets("Schedule and Metabolic").Range("SUM_Heat_Load") / 12000) * 1.51

' increase urine, feces and food consumption by food scaling factor
FoodScalingFactor = TotalFoodRequired / (Crew_Size * 1.51)

TotalAdditionalWaterInFood = (FoodScalingFactor * Worksheets("Schedule and Metabolic").Range("SUM_Food_Water") - Worksheets("Schedule and Metabolic").Range("SUM_Food_Water")) * Crew_Size
TotalAdditionalWaterToFeces = (FoodScalingFactor * 0.1 - 0.1) * Crew_Size
TotalAdditionalMetabolicWater = (FoodScalingFactor * Worksheets("Schedule and Metabolic").Range("SUM_Metabolic_Water") - Worksheets("Schedule and Metabolic").Range("SUM_Metabolic_Water")) * Crew_Size

AdditionalWaterToRemove = (TotalAdditionalWaterInFood + TotalAdditionalMetabolicWater - TotalAdditionalWaterToFeces)

For i = 3 To 50
    Cells(i, miScheduleColumns(5)).Value = Cells(i, miScheduleColumns(5)).Value * FoodScalingFactor
    Cells(i, miScheduleColumns(8)).Value = Cells(i, miScheduleColumns(8)).Value * FoodScalingFactor
    If Cells(i, miScheduleColumns(1)).Value = "Post Exercise" Then
        Cells(i, miScheduleColumns(6)).Value = Cells(i, miScheduleColumns(6)).Value - AdditionalWaterToRemove / (iPostExerciseTicks / 2) ' /2 because each ticks is 0.5 hours!
    End If
Next i

Application.ScreenUpdating = True

Worksheets("Schedule and Metabolic").Protect UserInterfaceOnly:=True

End Sub

Sub UpdateScheduleToNumberOfCrew()

Dim Crew_Size
Dim prev_crew_size
Dim rngData As Range
Dim String1 As String
Dim String2 As String
Dim EvaluationString As String

    Set rngData = ThisWorkbook.Worksheets("Schedule and Metabolic").Range("D3:L50")

    Crew_Size = Worksheets("User Interface").Range("Crew_Size").Value
    prev_crew_size = Cells(3, "D").Value / Cells(3, "P").Value

    ' update default schedule to crew size
    'Values = Range("D3:L50")
    String1 = "* "
    String2 = " / "
    EvaluationString = String1 & CStr(Crew_Size) & String2 & CStr(prev_crew_size)
    rngData = Evaluate(rngData.Address & EvaluationString)

'Worksheets("Schedule and Metabolic").Protect UserInterfaceOnly:=True

End Sub


Sub ComposeUserSchedule()
    
    ComposeScheduleForm.Width = 500
    ComposeScheduleForm.Height = 550
    ComposeScheduleForm.Show
    
End Sub

Sub CheckSchedule()
' this sub first checks whether all required elements of a schedule are present and then calculates the metabolic values for each time:

Worksheets("Schedule and Metabolic").Unprotect

Dim iScheduleHeaderRow As Integer
Dim iAdditionalHeaderRows As Integer
Dim iScheduleRows As Integer
Dim iRequiredSleep As Integer
Dim iRequiredPostSleep As Integer
Dim iRequiredPreSleep As Integer
Dim iRequiredHygiene As Integer
Dim iRequiredBreakfast As Integer
Dim iRequiredLunch As Integer
Dim iRequiredDinner As Integer
Dim iRequiredPostExercise As Integer
Dim lastCol As Integer
Dim iCrewInSchedule As Integer

Dim iSchedule As Integer
Dim iCell As Integer

Dim CurrentCell
Dim ErrorString As String
Dim colRange
Dim Crew_Size As Integer

iScheduleHeaderRow = 59
iAdditionalHeaderRows = 2
iScheduleRows = 48

iCrewInSchedule = 0
Crew_Size = Worksheets("User Interface").Range("Crew_Size").Value

With ThisWorkbook.Worksheets("Schedule and Metabolic")

Set colRange = Cells(iScheduleHeaderRow, Columns.Count).End(xlToLeft)
lastCol = colRange.Column

For iSchedule = 1 To lastCol - 1
    If .Cells(iScheduleHeaderRow + 1, iSchedule + 1) > 0 Then ' only check schedules with assigned crew members
    iRequiredSleep = 14 ' 7 hours minimum
    iRequiredPostSleep = 1
    iRequiredPreSleep = 2
    iRequiredHygiene = 1
    iRequiredBreakfast = 1
    iRequiredLunch = 2
    iRequiredDinner = 2
    iRequiredPreSleep = 2
    
    For iCell = 1 To iScheduleRows
        CurrentCell = .Cells(iScheduleHeaderRow + iAdditionalHeaderRows + iCell, iSchedule + 1)
        If CurrentCell = "Sleep" Then
        iRequiredSleep = iRequiredSleep - 1
        ElseIf CurrentCell = "Post Sleep" Then
        iRequiredPostSleep = iRequiredPostSleep - 1
        ElseIf CurrentCell = "Hygiene" Then
        iRequiredHygiene = iRequiredHygiene - 1
        ElseIf CurrentCell = "Breakfast" Then
        iRequiredBreakfast = iRequiredBreakfast - 1
        ElseIf CurrentCell = "Lunch" Then
        iRequiredLunch = iRequiredLunch - 1
        ElseIf CurrentCell = "Dinner" Then
        iRequiredDinner = iRequiredDinner - 1
        ElseIf CurrentCell = "Pre-Sleep" Then
        iRequiredPreSleep = iRequiredPreSleep - 1
        ' Check if exercise and post exercise are used correctly
        ElseIf CurrentCell = "Exercise" Then
            If .Cells(iScheduleHeaderRow + iAdditionalHeaderRows + iCell + 1, iSchedule + 1).Value <> "Exercise" Then
                iRequiredPostExercise = 2
                If .Cells(iScheduleHeaderRow + iAdditionalHeaderRows + iCell + 1, iSchedule + 1).Value = "Post Exercise" Then
                    iRequiredPostExercise = iRequiredPostExercise - 1
                End If
                If .Cells(iScheduleHeaderRow + iAdditionalHeaderRows + iCell + 2, iSchedule + 1).Value = "Post Exercise" Then
                    iRequiredPostExercise = iRequiredPostExercise - 1
                End If
                If .Cells(iScheduleHeaderRow + iAdditionalHeaderRows + iCell + 3, iSchedule + 1).Value = "Post Exercise" Then
                    iRequiredPostExercise = iRequiredPostExercise - 1
                End If
                
                If iRequiredPostExercise > 0 Then
                    ErrorString = "In user defined schedule " & CStr(iSchedule) & " Post Exercise is missing, please adjust the schedule!"
                    MsgBox (ErrorString)
                    Exit Sub
                ElseIf iRequiredPostExercise < 0 Then
                    ErrorString = "In user defined schedule " & CStr(iSchedule) & " too much Post Exercise is used, please adjust the schedule!"
                    MsgBox (ErrorString)
                    Exit Sub
                End If
            End If
            
        ElseIf CurrentCell = "Post Exercise" Then
            If iCell = 1 Then
                If .Cells(iScheduleHeaderRow + iAdditionalHeaderRows + iScheduleRows, iSchedule + 1).Value <> "Exercise" And .Cells(iScheduleHeaderRow + iAdditionalHeaderRows + iCell - 1, iSchedule + 1).Value <> "Post Exercise" Then
                    ErrorString = "In user defined schedule " & CStr(iSchedule) & " Post Exercise is used without a preceeding Exercise, please adjust the schedule!"
                    MsgBox (ErrorString)
                    Exit Sub
                End If
            Else
                If .Cells(iScheduleHeaderRow + iAdditionalHeaderRows + iCell - 1, iSchedule + 1).Value <> "Exercise" And .Cells(iScheduleHeaderRow + iAdditionalHeaderRows + iCell - 1, iSchedule + 1).Value <> "Post Exercise" Then
                    ErrorString = "In user defined schedule " & CStr(iSchedule) & " Post Exercise is used without a preceeding Exercise, please adjust the schedule!"
                    MsgBox (ErrorString)
                    Exit Sub
                End If
            End If
        End If
    Next iCell
    
    ' Check overall schedule requirements
    If iRequiredSleep > 0 Then
        ErrorString = "In user defined schedule " & CStr(iSchedule) & " not enough sleep is used, At least 7 hours is expected, please adjust the schedule!"
        MsgBox (ErrorString)
        Exit Sub
    End If
    If iRequiredPostSleep <> 0 Then
        ErrorString = "In user defined schedule " & CStr(iSchedule) & " the wrong amount of post sleep is used. 0.5 hours is expected, please adjust the schedule!"
        MsgBox (ErrorString)
        Exit Sub
    End If
    If iRequiredPreSleep <> 0 Then
        ErrorString = "In user defined schedule " & CStr(iSchedule) & " the wrong amount of pre sleep is used. 1 hours is expected, please adjust the schedule!"
        MsgBox (ErrorString)
        Exit Sub
    End If
    If iRequiredHygiene > 0 Then
        ErrorString = "In user defined schedule " & CStr(iSchedule) & " not enough hygiene is used. At least 0.5 hours is expected, please adjust the schedule!"
        MsgBox (ErrorString)
        Exit Sub
    End If
    If iRequiredBreakfast <> 0 Then
        ErrorString = "In user defined schedule " & CStr(iSchedule) & " the wrong amount of breakfast is used. 0.5 hours is expected, please adjust the schedule!"
        MsgBox (ErrorString)
        Exit Sub
    End If
    If iRequiredLunch <> 0 Then
        ErrorString = "In user defined schedule " & CStr(iSchedule) & " the wrong amount of lunch is used. 1 hours is expected, please adjust the schedule!"
        MsgBox (ErrorString)
        Exit Sub
    End If
    If iRequiredDinner <> 0 Then
        ErrorString = "In user defined schedule " & CStr(iSchedule) & " the wrong amount of dinner is used. 1 hours is expected, please adjust the schedule!"
        MsgBox (ErrorString)
        Exit Sub
    End If
    End If
    
    iCrewInSchedule = iCrewInSchedule + .Cells(iScheduleHeaderRow + 1, iSchedule + 1).Value
    
Next iSchedule
If Crew_Size = iCrewInSchedule Then
    MsgBox ("All schedules and assigned crew are correct!")
    Worksheets("Schedule and Metabolic").Range("Schedule_was_Checked").Value = 1
ElseIf Crew_Size > iCrewInSchedule Then
    MsgBox ("Not all crew members have an assigned schedule in the user schedule!")
    Worksheets("Schedule and Metabolic").Range("Schedule_was_Checked").Value = 0
Else
    MsgBox ("Too many crew members are assigned a schedule in the user schedule!")
    Worksheets("Schedule and Metabolic").Range("Schedule_was_Checked").Value = 0
End If
End With

protectSheet

End Sub

Sub CalculateUserScheduleMetabolicValues()

Worksheets("Schedule and Metabolic").Unprotect
Application.ScreenUpdating = False

Dim iScheduleHeaderRow As Integer
Dim iAdditionalHeaderRows As Integer
Dim iScheduleRows As Integer
Dim iHeaderRowMetabolicUserSchedule As Integer

Dim lastCol As Integer
Dim lastColMetabolic As Integer

Dim iSchedule As Integer
Dim iCell As Integer
Dim iCrewInSchedule As Integer

Dim colRange
Dim cellRange
Dim cell
Dim CurrentCell
Dim ErrorString As String

Dim iHeaderRow As Integer
Dim sNameColumn As String
Dim MetabolicHeaders
Dim miMetabolicColumn
Dim MetabolicRows
Dim miMetabolicRows
Dim ScheduleHeaders
Dim miScheduleColumns

Dim check As Variant
Dim pos As Integer
Dim posExercise As Integer
Dim posPreExercise As Integer
Dim posPostExercise As Integer
Dim posPotableWater As Integer
Dim posSweat As Integer
Dim posWork As Integer
Dim iMetabolic As Integer
Dim iExerciseTicks As Integer
Dim iBacksteps As Integer
Dim iCurrentScheduleCellRowIndex As Integer
Dim iCurrentMetabolicCellRowIndex As Integer
Dim sEventLabel As String
Dim iPostExerciseTicks As Integer

iScheduleHeaderRow = 59
iAdditionalHeaderRows = 2
iScheduleRows = 48
iHeaderRowMetabolicUserSchedule = 128

' Define the row in which the metabolic value headers are located
iHeaderRow = 1
' Define the column in which the different schedule events are located
sNameColumn = "N"

Set colRange = Cells(iScheduleHeaderRow, Columns.Count).End(xlToLeft)
lastCol = colRange.Column

MetabolicHeaders = Array("O2 Consumption", "CO2 Production", "Urine Water Production", "Feces Production", "Potable Water Consumption", "Hygiene Water Consumption", "Food Consumption", "Sweat/Water Vapor Production", "Heat Load")
miMetabolicColumn = findColumnByName(MetabolicHeaders, "Schedule and Metabolic", iHeaderRow)

MetabolicRows = Array("Sleep", "Pre-Sleep", "Post Sleep", "Personal Hygiene", "Work/Planning", "Recreation", "Pre Exercise", "Exercise", "Post Exercise", "Breakfast", "Lunch", "Dinner")
miMetabolicRows = findRowByName(MetabolicRows, "Schedule and Metabolic", sNameColumn)

' Urine production is considered fix for one day and not actually matched to an event like post exercise of which multiple can exist
' Metabolic Water and Water in Food are just required in internal calculation to derive potable water demand

ScheduleHeaders = Array("O2", "CO2", "urine", "feces", "potable", "hygiene", "food", "sweat", "heat")
miScheduleColumns = findColumnByName(ScheduleHeaders, "Schedule and Metabolic", iHeaderRowMetabolicUserSchedule)

posExercise = Application.Match("Exercise", MetabolicRows, False)
posPreExercise = Application.Match("Pre Exercise", MetabolicRows, False)
posPostExercise = Application.Match("Post Exercise", MetabolicRows, False)
posSweat = Application.Match("Sweat/Water Vapor Production", MetabolicHeaders, False)
posWork = Application.Match("Work/Planning", MetabolicRows, False)

posPotableWater = Application.Match("potable", ScheduleHeaders, False)


With ThisWorkbook.Worksheets("Schedule and Metabolic")
Set colRange = Cells(iHeaderRowMetabolicUserSchedule, Columns.Count).End(xlToLeft)
lastColMetabolic = colRange.Column

' empty the metabolic values
Set cellRange = .Range(.Cells(iHeaderRowMetabolicUserSchedule + 1, 2), .Cells(iHeaderRowMetabolicUserSchedule + iScheduleRows, lastColMetabolic))
For Each cell In cellRange
    With cell
        .Value = 0
        .Interior.color = RGB(255, 255, 255)
    End With
Next cell

iPostExerciseTicks = 0

For iCell = 1 To iScheduleRows
    For iSchedule = 1 To lastCol - 1
        If .Cells(iScheduleHeaderRow + 1, iSchedule + 1) > 0 Then ' only check schedules with assigned crew members
        
            iCrewInSchedule = .Cells(iScheduleHeaderRow + 1, iSchedule + 1).Value
            iCurrentScheduleCellRowIndex = iScheduleHeaderRow + iAdditionalHeaderRows + iCell
            iCurrentMetabolicCellRowIndex = iHeaderRowMetabolicUserSchedule + iCell
            CurrentCell = .Cells(iCurrentScheduleCellRowIndex, iSchedule + 1)
            
            If CurrentCell = "Hygiene" Then
                sEventLabel = "Personal Hygiene"
            ElseIf CurrentCell = "Work" Then
                sEventLabel = "Work/Planning"
            Else
                sEventLabel = CurrentCell
                ' Check if the current schedule entry is a valid schedule entry:
                check = Application.Match(sEventLabel, MetabolicRows, False)
                If IsError(check) Then
                    ErrorString = "In user defined schedule " & CStr(iSchedule) & " the unknown schedule entry '" & sEventLabel & "' was encountered. Please adjust the schedule!"
                    MsgBox (ErrorString)
                    Exit Sub
                End If
            End If
            
            If sEventLabel = "Exercise" Then
                ' check whether this is the first exercise entry, if so we have to mix the metabolic values with the pre exercise values, otherwise we can use the metabolic values for exercise
                If .Cells(iCurrentScheduleCellRowIndex - 1, iSchedule + 1).Value = "Exercise" Then
                    For iMetabolic = 1 To UBound(miMetabolicColumn)
                        .Cells(iCurrentMetabolicCellRowIndex, miScheduleColumns(iMetabolic)).Value = .Cells(iCurrentMetabolicCellRowIndex, miScheduleColumns(iMetabolic)).Value + iCrewInSchedule * .Cells(miMetabolicRows(posExercise), miMetabolicColumn(iMetabolic)).Value
                    Next iMetabolic
                Else
                    For iMetabolic = 1 To UBound(miMetabolicColumn)
                        .Cells(iCurrentMetabolicCellRowIndex, miScheduleColumns(iMetabolic)).Value = .Cells(iCurrentMetabolicCellRowIndex, miScheduleColumns(iMetabolic)).Value + iCrewInSchedule * (.Cells(miMetabolicRows(posExercise), miMetabolicColumn(iMetabolic)).Value + .Cells(miMetabolicRows(posPreExercise), miMetabolicColumn(iMetabolic)).Value) / 2
                    Next iMetabolic
                End If
                
            Else
                If sEventLabel = "Post Exercise" Then
                ' For post exercise, the metabolic values contain the amount of water that was consumed during the total exercise period. However, the value in the metabolic values table is only correct for half an hour of exercise.
                ' Therefore we check here how long the exercise before this block was and sum up the total water demand from that exercise
                
                    iPostExerciseTicks = iPostExerciseTicks + 1 * iCrewInSchedule
                    iExerciseTicks = 0
                    For iBacksteps = 1 To iCurrentScheduleCellRowIndex
                        If .Cells(iCurrentScheduleCellRowIndex - iBacksteps, iSchedule + 1).Value = "Exercise" Then
                            iExerciseTicks = iExerciseTicks + 1
                        ElseIf .Cells(iCurrentScheduleCellRowIndex - iBacksteps, iSchedule + 1).Value = "Post Exercise" Then
                            ' in this case do not exit for loop but also do not increase exercise counter
                        Else
                            Exit For
                        End If
                            
                    Next iBacksteps
                    
                    ' for only one exercise tick (meaning half an hour) the values in the metabolic table are correct
                    pos = Application.Match(Cells(iCurrentScheduleCellRowIndex, iSchedule + 1).Value, MetabolicRows, False)
                    For iMetabolic = 1 To UBound(miMetabolicColumn)
                        If iExerciseTicks > 1 And iMetabolic = posPotableWater Then
                        Else
                            .Cells(iCurrentMetabolicCellRowIndex, miScheduleColumns(iMetabolic)).Value = .Cells(iCurrentMetabolicCellRowIndex, miScheduleColumns(iMetabolic)).Value + iCrewInSchedule * .Cells(miMetabolicRows(pos), miMetabolicColumn(iMetabolic)).Value
                        End If
                    Next iMetabolic
                    If iExerciseTicks > 1 Then
                    ' if more exercise is done we have to recalculate the potable water consumption, but the other metabolic values are correct. We always have 15 minutes pre exercise and the remaining exericse plus the sweat from post exercise.
                    ' Also we have to consider the corresponding nominal sweat rate, for which water consumption is already planned
                        .Cells(iCurrentMetabolicCellRowIndex, miScheduleColumns(posPotableWater)).Value = .Cells(iCurrentMetabolicCellRowIndex, miScheduleColumns(posPotableWater)).Value + iCrewInSchedule * (.Cells(miMetabolicRows(posPreExercise), miMetabolicColumn(posSweat)).Value * 0.25 + .Cells(miMetabolicRows(posExercise), miMetabolicColumn(posSweat)).Value * ((iExerciseTicks / 2) - 0.25) + .Cells(miMetabolicRows(posPostExercise), miMetabolicColumn(posSweat)).Value - .Cells(miMetabolicRows(posWork), miMetabolicColumn(posSweat)).Value * ((iExerciseTicks / 2) + 1))
                    End If
                    
                Else
                    pos = Application.Match(sEventLabel, MetabolicRows, False)
                    
                    For iMetabolic = 1 To UBound(miMetabolicColumn)
                        Cells(iCurrentMetabolicCellRowIndex, miScheduleColumns(iMetabolic)).Value = .Cells(iCurrentMetabolicCellRowIndex, miScheduleColumns(iMetabolic)).Value + iCrewInSchedule * Cells(miMetabolicRows(pos), miMetabolicColumn(iMetabolic)).Value
                    Next iMetabolic
                End If
            End If
            
        
        End If
    Next iSchedule
Next iCell

' Now that we have the total heat output, we can also scale the total food consumption:
Dim TotalFoodRequired As Double
Dim AdditionalWaterToRemove As Double
Dim TotalAdditionalWaterInFood As Double
Dim TotalAdditionalWaterToFeces As Double
Dim TotalAdditionalMetabolicWater As Double
Dim FoodScalingFactor As Double
Dim Crew_Size As Integer

Crew_Size = Worksheets("User Interface").Range("Crew_Size").Value

TotalFoodRequired = (Worksheets("Schedule and Metabolic").Range("SUM_Heat_Load_2") / 12000) * 1.51

' increase urine, feces and food consumption by food scaling factor
FoodScalingFactor = TotalFoodRequired / (Crew_Size * 1.51)

TotalAdditionalWaterInFood = (FoodScalingFactor * Worksheets("Schedule and Metabolic").Range("SUM_Food_Water") - Worksheets("Schedule and Metabolic").Range("SUM_Food_Water")) * Crew_Size
TotalAdditionalWaterToFeces = (FoodScalingFactor * 0.1 - 0.1) * Crew_Size
TotalAdditionalMetabolicWater = (FoodScalingFactor * Worksheets("Schedule and Metabolic").Range("SUM_Metabolic_Water") - Worksheets("Schedule and Metabolic").Range("SUM_Metabolic_Water")) * Crew_Size

AdditionalWaterToRemove = (TotalAdditionalWaterInFood + TotalAdditionalMetabolicWater - TotalAdditionalWaterToFeces)

For iCell = 1 To iScheduleRows
    iCurrentMetabolicCellRowIndex = iHeaderRowMetabolicUserSchedule + iCell
    For iSchedule = 1 To lastCol - 1
        If .Cells(iScheduleHeaderRow + 1, iSchedule + 1) > 0 Then ' only check schedules with assigned crew members
        
            iCrewInSchedule = .Cells(iScheduleHeaderRow + 1, iSchedule + 1).Value
            iCurrentScheduleCellRowIndex = iScheduleHeaderRow + iAdditionalHeaderRows + iCell
            CurrentCell = .Cells(iCurrentScheduleCellRowIndex, iSchedule + 1)
            
            If CurrentCell = "Post Exercise" Then
                .Cells(iCurrentMetabolicCellRowIndex, miScheduleColumns(5)).Value = Cells(iCurrentMetabolicCellRowIndex, miScheduleColumns(5)).Value - AdditionalWaterToRemove * iCrewInSchedule / (iPostExerciseTicks / 2) ' /2 because each ticks is 0.5 hours!
            End If
            
        End If
    Next iSchedule
    
    .Cells(iCurrentMetabolicCellRowIndex, miScheduleColumns(4)).Value = Cells(iCurrentMetabolicCellRowIndex, miScheduleColumns(4)).Value * FoodScalingFactor
    .Cells(iCurrentMetabolicCellRowIndex, miScheduleColumns(7)).Value = Cells(iCurrentMetabolicCellRowIndex, miScheduleColumns(7)).Value * FoodScalingFactor
Next iCell


Worksheets("Schedule and Metabolic").Range("Metabolic_Values_up_to_date").Value = 1

Application.ScreenUpdating = True

protectSheet

End With
End Sub

Sub protectSheet()

Dim iScheduleHeaderRow As Integer
Dim iScheduleRows As Integer
Dim iLastScheduleRow As Integer

iScheduleHeaderRow = 59
iScheduleRows = 48

Worksheets("Schedule and Metabolic").Protect UserInterfaceOnly:=True

With ThisWorkbook.Worksheets("Schedule and Metabolic")
iLastScheduleRow = .Cells.Find(What:="*", _
                    After:=Range("A" & CStr(iScheduleHeaderRow)), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).row

.Range(.Cells(iScheduleHeaderRow + 1, 2), .Cells(iScheduleHeaderRow + 1, iLastScheduleRow)).Locked = False
.Range(.Cells(iScheduleHeaderRow + 3, 2), .Cells(iScheduleHeaderRow + 3 + iScheduleRows, iLastScheduleRow)).Locked = False

End With
End Sub
