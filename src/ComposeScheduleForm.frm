VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ComposeScheduleForm 
   Caption         =   "Compose Schedule"
   ClientHeight    =   8100
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   7785
   OleObjectBlob   =   "ComposeScheduleForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ComposeScheduleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
  
Worksheets("Schedule and Metabolic").Unprotect

Dim Crew_Size As Integer
Dim iHeaderRow As Integer
Dim iScheduleCells As Integer
Dim iAdditionalHeaderRows As Integer
Dim ScheduleString As String
Dim i As Integer
Dim iCell As Integer
Dim iScheduleEntries As Integer
Dim iStartCell As Integer
Dim iEndCell As Integer

Dim EventStartTimes(12) As Variant
Dim EventEndTimes(12) As Variant
Dim EventLabels(12) As Variant
Dim EventColors(12) As Variant
Dim EventDurations(12) As Variant

Dim StartTimes
Dim EndTimes

Dim CurrentCell
Dim NextCell

Dim cell
Dim cellRange
Dim cellRange1
Dim cellRange2


iHeaderRow = 59
iScheduleCells = 48
iAdditionalHeaderRows = 2
Crew_Size = Worksheets("User Interface").Range("Crew_Size").Value

ScheduleString = "Schedule "

EventLabels(1) = "Sleep"
EventLabels(2) = "Work"
EventLabels(3) = "Exercise"
EventLabels(4) = "Recreation"

EventLabels(5) = "Breakfast"
EventLabels(6) = "Lunch"
EventLabels(7) = "Dinner"
EventLabels(8) = "Hygiene"

EventLabels(9) = "Pre-Sleep"
EventLabels(10) = "Post Sleep"
EventLabels(11) = "Post Exercise"

With Worksheets("Schedule and Metabolic")
    EventColors(1) = .Range("N3").Interior.color
    EventColors(2) = .Range("N7").Interior.color
    EventColors(3) = .Range("N10").Interior.color
    EventColors(4) = .Range("N8").Interior.color
    
    EventColors(5) = .Range("N12").Interior.color
    EventColors(6) = .Range("N13").Interior.color
    EventColors(7) = .Range("N14").Interior.color
    EventColors(8) = .Range("N6").Interior.color
    EventDurations(5) = .Range("O12").Value
    EventDurations(6) = .Range("O13").Value
    EventDurations(7) = .Range("O14").Value
    EventDurations(8) = .Range("O6").Value
    
    ' Add pre and post events
    EventColors(9) = .Range("N5").Interior.color
    EventColors(10) = .Range("N4").Interior.color
    EventColors(11) = .Range("N11").Interior.color
    
    EventDurations(9) = .Range("O5").Value
    EventDurations(10) = .Range("O4").Value
    EventDurations(11) = .Range("O11").Value
End With

If ScheduleNumber.Value = "" Then
    MsgBox ("Schedule number cannot be empty! Please set a schedule number.")
    Exit Sub
End If
If CrewForSchedule.Value = "" Then
    CrewForSchedule.Value = 0
End If


' Split the inputs into the individual times:
EventStartTimes(1) = SleepStart.Value
EventStartTimes(2) = WorkStart.Value
EventStartTimes(3) = ExerciseStart.Value
EventStartTimes(4) = RecreationStart.Value
EventStartTimes(5) = BreakfastTime.Value
EventStartTimes(6) = LunchTime.Value
EventStartTimes(7) = DinnerTime.Value
EventStartTimes(8) = HygieneTime.Value

EventEndTimes(1) = SleepEnd.Value
EventEndTimes(2) = WorkEnd.Value
EventEndTimes(3) = ExerciseEnd.Value
EventEndTimes(4) = RecreationEnd.Value


With Worksheets("Schedule and Metabolic")
    ' Clear schedule content
    If CheckBoxDoNotClearSchedule.Value = False Then
        Set cellRange = .Range(.Cells(iHeaderRow + iAdditionalHeaderRows, ScheduleNumber.Value + 1), .Cells(iHeaderRow + iAdditionalHeaderRows + iScheduleCells, ScheduleNumber.Value + 1))
        For Each cell In cellRange
            With cell
                .Value = ""
                .Interior.color = RGB(255, 255, 255)
            End With
        Next cell
    End If

    ' Set headings and number of crew for this schedule
    .Cells(iHeaderRow, ScheduleNumber.Value + 1).Value = ScheduleString & CStr(ScheduleNumber.Value)
    .Cells(iHeaderRow + 1, ScheduleNumber.Value + 1).Value = CrewForSchedule.Value
    .Cells(iHeaderRow + 2, ScheduleNumber.Value + 1).Value = "task"
    
    For i = 0 To 2
        .Cells(iHeaderRow + i, ScheduleNumber.Value + 1).Font.Bold = True
        .Cells(iHeaderRow + i, ScheduleNumber.Value + 1).Font.color = vbWhite
        .Cells(iHeaderRow + i, ScheduleNumber.Value + 1).Interior.color = vbBlack
        .Cells(iHeaderRow + i, ScheduleNumber.Value + 1).HorizontalAlignment = xlCenter
    Next i
    
    For iScheduleEntry = 1 To 8
        StartTimes = Split(EventStartTimes(iScheduleEntry), ";")
        If iScheduleEntry < 5 Then
            EndTimes = Split(EventEndTimes(iScheduleEntry), ";")
        End If
        
        If iScheduleEntry < 5 And UBound(StartTimes) <> UBound(EndTimes) Then
            MsgBox ("For each start time an end time is required and vice versa!")
            Exit Sub
        End If
        
        ' Now start filling the schedule starting with sleep
        For i = 0 To UBound(StartTimes)
            StartTime = TimeValue(StartTimes(i))
            
            If iScheduleEntry < 5 Then
                EndTime = TimeValue(EndTimes(i))
            Else
                EndTime = StartTime + EventDurations(iScheduleEntry) / 24
            End If
            
            iStartCell = 0
            iEndCell = 0
            For iCell = 1 To iScheduleCells
                CurrentCell = Format$((.Cells(iHeaderRow + iAdditionalHeaderRows + iCell, 1).Value), "hh:mm:SS")
                CurrentCell = (TimeValue(CurrentCell))
                If CurrentCell > StartTime And iStartCell = 0 Then
                    iStartCell = iCell - 1
                End If
                
                ' This comparison is necessary because rounding errors in the ime prevent excel from noticing that 05:30:00 is identical to 05:30:00 ....
                If (CurrentCell - EndTime) > -(1 / (24 * 60)) And iEndCell = 0 Then
                    iEndCell = iCell - 1
                End If
            Next iCell
            If iStartCell = 0 Then
                ' in this case the start is the last cell of the previous day (event starts at 23:30)
                iStartCell = iScheduleCells
            End If
            If iEndCell > iScheduleCells Then
                iEndCell = iScheduleCells
            End If
            
            If iStartCell > iEndCell Then
                ' In this case the activity is from evening to morning (starts before 00:00 and ends after 00:00)
                Set cellRange1 = .Range(.Cells(iHeaderRow + iAdditionalHeaderRows + 1, ScheduleNumber.Value + 1), .Cells(iHeaderRow + iAdditionalHeaderRows + iEndCell, ScheduleNumber.Value + 1))
                Set cellRange2 = .Range(.Cells(iHeaderRow + iAdditionalHeaderRows + iStartCell, ScheduleNumber.Value + 1), .Cells(iHeaderRow + iAdditionalHeaderRows + iScheduleCells, ScheduleNumber.Value + 1))
                Set cellRange = Union(cellRange1, cellRange2)
            Else
                Set cellRange = .Range(.Cells(iHeaderRow + iAdditionalHeaderRows + iStartCell, ScheduleNumber.Value + 1), .Cells(iHeaderRow + iAdditionalHeaderRows + iEndCell, ScheduleNumber.Value + 1))
            End If
            For Each cell In cellRange
                With cell
                    .Value = EventLabels(iScheduleEntry)
                    .Interior.color = EventColors(iScheduleEntry)
                    .HorizontalAlignment = xlCenter
                End With
            Next cell
        Next i
    Next iScheduleEntry

    ' Add pre and post events:
    
    For iCell = 2 To iScheduleCells - 1
        iStartCell = iScheduleCells * 2
        iEndCell = 0
        CurrentCell = .Cells(iHeaderRow + iAdditionalHeaderRows + iCell, ScheduleNumber.Value + 1).Value
        NextCell = .Cells(iHeaderRow + iAdditionalHeaderRows + iCell + 1, ScheduleNumber.Value + 1).Value
        If CurrentCell <> "Sleep" And NextCell = "Sleep" Then
            iEndCell = iCell
            iScheduleEntry = 9
            iStartCell = iEndCell - ((EventDurations(iScheduleEntry) * 2) - 1)
        End If
        If CurrentCell = "Sleep" And NextCell <> "Sleep" Then
            iStartCell = iCell + 1
            iScheduleEntry = 10
            iEndCell = iStartCell + ((EventDurations(iScheduleEntry) * 2) - 1)
        End If
        If CurrentCell = "Exercise" And NextCell <> "Exercise" Then
            iStartCell = iCell + 1
            iScheduleEntry = 11
            iEndCell = iStartCell + ((EventDurations(iScheduleEntry) * 2) - 1)
        End If
        
        ' Only execute the following code if a pre or post event was found
        If iStartCell <> iScheduleCells * 2 Then
            ' Limit start and end times
            If iStartCell < 1 Then
                ' in this case, the start of the pre event should be before 00:00 (e.g. 23:30) so we have to set the start cells to the corresponding value
                iStartCell = iScheduleCells + (iStartCell - 1)
            End If
            If iEndCell > iScheduleCells Then
            ' in this case the end of the post event should be after 00:00 (e.g. 00:30) so we have to set the end cell accordingly
                iEndCell = iEndCell - iScheduleCells
            End If
            
            If iStartCell > iEndCell Then
                ' In this case the activity is from evening to morning (starts before 00:00 and ends after 00:00)
                Set cellRange1 = .Range(.Cells(iHeaderRow + iAdditionalHeaderRows + 1, ScheduleNumber.Value + 1), .Cells(iHeaderRow + iAdditionalHeaderRows + iEndCell, ScheduleNumber.Value + 1))
                Set cellRange2 = .Range(.Cells(iHeaderRow + iAdditionalHeaderRows + iStartCell, ScheduleNumber.Value + 1), .Cells(iHeaderRow + iAdditionalHeaderRows + iScheduleCells, ScheduleNumber.Value + 1))
                Set cellRange = Union(cellRange1, cellRange2)
            Else
                Set cellRange = .Range(.Cells(iHeaderRow + iAdditionalHeaderRows + iStartCell, ScheduleNumber.Value + 1), .Cells(iHeaderRow + iAdditionalHeaderRows + iEndCell, ScheduleNumber.Value + 1))
            End If
            For Each cell In cellRange
                With cell
                    .Value = EventLabels(iScheduleEntry)
                    .Interior.color = EventColors(iScheduleEntry)
                    .HorizontalAlignment = xlCenter
                End With
            Next cell
        End If
    Next iCell
    
End With


Worksheets("Schedule and Metabolic").protectSheet

End Sub
