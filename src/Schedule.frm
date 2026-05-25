VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Schedule 
   Caption         =   "Schedule Planning"
   ClientHeight    =   6480
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   4695
   OleObjectBlob   =   "Schedule.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Schedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheckBox_TS_Click()

    If Schedule.CheckBox_TS.Value = True Then
    
        Schedule.TextBox_E.Enabled = False
        Schedule.TextBox_P.Enabled = False
        Schedule.TextBox_W.Enabled = False
        Schedule.TextBox_R.Enabled = False
        Schedule.TextBox_H.Enabled = False
        Schedule.ComboBox_IS.Enabled = False
        Schedule.ComboBox_DT.Enabled = True

    Else: Schedule.TextBox_E.Enabled = True
          Schedule.TextBox_P.Enabled = True
          Schedule.TextBox_W.Enabled = True
          Schedule.TextBox_R.Enabled = True
          Schedule.TextBox_H.Enabled = True
          Schedule.ComboBox_IS.Enabled = True
          Schedule.ComboBox_DT.Enabled = False
          
    End If
    
End Sub

Private Sub ComboBox_IS_Change()

    If Schedule.ComboBox_IS.Value = "Yes" Then
    
        Schedule.TextBox_E.Enabled = False
        Schedule.TextBox_P.Enabled = False
        Schedule.TextBox_W.Enabled = False
        Schedule.TextBox_R.Enabled = False
        Schedule.TextBox_H.Enabled = False
        
    Else: Schedule.TextBox_E.Enabled = True
          Schedule.TextBox_P.Enabled = True
          Schedule.TextBox_W.Enabled = True
          Schedule.TextBox_R.Enabled = True
          Schedule.TextBox_E.Enabled = True
          
    End If

End Sub

Private Sub CommandButton_Cancel_Click()

    Unload Schedule

End Sub

Private Sub CommandButton_Proceed_Click()

Dim Crew_Size
Dim cell
Dim i
Dim Y
Dim db
Dim lastcolumn
Dim row 'row in which Crew Member number is shown
Dim curCombo
Dim rng As Range


With Worksheets("Schedule and Metabolic").Activate
Worksheets("Schedule and Metabolic").Protect UserInterfaceOnly:=True
Application.ScreenUpdating = False

lastcolumn = Cells(60, Columns.Count).End(xlToLeft).Column
Crew_Size = Worksheets("User Interface").Range("Crew_Size").Value
row = 59

If Worksheets("User Interface").Range("Schedule_Number").Value = 1 Then

    MsgBox ("Default Schedule has been selected in the User Interface!")
    Exit Sub
    
End If

Worksheets("Output Data").Range("J69").Value = 0

If Schedule.CheckBox_TS.Value = True Then

    If lastcolumn - 1 <> Crew_Size Then
    
        With Range(Cells(59, 2), Cells(108, lastcolumn))
            .Clear
            .Interior.ColorIndex = 2
        End With
        With Range(Cells(111, 2), Cells(115, lastcolumn))
            .Clear
            .Interior.ColorIndex = 2
        End With
        
    End If
    
    Range(Cells(111, 2), Cells(115, lastcolumn)).Interior.ColorIndex = 2
    
    With Range(Cells(117, 1), Cells(123, lastcolumn))
        .Clear
        .Interior.ColorIndex = 2
    End With
    
    Worksheets("Schedule and Metabolic").Shapes("Check Inputs").Visible = False
    
    Range(Cells(61, 2), Cells(108, lastcolumn)).ClearContents
    Worksheets("Schedule and Metabolic").DropDowns.Delete
    Range(Cells(111, 2), Cells(115, Crew_Size + 1)).ClearContents
    Range("B128:J176").ClearContents

    If Schedule.ComboBox_DT.Value = "Weekday" Then
    
        Call wsSchedule.Schedule_WD
        Range("A59").Value = "Weekday"
    
        For i = 1 To Crew_Size
    
            Cells(111, i + 1).Value = 6.5 'Work
            Cells(112, i + 1).Value = 3 'Exercise
            Cells(113, i + 1).Value = 2 'Planning/Preparations
            Cells(114, i + 1).Value = 0 'Recreation
            Cells(115, i + 1).Value = 0 'Homekeeping
            
            Range(Cells(111, 1), Cells(115, Crew_Size + 1)).Borders.Weight = xlThin
            
            If Cells(111, i + 1).Value = 0 Then
                Cells(111, i + 1).Interior.ColorIndex = 4
            Else: Cells(111, i + 1).Interior.ColorIndex = 3
            End If
            If Cells(112, i + 1).Value = 0 Then
                Cells(112, i + 1).Interior.ColorIndex = 4
            Else: Cells(112, i + 1).Interior.ColorIndex = 3
            End If
            If Cells(113, i + 1).Value = 0 Then
                Cells(113, i + 1).Interior.ColorIndex = 4
            Else: Cells(113, i + 1).Interior.ColorIndex = 3
            End If
            If Cells(114, i + 1).Value = 0 Then
                Cells(114, i + 1).Interior.ColorIndex = 4
            Else: Cells(114, i + 1).Interior.ColorIndex = 3
            End If
            If Cells(115, i + 1).Value = 0 Then
                Cells(115, i + 1).Interior.ColorIndex = 4
            Else: Cells(115, i + 1).Interior.ColorIndex = 3
            End If
            
        Next
        
    End If
    
    If Schedule.ComboBox_DT.Value = "Weekend Day" Then
    
        Call wsSchedule.Schedule_WendD
        Range("A59").Value = "Weekend Day"
        
        For i = 1 To Crew_Size
    
            Cells(111, i + 1).Value = 0.5 'Work
            Cells(112, i + 1).Value = 3 'Exercise
            Cells(113, i + 1).Value = 0 'Planning/Preparations
            Cells(114, i + 1).Value = 5.5 'Recreation
            Cells(115, i + 1).Value = 2.5 'Homekeeping
            
            Range(Cells(111, 1), Cells(115, Crew_Size + 1)).Borders.Weight = xlThin
            
            If Cells(111, i + 1).Value = 0 Then
                Cells(111, i + 1).Interior.ColorIndex = 4
            Else: Cells(111, i + 1).Interior.ColorIndex = 3
            End If
            If Cells(112, i + 1).Value = 0 Then
                Cells(112, i + 1).Interior.ColorIndex = 4
            Else: Cells(112, i + 1).Interior.ColorIndex = 3
            End If
            If Cells(113, i + 1).Value = 0 Then
                Cells(113, i + 1).Interior.ColorIndex = 4
            Else: Cells(113, i + 1).Interior.ColorIndex = 3
            End If
            If Cells(114, i + 1).Value = 0 Then
                Cells(114, i + 1).Interior.ColorIndex = 4
            Else: Cells(114, i + 1).Interior.ColorIndex = 3
            End If
            If Cells(115, i + 1).Value = 0 Then
                Cells(115, i + 1).Interior.ColorIndex = 4
            Else: Cells(115, i + 1).Interior.ColorIndex = 3
            End If
            
        Next
        
    End If
    
End If

'Check if  total variable crew time is 11.5 hrs if user defined schedule

If Schedule.CheckBox_TS.Value = False Then

    If Schedule.ComboBox_IS.Value = "No" Then

        If val(Schedule.TextBox_E) + val(Schedule.TextBox_H) + val(Schedule.TextBox_W) + val(Schedule.TextBox_R) + val(Schedule.TextBox_P) = 0 Then
           
           MsgBox ("Please enter values for crew tasks!")
           Exit Sub
           
        End If
        
        If val(Schedule.TextBox_E) + val(Schedule.TextBox_H) + val(Schedule.TextBox_W) + val(Schedule.TextBox_R) + val(Schedule.TextBox_P) < 11.5 Then
        
            MsgBox ("Total amount of selected crew tasks are lower than needed (11.5 hrs). Please adjust the values!")
            Exit Sub
            
        End If
        
        If val(Schedule.TextBox_E) + val(Schedule.TextBox_H) + val(Schedule.TextBox_W) + val(Schedule.TextBox_R) + val(Schedule.TextBox_P) > 11.5 Then
        
            MsgBox ("Total amount of selected crew tasks are exceeding the available crew time (11.5 hrs). Please adjust the values!")
            Exit Sub
            
        End If
    
        If val(Schedule.TextBox_E) + val(Schedule.TextBox_H) + val(Schedule.TextBox_W) + val(Schedule.TextBox_R) + val(Schedule.TextBox_P) = 11.5 Then
        
            If lastcolumn - 1 <> Crew_Size Then
        
            With Range(Cells(59, 2), Cells(108, lastcolumn))
                .Clear
                .Interior.ColorIndex = 2
            End With
            With Range(Cells(111, 2), Cells(115, lastcolumn))
                .Clear
                .Interior.ColorIndex = 2
            End With
            
            End If
            
            Range(Cells(111, 2), Cells(115, lastcolumn)).Interior.ColorIndex = 2
            
            With Range(Cells(117, 1), Cells(123, lastcolumn))
                .Clear
                .Interior.ColorIndex = 2
            End With
            
            Worksheets("Schedule and Metabolic").Shapes("Check Inputs").Visible = False
            
            Range(Cells(61, 2), Cells(108, lastcolumn)).ClearContents
            Worksheets("Schedule and Metabolic").DropDowns.Delete
            Range(Cells(111, 2), Cells(115, Crew_Size + 1)).ClearContents
            Range("B128:J176").ClearContents
    
            Call wsSchedule.Schedule_UD
            Range("A59").Value = "User defined"
            
                For i = 1 To Crew_Size
            
                    Cells(111, i + 1).Value = Schedule.TextBox_W.Value 'Work
                    Cells(112, i + 1).Value = Schedule.TextBox_E.Value 'Exercise
                    Cells(113, i + 1).Value = Schedule.TextBox_P.Value 'Planning/Preparations
                    Cells(114, i + 1).Value = Schedule.TextBox_R.Value 'Recreation
                    Cells(115, i + 1).Value = Schedule.TextBox_H.Value 'Homekeeping
                    
                    Range(Cells(111, 1), Cells(115, Crew_Size + 1)).Borders.Weight = xlThin
                    
                    If Cells(111, i + 1).Value = 0 Then
                        Cells(111, i + 1).Interior.ColorIndex = 4
                    Else: Cells(111, i + 1).Interior.ColorIndex = 3
                    End If
                    If Cells(112, i + 1).Value = 0 Then
                        Cells(112, i + 1).Interior.ColorIndex = 4
                    Else: Cells(112, i + 1).Interior.ColorIndex = 3
                    End If
                    If Cells(113, i + 1).Value = 0 Then
                        Cells(113, i + 1).Interior.ColorIndex = 4
                    Else: Cells(113, i + 1).Interior.ColorIndex = 3
                    End If
                    If Cells(114, i + 1).Value = 0 Then
                        Cells(114, i + 1).Interior.ColorIndex = 4
                    Else: Cells(114, i + 1).Interior.ColorIndex = 3
                    End If
                    If Cells(115, i + 1).Value = 0 Then
                        Cells(115, i + 1).Interior.ColorIndex = 4
                    Else: Cells(115, i + 1).Interior.ColorIndex = 3
                    End If
                    
                Next
                
                Worksheets("Output Data").Range("K81").Value = Schedule.TextBox_W.Value 'Work
                Worksheets("Output Data").Range("K82").Value = Schedule.TextBox_E.Value 'Exercise
                Worksheets("Output Data").Range("K83").Value = Schedule.TextBox_R.Value 'Recreation
                Worksheets("Output Data").Range("K84").Value = Schedule.TextBox_H.Value 'Homekeeping
                Worksheets("Output Data").Range("K85").Value = Schedule.TextBox_P.Value 'Planning
            
        End If
    
    End If
        
    If Schedule.ComboBox_IS = "Yes" Then
    
        If lastcolumn - 1 <> Crew_Size Then
        
            With Range(Cells(59, 2), Cells(108, lastcolumn))
                .Clear
                .Interior.ColorIndex = 2
            End With
            With Range(Cells(111, 2), Cells(115, lastcolumn))
                .Clear
                .Interior.ColorIndex = 2
            End With
            
        End If
        
        Range(Cells(111, 2), Cells(115, lastcolumn)).Interior.ColorIndex = 2
        
        With Range(Cells(117, 1), Cells(123, lastcolumn))
            .Clear
            .Interior.ColorIndex = 2
        End With
        
        Worksheets("Schedule and Metabolic").Shapes("Check Inputs").Visible = False
        
        Range(Cells(61, 2), Cells(108, lastcolumn)).ClearContents
        Worksheets("Schedule and Metabolic").DropDowns.Delete
        Range(Cells(111, 2), Cells(115, Crew_Size + 1)).ClearContents
        Range("B128:J176").ClearContents
    
        Call wsSchedule.Schedule_UD
        Range("A59").Value = "User defined"
        
        With Cells(117, 1)
                
            .Value = "Tasks left"
            .Font.Bold = True
            .Font.ColorIndex = 2
            .Interior.ColorIndex = 1
                
        End With
            
        Cells(118, 1).Value = "Work [h]"
        Cells(119, 1).Value = "Exercise [h]"
        Cells(120, 1).Value = "Planning/Prep. [h]"
        Cells(121, 1).Value = "Recreation [h]"
        Cells(122, 1).Value = "Homekeeping [h]"
        Cells(123, 1).Value = "Time left [h]"
            
        Range(Cells(118, 2), Cells(123, Crew_Size + 1)).Locked = False
        Range(Cells(118, 1), Cells(123, Crew_Size + 1)).Borders.Weight = xlThin
        Worksheets("Schedule and Metabolic").Shapes("Check Inputs").Visible = True
        
        For i = 1 To 23
        
            For Y = 1 To Crew_Size
                Worksheets("Schedule and Metabolic").DropDowns("DB " & i & "," & Y).Enabled = False
            Next Y
            
        Next i
            
        For i = 1 To Crew_Size

            Cells(123, i + 1).Value = 11.5
            Cells(123, i + 1).Interior.ColorIndex = 3

            For Y = 1 To 5
                Cells(110 + Y, i + 1).Borders.Weight = xlThin
            Next Y
        
            If Cells(111, i + 1).Value = 0 Then
                Cells(111, i + 1).Interior.ColorIndex = 4
            Else: Cells(111, i + 1).Interior.ColorIndex = 3
            End If
            If Cells(112, i + 1).Value = 0 Then
                Cells(112, i + 1).Interior.ColorIndex = 4
            Else: Cells(112, i + 1).Interior.ColorIndex = 3
            End If
            If Cells(113, i + 1).Value = 0 Then
                Cells(113, i + 1).Interior.ColorIndex = 4
            Else: Cells(113, i + 1).Interior.ColorIndex = 3
            End If
            If Cells(114, i + 1).Value = 0 Then
                Cells(114, i + 1).Interior.ColorIndex = 4
            Else: Cells(114, i + 1).Interior.ColorIndex = 3
            End If
            If Cells(115, i + 1).Value = 0 Then
                Cells(115, i + 1).Interior.ColorIndex = 4
            Else: Cells(115, i + 1).Interior.ColorIndex = 3
            End If
        
        Next i
        
    End If

End If


Worksheets("Output Data").Range("J69").Value = 1

Application.ScreenUpdating = True
End With

Unload Schedule

End Sub

Private Sub TextBox_E_Exit(ByVal Cancel As MSForms.ReturnBoolean)

Dim result

If IsNumeric(Schedule.TextBox_E.Value) = False Or Schedule.TextBox_E.Value < 0 Then
    
    MsgBox ("Please enter a valid value!")
    Schedule.TextBox_E.Value = ""
    Exit Sub
        
End If

If Schedule.TextBox_E.Value >= 11.5 Then
    
    MsgBox ("Please enter a valid value!")
    Schedule.TextBox_E.Value = ""
    Exit Sub
        
End If

result = val(Schedule.TextBox_E.Value) * 10 Mod 1.5 * 10

If result <> 0 Then

    MsgBox ("Exercise must be dividable by 1.5!")
    Schedule.TextBox_E.Value = ""
    Exit Sub
        
End If

End Sub

Private Sub TextBox_H_Change()

If IsNumeric(Schedule.TextBox_H) = False Or Schedule.TextBox_H.Value < 0 Then
    
    MsgBox ("Please enter a valid value!")
    Schedule.TextBox_H.Value = ""
    Exit Sub
        
End If

If Schedule.TextBox_H.Value >= 11.5 Then
    
    MsgBox ("Please enter a valid value!")
    Schedule.TextBox_H.Value = ""
    Exit Sub
        
End If
End Sub

Private Sub TextBox_P_Change()

If IsNumeric(Schedule.TextBox_P) = False Or Schedule.TextBox_P.Value < 0 Then
    
    MsgBox ("Please enter a valid value!")
    Schedule.TextBox_P.Value = ""
    Exit Sub
        
End If

If Schedule.TextBox_P.Value >= 11.5 Then
    
    MsgBox ("Please enter a valid value!")
    Schedule.TextBox_P.Value = ""
    Exit Sub
        
End If

End Sub

Private Sub TextBox_R_Change()

If IsNumeric(Schedule.TextBox_R) = False Or Schedule.TextBox_R.Value < 0 Then
    
    MsgBox ("Please enter a valid value!")
    Schedule.TextBox_R.Value = ""
    Exit Sub
        
End If

If Schedule.TextBox_R.Value >= 11.5 Then
    
    MsgBox ("Please enter a valid value!")
    Schedule.TextBox_R.Value = ""
    Exit Sub
        
End If
End Sub

Private Sub TextBox_W_Change()

If IsNumeric(Schedule.TextBox_W) = False Or Schedule.TextBox_W.Value < 0 Then
    
    MsgBox ("Please enter a valid value!")
    Schedule.TextBox_W.Value = ""
    Exit Sub
        
End If

If Schedule.TextBox_W.Value >= 11.5 Then
    
    MsgBox ("Please enter a valid value!")
    Schedule.TextBox_W.Value = ""
    Exit Sub
        
End If
End Sub

Private Sub UserForm_Initialize()

With Schedule.ComboBox_DT
    
    .AddItem "Weekday"
    .AddItem "Weekend Day"
    .Value = "Weekday"

End With

With Schedule.ComboBox_IS

    .AddItem "Yes"
    .AddItem "No"
    .Value = "No"
    
End With

Schedule.CheckBox_TS.Value = True

End Sub
