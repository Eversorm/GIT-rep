VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub AlleZeilenEinblenden()
  ActiveSheet.Rows.Hidden = False
End Sub

Sub calculatesystemproperties()

'Call the subs calculating entries in the technology sheet
Worksheets("User Interface").Protect UserInterfaceOnly:=True
Application.ScreenUpdating = False

    Call Modul1.calcsystemprops

Application.ScreenUpdating = True
With ThisWorkbook.Worksheets("User Interface").Activate
End With

EH:
End Sub

Sub PerformMCA()

    Call Modul1.InitializeMCA
    Call Tabelle5.ExecuteMCA
    
End Sub

Sub usersguide()

    Worksheets("User Interface").Protect UserInterfaceOnly:=True
    UG_Handbook.Width = 550
    UG_Handbook.Height = 500
    UG_Handbook.Show

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

'This sub informs the user, whether a user defined schedule is already available so that the user does not use a previously defined schedule that may be inappropriate

Worksheets("User Interface").Protect UserInterfaceOnly:=True
Worksheets("Schedule and Metabolic").Protect UserInterfaceOnly:=True

'Check if cell entry has changed
If Not Intersect(Target, Range("Schedule_Number")) Is Nothing Then
    If Range("Schedule_Number").Value = 1 Then Worksheets("Schedule and Metabolic").Range("A1") = "Schedule type:" & vbNewLine & vbNewLine & "Default Schedule"
'Inform the user that a user defined schedule is available or not, if a user defined schedule has been selected in the mission definition
    If Range("Schedule_Number").Value = 2 Then
        Worksheets("Schedule and Metabolic").Range("A1") = "Schedule type:" & vbNewLine & vbNewLine & "User defined" & vbNewLine & "Schedule"
        If Worksheets("Schedule and Metabolic").Range("B176") = "" Then
            MsgBox ("No user defined schedule in Schedule and Metabolic sheet currently available. It is advised to define the schedule now before proceeding!")
            Exit Sub
        End If
        If Worksheets("Schedule and Metabolic").Range("B176") <> "" Then
            MsgBox ("User defined schedule in Schedule and Metabolic sheet available. Adjust if needed!")
            Exit Sub
        End If
    End If

End If

' Check if crew has changed
If Target.Address = "$B$10" Then
    Call Tabelle3.UpdateScheduleToNumberOfCrew
    Worksheets("Schedule and Metabolic").Range("Schedule_was_Checked").Value = 0
    Worksheets("Schedule and Metabolic").Range("Metabolic_Values_up_to_date").Value = 0
End If

End Sub

