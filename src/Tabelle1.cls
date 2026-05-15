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

' If you want to change LiSTOT code, you need to know a couple of things which will be explained in the comments.
' One of these is that you have to press the right click command "show" for tall the hidden parts of the sheets to see
' where the datas are taken from.

' Show the hidden rows in this sheet, there is a table for the functions there.
Sub ShowHiddenRows()
  ActiveSheet.Rows.Hidden = False
End Sub

Sub calculatesystemproperties()

'Call the subs calculating entries in the technology sheet
'Worksheets("User Interface").Protect UserInterfaceOnly:=True
Application.ScreenUpdating = False
Call Modul1.calcsystemprops
Application.ScreenUpdating = True

End Sub

Sub PerformMCA()

    Call Modul1.InitializeMCA
    Call Tabelle5.ExecuteMCA
    
End Sub

Sub usersguide()

    'Worksheets("User Interface").Protect UserInterfaceOnly:=True
    UG_Handbook.Width = 550
    UG_Handbook.Height = 500
    UG_Handbook.Show

End Sub

Sub ResetExcel()        'sub to check if somethign remained blocked. to be erased
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Excel events are back online!"
End Sub


' This sub is called when the mission is selected and it changes the values of the equivalence factors based on assigned presets

Sub SetESMPreset(ByVal mission As String)
    
Dim ws As Worksheet
Set ws = Worksheets("User Interface")

On Error GoTo SafeExit      'safe mechanism
ws.Unprotect
    
    Select Case mission
    
        Case "ISS / LEO"
            ' Valori BVAD Table 4.1 - LEO/ISS
            ws.Range("ESM_LocationF").Value = 7.3       ' [kg/kg]
            ws.Range("ESM_VolumeF").Value = 215.5       ' [kg/m3]
            ws.Range("ESM_PowerF").Value = 87           ' [kg/kW]
            ws.Range("ESM_CoolingF").Value = 146        ' [kg/kW]
            ws.Range("ESM_Crew_TimeF").Value = 0.465    ' [kg/CM*h]
            
        Case "orbital station"
            ' Valori BVAD Table 4.1 - LEO/ISS
            ws.Range("ESM_LocationF").Value = 4447.3       ' [kg/kg]
            ws.Range("ESM_VolumeF").Value = 215.5       ' [kg/m3]
            ws.Range("ESM_PowerF").Value = 87           ' [kg/kW]
            ws.Range("ESM_CoolingF").Value = 146        ' [kg/kW]
            ws.Range("ESM_Crew_TimeF").Value = 0.465    ' [kg/CM*h]
            
        Case "Lunar Surface"
            ' Valori BVAD Table 4.1 - Lunar Surface
            ws.Range("ESM_LocationF").Value = 46        ' [kg/kg]
            ws.Range("ESM_VolumeF").Value = 170.5       ' [kg/m3]
            ws.Range("ESM_PowerF").Value = 108          ' [kg/kW]
            ws.Range("ESM_CoolingF").Value = 109        ' [kg/kW]
            ws.Range("ESM_Crew_TimeF").Value = 1.858    ' [kg/CM*h]
            
        Case "Mars trip"
            ' Valori BVAD Table 4.1 - Mars Surface
            ws.Range("ESM_LocationF").Value = 1853      ' [kg/kg]
            ws.Range("ESM_VolumeF").Value = 141         ' [kg/m3]
            ws.Range("ESM_PowerF").Value = 176          ' [kg/kW]
            ws.Range("ESM_CoolingF").Value = 176        ' [kg/kW]
            ws.Range("ESM_Crew_TimeF").Value = 7.432    ' [kg/CM*h]
    End Select
    MsgBox "ESM factors updated for: " & mission, vbInformation

SafeExit:
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

'This sub handles dynamic UI updates, worksheet protection, and triggers recalculations
'based on user selection (Mission Type, Schedule, and Crew size).

'Worksheets("User Interface").Protect UserInterfaceOnly:=True
Worksheets("Schedule and Metabolic").Protect UserInterfaceOnly:=True


'when the mission type is selected it changes the ESM factors
If Not Intersect(Target, Worksheets("User Interface").Range("MissionType")) Is Nothing Then
    'Worksheets("User Interface").Protect UserInterfaceOnly:=False
    Application.EnableEvents = False ' avoid infinite loops
    Call SetESMPreset(Worksheets("User Interface").Range("MissionType").Value) ' call the sub
    Application.EnableEvents = True
    'Worksheets("User Interface").Protect UserInterfaceOnly:=True
End If


'Check if Schedule number cell has changed
If Not Intersect(Target, Range("Schedule_Number")) Is Nothing Then
    If Range("Schedule_Number").Value = 1 Then
        Worksheets("Schedule and Metabolic").Range("A1") = "Schedule type:" & vbNewLine & vbNewLine & "Default Schedule"
'Inform the user that a user defined schedule is available or not, if a user defined schedule has been selected in the mission definition
    ElseIf Range("Schedule_Number").Value = 2 Then
        Worksheets("Schedule and Metabolic").Range("A1") = "Schedule type:" & vbNewLine & vbNewLine & "User defined" & vbNewLine & "Schedule"
        If Worksheets("Schedule and Metabolic").Range("B176") = "" Then
            MsgBox ("No user defined schedule in Schedule and Metabolic sheet currently available. It is advised to define the schedule now before proceeding!")
            Exit Sub
        ElseIf Worksheets("Schedule and Metabolic").Range("B176") <> "" Then
            MsgBox ("User defined schedule in Schedule and Metabolic sheet available. Adjust if needed!")
            Exit Sub
        End If
    End If
End If

' Check if crew has changed
If Not Intersect(Target, Worksheets("User Interface").Range("Crew_Size")) Is Nothing Then
    Call Tabelle3.UpdateScheduleToNumberOfCrew
    Worksheets("Schedule and Metabolic").Range("Schedule_was_Checked").Value = 0
    Worksheets("Schedule and Metabolic").Range("Metabolic_Values_up_to_date").Value = 0
End If

End Sub

