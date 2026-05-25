VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SensAnalysis 
   Caption         =   "Sensitivity Analysis"
   ClientHeight    =   7260
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   7005
   OleObjectBlob   =   "SensAnalysis.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SensAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBox_Ana_Change()

    If SensAnalysis.ComboBox_Ana.Value = "ESM" Then
        SensAnalysis.CheckBox_R.Enabled = False
        SensAnalysis.CheckBox_TRL.Enabled = False
        SensAnalysis.CheckBox_R.Value = 0
        SensAnalysis.CheckBox_TRL.Value = 0
    Else: SensAnalysis.CheckBox_R.Enabled = True
          SensAnalysis.CheckBox_TRL.Enabled = True
    End If

End Sub

Private Sub ComboBox_ATxp_Change()

    If SensAnalysis.ComboBox_ATxp.Value = "Iterative" Then
        SensAnalysis.TextBox_Per.Enabled = False
    Else: SensAnalysis.TextBox_Per.Enabled = True
    End If
    
    If SensAnalysis.ComboBox_ATxp.Value = "Manual" Then
        SensAnalysis.TextBox_Itstep.Enabled = False
    Else: SensAnalysis.TextBox_Itstep.Enabled = True
    End If
        
End Sub

Private Sub CommandButton_Perform_Click()

Dim iSign As Integer
Dim iColumnsToAdjust As Integer
Dim csColumnsToAdjust(0 To 6) As String
Dim miColumns
Dim AdjustmentValue As Double
Dim iRankColumn As Integer
Dim iNameColumn As Integer
Dim iColumn As Integer
Dim iRow As Integer
Dim iIteration As Integer
Dim iRowToAdjust As Integer
Dim na As Integer
Dim iHeaderRow As Integer

If SensAnalysis.ComboBox_ATxp = "Manual" And SensAnalysis.TextBox_Per.Value = "" Then

    MsgBox ("Enter required values: variation amount")
    Exit Sub

End If

If SensAnalysis.CheckBox_Mass = False And SensAnalysis.CheckBox_Vol = False And SensAnalysis.CheckBox_P = False And _
SensAnalysis.CheckBox_C = False And SensAnalysis.CheckBox_Main = False And SensAnalysis.CheckBox_R = False And _
SensAnalysis.CheckBox_TRL = False Then

    MsgBox ("At least one system value must be select before proceeding!")
    Exit Sub
    
End If

' check whether we should increase or decrease the values in the sensitivity analysis
If SensAnalysis.ComboBox_DecInc.Value = "Increase" Then
    iSign = 1
Else
    iSign = -1
End If

iHeaderRow = 4
' get the rank column we want to compare:
If SensAnalysis.ComboBox_Ana.Value = "MCA" Then
    miColumns = findColumnByName(Array("MCA Rank"), "MCA ESM", iHeaderRow)
    iRankColumn = miColumns(1)
ElseIf SensAnalysis.ComboBox_Ana.Value = "ESM" Then
    miColumns = findColumnByName(Array("ESM Rank"), "MCA ESM", iHeaderRow)
    iRankColumn = miColumns(1)
End If
' get the column containing the system names
miColumns = findColumnByName(Array("Name of Assembly"), "MCA ESM", iHeaderRow)
iNameColumn = miColumns(1)

' We check which columns we want to adjust and select the column array for the values accordingly
iColumnsToAdjust = 0
If SensAnalysis.CheckBox_Mass = True Then
    csColumnsToAdjust(iColumnsToAdjust) = "Mass"
    iColumnsToAdjust = iColumnsToAdjust + 1
End If
If SensAnalysis.CheckBox_Vol = True Then
    csColumnsToAdjust(iColumnsToAdjust) = "Volume"
    iColumnsToAdjust = iColumnsToAdjust + 1
End If
If SensAnalysis.CheckBox_P = True Then
    csColumnsToAdjust(iColumnsToAdjust) = "Power"
    iColumnsToAdjust = iColumnsToAdjust + 1
End If
If SensAnalysis.CheckBox_C = True Then
    csColumnsToAdjust(iColumnsToAdjust) = "Cooling"
    iColumnsToAdjust = iColumnsToAdjust + 1
End If
If SensAnalysis.CheckBox_Main = True Then
    csColumnsToAdjust(iColumnsToAdjust) = "Maintenance"
    iColumnsToAdjust = iColumnsToAdjust + 1
End If
If SensAnalysis.CheckBox_R = True Then
    csColumnsToAdjust(iColumnsToAdjust) = "Reliability"
    iColumnsToAdjust = iColumnsToAdjust + 1
End If
If SensAnalysis.CheckBox_TRL = True Then
    csColumnsToAdjust(iColumnsToAdjust) = "TRL"
    iColumnsToAdjust = iColumnsToAdjust + 1
End If
           
miColumns = findColumnByName(csColumnsToAdjust, "MCA ESM", iHeaderRow)

' check by how much percent the value should be adjusted
If SensAnalysis.ComboBox_ATxp = "Manual" Then
    AdjustmentValue = SensAnalysis.TextBox_Per.Value / 100
ElseIf SensAnalysis.ComboBox_ATxp = "Iterative" Then
    AdjustmentValue = SensAnalysis.TextBox_Itstep.Value / 100
End If

With ThisWorkbook.Worksheets("MCA ESM")
na = .Cells(Rows.Count, 1).End(xlUp).row

' check which system should be adjusted
For iRow = iHeaderRow + 2 To na
    If .Cells(iRow, iNameColumn).Value = SensAnalysis.ComboBox_Sysn.Value Then
        iRowToAdjust = iRow
        Exit For
    End If
Next iRow

' check the current ranks before the iteration
Dim miRank() As Integer

ReDim miRank(iHeaderRow + 2 To na)

For iRow = iHeaderRow + 2 To na
    miRank(iRow) = .Cells(iRow, iRankColumn).Value
Next iRow

iIteration = 0
' Flag to start an iteration in case the iterative approach is chosen
Iteration:
iIteration = iIteration + 1
' Now adjust the values
For iColumn = 1 To iColumnsToAdjust
    .Cells(iRowToAdjust, miColumns(iColumn)).Value = .Cells(iRowToAdjust, miColumns(iColumn)).Value * (1 + iSign * AdjustmentValue)
    ' Limit maximum value for TRL
    If csColumnsToAdjust(iColumn - 1) = "TRL" Then
        If .Cells(iRowToAdjust, miColumns(iColumn)).Value > 9 Then
            .Cells(iRowToAdjust, miColumns(iColumn)).Value = 9
        End If
    End If
Next iColumn

If SensAnalysis.ComboBox_Ana.Value = "MCA" Then
    ' for the sensitivity analysis of the MCA, we adjust the values in the MCA ESM sheet and then call the wsMCA.ExecuteMCA macro for the MCA calculation!
     wsMCA.ExecuteMCA
ElseIf SensAnalysis.ComboBox_Ana.Value = "ESM" Then
    ' for the ESM sensitvity analysis, we adjust the values in the MCA sheet, and then call the ESM sub but specify that we want to perform the ESM in the MCA sheet
    modCalculations.esmcalc "MCA ESM"
End If

If SensAnalysis.ComboBox_ATxp = "Iterative" Then
    ' check the current ranks, in case the first place changed, we finish the iteration by going to the EndIteration flag
    If miRank(iRowToAdjust) = .Cells(iRowToAdjust, iRankColumn).Value And iIteration < 1001 Then
    Else
        ' Inform the user after how much percent change the rank change occured
        If iIteration = 1000 Then
            MsgBox ("The rank did not change during the sensitivity analysis!")
        Else
            MsgBox ("The rank changed after a percentual change of " & CStr(iIteration * SensAnalysis.TextBox_Itstep.Value) & " %")
        End If
        Exit Sub
    End If
    ' If the first rank has not changed, we go to the iteration flag and adjust the values again
    GoTo Iteration
End If
End With

End Sub



Private Sub CommandButton_Cancel_Click()

Unload SensAnalysis

End Sub

Private Sub TextBox_Itstep_Change()

If SensAnalysis.TextBox_Itstep.Value < 1 Or SensAnalysis.TextBox_Itstep.Value > 10 Or IsNumeric(SensAnalysis.TextBox_Itstep.Value) = False Then

    MsgBox ("Invalid value inserted!")
    SensAnalysis.TextBox_Itstep.Value = ""
    Exit Sub
    
End If

End Sub

Private Sub TextBox_Per_Change()

If SensAnalysis.TextBox_Per.Value < 0 Or SensAnalysis.TextBox_Per.Value > 100 Or IsNumeric(SensAnalysis.TextBox_Per.Value) = False Then

    MsgBox ("Invalid value inserted!")
    SensAnalysis.TextBox_Per.Value = ""
    Exit Sub
    
End If

End Sub

Private Sub UserForm_Initialize()

Dim na
Dim i

With Worksheets("MCA ESM").Activate
na = Worksheets("MCA ESM").Cells(Rows.Count, 1).End(xlUp).row

With SensAnalysis.ComboBox_DecInc

    .AddItem "Increase"
    .AddItem "Decrease"
    .Value = "Increase"
        
End With

With SensAnalysis.ComboBox_ATxp

    .AddItem "Manual"
    .AddItem "Iterative"
    .Value = "Manual"
        
End With

With SensAnalysis.ComboBox_Sysn

    For i = 6 To na
        .AddItem Worksheets("MCA ESM").Cells(i, 1).Value
        .Value = Worksheets("MCA ESM").Cells(6, 1).Value
    Next
    
End With

With SensAnalysis.ComboBox_Ana

    .AddItem "ESM"
    .AddItem "MCA"
    .Value = "MCA"
    
End With

SensAnalysis.TextBox_Itstep.Value = 5

End With
Exit Sub
    
EH:
End Sub
