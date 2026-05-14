VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IniCons 
   Caption         =   "Initial Consumables"
   ClientHeight    =   7400
   ClientLeft      =   -255
   ClientTop       =   -975
   ClientWidth     =   4050
   OleObjectBlob   =   "IniCons.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "IniCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheckBox_C_Click()

If IniCons.CheckBox_C.Value = True Then
    IniCons.ComboBox_MN.Enabled = False
Else: IniCons.ComboBox_MN.Enabled = True
End If

End Sub

Private Sub Command_Proceed_Click()

Dim i
Dim m 'amount of missions
Dim mn
Dim cell
Dim answer
Dim Crew_Size
Dim mission_duration
Dim bConsumablesSelected
Dim mfInitialConsumables

Dim iInitialConsumablesColumn As Integer
Dim iECLSSHeaderRow As Integer
Dim iECLSSConsumablesColumn As Integer
Dim iResupplyConsumablesColumn As Integer
Dim iConsumable As Integer
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

Crew_Size = Worksheets("User Interface").Range("Crew_Size").Value
mission_duration = Worksheets("User Interface").Range("Mission_Duration").Value
m = Worksheets("Resupply Approach").Range("B7").Value

With Worksheets("Resupply Approach").Activate
Application.ScreenUpdating = False
Worksheets("Resupply Approach").Protect UserInterfaceOnly:=True

    
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

iECLSSColumn = findColumnByName(Array(Worksheets("Resupply Approach").Range("B5")), "Output Data", iECLSSHeaderRow)
iECLSSColumn = iECLSSColumn(1)
    
If IniCons.ComboBox_RM.Value = "Yes" Then

    answer = MsgBox("Are you sure to delete previous initial consumable masses?", vbYesNo)
    
    If answer = vbYes Then
        
        If IniCons.CheckBox_C.Value = True Then
            LowerBound = 1
            UpperBound = m
        Else:
            mn = IniCons.ComboBox_MN.Value
            LowerBound = mn
            UpperBound = mn
        End If
        
        For i = LowerBound To UpperBound
            With Worksheets("Resupply Approach")
            
                MissionDuration = .Cells(9, 6 + i)
                MissionCrewSize = .Cells(10, 6 + i)
                
                For iConsumable = 1 To UBound(miInitialConsumableRows)
                    .Cells(miResupplyConsumableRows(iConsumable), iResupplyConsumablesColumn + i).Value = Worksheets("Output Data").Cells(miECLSSConsumableRows(iConsumable), iECLSSColumn).Value * MissionDuration * MissionCrewSize / Crew_Size
                    If .Cells(miResupplyConsumableRows(iConsumable), iResupplyConsumablesColumn + i).Value < 0 Then
                        .Cells(miResupplyConsumableRows(iConsumable), iResupplyConsumablesColumn + i).Value = 0
                    End If
                Next iConsumable
            End With
        Next i
            
    Else: Exit Sub
    End If
Else
    mfInitialConsumables = Array(val(IniCons.TextBox_O2.Value), val(IniCons.TextBox_H2.Value), val(IniCons.TextBox_N2.Value), val(IniCons.TextBox_H2O.Value), val(IniCons.TextBox_Food.Value), val(IniCons.TextBox_CRM.Value), val(IniCons.TextBox_CRD.Value), val(IniCons.TextBox_WW.Value), val(IniCons.TextBox_URI.Value), val(IniCons.TextBox_Brine.Value), val(IniCons.TextBox_ELE.Value), val(IniCons.TextBox_Nutrients.Value), val(IniCons.TextBox_PGC.Value), val(IniCons.TextBox_PBR.Value), val(IniCons.TextBox_CH4.Value), val(IniCons.TextBox_Clothes.Value))
    
    bConsumablesSelected = False
    For i = 0 To UBound(mfInitialConsumables)
        If mfInitialConsumables(i) <> 0 Then
            bConsumablesSelected = True
            Exit For
        End If
    Next i
    If bConsumablesSelected = False Then
        MsgBox ("No initial consumables have been selected!")
        Exit Sub
    End If
    
    'set new values
    For iConsumable = 1 To UBound(miInitialConsumableRows)
        Worksheets("Output Data").Cells(miInitialConsumableRows(iConsumable), iInitialConsumablesColumn + 2).Value = mfInitialConsumables(iConsumable - 1)
    Next iConsumable
    
            
    If IniCons.CheckBox_C.Value = True Then
        
        LowerBound = 1
        UpperBound = m
            
    End If
    
    If IniCons.CheckBox_C.Value = False Then
    
        If IniCons.ComboBox_MN.Value = "" Then
            
            MsgBox ("Please select the mission number for which the initial consumables shall be added!")
            Exit Sub
            
        End If
        
        mn = IniCons.ComboBox_MN.Value
        LowerBound = mn
        UpperBound = mn
    End If
    
    For i = LowerBound To UpperBound
        With Worksheets("Resupply Approach")
        
            MissionDuration = .Cells(9, 6 + i)
            MissionCrewSize = .Cells(10, 6 + i)
            
            For iConsumable = 1 To UBound(miInitialConsumableRows)
                .Cells(miResupplyConsumableRows(iConsumable), iResupplyConsumablesColumn + i).Value = Worksheets("Output Data").Cells(miECLSSConsumableRows(iConsumable), iECLSSColumn).Value * MissionDuration * MissionCrewSize / Crew_Size - Worksheets("Output Data").Cells(miInitialConsumableRows(iConsumable), iInitialConsumablesColumn + 2).Value
                If .Cells(miResupplyConsumableRows(iConsumable), iResupplyConsumablesColumn + i).Value < 0 Then
                    .Cells(miResupplyConsumableRows(iConsumable), iResupplyConsumablesColumn + i).Value = 0
                End If
            Next iConsumable
        End With
    Next i
End If

Application.ScreenUpdating = True
End With
Unload IniCons

End Sub

Private Sub CommandButton_Cancel_Click()

    Unload IniCons

End Sub

Private Sub TextBox_CRD_Change()

If IsNumeric(IniCons.TextBox_CRD.Value) = False Or IniCons.TextBox_CRD.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        IniCons.TextBox_CRD.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub TextBox_CRM_Change()

If IsNumeric(IniCons.TextBox_CRM.Value) = False Or IniCons.TextBox_CRM.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        IniCons.TextBox_CRM.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub TextBox_ELE_Change()

If IsNumeric(IniCons.TextBox_ELE.Value) = False Or IniCons.TextBox_ELE.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        IniCons.TextBox_ELE.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub TextBox_Food_Change()

If IsNumeric(IniCons.TextBox_Food.Value) = False Or IniCons.TextBox_Food.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        IniCons.TextBox_Food.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub TextBox_H2O_Change()

If IsNumeric(IniCons.TextBox_H2O.Value) = False Or IniCons.TextBox_H2O.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        IniCons.TextBox_H2O.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub TextBox_N2_Change()

If IsNumeric(IniCons.TextBox_N2.Value) = False Or IniCons.TextBox_N2.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        IniCons.TextBox_N2.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub TextBox_O2_Change()

If IsNumeric(IniCons.TextBox_O2.Value) = False Or IniCons.TextBox_O2.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        IniCons.TextBox_O2.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub TextBox_URI_Change()

If IsNumeric(IniCons.TextBox_URI.Value) = False Or IniCons.TextBox_URI.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        IniCons.TextBox_URI.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub TextBox_WW_Change()

If IsNumeric(IniCons.TextBox_WW.Value) = False Or IniCons.TextBox_WW.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        IniCons.TextBox_WW.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub TextBox_CH4_Change()

If IsNumeric(IniCons.TextBox_CH4.Value) = False Or IniCons.TextBox_CH4.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        IniCons.TextBox_CH4.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub UserForm_Initialize()

Dim i
Dim m 'amount of missions

m = Worksheets("Resupply Approach").Range("B7").Value

With IniCons.ComboBox_MN
    For i = 1 To m
        .AddItem i
    Next
End With

With IniCons.ComboBox_RM
    .AddItem "Yes"
    .AddItem "No"
    .Value = "No"
End With

IniCons.CheckBox_C.Value = True

End Sub
