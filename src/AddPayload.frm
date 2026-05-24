VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddPayload 
   Caption         =   "Define Additional Payload"
   ClientHeight    =   5325
   ClientLeft      =   -150
   ClientTop       =   -465
   ClientWidth     =   4845
   OleObjectBlob   =   "AddPayload.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddPayload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBox1_Change()

End Sub

Private Sub Command_Proceed_Click()

Dim i As Integer
Dim na As Integer
Dim iCargoCapacityRow As Integer
Dim iCrewRow As Integer
Dim iOtherPayloadRow As Integer
Dim iSumRow As Integer
Dim LowerBound As Integer
Dim UpperBound As Integer
Dim m_total As Double
Dim mission_selected As Integer

na = Cells(8, Columns.Count).End(xlToLeft).Column - 6

iCargoCapacityRow = 11
iCrewRow = 28
iOtherPayloadRow = 29
iSumRow = 30

Worksheets("Resupply Approach").Protect UserInterfaceOnly:=True

If AddPayload.TextBox_CO.Value = "" Then AddPayload.TextBox_CO.Value = 0
If AddPayload.TextBox_CP.Value = "" Then AddPayload.TextBox_CP.Value = 0
If AddPayload.TextBox_ER.Value = "" Then AddPayload.TextBox_ER.Value = 0
If AddPayload.TextBox_HI.Value = "" Then AddPayload.TextBox_HI.Value = 0
If AddPayload.TextBox_M.Value = "" Then AddPayload.TextBox_M.Value = 0
If AddPayload.TextBox_MU.Value = "" Then AddPayload.TextBox_MU.Value = 0
If AddPayload.TextBox_TC.Value = "" Then AddPayload.TextBox_TC.Value = 0
If AddPayload.TextBox_PF.Value = "" Then AddPayload.TextBox_PF.Value = 0
If AddPayload.TextBox_SR.Value = "" Then AddPayload.TextBox_SR.Value = 0
If AddPayload.TextBox_WD.Value = "" Then AddPayload.TextBox_WD.Value = 0

If AddPayload.TextBox_AdditionalCrew.Value = "" Then AddPayload.TextBox_AdditionalCrew.Value = 0


m_total = val(AddPayload.TextBox_CO) + val(AddPayload.TextBox_CP) + val(AddPayload.TextBox_ER) + val(AddPayload.TextBox_HI) + _
          val(AddPayload.TextBox_M) + val(AddPayload.TextBox_MU) + val(AddPayload.TextBox_TC) + val(AddPayload.TextBox_PF) + _
          val(AddPayload.TextBox_SR) + val(AddPayload.TextBox_WD)

If AddPayload.ComboBox_MN.Value = "All" Then
    LowerBound = 1
    UpperBound = na
ElseIf AddPayload.ComboBox_MN.Value <> "All" Then
    LowerBound = AddPayload.ComboBox_MN.Value
    UpperBound = AddPayload.ComboBox_MN.Value
End If

If AddPayload.CheckBox_RemoveCrew.Value Then
    For i = LowerBound To UpperBound
            Worksheets("Resupply Approach").Cells(iSumRow, 6 + i).Value = Worksheets("Resupply Approach").Cells(iSumRow, 6 + i).Value - Worksheets("Resupply Approach").Cells(iCrewRow, 6 + i).Value
            Worksheets("Resupply Approach").Cells(iCrewRow, 6 + i).Value = 0
    Next i
End If

If AddPayload.CheckBox_RemovePayload.Value Then
    For i = LowerBound To UpperBound
            Worksheets("Resupply Approach").Cells(iSumRow, 6 + i).Value = Worksheets("Resupply Approach").Cells(iSumRow, 6 + i).Value - Worksheets("Resupply Approach").Cells(iOtherPayloadRow, 6 + i).Value
            Worksheets("Resupply Approach").Cells(iOtherPayloadRow, 6 + i).Value = 0
    Next i
Else
    For i = LowerBound To UpperBound
        If AddPayload.TextBox_AdditionalCrew.Value + m_total + Worksheets("Resupply Approach").Cells(iSumRow, 6 + i).Value < Worksheets("Resupply Approach").Cells(iCargoCapacityRow, 6 + i).Value Then
            Worksheets("Resupply Approach").Cells(iOtherPayloadRow, 6 + i).Value = m_total
            ' no check necessary, because if no crew mass is set, it simply is 0
            Worksheets("Resupply Approach").Cells(iCrewRow, 6 + i).Value = Worksheets("Resupply Approach").Cells(iCrewRow, 6 + i).Value + AddPayload.TextBox_AdditionalCrew.Value
            Worksheets("Resupply Approach").Cells(iSumRow, 6 + i).Value = Worksheets("Resupply Approach").Cells(iSumRow, 6 + i).Value + Worksheets("Resupply Approach").Cells(iOtherPayloadRow, 6 + i).Value + AddPayload.TextBox_AdditionalCrew.Value
        Else
            MsgBox ("The mission " & CStr(i) & " would go over the defined cargo capacity if this payload is added!")
            Exit Sub
        End If
    Next i

End If

Unload AddPayload

End Sub

Private Sub CommandButton_Cancel_Click()

    Unload AddPayload

End Sub

Private Sub TextBox_AdditionalCrew_Change()

If IsNumeric(AddPayload.TextBox_AdditionalCrew.Value) = False Or AddPayload.TextBox_AdditionalCrew.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        AddPayload.TextBox_AdditionalCrew.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub TextBox_CO_Change()

If IsNumeric(AddPayload.TextBox_CO.Value) = False Or AddPayload.TextBox_CO.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        AddPayload.TextBox_CO.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub TextBox_CP_Change()

AddPayload.TextBox_CP.Enabled = False

If IsNumeric(AddPayload.TextBox_CP.Value) = False Or AddPayload.TextBox_CP.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        AddPayload.TextBox_CP.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub TextBox_ER_Change()

If IsNumeric(AddPayload.TextBox_ER.Value) = False Or AddPayload.TextBox_ER.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        AddPayload.TextBox_ER.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub TextBox_HI_Change()

If IsNumeric(AddPayload.TextBox_HI.Value) = False Or AddPayload.TextBox_HI.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        AddPayload.TextBox_HI.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub TextBox_M_Change()

If IsNumeric(AddPayload.TextBox_M.Value) = False Or AddPayload.TextBox_M.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        AddPayload.TextBox_M.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub TextBox_MU_Change()

If IsNumeric(AddPayload.TextBox_MU.Value) = False Or AddPayload.TextBox_MU.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        AddPayload.TextBox_MU.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub TextBox_PF_Change()

If IsNumeric(AddPayload.TextBox_PF.Value) = False Or AddPayload.TextBox_PF.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        AddPayload.TextBox_PF.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub TextBox_SR_Change()

If IsNumeric(AddPayload.TextBox_SR.Value) = False Or AddPayload.TextBox_SR.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        AddPayload.TextBox_SR.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub TextBox_TC_Change()

If IsNumeric(AddPayload.TextBox_TC.Value) = False Or AddPayload.TextBox_TC.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        AddPayload.TextBox_TC.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub TextBox_WD_Change()

If IsNumeric(AddPayload.TextBox_WD.Value) = False Or AddPayload.TextBox_WD.Value < 0 Then
    
        MsgBox ("Please enter a valid value!")
        AddPayload.TextBox_WD.Value = ""
        Exit Sub
        
End If

End Sub

Private Sub UserForm_Initialize()

Dim i
Dim lastcolumn

lastcolumn = Worksheets("Resupply Approach").Cells(8, Columns.Count).End(xlToLeft).Column

    AddPayload.TextBox_WD.Value = ""
    AddPayload.TextBox_TC.Value = ""
    AddPayload.TextBox_SR.Value = ""
    AddPayload.TextBox_CO.Value = ""
    AddPayload.TextBox_CP.Value = 0
    AddPayload.TextBox_ER.Value = ""
    AddPayload.TextBox_HI.Value = ""
    AddPayload.TextBox_M.Value = ""
    AddPayload.TextBox_MU.Value = ""
    AddPayload.TextBox_PF.Value = ""
    AddPayload.TextBox_AdditionalCrew.Value = ""

    With AddPayload.ComboBox_MN
    
        .AddItem "All"
    
        For i = 1 To lastcolumn - 6
        
            .AddItem i
            .Value = "All"
            
        Next
    
    End With
End Sub
