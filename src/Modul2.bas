Attribute VB_Name = "Modul2"
Option Explicit
' some table appearances and check box for stuff that must remain enabled

Sub performsensitivityanalysis_up()

    SensAnalysis.Width = 500
    SensAnalysis.Height = 550
    SensAnalysis.Show

End Sub

Sub showaddpaylaod()
    AddPayload.Width = 500
    AddPayload.Height = 550
    AddPayload.Show

End Sub

Sub iniconsumables()
    IniCons.Width = 300
    IniCons.Height = 650
    IniCons.Show

End Sub

Sub info_schedule()
    Info_Schedule_1.Width = 500
    Info_Schedule_1.Height = 550
    Info_Schedule_1.Show

End Sub

Sub ws()

If Sheets("ECLSS Composition").Shapes("WS").OLEFormat.Object.Value <> 1 Then
    MsgBox ("This Checkbox must remain enabled!")
    Sheets("ECLSS Composition").Shapes("WS").OLEFormat.Object.Value = 1
End If

End Sub

Sub H2O()

If Sheets("ECLSS Composition").Shapes("H2O").OLEFormat.Object.Value <> 1 Then
    MsgBox ("This Checkbox must remain enabled!")
    Sheets("ECLSS Composition").Shapes("H2O").OLEFormat.Object.Value = 1
End If

End Sub

Sub algreact()

If Sheets("ECLSS Composition").Shapes("ALG").OLEFormat.Object.Value <> 1 Then
    MsgBox ("This Checkbox must remain enabled!")
    Sheets("ECLSS Composition").Shapes("ALG").OLEFormat.Object.Value = 1
End If

End Sub

Sub shapeclick()
Dim tmp
Dim r
Dim dd

Set dd = Worksheets("Schedule and Metabolic").Shapes(Application.Caller).OLEFormat.Object
tmp = dd.List(dd.ListIndex)

Set r = Worksheets("Schedule and Metabolic").Shapes(Application.Caller).TopLeftCell
Worksheets("Schedule and Metabolic").Range(Cells(r.row, r.Column), Cells(r.row, r.Column)).Value = tmp


End Sub
