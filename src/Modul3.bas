Attribute VB_Name = "Modul3"
Option Explicit

Sub plotresults()

'All the subs in this module are meant to plot graphs. The procedure is always the same
Application.ScreenUpdating = False

    Call ESMDiagramm
    Call ECLSSComposition
    Call CO2PPDiagramm
    Call CO2ProductionDiagramm
    Call O2ConDiagramm
    Call O2PPDiagramm
    Call PotHygwaterDiagramm
    Call WWurineDiagramm
    Call PowWWurineDiagramm

Application.ScreenUpdating = True

End Sub

Sub plotbiores()

Application.ScreenUpdating = False

    Call ECLSSComposition
    Call ESMDiagramm
    On Error Resume Next
    Worksheets("Graphical Output").ChartObjects("CO2 Partial Pressure").Delete
    Worksheets("Graphical Output").ChartObjects("O2 Partial Pressure").Delete
    Worksheets("Graphical Output").ChartObjects("Daily WRM Power").Delete
    
Application.ScreenUpdating = True

End Sub

Sub ESMDiagramm()

Dim MD
Dim Dia_ESM As Chart
Dim frame As ChartObject
Dim lastRow
Dim ws As Worksheet
Dim rng As Range
Dim rg As Range
Dim MinValue
Dim MaxValue
Dim ser1
Dim ser2
Dim ser3
Dim ser4
Dim ser5

Application.ScreenUpdating = False

MD = Worksheets("User Interface").Range("Mission_Duration").Value

On Error Resume Next
ThisWorkbook.Worksheets("Graphical Output").ChartObjects("ESM System Mass").Delete

    Set ws = Worksheets("Output Data")
        lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
            If lastRow < 7 Then
            MsgBox ("Nothing to plot!")
            Exit Sub
            End If

With Worksheets("Output Data").Activate
    
    If Worksheets("Output Data").Range("C7") <> "" Then Set rng = Worksheets("Output Data").Range(Cells(7, "C"), Cells(lastRow, "G"))
    If Worksheets("Output Data").Range("C7") = "" Then Set rng = Worksheets("Output Data").Range(Cells(7, "D"), Cells(lastRow, "G"))
    If Worksheets("Output Data").Range("C7") = "" And Worksheets("Output Data").Range("D7") = "" Then Set rng = Worksheets("Output Data").Range(Cells(7, "E"), Cells(lastRow, "G"))
    
    For Each rg In rng
        If rg <> "" Then
            MinValue = WorksheetFunction.Min(rng)
            MaxValue = WorksheetFunction.Max(rng)
        End If
    Next
    
End With

    Set frame = ThisWorkbook.Worksheets("Graphical Output").ChartObjects.Add(460, 1, 450, 300)
    Set Dia_ESM = frame.Chart

If Worksheets("Output Data").Range("C7") <> "" Then

'Quite a bit of code necessary only to format the caption
    With Dia_ESM
    
            .Parent.Name = "ESM System Mass"
            .PlotArea.Border.ColorIndex = 14
            .SetSourceData rng
            .ChartType = xlXYScatterSmoothNoMarkers
            .HasLegend = True
            If Worksheets("Output Data").Range("C7") <> "" Then
            .SeriesCollection(1).Values = Worksheets("Output Data").Range(Cells(7, "C"), Cells(lastRow, "C"))
            .SeriesCollection(1).Name = "=""Open Loop"""
            Else: Set ser1 = Dia_ESM.SeriesCollection(1)
                With ser1.Format.Line
                        .Visible = msoFalse
                End With
                ser1.Name = vbNullString
            End If
            If Worksheets("Output Data").Range("D7") <> "" Then
            .SeriesCollection(2).Name = "=""Partial Loop"""
            .SeriesCollection(2).Values = Worksheets("Output Data").Range(Cells(7, "D"), Cells(lastRow, "D"))
            Else: Set ser2 = Dia_ESM.SeriesCollection(2)
                With ser2.Format.Line
                        .Visible = msoFalse
                End With
                ser2.Name = vbNullString
            End If
            If Worksheets("Output Data").Range("E7") <> "" Then
            .SeriesCollection(3).Name = "=""Closed Loop"""
            .SeriesCollection(3).Values = Worksheets("Output Data").Range(Cells(7, "E"), Cells(lastRow, "E"))
            Else: Set ser3 = Dia_ESM.SeriesCollection(3)
                With ser3.Format.Line
                        .Visible = msoFalse
                End With
                ser3.Name = vbNullString
            End If
            If Worksheets("Output Data").Range("F7") <> "" Then
            .SeriesCollection(4).Name = "=""Bioregenerative"""
            .SeriesCollection(4).Values = Worksheets("Output Data").Range(Cells(7, "F"), Cells(lastRow, "F"))
            Else: Set ser4 = Dia_ESM.SeriesCollection(4)
                With ser4.Format.Line
                        .Visible = msoFalse
                End With
                ser4.Name = vbNullString
            End If
            If Worksheets("Output Data").Range("G7") <> "" Then
            .SeriesCollection(5).Name = "=""Hybrid"""
            .SeriesCollection(5).Values = Worksheets("Output Data").Range(Cells(7, "G"), Cells(lastRow, "G"))
            Else: Set ser5 = Dia_ESM.SeriesCollection(5)
                With ser5.Format.Line
                        .Visible = msoFalse
                End With
                ser5.Name = vbNullString
            End If
            .HasTitle = True
            .ChartTitle.Text = "ESM System Masses"
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).MinimumScale = Worksheets("Output Data").Range("B7").Value
            .Axes(xlCategory).MaximumScale = MD
            .Axes(xlCategory).AxisTitle.Text = "Mission Duration [d]"
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).MinimumScale = WorksheetFunction.RoundUp(MinValue * 0.8, -2)
            .Axes(xlValue).MaximumScale = WorksheetFunction.RoundUp(MaxValue * 1.1, -2)
            .Axes(xlValue).AxisTitle.Text = "ESM mass [kg]"
            
    End With
End If

If Worksheets("Output Data").Range("C7") = "" Then

    With Dia_ESM
    
            .Parent.Name = "ESM System Mass"
            .PlotArea.Border.ColorIndex = 14
            .SetSourceData rng
            .ChartType = xlXYScatterSmoothNoMarkers
            .HasLegend = True
            If Worksheets("Output Data").Range("D7") <> "" Then
            .SeriesCollection(1).Values = Worksheets("Output Data").Range(Cells(7, "D"), Cells(lastRow, "D"))
            .SeriesCollection(1).Name = "=""Partial Loop"""
            Else: Set ser1 = Dia_ESM.SeriesCollection(1)
                With ser1.Format.Line
                        .Visible = msoFalse
                End With
                ser1.Name = vbNullString
            End If
            If Worksheets("Output Data").Range("E7") <> "" Then
            .SeriesCollection(2).Name = "=""Closed Loop"""
            .SeriesCollection(2).Values = Worksheets("Output Data").Range(Cells(7, "E"), Cells(lastRow, "E"))
            Else: Set ser2 = Dia_ESM.SeriesCollection(2)
                With ser2.Format.Line
                        .Visible = msoFalse
                End With
                ser2.Name = vbNullString
            End If
            If Worksheets("Output Data").Range("F7") <> "" Then
            .SeriesCollection(3).Name = "=""Bioregenerative"""
            .SeriesCollection(3).Values = Worksheets("Output Data").Range(Cells(7, "F"), Cells(lastRow, "F"))
            Else: Set ser3 = Dia_ESM.SeriesCollection(3)
                With ser3.Format.Line
                        .Visible = msoFalse
                End With
                ser3.Name = vbNullString
            End If
            If Worksheets("Output Data").Range("G7") <> "" Then
            .SeriesCollection(4).Name = "=""Hybrid"""
            .SeriesCollection(4).Values = Worksheets("Output Data").Range(Cells(7, "G"), Cells(lastRow, "G"))
            Else: Set ser4 = Dia_ESM.SeriesCollection(4)
                With ser4.Format.Line
                        .Visible = msoFalse
                End With
                ser4.Name = vbNullString
            End If
            .HasTitle = True
            .ChartTitle.Text = "ESM System Masses"
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).MinimumScale = Worksheets("Output Data").Range("B7").Value
            .Axes(xlCategory).MaximumScale = MD
            .Axes(xlCategory).AxisTitle.Text = "Mission Duration [d]"
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).MinimumScale = WorksheetFunction.RoundUp(MinValue * 0.8, -2)
            .Axes(xlValue).MaximumScale = WorksheetFunction.RoundUp(MaxValue * 1.1, -2)
            .Axes(xlValue).AxisTitle.Text = "ESM mass [kg]"
            
    End With
End If

If Worksheets("Output Data").Range("C7") = "" And Worksheets("Output Data").Range("D7") = "" Then

    With Dia_ESM
    
            .Parent.Name = "ESM System Mass"
            .PlotArea.Border.ColorIndex = 14
            .SetSourceData rng
            .ChartType = xlXYScatterSmoothNoMarkers
            .HasLegend = True
            If Worksheets("Output Data").Range("E7") <> "" Then
            .SeriesCollection(1).Values = Worksheets("Output Data").Range(Cells(7, "E"), Cells(lastRow, "E"))
            .SeriesCollection(1).Name = "=""Closed Loop"""
            Else: Set ser1 = Dia_ESM.SeriesCollection(1)
                With ser1.Format.Line
                        .Visible = msoFalse
                End With
                ser1.Name = vbNullString
            End If
            If Worksheets("Output Data").Range("F7") <> "" Then
            .SeriesCollection(2).Name = "=""Bioregenerative"""
            .SeriesCollection(2).Values = Worksheets("Output Data").Range(Cells(7, "F"), Cells(lastRow, "F"))
            Else: Set ser2 = Dia_ESM.SeriesCollection(2)
                With ser2.Format.Line
                        .Visible = msoFalse
                End With
                ser2.Name = vbNullString
            End If
            If Worksheets("Output Data").Range("G7") <> "" Then
            .SeriesCollection(3).Name = "=""Hybrid"""
            .SeriesCollection(3).Values = Worksheets("Output Data").Range(Cells(7, "G"), Cells(lastRow, "G"))
            Else: Set ser3 = Dia_ESM.SeriesCollection(3)
                With ser3.Format.Line
                        .Visible = msoFalse
                End With
                ser3.Name = vbNullString
            End If
            .HasTitle = True
            .ChartTitle.Text = "ESM System Masses"
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).MinimumScale = Worksheets("Output Data").Range("B7").Value
            .Axes(xlCategory).MaximumScale = MD
            .Axes(xlCategory).AxisTitle.Text = "Mission Duration [d]"
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).MinimumScale = WorksheetFunction.RoundUp(MinValue * 0.8, -2)
            .Axes(xlValue).MaximumScale = WorksheetFunction.RoundUp(MaxValue * 1.1, -2)
            .Axes(xlValue).AxisTitle.Text = "ESM mass [kg]"
            
    End With
End If

Application.ScreenUpdating = True
With Worksheets("Graphical Output").Activate
End With
End Sub

Sub PotHygwaterDiagramm()

Dim Dia_PHW As Chart
Dim frame As ChartObject
Dim rng As Range

Application.ScreenUpdating = False
On Error Resume Next
ThisWorkbook.Worksheets("Graphical Output").ChartObjects("Potable and Hygiene Water Needs").Delete

With Worksheets("Schedule and Metabolic").Activate

    Set rng = Worksheets("Schedule and Metabolic").Range("H3:I50")
    
End With

    Set frame = ThisWorkbook.Worksheets("Graphical Output").ChartObjects.Add(920, 620, 450, 300)
    Set Dia_PHW = frame.Chart
    
With Dia_PHW

    .Parent.Name = "Potable and Hygiene Water Needs"
    .SetSourceData rng
    .PlotArea.Border.ColorIndex = 14
    .ChartType = xlXYScatterSmoothNoMarkers
    .HasLegend = True
    .SeriesCollection(1).XValues = Worksheets("Output Data").Range("H7:H54")
    .SeriesCollection(2).XValues = Worksheets("Output Data").Range("H7:H54")
    .SeriesCollection(1).Values = Worksheets("Schedule and Metabolic").Range("H3:H50")
    .SeriesCollection(1).Name = "=""Potable Water"""
    .SeriesCollection(2).Values = Worksheets("Schedule and Metabolic").Range("I3:I50")
    .SeriesCollection(2).Name = "=""Hygiene Water"""
    .HasTitle = True
    .ChartTitle.Text = "Daily Potable and Hygiene Water Needs"
    .Axes(xlCategory).MajorUnit = 0.125
    .Axes(xlCategory).HasTitle = True
    .Axes(xlCategory).CategoryType = xlTimeScale
    .Axes(xlCategory).MinimumScale = Worksheets("Schedule and Metabolic").Range("A3")
    .Axes(xlCategory).MaximumScale = 1
    .Axes(xlCategory).AxisTitle.Text = "Time"
    .Axes(xlValue).HasTitle = True
    .Axes(xlValue).MinimumScale = 0
    .Axes(xlValue).TickLabels.NumberFormat = "0.00"
    .Axes(xlValue).AxisTitle.Text = "Water Need [kg]"

End With
Application.ScreenUpdating = True
With Worksheets("Graphical Output").Activate
End With
End Sub

Sub CO2ProductionDiagramm()

Dim Dia_CO2P As Chart
Dim frame As ChartObject
Dim rng As Range

Application.ScreenUpdating = False
On Error Resume Next
ThisWorkbook.Worksheets("Graphical Output").ChartObjects("CO2 Production").Delete

With Worksheets("Output Data").Activate

    Set rng = Worksheets("Output Data").Range("I7:I54")
    
End With

    Set frame = ThisWorkbook.Worksheets("Graphical Output").ChartObjects.Add(920, 1, 450, 300)
    Set Dia_CO2P = frame.Chart
    
With Dia_CO2P

    .Parent.Name = "CO2 Production"
    .SetSourceData rng
    .PlotArea.Border.ColorIndex = 14
    .ChartType = xlXYScatterSmoothNoMarkers
    .HasLegend = True
    .SeriesCollection(1).XValues = Worksheets("Output Data").Range("H7:H54")
    .SeriesCollection(1).Values = Worksheets("Output Data").Range("I7:I54")
    .SeriesCollection(1).Name = "=""CO2 Production"""
    .HasTitle = True
    .ChartTitle.Text = "Daily CO2 Production"
    .Axes(xlCategory).MajorUnit = 0.125
    .Axes(xlCategory).HasTitle = True
    .Axes(xlCategory).CategoryType = xlTimeScale
    .Axes(xlCategory).MinimumScale = Worksheets("Schedule and Metabolic").Range("A3")
    .Axes(xlCategory).MaximumScale = 1
    .Axes(xlCategory).AxisTitle.Text = "Time"
    .Axes(xlValue).HasTitle = True
    .Axes(xlValue).MinimumScale = 0
    .Axes(xlValue).AxisTitle.Text = "CO2 [kg]"
    .Axes(xlValue).TickLabels.NumberFormat = "0.00"

End With
With Worksheets("Graphical Output").Activate
End With
Application.ScreenUpdating = True
End Sub

Sub CO2PowerDiagramm()

Dim Dia_CO2Pow As Chart
Dim frame As ChartObject
Dim rng As Range

Application.ScreenUpdating = False

Call Tabelle6.CO2Presssures

On Error Resume Next
ThisWorkbook.Worksheets("Graphical Output").ChartObjects("CO2 Removal Power Requirement").Delete

With Worksheets("Output Data").Activate

    Set rng = Worksheets("Output Data").Range("I7:I54")
    
End With

    Set frame = ThisWorkbook.Worksheets("Graphical Output").ChartObjects.Add(920, 310, 450, 300)
    Set Dia_CO2Pow = frame.Chart
    
With Dia_CO2Pow

    .Parent.Name = "CO2 Removal Power Requirement"
    .SetSourceData rng
    .ChartType = xlXYScatterSmoothNoMarkers
    .HasLegend = True
    .SeriesCollection(1).XValues = Worksheets("Output Data").Range("H7:H54")
    .SeriesCollection(1).Values = Worksheets("Output Data").Range("M7:M54")
    .SeriesCollection(1).Name = "=""Power"""
    .HasTitle = True
    .ChartTitle.Text = "CO2 Removal Power Requirement"
    .Axes(xlCategory).MajorUnit = 0.125
    .Axes(xlCategory).HasTitle = True
    .Axes(xlCategory).CategoryType = xlTimeScale
    .Axes(xlCategory).MinimumScale = Worksheets("Schedule and Metabolic").Range("A3")
    .Axes(xlCategory).MaximumScale = 1
    .Axes(xlCategory).AxisTitle.Text = "Time"
    .Axes(xlValue).HasTitle = True
    .Axes(xlValue).MinimumScale = 0
    .Axes(xlValue).AxisTitle.Text = "Power [W]"
    .Axes(xlValue).TickLabels.NumberFormat = "0.00"

End With
With Worksheets("Graphical Output").Activate
End With
Application.ScreenUpdating = True
End Sub

Sub CO2PPDiagramm()

Dim Dia_CO2PP As Chart
Dim frame As ChartObject
Dim rng As Range

Application.ScreenUpdating = False

On Error Resume Next
ThisWorkbook.Worksheets("Graphical Output").ChartObjects("CO2 Partial Pressure").Delete

With Worksheets("Output Data").Activate

    Set rng = Worksheets("Output Data").Range("R7:R150")
    
End With

    Set frame = ThisWorkbook.Worksheets("Graphical Output").ChartObjects.Add(1380, 1, 600, 300)
    Set Dia_CO2PP = frame.Chart
    
With Dia_CO2PP

    .Parent.Name = "CO2 Partial Pressure"
    .SetSourceData rng
    .PlotArea.Border.ColorIndex = 14
    .ChartType = xlXYScatterSmoothNoMarkers
    .HasLegend = True
    .SeriesCollection(1).Values = Worksheets("Output Data").Range("R7:R150")
    .SeriesCollection(1).XValues = Worksheets("Output Data").Range("O7:O150")
    .SeriesCollection(1).Name = "=""Partial CO2 Pressure"""
    .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
    .HasTitle = True
    .ChartTitle.Text = "CO2 Partial Pressure - 3 Day Scale"
    .Axes(xlCategory).MajorUnit = 10
    .Axes(xlCategory).HasTitle = True
    .Axes(xlCategory).MinimumScale = 0
    .Axes(xlCategory).MaximumScale = 71
    .Axes(xlCategory).TickLabels.NumberFormat = "##"
    .Axes(xlCategory).AxisTitle.Text = "Time [h]"
    .Axes(xlValue).HasTitle = True
    .Axes(xlValue).AxisTitle.Text = "CO2 Partial Pressure [kPa]"
    .Axes(xlValue).TickLabels.NumberFormat = "0.00"

End With
With Worksheets("Graphical Output").Activate
End With
Application.ScreenUpdating = True
End Sub

Sub ECLSSComposition()

Dim Dia_ECLSSComp As Chart
Dim frame As ChartObject
Dim rng As Range

Application.ScreenUpdating = False

On Error Resume Next
ThisWorkbook.Worksheets("Graphical Output").ChartObjects("ECLSS ESM Composition").Delete

With Worksheets("Output Data").Activate

If Worksheets("Graphical Output").Range("C33") <> "-" Then Set rng = Worksheets("Graphical Output").Range("A29:C33")
If Worksheets("Graphical Output").Range("C33") = "-" Then Set rng = Worksheets("Graphical Output").Range("A29:C32")
If Worksheets("ECLSS Composition").Range("B1") = "biological" Then Set rng = Worksheets("Graphical Output").Range("A28:C28,A33:C33")
    
End With

Set frame = ThisWorkbook.Worksheets("Graphical Output").ChartObjects.Add(460, 310, 450, 300)
    Set Dia_ECLSSComp = frame.Chart
    
With Dia_ECLSSComp
    
    .Parent.Name = "ECLSS ESM Composition"
    .SetSourceData rng
    .ChartType = xlPie
    If Worksheets("ECLSS Composition").Range("B1") = "biological" Then
    .HasLegend = True
    .SeriesCollection(1).XValues = Worksheets("Graphical Output").Range("A28,A33")
    .SeriesCollection(2).XValues = Worksheets("Graphical Output").Range("C28:C33")
    End If
    If Worksheets("ECLSS Composition").Range("B1") <> "biological" Then
    .HasLegend = True
    .SeriesCollection(1).XValues = Worksheets("Graphical Output").Range("A29:A33")
    .SeriesCollection(2).XValues = Worksheets("Graphical Output").Range("C29:C33")
    End If
    .HasTitle = True
    .ChartTitle.Text = "ECLSS ESM Composition"
    
End With
Application.ScreenUpdating = True
End Sub

Sub O2ConDiagramm()

Dim Dia_O2C As Chart
Dim frame As ChartObject
Dim rng As Range
Dim Schedule

Application.ScreenUpdating = False
On Error Resume Next
ThisWorkbook.Worksheets("Graphical Output").ChartObjects("O2 Needs").Delete
Schedule = Worksheets("User Interface").Range("Schedule_Number").Value

With Worksheets("Schedule and Metabolic").Activate

    If Schedule = 1 Then Set rng = Worksheets("Schedule and Metabolic").Range("D3:D50")
    If Schedule = 2 Then Set rng = Worksheets("Schedule and Metabolic").Range("B128:B175")
    
End With

    Set frame = ThisWorkbook.Worksheets("Graphical Output").ChartObjects.Add(920, 310, 450, 300)
    Set Dia_O2C = frame.Chart
    
With Dia_O2C

    .Parent.Name = "O2 Needs"
    .SetSourceData rng
    .PlotArea.Border.ColorIndex = 14
    .ChartType = xlXYScatterSmoothNoMarkers
    .HasLegend = True
    .SeriesCollection(1).XValues = Worksheets("Output Data").Range("H7:H54")
    .SeriesCollection(1).Values = rng
    .SeriesCollection(1).Name = "=""O2 Consumption"""
    .HasTitle = True
    .ChartTitle.Text = "Daily O2 Consumption"
    .Axes(xlCategory).MajorUnit = 0.125
    .Axes(xlCategory).HasTitle = True
    .Axes(xlCategory).CategoryType = xlTimeScale
    .Axes(xlCategory).MinimumScale = Worksheets("Schedule and Metabolic").Range("A3")
    .Axes(xlCategory).MaximumScale = 1
    .Axes(xlCategory).AxisTitle.Text = "Time"
    .Axes(xlValue).HasTitle = True
    .Axes(xlValue).MinimumScale = 0
    .Axes(xlValue).TickLabels.NumberFormat = "0.00"
    .Axes(xlValue).AxisTitle.Text = "O2 Need [kg]"

End With
With Worksheets("Graphical Output").Activate
End With
Application.ScreenUpdating = True
End Sub

Sub O2PPDiagramm()

Dim Dia_O2PP As Chart
Dim frame As ChartObject
Dim rng As Range

Application.ScreenUpdating = False
On Error Resume Next
ThisWorkbook.Worksheets("Graphical Output").ChartObjects("O2 Partial Pressure").Delete

With Worksheets("Output Data").Activate

    Set rng = Worksheets("Output Data").Range("X7:X150")
    
End With

    Set frame = ThisWorkbook.Worksheets("Graphical Output").ChartObjects.Add(1380, 310, 600, 300)
    Set Dia_O2PP = frame.Chart
    
With Dia_O2PP

    .Parent.Name = "O2 Partial Pressure"
    .SetSourceData rng
    .PlotArea.Border.ColorIndex = 14
    .ChartType = xlXYScatterSmoothNoMarkers
    .HasLegend = True
    .SeriesCollection(1).Values = Worksheets("Output Data").Range("X7:X150")
    .SeriesCollection(1).XValues = Worksheets("Output Data").Range("U7:U150")
    .SeriesCollection(1).Name = "=""Partial CO2 Pressure"""
    .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
    .HasTitle = True
    .ChartTitle.Text = "O2 Partial Pressure - 3 Day Scale"
    .Axes(xlCategory).MajorUnit = 10
    .Axes(xlCategory).HasTitle = True
    .Axes(xlCategory).MinimumScale = 0
    .Axes(xlCategory).MaximumScale = 71
    .Axes(xlCategory).TickLabels.NumberFormat = "##"
    .Axes(xlCategory).AxisTitle.Text = "Time [h]"
    .Axes(xlValue).HasTitle = True
    .Axes(xlValue).AxisTitle.Text = "O2 Partial Pressure [kPa]"
    .Axes(xlValue).TickLabels.NumberFormat = "0.00"

End With
With Worksheets("Graphical Output").Activate
End With
Application.ScreenUpdating = True
End Sub

Sub WWurineDiagramm()

Dim Dia_WWU As Chart
Dim frame As ChartObject
Dim rng As Range

Application.ScreenUpdating = False
On Error Resume Next
ThisWorkbook.Worksheets("Graphical Output").ChartObjects("Daily WW and Urine Production").Delete

With Worksheets("Schedule and Metabolic").Activate

    Set rng = Worksheets("Output Data").Range("AB7:AD54")
    
End With

    Set frame = ThisWorkbook.Worksheets("Graphical Output").ChartObjects.Add(460, 620, 450, 300)
    Set Dia_WWU = frame.Chart
    
With Dia_WWU

    .Parent.Name = "Daily WW and Urine Production"
    .SetSourceData rng
    .PlotArea.Border.ColorIndex = 14
    .ChartType = xlXYScatterSmoothNoMarkers
    .HasLegend = True
    .SeriesCollection(1).XValues = Worksheets("Output Data").Range("H7:H54")
    .SeriesCollection(2).XValues = Worksheets("Output Data").Range("H7:H54")
    .SeriesCollection(3).XValues = Worksheets("Output Data").Range("H7:H54")
    .SeriesCollection(1).Values = Worksheets("Output Data").Range("AB7:AB54")
    .SeriesCollection(1).Name = "=""Sweat Prodcution"""
    .SeriesCollection(2).Values = Worksheets("Output Data").Range("AC7:AC54")
    .SeriesCollection(2).Name = "=""Urine Production"""
    .SeriesCollection(3).Values = Worksheets("Output Data").Range("AD7:AD54")
    .SeriesCollection(3).Name = "=""Hygiene WW Production"""
    .HasTitle = True
    .ChartTitle.Text = "Daily WW and Urine Production"
    .Axes(xlCategory).MajorUnit = 0.125
    .Axes(xlCategory).HasTitle = True
    .Axes(xlCategory).CategoryType = xlTimeScale
    .Axes(xlCategory).MinimumScale = Worksheets("Schedule and Metabolic").Range("A3")
    .Axes(xlCategory).MaximumScale = 1
    .Axes(xlCategory).AxisTitle.Text = "Time"
    .Axes(xlValue).HasTitle = True
    .Axes(xlValue).MinimumScale = 0
    .Axes(xlValue).TickLabels.NumberFormat = "0.00"
    .Axes(xlValue).AxisTitle.Text = "WW and urine Production [kg]"

End With
With Worksheets("Graphical Output").Activate
End With
Application.ScreenUpdating = True
End Sub

Sub PowWWurineDiagramm()

Dim Dia_PWWU As Chart
Dim frame As ChartObject
Dim rng As Range

Application.ScreenUpdating = False
On Error Resume Next
ThisWorkbook.Worksheets("Graphical Output").ChartObjects("Daily WRM Power").Delete

With Worksheets("Schedule and Metabolic").Activate

    Set rng = Worksheets("Output Data").Range("AG7:AJ54")
    
End With

    Set frame = ThisWorkbook.Worksheets("Graphical Output").ChartObjects.Add(1380, 620, 600, 300)
    Set Dia_PWWU = frame.Chart
    
With Dia_PWWU

    .Parent.Name = "Daily WRM Power"
    .SetSourceData rng
    .PlotArea.Border.ColorIndex = 14
    .ChartType = xlXYScatterSmoothNoMarkers
    .HasLegend = True
    .SeriesCollection(1).XValues = Worksheets("Output Data").Range("H7:H54")
    .SeriesCollection(2).XValues = Worksheets("Output Data").Range("H7:H54")
    .SeriesCollection(3).XValues = Worksheets("Output Data").Range("H7:H54")
    .SeriesCollection(4).XValues = Worksheets("Output Data").Range("H7:H54")
    .SeriesCollection(1).Values = Worksheets("Output Data").Range("AG7:AG54")
    .SeriesCollection(1).Name = "=""Power: Urine Processing"""
    .SeriesCollection(2).Values = Worksheets("Output Data").Range("AH7:AH54")
    .SeriesCollection(2).Name = "=""Power: Wastewater"""
    .SeriesCollection(3).Values = Worksheets("Output Data").Range("AI7:AI54")
    .SeriesCollection(3).Name = "=""Power: Heating"""
    .SeriesCollection(4).Values = Worksheets("Output Data").Range("AJ7:AJ54")
    .SeriesCollection(4).Name = "=""Power: Total"""
    .HasTitle = True
    .ChartTitle.Text = "Daily WRM Power Needs"
    .Axes(xlCategory).MajorUnit = 0.125
    .Axes(xlCategory).HasTitle = True
    .Axes(xlCategory).CategoryType = xlTimeScale
    .Axes(xlCategory).MinimumScale = Worksheets("Schedule and Metabolic").Range("A3")
    .Axes(xlCategory).MaximumScale = 1
    .Axes(xlCategory).AxisTitle.Text = "Time"
    .Axes(xlValue).HasTitle = True
    .Axes(xlValue).MinimumScale = 0
    .Axes(xlValue).TickLabels.NumberFormat = "0.00"
    .Axes(xlValue).AxisTitle.Text = "Power [W]"

End With
With Worksheets("Graphical Output").Activate
End With
Application.ScreenUpdating = True
End Sub

Sub critattribute_plot()

Dim Dia_CA As Chart
Dim frame As ChartObject
Dim rng As Range

Application.ScreenUpdating = False
On Error Resume Next
ThisWorkbook.Worksheets("Graphical Output").ChartObjects("Relative Attribute Criticality").Delete

With Worksheets("Output Data").Activate

    Set rng = Worksheets("Output Data").Range("K70:L76")
    
End With

    Set frame = ThisWorkbook.Worksheets("Graphical Output").ChartObjects.Add(460, 930, 450, 300)
    Set Dia_CA = frame.Chart
    
With Dia_CA

    .Parent.Name = "Relative Attribute Criticality"
    .SetSourceData rng
    .PlotArea.Border.ColorIndex = 14
    .ChartType = xlColumnClustered
    .HasTitle = True
    .HasLegend = False
    .ChartTitle.Text = "Relative Attribute Criticality"
    .Axes(xlCategory).HasTitle = True
    .Axes(xlCategory).AxisTitle.Text = "Attribute"
    .Axes(xlValue).HasTitle = True
    .Axes(xlValue).AxisTitle.Text = "Relative Change [%]"

End With
With Worksheets("Graphical Output").Activate
End With
Application.ScreenUpdating = True
End Sub
