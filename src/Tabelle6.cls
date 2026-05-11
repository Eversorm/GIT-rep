VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub findBreakEvenPoints()

Dim iResult As Integer
Dim iResult2 As Integer
Dim iHeaderRow As Integer
Dim StartValues(10) As Double
Dim SecondValues(10) As Double
Dim Slope As Double
Dim Slope2 As Double
Dim Intersection As Double
Dim ChtRng

With Worksheets("Output Data")

Set ChtRng = .Range(.Cells(77, 3), .Cells(86, 12))
ChtRng.ClearContents

iHeaderRow = 92


For iResult = 1 To 10
    StartValues(iResult) = 0
    SecondValues(iResult) = 0
Next iResult


For iResult = 1 To 10
    If Not IsEmpty(.Cells(iHeaderRow + 1, 2 + iResult).Value) Then
        StartValues(iResult) = .Cells(iHeaderRow + 2, 2 + iResult).Value
        SecondValues(iResult) = .Cells(iHeaderRow + 3, 2 + iResult).Value
    End If
Next iResult


For iResult = 1 To 10
    If Not IsEmpty(.Cells(iHeaderRow + 1, 2 + iResult).Value) Then
        For iResult2 = 1 To 10
            If Not IsEmpty(.Cells(iHeaderRow + 1, 2 + iResult2).Value) Then
            
                ' Get intersection
                ' m*x+t = m2*x+t2
                ' x = (t2 - t) / (m - m2)
                Slope = (SecondValues(iResult) - StartValues(iResult))
                Slope2 = (SecondValues(iResult2) - StartValues(iResult2))
                If Slope <> Slope2 Then
                
                    Intersection = (StartValues(iResult2) - StartValues(iResult)) / (Slope - Slope2)
                    
                    If Intersection > 0 Then
                    ' only add the break even, if this system becomes better than the other system
                        If Slope < Slope2 Then
                            .Cells(76 + iResult2, 2 + iResult).Value = Intersection
                        End If
                    Else
                        If Slope < Slope2 Then
                            .Cells(76 + iResult2, 2 + iResult).Value = 1
                        Else
                            .Cells(76 + iResult2, 2 + iResult).Value = -1
                        End If
                    End If
                End If
            End If
        Next iResult2
    End If
Next iResult

End With
End Sub
