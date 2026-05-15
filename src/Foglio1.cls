VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Foglio1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


' get the plants properties for the "ECLSS composition" calculations
Public Function GetPlantProperties() As Variant
    ' read plants table and gives a 2x2 array with the corrispective properties (carbon content, O2 production...)
    With ThisWorkbook.Worksheets("Plant and PGC")
    'Dim ws As Worksheet
    'Set ws = Worksheets("Plant and PGC")
    Dim lastRow As Long
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).row
    GetPlantProperties = .Range("C6:K" & lastRow).Value ' A = plant name, B to I= properties
End Function

' get the areas of every plant for the specific PGC, used for the "ECLSS composition" calculations
Public Function GetPGCAreasByName(pgcName As String) As Variant
    Dim ws As Worksheet
    Set ws = Worksheets("Plants and PGC")
    ' cerca il nome PGC nella colonna della tabella configurazioni
    ' restituisce il vettore di aree corrispondente
    Dim configRange As Range
    Set configRange = ws.Range("PGC_Configs") ' Named range sulla tabella configurazioni
    Dim i As Long
    For i = 1 To configRange.Rows.Count
        If InStr(1, configRange.Cells(i, 1).Value, pgcName, vbTextCompare) > 0 Then
            GetPGCAreasByName = Application.Transpose( _
                configRange.Rows(i).Offset(0, 1).Resize(1, 24).Value)
            Exit Function
        End If
    Next i
    ' fallback: array di zeri
    GetPGCAreasByName = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
End Function

