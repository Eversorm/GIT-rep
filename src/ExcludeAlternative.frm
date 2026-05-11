VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExcludeAlternative 
   Caption         =   "Exclude Alternative"
   ClientHeight    =   2592
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   7500
   OleObjectBlob   =   "ExcludeAlternative.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExcludeAlternative"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_Cancel_Click()

    Unload ExcludeAlternative

End Sub

Private Sub CommandButton_Exclude_Click()
' this code is executed after the user clicks the button exclude alternative. In that case we have to first delete the selected alternative from the MCA sheet and then reporfm the MCA
Dim i As Integer
Dim na As Integer
Dim iHeaderRow As Integer
Dim miColumns
Dim iColumn As Integer

na = Worksheets("MCA ESM").Range("A1:A20").Cells.SpecialCells(xlCellTypeConstants).Count

iHeaderRow = 4
miColumns = findColumnByName(Array("Name of Assembly"), "MCA ESM", iHeaderRow)
iColumn = miColumns(1)

For i = 6 To na + 4

    If Worksheets("MCA ESM").Cells(i, iColumn).Value = ExcludeAlternative.ComboBox_ExclAlt.Value Then
        Worksheets("MCA ESM").Cells(i, iColumn).EntireRow.Delete
    End If
    
Next i

' repeform the MCA using the corresponding sub for it
Tabelle5.ExecuteMCA

Unload ExcludeAlternative

End Sub

Private Sub UserForm_Initialize()

Dim na As Integer
Dim i As Integer
Dim iHeaderRow As Integer
Dim miColumns

'On Error GoTo EH
na = Worksheets("MCA ESM").Range("A6:A20").Cells.SpecialCells(xlCellTypeConstants).Count

iHeaderRow = 4
miColumns = findColumnByName(Array("Name of Assembly"), "MCA ESM", iHeaderRow)

With Worksheets("MCA ESM").Activate

    With ExcludeAlternative.ComboBox_ExclAlt
    
        For i = 6 To na + 5
    
            .AddItem Worksheets("MCA ESM").Cells(i, miColumns(1)).Value
    
        Next
        
    End With
End With

End Sub
