VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Iniatialize the ECLSS composition scheme every time the file opens
Private Sub Workbook_Open()
    Tabelle4.CreateDropDownLists
End Sub

' Save all the VBA codes files in 'src' folder
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    Dim comp As Object
    Dim exportPath As String
    Dim fileName As String
    
    ' If the save operation was not successful, exit the procedure
    If Not Success Then Exit Sub
    
    ' Define the destination folder (subfolder "src" within the workbook's path)
    exportPath = ThisWorkbook.Path & "\src\"
    
    ' Create the "src" directory if it does not already exist
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    
    ' Loop through all components in the VBA project (modules, classes, forms)
    For Each comp In ThisWorkbook.VBProject.VBComponents
        ' Export only if the component contains code (avoids empty sheets)
        If comp.CodeModule.CountOfLines > 0 Then
            Select Case comp.Type
                Case 1 ' Standard Module
                    fileName = comp.Name & ".bas"
                Case 2 ' Class Module
                    fileName = comp.Name & ".cls"
                Case 3 ' Microsoft UserForm
                    fileName = comp.Name & ".frm"
                Case 100 ' Document objects (ThisWorkbook or Worksheets)
                    fileName = comp.Name & ".cls"
                Case Else
                    fileName = ""
            End Select
            
            ' If a valid filename was assigned, export the component
            If fileName <> "" Then
                comp.Export exportPath & fileName
            End If
        End If
    Next comp
    
    ' Note: The .xlsm file itself remains saved in the main repository folder
End Sub
