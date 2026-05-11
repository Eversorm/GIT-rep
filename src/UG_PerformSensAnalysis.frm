VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UG_PerformSensAnalysis 
   Caption         =   "LiSTOT User's Guide - Sensitivity Analysis"
   ClientHeight    =   5112
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9390.001
   OleObjectBlob   =   "UG_PerformSensAnalysis.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UG_PerformSensAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_Cancel_Click()

    Unload UG_PerformSensAnalysis

End Sub

Private Sub CommandButton_MCAESM_Click()

     Unload UG_PerformSensAnalysis
     UG_ECLSSComposition_1.Width = 500
     UG_ECLSSComposition_1.Height = 550
     UG_ECLSSComposition_1.Show

End Sub

Private Sub CommandButton1_Click()

     Unload UG_PerformSensAnalysis
     UG_MCAESM_4.Width = 500
     UG_MCAESM_4.Height = 550
     UG_MCAESM_4.Show

End Sub
