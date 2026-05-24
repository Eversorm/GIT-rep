VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UG_Handbook 
   Caption         =   "LiSTOT User's Guide"
   ClientHeight    =   10200
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   9210.001
   OleObjectBlob   =   "UG_Handbook.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UG_Handbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_MD_Click()

    Unload UG_Handbook
    UG_MissionDefinition.Width = 500
    UG_MissionDefinition.Height = 550
    UG_MissionDefinition.Show

End Sub

Private Sub CommandButton1_Click()

    Unload UG_Handbook

End Sub

Private Sub CommandButton2_Click()
    
    Unload UG_Handbook
    UG_MCAESM_1.Width = 500
    UG_MCAESM_1.Height = 550
    UG_MCAESM_1.Show

End Sub

Private Sub CommandButton3_Click()

    Unload UG_Handbook
    UG_PerformSensAnalysis.Width = 500
    UG_PerformSensAnalysis.Height = 550
    UG_PerformSensAnalysis.Show

End Sub

Private Sub CommandButton4_Click()

    Unload UG_Handbook
    UG_ECLSSComposition.Width = 500
    UG_ECLSSComposition.Height = 550
    UG_ECLSSComposition.Show
    
End Sub

Private Sub CommandButton5_Click()

    Unload UG_Handbook
    UG_ResupplyModelling.Width = 500
    UG_ResupplyModelling.Height = 350
    UG_ResupplyModelling.Show
    
End Sub
