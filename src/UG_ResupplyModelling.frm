VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UG_ResupplyModelling 
   Caption         =   "LiSTOT User's Guide - Resupply Modelling"
   ClientHeight    =   2688
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   6000
   OleObjectBlob   =   "UG_ResupplyModelling.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UG_ResupplyModelling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_Cancel_Click()

    Unload UG_ResupplyModelling

End Sub

Private Sub CommandButton2_Click()

    Unload UG_ResupplyModelling
    UG_ECLSSComposition_3.Widht = 500
    UG_ECLSSComposition_3.Height = 550
    UG_ECLSSComposition_3.Show

End Sub
