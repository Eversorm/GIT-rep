VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UG_MCAESM_2 
   Caption         =   "LiSTOT User's Guide - MCA/ESM Part 2"
   ClientHeight    =   10095
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9390.001
   OleObjectBlob   =   "UG_MCAESM_2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UG_MCAESM_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_MCAESM_Click()

    Unload UG_MCAESM_2
    UG_MCAESM_3.Width = 500
    UG_MCAESM_3.Height = 550
    UG_MCAESM_3.Show

End Sub

Private Sub CommandButton2_Click()

    Unload UG_MCAESM_2

End Sub

Private Sub CommandButton3_Click()

    Unload UG_MCAESM_2
    UG_MCAESM_1.Width = 500
    UG_MCAESM_1.Height = 550
    UG_MCAESM_1.Show

End Sub

Private Sub UserForm_Click()

End Sub
