VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UG_MCAESM_3 
   Caption         =   "LiSTOT User's Guide - MCA/ESM Part 3"
   ClientHeight    =   9480.001
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9390.001
   OleObjectBlob   =   "UG_MCAESM_3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UG_MCAESM_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_MCAESM_Click()

    Unload UG_MCAESM_3
    UG_MCAESM_4.Width = 500
    UG_MCAESM_4.Height = 550
    UG_MCAESM_4.Show

End Sub

Private Sub CommandButton2_Click()

    Unload UG_MCAESM_3

End Sub

Private Sub CommandButton3_Click()

    Unload UG_MCAESM_3
    UG_MCAESM_2.Width = 500
    UG_MCAESM_2.Height = 550
    UG_MCAESM_2.Show

End Sub


