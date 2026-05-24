VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UG_Plants 
   Caption         =   "UserForm1"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11805
   OleObjectBlob   =   "UG_Plants.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UG_Plants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_Cancel_Click()

    Unload UG_ResupplyModelling

End Sub

Private Sub CommandButton_Back_Click()

    Unload UG_ResupplyModelling
    UG_Handbook.Width = 500
    UG_Handbook.Height = 550
    UG_Handbook.Show

End Sub

Private Sub Label2_Click()

End Sub
