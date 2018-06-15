VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FuelPusher 
   Caption         =   "UserForm1"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5880
   OleObjectBlob   =   "FuelPusher.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FuelPusher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub appendbutton_Click()
    Me.Tag = 2
    Me.Hide
End Sub

Private Sub cancelbutton_Click()
    Me.Tag = 0
    Me.Hide
End Sub

Private Sub replacebutton_Click()
    Me.Tag = 1
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    Label1.Caption = "We found data for this month already in the fuel analyzer. What should be done with the data?"
    Me.Caption = "Conflicting Data Found!"
End Sub
