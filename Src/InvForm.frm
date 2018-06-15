VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InvForm 
   Caption         =   "Inventory"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3975
   OleObjectBlob   =   "InvForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InvForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancelbutton_Click()
    Unload Me
End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub okbutton_Click()
        
    Call InventoryProcessor
    monthval = Month(DateValue("01-" & monthbox.Value & "-1900"))
    
    If newbutton.Value Then
        invtype = ";n"
    Else
        invtype = ";u"
    End If
    
        Sheet2.Range("b1").Value = monthval & invtype
    
    If invholdingbox Then
        'If Not Sheet3.Range("A1").Value = "" Then Sheet2.Rows(1).Delete
        Call invholding
    Else
        Call inventorywriter("Paste Data Here")
        Sheet3.Cells.Delete
    End If
    
    Sheet2.Cells.Clear
    
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()
    With monthbox
        .AddItem "January"
        .AddItem "February"
        .AddItem "March"
        .AddItem "April"
        .AddItem "May"
        .AddItem "June"
        .AddItem "July"
        .AddItem "August"
        .AddItem "September"
        .AddItem "October"
        .AddItem "November"
        .AddItem "December"
    End With

    monthbox.Value = "January"
    newbutton.Value = True
    
End Sub
