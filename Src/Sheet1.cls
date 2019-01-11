VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    runerror = ExxonProcessor
    If pushbox And runerror <> 0 Then Call holding
End Sub
Private Sub CommandButton2_Click()
    Call FuelmanProcessor("FUELMAN")
    If pushbox Then Call holding
End Sub
Private Sub CommandButton3_Click()
    runerror = ChaseProcessor
    If pushbox And runerror <> 0 Then Call holding
End Sub
Private Sub commandbutton4_click()
    InvForm.Show
End Sub
Private Sub CommandButton5_Click()
    Call FuelmanProcessor("CHEVRON")
    If pushbox Then Call holding
End Sub

Private Sub fuelpushbutton_Click()
    Call FuelWriter
    MsgBox ("All Done!")
End Sub

Private Sub inventorypushbutton_Click()
    Call inventorywriter("Holding Data")
    Sheet3.Cells.Delete
End Sub
