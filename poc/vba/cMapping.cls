VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'Controller: Mapping
'Description: Controller of vMapping.

Private WithEvents vMapping As vMapping
Attribute vMapping.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set vMapping = New vMapping
End Sub

Private Sub vMapping_DoConfirm()
    ' When user confirm all selected name for each columns
    ' will raise this event, and process Mapping Service
    ' To map and set values for GL and TB entity attributes
End Sub

Private Sub vMapping_DoExit()
    Call vMapping.Hide
    Unload vMapping
    Set vMapping = Nothing
End Sub
