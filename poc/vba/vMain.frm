VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vMain 
   Caption         =   "Main"
   ClientHeight    =   8685.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5550
   OleObjectBlob   =   "vMain.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "vMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'View: Main
'Description: Main user form to interact with functions.

Option Explicit
Public Event DoExit()
Public Event DoImportGL()
Public Event DoImportTB()
Public Event DoMapping()
Public Event DoPreview()
Public Event GetTable()

Private Sub ButtonExit_Click()
    RaiseEvent DoExit
End Sub

Private Sub ButtonImportGL_Click()
    RaiseEvent DoImportGL
End Sub

Private Sub ButtonImportTB_Click()
    RaiseEvent DoImportTB
End Sub

Private Sub ButtonMap_Click()
    RaiseEvent DoMapping
End Sub

Private Sub ButtonPreview_Click()
    ' Load selected table from LisTable(comboBox) to worksheet "Preview" as view.
    RaiseEvent DoPreview
End Sub

Private Sub ListTable_DropButtonClick()
    ' Accquire all table names from available databases.
    RaiseEvent GetTable
End Sub

Private Sub UserForm_Initialize()
    Me.Show (vbModeless)
End Sub
