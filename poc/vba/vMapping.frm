VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vMapping 
   Caption         =   "Mapping"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7620
   OleObjectBlob   =   "vMapping.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "vMApping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'View: Mapping
'Description: 映射欄位介面，讓使用者操作下拉式選單來映射標準的欄位名稱。

Public Event DoExit()
Public Event DoConfirm()


Private Sub ButtonConfirm_Click()
    RaiseEvent DoConfirm
End Sub

Private Sub ButtonExit_Click()
    RaiseEvent DoExit
End Sub

Private Sub UserForm_Initialize()
    Me.Show (vbModeless)
    
    ' Gets GL and TB table columns as values
    
    ' Set values for each ComboBox, correspond to GLEntity and TBEntity
    
End Sub

