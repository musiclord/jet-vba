VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vCriteria 
   Caption         =   "Criteria"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "vCriteria.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "vCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Event DoExit()
Public Event DoConfirm()

Private Sub UserForm_Initialize()
    Me.Show (vbModeless)
End Sub
Private Sub btnConfirm_Click()
    RaiseEvent DoConfirm
End Sub

Private Sub btnExit_Click()
    RaiseEvent DoExit
End Sub
