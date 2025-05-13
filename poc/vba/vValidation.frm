VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vValidation 
   Caption         =   "Validation"
   ClientHeight    =   5412
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5760
   OleObjectBlob   =   "vValidation.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "vValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'View: Validation
'Description: Sets of data validation procedures.


Public Event TestCompleteness()
Public Event TestDocumentBalance()
Public Event TestRDE()
Public Event DoAccountMapping()
Public Event DoConfirm()
Public Event DoExit()


Private Sub UserForm_Initialize()
    Me.lblInfo.Caption = "����:" & vbCrLf & _
        "1. ��������ʴ��աA�~��i��ɶU�������դ� Account Mapping�C" & vbCrLf & _
        "2. �����ɶU�������աA�~��i��i�a�ʴ��աC"
    Me.Show (vbModeless)
End Sub

Private Sub ButtonCompleteness_Click()
    RaiseEvent TestCompleteness
End Sub

Private Sub ButtonBalance_Click()
    RaiseEvent TestDocumentBalance
End Sub

Private Sub ButtonRDE_Click()
    RaiseEvent TestRDE
End Sub

Private Sub ButtonAccountMapping_Click()
    RaiseEvent DoAccountMapping
End Sub

Private Sub ButtonConfim_Click()
    RaiseEvent DoConfirm
End Sub

Private Sub ButtonExit_Click()
    RaiseEvent DoExit
End Sub
