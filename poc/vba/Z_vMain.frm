VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Z_vMain 
   Caption         =   "Main"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5550
   OleObjectBlob   =   "Z_vMain.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "Z_vMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'View: Main
'�D�{�������A���ϥΪ̰��� :1.�פJ�ɮ�,2.���Ҹ��,3.�z�����,4.��X���i
'�ó]�p�w���\��A��ܨé�u�@���˵���ƪ��e1000����ơC

Public Event DoExit()
Public Event DoImportGL()
Public Event DoImportTB()
Public Event DoPreview()
Public Event OpenMapping()
Public Event GetTableNames()

Private Sub UserForm_Initialize()
    Me.Show (vbModeless)
End Sub

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
    RaiseEvent OpenMapping
End Sub

Private Sub ButtonPreview_Click()
    ' �N ListTable�襤����ƪ���J�e1000����Ʀܹw���u�@��
    RaiseEvent DoPreview
End Sub

Private Sub ListTable_Enter()
    ' �� ComboBox ��o�J�I��Ĳ�o�ƥ�
    RaiseEvent GetTableNames
End Sub
