VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vProject 
   Caption         =   "Project Metadata"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "vProject.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "vProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===== ����T =====
'�W��: vProject
'�y�z: �]�w�M�פ��ƾڡA�]�t�Ȥ��T�M�|�p�����C
'
'===== ��汱� =====
'��r���:
'  - txtClientName: �Ȥ�W�ٿ�J���
'  - txtbPeriodStart: �|�p�����}�l���
'  - txtbPeriodEnd: �|�p�����������
'  - txtbPeriodLast: �W�@�|�p�������
'
'�R�O���s:
'  - ButtonConfirm: �T�{���s�A�I����Ĳ�o DoConfirm �ƥ�
'  - ButtonExit: �h�X���s�A�I����Ĳ�o DoExit �ƥ�

Public Event DoConfirm()
Public Event DoExit()

Private Sub UserForm_Initialize()

End Sub

Private Sub ButtonConfirm_Click()
    RaiseEvent DoConfirm
End Sub

Private Sub ButtonExit_Click()
    RaiseEvent DoExit
End Sub
