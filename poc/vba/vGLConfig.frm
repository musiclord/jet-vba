VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vGLConfig 
   Caption         =   "GL Configuration"
   ClientHeight    =   8415.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   OleObjectBlob   =   "vGLConfig.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "vGLConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===== ����T =====
'�W��: vGLConfig
'�y�z: �]�w�`�b (GL) ��ƶפJ���������M�����ﶵ�C
'
'===== ��汱� =====
'�U�Ԧ���� (ComboBox):
'  - cboDocumentNo: �����ӷ��ɤ����u�ǲ����X�v���
'  - cboLineItem: �����ӷ��ɤ����u�ǲ������v���
'  - cbotPostDate: �����ӷ��ɤ����u�`�b����v���
'  - cboAccountNo: �����ӷ��ɤ����u�|�p��ؽs���v���
'  - cboAccountName: �����ӷ��ɤ����u�|�p��ئW�١v���
'  - cboDescription: �����ӷ��ɤ����u�ǲ��K�n�v���
'  - cboMethod: ����`�b���B���B�z�覡 (�Ҧp�G��@���B���B�ɶU���}��쵥)
'  - cboAmount: �����ӷ��ɤ����u�ǲ����B�v��� (�Y cboMethod ��ܳ�@���B)
'  - cboDebit: �����ӷ��ɤ����u�ɤ���B�v��� (�Y cboMethod ��ܭɶU���})
'  - cboCredit: �����ӷ��ɤ����u�U����B�v��� (�Y cboMethod ��ܭɶU���})
'  - cboDrCr: �����ӷ��ɤ����u�ɶU���O�v��� (�Y cboMethod ��ܭɶU�аO)
'  - cboDebitFlag: ��J�N��ɤ誺�аO�� (�Y cboMethod ��ܭɶU�аO)
'  - cboPostingStatus: �����ӷ��ɤ����u�L�b���A�v���
'  - cboPostedCode: ��J�N��w�L�b�����A�N�X
'  - cboPosted: ��ܬO�_�ȿz��w�L�b����� (�q�`�O CheckBox �� OptionButton�A�Y�� ComboBox �h�i�ണ�� Yes/No �ﶵ)
'
'�R�O���s (CommandButton):
'  - btnImport: Ĳ�o�פJ GL ��ƪ��ƥ� (RaiseEvent DoImport)
'  - btnPreview: Ĳ�o�w�� GL ��ƪ��ƥ� (RaiseEvent DoPreview)
'  - btnConfirm: �T�{�]�w��Ĳ�o DoConfirm �ƥ� (RaiseEvent DoConfirm)
'  - btnExit: ��������Ĳ�o DoExit �ƥ� (RaiseEvent DoExit)

Public Event DoImport()
Public Event DoPreview()
Public Event DoConfirm()
Public Event DoExit()

Private Sub btnConfirm_Click()
    RaiseEvent DoConfirm
End Sub

Private Sub btnExit_Click()
    RaiseEvent DoExit
End Sub

Private Sub btnImport_Click()
    RaiseEvent DoImport
End Sub

Private Sub btnPreview_Click()
    RaiseEvent DoPreview
End Sub

Private Sub cboCredit_Change()

End Sub

Private Sub lblDocumentNo_Click()

End Sub

Private Sub UserForm_Initialize()

End Sub

