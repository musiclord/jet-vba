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
'��r��J�� (TextBox):
'  - txtbPosted: ��ܬO�_�ȿz��w�L�b����� (�q�`�O CheckBox �� OptionButton�A�Y�� ComboBox �h�i�ണ�� Yes/No �ﶵ)
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

Private Sub cboDescription_Change()

End Sub

' --- ��l�ƨƥ� (�M�ũβ��������J�޿�) ---
Private Sub UserForm_Initialize()
    ' �T�O���B�����������J
    Debug.Print Me.Name & " Initialized."
End Sub
' --- �s�W: ���@��k�A�Ω���o�ϥΪ̳]�w�������� ---
Public Function GetFieldMappings() As Object ' Returns Scripting.Dictionary
    Const METHOD_NAME As String = "GetFieldMappings"
    Dim mappings As Object ' Scripting.Dictionary
    Dim ctrl As MSForms.Control
    Dim cbo As MSForms.ComboBox
    Dim comboBoxTag As String
    Dim standardField As String ' ComboBox �� Name�A�@���з����W (Key)
    Dim selectedField As String ' ComboBox �襤�� Value�A�@���ӷ����W (Value)

    On Error GoTo ErrorHandler

    ' �إ� Dictionary ����
    Set mappings = CreateObject("Scripting.Dictionary")
    mappings.CompareMode = vbTextCompare ' Key ���Ϥ��j�p�g
    Debug.Print Me.Name & "." & METHOD_NAME & " - Dictionary created."

    ' �M�����W���Ҧ����
    For Each ctrl In Me.Controls
        ' �ˬd�O�_�� ComboBox
        If TypeOf ctrl Is MSForms.ComboBox Then
            Set cbo = ctrl
            comboBoxTag = Trim(cbo.tag)
            ' �p�G Tag �ݩʳQ�аO�� "header"�A�h�B�z�� ComboBox
            If comboBoxTag = "header" Then
                If cbo.ListIndex > -1 Then          ' �T�O�����سQ�襤
                    standardField = cbo.Name        ' �ϥ� ComboBox �� Name �@���з����W (Key)
                    selectedField = cbo.value       ' �ϥ� ComboBox �襤���ȧ@���ӷ����W (Value)
                    mappings.Add standardField, selectedField
                    Debug.Print "  Mapping added: """ & standardField & """ (from ComboBox """ & cbo.Name & """ with Tag ""header"") -> """ & selectedField & """"
                Else
                    ' ComboBox �Q�аO�� "header"�A���ϥΪ̥���ܥ��󶵥�
                    Debug.Print "  Info: ComboBox """ & cbo.Name & """ (Tag: ""header"") has no selection."
                End If
            End If
            Set cbo = Nothing ' ���� ComboBox ����Ѧ�
        End If
    Next ctrl
    
    Debug.Print Me.Name & "." & METHOD_NAME & " - Mappings collected. Count: " & mappings.Count
    Set GetFieldMappings = mappings ' �^�ǫإߦn�� Dictionary
    Exit Function

ErrorHandler:
    Debug.Print "!!! ERROR in " & Me.Name & "." & METHOD_NAME & " !!!"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "Error Description: " & Err.Description
    MsgBox "Ū�� TB �������]�w�ɵo�Ϳ��~�G" & vbCrLf & Err.Description, vbCritical, "���~"
    Set GetFieldMappings = Nothing ' �o�Ϳ��~�ɦ^�� Nothing
    Set mappings = Nothing
    Set cbo = Nothing
End Function

Public Sub PopulateComboBoxes(fieldNames As Variant)
    Const METHOD_NAME As String = "PopulateComboBoxes"
    Dim cbo As MSForms.ComboBox
    Dim ctrl As MSForms.Control
    Dim fieldName As Variant
    Dim comboBoxTag As String ' �Ω��x�s ComboBox �� Tag �ݩ�

    On Error GoTo ErrorHandler
    
    Dim items_cboMethod As Variant: items_cboMethod = Array("�~���ܰʪ��B", "����������B", "�ɤ�U����B", "�ɶU�褧���쥽���B")
    Dim items_cboApproveDateAsGLDate As Variant: items_cboApproveDateAsGLDate = Array("Yes", "No")
    Dim items_cboEntryType As Variant: items_cboEntryType = Array("Auto", "Manual")
    
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.ComboBox Then
            Set cbo = ctrl
            
            Select Case cbo.Name
                Case "cboMethod"
                    cbo.Clear
                    cbo.List = items_cboMethod
                    cbo.ListIndex = -1  ' �w�]���襤
                
                Case "cboApproveDateAsGLDate"
                    cbo.Clear
                    cbo.List = items_cboApproveDateAsGLDate
                    cbo.ListIndex = -1
                Case "cboEntryType"
                    cbo.Clear
                    cbo.List = items_cboEntryType
                    cbo.ListIndex = -1
            End Select
            
            Set cbo = Nothing  ' ���񪫥�
        End If
    Next ctrl
    
    ' �ˬd�ǤJ���ѼƬO�_�����Ī��}�C
    If IsEmpty(fieldNames) Or Not IsArray(fieldNames) Then
        Debug.Print Me.Name & "." & METHOD_NAME & " - �����즳�Ī����W�ٰ}�C�C"
        ' �M�ũҦ� Tag �� "header" �� ComboBox
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is MSForms.ComboBox Then ' �T�O�O ComboBox
                Set cbo = ctrl
                comboBoxTag = Trim(cbo.tag)
                If comboBoxTag = "header" Then '
                    cbo.Clear
                End If
                Set cbo = Nothing
            End If
        Next ctrl
        Exit Sub
    End If

    ' �M�����W���Ҧ����
    For Each ctrl In Me.Controls
        ' �ˬd�O�_�� ComboBox
        If TypeOf ctrl Is MSForms.ComboBox Then
            Set cbo = ctrl ' �N����૬�� ComboBox
            comboBoxTag = Trim(cbo.tag) ' Ū���òM�z Tag �ݩ�
            
            ' �u��R Tag �� "header" �� ComboBox
            If comboBoxTag = "header" Then
                cbo.Clear
                ' �N���W�٥[�J ComboBox
                For Each fieldName In fieldNames
                    cbo.AddItem fieldName
                Next fieldName
                If cbo.ListCount > 0 Then
                    cbo.ListIndex = -1
                End If
                Debug.Print "  Populated ComboBox: """ & cbo.Name & """ (Tag: ""header"")"
            End If
            Set cbo = Nothing
        End If
    Next ctrl

    Debug.Print Me.Name & "." & METHOD_NAME & " - �w�N " & UBound(fieldNames) + 1 & " ������J Tag �� ""header"" �� ComboBox�C"
    Exit Sub

ErrorHandler:
    Debug.Print "--- �b " & Me.Name & "." & METHOD_NAME & " ���o�Ϳ��~ ---"
    Debug.Print "���~�N�X: " & Err.Number
    Debug.Print "���~�y�z: " & Err.Description
    MsgBox "��J '" & Me.Caption & "' ���U�Կ��ɵo�Ϳ��~�G" & vbCrLf & Err.Description, vbCritical, "���~"
    Set cbo = Nothing
End Sub

