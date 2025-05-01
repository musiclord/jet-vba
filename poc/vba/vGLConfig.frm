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

Private Sub cboDescription_Change()

End Sub

' --- ��l�ƨƥ� (�M�ũβ��������J�޿�) ---
Private Sub UserForm_Initialize()
    ' �T�O���B�����������J
    Debug.Print Me.Name & " Initialized."
End Sub

' --- �s�W: ���@��k�A�Ω�q�~����J ComboBox ---
Public Sub PopulateComboBoxes(fieldNames As Variant)
    Const METHOD_NAME As String = "PopulateComboBoxes"
    Dim cbo As MSForms.ComboBox
    Dim ctrl As MSForms.Control
    Dim fieldName As Variant

    On Error GoTo ErrorHandler

    ' �ˬd�ǤJ���ѼƬO�_�����Ī��}�C
    If IsEmpty(fieldNames) Or Not IsArray(fieldNames) Then
        Debug.Print Me.Name & "." & METHOD_NAME & " - �����즳�Ī����W�ٰ}�C�AComboBox �N���|�Q��J�C"
        ' �M�ũҦ� ComboBox
        For Each ctrl In Me.Controls
            If TypeName(ctrl) = "ComboBox" Then
                Set cbo = ctrl
                cbo.Clear
                Set cbo = Nothing
            End If
        Next ctrl
        Exit Sub
    End If

    ' �M�����W���Ҧ����
    For Each ctrl In Me.Controls
        ' �ˬd�O�_�� ComboBox
        If TypeName(ctrl) = "ComboBox" Then
            Set cbo = ctrl ' �N����૬�� ComboBox
            cbo.Clear ' �M���{������
            ' �N���W�٥[�J ComboBox
            For Each fieldName In fieldNames
                cbo.AddItem fieldName
            Next fieldName
            ' �i��G�]�w�w�]���襤���󶵥�
            If cbo.ListCount > 0 Then
                cbo.ListIndex = -1
            End If
            Set cbo = Nothing
        End If
    Next ctrl

    Debug.Print Me.Name & "." & METHOD_NAME & " - �w���\�N " & UBound(fieldNames) + 1 & " ������J ComboBox�C"
    Exit Sub

ErrorHandler:
    Debug.Print "--- �b " & Me.Name & "." & METHOD_NAME & " ���o�Ϳ��~ ---"
    Debug.Print "���~�N�X: " & Err.Number
    Debug.Print "���~�y�z: " & Err.Description
    MsgBox "��J '" & Me.Caption & "' ���U�Կ��ɵo�Ϳ��~�G" & vbCrLf & Err.Description, vbCritical, "���~"
    Set cbo = Nothing
End Sub

' --- �s�W: ���@��k�A�Ω���o�ϥΪ̳]�w�� GL ������ ---
Public Function GetFieldMappings() As Object ' Returns Scripting.Dictionary
    Const METHOD_NAME As String = "GetFieldMappings"
    Dim mappings As Object ' Scripting.Dictionary
    Dim selectedSourceField As String ' �ϥΪ̿�ܪ��ӷ����W
    Dim cbo As MSForms.ComboBox ' �Ω��ˬd����O�_�s�b
    Dim controlName As String
    Dim internalName As String

    On Error GoTo ErrorHandler

    ' �إ� Dictionary ����
    Set mappings = CreateObject("Scripting.Dictionary")
    mappings.CompareMode = vbTextCompare ' Key ���Ϥ��j�p�g
    Debug.Print Me.Name & "." & METHOD_NAME & " - Dictionary created."

    ' --- �B�z cboDocumentNo ---
    controlName = "cboDocumentNo"
    internalName = "DocumentNo"
    On Error Resume Next ' �����䤣�챱������~
    Set cbo = Me.Controls(controlName)
    If Err.Number = 0 Then ' ����s�b
        If cbo.ListIndex > -1 Then
            selectedSourceField = cbo.value
            mappings.Add internalName, selectedSourceField
            Debug.Print "  Mapping added: """ & internalName & """ -> """ & selectedSourceField & """"
        Else
            Debug.Print "  Warning: ComboBox '" & controlName & "' (for " & internalName & ") has no selection."
        End If
    Else
        Debug.Print "  Error: Control '" & controlName & "' not found on form " & Me.Name & "."
    End If
    Set cbo = Nothing
    Err.Clear
    On Error GoTo ErrorHandler ' ��_���`���~�B�z

    ' --- �B�z cboLineItem ---
    controlName = "cboLineItem"
    internalName = "LineItem"
    On Error Resume Next
    Set cbo = Me.Controls(controlName)
    If Err.Number = 0 Then
        If cbo.ListIndex > -1 Then
            selectedSourceField = cbo.value
            mappings.Add internalName, selectedSourceField
            Debug.Print "  Mapping added: """ & internalName & """ -> """ & selectedSourceField & """"
        Else
            Debug.Print "  Warning: ComboBox '" & controlName & "' (for " & internalName & ") has no selection."
        End If
    Else
        Debug.Print "  Error: Control '" & controlName & "' not found on form " & Me.Name & "."
    End If
    Set cbo = Nothing
    Err.Clear
    On Error GoTo ErrorHandler

    ' --- �B�z cbotPostDate ---
    controlName = "cbotPostDate"
    internalName = "PostDate"
    On Error Resume Next
    Set cbo = Me.Controls(controlName)
    If Err.Number = 0 Then
        If cbo.ListIndex > -1 Then
            selectedSourceField = cbo.value
            mappings.Add internalName, selectedSourceField
            Debug.Print "  Mapping added: """ & internalName & """ -> """ & selectedSourceField & """"
        Else
            Debug.Print "  Warning: ComboBox '" & controlName & "' (for " & internalName & ") has no selection."
        End If
    Else
        Debug.Print "  Error: Control '" & controlName & "' not found on form " & Me.Name & "."
    End If
    Set cbo = Nothing
    Err.Clear
    On Error GoTo ErrorHandler

    ' --- �B�z cboAccountNo ---
    controlName = "cboAccountNo"
    internalName = "AccountNo"
    On Error Resume Next
    Set cbo = Me.Controls(controlName)
    If Err.Number = 0 Then
        If cbo.ListIndex > -1 Then
            selectedSourceField = cbo.value
            mappings.Add internalName, selectedSourceField
            Debug.Print "  Mapping added: """ & internalName & """ -> """ & selectedSourceField & """"
        Else
            Debug.Print "  Warning: ComboBox '" & controlName & "' (for " & internalName & ") has no selection."
        End If
    Else
        Debug.Print "  Error: Control '" & controlName & "' not found on form " & Me.Name & "."
    End If
    Set cbo = Nothing
    Err.Clear
    On Error GoTo ErrorHandler

    ' --- �B�z cboAccountName ---
    controlName = "cboAccountName"
    internalName = "AccountName"
    On Error Resume Next
    Set cbo = Me.Controls(controlName)
    If Err.Number = 0 Then
        If cbo.ListIndex > -1 Then
            selectedSourceField = cbo.value
            mappings.Add internalName, selectedSourceField
            Debug.Print "  Mapping added: """ & internalName & """ -> """ & selectedSourceField & """"
        Else
            Debug.Print "  Warning: ComboBox '" & controlName & "' (for " & internalName & ") has no selection."
        End If
    Else
        Debug.Print "  Error: Control '" & controlName & "' not found on form " & Me.Name & "."
    End If
    Set cbo = Nothing
    Err.Clear
    On Error GoTo ErrorHandler

    ' --- �B�z cboDescription ---
    controlName = "cboDescription"
    internalName = "Description"
    On Error Resume Next
    Set cbo = Me.Controls(controlName)
    If Err.Number = 0 Then
        If cbo.ListIndex > -1 Then
            selectedSourceField = cbo.value
            mappings.Add internalName, selectedSourceField
            Debug.Print "  Mapping added: """ & internalName & """ -> """ & selectedSourceField & """"
        Else
            Debug.Print "  Warning: ComboBox '" & controlName & "' (for " & internalName & ") has no selection."
        End If
    Else
        Debug.Print "  Error: Control '" & controlName & "' not found on form " & Me.Name & "."
    End If
    Set cbo = Nothing
    Err.Clear
    On Error GoTo ErrorHandler

    Debug.Print Me.Name & "." & METHOD_NAME & " - Mappings collected. Count: " & mappings.Count
    Set GetFieldMappings = mappings ' �^�ǫإߦn�� Dictionary
    Exit Function

ErrorHandler:
    Debug.Print "!!! ERROR in " & Me.Name & "." & METHOD_NAME & " !!!"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "Error Description: " & Err.Description
    MsgBox "Ū�� GL �������]�w�ɵo�Ϳ��~�G" & vbCrLf & Err.Description, vbCritical, "���~"
    Set GetFieldMappings = Nothing ' �o�Ϳ��~�ɦ^�� Nothing
    Set mappings = Nothing
    Set cbo = Nothing ' �T�O cbo �]�Q�M�z
End Function
