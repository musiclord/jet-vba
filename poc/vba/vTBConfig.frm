VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vTBConfig 
   Caption         =   "TB Configuration"
   ClientHeight    =   8415.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   OleObjectBlob   =   "vTBConfig.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "vTBConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'View: TB Configuration
'Description:

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


Private Sub cboAccountName_Change()

End Sub

' --- ��l�ƨƥ� (�M�ũβ��������J�޿�) ---
Private Sub UserForm_Initialize()
    ' �T�O���B�����������J
    Debug.Print Me.Name & " Initialized."
End Sub

' --- �s�W: ���@��k�A�Ω���o�ϥΪ̳]�w�� TB ������ ---
Public Function GetFieldMappings() As Object ' Returns Scripting.Dictionary
    Const METHOD_NAME As String = "GetFieldMappings"
    Dim mappings As Object ' Scripting.Dictionary
    Dim conceptualFieldName As String ' �ڭ̤����w�q�����W
    Dim selectedSourceField As String ' �ϥΪ̿�ܪ��ӷ����W

    On Error GoTo ErrorHandler

    ' �إ� Dictionary ����
    Set mappings = CreateObject("Scripting.Dictionary")
    mappings.CompareMode = vbTextCompare ' Key ���Ϥ��j�p�g
    Debug.Print Me.Name & "." & METHOD_NAME & " - Dictionary created."

    ' --- �B�z���w�� ComboBox ---

    ' 1. AccountNumber (�|�p��ؽs��) - ���] vTBConfig �W�� cboAccountNo
    conceptualFieldName = "AccountNumber" ' �зǤ����W��
    If Me.cboAccountNo.ListIndex > -1 Then
        selectedSourceField = Me.cboAccountNo.value
        mappings.Add conceptualFieldName, selectedSourceField
        Debug.Print "  Mapping added: """ & conceptualFieldName & """ -> """ & selectedSourceField & """"
    Else
        Debug.Print "  Warning: ComboBox 'cboAccountNo' (for " & conceptualFieldName & ") has no selection."
        ' mappings.Add conceptualFieldName, "" ' �i��G�[�J�Ŧr���ܥ����
    End If

    ' 2. AccountName (�|�p��ئW��) - ���] vTBConfig �W�� cboAccountName
    conceptualFieldName = "AccountName" ' �зǤ����W��
    If Me.cboAccountName.ListIndex > -1 Then
        selectedSourceField = Me.cboAccountName.value
        mappings.Add conceptualFieldName, selectedSourceField
        Debug.Print "  Mapping added: """ & conceptualFieldName & """ -> """ & selectedSourceField & """"
    Else
        Debug.Print "  Warning: ComboBox 'cboAccountName' (for " & conceptualFieldName & ") has no selection."
        ' mappings.Add conceptualFieldName, "" ' �i��G�[�J�Ŧr���ܥ����
    End If

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
End Function

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
