VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vConfigTB 
   Caption         =   "TB Configuration"
   ClientHeight    =   8412.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   OleObjectBlob   =   "vConfigTB.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "vConfigTB"
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

Private Sub UserForm_Initialize()
    Debug.Print Me.Name & " Initialized."
End Sub

' --- �s�W: ���@��k�A�Ω���o�ϥΪ̳]�w�������� ---
Public Function GetFieldMappings() As Object ' Returns Scripting.Dictionary
    Const METHOD_NAME As String = "GetFieldMappings"
    Dim mappings As Object ' Scripting.Dictionary
    Dim ctrl As MSForms.Control
    Dim cbo As MSForms.ComboBox
    Dim tag As String
    Dim standardField As String
    Dim selectedField As String

    On Error GoTo ErrorHandler
    
    Set mappings = CreateObject("Scripting.Dictionary")
    mappings.CompareMode = vbTextCompare ' Key ���Ϥ��j�p�g

    ' �M�����W���Ҧ����
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.ComboBox Then
            Set cbo = ctrl
            standardField = Trim(cbo.tag)       ' Tag ���з����W��
            selectedField = Trim(cbo.value)     ' value ���t�����W��
            If cbo.tag <> "" Then
                If cbo.ListIndex > -1 Then ' �T�O�����سQ�襤
                    mappings.Add standardField, selectedField
                    Debug.Print "  Mapping added: [" & standardField & "] --> [" & selectedField & "], (Name:" & cbo.Name&; ",Tag:" & cbo.tag & ")"
                Else
                    ' ComboBox �Q�аO������ܥ��󶵥�
                    Debug.Print "  Info: ComboBox [" & cbo.Name & "] with tag [" & cbo.tag & "] has no selection."
                End If
            End If
            Set cbo = Nothing
        End If
    Next ctrl
    
    Debug.Print Me.Name & "." & METHOD_NAME & " - Mappings collected. Count: " & mappings.Count
    Set GetFieldMappings = mappings ' �^�ǫإߦn�� Dictionary
    Exit Function

ErrorHandler:
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

    On Error GoTo ErrorHandler
    
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.ComboBox Then
            Set cbo = ctrl
            If cbo.Name = "cboMethod" Then
                cbo.Clear
                cbo.List = Array("�~���ܰʪ��B", "����P�������B", "�ɤ�P�U����B", "�ɶU�褧����������B")
                cbo.ListIndex = -1 ' �w�]���襤
            End If
            Set cbo = Nothing ' �T�O�b�`��������
        End If
    Next ctrl
    
    ' �ˬd�ǤJ���ѼƬO�_�����Ī��}�C
    If IsEmpty(fieldNames) Or Not IsArray(fieldNames) Then
        Debug.Print Me.Name & "." & METHOD_NAME & " - �����즳�Ī����W�ٰ}�C�C"
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is MSForms.ComboBox Then
                Set cbo = ctrl
                If cbo.tag <> "" Then
                    cbo.Clear
                End If
                Set cbo = Nothing
            End If
        Next ctrl
        Exit Sub
    End If

    For Each ctrl In Me.Controls
        ' �ˬd�O�_�� ComboBox
        If TypeOf ctrl Is MSForms.ComboBox Then
            Set cbo = ctrl
            
            If cbo.tag <> "" Then
                cbo.Clear
                For Each fieldName In fieldNames
                    cbo.AddItem fieldName
                Next fieldName
                If cbo.ListCount > 0 Then
                    cbo.ListIndex = -1
                End If
                ' Debug.Print "Populated ComboBox: " & cbo.Name & "."
            End If
            Set cbo = Nothing
        End If
    Next ctrl
    
    ' ��J�w�]��
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.ComboBox Then
            Set cbo = ctrl
            If cbo.Name = "cboAccountName" Then
                cbo.value = "���ئW��"
            ElseIf cbo.Name = "cboAccountNo" Then
                cbo.value = "�|�p����"
            ElseIf cbo.Name = "cboMethod" Then
                cbo.value = "�ɤ�P�U����B"
            ElseIf cbo.Name = "cboCredit" Then
                cbo.value = "����U����B"
            ElseIf cbo.Name = "cboDebit" Then
                cbo.value = "����ɤ���B"
            End If
            Set cbo = Nothing
        End If
    Next ctrl

    Debug.Print Me.Name & "." & METHOD_NAME & " - �w�N " & UBound(fieldNames) + 1 & " ������J ComboBox�C"
    Exit Sub

ErrorHandler:
    Debug.Print "--- �b " & Me.Name & "." & METHOD_NAME & " ���o�Ϳ��~ ---"
    Debug.Print "���~�N�X: " & Err.Number
    Debug.Print "���~�y�z: " & Err.Description
    MsgBox "��J '" & Me.Caption & "' ���U�Կ��ɵo�Ϳ��~�G" & vbCrLf & Err.Description, vbCritical, "���~"
    Set cbo = Nothing
End Sub

