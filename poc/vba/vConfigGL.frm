VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vConfigGL 
   Caption         =   "GL Configuration"
   ClientHeight    =   8412.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   11760
   OleObjectBlob   =   "vConfigGL.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "vConfigGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===== ����T =====
'�W��: vConfigGL
'�y�z: �]�w�`�b (GL) ��ƶפJ���������M�����ﶵ�C
'

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

Public Function GetFieldMappings() As Object ' Scripting.Dictionary
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
    Debug.Print Me.Name & "." & METHOD_NAME & " - Dictionary created."

    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.ComboBox Then
            Set cbo = ctrl
            standardField = Trim(cbo.tag)
            selectedField = Trim(cbo.value)
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
    Debug.Print "!!! ERROR in " & Me.Name & "." & METHOD_NAME & " !!!"
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
            
            Select Case cbo.Name
                Case "cboMethod"
                    cbo.Clear
                    cbo.List = Array("�~���ܰʪ��B", "����������B", "�ɤ�U����B", "�ɶU�褧���쥽���B")
                    cbo.ListIndex = -1  ' �w�]���襤
                
                Case "cboApproveDateAsGLDate"
                    cbo.Clear
                    cbo.List = Array("Yes", "No")
                    cbo.ListIndex = -1
                Case "cboEntryType"
                    cbo.Clear
                    cbo.List = Array("Auto", "Manual")
                    cbo.ListIndex = -1
            End Select
            
            Set cbo = Nothing  ' ���񪫥�
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
            End If
            Set cbo = Nothing
        End If
    Next ctrl
    
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.ComboBox Then
            Set cbo = ctrl
            If cbo.Name = "cboDocumentNo" Then
                cbo.value = "�ǲ����X"
            ElseIf cbo.Name = "cboMethod" Then
                cbo.value = "�ɤ�P�U����B"
            ElseIf cbo.Name = "cboDebit" Then
                cbo.value = "�����ɤ���B"
            ElseIf cbo.Name = "cboCredit" Then
                cbo.value = "�����U����B"
            ElseIf cbo.Name = "cboDebitFlag" Then
                cbo.value = "�ɶU"
            ElseIf cbo.Name = "cboPostDate" Then
                cbo.value = "���"
            ElseIf cbo.Name = "cboAccountNo" Then
                cbo.value = "��إN�X"
            ElseIf cbo.Name = "cboAccountName" Then
                cbo.value = "�|�p���"
            ElseIf cbo.Name = "cboDescription" Then
                cbo.value = "�K�n"
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

