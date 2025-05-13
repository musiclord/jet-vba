VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vCriteria 
   Caption         =   "Criteria"
   ClientHeight    =   6576
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10680
   OleObjectBlob   =   "vCriteria.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "vCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Event DoExit()
Public Event DoConfirm()


Private Sub cboOperator1_Change()

End Sub

Private Sub UserForm_Initialize()
    ' Me.Show (vbModeless)
    Dim operators As Variant
    Dim op As Variant
    operators = Array("IN", ">", "<", "=")
    ' ��J cboOperator1
    Me.cboOperator1.Clear
    For Each op In operators
        Me.cboOperator1.AddItem op
    Next op
    Me.cboOperator1.ListIndex = -1 ' �w�]���襤

    ' ��J cboOperator2
    Me.cboOperator2.Clear
    For Each op In operators
        Me.cboOperator2.AddItem op
    Next op
    Me.cboOperator2.ListIndex = -1 ' �w�]���襤
End Sub
Private Sub btnConfirm_Click()
    RaiseEvent DoConfirm
End Sub

Private Sub btnExit_Click()
    RaiseEvent DoExit
End Sub

' --- ���@��k�A�Ω�q�~����J���w�� ComboBox ---
Public Sub PopulateComboBoxes(fieldNames As Variant)
    Const METHOD_NAME As String = "PopulateComboBoxes"
    Dim fieldName As Variant
    Dim populatedCount As Integer
    populatedCount = 0

    On Error GoTo ErrorHandler

    ' �ˬd�ǤJ���ѼƬO�_�����Ī��}�C
    If IsEmpty(fieldNames) Or Not IsArray(fieldNames) Then
        Debug.Print Me.Name & "." & METHOD_NAME & " - �����즳�Ī����W�ٰ}�C�AComboBox �N���|�Q��J�C"
        ' �M�ťؼ� ComboBox
        On Error Resume Next ' �����i�઺���~ (�Ҧp������s�b)
        Me.cboColumn1.Clear
        Me.cboColumn2.Clear
        On Error GoTo ErrorHandler ' ��_���`���~�B�z
        Exit Sub
    End If

    ' --- ��R cboColumn1 ---
    On Error Resume Next ' �Ȯɩ������~�A�H��������s�b
    Me.cboColumn1.Clear ' �M���{������
    If Err.Number = 0 Then ' �ˬd Clear �ާ@�O�_���\ (��ܱ���s�b)
        ' �N���W�٥[�J ComboBox
        For Each fieldName In fieldNames
            Me.cboColumn1.AddItem fieldName
        Next fieldName
        ' �i��G�]�w�w�]���襤���󶵥�
        If Me.cboColumn1.ListCount > 0 Then
            Me.cboColumn1.ListIndex = -1
        End If
        populatedCount = populatedCount + 1
    Else
        Debug.Print Me.Name & "." & METHOD_NAME & " - ĵ�i: ��� 'cboColumn1' �i�ण�s�b�εo�Ϳ��~�C"
        Err.Clear ' �M�����~
    End If
    On Error GoTo ErrorHandler ' ��_���`���~�B�z

    ' --- ��R cboColumn2 ---
    On Error Resume Next ' �Ȯɩ������~�A�H��������s�b
    Me.cboColumn2.Clear ' �M���{������
    If Err.Number = 0 Then ' �ˬd Clear �ާ@�O�_���\ (��ܱ���s�b)
        ' �N���W�٥[�J ComboBox
        For Each fieldName In fieldNames
            Me.cboColumn2.AddItem fieldName
        Next fieldName
        ' �i��G�]�w�w�]���襤���󶵥�
        If Me.cboColumn2.ListCount > 0 Then
            Me.cboColumn2.ListIndex = -1
        End If
        populatedCount = populatedCount + 1
    Else
        Debug.Print Me.Name & "." & METHOD_NAME & " - ĵ�i: ��� 'cboColumn2' �i�ण�s�b�εo�Ϳ��~�C"
        Err.Clear ' �M�����~
    End If
    On Error GoTo ErrorHandler ' ��_���`���~�B�z


    Debug.Print Me.Name & "." & METHOD_NAME & " - �w���ձN " & UBound(fieldNames) + 1 & " ������J " & populatedCount & " �ӥؼ� ComboBox�C"
    Exit Sub

ErrorHandler:
    Debug.Print "--- �b " & Me.Name & "." & METHOD_NAME & " ���o�Ϳ��~ ---"
    Debug.Print "���~�N�X: " & Err.Number
    Debug.Print "���~�y�z: " & Err.Description
    ' �קK�b PopulateComboBoxes ����� MsgBox�A���I�s�ݳB�z���Y�������~
    ' MsgBox "��J '" & Me.Caption & "' ���U�Կ��ɵo�Ϳ��~�G" & vbCrLf & Err.Description, vbCritical, "���~"
End Sub
