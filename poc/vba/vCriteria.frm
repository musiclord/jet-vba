VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vCriteria 
   Caption         =   "Criteria"
   ClientHeight    =   6576
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10680
   OleObjectBlob   =   "vCriteria.frx":0000
   StartUpPosition =   1  '所屬視窗中央
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
    ' 填入 cboOperator1
    Me.cboOperator1.Clear
    For Each op In operators
        Me.cboOperator1.AddItem op
    Next op
    Me.cboOperator1.ListIndex = -1 ' 預設不選中

    ' 填入 cboOperator2
    Me.cboOperator2.Clear
    For Each op In operators
        Me.cboOperator2.AddItem op
    Next op
    Me.cboOperator2.ListIndex = -1 ' 預設不選中
End Sub
Private Sub btnConfirm_Click()
    RaiseEvent DoConfirm
End Sub

Private Sub btnExit_Click()
    RaiseEvent DoExit
End Sub

' --- 公共方法，用於從外部填入指定的 ComboBox ---
Public Sub PopulateComboBoxes(fieldNames As Variant)
    Const METHOD_NAME As String = "PopulateComboBoxes"
    Dim fieldName As Variant
    Dim populatedCount As Integer
    populatedCount = 0

    On Error GoTo ErrorHandler

    ' 檢查傳入的參數是否為有效的陣列
    If IsEmpty(fieldNames) Or Not IsArray(fieldNames) Then
        Debug.Print Me.Name & "." & METHOD_NAME & " - 未收到有效的欄位名稱陣列，ComboBox 將不會被填入。"
        ' 清空目標 ComboBox
        On Error Resume Next ' 忽略可能的錯誤 (例如控制項不存在)
        Me.cboColumn1.Clear
        Me.cboColumn2.Clear
        On Error GoTo ErrorHandler ' 恢復正常錯誤處理
        Exit Sub
    End If

    ' --- 填充 cboColumn1 ---
    On Error Resume Next ' 暫時忽略錯誤，以防控制項不存在
    Me.cboColumn1.Clear ' 清除現有項目
    If Err.Number = 0 Then ' 檢查 Clear 操作是否成功 (表示控制項存在)
        ' 將欄位名稱加入 ComboBox
        For Each fieldName In fieldNames
            Me.cboColumn1.AddItem fieldName
        Next fieldName
        ' 可選：設定預設不選中任何項目
        If Me.cboColumn1.ListCount > 0 Then
            Me.cboColumn1.ListIndex = -1
        End If
        populatedCount = populatedCount + 1
    Else
        Debug.Print Me.Name & "." & METHOD_NAME & " - 警告: 控制項 'cboColumn1' 可能不存在或發生錯誤。"
        Err.Clear ' 清除錯誤
    End If
    On Error GoTo ErrorHandler ' 恢復正常錯誤處理

    ' --- 填充 cboColumn2 ---
    On Error Resume Next ' 暫時忽略錯誤，以防控制項不存在
    Me.cboColumn2.Clear ' 清除現有項目
    If Err.Number = 0 Then ' 檢查 Clear 操作是否成功 (表示控制項存在)
        ' 將欄位名稱加入 ComboBox
        For Each fieldName In fieldNames
            Me.cboColumn2.AddItem fieldName
        Next fieldName
        ' 可選：設定預設不選中任何項目
        If Me.cboColumn2.ListCount > 0 Then
            Me.cboColumn2.ListIndex = -1
        End If
        populatedCount = populatedCount + 1
    Else
        Debug.Print Me.Name & "." & METHOD_NAME & " - 警告: 控制項 'cboColumn2' 可能不存在或發生錯誤。"
        Err.Clear ' 清除錯誤
    End If
    On Error GoTo ErrorHandler ' 恢復正常錯誤處理


    Debug.Print Me.Name & "." & METHOD_NAME & " - 已嘗試將 " & UBound(fieldNames) + 1 & " 個欄位填入 " & populatedCount & " 個目標 ComboBox。"
    Exit Sub

ErrorHandler:
    Debug.Print "--- 在 " & Me.Name & "." & METHOD_NAME & " 中發生錯誤 ---"
    Debug.Print "錯誤代碼: " & Err.Number
    Debug.Print "錯誤描述: " & Err.Description
    ' 避免在 PopulateComboBoxes 中顯示 MsgBox，讓呼叫端處理更嚴重的錯誤
    ' MsgBox "填入 '" & Me.Caption & "' 的下拉選單時發生錯誤：" & vbCrLf & Err.Description, vbCritical, "錯誤"
End Sub
