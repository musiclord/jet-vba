VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vTBConfig 
   Caption         =   "TB Configuration"
   ClientHeight    =   8415.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   OleObjectBlob   =   "vTBConfig.frx":0000
   StartUpPosition =   1  '所屬視窗中央
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

' --- 初始化事件 (清空或移除欄位載入邏輯) ---
Private Sub UserForm_Initialize()
    ' 確保此處不執行欄位載入
    Debug.Print Me.Name & " Initialized."
End Sub

' --- 新增: 公共方法，用於取得使用者設定的 TB 欄位對應 ---
Public Function GetFieldMappings() As Object ' Returns Scripting.Dictionary
    Const METHOD_NAME As String = "GetFieldMappings"
    Dim mappings As Object ' Scripting.Dictionary
    Dim conceptualFieldName As String ' 我們內部定義的欄位名
    Dim selectedSourceField As String ' 使用者選擇的來源欄位名

    On Error GoTo ErrorHandler

    ' 建立 Dictionary 物件
    Set mappings = CreateObject("Scripting.Dictionary")
    mappings.CompareMode = vbTextCompare ' Key 不區分大小寫
    Debug.Print Me.Name & "." & METHOD_NAME & " - Dictionary created."

    ' --- 處理指定的 ComboBox ---

    ' 1. AccountNumber (會計科目編號) - 假設 vTBConfig 上有 cboAccountNo
    conceptualFieldName = "AccountNumber" ' 標準內部名稱
    If Me.cboAccountNo.ListIndex > -1 Then
        selectedSourceField = Me.cboAccountNo.value
        mappings.Add conceptualFieldName, selectedSourceField
        Debug.Print "  Mapping added: """ & conceptualFieldName & """ -> """ & selectedSourceField & """"
    Else
        Debug.Print "  Warning: ComboBox 'cboAccountNo' (for " & conceptualFieldName & ") has no selection."
        ' mappings.Add conceptualFieldName, "" ' 可選：加入空字串表示未選擇
    End If

    ' 2. AccountName (會計科目名稱) - 假設 vTBConfig 上有 cboAccountName
    conceptualFieldName = "AccountName" ' 標準內部名稱
    If Me.cboAccountName.ListIndex > -1 Then
        selectedSourceField = Me.cboAccountName.value
        mappings.Add conceptualFieldName, selectedSourceField
        Debug.Print "  Mapping added: """ & conceptualFieldName & """ -> """ & selectedSourceField & """"
    Else
        Debug.Print "  Warning: ComboBox 'cboAccountName' (for " & conceptualFieldName & ") has no selection."
        ' mappings.Add conceptualFieldName, "" ' 可選：加入空字串表示未選擇
    End If

    Debug.Print Me.Name & "." & METHOD_NAME & " - Mappings collected. Count: " & mappings.Count
    Set GetFieldMappings = mappings ' 回傳建立好的 Dictionary
    Exit Function

ErrorHandler:
    Debug.Print "!!! ERROR in " & Me.Name & "." & METHOD_NAME & " !!!"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "Error Description: " & Err.Description
    MsgBox "讀取 TB 欄位對應設定時發生錯誤：" & vbCrLf & Err.Description, vbCritical, "錯誤"
    Set GetFieldMappings = Nothing ' 發生錯誤時回傳 Nothing
    Set mappings = Nothing
End Function

' --- 新增: 公共方法，用於從外部填入 ComboBox ---
Public Sub PopulateComboBoxes(fieldNames As Variant)
    Const METHOD_NAME As String = "PopulateComboBoxes"
    Dim cbo As MSForms.ComboBox
    Dim ctrl As MSForms.Control
    Dim fieldName As Variant

    On Error GoTo ErrorHandler

    ' 檢查傳入的參數是否為有效的陣列
    If IsEmpty(fieldNames) Or Not IsArray(fieldNames) Then
        Debug.Print Me.Name & "." & METHOD_NAME & " - 未收到有效的欄位名稱陣列，ComboBox 將不會被填入。"
        ' 清空所有 ComboBox
        For Each ctrl In Me.Controls
            If TypeName(ctrl) = "ComboBox" Then
                Set cbo = ctrl
                cbo.Clear
                Set cbo = Nothing
            End If
        Next ctrl
        Exit Sub
    End If

    ' 遍歷表單上的所有控制項
    For Each ctrl In Me.Controls
        ' 檢查是否為 ComboBox
        If TypeName(ctrl) = "ComboBox" Then
            Set cbo = ctrl ' 將控制項轉型為 ComboBox
            cbo.Clear ' 清除現有項目
            ' 將欄位名稱加入 ComboBox
            For Each fieldName In fieldNames
                cbo.AddItem fieldName
            Next fieldName
            ' 可選：設定預設不選中任何項目
            If cbo.ListCount > 0 Then
                cbo.ListIndex = -1
            End If
            Set cbo = Nothing
        End If
    Next ctrl

    Debug.Print Me.Name & "." & METHOD_NAME & " - 已成功將 " & UBound(fieldNames) + 1 & " 個欄位填入 ComboBox。"
    Exit Sub

ErrorHandler:
    Debug.Print "--- 在 " & Me.Name & "." & METHOD_NAME & " 中發生錯誤 ---"
    Debug.Print "錯誤代碼: " & Err.Number
    Debug.Print "錯誤描述: " & Err.Description
    MsgBox "填入 '" & Me.Caption & "' 的下拉選單時發生錯誤：" & vbCrLf & Err.Description, vbCritical, "錯誤"
    Set cbo = Nothing
End Sub
