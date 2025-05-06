VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vGLConfig 
   Caption         =   "GL Configuration"
   ClientHeight    =   8415.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   OleObjectBlob   =   "vGLConfig.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "vGLConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===== 表單資訊 =====
'名稱: vGLConfig
'描述: 設定總帳 (GL) 資料匯入的欄位對應和相關選項。
'
'===== 表單控制項 =====
'下拉式選單 (ComboBox):
'  - cboDocumentNo: 對應來源檔中的「傳票號碼」欄位
'  - cboLineItem: 對應來源檔中的「傳票項次」欄位
'  - cbotPostDate: 對應來源檔中的「總帳日期」欄位
'  - cboAccountNo: 對應來源檔中的「會計科目編號」欄位
'  - cboAccountName: 對應來源檔中的「會計科目名稱」欄位
'  - cboDescription: 對應來源檔中的「傳票摘要」欄位
'  - cboMethod: 選擇總帳金額的處理方式 (例如：單一金額欄位、借貸分開欄位等)
'  - cboAmount: 對應來源檔中的「傳票金額」欄位 (若 cboMethod 選擇單一金額)
'  - cboDebit: 對應來源檔中的「借方金額」欄位 (若 cboMethod 選擇借貸分開)
'  - cboCredit: 對應來源檔中的「貸方金額」欄位 (若 cboMethod 選擇借貸分開)
'  - cboDrCr: 對應來源檔中的「借貸類別」欄位 (若 cboMethod 選擇借貸標記)
'  - cboDebitFlag: 輸入代表借方的標記值 (若 cboMethod 選擇借貸標記)
'  - cboPostingStatus: 對應來源檔中的「過帳狀態」欄位
'  - cboPostedCode: 輸入代表已過帳的狀態代碼
'文字輸入欄 (TextBox):
'  - txtbPosted: 選擇是否僅篩選已過帳的資料 (通常是 CheckBox 或 OptionButton，若為 ComboBox 則可能提供 Yes/No 選項)
'命令按鈕 (CommandButton):
'  - btnImport: 觸發匯入 GL 資料的事件 (RaiseEvent DoImport)
'  - btnPreview: 觸發預覽 GL 資料表的事件 (RaiseEvent DoPreview)
'  - btnConfirm: 確認設定並觸發 DoConfirm 事件 (RaiseEvent DoConfirm)
'  - btnExit: 關閉表單並觸發 DoExit 事件 (RaiseEvent DoExit)


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

' --- 初始化事件 (清空或移除欄位載入邏輯) ---
Private Sub UserForm_Initialize()
    ' 確保此處不執行欄位載入
    Debug.Print Me.Name & " Initialized."
End Sub
' --- 新增: 公共方法，用於取得使用者設定的欄位對應 ---
Public Function GetFieldMappings() As Object ' Returns Scripting.Dictionary
    Const METHOD_NAME As String = "GetFieldMappings"
    Dim mappings As Object ' Scripting.Dictionary
    Dim ctrl As MSForms.Control
    Dim cbo As MSForms.ComboBox
    Dim comboBoxTag As String
    Dim standardField As String ' ComboBox 的 Name，作為標準欄位名 (Key)
    Dim selectedField As String ' ComboBox 選中的 Value，作為來源欄位名 (Value)

    On Error GoTo ErrorHandler

    ' 建立 Dictionary 物件
    Set mappings = CreateObject("Scripting.Dictionary")
    mappings.CompareMode = vbTextCompare ' Key 不區分大小寫
    Debug.Print Me.Name & "." & METHOD_NAME & " - Dictionary created."

    ' 遍歷表單上的所有控制項
    For Each ctrl In Me.Controls
        ' 檢查是否為 ComboBox
        If TypeOf ctrl Is MSForms.ComboBox Then
            Set cbo = ctrl
            comboBoxTag = Trim(cbo.tag)
            ' 如果 Tag 屬性被標記為 "header"，則處理此 ComboBox
            If comboBoxTag = "header" Then
                If cbo.ListIndex > -1 Then          ' 確保有項目被選中
                    standardField = cbo.Name        ' 使用 ComboBox 的 Name 作為標準欄位名 (Key)
                    selectedField = cbo.value       ' 使用 ComboBox 選中的值作為來源欄位名 (Value)
                    mappings.Add standardField, selectedField
                    Debug.Print "  Mapping added: """ & standardField & """ (from ComboBox """ & cbo.Name & """ with Tag ""header"") -> """ & selectedField & """"
                Else
                    ' ComboBox 被標記為 "header"，但使用者未選擇任何項目
                    Debug.Print "  Info: ComboBox """ & cbo.Name & """ (Tag: ""header"") has no selection."
                End If
            End If
            Set cbo = Nothing ' 釋放 ComboBox 物件參考
        End If
    Next ctrl
    
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
    Set cbo = Nothing
End Function

Public Sub PopulateComboBoxes(fieldNames As Variant)
    Const METHOD_NAME As String = "PopulateComboBoxes"
    Dim cbo As MSForms.ComboBox
    Dim ctrl As MSForms.Control
    Dim fieldName As Variant
    Dim comboBoxTag As String ' 用於儲存 ComboBox 的 Tag 屬性

    On Error GoTo ErrorHandler
    
    Dim items_cboMethod As Variant: items_cboMethod = Array("年度變動金額", "期初期末金額", "借方貸方金額", "借貸方之期初末金額")
    Dim items_cboApproveDateAsGLDate As Variant: items_cboApproveDateAsGLDate = Array("Yes", "No")
    Dim items_cboEntryType As Variant: items_cboEntryType = Array("Auto", "Manual")
    
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.ComboBox Then
            Set cbo = ctrl
            
            Select Case cbo.Name
                Case "cboMethod"
                    cbo.Clear
                    cbo.List = items_cboMethod
                    cbo.ListIndex = -1  ' 預設不選中
                
                Case "cboApproveDateAsGLDate"
                    cbo.Clear
                    cbo.List = items_cboApproveDateAsGLDate
                    cbo.ListIndex = -1
                Case "cboEntryType"
                    cbo.Clear
                    cbo.List = items_cboEntryType
                    cbo.ListIndex = -1
            End Select
            
            Set cbo = Nothing  ' 釋放物件
        End If
    Next ctrl
    
    ' 檢查傳入的參數是否為有效的陣列
    If IsEmpty(fieldNames) Or Not IsArray(fieldNames) Then
        Debug.Print Me.Name & "." & METHOD_NAME & " - 未收到有效的欄位名稱陣列。"
        ' 清空所有 Tag 為 "header" 的 ComboBox
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is MSForms.ComboBox Then ' 確保是 ComboBox
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

    ' 遍歷表單上的所有控制項
    For Each ctrl In Me.Controls
        ' 檢查是否為 ComboBox
        If TypeOf ctrl Is MSForms.ComboBox Then
            Set cbo = ctrl ' 將控制項轉型為 ComboBox
            comboBoxTag = Trim(cbo.tag) ' 讀取並清理 Tag 屬性
            
            ' 只填充 Tag 為 "header" 的 ComboBox
            If comboBoxTag = "header" Then
                cbo.Clear
                ' 將欄位名稱加入 ComboBox
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

    Debug.Print Me.Name & "." & METHOD_NAME & " - 已將 " & UBound(fieldNames) + 1 & " 個欄位填入 Tag 為 ""header"" 的 ComboBox。"
    Exit Sub

ErrorHandler:
    Debug.Print "--- 在 " & Me.Name & "." & METHOD_NAME & " 中發生錯誤 ---"
    Debug.Print "錯誤代碼: " & Err.Number
    Debug.Print "錯誤描述: " & Err.Description
    MsgBox "填入 '" & Me.Caption & "' 的下拉選單時發生錯誤：" & vbCrLf & Err.Description, vbCritical, "錯誤"
    Set cbo = Nothing
End Sub

