VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===============================================================================
' Module:    cApplication
' Purpose:   應用程式的主要控制器 (Controller)。
'            負責處理來自 vMain 使用者介面的事件，
'            協調各個服務 (Service) 類別來執行業務邏輯，
'            並管理應用程式的生命週期。
'===============================================================================

' --- 常數 ---
Private Const MODULE_NAME As String = "cApplication"
Private Const DATABASE_NAME As String = "default.accdb"
Private Const MAX_ROWS_TO_SHOW As Long = 1000

' --- 模組變數 ---
Private WithEvents vMain As vMain
Attribute vMain.VB_VarHelpID = -1
Private GLEntity As GLEntity
Private TBEntity As TBEntity
Private GLService As GLService
Private TBService As TBService
Private ImportService As ImportService
Private PreviewService As PreviewService
Private MappingService As MappingService
Private ValidationService As ValidationService
' --- Step 1 ---
Private WithEvents vProject As vProject
Attribute vProject.VB_VarHelpID = -1
Private WithEvents vTBConfig As vTBConfig
Attribute vTBConfig.VB_VarHelpID = -1
Private WithEvents vGLConfig As vGLConfig
Attribute vGLConfig.VB_VarHelpID = -1
' --- Step 2 ---
Private WithEvents vValidation As vValidation
Attribute vValidation.VB_VarHelpID = -1
' --- Step 3 ---
Private WithEvents vCriteria As vCriteria
Attribute vCriteria.VB_VarHelpID = -1
' --- Step 4 ---

Private Sub Class_Initialize()
    Set ImportService = New ImportService
    Set PreviewService = New PreviewService
    Set MappingService = New MappingService
    Set ValidationService = New ValidationService
    Set GLEntity = New GLEntity
    Set GLService = New GLService
    Set TBEntity = New TBEntity
    Set TBService = New TBService
    Set vMain = New vMain
End Sub

' ====== vMain ======
Private Sub vMain_DoExit()
    ' -> Save AppConfig <-
    ' Unload vMain
    ' Set vMain = Nothing
    vMain.Hide
End Sub

Private Sub vMain_DoStep1()
    Set vProject = New vProject
    Set vTBConfig = New vTBConfig
    Set vGLConfig = New vGLConfig
    
    Call vProject.Show(vbModeless)
End Sub

Private Sub vMain_DoStep2()
    ' -> vValidation
    Set vValidation = New vValidation
End Sub

Private Sub vMain_DoStep3()
    ' -> vCriteria
    Set vCriteria = New vCriteria
    Call PopulateFormComboBoxes(vCriteria, "GL")
    Call vCriteria.Show(vbModeless)
End Sub

Private Sub vMain_DoStep4()
    ' ...
End Sub

' ====== vProject ======
Private Sub vProject_DoExit()
    Unload vProject
    Set vProject = Nothing
End Sub

Private Sub vProject_DoConfirm()
    Call PopulateFormComboBoxes(vTBConfig, "TB") ' 將 TB 欄位注入至 vTBConfig
    ' 移至下個視窗
    Call vProject.Hide
    Call vTBConfig.Show(vbModeless)
    vMain.lblProject.BackColor = vbYellow
End Sub

' ====== vTBConfig ======
Private Sub vTBConfig_DoExit()
    Unload vTBConfig
    Set vTBConfig = Nothing
End Sub

Private Sub vTBConfig_DoImport()
    Call ImportCSV("TB")
End Sub

Private Sub vTBConfig_DoPreview()
    Call PreviewTable("TB") ' 在工作表預覽資料表 TB
End Sub

Private Sub vTBConfig_DoConfirm()
    Call ProcessAndStoreFormMappings(vTBConfig, "TB") ' 處理並儲存 TB 欄位配對
    Call PopulateFormComboBoxes(vGLConfig, "GL") ' 將 GL 欄位注入至 vGLConfig
    ' 移至下個視窗
    Call vTBConfig.Hide
    Call vGLConfig.Show(vbModeless)
    vMain.lblImportTB.BackColor = vbYellow
End Sub

' ====== vGLConfig ======
Private Sub vGLConfig_DoExit()
    Unload vGLConfig
    Set vGLConfig = Nothing
End Sub

Private Sub vGLConfig_DoImport()
    Call ImportCSV("GL")
End Sub

Private Sub vGLConfig_DoPreview()
    Call PreviewTable("GL") ' 在工作表預覽資料表 GL
End Sub

Private Sub vGLConfig_DoConfirm()
    Call ProcessAndStoreFormMappings(vGLConfig, "GL") ' 處理並儲存 GL 欄位配對
    ' 移至下個視窗
    Call vGLConfig.Hide
    vMain.lblImportGL.BackColor = vbYellow
End Sub

' ====== vValidation ======
Private Sub vValidation_DoExit()
    Unload vValidation
    Set vValidation = Nothing
End Sub

Private Sub vValidation_TestCompleteness()
    ' Call ValidationService.TestCompleteness
    vMain.lblCompletenessTest.BackColor = vbYellow
End Sub

Private Sub vValidation_TestDocumentBalance()
    vMain.lblDocumentBalance.BackColor = vbYellow
End Sub

Private Sub vValidation_TestRDE()
    vMain.lblRDETest.BackColor = vbYellow
End Sub

Private Sub vValidation_DoAccountMapping()
    vMain.lblAccountMapping.BackColor = vbYellow
End Sub

Private Sub vValidation_DoConfirm()
    Call vValidation.Hide
End Sub

' ====== vCriteria ======
Private Sub vCriteria_DoConfirm()
    Call vCriteria.Hide
    vMain.lblCriteria.BackColor = vbYellow
End Sub

Private Sub vCriteria_DoExit()
    Unload vCriteria
    Set vCriteria = Nothing
End Sub

'-------------------------------------------------------------------------------
' Procedure: ImportCSV
' Purpose:   處理將 CSV 檔案匯入到 Access 資料庫的通用流程。
'            包括選擇檔案、偵測編碼、調用 ImportService 匯入資料，
'            以及調用 PreviewService 顯示匯入後的資料預覽。
' Arguments:
'   @param importType (String): 指定匯入的類型 ("GL" 或 "TB")，用於決定目標資料表名稱等。
' Note:      會使用模組級常數 DATABASE_NAME 和 MAX_ROWS_TO_SHOW。
'            依賴於標準模組 mod_Utility 中的 DetectCSVEncoding 函數。
'-------------------------------------------------------------------------------
Private Sub ImportCSV(ByVal importType As String)
    ' 控制器：協調 CSV 匯入流程 (GL 或 TB)
    Dim csvFilePath As String
    Dim accessDbPath As String
    Dim tableName As String
    Dim worksheetName As String ' 用於預覽的工作表名稱
    Dim csvEncoding As Long
    Const DEFAULT_ENCODING As Long = 950 ' <-- 保持局部
    ' 1. 選擇 CSV 檔案
    csvFilePath = Application.GetOpenFilename("CSV 檔案 (*.csv),*.csv", , "選擇要匯入的 " & importType & " CSV 檔案")
    If csvFilePath = "False" Then Exit Sub ' 使用者取消
    
    ' 2. 設定路徑和名稱 (使用模組級常數，統一工作表命名)
    accessDbPath = ThisWorkbook.Path & "\" & DATABASE_NAME ' <-- 使用模組級常數
    Select Case importType
        Case "GL"
            tableName = "GL"
        Case "TB"
            tableName = "TB"
        Case Else
            MsgBox "不支援的匯入類型：" & importType, vbCritical, "內部錯誤"
            Exit Sub
    End Select
    worksheetName = tableName ' & "_Preview"
    
    ' 3. 執行匯入流程
    Application.StatusBar = "正在處理 " & importType & " 資料到資料庫 [" & DATABASE_NAME & "]，請稍候..."
    Application.Cursor = xlWait
    On Error GoTo ErrorHandler
    
    ' 3a. 偵測編碼
    ' 注意: DetectCSVEncoding 位於 mod_Utility.bas
    On Error Resume Next
    csvEncoding = mod_Utility.DetectCSVEncoding(csvFilePath)
    If Err.Number <> 0 Then
        Debug.Print "Warning: Failed to detect CSV encoding for " & csvFilePath & ". Using default: " & DEFAULT_ENCODING
        csvEncoding = DEFAULT_ENCODING
        Err.Clear
    End If
    On Error GoTo ErrorHandler ' 恢復正常錯誤處理
    Debug.Print importType & " CSV 編碼偵測結果: " & csvEncoding

    ' 3b. 匯入至 Access (調用 ImportService)
    Application.StatusBar = "正在將 " & importType & " CSV 匯入 Access 資料表 [" & tableName & "]..."
    ' 假設 ImportService.ImportToAccess 簽名包含 csvEncoding
    If Not ImportService.ImportToAccess(csvFilePath, accessDbPath, tableName, CLng(csvEncoding)) Then
        GoTo ErrorHandler
    End If

    ' 3c. 從 Access 載入至 Excel 進行預覽 (調用 PreviewService)
    Application.StatusBar = "正在從 Access 將資料表 [" & tableName & "] 載入 Excel 預覽..."
    If Not PreviewService.ShowPreview(accessDbPath, tableName, worksheetName, MAX_ROWS_TO_SHOW) Then ' <-- 使用模組級常數
         GoTo ErrorHandler
    End If
    
    ' 4. 完成
    Application.StatusBar = False
    Application.Cursor = xlDefault
    MsgBox importType & " 匯入與預覽完成！" & vbCrLf & "資料預覽已顯示在工作表 [" & worksheetName & "]", vbInformation, "處理完成"
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.Cursor = xlDefault
    MsgBox importType & " 處理過程中發生錯誤：" & vbCrLf & _
           "來源: " & Err.Source & vbCrLf & _
           "描述: " & Err.Description, vbCritical, "錯誤"
End Sub

'-------------------------------------------------------------------------------
' Procedure: PreviewTable
' Purpose:   從 Access 資料庫讀取指定的資料表，
'            並使用 PreviewService 將其內容顯示在新的 Excel 工作表中。
'            工作表名稱會根據資料表名稱加上 "_Preview" 後綴。
' Arguments:
'   @param tableName (String): 要預覽的 Access 資料表名稱。
' Note:      會使用模組級常數 DATABASE_NAME 和 MAX_ROWS_TO_SHOW。
'-------------------------------------------------------------------------------
Private Sub PreviewTable(ByVal tableName As String)
    Const METHOD_NAME As String = "PreviewTable"
    ' Dim tableName As String ' tableName 現在是參數
    Dim worksheetName As String
    Dim accessDbPath As String
    On Error GoTo ErrorHandler

    ' 1. 驗證輸入的 tableName (基本檢查)
    If Trim$(tableName) = "" Then
        MsgBox "必須提供有效的資料表名稱才能進行預覽。", vbExclamation, "未指定資料表"
        Exit Sub
    End If

    ' 2. 設定參數 (使用模組級常數)
    accessDbPath = ThisWorkbook.Path & "\" & DATABASE_NAME
    worksheetName = tableName & "_Preview" ' 維持原本的命名邏輯

    ' 3. 執行預覽 (調用 PreviewService, 使用模組級常數)
    Application.StatusBar = "正在載入資料表 [" & tableName & "] 的預覽..."
    Application.Cursor = xlWait
    If PreviewService.ShowPreview(accessDbPath, tableName, worksheetName, MAX_ROWS_TO_SHOW) Then ' <-- Use module const
        Application.StatusBar = False
        Application.Cursor = xlDefault
        ' 成功訊息現在使用傳入的 tableName 和衍生的 worksheetName
        MsgBox "資料表 [" & tableName & "] 預覽完成！" & vbCrLf & "資料已顯示在工作表 [" & worksheetName & "]", vbInformation, "預覽完成"
    Else
        ' 如果 PreviewService.ShowPreview 傳回 False (雖然目前設計是錯誤時引發 Error)
        ' 或是發生其他未預期的狀況跳到 Else
        GoTo ErrorHandler ' 轉到通用錯誤處理
    End If
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.Cursor = xlDefault
    ' 錯誤訊息現在使用傳入的 tableName
    MsgBox "預覽資料表 [" & tableName & "] 時發生錯誤：" & vbCrLf & _
           "來源: " & Err.Source & vbCrLf & _
           "描述: " & Err.Description, vbCritical, "預覽錯誤"
End Sub

' --- 通用方法，準備並填充指定表單的 ComboBox ---
Private Sub PopulateFormComboBoxes(ByVal targetForm As Object, ByVal tableName As String)
    Const METHOD_NAME As String = "PrepareAndPopulateFormComboBoxes"
    Dim fieldNames As Variant
    Dim dal As AccessDAL ' 臨時 DAL 物件
    Dim dbPath As String

    On Error GoTo ErrorHandler

    ' 檢查 targetForm 是否有效
    If targetForm Is Nothing Then
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 錯誤: targetForm 為 Nothing。"
        Exit Sub
    End If

    ' --- 1. 取得指定資料表的欄位名稱 ---
    dbPath = ThisWorkbook.Path & "\" & DATABASE_NAME ' 取得資料庫路徑
    If Dir(dbPath) = "" Then
        MsgBox "找不到資料庫檔案：" & vbCrLf & dbPath, vbCritical, MODULE_NAME & "." & METHOD_NAME
        fieldNames = Empty ' 設定為 Empty
    Else
        Set dal = New AccessDAL ' 建立 DAL 實例
        dal.DatabasePath = dbPath ' 設定路徑
        fieldNames = dal.GetTableFieldNames(tableName) ' 呼叫函數
        Set dal = Nothing ' 立即釋放 DAL 物件
    End If

    ' --- 2. 將欄位名稱注入目標表單 ---
    '    這裡假設 targetForm 有一個名為 PopulateComboBoxes 的 Public Sub
    '    使用 On Error Resume Next 處理 targetForm 可能沒有此方法的情況 (雖然最好確保有)
    On Error Resume Next
    CallByName targetForm, "PopulateComboBoxes", VbMethod, fieldNames
    If Err.Number <> 0 Then
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 警告: 無法呼叫 '" & TypeName(targetForm) & ".PopulateComboBoxes'。錯誤: " & Err.Description
        Err.Clear
    End If
    On Error GoTo ErrorHandler ' 恢復正常的錯誤處理

CleanUp:
    Set dal = Nothing ' 確保錯誤時也釋放
    Exit Sub

ErrorHandler:
    MsgBox "準備填充表單 '" & TypeName(targetForm) & "' 時發生錯誤 (資料表: " & tableName & ")：" & vbCrLf & Err.Description, vbCritical, MODULE_NAME & "." & METHOD_NAME
    ' 嘗試填充空值
    On Error Resume Next
    CallByName targetForm, "PopulateComboBoxes", VbMethod, Empty
    Err.Clear
    On Error GoTo 0 ' 關閉此處的錯誤處理
    GoTo CleanUp
End Sub

' --- 新增: 通用方法，處理並儲存指定表單的欄位對應 ---
Private Function ProcessAndStoreFormMappings(ByVal sourceForm As Object, ByVal mappingType As String) As Boolean
    Const METHOD_NAME As String = "ProcessAndStoreFormMappings"
    Dim formMappings As Object ' Scripting.Dictionary

    On Error GoTo ErrorHandler

    ' --- 1. 檢查輸入參數 ---
    If sourceForm Is Nothing Then
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - Error: sourceForm object is Nothing."
        MsgBox "來源表單物件無效。", vbCritical, MODULE_NAME & "." & METHOD_NAME
        GoTo CleanUp ' 返回 False
    End If
    If MappingService Is Nothing Then
         Debug.Print MODULE_NAME & "." & METHOD_NAME & " - Error: MappingService object is Nothing."
         MsgBox "對應服務尚未初始化。", vbCritical, MODULE_NAME & "." & METHOD_NAME
         GoTo CleanUp ' 返回 False
    End If

    ' --- 2. 從來源表單取得欄位對應 ---
    '    假設 sourceForm 有 GetFieldMappings 方法
    On Error Resume Next ' 暫時忽略錯誤，以便檢查
    Set formMappings = sourceForm.GetFieldMappings()
    If Err.Number <> 0 Then
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - Error calling GetFieldMappings on " & TypeName(sourceForm) & ": " & Err.Description
        MsgBox "無法從 '" & TypeName(sourceForm) & "' 表單讀取欄位對應。", vbCritical, MODULE_NAME & "." & METHOD_NAME
        Err.Clear
        GoTo CleanUp ' 返回 False
    End If
    On Error GoTo ErrorHandler ' 恢復正常錯誤處理

    ' --- 3. 驗證對應是否成功取得 ---
    If formMappings Is Nothing Then
        ' GetFieldMappings 內部應該已經顯示錯誤訊息
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - Error: Failed to get mappings from " & TypeName(sourceForm) & "."
        GoTo CleanUp ' 返回 False
    End If

    ' --- 4. 將對應關係傳遞給 MappingService ---
    Select Case UCase$(mappingType) ' 根據類型呼叫不同的 Set 方法
        Case "TB"
            Call MappingService.SetTBFieldMappings(formMappings)
            Debug.Print MODULE_NAME & "." & METHOD_NAME & " - TB field mappings passed to MappingService."
        Case "GL"
            Call MappingService.SetGLFieldMappings(formMappings)
            Debug.Print MODULE_NAME & "." & METHOD_NAME & " - GL field mappings passed to MappingService."
        Case Else
            Debug.Print MODULE_NAME & "." & METHOD_NAME & " - Error: Unknown mappingType '" & mappingType & "'."
            MsgBox "不支援的對應類型：" & mappingType, vbCritical, MODULE_NAME & "." & METHOD_NAME
            GoTo CleanUp ' 返回 False
    End Select

CleanUp:
    Set formMappings = Nothing
    Exit Function

ErrorHandler:
    MsgBox "處理 '" & mappingType & "' 欄位對應時發生錯誤：" & vbCrLf & Err.Description, vbCritical, MODULE_NAME & "." & METHOD_NAME
    GoTo CleanUp
End Function
