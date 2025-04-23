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
' Controller: Application Controller
' Description: Main controller for the application, handling UI events and coordinating services.

Private Const MODULE_NAME As String = "cApplication"
Private Const DATABASE_NAME As String = "default.accdb"
Private Const MAX_ROWS_TO_SHOW As Long = 1000
Private WithEvents vMain As vMain
Attribute vMain.VB_VarHelpID = -1
Private cMapping As cMapping
Private GLEntity As GLEntity
Private TBEntity As TBEntity
Private GLService As GLService
Private TBService As TBService
Private ImportService As ImportService
Private PreviewService As PreviewService
Private MappingService As MappingService

Private Sub Class_Initialize()
    Set ImportService = New ImportService
    Set PreviewService = New PreviewService
    Set MappingService = New MappingService
    Set GLEntity = New GLEntity
    Set GLService = New GLService
    Set TBEntity = New TBEntity
    Set TBService = New TBService
    Set vMain = New vMain
End Sub

Private Sub vMain_DoExit()
    Call vMain.Hide
    ' --> Save checkpoint <--
    Unload vMain
    Set vMain = Nothing
End Sub

Private Sub vMain_DoImportGL()
    Call ImportCSV("GL")
End Sub

Private Sub vMain_DoImportTB()
    Call ImportCSV("TB")
End Sub

Private Sub vMain_DoMapping()
    Set cMapping = New cMapping
End Sub

Private Sub vMain_DoPreview()
    Call PreviewTable
End Sub

Private Sub vMain_GetTableNames()
    Call GetTableNames
End Sub
Private Sub PreviewTable()
    ' Controller: Handles the request to preview a selected table from vMain.
    Const METHOD_NAME As String = "PreviewTable"
    Dim tableName As String
    Dim worksheetName As String
    Dim accessDbPath As String

    On Error GoTo ErrorHandler

    ' 1. 從 vMain 的 ComboBox (命名為 ListTable) 獲取選取的資料表名稱
    ' 修改：使用更可靠的方式獲取 ComboBox 的值
    On Error Resume Next ' Handle ListTable potentially not existing or no selection
    If vMain.ListTable.ListIndex >= 0 Then ' 確保有選取項目
        tableName = vMain.ListTable.List(vMain.ListTable.ListIndex)
    Else
        tableName = ""
    End If
    
    If Err.Number <> 0 Or tableName = "" Or tableName = "(無可用資料表)" Or tableName = "(讀取資料表錯誤)" Or tableName = "(讀取資料表時傳回非預期類型)" Then
        MsgBox "請先從下拉選單中選擇一個有效的資料表進行預覽。", vbExclamation, "未選擇資料表"
        Exit Sub
    End If
    On Error GoTo ErrorHandler ' Restore default error handling

    ' 2. 設定參數 (使用模組級常數)
    accessDbPath = ThisWorkbook.Path & "\" & DATABASE_NAME ' <-- Use module const
    worksheetName = tableName & "_Preview" ' 修改：統一工作表命名規則

    ' 3. 執行預覽 (調用 PreviewService, 使用模組級常數)
    Application.StatusBar = "正在載入資料表 [" & tableName & "] 的預覽..."
    Application.Cursor = xlWait

    ' Ensure PreviewService is initialized
    If PreviewService Is Nothing Then Set PreviewService = New PreviewService

    If PreviewService.ShowPreview(accessDbPath, tableName, worksheetName, MAX_ROWS_TO_SHOW) Then ' <-- Use module const
        Application.StatusBar = False
        Application.Cursor = xlDefault
        MsgBox "資料表 [" & tableName & "] 預覽完成！" & vbCrLf & "資料已顯示在工作表 [" & worksheetName & "]", vbInformation, "預覽完成"
    Else
        GoTo ErrorHandler
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.Cursor = xlDefault
    MsgBox "預覽資料表 [" & tableName & "] 時發生錯誤：" & vbCrLf & _
           "來源: " & Err.Source & vbCrLf & _
           "描述: " & Err.Description, vbCritical, "預覽錯誤"
End Sub

Private Sub GetTableNames()
    ' Controller: Gets table names via PreviewService and populates vMain.ListTable
    Dim tableNamesVariant As Variant ' Service returns a Variant Array
    Dim tableName As Variant
    Dim accessDbPath As String

    On Error GoTo ErrorHandler

    ' 1. 設定資料庫路徑 (使用模組級常數)
    accessDbPath = ThisWorkbook.Path & "\" & DATABASE_NAME

    ' 2. 調用 PreviewService 獲取資料表名稱
    ' Ensure PreviewService is initialized (should be in Class_Initialize)
    If PreviewService Is Nothing Then Set PreviewService = New PreviewService
    tableNamesVariant = PreviewService.GetAccessTableNames(accessDbPath)

    ' 3. 清空並禁用 ComboBox (預設)
    ' vMain.ListTable.MatchRequired = False ' <-- 移除此行
    vMain.ListTable.Clear
    vMain.ListTable.Enabled = False ' 禁用以進行填充

    ' 4. 填充 ComboBox
    If IsEmpty(tableNamesVariant) Then
         vMain.ListTable.AddItem "(讀取資料表錯誤或無資料表)"
         ' 保持禁用狀態
    ElseIf IsArray(tableNamesVariant) Then
        If UBound(tableNamesVariant) >= LBound(tableNamesVariant) Then ' Check if array is not empty
            For Each tableName In tableNamesVariant
                vMain.ListTable.AddItem tableName
            Next tableName

            ' 在設定 ListIndex 前檢查 ListCount
            If vMain.ListTable.ListCount > 0 Then
                vMain.ListTable.Enabled = True ' *** 啟用 ComboBox ***
                vMain.ListTable.ListIndex = 0 ' 再設定預設選項
                ' vMain.ListTable.MatchRequired = True ' <-- 移除此行
            Else
                ' 如果迴圈後列表仍為空
                If vMain.ListTable.ListCount = 0 Then
                     vMain.ListTable.AddItem "(無有效資料表)"
                End If
                ' 保持禁用狀態
            End If
        Else
            vMain.ListTable.AddItem "(無可用資料表)"
            ' 保持禁用狀態
        End If
    Else
         vMain.ListTable.AddItem "(讀取資料表時傳回非預期類型)"
         ' 保持禁用狀態
    End If

CleanUp:
    Exit Sub
ErrorHandler:
     ' 確保 ComboBox 在錯誤時保持禁用狀態
     If Not vMain Is Nothing Then ' 避免在卸載過程中出錯
         On Error Resume Next ' 忽略設置屬性時可能出現的錯誤
         vMain.ListTable.Clear
         vMain.ListTable.AddItem "(讀取資料表錯誤)"
         vMain.ListTable.Enabled = False
         ' vMain.ListTable.MatchRequired = False ' <-- 移除此行
         On Error GoTo 0
     End If
     MsgBox "讀取資料庫資料表列表時發生錯誤。" & vbCrLf & "請檢查資料庫連接或聯繫管理員。", vbCritical, "讀取錯誤"
     GoTo CleanUp
End Sub

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
    ' 注意: DetectCSVEncoding 應位於 mod_Utility
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
    If Not ImportService.ImportToAccess(csvFilePath, accessDbPath, tableName, csvEncoding) Then
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
    ' ... (ErrorHandler 保持不變) ...
    Application.StatusBar = False
    Application.Cursor = xlDefault
    MsgBox importType & " 處理過程中發生錯誤：" & vbCrLf & _
           "來源: " & Err.Source & vbCrLf & _
           "描述: " & Err.Description, vbCritical, "錯誤"
End Sub


