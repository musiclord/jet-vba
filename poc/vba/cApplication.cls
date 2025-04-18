VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Controller: Application
'Description: The main controller of all interfaces and classes.
Option Explicit

Private WithEvents vMain As vMain
Private cMapping As cMapping
Private AccessDAL As AccessDAL
Private GLEntity As GLEntity
Private GLService As GLService
Private TBEntity As TBEntity
Private TBService As TBService
Private ImportService As ImportService
Private MappingService As MappingService

Private Sub Class_Initialize()
    Set ImportService = New ImportService
    Set MappingService = New MappingService
    Set GLEntity = New GLEntity
    Set GLService = New GLService
    Set TBEntity = New TBEntity
    Set TBService = New TBService
    ' AccessDAL 應該由需要它的 Service 內部建立和管理，或者透過依賴注入傳遞
    ' Set AccessDAL = New AccessDAL ' <-- 考慮移除，讓 Service 管理自己的 DAL
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
    Dim tableName As String
    ' Set tableName as vMain.ListTable's value (comboBox).
    Call PreviewTable(tableName)
End Sub

Private Sub vMain_GetTableNames()
    ' Update database table names to vMain.ListTable
End Sub
Private Sub PreviewTable(ByVal tableName As String)
    ' Let user preview specific table, with maximum 1000 records.
    Const MAX As Long = 1000  ' limit should be define by PrevieTable,
    Call ImportService.LoadToExcel(accessDbPath, tableName, worksheetName, MAX_ROWS_TO_SHOW)
    
End Sub
Private Sub ImportCSV(ByVal importType As String)
    ' 控制器：協調 CSV 匯入流程 (GL 或 TB)
    ' 先將CSV匯入至資料庫(ImportService.ImportToAccess)
    ' 再將匯入的資料表預覽(...PreviewTable)
    ' 因此這裡不會呼叫 ImportService.LoadToExcel，而是由 PreviewTable 處理
    Dim csvFilePath As String
    Dim accessDbPath As String
    Dim tableName As String
    Dim worksheetName As String ' 用於預覽的工作表名稱
    Dim csvEncoding As Long
    Const MAX_ROWS_TO_SHOW As Long = 1000      ' 預覽顯示的最大列數
    Const DATABASE_NAME As String = "default.accdb" ' <-- 定義預設的資料庫名稱
    Const DEFAULT_ENCODING As Long = 950       ' <-- 定義預設編碼
    
    ' 1. 選擇 CSV 檔案
    csvFilePath = Application.GetOpenFilename("CSV 檔案 (*.csv),*.csv", , "選擇要匯入的 " & importType & " CSV 檔案")
    If csvFilePath = "False" Then Exit Sub ' 使用者取消
    
    ' 2. 設定路徑和名稱
    accessDbPath = ThisWorkbook.path & "\" & DATABASE_NAME ' <-- 使用預設的資料庫名稱
    tableName = importType                                 ' <-- 資料表名稱由 importType 決定
    worksheetName = importType                             ' <-- 預覽工作表名稱也使用 importType
    
    ' 3. 執行匯入流程
    Application.StatusBar = "正在處理 " & importType & " 資料到資料庫 [" & DATABASE_NAME & "]，請稍候..."
    Application.Cursor = xlWait
    On Error GoTo ErrorHandler
    
    ' 3a. 偵測編碼 (呼叫 mod_Utility 中的函數)
    csvEncoding = DetectCSVEncoding(csvFilePath)
    Debug.Print importType & " CSV 編碼偵測結果: " & csvEncoding
    
    ' 3b. 匯入至 Access (調用 ImportService)
    Application.StatusBar = "正在將 " & importType & " CSV 匯入 Access 資料表 [" & tableName & "]..."
    ' 傳遞正確的參數給 ImportService
    If Not ImportService.ImportToAccess(csvFilePath, accessDbPath, tableName, csvEncoding) Then
        ' ImportToAccess 內部應處理具體錯誤，這裡只處理是否成功完成
        ' 錯誤訊息已由 ImportToAccess 的 ErrorHandler 拋出並在此處被捕獲
        GoTo ErrorHandler ' 如果匯入 Access 失敗，則跳轉到通用錯誤處理
    End If
    MsgBox "成功將 " & importType & " CSV 匯入資料庫 [" & DATABASE_NAME & "] 的資料表 [" & tableName & "]。", vbInformation, "匯入成功" ' <-- 更新成功訊息
    
    ' 3c. 從 Access 載入至 Excel (可選的立即預覽)
    Application.StatusBar = "正在從 Access 將資料表 [" & tableName & "] 載入 Excel 預覽..." ' <-- 更新狀態列訊息
    If Not ImportService.LoadToExcel(accessDbPath, tableName, worksheetName, MAX_ROWS_TO_SHOW) Then
         ' LoadToExcel 內部應處理具體錯誤
         GoTo ErrorHandler ' 如果載入 Excel 失敗，則跳轉到通用錯誤處理
    End If
    
    ' 4. 完成
    Application.StatusBar = False
    Application.Cursor = xlDefault
    MsgBox importType & " 匯入與預覽完成！" & vbCrLf & "資料預覽已顯示在工作表 [" & worksheetName & "]", vbInformation, "處理完成"
    Exit Sub

ErrorHandler:
    ' 通用錯誤處理：恢復狀態列和滑鼠指標，顯示錯誤訊息
    Application.StatusBar = False
    Application.Cursor = xlDefault
    ' Err.Source 會顯示錯誤發生的模組/類別 (例如 "ImportCSV.ImportToAccess")
    MsgBox importType & " 處理過程中發生錯誤：" & vbCrLf & _
           "來源: " & Err.Source & vbCrLf & _
           "描述: " & Err.Description, vbCritical, "錯誤"
    ' 注意：這裡不需要再次 Debug.Print，因為 Service 層的 ErrorHandler 應該已經記錄了詳細資訊
End Sub
