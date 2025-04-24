VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImportService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'Model: Import Service
'Description: Handle any data import logics and uses AccessDAL for database operation.

Private Const adDate As Long = 7 ' <-- 用於 LoadToExcel 的日期格式化
Private Const MODULE_NAME As String = "ImportService" ' <-- 加入模組名稱常數

Private Sub Class_Initialize()

End Sub

Public Function ImportToAccess(ByVal csvFilePath As String, ByVal accessDbPath As String, ByVal tableName As String, ByVal csvEncoding As Long) As Boolean
    ' 重構後：使用 AccessDAL 處理資料庫操作，利用 Text Driver 進行匯入
    Const METHOD_NAME As String = "ImportToAccess"
    Dim dal As AccessDAL
    Dim fso As Object ' FileSystemObject
    Dim csvFolder As String
    Dim csvFile As String
    Dim sql As String
    Dim schemaIniPath As String ' Schema.ini 路徑

    On Error GoTo ErrorHandler
    ImportToAccess = False ' 預設為失敗

    ' 1. 驗證輸入與檔案存在
    If csvFilePath = "" Or accessDbPath = "" Or tableName = "" Then
        Err.Raise vbObjectError + 513, MODULE_NAME & "." & METHOD_NAME, "CSV file path, Access DB path, and target table name cannot be empty."
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(csvFilePath) Then
        Err.Raise vbObjectError + 514, MODULE_NAME & "." & METHOD_NAME, "CSV file not found: " & csvFilePath
    End If

    csvFolder = fso.GetParentFolderName(csvFilePath)
    csvFile = fso.GetFileName(csvFilePath)
    schemaIniPath = csvFolder & "\Schema.ini" ' 假設 Schema.ini 在 CSV 同目錄

    ' 2. 檢查 Schema.ini (重要：Text Driver 需要此檔案來正確解析 CSV)
    If Not fso.FileExists(schemaIniPath) Then
         ' 實際應用中可能需要更複雜的邏輯來創建或更新 Schema.ini
         ' 例如，讀取 CSV 標頭並寫入欄位定義。
         ' 目前假設 Schema.ini 已存在且配置正確。
         Debug.Print "警告: Schema.ini 未在 CSV 資料夾中找到: " & schemaIniPath & ". Text Driver 可能無法正確解析 CSV。"
         ' 可以選擇在這裡報錯退出，或繼續嘗試 (可能失敗或資料類型錯誤)
         ' Err.Raise vbObjectError + 516, MODULE_NAME & "." & METHOD_NAME, "Schema.ini not found in CSV folder: " & csvFolder
    End If
    Set fso = Nothing

    ' 3. 建立並設定 DAL
    Set dal = New AccessDAL
    dal.DatabasePath = accessDbPath ' 設定目標資料庫路徑

    ' 4. 刪除舊表 (如果存在) - DAL 內部會檢查是否存在
    Debug.Print "Attempting to drop table: " & tableName
    dal.DropTable tableName

    ' 5. 使用 Text Driver 透過 SQL 直接匯入
    '    確保 Schema.ini 中 CharacterSet 與 csvEncoding 匹配，或在此處指定
    '    HDR=YES 表示第一行為標頭
    '    FMT=Delimited 表示逗號分隔 (如果不是逗號，需修改)
    '    注意：欄位名稱將由 Text Driver 根據 CSV 標頭或 Schema.ini 決定
    sql = "SELECT * INTO [" & tableName & "] " & _
          "FROM [Text;Database=" & csvFolder & ";HDR=YES;FMT=Delimited;CharacterSet=" & csvEncoding & ";IMEX=1].[" & csvFile & "]"
          ' IMEX=1 有助於將混合數據類型作為文本讀取，減少類型轉換錯誤

    Debug.Print "Executing SQL for import: " & sql
    dal.ExecuteSQL sql ' DAL 負責連接和執行

    ' 6. 成功
    ImportToAccess = True
    Debug.Print "Successfully imported CSV to Access table: " & tableName

CleanUp:
    ' 清理 DAL 物件 (其 Terminate 事件會處理 Disconnect)
    Set dal = Nothing
    Set fso = Nothing
    Exit Function

ErrorHandler:
    ImportToAccess = False ' 確認設為失敗
    Debug.Print "--- ImportToAccess Error ---"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "Error Description: " & Err.Description
    ' 將錯誤向上層 (cApplication) 拋出
    Err.Raise Err.Number, MODULE_NAME & "." & METHOD_NAME, Err.Description
    GoTo CleanUp ' 確保資源被釋放
End Function

Private Sub Class_Terminate()

End Sub
