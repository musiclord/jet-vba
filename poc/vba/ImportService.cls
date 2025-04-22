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

Private Const adDate As Long = 7 ' <-- 暫時保留，用於 LoadToExcel 的日期格式化
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

Public Function LoadToExcel(ByVal accessDbPath As String, ByVal tableName As String, ByVal worksheetName As String, ByVal maxRowsToShow As Long) As Boolean
    ' 重構後：使用 AccessDAL 獲取資料，保留 Excel 操作部分 (待進一步分離)
    Const METHOD_NAME As String = "LoadToExcel"
    Dim dal As AccessDAL
    Dim rs As Object   ' ADODB.Recordset (從 DAL 獲取)
    Dim ws As Worksheet
    Dim sql As String
    Dim i As Long
    Dim targetSheetExists As Boolean
    Dim fieldCount As Long
    Dim recordCount As Long ' 用於 CopyFromRecordset

    On Error GoTo ErrorHandler
    LoadToExcel = False ' 預設為失敗

    ' --- Excel 工作表處理部分 (暫時保留，理想情況應移至 Controller/View) ---
    Application.DisplayAlerts = False
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(worksheetName)
    targetSheetExists = (Err.Number = 0)
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True

    If targetSheetExists Then
        ws.Cells.Clear
        Debug.Print "已清空現有工作表: " & worksheetName
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = worksheetName
        On Error Resume Next ' 嘗試修改 CodeName
        ThisWorkbook.VBProject.VBComponents(ws.CodeName).name = worksheetName
        If Err.Number <> 0 Then Debug.Print "警告: 無法修改工作表 '" & worksheetName & "' 的 VBA CodeName。": Err.Clear
        On Error GoTo ErrorHandler
        Debug.Print "已建立新工作表: " & worksheetName
    End If
    ' --- Excel 工作表處理部分結束 ---


    ' 2. 建立 DAL 並獲取資料
    Set dal = New AccessDAL
    dal.DatabasePath = accessDbPath

    ' 3. 建立 SQL 語句
    If maxRowsToShow > 0 Then
        sql = "SELECT TOP " & maxRowsToShow & " * FROM [" & tableName & "]"
    Else
        sql = "SELECT * FROM [" & tableName & "]"
    End If

    ' 4. 從 DAL 獲取 Recordset (建議獲取 disconnected recordset)
    Set rs = dal.GetRecordset(sql, disconnectRecordset:=True) ' 獲取後即可斷開連接

    ' 5. 檢查 Recordset 是否有效
    If rs Is Nothing Then
        ' GetRecordset 內部出錯會 Raise，理論上不會執行到這裡，除非 DAL 返回 Nothing
        Err.Raise vbObjectError + 517, MODULE_NAME & "." & METHOD_NAME, "Failed to retrieve data from Access."
    End If

    ' --- 將資料寫入 Excel (暫時保留) ---
    If Not (rs.EOF And rs.BOF) Then ' 檢查是否有記錄
        rs.MoveFirst ' 確保指針在開頭 (對於 disconnected recordset 也是好習慣)
        fieldCount = rs.Fields.Count
        recordCount = rs.recordCount ' Disconnected recordset 可以直接獲取 RecordCount

        ' 6. 寫入標題
        For i = 0 To fieldCount - 1
            ws.Cells(1, i + 1).value = rs.Fields(i).name
        Next i

        ' 7. 複製資料
        If recordCount > 0 Then
            ws.Range("A2").CopyFromRecordset rs, recordCount ' 指定複製的記錄數
        End If

        ' 8. 處理日期格式 (可選)
        For i = 0 To fieldCount - 1
            If rs.Fields(i).Type = adDate Then ' adDate = 7
                If recordCount > 0 Then ' 確保有資料行
                     ws.Range("A2").Offset(0, i).Resize(recordCount, 1).NumberFormat = "yyyy/mm/dd"
                End If
            End If
        Next i

        ' 9. 調整欄寬
        ws.Columns.AutoFit
    Else
        ' 如果 Access 表格是空的，嘗試從 DAL 獲取欄位資訊 (如果 GetRecordset 即使空也返回帶欄位的 rs)
        If Not rs Is Nothing Then
             If rs.Fields.Count > 0 Then
                 For i = 0 To rs.Fields.Count - 1
                    ws.Cells(1, i + 1).value = rs.Fields(i).name
                Next i
                ws.Columns.AutoFit
             Else
                 ws.Cells(1, 1).value = "資料表 [" & tableName & "] 不存在或無法讀取欄位。"
             End If
        Else
             ws.Cells(1, 1).value = "無法讀取資料表 [" & tableName & "]。"
        End If
        MsgBox "Access 資料表 [" & tableName & "] 中沒有資料可載入。", vbInformation
    End If
    ' --- Excel 寫入部分結束 ---

    ' 10. 成功
    LoadToExcel = True

CleanUp:
    ' 清理資源
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close ' adStateOpen = 1, 雖然是 disconnected，以防萬一
    Set rs = Nothing
    Set dal = Nothing ' DAL 的 Terminate 會處理連接
    Set ws = Nothing
    Exit Function

ErrorHandler:
    LoadToExcel = False
    Debug.Print "--- LoadToExcel Error ---"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "Error Description: " & Err.Description
    Application.DisplayAlerts = True ' 確保提示被恢復
    Err.Raise Err.Number, MODULE_NAME & "." & METHOD_NAME, Err.Description
    GoTo CleanUp
End Function

