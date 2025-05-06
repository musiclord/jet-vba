VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "ValidationService"

Private m_AccessDAL As AccessDAL
Private m_MappingService As MappingService
Private m_GLTableName As String ' Access 中 GL 資料的實際資料表名稱
Private m_TBTableName As String ' Access 中 TB 資料的實際資料表名稱

' 初始化方法，用於注入依賴項和資料表名稱
Public Sub Initialize(ByVal dal As AccessDAL, ByVal ms As MappingService, ByVal glImportTableName As String, ByVal tbImportTableName As String)
    Const METHOD_NAME As String = "Initialize"
    On Error GoTo ErrorHandler

    Set m_AccessDAL = dal
    Set m_MappingService = ms
    m_GLTableName = glImportTableName
    m_TBTableName = tbImportTableName

    If m_AccessDAL Is Nothing Then Err.Raise vbObjectError + 513, MODULE_NAME & "." & METHOD_NAME, "AccessDAL 物件未提供。"
    If m_MappingService Is Nothing Then Err.Raise vbObjectError + 513, MODULE_NAME & "." & METHOD_NAME, "MappingService 物件未提供。"
    If Len(m_GLTableName) = 0 Then Err.Raise vbObjectError + 514, MODULE_NAME & "." & METHOD_NAME, "GL 資料表名稱未提供。"
    If Len(m_TBTableName) = 0 Then Err.Raise vbObjectError + 515, MODULE_NAME & "." & METHOD_NAME, "TB 資料表名稱未提供。"

    Debug.Print MODULE_NAME & "." & METHOD_NAME & " 初始化成功。"
    Exit Sub
ErrorHandler:
    Debug.Print "錯誤於 " & MODULE_NAME & "." & METHOD_NAME & ": " & Err.Description
    Err.Raise Err.Number, MODULE_NAME & "." & METHOD_NAME, Err.Description
End Sub

' 執行完整性測試
' 如果完整 (無差異) 則返回 True，否則返回 False。
' 如果存在差異，則會產生一個名為 COMPLETENESS_DIFFERENCES_TABLE_NAME 的資料表，其中包含差異明細。
Public Function TestCompleteness(ByVal startDate As Date, ByVal endDate As Date) As Boolean
    Const METHOD_NAME As String = "TestCompleteness"
    Dim glMappings As Object 'Scripting.Dictionary
    Dim tbMappings As Object 'Scripting.Dictionary
    Dim glAccountField As String, glAmountField As String, glDateField As String
    Dim tbAccountField As String, tbChangeField As String ' 假設這是 TB 中的 "變動金額" 欄位
    Dim sql As String
    Dim recordsAffected As Long
    Dim differenceCount As Long

    On Error GoTo ErrorHandler
    TestCompleteness = False ' 預設為失敗或找到差異

    ' --- 1. 驗證依賴項 ---
    If m_AccessDAL Is Nothing Or m_MappingService Is Nothing Then
        Err.Raise vbObjectError + 512, MODULE_NAME & "." & METHOD_NAME, "服務尚未初始化，請先呼叫 Initialize。"
        GoTo CleanUp
    End If
    If Not m_AccessDAL.IsConnected Then
        MsgBox "資料庫尚未連接，無法執行完整性測試。", vbCritical, "資料庫錯誤"
        GoTo CleanUp
    End If

    ' --- 2. 獲取欄位對應 ---
    Set glMappings = m_MappingService.GetGLFieldMappings()
    Set tbMappings = m_MappingService.GetTBFieldMappings()

    If glMappings Is Nothing Or glMappings.Count = 0 Then
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 錯誤: GL 欄位對應未設定。"
        MsgBox "GL 欄位對應尚未設定，無法執行完整性測試。", vbExclamation, "欄位對應錯誤"
        GoTo CleanUp
    End If
    If tbMappings Is Nothing Or tbMappings.Count = 0 Then
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 錯誤: TB 欄位對應未設定。"
        MsgBox "TB 欄位對應尚未設定，無法執行完整性測試。", vbExclamation, "欄位對應錯誤"
        GoTo CleanUp
    End If

    ' 檢索對應的欄位名稱 (確保這些標準鍵在 MappingService 中存在)
    ' 這些標準鍵 ("AccountNumber", "Amount", "PostDate", "ChangeAmount") 必須在
    ' cApplication 呼叫 MappingService.SetGLFieldMappings 和 SetTBFieldMappings 時一致使用。
    On Error Resume Next ' 暫時忽略字典鍵不存在的錯誤，稍後檢查
    glAccountField = glMappings("AccountNumber")
    glAmountField = glMappings("Amount")
    glDateField = glMappings("PostDate")
    tbAccountField = tbMappings("AccountNumber")
    tbChangeField = tbMappings("ChangeAmount") ' 重要: 假設 TB 對應中有 "ChangeAmount"
    On Error GoTo ErrorHandler ' 恢復正常錯誤處理

    If Len(glAccountField) = 0 Or Len(glAmountField) = 0 Or Len(glDateField) = 0 Or Len(tbAccountField) = 0 Or Len(tbChangeField) = 0 Then
        Dim missingFields As String
        If Len(glAccountField) = 0 Then missingFields = missingFields & "GL AccountNumber, "
        If Len(glAmountField) = 0 Then missingFields = missingFields & "GL Amount, "
        If Len(glDateField) = 0 Then missingFields = missingFields & "GL PostDate, "
        If Len(tbAccountField) = 0 Then missingFields = missingFields & "TB AccountNumber, "
        If Len(tbChangeField) = 0 Then missingFields = missingFields & "TB ChangeAmount, "
        If Len(missingFields) > 0 Then missingFields = Left(missingFields, Len(missingFields) - 2) ' Remove trailing comma

        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 錯誤: 以下必要的欄位對應缺失: " & missingFields
        MsgBox "部分必要的 GL 或 TB 欄位對應缺失 (" & missingFields & ")，無法執行完整性測試。", vbExclamation, "欄位對應缺失"
        GoTo CleanUp
    End If

    ' --- 3. 準備資料庫: 如果存在舊的差異表，則刪除 ---
    If m_AccessDAL.TableExists(COMPLETENESS_DIFFERENCES_TABLE_NAME) Then
        If Not m_AccessDAL.ExecuteSQL("DROP TABLE [" & COMPLETENESS_DIFFERENCES_TABLE_NAME & "]") Then
            Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 錯誤: 無法刪除已存在的資料表 '" & COMPLETENESS_DIFFERENCES_TABLE_NAME & "'。"
            MsgBox "無法刪除舊的差異表 '" & COMPLETENESS_DIFFERENCES_TABLE_NAME & "'。", vbCritical, "資料庫錯誤"
            GoTo CleanUp
        End If
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 已刪除存在的資料表: " & COMPLETENESS_DIFFERENCES_TABLE_NAME
    End If

    ' --- 4. 建構並執行 SQL 以查找並儲存差異 ---
    ' 使用 UNION ALL 合併 GL 總和與 TB 變動，然後按科目分組，
    ' 並將存在差異的科目選入新資料表。
    sql = "SELECT " & _
          "    Combined.AccountKey AS AccountNo, " & _
          "    SUM(Combined.GL_Sum_Period) AS GL_Total_Amount, " & _
          "    SUM(Combined.TB_Change_Period) AS TB_Total_Change, " & _
          "    Nz(SUM(Combined.GL_Sum_Period), 0) - Nz(SUM(Combined.TB_Change_Period), 0) AS Difference " & _
          "INTO [" & COMPLETENESS_DIFFERENCES_TABLE_NAME & "] " & _
          "FROM ( " & _
          "    SELECT " & _
          "        [" & EscapeSQLField(glAccountField) & "] AS AccountKey, " & _
          "        SUM([" & EscapeSQLField(glAmountField) & "]) AS GL_Sum_Period, " & _
          "        0 AS TB_Change_Period " & _
          "    FROM [" & m_GLTableName & "] " & _
          "    WHERE [" & EscapeSQLField(glDateField) & "] >= #" & Format(startDate, "yyyy-mm-dd") & "# AND [" & EscapeSQLField(glDateField) & "] <= #" & Format(endDate, "yyyy-mm-dd") & "# " & _
          "    GROUP BY [" & EscapeSQLField(glAccountField) & "] " & _
          "    UNION ALL " & _
          "    SELECT " & _
          "        [" & EscapeSQLField(tbAccountField) & "] AS AccountKey, " & _
          "        0 AS GL_Sum_Period, " & _
          "        SUM([" & EscapeSQLField(tbChangeField) & "]) AS TB_Change_Period " & _
          "    FROM [" & m_TBTableName & "] " & _
          "    GROUP BY [" & EscapeSQLField(tbAccountField) & "] " & _
          ") AS Combined " & _
          "GROUP BY Combined.AccountKey " & _
          "HAVING Nz(SUM(Combined.GL_Sum_Period), 0) - Nz(SUM(Combined.TB_Change_Period), 0) <> 0"

    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 執行 SQL: " & sql

    If Not m_AccessDAL.ExecuteSQL(sql, recordsAffected) Then
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 錯誤: SQL 執行失敗，無法建立差異表。"
        MsgBox "執行完整性測試 SQL 失敗，無法產生差異表。", vbCritical, "SQL 執行錯誤"
        GoTo CleanUp
    End If
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 差異表 SQL 已執行。SELECT INTO 影響的記錄數: " & recordsAffected

    ' --- 5. 檢查是否找到差異 ---
    differenceCount = m_AccessDAL.GetRecordCount(COMPLETENESS_DIFFERENCES_TABLE_NAME)
    If differenceCount = -1 Then ' GetRecordCount 在錯誤時返回 -1 (例如，資料表未找到)
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 錯誤: 無法獲取 '" & COMPLETENESS_DIFFERENCES_TABLE_NAME & "' 的記錄數。"
        MsgBox "無法讀取差異表 '" & COMPLETENESS_DIFFERENCES_TABLE_NAME & "' 的記錄數量。", vbCritical, "資料庫錯誤"
        TestCompleteness = False ' 指示失敗
    ElseIf differenceCount = 0 Then
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 測試完成。未發現差異。"
        TestCompleteness = True ' 測試通過，無差異
    Else
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 測試完成。在 '" & COMPLETENESS_DIFFERENCES_TABLE_NAME & "' 中找到 " & differenceCount & " 筆差異。"
        TestCompleteness = False ' 測試失敗，找到差異
    End If

CleanUp:
    Set glMappings = Nothing
    Set tbMappings = Nothing
    Exit Function

ErrorHandler:
    TestCompleteness = False ' 確保錯誤時返回 False
    Debug.Print "!!! 錯誤於 " & MODULE_NAME & "." & METHOD_NAME & " !!!"
    Debug.Print "錯誤號碼: " & Err.Number
    Debug.Print "錯誤來源: " & Err.Source
    Debug.Print "錯誤描述: " & Err.Description
    MsgBox "在 " & MODULE_NAME & "." & METHOD_NAME & " 中發生錯誤：" & vbCrLf & Err.Description, vbCritical, "執行錯誤"
    GoTo CleanUp
End Function

' 輔助函數，用於轉義可能包含空格或特殊字元的欄位名稱
Private Function EscapeSQLField(ByVal fieldName As String) As String
    If InStr(1, fieldName, " ") > 0 Or InStr(1, fieldName, "[") > 0 Or InStr(1, fieldName, "]") > 0 Or InStr(1, fieldName, """") > 0 Or InStr(1, fieldName, "'") > 0 Or InStr(1, fieldName, "-") > 0 Or InStr(1, fieldName, ".") > 0 Then
        EscapeSQLField = "[" & Replace(fieldName, "]", "]]") & "]" ' 基本的括號處理，可以更穩健
    Else
        EscapeSQLField = fieldName
    End If
End Function

Private Sub Class_Initialize()
    ' m_AccessDAL, m_MappingService 和資料表名稱預計會透過 Public Initialize Sub 設定。
    Debug.Print MODULE_NAME & " Class_Initialize"
End Sub

Private Sub Class_Terminate()
    Set m_AccessDAL = Nothing
    Set m_MappingService = Nothing
    Debug.Print MODULE_NAME & " Class_Terminate"
End Sub

