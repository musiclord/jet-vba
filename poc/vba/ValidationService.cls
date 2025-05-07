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
Private m_GLTable As String ' Access 中 GL 資料的實際資料表名稱
Private m_TBTable As String ' Access 中 TB 資料的實際資料表名稱

' 提供一個公共的唯讀屬性來訪問差異資料表的名稱
Public Property Get COMPLETE_DIFF() As String
    COMPLETE_DIFF = "COMPLETE_DIFF" ' 直接在此處定義字串值
End Property

' 初始化方法，用於注入依賴項和資料表名稱
Public Sub Initialize(ByVal dal As AccessDAL, ByVal ms As MappingService, ByVal glImportTableName As String, ByVal tbImportTableName As String)
    Const METHOD_NAME As String = "Initialize"
    On Error GoTo ErrorHandler

    Set m_AccessDAL = dal
    Set m_MappingService = ms
    m_GLTable = glImportTableName ' 使用傳入的 GL 資料表名稱
    m_TBTable = tbImportTableName ' 使用傳入的 TB 資料表名稱

    Debug.Print MODULE_NAME & "." & METHOD_NAME & " 初始化成功，GL 資料表: " & m_GLTable & ", TB 資料表: " & m_TBTable
    Exit Sub
ErrorHandler:
    Debug.Print "錯誤於 " & MODULE_NAME & "." & METHOD_NAME & ": " & Err.Description
    Err.Raise Err.Number, MODULE_NAME & "." & METHOD_NAME, Err.Description ' 將錯誤向上拋給呼叫者
End Sub

' 執行完整性測試
' 如果完整 (無差異) 則返回 True，否則返回 False。
' 如果存在差異，則會產生一個名為 COMPLETE_DIFF (透過 Property Get 訪問) 的資料表，其中包含差異明細。
Public Function TestCompleteness(ByVal startDate As Date, ByVal endDate As Date) As Boolean
    Const METHOD_NAME As String = "TestCompleteness"
    Dim glMappings As Object 'Scripting.Dictionary
    Dim tbMappings As Object 'Scripting.Dictionary
    Dim glAccountNumber As String, glPostDate As String, glDebit As String, glCredit As String
    Dim tbAccountNumber As String, tbOpening As String, tbClosing As String, tbAccountName As String
    
    Dim sql As String
    Dim diff As Long ' 用於儲存差異資料表的記錄數
    Dim recordsAffected As Long ' 用於檢查 ExecuteSQL 的結果
    Dim diffTableName As String

    On Error GoTo ErrorHandler
    TestCompleteness = False ' 預設為失敗或找到差異
    diff = -1 ' 預設為錯誤或未執行的狀態
    diffTableName = Me.COMPLETE_DIFF ' 透過 Property Get 獲取差異資料表的名稱

    ' --- 0. 驗證依賴項和資料庫連接 ---
    If m_AccessDAL Is Nothing Or m_MappingService Is Nothing Then
        Err.Raise vbObjectError + 512, MODULE_NAME & "." & METHOD_NAME, "服務未正確初始化。請先呼叫 Initialize。"
        GoTo CleanUp
    End If
    If Not m_AccessDAL.IsConnected Then
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - DAL 未連接，嘗試連接..."
        If Not m_AccessDAL.Connect Then
             Err.Raise vbObjectError + 513, MODULE_NAME & "." & METHOD_NAME, "透過 AccessDAL 連接資料庫失敗。"
             GoTo CleanUp
        End If
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - DAL 連接成功。"
    End If

    ' --- 1. 獲取欄位映射 ---
    Set glMappings = m_MappingService.GetGLFieldMappings()
    If glMappings Is Nothing Then Err.Raise vbObjectError + 516, MODULE_NAME & "." & METHOD_NAME, "從 MappingService 獲取 GL 欄位映射失敗。"
    If Not glMappings.Exists("AccountNumber") Then Err.Raise vbObjectError + 517, MODULE_NAME & "." & METHOD_NAME, "GL 欄位映射 'AccountNumber' 缺失。"
    If Not glMappings.Exists("PostDate") Then Err.Raise vbObjectError + 517, MODULE_NAME & "." & METHOD_NAME, "GL 欄位映射 'PostDate' 缺失。"
    If Not glMappings.Exists("Debit") Then Err.Raise vbObjectError + 517, MODULE_NAME & "." & METHOD_NAME, "GL 欄位映射 'Debit' 缺失。"
    If Not glMappings.Exists("Credit") Then Err.Raise vbObjectError + 517, MODULE_NAME & "." & METHOD_NAME, "GL 欄位映射 'Credit' 缺失。"
    glAccountNumber = glMappings("AccountNumber")
    glPostDate = glMappings("PostDate")
    glDebit = glMappings("Debit")
    glCredit = glMappings("Credit")
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - GL 映射: 科目=[" & glAccountNumber & "], 日期=[" & glPostDate & "], 借方=[" & glDebit & "], 貸方=[" & glCredit & "]"

    Set tbMappings = m_MappingService.GetTBFieldMappings()
    If tbMappings Is Nothing Then Err.Raise vbObjectError + 518, MODULE_NAME & "." & METHOD_NAME, "從 MappingService 獲取 TB 欄位映射失敗。"
    If Not tbMappings.Exists("AccountNumber") Then Err.Raise vbObjectError + 519, MODULE_NAME & "." & METHOD_NAME, "TB 欄位映射 'AccountNumber' 缺失。"
    If Not tbMappings.Exists("Opening") Then Err.Raise vbObjectError + 519, MODULE_NAME & "." & METHOD_NAME, "TB 欄位映射 'Opening' 缺失。"
    If Not tbMappings.Exists("Closing") Then Err.Raise vbObjectError + 519, MODULE_NAME & "." & METHOD_NAME, "TB 欄位映射 'Closing' 缺失。"
    tbAccountNumber = tbMappings("AccountNumber")
    tbOpening = tbMappings("Opening")
    tbClosing = tbMappings("Closing")
    tbAccountName = tbMappings("AccountName")
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - TB 映射: 科目=[" & tbAccountNumber & "], 期初=[" & tbOpening & "], 期末=[" & tbClosing & "]"

    ' --- 2. 準備資料庫：刪除已存在的差異資料表 ---
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 嘗試刪除資料表: " & diffTableName
    m_AccessDAL.DropTable diffTableName ' DropTable 應能處理資料表不存在的情況而不會引發錯誤

    ' --- 3. 建構並執行 SQL 查詢 (改用 IIF(IsNull(...),0,...) 取代 Nz) ---
    Dim sql2 As String
    sql2 = _
      "SELECT " & _
      "  TB.[" & EscapeSQLField(tbAccountNumber) & "] AS ACCOUNT_NUM_ALL, " & _
      "  TB.[" & EscapeSQLField(tbAccountName) & "] AS 會計科目名稱_TB, " & _
      "  (IIF(TB.[" & EscapeSQLField(tbClosing) & "] IS NULL, 0, TB.[" & EscapeSQLField(tbClosing) & "]) " & _
      "   - IIF(TB.[" & EscapeSQLField(tbOpening) & "] IS NULL, 0, TB.[" & EscapeSQLField(tbOpening) & "])) AS Period_Change_TB, " & _
      "  IIF(GL.GL_Sum IS NULL, 0, GL.GL_Sum) AS Net_Amount_GL, " & _
      "  (IIF(TB.[" & EscapeSQLField(tbClosing) & "] IS NULL, 0, TB.[" & EscapeSQLField(tbClosing) & "]) " & _
      "   - IIF(TB.[" & EscapeSQLField(tbOpening) & "] IS NULL, 0, TB.[" & EscapeSQLField(tbOpening) & "])) " & _
      "   - IIF(GL.GL_Sum IS NULL, 0, GL.GL_Sum) AS DIFF " & _
      "INTO [" & diffTableName & "] " & _
      "FROM [" & m_TBTable & "] AS TB " & _
      "LEFT JOIN ( " & _
      "    SELECT [" & EscapeSQLField(glAccountNumber) & "] AS AccountKey, " & _
      "           SUM( IIF([" & EscapeSQLField(glDebit) & "] IS NULL, 0, [" & EscapeSQLField(glDebit) & "]) " & _
      "             - IIF([" & EscapeSQLField(glCredit) & "] IS NULL, 0, [" & EscapeSQLField(glCredit) & "]) ) AS GL_Sum " & _
      "    FROM [" & m_GLTable & "] " & _
      "    WHERE [" & EscapeSQLField(glPostDate) & "] BETWEEN #" & Format(startDate, "yyyy\-MM\-dd") & "# AND #" & Format(endDate, "yyyy\-MM\-dd") & "# " & _
      "    GROUP BY [" & EscapeSQLField(glAccountNumber) & "] " & _
      ") AS GL ON TB.[" & EscapeSQLField(tbAccountNumber) & "] = GL.AccountKey"

    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 執行 SQL: " & sql2
    recordsAffected = m_AccessDAL.ExecuteSQL(sql2)
    ' --- 4. 透過計算新資料表中的記錄數來檢查是否找到差異 ---
    diff = m_AccessDAL.GetRecordCount(diffTableName)
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - " & diffTableName & " 中的記錄數: " & diff
    
    If diff = -1 Then
        TestCompleteness = False ' 保持 TestCompleteness 為 False，因為獲取記錄數失敗
        Err.Raise vbObjectError + 520, MODULE_NAME & "." & METHOD_NAME, "獲取差異資料表 '" & diffTableName & "' 的記錄數失敗。"
    ElseIf diff = 0 Then
        TestCompleteness = True ' 沒有差異記錄，測試通過
    Else
        TestCompleteness = False ' 存在差異記錄，測試未通過
    End If

CleanUp:
    Set glMappings = Nothing
    Set tbMappings = Nothing
    Exit Function

ErrorHandler:
    TestCompleteness = False ' 確保在任何錯誤情況下都返回 False
    Debug.Print "錯誤於 " & MODULE_NAME & "." & METHOD_NAME & " (行號: " & Erl & "): " & Err.Description & " (錯誤碼: " & Err.Number & ")"
    Err.Raise Err.Number, MODULE_NAME & "." & METHOD_NAME, Err.Description ' 將錯誤向上拋給呼叫者
    GoTo CleanUp ' 雖然 Err.Raise 會中止，但保留 GoTo CleanUp 以防萬一
End Function

' 輔助函數，用於轉義可能包含空格或特殊字元的欄位名稱
Private Function EscapeSQLField(ByVal fieldName As String) As String
    Dim needsEscaping As Boolean
    needsEscaping = False ' 預設為不需要轉義

    If fieldName = "" Then
        EscapeSQLField = "" ' 處理空字串的情況
        Exit Function
    End If

    ' 檢查是否包含任何需要轉義的字元
    Dim specialChars As Variant
    Dim i As Long
    specialChars = Array(" ", "[", "]", """", "'", "-", ".", "#", "$", "%", "&", "*", "+", "/", ":", ";", "<", "=", ">", "?", "@", "\", "^", "`", "{", "|", "}", "~")

    For i = LBound(specialChars) To UBound(specialChars)
        If InStr(1, fieldName, CStr(specialChars(i))) > 0 Then
            needsEscaping = True
            Exit For ' 找到任何一個特殊字元就足夠了
        End If
    Next i
    
    If needsEscaping Then
        EscapeSQLField = "[" & Replace(fieldName, "]", "]]") & "]"
    Else
        EscapeSQLField = fieldName
    End If
End Function

Private Sub Class_Initialize()
    Debug.Print MODULE_NAME & " Class_Initialize"
End Sub

Private Sub Class_Terminate()
    Set m_AccessDAL = Nothing
    Set m_MappingService = Nothing
    Debug.Print MODULE_NAME & " Class_Terminate"
End Sub

