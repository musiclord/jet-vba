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
    
    Dim diff As Long ' 用於儲存差異資料表的記錄數
    Dim recordsAffected As Long ' 用於檢查 ExecuteSQL 的結果
    Dim diffTableName As String
    diffTableName = Me.COMPLETE_DIFF
    
    On Error GoTo ErrorHandler
    TestCompleteness = False ' 預設失敗或找到差異
    diff = -1 ' 預設錯誤或未執行
    
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " : 嘗試刪除資料表: " & diffTableName
    m_AccessDAL.DropTable diffTableName

    ' --- 1. 獲取欄位映射 ---
    Set glMappings = m_MappingService.GetGLFieldMappings()
    glAccountNumber = glMappings("AccountNumber")
    glPostDate = glMappings("PostDate")
    glDebit = glMappings("Debit")
    glCredit = glMappings("Credit")
    Debug.Print MODULE_NAME & "." & METHOD_NAME & "-->GL 映射: 科目=[" & glAccountNumber & "], 日期=[" & glPostDate & "], 借方=[" & glDebit & "], 貸方=[" & glCredit & "]"

    Set tbMappings = m_MappingService.GetTBFieldMappings()
    tbAccountNumber = tbMappings("AccountNumber")
    tbOpening = tbMappings("Opening")
    tbClosing = tbMappings("Closing")
    tbAccountName = tbMappings("AccountName")
    Debug.Print MODULE_NAME & "." & METHOD_NAME & "-->TB 映射: 科目=[" & tbAccountNumber & "], 期初=[" & tbOpening & "], 期末=[" & tbClosing & "]"

    ' --- 2. 準備資料 ---
    Dim sStart As String
    sStart = Format(startDate, "yyyy\-MM\-dd")
    Dim sEnd As String
    sEnd = Format(endDate, "yyyy\-MM\-dd")
    
    ' --- 3. SQL查詢 ---
    Dim sql As String
    sql = "SELECT * INTO [" & diffTableName & "] FROM [" & m_GLTable & "] AS GLD " & _
          "WHERE NOT (GLD.[" & glPostDate & "] BETWEEN #" & sStart & "# AND #" & sEnd & "#)"
    
    Debug.Print MODULE_NAME & ".TestCompleteness - 執行 SQL: " & vbCrLf & sql
    recordsAffected = m_AccessDAL.ExecuteSQL(sql)
    
    ' --- 4. 透過計算新資料表中的記錄數來檢查是否找到差異 ---
    diff = m_AccessDAL.GetRecordCount(diffTableName)
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - " & diffTableName & " 中的記錄數: " & diff
    
    If diff = -1 Then
        TestCompleteness = False
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
    TestCompleteness = False
    Debug.Print "錯誤於 " & MODULE_NAME & "." & METHOD_NAME & " (行號: " & Erl & "): " & Err.Description & " (錯誤碼: " & Err.Number & ")"
    Err.Raise Err.Number, MODULE_NAME & "." & METHOD_NAME, Err.Description ' 將錯誤向上拋給呼叫者
    GoTo CleanUp
End Function

Private Sub Class_Terminate()
    Set m_AccessDAL = Nothing
    Set m_MappingService = Nothing
    Debug.Print MODULE_NAME & " Class_Terminate"
End Sub

