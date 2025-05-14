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

Public Sub addLineItem()
    Dim sql As String
    Dim affected As Long
    
    ' 新增欄位
    'sql = "ALTER TABLE GL ADD COLUMN 項次 INTEGER;"
    'affected = m_AccessDAL.ExecuteSQL(sql)
    ' 新增項次
    'sql = ""
    'affected = m_AccessDAL.ExecuteSQL(sql)
    ' 借-貸
    sql = "ALTER TABLE GL ADD COLUMN 傳票金額 DOUBLE;"
    affected = m_AccessDAL.ExecuteSQL(sql)
    sql = "UPDATE GL SET GL.[傳票金額] = GL.[本幣借方金額] - GL.[本幣貸方金額]"
    affected = m_AccessDAL.ExecuteSQL(sql)
End Sub
' 執行完整性測試
' 如果完整 (無差異) 則返回 True，否則返回 False。
' 如果存在差異，則會產生一個名為 COMPLETE_DIFF (透過 Property Get 訪問) 的資料表，其中包含差異明細。
Public Function TestCompleteness(ByVal startDate As Date, ByVal endDate As Date) As Boolean
    Const METHOD_NAME As String = "TestCompleteness"
    Dim glMappings As Object 'Scripting.Dictionary
    Dim tbMappings As Object 'Scripting.Dictionary
    Dim diffTableName As String
    Dim diff As Long ' 用於儲存差異資料表的記錄數
    Dim affected As Long ' 用於檢查 ExecuteSQL 的結果
    diffTableName = Me.COMPLETE_DIFF
    
    On Error GoTo ErrorHandler
    TestCompleteness = False ' 預設失敗或找到差異
    diff = -1 ' 預設錯誤或未執行
    
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " : 嘗試刪除資料表: " & diffTableName
    m_AccessDAL.DropTable diffTableName

    ' --- 1. 獲取欄位映射 ---
    ' GL 欄位變數
    Set glMappings = m_MappingService.GetGLFieldMappings()
    Dim glDocumentNumber As String
    glDocumentNumber = glMappings("DocumentNumber")
    Dim glPostDate As String
    glPostDate = glMappings("PostDate")
    Dim glAccountNumber As String
    glAccountNumber = glMappings("AccountNumber")
    Dim glAccountName As String
    glAccountName = glMappings("AccountName")
    Dim glDescription As String
    glDescription = glMappings("Description")
    Dim glAmount As String
    glAmount = glMappings("Amount")
    Dim glDebit As String
    glDebit = glMappings("Debit")
    Dim glCredit As String
    glCredit = glMappings("Credit")
    Dim glDebitFlag As String
    glDebitFlag = glMappings("DebitFlag")

    ' TB 欄位變數
    Set tbMappings = m_MappingService.GetTBFieldMappings()
    Dim tbAccountNumber As String
    tbAccountNumber = tbMappings("AccountNumber")
    Dim tbAccountName As String
    tbAccountName = tbMappings("AccountName")
    Dim tbCredit As String
    tbCredit = tbMappings("Credit")
    Dim tbDebit As String
    tbDebit = tbMappings("Debit")

    ' --- 2. 準備資料 ---
    Dim sStart As String
    sStart = Format(startDate, "yyyy\-MM\-dd")
    Dim sEnd As String
    sEnd = Format(endDate, "yyyy\-MM\-dd")
    
    ' --- 3. SQL查詢 ---
    Dim sql As String
    Dim GL_In_Period As String
    
    ' -------> Record Extraction - @BetweenDate()
    GL_In_Period = "GL_In_Period"
    ' SELECT GL.* INTO GL_In_Period FROM GL
    ' WHERE CDate(GL.[日期]) >= 2024-01-01AND CDate(GL.[日期]) <= 2024-12-31;
    sql = "SELECT " & m_GLTable & ".* INTO " & GL_In_Period & _
            " FROM " & m_GLTable & " WHERE CDate(" & m_GLTable & ".[" & glPostDate & "]) " & _
            ">= " & sStart & "AND CDate(" & m_GLTable & ".[" & glPostDate & "]) " & _
            "<= " & sEnd & ";"
    
    If m_AccessDAL.TableExists(GL_In_Period) Then
        Debug.Print MODULE_NAME & ".TestCompleteness - 資料表 '" & GL_In_Period & "' 已存在，將其刪除。"
        m_AccessDAL.DropTable GL_In_Period
    Else
        Debug.Print MODULE_NAME & ".TestCompleteness - 資料表 '" & GL_In_Period & "' 不存在，無需刪除。"
    End If
    Debug.Print MODULE_NAME & ".TestCompleteness - 執行 SQL: " & vbCrLf & sql
    affected = m_AccessDAL.ExecuteSQL(sql)
    Debug.Print MODULE_NAME & ".TestCompleteness - 不在日期區間: " & affected
    
    ' -------> Summarization - AddFieldToSummarize "會計科目編號_JE", AddFieldToTotal "傳票金額_JE"
    Dim GL_Account_Sum As String, Sum_傳票金額 As String
    GL_Account_Sum = "GL_Account_Sum"
    Sum_傳票金額 = "Sum_傳票金額"
    ' SELECT GL.科目代碼, SUM(GL.傳票金額) AS Sum_傳票金額
    ' INTO GL_Account_Sum FROM GL GROUP BY GL.科目代碼;
    sql = "SELECT " & m_GLTable & "." & glAccountNumber & _
          ", SUM(" & m_GLTable & "." & glAmount & ") AS " & Sum_傳票金額 & _
          " INTO " & GL_Account_Sum & " FROM " & m_GLTable & _
          " GROUP BY " & m_GLTable & "." & glAccountNumber & ";"
    
    If m_AccessDAL.TableExists(GL_Account_Sum) Then
        Debug.Print MODULE_NAME & ".TestCompleteness - 資料表 '" & GL_Account_Sum & "' 已存在，將其刪除。"
        m_AccessDAL.DropTable GL_Account_Sum
    Else
        Debug.Print MODULE_NAME & ".TestCompleteness - 資料表 '" & GL_Account_Sum & "' 不存在，無需刪除。"
    End If
    Debug.Print MODULE_NAME & ".TestCompleteness - 執行 SQL: " & vbCrLf & sql
    affected = m_AccessDAL.ExecuteSQL(sql)
    Debug.Print MODULE_NAME & ".TestCompleteness - 依科目代碼分類加總: " & affected
    
    ' -------> Join Databases, primary: #GL_Account_Sum.IDM, secondary: #TB#.IDM
    Dim Completeness_calculate As String, LeftJ As String, RightJ As String
    Completeness_calculate = "Completeness_calculate"
    LeftJ = "temp_Left"
    RightJ = "temp_Right"
    If m_AccessDAL.TableExists(Completeness_calculate) Then
        m_AccessDAL.DropTable Completeness_calculate
    ElseIf m_AccessDAL.TableExists(LeftJ) Then
        m_AccessDAL.DropTable LeftJ
    ElseIf m_AccessDAL.TableExists(RightJ) Then
        m_AccessDAL.DropTable RightJ
    Else
        Debug.Print MODULE_NAME & ".TestCompleteness - 聯集用資料表無需刪除。"
    End If
    '->LEFT JOIN
    'SELECT GL_Account_Sum.*, TB.* INTO LeftJ
    'FROM GL_Account_Sum LEFT JOIN TB ON GL_Account_Sum.科目代碼 = TB.會計項目;
    sql = "SELECT " & GL_Account_Sum & " .*, " & m_TBTable & ".* INTO " & LeftJ & vbCrLf & _
          "FROM " & GL_Account_Sum & " LEFT JOIN " & m_TBTable & " ON " & GL_Account_Sum & "." & glAccountNumber & " = " & m_TBTable & "." & tbAccountNumber & ";"
    affected = m_AccessDAL.ExecuteSQL(sql)
    '->RIGHT JOIN
    'SELECT GL_Account_Sum.*, TB.* INTO RightJ
    'FROM GL_Account_Sum RIGHT JOIN TB ON GL_Account_Sum.科目代碼 = TB.會計項目
    'WHERE GL_Account_Sum.科目代碼 IS NULL;
    sql = "SELECT " & GL_Account_Sum & ".*, " & m_TBTable & ".* INTO " & RightJ & vbCrLf & _
          "FROM " & GL_Account_Sum & " RIGHT JOIN " & m_TBTable & " ON " & GL_Account_Sum & "." & glAccountNumber & " = " & m_TBTable & "." & tbAccountNumber & vbCrLf & _
          "WHERE " & GL_Account_Sum & "." & glAccountNumber & " IS NULL;"
    affected = m_AccessDAL.ExecuteSQL(sql)
    '->UNION
    'SELECT * INTO Completeness_calculate FROM
    '(SELECT * FROM LJOIN_Account UNION SELECT * FROM RJOIN_Account);
    sql = "SELECT * INTO " & Completeness_calculate & " FROM" & vbCrLf & _
          "(SELECT * FROM " & LeftJ & " UNION SELECT * FROM " & RightJ & " );"
    affected = m_AccessDAL.ExecuteSQL(sql)
    
    
    
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

