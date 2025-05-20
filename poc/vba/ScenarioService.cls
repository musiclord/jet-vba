VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScenarioService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "ScenarioService"
Private m_DbPath As String
Private m_dal As AccessDAL
Private m_JE1 As String
Private m_JE2 As String
Private m_JE3 As String

Public Sub Initialize(ByVal DbPath As String, ByVal je1 As String, ByVal je2 As String, ByVal je3 As String)
    Dim METHOD_NAME As String
    METHOD_NAME = "Initialize"
    
    On Error GoTo ErrorHandler
    Set m_dal = New AccessDAL
    m_DbPath = DbPath
    m_dal.DatabasePath = m_DbPath
    m_JE1 = je1
    m_JE2 = je2
    m_JE3 = je3
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " 初始化成功。使用資料庫: " & m_DbPath
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Description ' 將錯誤向上拋給呼叫者
End Sub

Public Function ExecuteScenario1()
    ' 短時間內高頻交易
    Const METHOD_NAME As String = "ExecuteScenario1"
    Dim success As Boolean
    
    If Not m_dal.Connect Then
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 錯誤: 無法連接到案例資料庫: " & m_dal.DatabasePath
        Exit Function
    End If '若無例外則連接成功
    
    ' Execute procedure... start
    Dim sql As String
    Dim sttDate As Date
    Dim num As Integer
    Dim currentStartDate As Date
    Dim currentEndDate As Date
    Dim tableToDrop As Variant
    Dim tbl As Variant
    
    On Error GoTo ErrorHandler
    
    ' --- 步驟 0: 清理/準備可能已存在的表格 ---
    tableToDrop = Array("deposit_credit", "date_counterparty_count", "day_error_count", "date_counterparty_sum", "day_error_sum", "week7+2cycle", "week_countandsum", "week_error_count", "week_error_sum", "month30+8cycle", "month_countandsum", "month_error_count", "month_error_sum")
    
    For Each tbl In tableToDrop
        On Error Resume Next
        m_dal.ExecuteSQL "DROP TABLE [" & tbl & "];"
    Next tbl
    Debug.Print MODULE_NAME & "." & METHOD_NAME & "- 舊表格清理完畢。"
    
    ' --- 步驟 1. 篩選銀行存款貸方交易紀錄 ---
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 1. 查詢銀行存款及貸方餘額:" & vbCrLf & _
                "FROM [JE_1] INTO [deposit_credit]"
    sql = "SELECT * INTO deposit_credit " & _
          "FROM [" & m_JE1 & "] " & _
          "WHERE [會計科目] = '銀行存款' AND [本幣貸方金額] <> 0;"
    m_dal.ExecuteSQL sql
    
    ' --- 步驟 2.1. 計算每日各交易對手交易筆數 ---
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 2.1. 同一日同一交易對手交易筆數:" & vbCrLf & _
                "FROM [deposit_credit] INTO [date_counterparty_count]"
    sql = "SELECT [日期], [交易對手帳戶], COUNT(*) AS [同一日同一交易對手交易筆數] " & _
          "INTO date_counterparty_count " & _
          "FROM deopsit_credit " & _
          "GROUP BY [日期], [交易對手帳戶];"
    m_dal.ExecuteSQL sql
    
    ' --- 步驟 2.2. 標記每日交易筆數異常分錄 ---
    ' 門檻值: 交易筆數 > 2
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 2.2. 標記筆數異常分錄(同一天):" & vbCrLf & _
                "FROM [deposit_credit], [date_counterparty_count] INTO [day_error_count]"
    sql = "SELECT [日期], [交易對手帳戶], COUNT(*) AS [同一日同一交易對手交易筆數] " & _
          "INTO date_counterparty_count " & _
          "FROM deopsit_credit " & _
          "GROUP BY [日期], [交易對手帳戶];"
    m_dal.ExecuteSQL sql
    
    ' --- 步驟 3.1. 計算每日各交易對手總交易金額 ---
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 3.1. 同一日同一交易對手交易金額:" & vbCrLf & _
                "FROM [deposit_credit] INTO [date_counterparty_sum]"
    sql = "SELECT [日期], [交易對手帳戶], SUM([本幣貸方金額]) AS [同一日同一交易對手交易金額 ] " & _
          "INTO date_counterparty_sum " & _
          "FROM deposit_credit " & _
          "GROUP BY [日期], [交易對手帳戶];"
    m_dal.ExecuteSQL sql
    
    ' --- 步驟 3.2. 標記每日交易金額異常分錄 ---
    ' 門檻值: 交易金額 > 500000
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 3.2. 標記金額異常分錄(同一天):" & vbCrLf & _
                "FROM [deposit_credit], [date_counterparty_sum] INTO [day_error_sum]"
    sql = "SELECT * INTO day_error_sum " & _
          "FROM (SELECT d1.*, t8.[同一日同一交易對手交易金額 ], " & _
          "IIF(t8.[同一日同一交易對手交易金額 ] > 500000, 'V', '') AS 超過金額門檻之異常分錄_日 " & _
          "FROM deposit_credit AS d1 LEFT JOIN date_counterparty_sum AS t8 " & _
          "ON (d1.[日期] = t8.[日期]) AND (d1.[交易對手帳戶] = t8.[交易對手帳戶])) AS A " & _
          "WHERE A.超過金額門檻之異常分錄_日='V';"
          
    m_dal.ExecuteSQL sql
    
    ' --- 步驟 4.1. 產生週滾動時間區間輔助表 ---
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 4.1. 產出7+2天迴圈:" & vbCrLf & _
                "INSERT INTO [week7+2cycle]"
    sql = "CREATE TABLE [week7+2cycle] (StartDate Date, EndDate Date);"
    m_dal.ExecuteSQL sql
    
    sttDate = #1/1/2024# ' 起始日期
    num = 0 ' 計數器
    Do While (DateAdd("d", num, sttDate) <= #12/31/2024#)
        currentStartDate = DateAdd("d", num, sttDate)
        currentEndDate = DateAdd("d", num + 8, sttDate) ' 9天區間 (0 到 8 共9天)
        If currentEndDate > #12/31/2024# And currentStartDate <= #12/31/2024# Then
             If DateDiff("d", currentStartDate, #12/31/2024#) >= 0 Then
                currentEndDate = #12/31/2024#
             Else
                Exit Do
             End If
        ElseIf currentStartDate > #12/31/2024# Then
            Exit Do
        End If
        sql = "INSERT INTO [week7+2cycle] (StartDate, EndDate) " & _
              "VALUES (#" & Format(currentStartDate, "yyyy-mm-dd") & "#, #" & Format(currentEndDate, "yyyy-mm-dd") & "#);"
        m_dal.ExecuteSQL sql
        num = num + 1 ' 每天滾動
    Loop
    
    ' --- 步驟 4.2. 計算週滾動區間交易筆數與金額 ---
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " -  4.2. week7+2_count&sum:" & vbCrLf & _
                "FROM [deposit_credit], [week7+2cycle] INTO [week_countandsum]"
    sql = "SELECT t1.[交易對手帳戶], t4.StartDate, t4.EndDate, " & _
          "COUNT(*) AS [連續7日同一交易對手交易筆數 ], " & _
          "SUM(t1.[本幣貸方金額]) AS [連續7日同一交易對手交易金額 ] " & _
          "INTO week_countandsum " & _
          "FROM deposit_credit AS t1 INNER JOIN [week7+2cycle] AS t4 " & _
          "ON (t1.[日期] >= t4.StartDate) AND (t1.[日期] <= t4.EndDate) " & _
          "GROUP BY t1.[交易對手帳戶], t4.StartDate, t4.EndDate;"
    m_dal.ExecuteSQL sql
    
    ' --- 步驟 4.3. 標記週區間交易筆數異常分錄 ---
    ' 門檻值: 交易筆數 > 10
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 4.3. 標記筆數異常分錄(週):" & vbCrLf & _
                "FROM [deposit_credit], [week_countandsum] INTO [week_error_count]"
    sql = "SELECT DISTINCT A.* INTO week_error_count " & _
          "FROM (SELECT t1.*, t5.[連續7日同一交易對手交易筆數 ], " & _
          "IIF(t5.[連續7日同一交易對手交易筆數 ] > 10, 'V', '') AS 超過筆數門檻之異常分錄_週 " & _
          "FROM deposit_credit AS t1 LEFT JOIN week_countandsum AS t5 " & _
          "ON (t1.[交易對手帳戶] = t5.[交易對手帳戶]) AND (t1.[日期] >= t5.StartDate) AND (t1.[日期] <= t5.EndDate)) AS A " & _
          "WHERE A.超過筆數門檻之異常分錄_週='V';"
    m_dal.ExecuteSQL sql
    
    ' --- 步驟 4.4. 標記週區間交易金額異常分錄 ---
    ' 門檻值: 交易金額 > 1000000
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 4.4. 標記金額異常分錄(週):" & vbCrLf & _
                "FROM [deposit_credit], [week_countandsum] INTO [week_error_sum]"
    sql = "SELECT DISTINCT A.* INTO week_error_sum " & _
          "FROM (SELECT t1.*, t5.[連續7日同一交易對手交易金額 ], " & _
          "IIF(t5.[連續7日同一交易對手交易金額 ] > 1000000, 'V', '') AS 超過金額門檻之異常分錄_週 " & _
          "FROM deposit_credit AS t1 LEFT JOIN week_countandsum AS t5 " & _
          "ON (t1.[交易對手帳戶] = t5.[交易對手帳戶]) AND (t1.[日期] >= t5.StartDate) AND (t1.[日期] <= t5.EndDate)) AS A " & _
          "WHERE A.超過金額門檻之異常分錄_週='V';"
    m_dal.ExecuteSQL sql
    
    ' --- 步驟 5.1. 產生月滾動時間區間輔助表 ---
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 5.1. 產出30+8天迴圈:" & vbCrLf & _
                "INSERT INTO [month30+8cycle]"
    sql = "CREATE TABLE [month30+8cycle] (StartDate Date, EndDate Date);"
    m_dal.ExecuteSQL sql
    
    sttDate = #1/1/2024# ' 起始日期
    num = 0 ' 計數器
    Do While (DateAdd("d", num, sttDate) <= #12/31/2024#)
        currentStartDate = DateAdd("d", num, sttDate)
        currentEndDate = DateAdd("d", num + 38, sttDate) ' 39天區間
        If currentEndDate > #12/31/2024# And currentStartDate <= #12/31/2024# Then
            If DateDiff("d", currentStartDate, #12/31/2024#) >= 0 Then
                currentEndDate = #12/31/2024#
            Else
                Exit Do
            End If
        ElseIf currentStartDate > #12/31/2024# Then
            Exit Do
        End If
        sql = "INSERT INTO [month30+8cycle] (StartDate, EndDate) " & _
              "VALUES (#" & Format(currentStartDate, "yyyy-mm-dd") & "#, #" & Format(currentEndDate, "yyyy-mm-dd") & "#);"
        m_dal.ExecuteSQL sql
        num = num + 1 ' 每天滾動
    Loop
    
    ' --- 步驟 5.2. 計算月滾動區間交易筆數與金額 ---
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 5.2. month30+8_count&sum:" & vbCrLf & _
                "FROM [deposit_credit], [month30+8cycle] INTO [month_countandsum]"
    sql = "SELECT t1.[交易對手帳戶], t6.StartDate, t6.EndDate, " & _
          "COUNT(*) AS [連續30日同一交易對手交易筆數 ], " & _
          "SUM(t1.[本幣貸方金額]) AS [連續30日同一交易對手交易金額 ] " & _
          "INTO month_countandsum " & _
          "FROM deposit_credit AS t1 INNER JOIN [month30+8cycle] AS t6 " & _
          "ON (t1.[日期] >= t6.StartDate) AND (t1.[日期] <= t6.EndDate) " & _
          "GROUP BY t1.[交易對手帳戶], t6.StartDate, t6.EndDate;"
    m_dal.ExecuteSQL sql
    
    ' --- 步驟 5.3. 標記月區間交易筆數異常分錄 ---
    ' 門檻值: 交易筆數 > 20
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 5.3. 標記筆數異常分錄(月):" & vbCrLf & _
                "FROM [deposit_credit], [month_countandsum] INTO [month_error_count]"
    sql = "SELECT DISTINCT A.* INTO month_error_count " & _
          "FROM (SELECT t1.*, t7.[連續30日同一交易對手交易筆數 ], " & _
          "IIF(t7.[連續30日同一交易對手交易筆數 ] > 20, 'V', '') AS 超過筆數門檻之異常分錄_月 " & _
          "FROM deposit_credit AS t1 LEFT JOIN month_countandsum AS t7 " & _
          "ON (t1.[交易對手帳戶] = t7.[交易對手帳戶]) AND (t1.[日期] >= t7.StartDate) AND (t1.[日期] <= t7.EndDate)) AS A " & _
          "WHERE A.超過筆數門檻之異常分錄_月='V';"
    m_dal.ExecuteSQL sql
    
    ' --- 步驟 5.4. 標記月區間交易金額異常分錄 ---
    ' 門檻值: 交易金額 > 10000000
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 5.4. 標記金額異常分錄(月):" & vbCrLf & _
                "FROM [deposit_credit], [month_countandsum] INTO [month_error_sum]"
    sql = "SELECT DISTINCT A.* INTO month_error_sum " & _
          "FROM (SELECT t1.*, t7.[連續30日同一交易對手交易金額 ], " & _
          "IIF(t7.[連續30日同一交易對手交易金額 ] > 10000000, 'V', '') AS 超過金額門檻之異常分錄_月 " & _
          "FROM deposit_credit AS t1 LEFT JOIN month_countandsum AS t7 " & _
          "ON (t1.[交易對手帳戶] = t7.[交易對手帳戶]) AND (t1.[日期] >= t7.StartDate) AND (t1.[日期] <= t7.EndDate)) AS A " & _
          "WHERE A.超過金額門檻之異常分錄_月 = 'V';"
    m_dal.ExecuteSQL sql
    
    ' Execute procedure... end
    
    MsgBox "Scenario 1 執行完畢。", vbInformation
    'success = True
    m_dal.Disconnect
    ExecuteScenario1 = success

Cleanup:
    If Not m_dal Is Nothing Then
        If m_dal.IsConnected() Then
            m_dal.Disconnect
            Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 資料庫連接已斷開。"
        End If
    End If
    ExecuteScenario1 = success
    Exit Function

ErrorHandler:
    Debug.Print "錯誤於 " & MODULE_NAME & "." & METHOD_NAME & " - SQL: " & sql
    Debug.Print "錯誤碼 " & Err.Number & ", 描述: " & Err.Description
    Resume Cleanup ' 跳至清理程序以斷開連接並退出
    
End Function

Public Function ExecuteScenario2()
    ' 單筆大額款項拆分成多筆小金額款項
    Const METHOD_NAME As String = "ExecuteScenario2"
    Dim success As Boolean
    
    If Not m_dal.Connect Then
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 錯誤: 無法連接到案例資料庫: " & m_dal.DatabasePath
        Exit Function
    End If '若無例外則連接成功
    
    ' Execute procedure... start
    
    ' Execute procedure... end
    
    MsgBox "Scenario 2 執行完畢。", vbInformation
    'success = True
    m_dal.Disconnect
    ExecuteScenario1 = success
    
End Function

Public Function ExecuteScenario3()
    ' 涉及外幣交易
    Const METHOD_NAME As String = "ExecuteScenario3"
    Dim success As Boolean
    
    If Not m_dal.Connect Then
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 錯誤: 無法連接到案例資料庫: " & m_dal.DatabasePath
        Exit Function
    End If '若無例外則連接成功
    
    ' Execute procedure... start
    
    ' Execute procedure... end
    
    MsgBox "Scenario 3 執行完畢。", vbInformation
    'success = True
    m_dal.Disconnect
    ExecuteScenario1 = success
    
End Function
    
