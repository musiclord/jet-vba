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

' --- 假設的初始化 ---
Private Sub Class_Initialize()
    ' 在實際應用中，DAL 和 MappingService 需要被正確實例化和注入
    ' Set dal = New AccessDAL
    ' dal.DatabasePath = ... ' 設定路徑
    ' Set MappingService = ... ' 獲取 MappingService 實例
    Debug.Print MODULE_NAME & " Initialized (DAL/MappingService should be set)."
End Sub

'-------------------------------------------------------------------------------
' Function:  TestCompleteness (當前版本 - 僅計算差異數量)
' Purpose:   執行 GL 與 TB 的基礎完整性對平測試。
'            此版本旨在計算 GL 彙總變動與 TB 記錄變動之間存在顯著差異的科目數量。
'            其核心邏輯與先前討論的 IDEA 腳本 (#file:caseware_idea_vba.bas)
'            中的步驟相似，但目前**簡化**了實現：
'            1. 彙總 GL 變動 (對應 IDEA Summarization)。
'            2. 與 TB 資料連接 (對應 IDEA Join，但目前使用 LEFT JOIN)。
'            3. 計算差異 (對應 IDEA Append Field 'DIFF')。
'            4. 篩選並計數差異 (對應 IDEA Extraction with Criteria)。
'            **注意：此版本目前使用硬編碼的欄位名稱和表名，且未實現日期篩選和差異明細輸出。**
' Arguments: 無 (未來應加入 StartDate, EndDate)。
' Returns:   Boolean - True 表示測試通過 (差異數量為 0)，False 表示失敗 (差異數量 > 0) 或執行出錯。
'-------------------------------------------------------------------------------
Public Function TestCompleteness() As Boolean
    Const METHOD_NAME As String = "TestCompleteness" ' 方法名稱，方便除錯
    Dim sql As String
    Dim differenceCount As Long
    Dim result As Variant ' **重要：使用 Variant 接收 GetScalar 結果**
    Const TOLERANCE As Double = 0.01 ' 定義可接受的差異容錯值

    On Error GoTo ErrorHandler ' **啟用結構化錯誤處理**
    TestCompleteness = False ' 預設為失敗

    ' --- 1. 確保 DAL 實例存在 ---
    ' 檢查資料存取層物件是否已準備就緒。
    If dal Is Nothing Then
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 錯誤：DAL 物件尚未初始化。"
        Exit Function ' 或者觸發錯誤 Err.Raise ...
    End If

    ' --- 2. 構建 SQL 查詢 ---
    ' 核心邏輯：計算 GL 變動與 TB 變動之間差異大於容錯值的科目數量。
    ' # AI Guidance: 目前 SQL 使用硬編碼欄位 (GL.AccountNo, GL.Amount, TB.AccountNo, TB.Change)。
    ' # AI Guidance: 未來應修改此 SQL，使其動態使用從 MappingService 獲取的欄位名稱。
    ' # AI Guidance: 例如，將 "GL.AccountNo" 替換為 "[" & glMappings("AccountNo") & "]"。
    ' # AI Guidance: 同時，GL.Amount 的計算邏輯 (Sum(GL.Amount)) 需要根據映射的金額欄位 (單一金額+借貸標記 或 借/貸分開) 動態生成，
    ' # AI Guidance: 類似於之前討論的 glAmtExpr 變數。
    ' # AI Guidance: TB.Change 也應動態計算 (Nz(EndBalance, 0) - Nz(BeginBalance, 0))。
    ' # AI Guidance: 未加入日期篩選，目前是對 GL 表所有記錄進行彙總。未來應加入 WHERE GL.PostDate BETWEEN StartDate AND EndDate。
    sql = "SELECT Count(*) AS DifferenceCount " & _
          "FROM ( " & _
              "SELECT " & _
                  ' 計算差異：TB 變動 - GL 變動。Nz() 用於處理某科目只存在於一個表的情況 (因 LEFT JOIN)。
                  "(Nz(TB.Change, 0) - Nz(QGL.GL_Movement, 0)) AS Difference " & _
              "FROM ( " & _
                  ' 子查詢 QGL：彙總 GL 表，按科目編號加總金額，計算 GL 方的淨變動。
                  ' 對應 IDEA 步驟：篩選期間 (目前缺少) -> 計算帶號金額 (目前假設 Amount 已帶號) -> Summarization。
                  "SELECT GL.AccountNo, Sum(GL.Amount) AS GL_Movement " & _
                  "FROM GL GROUP BY GL.AccountNo " & _
              ") AS QGL " & _
              ' 將 GL 彙總結果 (QGL) 左連接到 TB 表。
              ' # AI Guidance: 為完全模擬 IDEA 的 "All records in both files"，應改用更複雜的 FULL OUTER JOIN 模擬 (如之前討論的 UNION ALL 方法)。
              ' # AI Guidance: 目前 LEFT JOIN 可能會遺漏只存在於 TB 而不存在於 GL 的科目差異 (雖然理論上不應發生)。
              "LEFT JOIN TB ON QGL.AccountNo = TB.AccountNo " & _
          ") AS Comp " & _
          ' 篩選條件：只計算差異絕對值大於容錯值的記錄數。
          "WHERE Abs(Comp.Difference) > " & CStr(TOLERANCE) & ";" ' 使用 CStr 確保數值轉字串正確

    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 執行 SQL: " & sql

    ' --- 3. 調用 DAL 的 GetScalar 方法 ---
    ' 執行 SQL 查詢並獲取單一結果 (差異計數)。
    result = dal.GetScalar(sql) ' 接收為 Variant 以處理可能的錯誤或非預期類型

    ' --- 4. **健壯地處理 GetScalar 返回值** ---
    ' 檢查 GetScalar 的返回結果是否有效。
    If IsError(result) Then ' 檢查是否返回了 VBA 錯誤值 (例如 CVErr)
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - GetScalar 方法返回錯誤值。"
        ' 假設錯誤已在 DAL 層記錄，此處不再重複 MsgBox，直接返回 False
        GoTo CleanUp
    ElseIf IsNull(result) Then ' Count(*) 正常不應返回 Null，但以防萬一
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - GetScalar 方法返回 Null。"
        GoTo CleanUp ' 視為異常，返回 False
    ElseIf Not IsNumeric(result) Then ' 確保返回的是數值
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - GetScalar 方法返回非數值: '" & CStr(result) & "'"
        GoTo CleanUp ' 視為異常，返回 False
    Else
        ' --- 5. 安全轉換並判斷結果 ---
        ' 確認是有效數值後，轉換為 Long 並判斷差異。
        differenceCount = CLng(result)
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 發現差異記錄數量: " & differenceCount

        If differenceCount = 0 Then
            TestCompleteness = True ' 沒有顯著差異，測試通過
            Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 測試通過：未發現顯著差異。"
        Else
            TestCompleteness = False ' 存在顯著差異，測試失敗
            Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 測試失敗：發現 " & differenceCount & " 筆顯著差異。"
            ' # AI Guidance: 測試失敗時，應觸發生成差異明細表 (COMPLETENESS) 的邏輯。
            ' # AI Guidance: 可以呼叫另一個私有方法，該方法執行之前討論的 "SELECT INTO COMPLETENESS" SQL。
        End If
    End If

CleanUp: ' 清理出口點
    Exit Function ' 正常或經過錯誤處理後退出

ErrorHandler: ' 錯誤處理標籤
    TestCompleteness = False ' 確保錯誤時函數返回 False
    ' 記錄詳細錯誤信息，方便追蹤問題。
    Debug.Print "--- " & MODULE_NAME & "." & METHOD_NAME & " 錯誤 ---"
    Debug.Print "SQL (若適用): " & sql ' 記錄當時執行的 SQL
    Debug.Print "錯誤描述: " & Err.Description & " (" & Err.Number & ")"
    ' MsgBox "執行完整性測試時發生未預期的錯誤。", vbCritical, "內部錯誤" ' 可選：顯示通用錯誤給使用者
    GoTo CleanUp ' 跳轉到清理步驟
End Function

' --- 其他 ValidationService 方法 ---

' # AI Guidance: 未來可添加一個私有方法來生成差異明細表
' Private Sub CreateCompletenessTable(ByVal StartDate As Date, ByVal EndDate As Date)
'     ' 實作之前討論的 "SELECT INTO COMPLETENESS" 邏輯
'     ' 需要動態構建 SQL，包含所有映射欄位、日期篩選、FULL JOIN 模擬等
'     Const METHOD_NAME As String = "CreateCompletenessTable"
'     Dim sqlDrop As String, sqlMakeTable As String
'     ' ... 獲取映射 ...
'     ' ... 構建 SQL ...
'     sqlDrop = "DROP TABLE COMPLETENESS;"
'     sqlMakeTable = "SELECT ... INTO COMPLETENESS FROM ... WHERE Abs(Difference) > " & CStr(TOLERANCE) & ";"
'     On Error Resume Next
'     dal.ExecuteSQL sqlDrop
'     On Error GoTo ErrorHandler_CreateTable
'     If dal.ExecuteSQL(sqlMakeTable) Then
'         Debug.Print MODULE_NAME & "." & METHOD_NAME & " - COMPLETENESS 表已成功創建/更新。"
'     Else
'         Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 創建 COMPLETENESS 表失敗。"
'     End If
'     Exit Sub
' ErrorHandler_CreateTable:
'     Debug.Print "--- ERROR in " & MODULE_NAME & "." & METHOD_NAME & " ---"
'     ' ... 記錄錯誤 ...
' End Sub

