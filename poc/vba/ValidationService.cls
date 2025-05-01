VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' Function:  TestCompleteness (完善版)
' Purpose:   執行 GL 與 TB 的完整性對平測試。
'            依賴 AccessDAL.GetScalar 方法執行 SQL 並獲取差異計數。
' Returns:   Boolean - True 表示測試通過 (無顯著差異)，False 表示失敗或出錯。
' Note:      已修正 Difference 計算順序為 TB - GL。
'            包含對 GetScalar 返回值的健壯性檢查和改進的錯誤處理。
'-------------------------------------------------------------------------------
Public Function TestCompleteness() As Boolean
    Const METHOD_NAME As String = "TestCompleteness" ' 方法名稱，方便除錯
    Dim sql As String
    Dim differenceCount As Long
    Dim result As Variant ' **重要：使用 Variant 接收 GetScalar 結果**
    Const TOLERANCE As Double = 0.01 ' 定義容錯值

    On Error GoTo ErrorHandler ' **啟用結構化錯誤處理**
    TestCompleteness = False ' 預設為失敗

    ' --- 1. 確保 DAL 實例存在 ---
    ' 假設 dal 在 Class_Initialize 中已成功 Set
    If dal Is Nothing Then
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 錯誤：DAL 物件尚未初始化。"
        Exit Function ' 或者觸發錯誤 Err.Raise ...
    End If

    ' --- 2. 構建最終確認的 SQL 查詢 ---
    ' 計算差異記錄數，Difference 計算已調整為 TB - GL
    sql = "SELECT Count(*) AS DifferenceCount " & _
          "FROM ( " & _
              "SELECT " & _
                  "(Nz(TB.Change, 0) - Nz(QGL.GL_Movement, 0)) AS Difference " & _
              "FROM ( " & _
                  "SELECT GL.AccountNo, Sum(GL.Amount) AS GL_Movement " & _
                  "FROM GL GROUP BY GL.AccountNo " & _
              ") AS QGL " & _
              "LEFT JOIN TB ON QGL.AccountNo = TB.AccountNo " & _
          ") AS Comp " & _
          "WHERE Abs(Comp.Difference) > " & CStr(TOLERANCE) & ";" ' 使用 CStr

    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 執行 SQL: " & sql

    ' --- 3. 調用 DAL 的 GetScalar 方法 ---
    result = dal.GetScalar(sql) ' 接收為 Variant

    ' --- 4. **健壯地處理 GetScalar 返回值** ---
    If IsError(result) Then ' 檢查 GetScalar 是否返回了 VBA 錯誤值
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - GetScalar 方法返回錯誤值。"
        ' 錯誤已在 DAL 層記錄，此處不再重複 MsgBox，直接返回 False
        GoTo CleanUp
    ElseIf IsNull(result) Then ' 檢查是否返回 Null (Count(*) 正常不應返回 Null)
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - GetScalar 方法返回 Null。"
        GoTo CleanUp ' 視為異常，返回 False
    ElseIf Not IsNumeric(result) Then ' 檢查是否為數值
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - GetScalar 方法返回非數值: '" & CStr(result) & "'"
        GoTo CleanUp ' 視為異常，返回 False
    Else
        ' --- 5. 安全轉換並判斷結果 ---
        differenceCount = CLng(result) ' 確認是數值後再轉換
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 發現差異記錄數量: " & differenceCount

        If differenceCount = 0 Then
            TestCompleteness = True ' 沒有差異，測試通過
        Else
            TestCompleteness = False ' 存在差異，測試失敗
        End If
    End If

CleanUp: ' 清理出口點
    Exit Function ' 正常或經過錯誤處理後退出

ErrorHandler: ' 錯誤處理標籤
    TestCompleteness = False ' 確保錯誤時函數返回 False
    ' 使用 Debug.Print 記錄詳細錯誤信息
    Debug.Print "--- " & MODULE_NAME & "." & METHOD_NAME & " 錯誤 ---"
    Debug.Print "SQL (若適用): " & sql ' 記錄當時執行的 SQL
    Debug.Print "錯誤描述: " & Err.Description & " (" & Err.Number & ")"
    ' 根據您的錯誤處理策略，可選擇性地記錄到日誌文件或顯示通用錯誤信息給使用者
    ' MsgBox "執行完整性測試時發生未預期的錯誤。", vbCritical, "內部錯誤"
    GoTo CleanUp ' 跳轉到清理步驟
End Function

