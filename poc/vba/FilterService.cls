VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "FilterService"
Private m_AccessDAL As AccessDAL
Private m_MappingService As MappingService
Private m_GLTable As String ' Access 中 GL 資料的實際資料表名稱
Private m_TBTable As String ' Access 中 TB 資料的實際資料表名稱

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

Public Function ApplyFilter(ByVal criteriaCollection As Collection, ByVal targetTableName As String) As Object ' ADODB.Recordset
    Const METHOD_NAME As String = "ApplyFilter"
    Dim sql As String
    Dim whereClause As String
    Dim criterion As Object ' Scripting.Dictionary
    Dim fieldName As String
    Dim operator As String
    Dim value As String
    Dim formattedValue As String
    Dim firstCondition As Boolean

    On Error GoTo ErrorHandler
    Set ApplyFilter = Nothing
    firstCondition = True
    whereClause = ""
    ' 確保資料庫已連接
    If Not m_AccessDAL.IsConnected() Then
        If Not m_AccessDAL.Connect() Then
             Err.Raise vbObjectError + 502, MODULE_NAME & "." & METHOD_NAME, "無法連接到資料庫。"
            Exit Function
        End If
    End If
    If criteriaCollection Is Nothing Or criteriaCollection.Count = 0 Then
        ' 如果沒有提供篩選條件，則選取所有記錄
        sql = "SELECT * FROM [" & targetTableName & "];"
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 無篩選條件，執行的 SQL: " & sql
        Set ApplyFilter = m_AccessDAL.GetRecordset(sql)
        Exit Function
    End If

    For Each criterion In criteriaCollection
        fieldName = Trim$(CStr(criterion("Field")))
        operator = Trim$(CStr(criterion("Operator")))
        value = Trim$(CStr(criterion("Value")))

        ' 忽略無效的條件部分
        If fieldName = "" Or operator = "" Or value = "" Then GoTo NextCriterion

        ' 處理值的格式 (數字 vs 字串)
        If IsNumeric(value) Then
            formattedValue = value
        Else
            ' 對字串值中的單引號進行轉義，並用單引號包圍
            formattedValue = "'" & Replace(value, "'", "''") & "'"
        End If

        If Not firstCondition Then
            whereClause = whereClause & " AND "
        Else
            firstCondition = False
        End If

        ' 根據運算子建構條件
        Select Case operator
            Case "==" ' SQL 中的等於是 =
                whereClause = whereClause & "[" & fieldName & "] = " & formattedValue
            Case "HAS" ' SQL 中的包含是 LIKE
                ' 對於 HAS，通常用於文字欄位。如果值本身包含 SQL 的萬用字元，可能需要進一步處理。
                ' 此處假設 value 是要搜尋的純文字。
                If IsNumeric(value) Then ' 如果原值是數字，但要用 HAS，則將欄位轉為文字比較
                    whereClause = whereClause & "CStr([" & fieldName & "]) LIKE '%" & Replace(value, "'", "''") & "%'"
                Else
                    whereClause = whereClause & "[" & fieldName & "] LIKE '%" & Replace(value, "'", "''") & "%'"
                End If
            Case ">="
                whereClause = whereClause & "[" & fieldName & "] >= " & formattedValue
            Case "<="
                whereClause = whereClause & "[" & fieldName & "] <= " & formattedValue
            Case Else
                Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 警告: 未知的運算子 '" & operator & "'，此條件將被忽略。"
                If whereClause Like "* AND " Then ' 移除多餘的 AND
                    whereClause = Left$(whereClause, Len(whereClause) - 5)
                End If
                firstCondition = (whereClause = "") ' 如果移除了唯一的條件，重設 firstCondition
                GoTo NextCriterion ' 忽略此條件
        End Select
NextCriterion:
    Next criterion

    If Trim$(whereClause) <> "" Then
        sql = "SELECT * FROM [" & targetTableName & "] WHERE " & whereClause & ";"
    Else
        ' 如果所有條件都無效或被忽略，則選取所有記錄
        sql = "SELECT * FROM [" & targetTableName & "];"
    End If

    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 執行的 SQL: " & sql
    Set ApplyFilter = m_AccessDAL.GetRecordset(sql)

Exit Function
ErrorHandler:
    Debug.Print "錯誤於 " & MODULE_NAME & "." & METHOD_NAME & ": " & Err.Description
    Set ApplyFilter = Nothing ' 確保錯誤時回傳 Nothing
End Function

Public Function CreateTableFromFilter(ByVal criteriaCollection As Collection, ByVal sourceTableName As String, ByVal destinationTableName As String) As Boolean
    Const METHOD_NAME As String = "CreateTableFromFilter"
    Dim sql As String
    Dim whereClause As String
    Dim criterion As Object ' Scripting.Dictionary
    Dim fieldName As String
    Dim operator As String
    Dim value As String
    Dim formattedValue As String
    Dim firstCondition As Boolean
    Dim recordsAffected As Long

    On Error GoTo ErrorHandler
    CreateTableFromFilter = False ' 預設失敗
    firstCondition = True
    whereClause = ""
    recordsAffected = -1

    If m_AccessDAL Is Nothing Then
        Err.Raise vbObjectError + 501, MODULE_NAME & "." & METHOD_NAME, "AccessDAL 未初始化。"
        Exit Function
    End If

    ' 確保資料庫已連接
    If Not m_AccessDAL.IsConnected() Then
        If Not m_AccessDAL.Connect() Then
             Err.Raise vbObjectError + 502, MODULE_NAME & "." & METHOD_NAME, "無法連接到資料庫。"
            Exit Function
        End If
    End If

    ' --- 1. 建構 WHERE 子句 (與 ApplyFilter 邏輯類似) ---
    If Not (criteriaCollection Is Nothing Or criteriaCollection.Count = 0) Then
        For Each criterion In criteriaCollection
            fieldName = Trim$(CStr(criterion("Field")))
            operator = Trim$(CStr(criterion("Operator")))
            value = Trim$(CStr(criterion("Value")))

            If fieldName = "" Or operator = "" Or value = "" Then GoTo NextCriterion

            ' === 修改這裡的值格式處理邏輯 ===
            ' 特殊處理科目代碼欄位：強制使用字串格式，不論值是否看起來像數字
            If fieldName = "科目代碼" Or fieldName = "傳票號碼" Or _
               fieldName = "帳戶" Or fieldName = "統一編號" Or _
               fieldName = "文件號碼" Then
            ' 這些欄位一律視為字串處理
                formattedValue = "'" & Replace(value, "'", "''") & "'"
            ' 如果有其他需要特殊處理的欄位，在這裡添加 ElseIf 條件
            Else
                ' 一般欄位處理邏輯
                If IsNumeric(value) Then
                    formattedValue = value
                Else
                    formattedValue = "'" & Replace(value, "'", "''") & "'"
                End If
            End If

            If Not firstCondition Then
                whereClause = whereClause & " AND "
            Else
                firstCondition = False
            End If

            Select Case operator
                Case "=="
                    whereClause = whereClause & "[" & fieldName & "] = " & formattedValue
                Case ">="
                    whereClause = whereClause & "[" & fieldName & "] >= " & formattedValue
                Case "<="
                    whereClause = whereClause & "[" & fieldName & "] <= " & formattedValue
                Case Else
                    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 警告: 未知的運算子 '" & operator & "'，此條件將被忽略。"
                    If Right$(whereClause, 5) = " AND " Then
                        whereClause = Left$(whereClause, Len(whereClause) - 5)
                    End If
                    firstCondition = (Trim$(whereClause) = "")
                    GoTo NextCriterion
            End Select
NextCriterion:
        Next criterion
    End If

    ' --- 2. 嘗試刪除已存在的目標資料表 ---
    ' DropTable 內部有自己的錯誤處理，如果資料表不存在，它不會引發錯誤使此處中斷
    m_AccessDAL.DropTable destinationTableName
    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 嘗試刪除舊的資料表: " & destinationTableName

    ' --- 3. 建構 SELECT INTO SQL ---
    sql = "SELECT * INTO [" & destinationTableName & "] FROM [" & sourceTableName & "]"

    If Trim$(whereClause) <> "" Then
        sql = sql & " WHERE " & whereClause
    End If
    sql = sql & ";"

    Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 執行的 SQL: " & sql

    ' --- 4. 執行 SQL ---
    recordsAffected = m_AccessDAL.ExecuteSQL(sql)

    If recordsAffected <> -1 Then
        CreateTableFromFilter = True ' SQL 執行成功 (對於 SELECT INTO，影響的記錄數可能為 0 或更多)
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 資料表 '" & destinationTableName & "' 建立成功，影響記錄數: " & recordsAffected
    Else
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - 建立資料表 '" & destinationTableName & "' 失敗。"
    End If

    Exit Function
ErrorHandler:
    Debug.Print "錯誤於 " & MODULE_NAME & "." & METHOD_NAME & ": " & Err.Description
    CreateTableFromFilter = False
End Function
