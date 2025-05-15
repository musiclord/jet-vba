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
