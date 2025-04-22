VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccessDAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "AccessDAL"
Private Const DB_PROVIDER As String = "Microsoft.ACE.OLEDB.12.0" ' Or Microsoft.Jet.OLEDB.4.0 for .mdb

Private m_conn As Object ' ADODB.Connection - Late binding
Private m_dbPath As String

' --- Public Properties ---

Public Property Get DatabasePath() As String
    DatabasePath = m_dbPath
End Property

Public Property Let DatabasePath(ByVal value As String)
    If m_conn Is Nothing Then
        m_dbPath = value
    ElseIf IsConnected() Then ' Check connection state safely
        Err.Raise vbObjectError + 513, MODULE_NAME & ".DatabasePath", "Cannot change DatabasePath while connection is open."
    Else
        m_dbPath = value
    End If
End Property

' --- Connection Management ---

Public Function Connect() As Boolean
    Const METHOD_NAME As String = "Connect"
    On Error GoTo ErrorHandler

    If m_conn Is Nothing Then
        Set m_conn = CreateObject("ADODB.Connection") ' Late binding
    End If

    If Not IsConnected() Then
        If m_dbPath = "" Then
             Err.Raise vbObjectError + 514, MODULE_NAME & "." & METHOD_NAME, "DatabasePath is not set."
        End If
        Dim connStr As String
        connStr = "Provider=" & DB_PROVIDER & ";" & _
                  "Data Source=" & m_dbPath & ";" & _
                  "Persist Security Info=False;"
        m_conn.ConnectionString = connStr
        m_conn.Open
    End If
    Connect = IsConnected()
    Exit Function

ErrorHandler:
    Connect = False
    ' Log error details if needed
    Err.Raise Err.Number, MODULE_NAME & "." & METHOD_NAME, Err.Description
End Function

Public Sub Disconnect()
    On Error Resume Next ' Ignore errors if already closed or not initialized
    If Not m_conn Is Nothing Then
        If IsConnected() Then
            m_conn.Close
        End If
        Set m_conn = Nothing
    End If
    On Error GoTo 0
End Sub

Public Function IsConnected() As Boolean
    On Error Resume Next
    If m_conn Is Nothing Then
        IsConnected = False
    Else
        IsConnected = (m_conn.State = 1) ' adStateOpen = 1
    End If
    On Error GoTo 0
End Function


' --- Data Manipulation (DML) & Data Definition (DDL) ---

Public Function ExecuteSQL(ByVal sql As String, Optional ByVal useTransaction As Boolean = False) As Long
    Const METHOD_NAME As String = "ExecuteSQL"
    Dim recordsAffected As Long
    On Error GoTo ErrorHandler

    If Not Connect() Then
        Err.Raise vbObjectError + 515, MODULE_NAME & "." & METHOD_NAME, "Failed to connect to database."
    End If

    If useTransaction Then m_conn.BeginTrans

    m_conn.Execute sql, recordsAffected, 128 ' adExecuteNoRecords = 128

    If useTransaction Then m_conn.CommitTrans

    ExecuteSQL = recordsAffected
    Exit Function

ErrorHandler:
    If useTransaction And Not m_conn Is Nothing Then
        If IsConnected() Then
            On Error Resume Next ' Avoid error during rollback if connection lost
            m_conn.RollbackTrans
            On Error GoTo 0
        End If
    End If
    ExecuteSQL = -1 ' Indicate error
    ' Log error details if needed
    Err.Raise Err.Number, MODULE_NAME & "." & METHOD_NAME, "Error executing SQL: " & sql & " - " & Err.Description
End Function

' --- Data Query (DQL) ---

Public Function GetRecordset(ByVal sql As String, Optional ByVal disconnectRecordset As Boolean = False) As Object ' ADODB.Recordset - Late binding
    Const METHOD_NAME As String = "GetRecordset"
    Dim rs As Object ' ADODB.Recordset
    On Error GoTo ErrorHandler

    If Not Connect() Then
        Err.Raise vbObjectError + 515, MODULE_NAME & "." & METHOD_NAME, "Failed to connect to database."
    End If

    Set rs = CreateObject("ADODB.Recordset") ' Late binding
    rs.CursorLocation = 3 ' adUseClient = 3
    rs.Open sql, m_conn, 3, 1 ' adOpenStatic = 3, adLockReadOnly = 1

    If disconnectRecordset Then
        Set rs.ActiveConnection = Nothing ' Disconnect the recordset
        ' Note: Connection can be closed here if ONLY disconnected recordsets are needed by the caller
        ' Disconnect ' Uncomment if appropriate
    End If

    Set GetRecordset = rs
    Set rs = Nothing ' Release local variable
    Exit Function

ErrorHandler:
    Set GetRecordset = Nothing
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close ' adStateOpen = 1
        Set rs = Nothing
    End If
    ' Log error details if needed
    Err.Raise Err.Number, MODULE_NAME & "." & METHOD_NAME, "Error getting Recordset for SQL: " & sql & " - " & Err.Description
End Function

' --- Schema Operations ---

Public Function TableExists(ByVal tableName As String) As Boolean
    Const METHOD_NAME As String = "TableExists"
    Dim rs As Object ' ADODB.Recordset
    On Error GoTo ErrorHandler

    If Not Connect() Then
         Err.Raise vbObjectError + 515, MODULE_NAME & "." & METHOD_NAME, "Failed to connect to database."
    End If

    ' adSchemaTables = 20
    Set rs = m_conn.OpenSchema(20, Array(Empty, Empty, tableName, "TABLE"))
    TableExists = Not rs.EOF
    rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    TableExists = False ' Assume false on error
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close ' adStateOpen = 1
        Set rs = Nothing
    End If
    ' Log error details if needed, but might just return False
    ' Err.Raise Err.Number, MODULE_NAME & "." & METHOD_NAME, Err.Description
End Function

Public Sub DropTable(ByVal tableName As String)
    Const METHOD_NAME As String = "DropTable"
    Dim sql As String
    On Error GoTo ErrorHandler

    If TableExists(tableName) Then
        sql = "DROP TABLE [" & tableName & "]"
        ExecuteSQL sql
    End If
    Exit Sub

ErrorHandler:
    ' Log error details if needed
    Err.Raise Err.Number, MODULE_NAME & "." & METHOD_NAME, "Error dropping table: " & tableName & " - " & Err.Description
End Sub

Public Function GetTableNames() As Variant ' Returns array of table names or Empty if none/error
    Const METHOD_NAME As String = "GetTableNames"
    Dim rs As Object ' ADODB.Recordset
    Dim tblList As Object ' Collection
    Dim tblName As String
    Dim result() As String
    Dim i As Long
    On Error GoTo ErrorHandler

    If Not Connect() Then
        Err.Raise vbObjectError + 515, MODULE_NAME & "." & METHOD_NAME, "Failed to connect to database."
    End If

    Set tblList = CreateObject("Scripting.Collection") ' Use Collection for simplicity

    ' Get user tables (excluding system and temp tables)
    ' adSchemaTables = 20
    Set rs = m_conn.OpenSchema(20, Array(Empty, Empty, Empty, "TABLE"))
    Do While Not rs.EOF
        tblName = rs.Fields("TABLE_NAME").value ' Use Fields collection by name
        ' Filter out system tables (MSys*) and temporary tables (~*)
        If Not (Left$(tblName, 4) = "MSys" Or Left$(tblName, 1) = "~") Then
            tblList.Add tblName
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    If tblList.Count > 0 Then
        ReDim result(1 To tblList.Count)
        For i = 1 To tblList.Count
            result(i) = tblList(i)
        Next i
        GetTableNames = result
    Else
        GetTableNames = Empty ' Return Empty if no user tables found
    End If

    Set tblList = Nothing
    Exit Function

ErrorHandler:
    GetTableNames = Empty ' Return Empty on error
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close ' adStateOpen = 1
        Set rs = Nothing
    End If
    Set tblList = Nothing
    ' Log error details if needed
    Err.Raise Err.Number, MODULE_NAME & "." & METHOD_NAME, Err.Description
End Function

Public Function GetTableFieldNames(ByVal tableName As String) As Variant ' Returns array or Empty
    Const METHOD_NAME As String = "GetTableFieldNames"
    Dim rs As Object ' ADODB.Recordset
    Dim fieldList As Object ' Collection
    Dim result() As String
    Dim i As Long
    On Error GoTo ErrorHandler

    If Not Connect() Then
        Err.Raise vbObjectError + 515, MODULE_NAME & "." & METHOD_NAME, "Failed to connect to database."
    End If

    Set fieldList = CreateObject("Scripting.Collection")

    ' Open schema for columns of the specific table
    ' adSchemaColumns = 4
    Set rs = m_conn.OpenSchema(4, Array(Empty, Empty, tableName, Empty))

    If rs.EOF Then ' Table might not exist or has no columns
        rs.Close
        Set rs = Nothing
        GetTableFieldNames = Empty
        Set fieldList = Nothing
        Exit Function
    End If

    Do While Not rs.EOF
        fieldList.Add rs.Fields("COLUMN_NAME").value ' Use Fields collection by name
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    If fieldList.Count > 0 Then
        ReDim result(1 To fieldList.Count)
        For i = 1 To fieldList.Count
            result(i) = fieldList(i)
        Next i
        GetTableFieldNames = result
    Else
        GetTableFieldNames = Empty
    End If

    Set fieldList = Nothing
    Exit Function

ErrorHandler:
    GetTableFieldNames = Empty ' Return Empty on error
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close ' adStateOpen = 1
        Set rs = Nothing
    End If
    Set fieldList = Nothing
    ' Log error details if needed
    Err.Raise Err.Number, MODULE_NAME & "." & METHOD_NAME, "Error getting field names for table: " & tableName & " - " & Err.Description
End Function


' --- Initialization / Termination ---

Private Sub Class_Initialize()
    ' Set default database path, or leave empty to force setting it
    On Error Resume Next ' Avoid error if ThisWorkbook is not available (e.g., during compilation)
    m_dbPath = ThisWorkbook.Path & "\default.accdb"
    If Err.Number <> 0 Then m_dbPath = "" ' Reset if error occurred
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Disconnect ' Ensure connection is closed when object is destroyed
End Sub

