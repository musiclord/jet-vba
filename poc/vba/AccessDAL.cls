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
    Dim fso As Object ' Scripting.FileSystemObject - Late Binding
    Dim dbCreated As Boolean
    dbCreated = False ' Flag to indicate if DB was created in this call

    On Error GoTo ErrorHandler

    ' --- 1. Check Database Path ---
    If m_dbPath = "" Then
         Err.Raise vbObjectError + 514, MODULE_NAME & "." & METHOD_NAME, "DatabasePath is not set."
    End If

    ' --- 2. Check if Database File Exists ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(m_dbPath) Then
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - Database file not found at: " & m_dbPath & ". Attempting to create..."
        ' --- 3. Create Database if Not Exists (using ADOX - Late Binding) ---
        Dim cat As Object ' ADOX.Catalog - Late Binding
        Dim createConnStr As String

        On Error Resume Next ' Temporarily disable error handling for CreateObject
        Set cat = CreateObject("ADOX.Catalog")
        If Err.Number <> 0 Then
            Set fso = Nothing
            Err.Raise vbObjectError + 518, MODULE_NAME & "." & METHOD_NAME, "Failed to create ADOX.Catalog object. Ensure 'Microsoft ADO Ext. for DDL and Security' components are available or registered."
            GoTo CleanUp ' Exit if ADOX object creation fails
        End If
        On Error GoTo ErrorHandler ' Restore default error handling

        ' Connection string for creating the database
        createConnStr = "Provider=" & DB_PROVIDER & ";" & _
                        "Data Source=" & m_dbPath & ";"
                        ' No need for Persist Security Info for creation

        On Error Resume Next ' Handle potential errors during DB creation
        cat.Create createConnStr
        If Err.Number <> 0 Then
            Dim createErrorDesc As String
            createErrorDesc = Err.Description ' Capture error description before clearing
            Set cat = Nothing
            Set fso = Nothing
            Err.Clear ' Clear the error before raising a new one
            Err.Raise vbObjectError + 519, MODULE_NAME & "." & METHOD_NAME, "Failed to create new database file at '" & m_dbPath & "'. Error: " & createErrorDesc
            GoTo CleanUp ' Exit if creation fails
        End If
        On Error GoTo ErrorHandler ' Restore default error handling

        Set cat = Nothing ' Release ADOX object
        dbCreated = True
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - Successfully created new database file: " & m_dbPath
    End If
    Set fso = Nothing ' Release FileSystemObject

    ' --- 4. Initialize and Open Connection ---
    If m_conn Is Nothing Then
        Set m_conn = CreateObject("ADODB.Connection") ' Late binding
    End If

    If Not IsConnected() Then
        Dim connStr As String
        connStr = "Provider=" & DB_PROVIDER & ";" & _
                  "Data Source=" & m_dbPath & ";" & _
                  "Persist Security Info=False;"
        m_conn.ConnectionString = connStr
        m_conn.Open
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - Connection opened successfully to: " & m_dbPath
    Else
        Debug.Print MODULE_NAME & "." & METHOD_NAME & " - Connection was already open to: " & m_dbPath
    End If

    Connect = IsConnected()

    ' Optional: Inform user if DB was just created
    If dbCreated Then
        ' MsgBox "已成功建立新的資料庫檔案：" & vbCrLf & m_dbPath, vbInformation, "資料庫建立"
        ' Consider if a MsgBox is appropriate here, or just logging. Logging is less intrusive.
    End If

CleanUp: ' Added CleanUp label
    Set fso = Nothing ' Ensure FSO is released on exit/error
    ' Set cat = Nothing ' Ensure ADOX object is released (already done above, but safe)
    Exit Function

ErrorHandler:
    Connect = False
    Debug.Print "--- " & MODULE_NAME & "." & METHOD_NAME & " Error ---"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "Error Description: " & Err.Description
    ' Raise the error to the caller
    Err.Raise Err.Number, MODULE_NAME & "." & METHOD_NAME, Err.Description
    GoTo CleanUp ' Go to cleanup after raising error
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
    Dim dbPathFromConn As String ' For verification
    ExecuteSQL = -1 ' Default to error

    Debug.Print MODULE_NAME & ".ExecuteSQL - Received SQL: " & sql

    On Error GoTo ErrorHandler

    ' --- Enhanced Connection Check ---
    ' Check if connection object exists and is open. If not, try to connect.
    If m_conn Is Nothing Then ' <-- Use m_conn
        Debug.Print MODULE_NAME & ".ExecuteSQL - Connection object is Nothing. Attempting to connect..."
        If Not Me.Connect Then ' Assumes Connect method uses m_DatabasePath
            Err.Raise vbObjectError + 514, MODULE_NAME & "." & METHOD_NAME, "Failed to establish database connection before executing SQL. Check DatabasePath property."
            GoTo CleanUp ' Exit if connection fails
        End If
    ElseIf Not IsConnected() Then ' <-- Use IsConnected() which checks m_conn.State
        Debug.Print MODULE_NAME & ".ExecuteSQL - Connection state is not open. Attempting to reconnect..."
        If Not Me.Connect Then ' Try to reconnect
             Err.Raise vbObjectError + 515, MODULE_NAME & "." & METHOD_NAME, "Failed to re-establish database connection before executing SQL. Check DatabasePath property."
             GoTo CleanUp ' Exit if reconnection fails
        End If
    End If
    ' --- End Enhanced Connection Check ---

    ' --- Verify Connection String and Path Existence (Optional but recommended) ---
    Debug.Print MODULE_NAME & ".ExecuteSQL - Connection State before Execute: " & CStr(m_conn.State) ' <-- Use m_conn
    Debug.Print MODULE_NAME & ".ExecuteSQL - Connection String before Execute: " & m_conn.ConnectionString ' <-- Use m_conn

    ' Extract database path from the current connection string for verification
    On Error Resume Next ' In case connection string parsing fails or lacks "Data Source"
    dbPathFromConn = "" ' Reset path
    If InStr(1, m_conn.ConnectionString, "Data Source=", vbTextCompare) > 0 Then ' <-- Use m_conn
        dbPathFromConn = Mid(m_conn.ConnectionString, InStr(1, m_conn.ConnectionString, "Data Source=", vbTextCompare) + Len("Data Source=")) ' <-- Use m_conn
        If InStr(dbPathFromConn, ";") > 0 Then ' Check if there are more parts after the path
             dbPathFromConn = Split(dbPathFromConn, ";")(0)
        End If
        ' Trim potential quotes
         If Left(dbPathFromConn, 1) = """" And Right(dbPathFromConn, 1) = """" Then
             dbPathFromConn = Mid(dbPathFromConn, 2, Len(dbPathFromConn) - 2)
         End If
    End If
    On Error GoTo ErrorHandler ' Restore default error handling
    
    If dbPathFromConn = "" Then
         Debug.Print MODULE_NAME & ".ExecuteSQL - WARNING: Could not extract database path from connection string: " & m_conn.ConnectionString ' <-- Use m_conn
    Else
        Debug.Print MODULE_NAME & ".ExecuteSQL - Extracted DB Path from Connection: '" & dbPathFromConn & "'"
        ' Use FileSystemObject for a more robust check
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FileExists(dbPathFromConn) Then
            Debug.Print MODULE_NAME & ".ExecuteSQL - ERROR: Database file NOT FOUND at path from connection string: " & dbPathFromConn
            Set fso = Nothing ' Release before raising error
            Err.Raise vbObjectError + 516, MODULE_NAME & "." & METHOD_NAME, "Database file specified in connection string not found: " & dbPathFromConn
            GoTo CleanUp ' Should not be reached if Err.Raise works
        Else
            Debug.Print MODULE_NAME & ".ExecuteSQL - Database file confirmed FOUND at path from connection string."
        End If
        Set fso = Nothing
    End If
    ' --- End Verification ---
    
    Debug.Print MODULE_NAME & ".ExecuteSQL - Executing SQL on connection..."
    If useTransaction Then m_conn.BeginTrans ' <-- Use m_conn
    
    m_conn.Execute sql, recordsAffected, 128 ' adExecuteNoRecords = 128 <-- Use m_conn
    
    If useTransaction Then m_conn.CommitTrans ' <-- Use m_conn
    
    ExecuteSQL = recordsAffected ' Success

CleanUp:
    ' Optional: Decide if connection should be closed here or managed elsewhere
    Exit Function

ErrorHandler:
    If useTransaction And Not m_conn Is Nothing Then ' <-- Use m_conn
        If IsConnected() Then ' <-- Use IsConnected() which checks m_conn
            On Error Resume Next ' Avoid error during rollback if connection lost
            m_conn.RollbackTrans ' <-- Use m_conn
            On Error GoTo 0
        End If
    End If
    ExecuteSQL = -1 ' Indicate error
    Dim errorDesc As String
    errorDesc = "Error executing SQL: " & sql & " - " & Err.Description
    Debug.Print "--- " & MODULE_NAME & "." & METHOD_NAME & " Error ---"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Source: " & Err.Source & " (Original Error: " & Err.Description & ")"
    ' Raise a more informative error, including the original description
    Err.Raise Err.Number, MODULE_NAME & "." & METHOD_NAME, errorDesc
    GoTo CleanUp ' Should not be reached if Err.Raise works
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
    Dim tblList As Collection ' Changed type hint for clarity
    Dim tblName As String
    Dim result() As String
    Dim i As Long
    On Error GoTo ErrorHandler
    
    If Not Connect() Then
        Err.Raise vbObjectError + 515, MODULE_NAME & "." & METHOD_NAME, "Failed to connect to database."
    End If
    
    Set tblList = New Collection ' Use VBA's built-in Collection
    
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

