VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PreviewService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'Model: Preview Service
'Description: Load maximum 1000 records of selected table to worksheet, let user preview data without opening Access.

Private Const MODULE_NAME As String = "PreviewService"
Private Const adDate As Long = 7 ' ADODB Date type constant


Public Function GetAccessTableNames(ByVal accessDbPath As String) As Variant
    ' Service: Retrieves a list of table names from the specified Access database.
    Const METHOD_NAME As String = "GetAccessTableNames"
    Dim dal As AccessDAL
    Dim tableNames As Variant

    On Error GoTo ErrorHandler
    GetAccessTableNames = Empty ' Default return value

    Set dal = New AccessDAL
    dal.DatabasePath = accessDbPath
    tableNames = dal.GetTableNames ' Delegate to DAL

    GetAccessTableNames = tableNames ' Return the variant array

CleanUp:
    Set dal = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "--- " & MODULE_NAME & "." & METHOD_NAME & " Error ---"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "Error Description: " & Err.Description
    ' Let the caller (cApplication) handle the error message to the user
    ' Return Empty to indicate failure
    GetAccessTableNames = Empty
    GoTo CleanUp
End Function

Public Function ShowPreview(ByVal accessDbPath As String, ByVal tableName As String, ByVal worksheetName As String, ByVal maxRowsToShow As Long) As Boolean
    Const METHOD_NAME As String = "ShowPreview"
    Dim dal As AccessDAL
    Dim rs As Object   ' ADODB.Recordset
    Dim ws As Worksheet
    Dim sql As String
    Dim i As Long
    Dim targetSheetExists As Boolean
    Dim fieldCount As Long
    Dim recordCount As Long

    On Error GoTo ErrorHandler
    ShowPreview = False ' Default to failure

    ' --- Excel Worksheet Handling ---
    Application.ScreenUpdating = False ' Prevent screen flicker
    Application.DisplayAlerts = False
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(worksheetName)
    targetSheetExists = (Err.Number = 0)
    On Error GoTo ErrorHandler ' Restore default error handling
    Application.DisplayAlerts = True

    If targetSheetExists Then
        ws.Cells.Clear
        Debug.Print "Cleared existing worksheet: " & worksheetName
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = worksheetName
        On Error Resume Next ' Attempt to modify CodeName
        ThisWorkbook.VBProject.VBComponents(ws.CodeName).name = worksheetName
        If Err.Number <> 0 Then Debug.Print "Warning: Could not modify CodeName for worksheet '" & worksheetName & "'.": Err.Clear
        On Error GoTo ErrorHandler
        Debug.Print "Created new worksheet: " & worksheetName
    End If
    ws.Activate ' Activate the sheet for the user
    ' --- End Excel Worksheet Handling ---

    ' 2. Create DAL and get data
    Set dal = New AccessDAL
    dal.DatabasePath = accessDbPath

    ' 3. Build SQL statement
    If maxRowsToShow > 0 Then
        sql = "SELECT TOP " & maxRowsToShow & " * FROM [" & tableName & "]"
    Else
        sql = "SELECT * FROM [" & tableName & "]"
    End If

    ' 4. Get Recordset from DAL (disconnected recommended)
    Set rs = dal.GetRecordset(sql, disconnectRecordset:=True)

    ' 5. Check if Recordset is valid
    If rs Is Nothing Then
        Err.Raise vbObjectError + 517, MODULE_NAME & "." & METHOD_NAME, "Failed to retrieve data from Access table '" & tableName & "'."
    End If

    ' --- Write data to Excel ---
    If Not (rs.EOF And rs.BOF) Then ' Check if there are records
        rs.MoveFirst
        fieldCount = rs.Fields.Count
        recordCount = rs.recordCount ' Can get RecordCount from disconnected recordset

        ' 6. Write Headers
        For i = 0 To fieldCount - 1
            ws.Cells(1, i + 1).value = rs.Fields(i).name
        Next i

        ' 7. Copy Data
        If recordCount > 0 Then
            ws.Range("A2").CopyFromRecordset rs, recordCount
        End If

        ' 8. Format Dates (Optional)
        For i = 0 To fieldCount - 1
            If rs.Fields(i).Type = adDate Then
                If recordCount > 0 Then
                     ws.Range("A2").Offset(0, i).Resize(recordCount, 1).NumberFormat = "yyyy/mm/dd"
                End If
            End If
        Next i

        ' 9. AutoFit Columns
        ws.Columns.AutoFit
    Else
        ' Handle empty table: Write headers if possible
        If Not rs Is Nothing Then
             If rs.Fields.Count > 0 Then
                 For i = 0 To rs.Fields.Count - 1
                    ws.Cells(1, i + 1).value = rs.Fields(i).name
                Next i
                ws.Columns.AutoFit
             Else
                 ws.Cells(1, 1).value = "Table [" & tableName & "] exists but has no columns or could not be read."
             End If
        Else
             ws.Cells(1, 1).value = "Could not read table [" & tableName & "]."
        End If
        MsgBox "No data to load from Access table [" & tableName & "].", vbInformation
    End If
    ' --- End Write data to Excel ---

    ' 10. Success
    ShowPreview = True

CleanUp:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    Set rs = Nothing
    Set dal = Nothing ' DAL's Terminate handles disconnection
    Set ws = Nothing
    Application.ScreenUpdating = True ' Restore screen updating
    Exit Function

ErrorHandler:
    ShowPreview = False
    Debug.Print "--- " & MODULE_NAME & "." & METHOD_NAME & " Error ---"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "Error Description: " & Err.Description
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ' Raise the error to the caller (cApplication)
    Err.Raise Err.Number, MODULE_NAME & "." & METHOD_NAME, Err.Description
    GoTo CleanUp
End Function
