VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImportService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
 'Model: Import Service
'Description: Handle any data import logics and uses AccessDAL for database operation.

Private Const adStateOpen As Long = 1
Private Const adDate As Long = 7
Private Const adOpenForwardOnly As Long = 0
Private Const adLockReadOnly As Long = 1
Private Const adOpenStatic As Long = 3
Private Const adLockOptimistic As Long = 3
Private Const adTypeBinary As Long = 1

Private Sub Class_Initialize()

End Sub

Public Function ImportToAccess(ByVal csvFilePath As String, ByVal accessDbPath As String, ByVal tableName As String, ByVal csvEncoding As Long) As Boolean
    ' 將 CSV 資料匯入 Access 資料庫
    ' 返回 True 表示成功，False 表示失敗
    Dim csvConn As Object  ' ADODB.Connection for CSV
    Dim accConn As Object  ' ADODB.Connection for Access
    Dim rs As Object       ' ADODB.Recordset for CSV data
    Dim destRs As Object   ' ADODB.Recordset for Access table
    Dim fso As Object      ' FileSystemObject
    Dim cat As Object      ' ADOX.Catalog
    Dim fldDef As String
    Dim i As Integer
    Dim fieldName As String

    On Error GoTo ErrorHandler
    ImportToAccess = False ' 預設為失敗

    ' 1. 確保 Access 資料庫存在
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(accessDbPath) Then
        ' 嘗試建立 Access 檔案
        On Error Resume Next ' 暫時忽略建立過程中的錯誤
        Set cat = CreateObject("ADOX.Catalog")
        If Err.Number = 0 Then
            cat.Create "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessDbPath
            Set cat = Nothing
            Debug.Print "已建立新的 Access 資料庫: " & accessDbPath
        Else
            ' 如果 ADOX 失敗，嘗試用 ADODB 連接 (可能已存在但無法被 ADOX 識別)
            Err.Clear
            Set accConn = CreateObject("ADODB.Connection")
            accConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessDbPath & ";"
            accConn.Close
            Set accConn = Nothing
            Debug.Print "Access 資料庫已存在 (或透過 ADODB 連接建立): " & accessDbPath
        End If
        On Error GoTo ErrorHandler ' 恢復錯誤處理
        If Err.Number <> 0 Then ' 如果兩種方法都失敗
             MsgBox "無法建立或連接 Access 資料庫 '" & accessDbPath & "'。" & vbCrLf & Err.Description, vbCritical
             GoTo CleanUp ' 直接跳到清理
        End If
    End If
    Set fso = Nothing

    ' 2. 連接 CSV 檔案
    Set csvConn = CreateObject("ADODB.Connection")
    csvConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                 "Data Source=" & Left(csvFilePath, InStrRev(csvFilePath, "\")) & ";" & _
                 "Extended Properties=""text;HDR=Yes;FMT=Delimited;CharacterSet=" & csvEncoding & ";IMEX=1;"""
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM [" & Mid(csvFilePath, InStrRev(csvFilePath, "\") + 1) & "]", csvConn

    ' 3. 連接 Access 資料庫
    Set accConn = CreateObject("ADODB.Connection")
    accConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessDbPath & ";"

    ' 4. 刪除舊表 (如果存在)
    On Error Resume Next ' 忽略 "Table does not exist" 錯誤
    accConn.Execute "DROP TABLE [" & tableName & "];"
    On Error GoTo ErrorHandler

    ' 5. 根據 CSV 標頭建立新表
    fldDef = ""
    For i = 0 To rs.fields.Count - 1
        If i > 0 Then fldDef = fldDef & ", "
        ' 清理欄位名稱中的非法字元
        fieldName = Replace(rs.fields(i).name, ".", "_")
        fieldName = Replace(fieldName, "#", "_")
        fieldName = Replace(fieldName, "$", "_")
        fieldName = Replace(fieldName, "[", "_")
        fieldName = Replace(fieldName, "]", "_")
        ' 確保名稱不為空或純數字等 Access 不允許的名稱
        If Len(Trim(fieldName)) = 0 Then fieldName = "Column" & i
        If IsNumeric(Left(fieldName, 1)) Then fieldName = "_" & fieldName
        ' 簡單地將所有類型設為 TEXT，後續可根據需要調整
        fldDef = fldDef & "[" & fieldName & "] TEXT"
    Next i
    If Len(fldDef) = 0 Then
        MsgBox "CSV 檔案沒有有效的標頭行。", vbExclamation
        GoTo CleanUp
    End If
    accConn.Execute "CREATE TABLE [" & tableName & "] (" & fldDef & ")"

    ' 6. 複製資料
    Set destRs = CreateObject("ADODB.Recordset")
    destRs.Open "SELECT * FROM [" & tableName & "]", accConn, 3, 3 ' adOpenStatic, adLockOptimistic

    If Not rs.EOF Then
        rs.MoveFirst
        Do Until rs.EOF
            destRs.AddNew
            For i = 0 To rs.fields.Count - 1
                On Error Resume Next ' 處理單一欄位寫入錯誤
                destRs.fields(i).value = rs.fields(i).value
                If Err.Number <> 0 Then
                    Debug.Print "寫入 Access 錯誤: Row " & (destRs.AbsolutePosition + rs.AbsolutePosition - 1) & ", Field " & i & " (" & destRs.fields(i).name & "), Value: '" & rs.fields(i).value & "', Error: " & Err.Description
                    destRs.fields(i).value = Null ' 寫入錯誤時設為 Null
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
            Next i
            destRs.Update
            rs.MoveNext
        Loop
    End If

    ' 7. 成功
    ImportToAccess = True

CleanUp:
    ' 清理資源
    On Error Resume Next ' 忽略關閉時可能發生的錯誤
    If Not rs Is Nothing Then If rs.State = adStateOpen Then rs.Close
    If Not destRs Is Nothing Then If destRs.State = adStateOpen Then destRs.Close
    If Not csvConn Is Nothing Then If csvConn.State = adStateOpen Then csvConn.Close
    If Not accConn Is Nothing Then If accConn.State = adStateOpen Then accConn.Close
    Set rs = Nothing
    Set destRs = Nothing
    Set csvConn = Nothing
    Set accConn = Nothing
    Set fso = Nothing
    Set cat = Nothing
    Exit Function ' 從這裡退出函數

ErrorHandler:
    ' 記錄詳細錯誤信息
    Debug.Print "--- ImportToAccess Error ---"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "CSV File: " & csvFilePath
    Debug.Print "Access DB: " & accessDbPath
    Debug.Print "Table Name: " & tableName
    Debug.Print "CSV Encoding: " & csvEncoding
    ' 將錯誤傳遞給調用者 (ImportCSV)
    Err.Raise Err.Number, "ImportCSV.ImportToAccess", Err.Description
    GoTo CleanUp ' 確保資源被釋放
End Function

Public Function LoadToExcel(ByVal accessDbPath As String, ByVal tableName As String, ByVal worksheetName As String, ByVal maxRowsToShow As Long) As Boolean
    ' 從 Access 資料庫讀取資料並載入 Excel 工作表
    ' 返回 True 表示成功，False 表示失敗
    Dim conn As Object ' ADODB.Connection
    Dim rs As Object   ' ADODB.Recordset
    Dim ws As Worksheet
    Dim sql As String
    Dim i As Long
    Dim targetSheetExists As Boolean

    On Error GoTo ErrorHandler
    LoadToExcel = False ' 預設為失敗

    ' 1. 檢查並處理目標工作表
    Application.DisplayAlerts = False ' 暫時關閉提示
    On Error Resume Next ' 檢查工作表是否存在
    Set ws = ThisWorkbook.Worksheets(worksheetName)
    targetSheetExists = (Err.Number = 0)
    On Error GoTo ErrorHandler ' 恢復錯誤處理
    Application.DisplayAlerts = True ' 恢復提示

    If targetSheetExists Then
        ws.Cells.Clear ' 清空現有工作表內容
        Debug.Print "已清空現有工作表: " & worksheetName
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = worksheetName
        ' 嘗試修改 VBA CodeName (如果專案未受保護)
        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents(ws.CodeName).name = worksheetName ' 避免特殊字元問題
        If Err.Number <> 0 Then
             Debug.Print "警告: 無法修改工作表 '" & worksheetName & "' 的 VBA CodeName。"
             Err.Clear
        End If
        On Error GoTo ErrorHandler
        Debug.Print "已建立新工作表: " & worksheetName
    End If

    ' 2. 連接 Access 資料庫
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessDbPath & ";"

    ' 3. 建立記錄集 (限制行數)
    Set rs = CreateObject("ADODB.Recordset")
    If maxRowsToShow > 0 Then
        sql = "SELECT TOP " & maxRowsToShow & " * FROM [" & tableName & "]"
    Else
        sql = "SELECT * FROM [" & tableName & "]"
    End If
    rs.Open sql, conn, 0, 1 ' adOpenForwardOnly, adLockReadOnly

    ' 4. 寫入標題
    If Not rs.EOF Then ' 確保有資料
        For i = 0 To rs.fields.Count - 1
            ws.Cells(1, i + 1).value = rs.fields(i).name
        Next i

        ' 5. 複製資料
        ws.Range("A2").CopyFromRecordset rs

        ' 6. 處理日期格式 (如果需要)
        For i = 0 To rs.fields.Count - 1
            If rs.fields(i).Type = adDate Then ' adDate = 7
                ' 檢查是否有資料行被複製
                If ws.Range("A2").value <> "" Or ws.Range("A2").Offset(1, 0).value <> "" Then ' 簡單檢查
                     ws.Range("A2").Offset(0, i).Resize(WorksheetFunction.Min(rs.RecordCount, maxRowsToShow), 1).NumberFormat = "yyyy/mm/dd"
                End If
            End If
        Next i

        ' 7. 調整欄寬
        ws.Columns.AutoFit
    Else
        ' 如果 Access 表格是空的，只寫入標題
         For i = 0 To rs.fields.Count - 1
            ws.Cells(1, i + 1).value = rs.fields(i).name
        Next i
        ws.Columns.AutoFit
        MsgBox "Access 資料表 [" & tableName & "] 中沒有資料可載入。", vbInformation
    End If

    ' 8. 成功
    LoadToExcel = True

CleanUp:
    ' 清理資源
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = adStateOpen Then rs.Close
    If Not conn Is Nothing Then If conn.State = adStateOpen Then conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Set ws = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "--- LoadToExcel Error ---"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "Access DB: " & accessDbPath
    Debug.Print "Table Name: " & tableName
    Debug.Print "Worksheet Name: " & worksheetName
    Application.DisplayAlerts = True ' 確保提示被恢復
    Err.Raise Err.Number, "ImportCSV.LoadToExcel", Err.Description
    GoTo CleanUp
End Function

